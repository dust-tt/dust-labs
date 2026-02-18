import hashlib
import json
import logging
import re
import uuid
from collections import defaultdict
from typing import Any, Iterable, List, Mapping, Optional

from airbyte_cdk.destinations import Destination
from airbyte_cdk.models import (
    AirbyteConnectionStatus,
    AirbyteLogMessage,
    AirbyteMessage,
    AirbyteStateMessage,
    ConfiguredAirbyteCatalog,
    ConnectorSpecification,
    Level,
    Status,
    Type,
)

from destination_dust.client import DustClient

logger = logging.getLogger("airbyte")

# Default batch size for table row upserts (configurable via table_batch_size)
DEFAULT_TABLE_BATCH_SIZE = 500


def _create_log_message(level: Level, message: str) -> AirbyteMessage:
    """Create an AirbyteLogMessage wrapped in AirbyteMessage."""
    return AirbyteMessage(
        type=Type.LOG,
        log=AirbyteLogMessage(level=level, message=message)
    )


def _ensure_state_has_id(message: AirbyteMessage) -> AirbyteMessage:
    """
    Ensure state message has an id field. If missing, add one.
    
    Airbyte requires state messages to have an id field for proper state tracking.
    The id should be on the AirbyteStateMessage object.
    """
    if message.type != Type.STATE:
        return message
    
    # Check if state message already has an id
    if message.state:
        # Try to get id from state message
        state_id = getattr(message.state, 'id', None)
        if state_id:
            return message
        
        # Get state data
        state_data = getattr(message.state, 'data', {}) or {}
    else:
        state_data = {}
    
    # Generate an id from state data hash if no id exists
    state_id = hashlib.sha256(json.dumps(state_data, sort_keys=True, default=str).encode()).hexdigest()[:16]
    
    # Create new state message with id
    # Try to preserve other attributes if they exist
    try:
        new_state = AirbyteStateMessage(
            data=state_data,
            id=state_id
        )
    except TypeError:
        # If AirbyteStateMessage doesn't accept id parameter directly, 
        # try creating it and setting id as attribute
        new_state = AirbyteStateMessage(data=state_data)
        if hasattr(new_state, 'id'):
            new_state.id = state_id
    
    return AirbyteMessage(
        type=Type.STATE,
        state=new_state
    )


class DestinationDust(Destination):
    def spec(self, *args: Any, **kwargs: Any) -> ConnectorSpecification:
        return super().spec(*args, **kwargs)

    def check(
        self, logger: logging.Logger, config: Mapping[str, Any]
    ) -> AirbyteConnectionStatus:
        try:
            client = DustClient(config)
            data_format = config.get("data_format", "documents")
            client.check_connection(data_format=data_format)
            return AirbyteConnectionStatus(status=Status.SUCCEEDED)
        except Exception as e:
            return AirbyteConnectionStatus(
                status=Status.FAILED,
                message=f"Connection check failed: {str(e)}",
            )

    def write(
        self,
        config: Mapping[str, Any],
        configured_catalog: ConfiguredAirbyteCatalog,
        input_messages: Iterable[AirbyteMessage],
    ) -> Iterable[AirbyteMessage]:
        data_format = config.get("data_format", "documents")
        
        # Create log callback to yield log messages
        log_messages = []
        def log_callback(message: str, level: str) -> None:
            log_level = Level.INFO if level == "INFO" else Level.DEBUG if level == "DEBUG" else Level.ERROR
            log_messages.append(_create_log_message(log_level, message))
        
        client = DustClient(config, log_callback=log_callback)
        
        yield _create_log_message(Level.INFO, f"Starting sync to Dust (format: {data_format})")
        
        if data_format == "tables":
            yield from self._write_tables(client, config, configured_catalog, input_messages, log_messages)
        else:
            yield from self._write_documents(client, configured_catalog, input_messages, log_messages)
        
        yield _create_log_message(Level.INFO, "Sync to Dust completed successfully")

    def _write_documents(
        self,
        client: DustClient,
        configured_catalog: ConfiguredAirbyteCatalog,
        input_messages: Iterable[AirbyteMessage],
        log_messages: List[AirbyteMessage],
    ) -> Iterable[AirbyteMessage]:
        """Write records as documents (original behavior)."""
        streams = {
            stream.stream.name: stream for stream in configured_catalog.streams
        }
        
        record_count = 0
        stream_counts: dict[str, int] = {}

        for message in input_messages:
            if message.type == Type.STATE:
                # Yield any pending log messages before state
                yield from log_messages
                log_messages.clear()
                # Ensure state message has an id field
                yield _ensure_state_has_id(message)

            elif message.type == Type.RECORD:
                record = message.record
                stream_name = record.stream
                data = record.data
                
                record_count += 1
                stream_counts[stream_name] = stream_counts.get(stream_name, 0) + 1

                configured_stream = streams.get(stream_name)
                document_id = self._build_document_id(
                    stream_name, data, configured_stream
                )
                title = self._build_title(stream_name, data)
                text = json.dumps(data, indent=2, default=str)
                tags = [f"airbyte:stream:{stream_name}"]
                timestamp = record.emitted_at

                client.upsert_document(
                    document_id=document_id,
                    title=title,
                    text=text,
                    tags=tags,
                    timestamp=timestamp,
                )
        
        # Yield final log messages
        yield from log_messages
        log_messages.clear()
        yield _create_log_message(Level.INFO, f"Processed {record_count} documents across {len(stream_counts)} stream(s)")

    def _write_tables(
        self,
        client: DustClient,
        config: Mapping[str, Any],
        configured_catalog: ConfiguredAirbyteCatalog,
        input_messages: Iterable[AirbyteMessage],
        log_messages: List[AirbyteMessage],
    ) -> Iterable[AirbyteMessage]:
        """Write records as table rows with batching."""
        streams = {
            stream.stream.name: stream for stream in configured_catalog.streams
        }
        batch_size = config.get("table_batch_size", DEFAULT_TABLE_BATCH_SIZE)

        yield _create_log_message(Level.INFO, f"Processing {len(streams)} stream(s) in tables mode")

        # Collect records by stream
        stream_rows: dict[str, List[dict[str, Any]]] = defaultdict(list)
        table_ids: dict[str, str] = {}  # stream_name -> table_id (looked up by table title)
        record_count = 0

        for message in input_messages:
            if message.type == Type.STATE:
                # Flush any pending rows before yielding state
                yield from log_messages
                log_messages.clear()
                self._flush_table_batches(
                    client, stream_rows, table_ids, streams, batch_size
                )
                stream_rows.clear()
                # Ensure state message has an id field
                yield _ensure_state_has_id(message)

            elif message.type == Type.RECORD:
                record = message.record
                stream_name = record.stream
                data = record.data
                
                record_count += 1

                if stream_name not in stream_rows:
                    yield _create_log_message(Level.INFO, f"Discovered stream: {stream_name}")

                # Flatten nested objects to JSON strings for now
                flattened_data = self._flatten_record(data)
                stream_rows[stream_name].append(flattened_data)

                # Batch and flush when batch size reached
                if len(stream_rows[stream_name]) >= batch_size:
                    if stream_name not in table_ids:
                        # Lookup table by title (stream_name), create if not found
                        table_ids[stream_name] = self._ensure_table_exists(
                            client, stream_name, streams.get(stream_name)
                        )

                    rows_batch = stream_rows[stream_name][:batch_size]
                    client.upsert_rows(table_ids[stream_name], rows_batch)
                    stream_rows[stream_name] = stream_rows[stream_name][batch_size:]

        # Flush remaining rows
        yield from log_messages
        log_messages.clear()
        self._flush_table_batches(
            client, stream_rows, table_ids, streams, batch_size
        )
        yield _create_log_message(Level.INFO, f"Processed {record_count} records across {len(stream_rows)} stream(s)")

    def _flush_table_batches(
        self,
        client: DustClient,
        stream_rows: dict[str, List[dict[str, Any]]],
        table_ids: dict[str, str],
        streams: dict[str, Any],
        batch_size: int,
    ) -> None:
        """Flush all pending rows for all streams."""
        for stream_name, rows in stream_rows.items():
            if not rows:
                continue

            if stream_name not in table_ids:
                # Lookup table by title (stream_name), create if not found
                table_ids[stream_name] = self._ensure_table_exists(
                    client, stream_name, streams.get(stream_name)
                )

            # Flush in batches
            for i in range(0, len(rows), batch_size):
                batch = rows[i:i + batch_size]
                client.upsert_rows(table_ids[stream_name], batch)

    def _ensure_table_exists(
        self,
        client: DustClient,
        stream_name: str,
        configured_stream: Any,
    ) -> str:
        """
        Ensure table exists, using table title as identifier.

        First looks up the table by title (stream_name). If found, returns its ID.
        If not found, creates the table and returns the generated ID.
        Dust infers table schema from row data; we do not pass column types.

        Args:
            client: Dust client instance
            stream_name: Name of the stream/table (used as table title)
            configured_stream: Configured stream (unused, kept for compatibility)

        Returns:
            The table_id (looked up by title or generated by Dust if creating new table)
        """
        table_title = stream_name
        
        # First, try to find existing table by title
        existing_table_id = client.find_table_by_title(table_title)
        if existing_table_id:
            return existing_table_id
        
        # Table doesn't exist, create it
        # Note: Dust API infers table schema from row data, so we don't pass columns
        # Primary keys are handled by Dust API based on row data
        response = client.upsert_table(
            name=stream_name,
            title=table_title,  # Use stream name as table title
            description=f"Airbyte stream: {stream_name}",
            table_id=None,  # Let Dust generate the ID
        )
        
        # Extract table_id from response
        table_id = None
        if isinstance(response, dict):
            if "table" in response and isinstance(response["table"], dict):
                table_id = response["table"].get("table_id") or response["table"].get("id")
            if not table_id:
                table_id = response.get("id") or response.get("table_id")
        
        if not table_id:
            raise RuntimeError(f"Failed to extract table_id from API response: {response}")
        
        return table_id

    @staticmethod
    def _build_document_id(
        stream_name: str,
        data: Mapping[str, Any],
        configured_stream: Any,
    ) -> str:
        """
        Build a deterministic document ID from stream name and primary key.

        Falls back to a SHA-256 hash prefix of the record data when no
        primary key is defined.
        """
        pk_parts: list[str] = []
        if (
            configured_stream
            and configured_stream.primary_key
            and len(configured_stream.primary_key) > 0
        ):
            for key_path in configured_stream.primary_key:
                value: Any = data
                for key in key_path:
                    if isinstance(value, dict):
                        value = value.get(key, "")
                    else:
                        value = ""
                        break
                pk_parts.append(str(value))

        if pk_parts:
            raw_id = f"{stream_name}-{'-'.join(pk_parts)}"
        else:
            data_str = json.dumps(data, sort_keys=True, default=str)
            data_hash = hashlib.sha256(data_str.encode()).hexdigest()[:16]
            raw_id = f"{stream_name}-{data_hash}"

        return re.sub(r"[^a-zA-Z0-9_\-]", "_", raw_id)

    @staticmethod
    def _build_title(stream_name: str, data: Mapping[str, Any]) -> str:
        """
        Extract a human-readable title from the record data.

        Checks common title-like field names; falls back to stream name.
        """
        for candidate in ("title", "name", "subject", "headline", "label"):
            if candidate in data and data[candidate]:
                return str(data[candidate])
        return f"{stream_name} record"

    @staticmethod
    def _build_table_id(stream_name: str, prefix: str) -> str:
        """Build a table ID from stream name and prefix. Uses UUID as default if prefix is empty."""
        if not prefix:
            return str(uuid.uuid4())
        raw_id = f"{prefix}{stream_name}"
        return re.sub(r"[^a-zA-Z0-9_\-]", "_", raw_id)

    @staticmethod
    def _flatten_record(data: Mapping[str, Any]) -> dict[str, Any]:
        """
        Flatten a record for table storage.

        Nested objects and arrays are serialized to JSON strings.
        """
        flattened = {}
        for key, value in data.items():
            if isinstance(value, (dict, list)):
                flattened[key] = json.dumps(value, default=str)
            else:
                flattened[key] = value
        return flattened
