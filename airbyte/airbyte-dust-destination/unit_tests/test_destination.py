#
# Copyright (c) 2023 Airbyte, Inc., all rights reserved.
#

import json
from typing import Any, Dict
from unittest import mock
from unittest.mock import Mock

from destination_dust.destination import DestinationDust

from airbyte_cdk.models import AirbyteConnectionStatus, AirbyteMessage, Status, Type
from airbyte_cdk.models.airbyte_protocol import (
    AirbyteRecordMessage,
    AirbyteStateMessage,
    AirbyteStream,
    ConfiguredAirbyteCatalog,
    ConfiguredAirbyteStream,
    DestinationSyncMode,
    SyncMode,
)


config = {
    "api_key": "sk-test",
    "workspace_id": "w1",
    "space_id": "s1",
    "data_source_id": "ds1",
    "base_url": "https://dust.tt",
}


def _init_mocks(client_init):
    """Patch DustClient and return a mock client."""
    mock_client = Mock()
    client_init.return_value = mock_client
    return mock_client


def _state() -> AirbyteMessage:
    return AirbyteMessage(type=Type.STATE, state=AirbyteStateMessage(data={}))


def _record(stream: str, data: Dict[str, Any]) -> AirbyteMessage:
    return AirbyteMessage(
        type=Type.RECORD,
        record=AirbyteRecordMessage(stream=stream, data=data, emitted_at=0),
    )


def _configured_catalog(stream_name: str = "people", primary_key: list = None) -> ConfiguredAirbyteCatalog:
    """Build a catalog with a single stream (e.g. 'people')."""
    stream_schema = {
        "type": "object",
        "properties": {"name": {"type": "string"}, "email": {"type": "string"}},
    }
    airbyte_stream = AirbyteStream(
        name=stream_name,
        json_schema=stream_schema,
        supported_sync_modes=[SyncMode.incremental],
    )
    append_stream = ConfiguredAirbyteStream(
        stream=airbyte_stream,
        sync_mode=SyncMode.incremental,
        destination_sync_mode=DestinationSyncMode.append,
        primary_key=primary_key or [["id"]],
    )
    return ConfiguredAirbyteCatalog(streams=[append_stream])


# --- Check ---


@mock.patch("destination_dust.destination.DustClient")
def test_check_succeeds(client_init):
    mock_client = _init_mocks(client_init)
    destination = DestinationDust()
    status = destination.check(logger=Mock(), config=config)
    assert status.status == Status.SUCCEEDED
    mock_client.check_connection.assert_called_once_with(data_format="documents")


@mock.patch("destination_dust.destination.DustClient")
def test_check_fails_on_connection(client_init):
    mock_client = _init_mocks(client_init)
    mock_client.check_connection.side_effect = Exception("Connection failed")
    destination = DestinationDust()
    status = destination.check(logger=Mock(), config=config)
    assert status.status == Status.FAILED
    assert "Connection failed" in status.message


@mock.patch("destination_dust.destination.DustClient")
def test_check_with_tables_format(client_init):
    mock_client = _init_mocks(client_init)
    destination = DestinationDust()
    tables_config = {**config, "data_format": "tables"}
    status = destination.check(logger=Mock(), config=tables_config)
    assert status.status == Status.SUCCEEDED
    mock_client.check_connection.assert_called_once_with(data_format="tables")


# --- Write (documents mode) ---


@mock.patch("destination_dust.destination.DustClient")
def test_write_succeeds(client_init):
    stream = "people"
    data = {"id": 1, "name": "John Doe", "email": "john.doe@example.com"}
    mock_client = _init_mocks(client_init)
    input_messages = [_record(stream=stream, data=data), _state()]
    destination = DestinationDust()
    messages = list(
        destination.write(
            config=config,
            configured_catalog=_configured_catalog(stream_name=stream),
            input_messages=input_messages,
        )
    )
    # LOG, (optional logs), STATE
    state_msgs = [m for m in messages if m.type == Type.STATE]
    assert len(state_msgs) >= 1
    mock_client.upsert_document.assert_called_once()
    call_kwargs = mock_client.upsert_document.call_args.kwargs
    assert call_kwargs["document_id"] == "people-1"
    assert call_kwargs["title"] == "John Doe"
    assert call_kwargs["tags"] == ["airbyte:stream:people"]
    assert json.loads(call_kwargs["text"]) == data


@mock.patch("destination_dust.destination.DustClient")
def test_write_succeeds_with_custom_base_url(client_init):
    stream = "people"
    data = {"id": 1, "name": "Jane"}
    mock_client = _init_mocks(client_init)
    custom_config = {**config, "base_url": "https://eu.dust.tt"}
    input_messages = [_record(stream=stream, data=data), _state()]
    destination = DestinationDust()
    list(
        destination.write(
            config=custom_config,
            configured_catalog=_configured_catalog(stream_name=stream),
            input_messages=input_messages,
        )
    )
    mock_client.upsert_document.assert_called_once()
    # Client was constructed with custom config (base_url is used in client init)
    client_init.assert_called_once()
    assert client_init.call_args[0][0]["base_url"] == "https://eu.dust.tt"  # first positional arg is config


@mock.patch("destination_dust.destination.DustClient")
def test_write_processes_message_from_unknown_stream(client_init):
    """When stream is not in catalog, Dust still writes the record (hash-based document id)."""
    stream = "shapes"
    data = {"name": "Rectangle", "color": "blue"}
    mock_client = _init_mocks(client_init)
    # Catalog only has "people", not "shapes"
    input_messages = [_record(stream=stream, data=data), _state()]
    destination = DestinationDust()
    list(
        destination.write(
            config=config,
            configured_catalog=_configured_catalog(stream_name="people"),
            input_messages=input_messages,
        )
    )
    # Dust writes every record; document_id is stream-hash when stream not in catalog
    mock_client.upsert_document.assert_called_once()
    call_kwargs = mock_client.upsert_document.call_args.kwargs
    assert call_kwargs["document_id"].startswith("shapes-")
    assert call_kwargs["tags"] == ["airbyte:stream:shapes"]


# --- Write (tables mode) ---


@mock.patch("destination_dust.destination.DustClient")
def test_write_succeeds_tables_mode(client_init):
    stream = "people"
    data = {"id": 1, "name": "John Doe", "email": "john.doe@example.com"}
    mock_client = _init_mocks(client_init)
    # No existing table -> will call upsert_table; must return table_id for upsert_rows
    mock_client.find_table_by_title.return_value = None
    mock_client.upsert_table.return_value = {"table_id": "test-table-id"}
    tables_config = {**config, "data_format": "tables", "table_id_prefix": "airbyte_"}
    input_messages = [_record(stream=stream, data=data), _state()]
    destination = DestinationDust()
    list(
        destination.write(
            config=tables_config,
            configured_catalog=_configured_catalog(stream_name=stream),
            input_messages=input_messages,
        )
    )
    mock_client.upsert_table.assert_called_once()
    mock_client.upsert_rows.assert_called_once()
    table_call = mock_client.upsert_table.call_args.kwargs
    assert table_call["name"] == stream
    assert table_call["title"] == stream
    rows = mock_client.upsert_rows.call_args[0][1]  # positional: (table_id, rows)
    assert len(rows) == 1
    assert rows[0]["name"] == "John Doe"


# --- Helpers: _build_document_id ---


def test_build_document_id_with_single_primary_key():
    stream = Mock()
    stream.primary_key = [["id"]]
    data = {"id": 42, "name": "Alice"}
    result = DestinationDust._build_document_id("users", data, stream)
    assert result == "users-42"


def test_build_document_id_with_composite_primary_key():
    stream = Mock()
    stream.primary_key = [["org_id"], ["user_id"]]
    data = {"org_id": "acme", "user_id": 7}
    result = DestinationDust._build_document_id("members", data, stream)
    assert result == "members-acme-7"


def test_build_document_id_without_primary_key():
    stream = Mock()
    stream.primary_key = []
    data = {"foo": "bar", "num": 1}
    result = DestinationDust._build_document_id("events", data, stream)
    assert result.startswith("events-")
    assert len(result) == len("events-") + 16


def test_build_document_id_without_primary_key_is_deterministic():
    stream = Mock()
    stream.primary_key = []
    data = {"a": 1, "b": 2}
    id1 = DestinationDust._build_document_id("s", data, stream)
    id2 = DestinationDust._build_document_id("s", data, stream)
    assert id1 == id2


def test_build_document_id_sanitizes_special_characters():
    stream = Mock()
    stream.primary_key = [["id"]]
    data = {"id": "hello world/foo@bar"}
    result = DestinationDust._build_document_id("s", data, stream)
    assert result == "s-hello_world_foo_bar"


# --- Helpers: _build_title ---


def test_build_title_uses_name_field():
    assert DestinationDust._build_title("s", {"name": "Alice"}) == "Alice"


def test_build_title_prefers_title_over_name():
    assert DestinationDust._build_title("s", {"title": "T", "name": "N"}) == "T"


def test_build_title_fallback_to_stream_name():
    assert DestinationDust._build_title("users", {"id": 1}) == "users record"


# --- Helpers: _build_table_id ---


def test_build_table_id_basic():
    assert DestinationDust._build_table_id("users", "airbyte_") == "airbyte_users"


def test_build_table_id_sanitizes_unsafe_characters():
    assert DestinationDust._build_table_id("stream@name", "p_") == "p_stream_name"


# --- Helpers: _flatten_record ---


def test_flatten_record_preserves_primitives():
    data = {"id": 1, "name": "Alice", "active": True}
    result = DestinationDust._flatten_record(data)
    assert result == data


def test_flatten_record_flattens_nested_objects():
    data = {"id": 1, "meta": {"key": "value"}}
    result = DestinationDust._flatten_record(data)
    assert result["id"] == 1
    assert result["meta"] == '{"key": "value"}'
