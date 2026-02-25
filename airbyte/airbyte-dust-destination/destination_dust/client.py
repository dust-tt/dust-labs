import json
import logging
from typing import Any, Callable, List, Mapping, Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

logger = logging.getLogger("airbyte")

RETRY_TOTAL = 3
RETRY_BACKOFF_FACTOR = 1.0  # 1s, 2s, 4s
RETRY_STATUS_FORCELIST = [429, 500, 502, 503, 504]


class DustClient:
    """HTTP client for the Dust document and table upsert APIs."""

    def __init__(self, config: Mapping[str, Any], log_callback=None):
        """
        Initialize Dust client.
        
        Args:
            config: Configuration dictionary
            log_callback: Optional callback function(message: str, level: str) for logging
        """
        self.api_key = config["api_key"]
        self.workspace_id = config["workspace_id"]
        self.space_id = config["space_id"]
        self.data_source_id = config["data_source_id"]
        self.base_url = config.get("base_url", "https://dust.tt").rstrip("/")
        self.log_callback = log_callback

        self._documents_base = (
            f"{self.base_url}/api/v1/w/{self.workspace_id}"
            f"/spaces/{self.space_id}"
            f"/data_sources/{self.data_source_id}"
            f"/documents"
        )

        # Tables API uses underscores in path (data_sources not data-sources)
        self._tables_base = (
            f"{self.base_url}/api/v1/w/{self.workspace_id}"
            f"/spaces/{self.space_id}"
            f"/data_sources/{self.data_source_id}"
            f"/tables"
        )

        self._session = requests.Session()
        retry = Retry(
            total=RETRY_TOTAL,
            backoff_factor=RETRY_BACKOFF_FACTOR,
            status_forcelist=RETRY_STATUS_FORCELIST,
            allowed_methods=["GET", "POST"],
        )
        adapter = HTTPAdapter(max_retries=retry)
        self._session.mount("https://", adapter)
        self._session.mount("http://", adapter)
        self._session.headers.update(
            {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json",
            }
        )

    def check_connection(self, data_format: str = "documents") -> None:
        """
        Verify credentials and data source existence.

        For documents: GETs /documents?limit=1
        For tables: GETs /tables (list tables)
        """
        if data_format == "tables":
            url = self._tables_base
            response = self._session.get(url, timeout=30)
        else:
            url = f"{self._documents_base}?limit=1"
            response = self._session.get(url, timeout=30)

        if response.status_code == 401:
            raise ConnectionError("Authentication failed. Check your API key.")
        if response.status_code == 404:
            raise ConnectionError(
                f"Data source not found. Verify workspace_id='{self.workspace_id}', "
                f"space_id='{self.space_id}', data_source_id='{self.data_source_id}'."
            )
        if not response.ok:
            raise ConnectionError(
                f"Dust API returned status {response.status_code}: "
                f"{response.text[:500]}"
            )

        logger.info("Dust connection check succeeded.")

    def list_tables(self) -> List[dict[str, Any]]:
        """
        List all tables in the configured Dust data source.

        Returns:
            List of table dictionaries containing table metadata (id, name, title, etc.)

        Raises RuntimeError on API errors after retries are exhausted.
        """
        url = self._tables_base

        # Log request
        request_log = "Listing tables"
        logger.debug(request_log)
        if self.log_callback:
            self.log_callback(f"Request: GET {url}", "DEBUG")

        response = self._session.get(url, timeout=30)

        # Log response
        response_log = f"Listed tables successfully (status: {response.status_code})"
        logger.debug(response_log)
        if self.log_callback:
            self.log_callback(
                f"Response: {response.status_code}\nBody: {response.text[:500]}",
                "DEBUG"
            )

        if response.status_code == 429:
            raise RuntimeError(
                f"Rate limited by Dust API after {RETRY_TOTAL} retries. "
                "Consider reducing sync frequency."
            )

        if not response.ok:
            raise RuntimeError(
                f"Failed to list tables: "
                f"status={response.status_code}, body={response.text[:500]}"
            )

        result = response.json()
        # Handle different response formats: could be {"tables": [...]} or just [...]
        if isinstance(result, dict) and "tables" in result:
            return result["tables"]
        elif isinstance(result, list):
            return result
        else:
            logger.warning(f"Unexpected list_tables response format: {result}")
            return []

    def find_table_by_title(self, title: str) -> Optional[str]:
        """
        Find a table ID by its title.

        Args:
            title: The table title to search for

        Returns:
            The table ID if found, None otherwise
        """
        tables = self.list_tables()
        for table in tables:
            # Check both 'title' and 'name' fields (API might use either)
            table_title = table.get("title") or table.get("name", "")
            if table_title == title:
                table_id = table.get("id") or table.get("table_id")
                if table_id:
                    logger.info(f"Found table '{title}' with ID: {table_id}")
                    if self.log_callback:
                        self.log_callback(
                            f"Found table '{title}' with ID: {table_id}",
                            "INFO"
                        )
                    return table_id
        logger.debug(f"Table with title '{title}' not found")
        return None

    def upsert_document(
        self,
        document_id: str,
        title: str,
        text: str,
        source_url: str = "",
        tags: Optional[List[str]] = None,
        timestamp: Optional[int] = None,
    ) -> dict:
        """
        Upsert a document into the configured Dust data source.

        Raises RuntimeError on API errors after retries are exhausted.
        """
        url = f"{self._documents_base}/{document_id}"
        payload: dict[str, Any] = {
            "title": title,
            "mime_type": "application/json",
            "text": text,
            "source_url": source_url,
            "tags": tags or [],
            "light_document_output": True,
            "async": False,
        }
        if timestamp is not None:
            payload["timestamp"] = timestamp

        # Log request
        request_log = f"Upserting document '{document_id}' (title: {title})"
        logger.debug(request_log)
        if self.log_callback:
            self.log_callback(f"Request: POST {url}\nDocument ID: {document_id}\nTitle: {title}", "DEBUG")

        response = self._session.post(url, json=payload, timeout=60)

        # Log response
        response_log = f"Document '{document_id}' upserted successfully (status: {response.status_code})"
        logger.debug(response_log)
        if self.log_callback:
            self.log_callback(f"Response: {response.status_code}\nBody: {response.text[:200]}", "DEBUG")

        if response.status_code == 429:
            raise RuntimeError(
                f"Rate limited by Dust API after {RETRY_TOTAL} retries. "
                "Consider reducing sync frequency."
            )

        if not response.ok:
            raise RuntimeError(
                f"Failed to upsert document '{document_id}': "
                f"status={response.status_code}, body={response.text[:500]}"
            )

        return response.json()

    def upsert_table(
        self,
        name: str,
        title: str = "",
        description: str = "",
        table_id: Optional[str] = None,
    ) -> dict:
        """
        Create or update a table definition in the configured Dust data source.

        Args:
            name: Human-readable table name
            title: Table title (optional)
            description: Optional table description
            table_id: Optional unique identifier for the table. If not provided, Dust will generate one.

        Returns:
            API response containing table metadata including table ID

        Raises RuntimeError on API errors after retries are exhausted.
        """
        url = self._tables_base
        # Only include fields specified in API docs: id, name, title, description
        payload: dict[str, Any] = {
            "name": name,
        }
        # Only include id if provided - let Dust generate it otherwise
        if table_id is not None:
            payload["id"] = table_id
        # Only include title if explicitly provided (not empty)
        if title:
            payload["title"] = title
        if description:
            payload["description"] = description

        # Log request
        request_log = f"Creating/updating table '{name}' (id: {table_id or 'auto-generated'})"
        logger.info(request_log)
        if self.log_callback:
            self.log_callback(request_log, "INFO")
        
        logger.debug(f"upsert_table: POST {url}")
        logger.debug(f"upsert_table: payload={payload}")
        if self.log_callback:
            self.log_callback(f"Request: POST {url}\nPayload: {json.dumps(payload, indent=2)}", "DEBUG")

        response = self._session.post(url, json=payload, timeout=60)

        # Log response
        response_log = f"Table '{name}' created/updated successfully (status: {response.status_code})"
        logger.info(response_log)
        if self.log_callback:
            self.log_callback(response_log, "INFO")
        
        logger.debug(f"upsert_table: status_code={response.status_code}")
        logger.debug(f"upsert_table: response_headers={dict(response.headers)}")
        logger.debug(f"upsert_table: response_body={response.text}")
        if self.log_callback:
            self.log_callback(
                f"Response: {response.status_code}\nBody: {response.text[:500]}", 
                "DEBUG"
            )

        if response.status_code == 429:
            raise RuntimeError(
                f"Rate limited by Dust API after {RETRY_TOTAL} retries. "
                "Consider reducing sync frequency."
            )

        if not response.ok:
            error_msg = f"Failed to upsert table"
            if table_id:
                error_msg += f" '{table_id}'"
            error_msg += f": status={response.status_code}, body={response.text[:500]}"
            raise RuntimeError(error_msg)

        return response.json()

    def upsert_rows(
        self,
        table_id: str,
        rows: List[dict[str, Any]],
    ) -> dict:
        """
        Upsert rows into an existing table.

        Args:
            table_id: The table ID to upsert rows into
            rows: List of row objects (dicts with column names as keys)

        Returns:
            API response

        Raises RuntimeError on API errors after retries are exhausted.
        """
        url = f"{self._tables_base}/{table_id}/rows"
        
        # Format rows for Dust API: each row needs row_id and value fields
        formatted_rows = []
        for row in rows:
            # Use 'id' field as row_id if present, otherwise generate one
            row_id = str(row.get("id", ""))
            if not row_id:
                # Generate row_id from first non-empty field value
                for key, value in row.items():
                    if value is not None and str(value).strip():
                        row_id = str(value)
                        break
                if not row_id:
                    row_id = str(hash(str(row)))[:16]  # Fallback to hash
            
            formatted_rows.append({
                "row_id": row_id,
                "value": row
            })
        
        payload = {"rows": formatted_rows}

        # Log request
        request_log = f"Upserting {len(formatted_rows)} rows into table '{table_id}'"
        logger.info(request_log)
        if self.log_callback:
            self.log_callback(request_log, "INFO")
        
        logger.debug(f"upsert_rows: POST {url}")
        logger.debug(f"upsert_rows: table_id={table_id}, row_count={len(formatted_rows)}")
        logger.debug(f"upsert_rows: payload={payload}")
        if self.log_callback:
            # Show sample of first row for debugging
            sample_payload = {"rows": formatted_rows[:1]} if formatted_rows else {"rows": []}
            self.log_callback(
                f"Request: POST {url}\nRow count: {len(formatted_rows)}\nPayload sample: {json.dumps(sample_payload, indent=2)[:500]}",
                "DEBUG"
            )

        response = self._session.post(url, json=payload, timeout=60)

        # Log response
        response_log = f"Successfully upserted {len(formatted_rows)} rows into table '{table_id}' (status: {response.status_code})"
        logger.info(response_log)
        if self.log_callback:
            self.log_callback(response_log, "INFO")
        
        logger.debug(f"upsert_rows: status_code={response.status_code}")
        logger.debug(f"upsert_rows: response_headers={dict(response.headers)}")
        logger.debug(f"upsert_rows: response_body={response.text}")
        if self.log_callback:
            self.log_callback(
                f"Response: {response.status_code}\nBody: {response.text[:500]}",
                "DEBUG"
            )

        if response.status_code == 429:
            raise RuntimeError(
                f"Rate limited by Dust API after {RETRY_TOTAL} retries. "
                "Consider reducing sync frequency."
            )

        if not response.ok:
            raise RuntimeError(
                f"Failed to upsert rows into table '{table_id}': "
                f"status={response.status_code}, body={response.text[:500]}"
            )

        return response.json()
