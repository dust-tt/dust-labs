#!/usr/bin/env python3
"""
CLI script to import CSV files into Dust tables.

Usage:
    python scripts/csv_to_dust.py <csv_file> [--table-id TABLE_ID] [--table-name TABLE_NAME]

The script reads configuration from a .env file in the project root.
"""

import argparse
import csv
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any, Dict, List

try:
    from dotenv import load_dotenv
except ImportError:
    print("Error: python-dotenv is required. Install it with: pip install python-dotenv")
    sys.exit(1)

# Add parent directory to path to import destination_dust
sys.path.insert(0, str(Path(__file__).parent.parent))

from destination_dust.client import DustClient

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


def infer_column_type(value: str) -> str:
    """
    Infer Dust column type from a CSV value string.

    Returns: "string", "number", "boolean", or "json"
    """
    if not value or value.strip() == "":
        return "string"  # Default for empty values

    value = value.strip()

    # Try boolean
    if value.lower() in ("true", "false", "yes", "no", "1", "0"):
        return "boolean"

    # Try number (int or float)
    try:
        float(value)
        # Check if it's an integer
        if "." not in value:
            return "number"
        return "number"
    except ValueError:
        pass

    # Try JSON
    if value.startswith("{") or value.startswith("["):
        try:
            json.loads(value)
            return "json"
        except (json.JSONDecodeError, ValueError):
            pass

    # Default to string
    return "string"


def infer_schema_from_csv(csv_file: str) -> Dict[str, str]:
    """
    Read the first few rows of CSV to infer column types.

    Returns: dict mapping column names to types
    """
    schema = {}
    sample_size = 100  # Sample first 100 rows for type inference

    with open(csv_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows_read = 0

        for row in reader:
            rows_read += 1
            for col_name, value in row.items():
                if col_name not in schema:
                    schema[col_name] = None

                # Infer type from this value
                inferred_type = infer_column_type(value)

                # Update schema with most specific type seen so far
                if schema[col_name] is None:
                    schema[col_name] = inferred_type
                elif schema[col_name] == "string":
                    # Can upgrade from string to more specific type
                    schema[col_name] = inferred_type
                elif schema[col_name] == "boolean" and inferred_type == "number":
                    # Number is more specific than boolean
                    schema[col_name] = inferred_type

            if rows_read >= sample_size:
                break

    # Ensure all columns have a type (default to string)
    for col_name in schema:
        if schema[col_name] is None:
            schema[col_name] = "string"

    return schema


def read_csv_rows(csv_file: str) -> List[Dict[str, Any]]:
    """
    Read all rows from CSV file.

    Returns: List of dictionaries, one per row
    """
    rows = []
    with open(csv_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            # Convert empty strings to None for cleaner data
            processed_row = {}
            for key, value in row.items():
                if value.strip() == "":
                    processed_row[key] = None
                else:
                    # Try to convert to appropriate type
                    value_type = infer_column_type(value)
                    if value_type == "number":
                        try:
                            if "." in value:
                                processed_row[key] = float(value)
                            else:
                                processed_row[key] = int(value)
                        except ValueError:
                            processed_row[key] = value
                    elif value_type == "boolean":
                        processed_row[key] = value.lower() in ("true", "yes", "1")
                    else:
                        processed_row[key] = value
            rows.append(processed_row)

    return rows


def ensure_title_column(rows: List[Dict[str, Any]], table_name: str) -> None:
    """
    Ensure every row has a 'title' column (mandatory for Dust tables).
    Uses first available title-like column or generates one.
    """
    title_candidates = ["title", "name", "subject", "headline", "label"]

    for row in rows:
        if "title" not in row or not row["title"]:
            # Try to find a title from candidates
            title = None
            for candidate in title_candidates:
                if candidate in row and row[candidate]:
                    title = str(row[candidate])
                    break

            # Fallback: use first non-empty column or table name
            if not title:
                for key, value in row.items():
                    if value and key != "title":
                        title = str(value)
                        break
                if not title:
                    title = f"{table_name} row"

            row["title"] = title


def main():
    parser = argparse.ArgumentParser(
        description="Import a CSV file into a Dust table",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python scripts/csv_to_dust.py data.csv
  python scripts/csv_to_dust.py data.csv --table-id my_table
  python scripts/csv_to_dust.py data.csv --table-name "My Table" --table-id my_table

Configuration is read from .env file in the project root.
Required variables: DUST_API_KEY, DUST_WORKSPACE_ID, DUST_SPACE_ID, DUST_DATA_SOURCE_ID
Optional variables: DUST_BASE_URL (default: https://dust.tt)
        """
    )
    parser.add_argument(
        "csv_file",
        help="Path to CSV file to import"
    )
    parser.add_argument(
        "--table-id",
        help="Table ID (default: derived from CSV filename)",
        default=None
    )
    parser.add_argument(
        "--table-name",
        help="Table name/title (default: derived from CSV filename)",
        default=None
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=500,
        help="Number of rows to batch per API request (default: 500)"
    )

    args = parser.parse_args()

    # Load environment variables from .env file
    env_path = Path(__file__).parent.parent / ".env"
    if not env_path.exists():
        logger.error(f".env file not found at {env_path}")
        logger.error("Please create a .env file based on .env.example")
        sys.exit(1)

    load_dotenv(env_path)

    # Read required configuration
    api_key = os.getenv("DUST_API_KEY")
    workspace_id = os.getenv("DUST_WORKSPACE_ID")
    space_id = os.getenv("DUST_SPACE_ID")
    data_source_id = os.getenv("DUST_DATA_SOURCE_ID")
    base_url = os.getenv("DUST_BASE_URL", "https://dust.tt")

    if not all([api_key, workspace_id, space_id, data_source_id]):
        logger.error("Missing required environment variables:")
        missing = []
        if not api_key:
            missing.append("DUST_API_KEY")
        if not workspace_id:
            missing.append("DUST_WORKSPACE_ID")
        if not space_id:
            missing.append("DUST_SPACE_ID")
        if not data_source_id:
            missing.append("DUST_DATA_SOURCE_ID")
        logger.error(f"Missing: {', '.join(missing)}")
        sys.exit(1)

    # Check CSV file exists
    csv_path = Path(args.csv_file)
    if not csv_path.exists():
        logger.error(f"CSV file not found: {csv_path}")
        sys.exit(1)

    # Determine table ID and name
    if args.table_id:
        table_id = args.table_id
    else:
        # Derive from filename: "data.csv" -> "data"
        table_id = csv_path.stem.lower().replace(" ", "_").replace("-", "_")
        # Sanitize to safe characters
        table_id = "".join(c for c in table_id if c.isalnum() or c in ("_", "-"))

    if args.table_name:
        table_name = args.table_name
    else:
        table_name = csv_path.stem.replace("_", " ").replace("-", " ").title()

    logger.info(f"Reading CSV file: {csv_path}")
    logger.info(f"Table ID: {table_id}")
    logger.info(f"Table name: {table_name}")

    # Infer schema from CSV
    logger.info("Inferring schema from CSV...")
    schema = infer_schema_from_csv(str(csv_path))
    logger.info(f"Inferred {len(schema)} columns: {', '.join(schema.keys())}")

    # Note: title column is not added to schema as it causes API errors
    # The table title is set via the 'title' parameter in upsert_table, not as a column

    # Read all rows
    logger.info("Reading CSV rows...")
    rows = read_csv_rows(str(csv_path))
    logger.info(f"Read {len(rows)} rows")

    if len(rows) == 0:
        logger.warning("No rows found in CSV file")
        sys.exit(0)

    # Note: Title column is not added to rows as it causes API errors
    # The table title is set via the 'title' parameter in upsert_table

    # Initialize Dust client
    config = {
        "api_key": api_key,
        "workspace_id": workspace_id,
        "space_id": space_id,
        "data_source_id": data_source_id,
        "base_url": base_url,
    }
    client = DustClient(config)

    # Test connection
    logger.info("Testing connection to Dust...")
    try:
        client.check_connection(data_format="tables")
        logger.info("Connection successful")
    except Exception as e:
        logger.error(f"Connection failed: {e}")
        sys.exit(1)

    # Note: Dust API infers table schema from row data, so we don't pass columns
    logger.info(f"Creating/updating table '{table_name}'...")
    try:
        # Let Dust generate the table_id automatically
        response = client.upsert_table(
            name=table_name,
            title=table_name,
            description=f"Imported from CSV: {csv_path.name}",
        )
        # Extract the table_id from the response
        # Dust API returns: {"table": {"table_id": "...", ...}}
        table_id = None
        if isinstance(response, dict):
            if "table" in response and isinstance(response["table"], dict):
                table_id = response["table"].get("table_id") or response["table"].get("id")
            if not table_id:
                table_id = response.get("id") or response.get("table_id")
        
        if not table_id:
            raise RuntimeError(f"Failed to extract table_id from API response: {response}")
        
        logger.info(f"Table created/updated successfully with ID: {table_id}")
    except Exception as e:
        logger.error(f"Failed to create table: {e}")
        sys.exit(1)

    # Upsert rows in batches
    batch_size = args.batch_size
    total_batches = (len(rows) + batch_size - 1) // batch_size

    logger.info(f"Uploading {len(rows)} rows in {total_batches} batch(es)...")

    for i in range(0, len(rows), batch_size):
        batch = rows[i:i + batch_size]
        batch_num = (i // batch_size) + 1

        logger.info(f"Uploading batch {batch_num}/{total_batches} ({len(batch)} rows)...")

        try:
            client.upsert_rows(table_id, batch)
            logger.info(f"Batch {batch_num} uploaded successfully")
        except Exception as e:
            logger.error(f"Failed to upload batch {batch_num}: {e}")
            sys.exit(1)

    logger.info(f"Successfully imported {len(rows)} rows into table '{table_id}'")


if __name__ == "__main__":
    main()
