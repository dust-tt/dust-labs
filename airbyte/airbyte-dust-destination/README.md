# Airbyte Destination Connector for Dust

[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

An [Airbyte](https://airbyte.com/) destination connector that syncs data to [Dust](https://dust.tt/) data sources. Supports both **document-based** and **table-based** data formats, allowing you to choose the best storage model for your use case.

## Features

- ✅ **Dual Format Support**: Choose between documents (unstructured) or tables (structured) data storage
- ✅ **Automatic Schema Inference**: Tables mode automatically infers column types from your data
- ✅ **Batch Processing**: Tables mode batches rows for efficient API usage
- ✅ **Incremental Syncs**: Supports Airbyte's incremental sync modes
- ✅ **Robust Error Handling**: Automatic retries with exponential backoff
- ✅ **Primary Key Support**: Uses Airbyte primary keys for document/row identification

## Installation

### Prerequisites

- Python 3.9 or higher
- An Airbyte instance (OSS or Cloud)
- A Dust workspace with API access

### Install from Source

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Development Dependencies

```bash
pip install pytest pytest-mock requests-mock ruff mypy
```

## Configuration

### Required Fields

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `api_key` | string | Yes | Dust API Bearer token (found in Dust workspace developer settings) |
| `workspace_id` | string | Yes | Dust workspace ID (wId in API URLs) |
| `space_id` | string | Yes | Dust space ID where the data source lives |
| `data_source_id` | string | Yes | Dust data source ID to write into |

### Optional Fields

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `base_url` | string | `https://dust.tt` | Dust API base URL. Use `https://eu.dust.tt` for Europe |
| `data_format` | string | `documents` | Data format: `"documents"` or `"tables"` |
| `table_id_prefix` | string | `airbyte_` | Prefix for table names (only used in tables mode) |

### Configuration Examples

#### Documents Mode (Default)

Best for: Unstructured data, text-heavy records, semantic search use cases

```json
{
  "api_key": "sk-...",
  "workspace_id": "your-workspace-id",
  "space_id": "your-space-id",
  "data_source_id": "your-datasource-id",
  "base_url": "https://dust.tt"
}
```

#### Tables Mode

Best for: Structured data, database tables, CSV-like data, analytics

```json
{
  "api_key": "sk-...",
  "workspace_id": "your-workspace-id",
  "space_id": "your-space-id",
  "data_source_id": "your-datasource-id",
  "base_url": "https://dust.tt",
  "data_format": "tables",
  "table_id_prefix": "airbyte_"
}
```

## Usage

### Running Locally

```bash
# Print connector specification
python main.py spec

# Test connection
python main.py check --config secrets/config.json

# Run a sync (pipe Airbyte messages from stdin)
cat messages.jsonl | python main.py write --config secrets/config.json --catalog catalog.json
```

### CSV Import Tool

A standalone CLI script is available to import CSV files directly into Dust tables:

```bash
# Install additional dependency
pip install python-dotenv

# Set up configuration (copy .env.example to .env and fill in values)
cp .env.example .env

# Import a CSV file
python scripts/csv_to_dust.py data.csv

# With custom table ID and name
python scripts/csv_to_dust.py data.csv --table-id my_table --table-name "My Table"
```

See [scripts/README.md](scripts/README.md) for detailed documentation.

### Example Config File (`secrets/config.json`)

```json
{
  "api_key": "sk-your-api-key-here",
  "workspace_id": "w1234567890",
  "space_id": "s0987654321",
  "data_source_id": "ds_abcdef123456",
  "base_url": "https://dust.tt",
  "data_format": "documents"
}
```

## Data Formats

### Documents Mode

Each Airbyte record becomes a Dust document:

- **Document ID**: Generated from stream name and primary key (or hash if no primary key)
- **Content**: JSON-serialized record data
- **Tags**: Automatically tagged with `airbyte:stream:{stream_name}`
- **Use Case**: Best for unstructured content, semantic search, text-heavy data

**Example**: A record `{"id": 1, "name": "Alice", "bio": "..."}` becomes a document with ID `users-1`.

### Tables Mode

Each Airbyte stream becomes a Dust table, with records as rows:

- **Table Name**: `{table_id_prefix}{stream_name}` (e.g., `airbyte_users`)
- **Schema**: Automatically inferred from record structure
- **Batching**: Rows are batched (500 per request) for efficiency
- **Use Case**: Best for structured relational data, analytics, database-like queries

**Example**: Stream `users` with records `[{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}]` creates table `airbyte_users` with columns `id` (number) and `name` (string).

#### Schema Inference

Tables mode automatically infers column types:

- `string` - Text values
- `number` - Integers and floats
- `boolean` - True/false values
- `json` - Nested objects and arrays (serialized as JSON strings)

Nested objects and arrays are automatically flattened to JSON strings for storage.

## Architecture

### File Structure

```
.
├── main.py                          # Entry point
├── destination_dust/
│   ├── __init__.py
│   ├── destination.py              # Core connector logic
│   ├── client.py                    # HTTP client with retry logic
│   └── spec.json                    # Connector configuration schema
├── unit_tests/
│   └── test_destination.py         # Test suite
├── Dockerfile                       # Docker build configuration
├── requirements.txt                 # Python dependencies
└── README.md                        # This file
```

### Key Components

- **`DestinationDust`**: Main connector class implementing Airbyte's Destination interface
- **`DustClient`**: HTTP client handling API communication with automatic retries
- **Schema Inference**: Automatically builds table schemas from record structure
- **Batch Processing**: Efficiently batches table row upserts

## Development

### Setup Development Environment

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install pytest pytest-mock requests-mock ruff mypy
```

### Running Tests

```bash
# Run all tests
pytest unit_tests/ -v

# Run specific test file
pytest unit_tests/test_destination.py -v

# Run with coverage
pytest unit_tests/ --cov=destination_dust --cov-report=html
```

### Code Quality

```bash
# Lint with ruff
ruff check destination_dust/

# Type check with mypy
mypy destination_dust/
```

### Building Docker Image

```bash
docker build -t airbyte/destination-dust:dev .
```

## API Details

### Documents API

- **Base URL**: `https://dust.tt` (US) or `https://eu.dust.tt` (EU)
- **Upsert Endpoint**: `POST /api/v1/w/{wId}/spaces/{spaceId}/data-sources/{dsId}/documents/{documentId}`
- **List Endpoint**: `GET /api/v1/w/{wId}/spaces/{spaceId}/data-sources/{dsId}/documents`

**Note**: Documents API uses hyphens (`data-sources`) in URL paths.

### Tables API

- **Upsert Table**: `POST /api/v1/w/{wId}/spaces/{spaceId}/data_sources/{dsId}/tables`
- **Upsert Rows**: `POST /api/v1/w/{wId}/spaces/{spaceId}/data_sources/{dsId}/tables/{tId}/rows`

**Note**: Tables API uses underscores (`data_sources`) in URL paths.

### Authentication

All requests use Bearer token authentication:

```
Authorization: Bearer sk-your-api-key
```

## Error Handling

- **Retries**: 3 attempts with exponential backoff (1s, 2s, 4s)
- **Retryable Errors**: HTTP 429 (rate limit) and 5xx server errors
- **Failures**: Failed upserts raise `RuntimeError`, causing sync to fail
- **State Management**: Airbyte resumes from last checkpointed STATE on retry
- **No Silent Failures**: All errors are surfaced - no records are silently dropped

## Document ID Strategy

### With Primary Key

If the stream has a primary key, document IDs are: `{stream_name}-{pk_value1}-{pk_value2}`

**Example**: Stream `users` with primary key `["id"]` → Document ID `users-42`

### Without Primary Key

If no primary key exists, document IDs use a hash: `{stream_name}-{sha256_hash_first_16_chars}`

**Example**: Stream `events` → Document ID `events-a1b2c3d4e5f6g7h8`

### Sanitization

All IDs are sanitized to `[a-zA-Z0-9_-]` for URL safety.

## Limitations

- **Tables Mode**: Nested objects and arrays are stored as JSON strings (not relational)
- **Schema Evolution**: Tables mode infers schema from first records seen (may need manual updates for schema changes)
- **Batch Size**: Table row batches are fixed at 500 rows (not configurable)
- **Mixed Formats**: Cannot mix documents and tables formats in a single sync

## Troubleshooting

### Connection Check Fails

1. Verify your API key is correct and has proper permissions
2. Check that workspace_id, space_id, and data_source_id are correct
3. Ensure the data source exists in the specified space
4. Verify base_url matches your region (US vs EU)

### Table Creation Fails

1. Ensure `data_format` is set to `"tables"` in configuration
2. Check that table names are valid (sanitized automatically)
3. Verify API key has write permissions for tables

### Rate Limiting

If you encounter rate limits:
- Reduce sync frequency
- Consider using tables mode (batches are more efficient)
- Contact Dust support to increase rate limits

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass (`pytest unit_tests/`)
6. Run linting (`ruff check`)
7. Commit your changes (`git commit -m 'Add amazing feature'`)
8. Push to the branch (`git push origin feature/amazing-feature`)
9. Open a Pull Request

## License

MIT License - see LICENSE file for details

## Support

- **Documentation**: [Dust API Docs](https://docs.dust.tt/reference/developer-platform-overview)
- **Issues**: Open an issue on GitHub
- **Dust Support**: [Dust Community](https://dust-community.tightknit.community/)

## Changelog

### v0.1.0 (Current)
- Initial release with documents mode support
- Added tables mode support
- Automatic schema inference
- Batch row processing for tables
- Comprehensive test coverage
