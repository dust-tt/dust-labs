# Scripts

## csv_to_dust.py

CLI tool to import CSV files directly into Dust tables.

### Setup

1. Install dependencies:
   ```bash
   pip install python-dotenv
   ```

2. Create a `.env` file in the project root (copy from `.env.example`):
   ```bash
   cp .env.example .env
   ```

3. Edit `.env` and fill in your Dust credentials:
   ```env
   DUST_API_KEY=sk-your-api-key-here
   DUST_WORKSPACE_ID=your-workspace-id
   DUST_SPACE_ID=your-space-id
   DUST_DATA_SOURCE_ID=your-datasource-id
   DUST_BASE_URL=https://dust.tt
   ```

### Usage

```bash
# Basic usage (table ID and name derived from CSV filename)
python scripts/csv_to_dust.py data.csv

# Specify table ID and name
python scripts/csv_to_dust.py data.csv --table-id my_table --table-name "My Table"

# Custom batch size (default: 500)
python scripts/csv_to_dust.py data.csv --batch-size 1000
```

### Features

- **Automatic schema inference**: Infers column types from CSV data
- **Type conversion**: Automatically converts strings to numbers/booleans where appropriate
- **Mandatory title column**: Ensures every row has a title field
- **Batch processing**: Uploads rows in configurable batches for efficiency
- **Error handling**: Validates connection and provides clear error messages

### Example

```bash
# Import a CSV file
python scripts/csv_to_dust.py products.csv --table-id products --table-name "Products"

# Output:
# Reading CSV file: products.csv
# Table ID: products
# Table name: Products
# Inferring schema from CSV...
# Inferred 5 columns: id, name, price, category, description
# Reading CSV rows...
# Read 1000 rows
# Testing connection to Dust...
# Connection successful
# Creating/updating table 'products'...
# Table created/updated successfully
# Uploading 1000 rows in 2 batch(es)...
# Uploading batch 1/2 (500 rows)...
# Batch 1 uploaded successfully
# Uploading batch 2/2 (500 rows)...
# Batch 2 uploaded successfully
# Successfully imported 1000 rows into table 'products'
```

### CSV Format

The script expects a standard CSV file with:
- Header row with column names
- UTF-8 encoding
- Standard CSV formatting (comma-separated)

The script will:
- Infer column types (string, number, boolean, json)
- Add a mandatory `title` column if missing
- Handle empty values appropriately
- Convert types where possible
