# Dropbox to Dust Data Sync

This script is a Node.js application designed to sync Dropbox files (including Paper docs) into Dust datasources. It fetches all files from your Dropbox account, optionally filtering by extension (e.g., `.paper`, `.md`, `.docx`), exports their content, and uploads them to a specified Dust datasource while maintaining their structure and content.

## Features

- Bulk synchronization of all Dropbox files (optionally filter by extension)
- Supports Dropbox Paper docs, Markdown, Word, and other file types
- Preserves file metadata including:
  - File ID and name
  - Path and parent folder
  - Created and updated dates
- Rate limiting for both Dropbox and Dust APIs
- Concurrent processing with configurable limits

## Prerequisites

- Node.js (version 14 or higher)
- npm (Node Package Manager)
- Dropbox account with API access
- Dust account with API access

## Installation

1. Clone the repository:
   ```bash
   git clone git@github.com:dust-tt/dust-labs.git
   cd dust-labs
   cd dropbox
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Create a `.env` file.

**Main script file:** The entry point for syncing is `dropbox-to-dust.ts`. All commands and npm scripts reference this file.

## Creating Dropbox API Credentials and Required Permissions

To use this script, you need to create a Dropbox API app and generate an access token with the correct permissions.

### Steps to Create Dropbox API Credentials

1. **Go to the Dropbox App Console:**
   - Visit https://www.dropbox.com/developers/apps

2. **Create a new app:**
   - Click "Create App".
   - Choose "Scoped access".
   - Choose "Full dropbox" (or "App folder" if you want to restrict access to a specific folder).
   - Name your app and create it.

3. **Configure permissions (scopes):**
   - In your app's settings, go to the "Permissions" tab.
   - Enable the following scopes:
     - `files.metadata.read`
     - `files.content.read`
     - `files.content.write` (if you want to support future write operations)
     - `files.export` (for exporting Dropbox Paper docs)
   - Click "Submit" at the bottom to save your changes.

4. **Generate an access token:**
   - In the "Settings" tab, scroll down to "OAuth 2".
   - Click "Generate access token" (for personal use/testing) or implement the OAuth flow for production.
   - Copy the generated token and use it as your `DROPBOX_API_KEY` in the `.env` file.

### Environment Variables Explained

- `DROPBOX_API_KEY`: Your Dropbox API token
- `DUST_API_KEY`: Your Dust API key
- `DUST_WORKSPACE_ID`: Your Dust workspace identifier
- `DUST_DATASOURCE_ID`: Your Dust datasource identifier. This can be found by clicking `...` on your folder in the Dust console then `Use from API`.
- `DUST_SPACE_ID`: Your Dust space identifier
- `DROPBOX_EXTENSION_FILTER`: (Optional) File extension to filter by (e.g., `.paper`, `.md`, `.docx`). If not set, all files are synced.
- `DROPBOX_ROOT_PATH`: (Optional) Dropbox folder path to start from. Defaults to root (`""`).
- `DUST_RATE_LIMIT`: (Optional, default 120) Maximum requests per minute to Dust API.
- `DROPBOX_MAX_CONCURRENT`: (Optional) Maximum concurrent Dropbox file processing operations

**To sync only Dropbox Paper docs:**
- Set `DROPBOX_EXTENSION_FILTER=.paper` in your `.env`, or
- Run the script with `npm run sync -- --ext .paper`

## Usage

To run the script (sync all files):

```bash
npx tsx dropbox-to-dust.ts
```

Or, using npm (recommended):

```bash
npm run sync
```

To filter by extension (e.g., only Dropbox Paper docs):

```bash
npx tsx dropbox-to-dust.ts --ext .paper
```

Or, using npm:

```bash
npm run sync -- --ext .paper
```

The script will:
1. Validate environment variables
2. Connect to Dropbox and fetch all files (optionally filter by extension)
3. Export and process each file
4. Upload formatted data to your Dust datasource

## How It Works

1. The script authenticates with Dropbox using the provided API token
2. It lists all files in your Dropbox (optionally under a specific folder)
3. Optionally filters files by the specified extension
4. Exports the content of each file (using Dropbox export API for Paper docs)
5. Formats and uploads the data to Dust using unique document IDs
6. Rate limiting is applied to respect both Dropbox and Dust API constraints

## Error Handling

The script includes comprehensive error handling:
- Validates required environment variables
- Handles API rate limits and retries
- Provides detailed error logging
- Continues processing on individual file failures

## License

This project is licensed under the ISC License. 