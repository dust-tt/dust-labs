# Front to Dust

Import Front conversations into Dust as searchable documents. Each conversation is stored as a single document containing all messages.

## Features

- **Conversation-level storage**: Each Front conversation stored as a single Dust document
- **Size optimization**: Total document size limited to 2MB
- **Hierarchical processing**: Fetches inboxes → conversations → messages
- **Incremental processing**: Handles large datasets with pagination
- **Error handling**: Stops immediately on any error to prevent partial imports
- **Rate limit monitoring**: Real-time monitoring of Front API rate limits
- **Automatic ID conversion**: Converts Front URL inbox IDs to API format automatically

## Prerequisites

### Front API Token

You must be a Front company admin to create an API token.

1. Log into [Front](https://app.frontapp.com)
2. Go to **Settings** → **Company** → **Developers** -> **API Tokens**
3. Click **"Create API token"**
4. Select scope, which determines the permissions of the API token
5. Click **"Create"**

### Dust Configuration
1. **Workspace ID**: Found in your Dust workspace URL: `https://dust.tt/w/{WORKSPACE_ID}`
2. **Space ID**: Found in your Dust space URL: `https://dust.tt/w/{WORKSPACE_ID}/spaces/{SPACE_ID}`
3. **Datasource ID**: Your Dust datasource identifier. This can be found by clicking `...` on your folder in the Dust console then `Use from API`.

### Front Inbox ID
The script converts FRONT URL inbox IDs to the required API format automatically: `https://app.frontapp.com/inboxes/teams/folders/{INBOX_ID}`

## Environment Variables

Create a `.env` file using `example.env`.

## Installation

```bash
npm install
```

## Usage

```bash
npm run sync
```

## Filtering Options

### Inbox Filtering
The `FRONT_FILTER` environment variable supports inbox filtering only. This allows you to import conversations from a specific Front inbox.

```env
FRONT_FILTER=inbox=inb_123  # Conversations from specific inbox ID (API format)
FRONT_FILTER=inbox=30231326 # Conversations from specific inbox (URL ID - auto-converted)
```

**Note**: If no `FRONT_FILTER` is specified, the script will process all available inboxes.

### Time Filtering
Enable time-based filtering to only process recent conversations and reduce processing time:

```env
DAYS_BACK=30               # Process conversations from last 30 days (set to 0 to disable)
```

**How it works:**
- Sorts conversations by `updated_at` (most recent first)
- Sorts messages and comments by `created_at` (most recent first)  
- Filters out conversations updated before the cutoff date

## How it Works

1. **Fetch Inboxes**: Retrieves all available Front inboxes
2. **Process Conversations**: For each inbox, fetches conversations (no time filtering applied)
3. **Fetch Messages**: For each conversation, retrieves all messages
4. **Create Documents**: Combines conversation metadata and messages into a single document
5. **Upload to Dust**: Stores each conversation as a document with appropriate tags

## Document Structure

Each conversation document contains:

- **Conversation Summary**: Subject, status, assignee, timestamps, message count
- **Messages Section**: All messages with sender, timestamp, direction, and content
- **Size Limits**: Total document size limited to 2MB
- **Tags**: Source, type, conversation ID, status, message count, timestamps, assignee, subject

## Rate Limit Monitoring

The script provides real-time monitoring of Front API rate limits:

- **Live tracking**: Shows remaining requests and reset time
- **Automatic warnings**: Warns when approaching rate limits
- **Configurable limits**: Default 50 requests per minute (configurable via `FRONT_RATE_LIMIT`)

## Error Handling

The script is designed to fail fast:
- Stops immediately on any API error
- Exits with code 1 on failure
- Provides detailed error messages and response data
- No partial imports to avoid data inconsistency 