# Grain to Dust Data Sync

This script is a Node.js application designed to sync Grain recording data into Dust datasources. It fetches recordings (and their details) from Grain, enriches them with highlights, participants, and more, and uploads them to a specified Dust datasource.

## Usage Example

Example of a recording summary added to the Dust datasource:

```
Grain Recording Summary for b5185ccb-9a08-458c-9be1-db17a03fb14c: Sample Recording

Metadata:
Recording Details:
ID: b5185ccb-9a08-458c-9be1-db17a03fb14c
Title: Sample Recording
URL: https://grain.com/recordings/b5185ccb-9a08-458c-9be1-db17a03fb14c/Kz5t1kAyPtt78hcxbSOJHJzFiPpZmUIeDVFXWzP0
Start: 2021-07-29T23:13:17Z
End: 2021-07-29T23:16:18Z
Tags: sample
Thumbnail: None

People:
Owners: alice@example.com
Participants: Bob Example <bob@example.com> (external)

Highlights:
Highlight vjQRUKsWw0aFpCT3531eGbr8V0HJrMjKMEIcAUmP
Text: testing 123
Transcript: expected, that there was a mews in a lane...
Speakers: Andy Arbol
Timestamp: 3080
Duration: 15000
Created: 2021-07-29T23:16:34Z
URL: https://grain.com/highlight/vjQRUKsWw0aFpCT3531eGbr8V0HJrMjKMEIcAUmP
Thumbnail: https://media.grain.com/clips/v1/...
Tags: test
```

## Features

- Configurable recording fetching by date range (`GRAIN_START_DATE`, `GRAIN_END_DATE`)
- Enriches recordings with highlights, participants, owners, and tags
- Collects comprehensive recording metadata, including:
  - Title, URL, start/end times, tags, thumbnail
  - People involved (owners, participants)
  - Highlights (text, transcript, speakers, etc.)
- Rate limiting for both Grain and Dust APIs
- Detailed logging and error handling

## Prerequisites

- Node.js (version 14 or higher)
- npm (Node Package Manager)
- Grain account with API access (OAuth2 or Personal Access Token)
- Dust account with API access
- Create a folder in Dust and fetch its data source id (starts with dsv\_)

## Installation

1. Clone the repository:
   ```bash
   git clone git@github.com:dust-tt/dust-labs.git
   cd dust-labs
   cd grain
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Create a `.env` file with the following variables:

   ```
   # Required variables
   GRAIN_API_TOKEN=your_grain_api_token
   DUST_API_KEY=your_dust_api_key
   DUST_WORKSPACE_ID=your_workspace_id
   DUST_DATASOURCE_ID=your_datasource_id
   DUST_RATE_LIMIT=100
   GRAIN_MAX_CONCURRENT=5

   # Optional filters
   GRAIN_START_DATE=2025-04-20T00:00:00Z
   GRAIN_END_DATE=2025-04-22T23:59:59Z

   # Optional feature flags
   FETCH_HIGHLIGHTS=true
   FETCH_TRANSCRIPT=true
   FETCH_PARTICIPANTS=true
   FETCH_TAGS=true
   FETCH_INTELLIGENCE_NOTES=true
   ```

## Usage

To run the script:

```bash
npm run meetings
```

The script will:
1. Validate environment variables and configuration
2. Fetch Grain recordings (with optional date filters)
3. Fetch recording details (highlights, participants, etc.)
4. Upload formatted data to your Dust datasource

## How It Works

1. The script connects to Grain using the provided API token
2. It fetches recordings within the specified date range (if configured)
3. It fetches all details for each recording (highlights, participants, owners, tags)
4. The data is formatted into a hierarchical document structure
5. The formatted documents are upserted to the specified Dust datasource
6. Rate limiting is applied to respect both Grain and Dust API constraints

## Error Handling

The script includes comprehensive error handling:
- Validates required environment variables
- Handles API rate limits
- Provides detailed error logging
- Continues processing on individual recording failures

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the ISC License.

---

Let me know if you want to add usage for specific feature flags, or example outputs for highlights, transcripts, etc.!