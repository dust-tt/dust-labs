# Modjo Transcripts to Dust Datasource

This script imports Modjo call transcripts into a Dust datasource. It fetches transcripts from the Modjo API and upserts them into a specified Dust datasource.

## Installation

1. Clone the repository:

   ```
   git clone git@github.com:dust-tt/dust-labs.git
   cd dust-labs/modjo
   ```

2. Install dependencies:

   ```
   npm install
   ```

3. Create a `.env` file in the root directory with the following content:
   ```
   MODJO_BASE_URL=https://api.modjo.ai
   MODJO_API_KEY=your_modjo_api_key
   DUST_API_KEY=your_dust_api_key
   DUST_WORKSPACE_ID=your_dust_workspace_id
   DUST_VAULT_ID=your_dust_vault_id
   DUST_DATASOURCE_ID=your_dust_datasource_id

   # TRANSCRIPTS_SINCE=YYYY-MM-DD # or "null" if you want to fetch all transcripts
   # INCLUDE_CONTACT_DETAILS=true # or "false" to skip contact details
   # INCLUDE_RECORDING_URL=true # or "false" if you don't want the recording URL to appear in Dust
   ```
   Replace the placeholder values with your actual API keys and IDs.

## Usage

To run the script:

```
npm run transcripts
```

This command executes the `modjo-transcripts-to-dust.ts` file using `ts-node`.

## Configuration

- `TRANSCRIPTS_SINCE`: You can set this environment variable to a date string (e.g., "2024-01-01") to fetch transcripts from that date onwards. Set it to `null` to fetch all transcripts.
- `INCLUDE_CONTACT_DETAILS`: Set this environment variable to `false` if you don't want to ingest contact details (ie: email and phone number) in Dust.
- `INCLUDE_RECORDING_URL`: Set this environment variable to `false` if you don't want to ingest the recording URL in Dust.

## What the Script Does

1. **Environment Setup**: The script uses `dotenv` to load environment variables from the `.env` file.

2. **API Clients**: It sets up API clients for both Modjo and Dust using Axios.

3. **Rate Limiting**: A rate limiter is implemented using Bottleneck to prevent overwhelming the Dust API.

4. **Fetching Transcripts**: The script fetches transcripts from Modjo in batches, handling pagination.

5. **Processing Transcripts**: Each transcript is processed and formatted into a structured text content.

6. **Upserting to Dust**: The formatted transcripts are upserted to the specified Dust datasource.

## Main Functions

- `getModjoTranscripts()`: Fetches transcripts from Modjo API.
- `upsertToDustDatasource(transcript)`: Formats and upserts a single transcript to Dust.
- `formatTime(seconds)`: Helper function to format time in MM:SS format.
- `main()`: The main execution function that orchestrates the entire process.

## Transcript Format in Dust

Each transcript is formatted as follows in the Dust datasource:

1. Call metadata (ID, title, date, duration, etc.)
2. Recording URL and highlights (if available)
3. List of speakers with their details
4. Full transcript with timestamps, speaker names, and topics

## Error Handling

The script includes error handling for API requests and logs errors to the console.

## Dependencies

- axios: For making HTTP requests
- dotenv: For loading environment variables
- bottleneck: For rate limiting API calls
- typescript and ts-node: For TypeScript support

## Building

To compile the TypeScript code to JavaScript:

```
npm run build
```

This will create a `dist` directory with the compiled JavaScript files.

## Notes

- Ensure you have the necessary permissions and API keys for both Modjo and Dust.
- The script assumes a specific structure for Modjo transcripts. If the Modjo API changes, you may need to update the script accordingly.
- Be mindful of API rate limits and adjust the Bottleneck configuration if necessary.

## License

This project is licensed under the ISC License.
