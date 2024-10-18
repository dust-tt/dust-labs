# Zendesk to Dust Datasource Import

This script imports Zendesk data into a Dust datasource. 
Two types of data can be fetched: 

### Tickets import
By default, the script will import tickets updated in the last 24 hours 
It fetches ticket details, comments, user information, and custom statuses from Zendesk, then formats and uploads this data to a specified Dust datasource.

### Knowledge Base import
The script can also import articles from Zendesk's Knowledge Base. (Help center)
If fetches all articles at once, then formats and uploads this data to a specified Dust datasource.

## Installation

1. Ensure you have Node.js (version 14 or higher) and npm installed on your system.

2. Clone this repository:
   ```
   git git@github.com:dust-tt/dust-labs.git
   cd zendesk
   ```

3. Install the dependencies:
   ```
   npm install
   ```

## Environment Setup

Create a `.env` file in the root directory of the project with the following variables:

```
ZENDESK_SUBDOMAIN=your_zendesk_subdomain
ZENDESK_EMAIL=your_zendesk_email
ZENDESK_API_TOKEN=your_zendesk_api_token
DUST_API_KEY=your_dust_api_key
DUST_WORKSPACE_ID=your_dust_workspace_id
DUST_VAULT_ID=your_dust_vault_id
DUST_DATASOURCE_ID=your_dust_datasource_id
```

Replace the placeholder values with your actual Zendesk and Dust credentials.

## Usage

To run the script:

*To import tickets updated in the last 24h:*
```
npm run tickets
```

*To import all articles from Zendesk's Knowledge Base:*
```
npm run articles
```

## How It Works

1. The script fetches all ticket IDs that have been updated in the last 24 hours using Zendesk's incremental export API.

2. It then retrieves detailed information for these tickets in batches of 100.

3. For each ticket, it fetches:
   - Ticket details
   - All comments on the ticket
   - User information for the ticket assignee and comment authors
   - Custom status information (if applicable)

4. The collected data is formatted into a single text document.

5. The formatted data is then upserted to the specified Dust datasource, using the ticket ID as the document ID.

6. The process is parallelized using `p-limit` to handle multiple tickets simultaneously, respecting Zendesk's rate limits.

## Configuration

- `THREADS_NUMBER`: Set to 5 by default. This determines the number of parallel operations.
- `TICKETS_UPDATED_SINCE`: Set to fetch tickets updated in the last 24 hours. Modify this value in the script if you need a different time range.

## Error Handling

The script includes error handling for rate limiting. If a rate limit is exceeded, it will wait before retrying the request.

## Building

To compile the TypeScript code to JavaScript:

```
npm run build
```

This will create a `dist` directory with the compiled JavaScript files.

## Dependencies

- axios: For making HTTP requests to Zendesk and Dust APIs
- dotenv: For loading environment variables
- p-limit: For limiting the number of concurrent operations

## Dev Dependencies

- @types/node: TypeScript definitions for Node.js
- ts-node: For running TypeScript files directly
- typescript: The TypeScript compiler

## Notes

- Ensure you have the necessary permissions in both Zendesk and Dust to perform these operations.
- Be mindful of your Zendesk API usage limits when running this script frequently or with large datasets.
- The script currently fetches tickets from the last 24 hours. Modify the `TICKETS_UPDATED_SINCE` constant if you need a different time range.

## License

This project is licensed under the ISC License.
