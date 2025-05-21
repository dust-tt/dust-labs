# HubSpot to Dust Import

This script is designed to import data from HubSpot to Dust datasources. It fetches recently updated companies from HubSpot, along with their associated contacts and deals, and upserts this information into a specified Dust datasource.

## Usage example

```
➜  hubspot git:(main) ✗ npm run import

> hubspot-to-dust-import@1.0.0 import
> node --no-warnings --loader ts-node/esm hubspot-to-dust.ts

Found 22 companies with updates in the last 1 day(s).
Worker 0: Starting to process 8 companies
Worker 1: Starting to process 8 companies
Worker 2: Starting to process 6 companies
Upserted company 13519965149 to Dust datasource
Upserted company 10979757009 to Dust datasource
[...]
```

Example of account summary added to the Dust datasource: 

```
Company Summary for Evil Corp

Basic Company Details:
Company Name: Evil Corp

Key Contacts:
- Email: dr.evil@evilcorp.com
- Dr Maboul, Title: Head of AI, Email: dr.maboul@evilcorp.com
- Dr Zeuss, Email: dr.zeuss@evilcorp.com, Phone: +33 7 66 66 66 66

Deals:
- Evil Corp, Stage: decisionmakerboughtin, Amount: 42000, Close Date: 2024-07-31T10:21:55.017Z

Notes:
- 2024-09-04: Adding some notes to be more Evil.
```

## Features

- Multi-threaded processing for improved performance
- Rate limiting to respect API constraints
- Detailed error logging
- Configurable lookback period for recent updates

## Prerequisites

- Node.js (v14 or later recommended)
- npm (comes with Node.js)
- A HubSpot account with API access
- A Dust account with API access

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/your-username/hubspot-to-dust-import.git
   cd hubspot-to-dust-import
   ```

2. Install dependencies:
   ```
   npm install
   ```

3. Create a `.env` file in the root directory with the following content:
   ```
   HUBSPOT_ACCESS_TOKEN=your_hubspot_access_token
   HUBSPOT_PORTAL_ID=your_hubspot_portal_id
   DUST_API_KEY=your_dust_api_key
   DUST_WORKSPACE_ID=your_dust_workspace_id
   DUST_DATASOURCE_ID=your_dust_datasource_id
   DUST_VAULT_ID=your-dust-space-id
   ```
   Replace the placeholders with your actual credentials.

## Configuration

You can adjust the following variables in the script to customize its behavior:

- `UPDATED_SINCE_DAYS`: Number of days to look back for updates (default: 1)
- `THREADS_NUMBER`: Number of worker threads to use (default: 3)

## Usage

To run the script:

```
npm run import
```

This command will execute the TypeScript script using ts-node.

## How It Works

1. The script first fetches the IDs of companies that have been updated within the specified time frame.
2. It then divides these companies among multiple worker threads for parallel processing.
3. For each company, the script:
   - Fetches detailed company information
   - Retrieves associated contacts
   - Retrieves associated deals
   - Compiles this information into a structured format
   - Upserts the compiled data into the specified Dust datasource

## Error Handling and Logging

The script includes comprehensive error handling and logging. All errors are caught and logged, allowing the script to continue processing other companies even if an error occurs with one.

## Rate Limiting

The script uses Bottleneck to implement rate limiting for both the HubSpot and Dust APIs, ensuring that we don't exceed the allowed request rates.

## Building and Linting

To compile the TypeScript code to JavaScript:

```
npm run build
```

To run the linter:

```
npm run lint
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the ISC License.

## Disclaimer

This script is provided as-is, without any guarantees or warranty. The authors are not responsible for any damage or data loss that may occur from its use.

## Support

If you encounter any issues or have questions, ask in the Dust community support: https://community.dust.tt
