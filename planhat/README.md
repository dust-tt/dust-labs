# Planhat to Dust Integration

This script syncs data from Planhat to a Dust datasource, creating one document per company with all associated data (conversations, endusers, NPS, projects, and assets).

## Prerequisites

- Node.js 18 or higher
- A Planhat account with API access
- A Dust workspace with a datasource created
- API tokens for both Planhat and Dust

## Installation

```bash
cd planhat
npm install
```

## Configuration

Create a `.env` file in the planhat directory with the following variables:

```env
# Planhat API configuration
PLANHAT_API_TOKEN=your_planhat_api_token

# Dust API configuration
DUST_API_KEY=your_dust_api_key
DUST_WORKSPACE_ID=your_dust_workspace_id
DUST_DATASOURCE_ID=your_dust_datasource_id
DUST_SPACE_ID=your_dust_space_id

# Sync configuration
LOOKBACK_DAYS=7  # Number of days to look back for updated companies
THREADS_NUMBER=1  # Number of parallel threads for processing (for Dust API rate limits, better to keep this at 1)
```

## Usage

### Development

```bash
npm run dev
```

### Production

```bash
npm run build
npm start
```

## How it works

1. **Fetches updated companies**: The script queries Planhat for all companies updated within the last `LOOKBACK_DAYS` days.

2. **Enriches company data**: For each company, it fetches:

   - Conversations
   - End users
   - NPS scores
   - Projects
   - Assets

3. **Creates structured documents**: Each company becomes a Dust document with hierarchical sections:

   ```
   Company Document
   ├── Basic Information
   ├── Custom Fields
   ├── Conversations
   │   ├── Conversation 1
   │   ├── Conversation 2
   │   └── ...
   ├── End Users
   │   ├── End User 1
   │   ├── End User 2
   │   └── ...
   ├── NPS
   │   ├── NPS Response 1
   │   ├── NPS Response 2
   │   └── ...
   ├── Projects
   │   ├── Project 1
   │   ├── Project 2
   │   └── ...
   └── Assets
       ├── Asset 1
       ├── Asset 2
       └── ...
   ```

4. **Upserts to Dust**: Documents are upserted to the specified Dust datasource with:
   - Document ID based on company slug or ID
   - Tags for status, health, phase, and owner
   - Timestamp from the company's last update

## Example Output

```
Starting Planhat to Dust sync...
Looking back 7 days for updated companies
Using 4 threads for processing
Found 42 companies to process
Processing company: Acme Corp (5f4d3a2b1c9d4e0001234567)
✓ Successfully processed company: Acme Corp
Processing company: Tech Solutions Inc (5f4d3a2b1c9d4e0001234568)
✓ Successfully processed company: Tech Solutions Inc
...
✅ Sync completed successfully!
```

## Rate Limiting

The script implements rate limiting to respect API limits:

- **Planhat**: 5 requests per second (200ms between requests)
- **Dust**: Configurable based on thread count

## Error Handling

- API errors are logged but don't stop the entire sync
- Failures on individual Planhat resource fetches (e.g. 403 Forbidden on `m_asset`) are logged on a **single line** and the script continues with an empty list for that resource
- Failed company syncs are reported but other companies continue processing
- Worker thread errors are caught and logged

## Development

### Building

```bash
npm run build
```

This creates a `dist` directory with the compiled JavaScript.

### TypeScript Configuration

The project uses ES modules with TypeScript. The `tsconfig.json` is configured for Node.js 18+ compatibility.

## Troubleshooting

### Missing environment variables

Ensure all required environment variables are set in your `.env` file.

### API rate limits

If you encounter rate limit errors, reduce the `THREADS_NUMBER` or increase the rate limiter delays in the code.

### Authentication errors

Verify your API tokens are correct and have the necessary permissions:

- Planhat token needs read access to companies, conversations, endusers, NPS, projects, and assets
- Dust API key needs write access to the specified datasource
