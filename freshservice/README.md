# Freshservice to Dust Integration

This integration imports Freshservice tickets into your Dust workspace, making your support tickets searchable and accessible through Dust's AI assistant.

![Freshservice Integration Example](https://via.placeholder.com/800x400/0052CC/FFFFFF?text=Freshservice+Tickets+in+Dust)

## Features

- 🎫 **Comprehensive Ticket Import**: Imports all ticket details including descriptions, notes, attachments, and custom fields
- 🔄 **Incremental Sync**: Only syncs tickets updated in the last 24 hours by default
- 🏷️ **Rich Metadata**: Preserves ticket status, priority, agent assignments, and custom fields
- 🔗 **Smart Linking**: Maintains relationships between tickets and related objects
- ⚡ **Rate Limited**: Respects Freshservice API limits with automatic retry logic
- 🎯 **Time-based Filtering**: Filter tickets by when they were last updated

## Prerequisites

- Node.js (v18 or higher)
- A Freshservice account with API access
- A Dust workspace

## Installation

1. Clone this repository or download the files
2. Navigate to the freshservice directory:
   ```bash
   cd freshservice
   ```
3. Install dependencies:
   ```bash
   npm install
   ```

## Configuration

### 1. Freshservice API Setup

1. Log into your Freshservice account
2. Go to **Admin** > **API Keys**
3. Generate a new API key or use an existing one
4. Note your Freshservice domain (e.g., `yourcompany` from `yourcompany.freshservice.com`)

### 2. Dust Workspace Setup

1. Log into your Dust workspace at [dust.tt](https://dust.tt)
2. Create a new space or use an existing one for storing Freshservice tickets
3. Create a new datasource within that space:
   - Name: `freshservice-tickets` (or your preferred name)
   - Type: Managed
4. Note down your workspace ID, space ID, and datasource ID from the URLs

### 3. Environment Configuration

1. Copy the example environment file:
   ```bash
   cp example.env .env
   ```

2. Edit `.env` with your configuration:

   **For Production (standard setup):**
   ```env
   # Freshservice Configuration
   FRESHSERVICE_DOMAIN=yourcompany
   FRESHSERVICE_API_KEY=your_api_key_here
   FRESHSERVICE_FILTER=updated_since=24h

   # Dust Configuration
   DUST_API_KEY=your_dust_api_key
   DUST_WORKSPACE_ID=your_workspace_id
   DUST_SPACE_ID=your_space_id
   DUST_DATASOURCE_ID=freshservice-tickets
   ```

## Usage

### Basic Usage

Import tickets updated in the last 24 hours:

```bash
npm run sync
```

Or, using tsx directly:

```bash
npx tsx freshservice-tickets-to-dust.ts
```

### Time-based Filtering

The integration supports filtering tickets by when they were last updated using the `updated_since` parameter:

#### Time-based filters:
```env
FRESHSERVICE_FILTER=updated_since=1h    # Last hour
FRESHSERVICE_FILTER=updated_since=24h   # Last 24 hours (default)
FRESHSERVICE_FILTER=updated_since=7d    # Last 7 days
FRESHSERVICE_FILTER=updated_since=2024-01-01T00:00:00Z  # Since specific date
```

### Automation

To run this integration regularly, you can set up a cron job:

```bash
# Run every hour
0 * * * * cd /path/to/freshservice && npm run sync

# Run every 6 hours
0 */6 * * * cd /path/to/freshservice && npm run sync

# Run daily at 2 AM
0 2 * * * cd /path/to/freshservice && npm run sync
```

### What Gets Imported

For each ticket, the integration imports:

### Core Information
- **Ticket ID**: Unique identifier
- **Subject and Description**: Main ticket content
- **Status**: Open, Pending, Resolved, Closed
- **Priority**: Low, Medium, High, Urgent
- **Type and Source**: Ticket categorization
- **Created and Updated dates**: Temporal information

### People and Assignment
- **Requester Details**: Name, email, department, location
- **Agent Assignment**: Assigned agent information
- **Group Assignment**: Support group details

### Rich Content
- **Ticket Notes**: All notes and replies on the ticket, including:
  - Public and private notes
  - Incoming and outgoing communications
  - Note timestamps and authors
  - Attachments associated with notes
- **Attachments**: Links to attached files
- **Custom Fields**: All custom field values
- **Tags**: Ticket labels and categorization
- **CC Recipients**: Additional stakeholders

### Relationships
- **Associated Items**: Related configuration items
- **Escalation Information**: Escalation status and details
- **Email Configuration**: Email notification settings

## Troubleshooting

### Common Issues

#### 1. Authentication Error
```
Error: Request failed with status code 401
```
**Solution**: Check your API key and domain:
- Verify `FRESHSERVICE_API_KEY` is correct
- Ensure `FRESHSERVICE_DOMAIN` matches your Freshservice URL
- Confirm the API key has proper permissions

#### 2. Rate Limiting
```
Error: Request failed with status code 429
```
**Solution**: The integration includes automatic retry logic, but if you see persistent rate limiting:
- Reduce the frequency of runs
- Consider using more specific time filters to reduce the number of tickets processed

#### 3. No Tickets Found
```
No tickets found matching the criteria
```
**Possible causes**:
- No tickets updated in the specified time range
- Time filter too restrictive
- Check if tickets exist in Freshservice with the current filter

#### 4. Dust API Errors
```
Error uploading to Dust datasource
```
**Solution**: Verify Dust configuration:
- Check `DUST_API_KEY` has write permissions
- Confirm `DUST_WORKSPACE_ID`, `DUST_SPACE_ID`, and `DUST_DATASOURCE_ID` are correct
- Ensure the datasource exists and is accessible

### Freshservice API Limits

- **Rate Limit**: 1000 requests per hour
- **Pagination**: 30 tickets per page (configurable up to 100)
- **Timeout**: 30 seconds per request

## Data Privacy

This integration:
- Only reads ticket data (no modifications to Freshservice)
- Transfers data directly between Freshservice and Dust
- Includes all ticket data including private notes (review your privacy requirements)
- Respects Freshservice user permissions (only accessible tickets are imported)

## Support

For issues specific to this integration:
1. Check the troubleshooting section above
2. Verify your environment configuration
3. Test API connectivity to both Freshservice and Dust

For Freshservice API questions, consult the [Freshservice API documentation](https://api.freshservice.com/).
For Dust API questions, consult the [Dust API documentation](https://docs.dust.tt/).

## License

MIT License - see LICENSE file for details. 