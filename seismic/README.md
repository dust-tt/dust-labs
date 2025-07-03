# Seismic to Dust Sync

Synchronizes content from Seismic to Dust data sources with intelligent content processing for text documents, PDFs, and other file formats.

## Quick Start

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Configure environment:**
   ```bash
   cp example.env .env
   # Edit .env with your credentials
   ```

3. **Run sync:**
   ```bash
   npm start
   ```

## Configuration

### Required Environment Variables

#### Seismic API Setup
- `SEISMIC_CLIENT_ID` - Your Seismic OAuth2 client ID
- `SEISMIC_CLIENT_SECRET` - Your Seismic OAuth2 client secret  
- `SEISMIC_TENANT` - Your Seismic tenant ID (e.g., "acme" for acme.seismic.com)
- `SEISMIC_DELEGATION_USER_ID` - User ID for delegation authentication

#### Dust API Setup  
- `DUST_API_KEY` - Your Dust API key
- `DUST_WORKSPACE_ID` - Target Dust workspace ID
- `DUST_SPACE_ID` - Target Dust space ID
- `DUST_DATASOURCE_ID` - Target Dust data source ID

### Optional Filters
- `CONTENT_TYPE_FILTER` - Filter by content format (e.g., "PDF", "TXT")
- `DAYS_BACK_FILTER` - Only sync content created in the last X days

## Setting Up Seismic Authentication

### Creating Seismic OAuth2 App

1. **Access Seismic Developer Portal:**
   - Contact your Seismic administrator for access
   - Or visit [developer.seismic.com](https://developer.seismic.com)

2. **Create New App:**
   - Choose **Client Credentials** authentication method
   - Request **Library Contents API** permissions
   - Generate client ID and secret

3. **Enable User Delegation:**
   - In Seismic: Settings > System Settings > Manage Apps > My Apps
   - Enable your app with the toggle switch
   - Select a delegation user (used for all API calls)
   - Copy the delegation user ID

**📝 Note:** Save your `SEISMIC_CLIENT_ID`, `SEISMIC_CLIENT_SECRET`, and `SEISMIC_DELEGATION_USER_ID` for the .env file.

## Finding Dust IDs

### Workspace ID
1. Go to your Dust workspace
2. Look at the URL: `https://dust.tt/w/[WORKSPACE_ID]/...`
3. Copy the workspace identifier from the URL

### Space ID  
1. Navigate to the target space in Dust
2. Check the URL: `https://dust.tt/w/[workspace]/spaces/[SPACE_ID]`
3. Copy the space identifier

### Data Source ID
1. Go to your space in Dust
2. Click on "Data Sources" tab
3. Click `...` on your folder in the Dust console
4. Click `Use from API` to find the data source ID

### API Key
1. In Dust, go to your profile/settings
2. Navigate to "API Keys" section  
3. Generate a new API key for this integration

## How It Works

1. **Content Discovery:** Searches Seismic for content matching your filters
2. **Smart Processing:**
   - **Text files:** Direct text extraction 
   - **PDFs:** Text extraction using pdf-parse library
   - **Other formats:** Metadata-only processing
3. **Upload to Dust:** Creates structured documents with metadata and content
4. **Large Content:** Automatically splits documents >2MB into multiple parts

## Content Structure

Each synced document includes:
- **Metadata:** Content name, ID, version, repository, type, dates, properties
- **Content:** Extracted text (PDFs/text files) or metadata descriptions  
- **Source URL:** Link back to original Seismic content
- **Title:** Clear identification in Dust

## Troubleshooting

- **PDF warnings:** The script suppresses PDF parsing warnings for encrypted/complex files
- **Rate limits:** Adjust `SEISMIC_RATE_LIMIT_PER_MINUTE` and `DUST_RATE_LIMIT_PER_MINUTE` if needed
- **Large files:** Content >2MB is automatically split into numbered parts
- **Authentication errors:** Verify your Seismic app is enabled and delegation user is set

## License

This project is licensed under the ISC License. 