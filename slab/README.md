# Slab to Dust Sync

Synchronizes all documents from Slab to Dust data sources with intelligent content processing and chunking.

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
   npm run sync
   ```

## Configuration

### Required Environment Variables

#### Slab API Setup

- `SLAB_API_TOKEN` - Your Slab API token (obtain from Developer settings in Slab)
- `SLAB_DOMAIN` - Your Slab domain (e.g., `app.slab.com` or `yourcompany.slab.com`, defaults to `app.slab.com`)

#### Dust API Setup

- `DUST_API_KEY` - Your Dust API key
- `DUST_WORKSPACE_ID` - Target Dust workspace ID
- `DUST_SPACE_ID` - Target Dust space ID
- `DUST_DATASOURCE_ID` - Target Dust data source ID

### Optional Configuration

- `DUST_API_BASE_URL` - Dust API base URL (defaults to `https://dust.tt/api/v1`, use `https://eu.dust.tt/api/v1` for EU region)
- `DUST_RATE_LIMIT_PER_MINUTE` - Rate limit for Dust API (default: 120)
- `SLAB_RATE_LIMIT_PER_MINUTE` - Rate limit for Slab API (default: 120)
- `SLAB_DOMAIN` - Your Slab domain (e.g., `app.slab.com` or `yourcompany.slab.com`, defaults to `app.slab.com`)

## Setting Up Slab Authentication

### Obtaining Slab API Token

1. **Access Slab Developer Settings:**

   - Log in to your Slab workspace
   - Navigate to Settings > Developer Tools
   - Ensure your team is on Business or Enterprise plan (API access required)

2. **Generate API Token:**
   - Click "Generate API Token" or "Create Token"
   - Copy the generated token
   - **Note:** Store this token securely as it won't be shown again

**📝 Note:** Save your `SLAB_API_TOKEN` for the .env file.

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

1. **Incremental Processing:** Fetches and processes posts one at a time, upserting immediately (no need to load all posts into memory)
2. **Checkpointing (Optional):** If `CHECKPOINT_FILE` is set, the script saves progress and can resume from where it left off if interrupted
3. **Content Processing:**
   - Extracts document content and metadata
   - Formats metadata with frontmatter-style headers
   - Builds topic hierarchy paths
4. **Intelligent Chunking:**
   - Automatically splits documents >1MB into multiple parts
   - Uses heading-aware splitting to preserve document hierarchy
   - Prefers breaking at section boundaries to keep content together
   - Includes 200-character overlap between chunks for continuity
5. **Upload to Dust:** Creates structured documents with hierarchical sections, metadata, content, and source URLs

## Content Structure

Each synced document includes:

- **Metadata:** Title, Post ID, URL, creation/update dates, author, topic
- **Hierarchy Context:** Section path showing where the chunk sits in the document structure
- **Section Information:** Current section heading and level, plus next section preview
- **Content:** Full document body text with preserved structure
- **Source URL:** Link back to original Slab document
- **Title:** Clear identification in Dust (with part numbers and section names for multi-part documents)

## License

This project is licensed under the ISC License.
