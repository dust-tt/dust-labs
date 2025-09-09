# Dust Excel Add-in

Use AI Agents directly in your Excel spreadsheet with Dust.

## Features

- Call Dust AI agents directly from Excel cells
- Process single cells or multiple rows/columns at once
- Batch processing with real-time progress tracking
- Support for both US and EU Dust regions
- Conversation URLs included with responses for full context

## Installation

### Method 1: Sideload in Excel Desktop (Windows/Mac)

1. **Download the manifest file**: 
   - Download `manifest.xml` from this repository

2. **Open Excel** and create or open any workbook

3. **Load the add-in**:
   - Go to **Insert** tab → **Office Add-ins** (or **My Add-ins** on Mac)
   - Click **Upload My Add-in** (you may need to click "Manage My Add-ins" first)
   - Browse and select the `manifest.xml` file
   - Click **Upload**

4. **Access the add-in**:
   - Look for the **Dust** group in the **Home** tab
   - Click **Call an Agent** to open the task pane

### Method 2: Sideload in Excel Online

1. **Download the manifest file**: 
   - Download `manifest.xml` from this repository

2. **Open Excel Online** in your browser

3. **Load the add-in**:
   - Go to **Insert** tab → **Office Add-ins**
   - Select **Upload My Add-in**
   - Browse and select the `manifest.xml` file
   - Click **Upload**

4. **Access the add-in**:
   - The Dust add-in will appear in the **Home** tab
   - Click **Call an Agent** to open the task pane

### Method 3: Network Share (For Organizations)

For deployment across multiple users in an organization:

1. **Place the manifest on a network share**:
   - Copy `manifest.xml` to a shared network folder
   - Ensure all users have read access to this location

2. **Configure trusted catalogs** in Excel:
   - Go to **File** → **Options** → **Trust Center** → **Trust Center Settings**
   - Select **Trusted Add-in Catalogs**
   - Add the network path to the catalog
   - Check "Show in Menu"
   - Click **OK**

3. **Users can then install**:
   - Go to **Insert** → **Office Add-ins** → **SHARED FOLDER** tab
   - Select the Dust add-in and click **Add**

## First-Time Setup

1. **Open the task pane**: Click **Call an Agent** in the Dust ribbon

2. **Configure your Dust credentials**:
   - Enter your **Workspace ID** (found in Dust workspace settings)
   - Enter your **API Key** (generate from Dust workspace settings)
   - Optionally specify **Region** ('eu' for Europe, leave blank for US)
   - Click **Save Credentials**

3. **Verify connection**: The add-in will validate your credentials and load available agents

## Usage Guide

### Basic Workflow

1. **Select an agent** from the dropdown list
2. **Choose input range**:
   - Type a range (e.g., `A1:A10`) 
   - Or select cells and click "Use selection"
3. **Set output column** (e.g., `B` to write results in column B)
4. **Add instructions** (optional) to customize agent behavior
5. Click **Run** to process

### Working with Data

#### Single Column Input
- Select a range like `A1:A10`
- Each cell will be processed individually
- Results appear in the specified output column

#### Multiple Column Input
- Select a range like `A1:C10`
- Specify which row contains headers (default: row 1)
- Each row is processed with column context
- Format sent to agent: "Header1: value1, Header2: value2"

### Understanding Results

- **Blue background**: Successfully processed cells
- **Red background**: Cells with errors
- **Conversation URLs**: Each result includes a link to view the full Dust conversation

### Managing Credentials

- Click **Update Credentials** at the bottom of the task pane to change settings
- Use **Remove Credentials** to completely clear stored credentials

## Troubleshooting

### Common Issues

**"Please configure Dust credentials first"**
- Ensure you've entered both Workspace ID and API Key
- Verify credentials are correct by checking them in your Dust workspace

**"Invalid credentials"**
- Double-check your Workspace ID and API Key
- Ensure your API key has the necessary permissions
- Verify the region setting (US vs EU)

**Add-in not appearing in ribbon**
- Try reloading Excel
- Remove and re-upload the manifest file
- Clear Excel's add-in cache (File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → Clear)

**Processing errors**
- Check that selected cells contain data
- Verify the agent has access to necessary resources
- Review the error message for specific issues

### Getting Your Dust Credentials

1. **Workspace ID**:
   - Log into Dust
   - Go to Settings → Workspace
   - Copy the Workspace ID

2. **API Key**:
   - Go to Settings → API Keys
   - Click "Create new key"
   - Copy the generated key (save it securely!)

## Data Privacy & Security

- Credentials are stored locally in your browser/Excel client
- Data is sent directly to Dust's API over HTTPS
- No data is stored by the add-in beyond your local credentials
- Each user must configure their own credentials

## Requirements

- **Excel 2016 or later** (Windows/Mac) or **Excel Online**
- Active **Dust workspace** with API access
- **Internet connection** for API calls

## Support

- Dust Documentation: https://docs.dust.tt
- Excel Add-ins Documentation: https://docs.microsoft.com/office/dev/add-ins/
- Report Issues: Create an issue in this repository

## License

MIT