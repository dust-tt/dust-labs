# Dust Excel Add-in

Use AI Agents in your Excel spreadsheet - A port of the Dust Google Sheets integration to Excel.

## Features

- Call Dust AI agents directly from Excel
- Process single cells or multiple rows/columns
- Support for batch processing with progress tracking  
- Flexible input/output range selection
- Additional instructions for customizing agent behavior
- Support for both US and EU Dust regions

## Setup Instructions

### Prerequisites

1. **Node.js**: Install Node.js (version 14 or higher)
2. **Excel**: Microsoft Excel 2016 or later (or Excel Online)
3. **Dust Account**: You'll need a Dust workspace ID and API key

### Installation

1. **Clone or download this repository**

2. **Install dependencies**:
   ```bash
   cd dust-excel-addin
   npm install
   ```

3. **Generate SSL certificates** (required for local development):
   ```bash
   npm run generate-cert
   ```
   This will create self-signed certificates for localhost development.

4. **Start the development server**:
   ```bash
   npm start
   ```
   The server will run on https://localhost:3000

### Loading the Add-in in Excel

#### Option 1: Excel Desktop (Windows/Mac)

1. Open Excel
2. Go to **Insert** > **Office Add-ins** (or **Insert** > **My Add-ins** on Mac)
3. Click **Upload My Add-in** (or **Manage My Add-ins** > **Upload My Add-in**)
4. Browse and select the `manifest.xml` file from this project
5. Click **Upload**

#### Option 2: Excel Online

1. Open Excel Online
2. Go to **Insert** > **Office Add-ins**
3. Click **Upload My Add-in**
4. Browse and select the `manifest.xml` file
5. Click **Upload**

### First-time Setup

1. After loading the add-in, click the **Dust** button in the Home tab
2. Click **Setup Dust Credentials**
3. Enter your:
   - **Workspace ID**: Found in your Dust workspace settings
   - **API Key**: Generate one from your Dust workspace settings
   - **Region** (optional): Leave empty for US, or enter 'eu' for European region

## Usage

### Basic Usage

1. **Open the Dust task pane**: Click "Call an Agent" in the Dust menu
2. **Select an agent**: Choose from your available Dust agents
3. **Select input range**: 
   - Type a range (e.g., A1:A10) or
   - Select cells and click "Use selection"
4. **Choose output column**: Where results will be written
5. **Add instructions** (optional): Provide additional context for the agent
6. Click **Run**

### Processing Multiple Columns

When selecting multiple columns:
- The add-in will treat the first row as headers (configurable)
- Each row will be processed with column headers as context
- Format: "Header1: value1, Header2: value2, ..."

### Tips

- **Batch Processing**: The add-in processes up to 10 rows simultaneously for better performance
- **Progress Tracking**: Watch the real-time progress counter during processing
- **Error Handling**: Cells that fail will show error messages with red background
- **Results**: Successfully processed cells have a light blue background

## Differences from Google Sheets Version

This Excel port maintains feature parity with the Google Sheets version with these adaptations:

1. **Storage**: Uses browser localStorage instead of Google's PropertiesService
2. **Range Selection**: Uses Excel's Office.js API instead of Google Apps Script
3. **Deployment**: Requires manifest.xml and local/web hosting instead of Google's built-in deployment
4. **Authentication**: Credentials stored locally per browser/device

## Development

### Project Structure

```
dust-excel-addin/
├── manifest.xml          # Excel Add-in manifest
├── package.json          # Node.js dependencies
├── src/
│   ├── taskpane.html    # Main UI
│   ├── taskpane.js      # Main logic
│   ├── taskpane.css     # Styles
│   ├── commands.html    # Ribbon commands handler
│   └── commands.js      # Ribbon commands logic
└── README.md
```

### Key Technologies

- **Office.js**: Microsoft's JavaScript API for Office Add-ins
- **Excel JavaScript API**: For Excel-specific operations
- **Dust API**: For AI agent interactions
- **Select2**: Enhanced select dropdown
- **localStorage**: For credential storage

### Testing

To validate your manifest:
```bash
npm run validate
```

### Deployment

For production deployment:

1. Host the add-in files on a web server with HTTPS
2. Update the manifest.xml `SourceLocation` URLs to point to your hosted files
3. Distribute the manifest.xml file or publish to Microsoft AppSource

## Troubleshooting

### Common Issues

1. **"Please configure Dust credentials first"**
   - Click "Setup Dust Credentials" and enter your workspace ID and API key

2. **SSL Certificate Errors**
   - Run `npm run generate-cert` to create certificates
   - Accept the self-signed certificate in your browser

3. **Add-in not loading**
   - Ensure the development server is running (`npm start`)
   - Check that Excel can access https://localhost:3000
   - Try clearing Excel's add-in cache

4. **API Errors**
   - Verify your Dust API key is valid
   - Check your workspace ID is correct
   - Ensure you have the right region setting (US/EU)

## Support

For issues or questions:
- Dust documentation: https://dust.tt/docs
- Excel Add-ins documentation: https://docs.microsoft.com/office/dev/add-ins/

## License

MIT