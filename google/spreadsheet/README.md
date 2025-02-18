# Dust Agent for Google Sheets

This Google Apps Script allows you to easily integrate Dust AI agents with Google Sheets, enabling you to process cell content through Dust agents directly from your spreadsheet.

## Installation

1. Copy the content of dust-app-script.js
2. Open your Google Spreadsheet
3. Go to `Extensions > Apps Script`
4. Delete any existing code in the editor
5. Copy and paste the entire script content into the editor
6. Save the project (Ctrl/Cmd + S)
7. Refresh your Google Spreadsheet

## Setup

Before using the script, you need to configure your Dust credentials:

1. Click on the new "Dust" menu item in your spreadsheet
2. Select "Setup"
3. Enter your:
   - Dust Workspace ID (found in your workspace URL: dust.tt/w/[workspace-id])
   - Dust API Key (generate one at dust.tt/settings/developers)

## Usage

### Basic Operation

1. Select the cells you want to process
2. Click `Dust > Call an Assistant`
3. In the sidebar that appears:
   - Select your Dust agent
   - Verify the input cell range (or use the "Use Selection" button)
   - Specify the target column where results should appear (e.g., "B")
   - Optionally add specific instructions
4. Click "Process"

### Features

- **Batch Processing**: Process multiple cells at once
- **Custom Instructions**: Add specific instructions for the agent
- **Progress Tracking**: Monitor processing progress in real-time
- **Error Handling**: Clear error messages if something goes wrong
- **Result Links**: Each processed cell includes a note with a link to view the full conversation on Dust

## Requirements

- A Dust.tt account with API access
- Google Sheets access
- Script authorization (you'll be prompted on first use)

## Permissions

The script requires the following permissions:

- Read/write access to your spreadsheet
- Internet access (to communicate with Dust API)
- Script properties storage (to save your credentials)

## Limitations

- API rate limits apply based on your Dust plan
- Processing time increases with the number of cells
- Maximum cell content length is determined by Dust API limits

## Troubleshooting

If you encounter issues:

1. Verify your credentials in the Setup
2. Check if you have selected valid input cells
3. Ensure the target column is a valid letter (A-Z)
4. Check the browser console for detailed error messages

## Security Note

Your Dust credentials are stored securely in Google's Script Properties and are only accessible within your spreadsheet.

## Contributing

Feel free to submit issues and enhancement requests through the GitHub repository.
