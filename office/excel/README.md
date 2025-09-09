# Dust Excel Add-in

This add-in allows you to use Dust AI Agents directly in your Excel spreadsheets.

## Installation Instructions

### For Excel Desktop (Windows/Mac)

1. Download the `manifest.xml` file from this folder
2. Open Excel
3. Go to **Insert** → **My Add-ins** (or **Office Add-ins**)
4. Click **Upload My Add-in** (you may need to click "More Add-ins" first)
5. Browse and select the `manifest.xml` file
6. Click **Upload**
7. The Dust add-in will appear in the **Home** tab

### For Excel Online (Web)

1. Download the `manifest.xml` file from this folder
2. Open Excel in your browser
3. Go to **Insert** → **Office Add-ins**
4. Select **Upload My Add-in** in the top right
5. Choose **Browse** and select the `manifest.xml` file
6. Click **Upload**
7. The Dust add-in will appear in the **Home** tab

## First Time Setup

1. Click **Call an Agent** in the Home tab to open the Dust panel
2. Enter your Dust credentials:
   - **Workspace ID**: Your Dust workspace identifier
   - **API Key**: Your Dust API key (get it from dust.tt → Settings → API Keys)
   - **Region**: Select your region (US or EU)
3. Click **Save Credentials**

## Using the Add-in

1. Select the cells you want to process
2. Open the Dust panel from the Home tab
3. Choose an agent from the dropdown
4. Select your data range options:
   - Include/exclude header row
   - Process by rows or columns
5. Add any additional instructions
6. Click **Run Agent**

## Features

- Process selected cells with AI agents
- Include or exclude header rows
- Process data by rows or columns
- Add custom instructions for agents
- Works with any Dust agent in your workspace

## Troubleshooting

### Add-in doesn't appear
- Make sure you're looking in the **Home** tab
- Try reloading Excel
- Check that the manifest.xml file was uploaded correctly

### Can't connect to Dust
- Verify your Workspace ID and API Key are correct
- Check your internet connection
- Ensure you've selected the correct region (US or EU)

### Agent list is empty
- Confirm you have agents configured in your Dust workspace
- Check that your API key has the necessary permissions

## Support

For issues or questions, visit: https://dust.tt/support