# Dust PowerPoint Add-in

This add-in allows you to use Dust AI Agents directly in your PowerPoint presentations.

## Installation Instructions

### For PowerPoint Desktop (Windows/Mac)

1. Download the `manifest.xml` file from this folder
2. Open PowerPoint
3. Go to **Insert** → **My Add-ins** (or **Office Add-ins**)
4. Click **Upload My Add-in** (you may need to click "More Add-ins" first)
5. Browse and select the `manifest.xml` file
6. Click **Upload**
7. The Dust add-in will appear in the **Home** tab

### For PowerPoint Online (Web)

1. Download the `manifest.xml` file from this folder
2. Open PowerPoint in your browser
3. Go to **Insert** → **Office Add-ins**
4. Select **Upload My Add-in** in the top right
5. Choose **Browse** and select the `manifest.xml` file
6. Click **Upload**
7. The Dust add-in will appear in the **Home** tab

## First Time Setup

1. Click **Open Dust** in the Home tab to open the Dust panel
2. Enter your Dust credentials:
   - **Workspace ID**: Your Dust workspace identifier
   - **API Key**: Your Dust API key (get it from dust.tt → Settings → API Keys)
   - **Region**: Select your region (US or EU)
3. Click **Save Credentials**

## Using the Add-in

1. Open the Dust panel from the Home tab
2. Choose an agent from the dropdown
3. Select the scope for processing:
   - **Entire Presentation**: Process all slides
   - **Current Slide**: Process only the active slide
   - **Selected Text/Shape**: Process only selected content
4. Add any additional instructions (optional)
5. Click **Run Agent**
6. Review the results
7. Click **Apply to Presentation** to update your content

## Features

- Process entire presentations with AI agents
- Work on individual slides
- Transform selected text or shapes
- Add custom instructions for agents
- Preview results before applying
- Works with any Dust agent in your workspace

## Use Cases

- **Content Enhancement**: Improve slide text and bullet points
- **Translation**: Translate presentations to different languages
- **Summarization**: Create executive summaries from detailed slides
- **Formatting**: Standardize content across slides
- **Fact-Checking**: Verify information in presentations
- **Content Generation**: Generate new slide content based on prompts

## Troubleshooting

### Add-in doesn't appear
- Make sure you're looking in the **Home** tab
- Try reloading PowerPoint
- Check that the manifest.xml file was uploaded correctly

### Can't connect to Dust
- Verify your Workspace ID and API Key are correct
- Check your internet connection
- Ensure you've selected the correct region (US or EU)

### Agent list is empty
- Confirm you have agents configured in your Dust workspace
- Check that your API key has the necessary permissions

### Can't select content
- Ensure you have text or shapes selected when using "Selected Text/Shape" option
- For slides, make sure you're on the slide you want to process

### Results not applying
- Check that you have appropriate permissions to edit the presentation
- Try applying to a smaller selection first
- Ensure the content type is compatible (text/shapes)

## Support

For issues or questions, visit: https://dust.tt/support