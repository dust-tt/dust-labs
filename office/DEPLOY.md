# Dust Office Add-ins Deployment Guide

This project contains both Excel and PowerPoint add-ins for Dust, served from a single Vercel deployment.

## Project Structure

```
/
├── src/
│   ├── shared/
│   │   └── dust-api.js         # Shared API logic for both add-ins
│   ├── excel/
│   │   ├── taskpane.html       # Excel add-in UI
│   │   ├── taskpane.js         # Excel-specific logic
│   │   ├── taskpane.css        # Excel styles
│   │   ├── commands.html       # Excel ribbon commands
│   │   └── commands.js         # Excel command handlers
│   └── powerpoint/
│       ├── taskpane.html       # PowerPoint add-in UI
│       ├── taskpane.js         # PowerPoint-specific logic
│       ├── taskpane.css        # PowerPoint styles
│       ├── commands.html       # PowerPoint ribbon commands
│       └── commands.js         # PowerPoint command handlers
├── excel/
│   └── manifest.xml            # Excel add-in manifest
├── powerpoint/
│   └── manifest.xml            # PowerPoint add-in manifest
├── vercel.json                 # Vercel routing configuration
└── package.json                # Project dependencies
```

## URL Structure

The Vercel deployment serves both add-ins with the following URL structure:
- Excel add-in: `https://dust-office-addins.vercel.app/excel/*`
- PowerPoint add-in: `https://dust-office-addins.vercel.app/powerpoint/*`

Shared resources (like dust-api.js) are automatically routed from both paths.

## Local Development

1. Install dependencies:
   ```bash
   npm install
   ```

2. Generate SSL certificates (first time only):
   ```bash
   npm run generate-cert
   ```

3. Start the development server:
   ```bash
   npm run dev
   ```

4. The server will run on https://localhost:3000 with:
   - Excel add-in at: https://localhost:3000/excel/taskpane.html
   - PowerPoint add-in at: https://localhost:3000/powerpoint/taskpane.html

## Deployment to Vercel

1. Push your changes to the repository

2. Deploy to Vercel:
   ```bash
   vercel --prod
   ```

3. The deployment will be available at:
   - Production: https://dust-office-addins.vercel.app

## Manifest Files

### Excel Manifest
- Location: `/excel/manifest.xml`
- Host: Workbook
- Base URL: `https://dust-office-addins.vercel.app/excel/`

### PowerPoint Manifest
- Location: `/powerpoint/manifest.xml`
- Host: Presentation
- Base URL: `https://dust-office-addins.vercel.app/powerpoint/`

## Testing the Add-ins

### Excel Add-in
1. Open Excel (desktop or web)
2. Go to Insert > Add-ins > Other Add-ins > My Add-ins > Upload My Add-in
3. Upload `/excel/manifest.xml`
4. The add-in will appear in the Home tab

### PowerPoint Add-in
1. Open PowerPoint (desktop or web)
2. Go to Insert > Add-ins > Advanced > Upload
3. Upload `/powerpoint/manifest.xml`
4. The add-in will appear in the Home tab

## Validation

Validate the manifest files:
```bash
# Validate Excel manifest
npm run validate:excel

# Validate PowerPoint manifest
npm run validate:powerpoint

# Validate both
npm run validate:all
```

## Environment Variables

No environment variables are needed for deployment. The add-ins store user credentials locally in Office document settings.

## CORS Configuration

The `vercel.json` file includes CORS headers to allow the add-ins to work properly:
- Allows all origins (`*`)
- Supports GET, POST, and OPTIONS methods
- Allows necessary headers for API communication

## Troubleshooting

### Certificate Issues
If you encounter SSL certificate issues in development:
1. Delete the `certs` folder
2. Run `npm run generate-cert` again
3. Trust the new certificate in your system

### Manifest Loading Issues
- Ensure the manifest URLs match your Vercel deployment URL
- Check that all referenced resources are accessible
- Validate the manifest using `npm run validate:all`

### API Connection Issues
- Verify the Dust API key and workspace ID are correct
- Check the region setting (US or EU)
- Ensure the API endpoints are accessible from your network