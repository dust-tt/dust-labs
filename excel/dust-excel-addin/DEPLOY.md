# Deploying Dust Excel Add-in to Vercel

## Prerequisites
- Vercel account (create one at https://vercel.com)
- Vercel CLI installed: `npm i -g vercel`

## Deployment Steps

### 1. Deploy to Vercel

```bash
# Login to Vercel (first time only)
vercel login

# Deploy the project
vercel

# For production deployment
vercel --prod
```

### 2. Update the Manifest

After deployment, Vercel will provide you with a URL like `https://your-project-name.vercel.app`

1. Open `manifest-prod.xml`
2. Replace all instances of `YOUR-PROJECT-NAME` with your actual Vercel project name
3. The URLs should look like:
   - `https://your-project-name.vercel.app/src/taskpane.html`
   - `https://your-project-name.vercel.app/src/commands.html`

### 3. Load the Add-in in Excel

#### Option A: Sideload for Testing (Excel on Windows/Mac)

1. Open Excel
2. Go to **Insert** > **My Add-ins** > **Manage My Add-ins**
3. Click **Upload My Add-in**
4. Browse and select your `manifest-prod.xml` file
5. Click **Upload**

#### Option B: Sideload for Excel Online

1. Open Excel Online (https://office.live.com/start/Excel.aspx)
2. Open any workbook
3. Go to **Insert** > **Add-ins**
4. Select **Manage My Add-ins** > **Upload My Add-in**
5. Browse and select your `manifest-prod.xml` file
6. Click **Upload**

#### Option C: Deploy to Microsoft AppSource (Production)

For production deployment to all users:
1. Register as a Microsoft Partner
2. Submit your add-in through Partner Center
3. Follow Microsoft's certification process

### 4. Test the Add-in

1. After loading the manifest, go to the **Home** tab in Excel
2. Look for the **Dust** group in the ribbon
3. Click **Call an Agent** to open the task pane
4. The add-in should load from your Vercel deployment

## Important Notes

- **HTTPS Required**: Excel add-ins must be served over HTTPS. Vercel provides this automatically.
- **CORS Headers**: The `vercel.json` configuration includes necessary CORS headers for Excel to load the add-in.
- **Domain Verification**: Make sure your Vercel domain is added to the `<AppDomains>` section in the manifest.
- **API Proxy**: The add-in uses a Vercel serverless function (`/api/dust-proxy`) to proxy requests to the Dust API, avoiding CORS issues.
- **Local Development with Proxy**: To test the proxy locally, use `vercel dev` instead of `npm start`.

## Troubleshooting

### Add-in doesn't load
- Check that all URLs in the manifest use HTTPS
- Verify the Vercel deployment is successful: `vercel ls`
- Check browser console for errors (F12 in Excel Desktop)

### CORS errors
- Ensure the `vercel.json` file is deployed with proper headers
- Verify the domain is listed in `<AppDomains>` in the manifest

### Updates not reflecting
- Clear the Office cache:
  - Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
  - Mac: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
- Remove and re-add the add-in in Excel

## Environment Variables

If your add-in needs environment variables:

1. Set them in Vercel Dashboard or CLI:
```bash
vercel env add VARIABLE_NAME
```

2. Access them in your code (if using a build step):
```javascript
const apiKey = process.env.VARIABLE_NAME;
```

## Automatic Deployments

Connect your GitHub repository to Vercel for automatic deployments:

1. Go to Vercel Dashboard
2. Import your Git repository
3. Every push to main will trigger a deployment
4. Pull requests get preview deployments