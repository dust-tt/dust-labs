# OAuth Setup for Zendesk App

## Problem: Changing Redirect URIs

Zendesk deploys app assets to paths that include a hash that changes on every deployment:
```
https://yourdomain.zendesk.com/1073173/assets/1765878898-ec18c462b053cebc3858fbb5e43625e1/oauth-callback.html
```

Where `1765878898-ec18c462b053cebc3858fbb5e43625e1` changes with each deployment.

**Note:** WorkOS does not support wildcard characters in URL paths (only in subdomains), so the redirect URI must be updated in WorkOS after each deployment.

## Setup Instructions

### 1. Deploy the Zendesk App

Deploy your Zendesk app using the standard deployment process.

### 2. Get the Redirect URI

After deployment:

1. Open the app in a Zendesk ticket sidebar
2. Open browser console (F12 or right-click → Inspect)
3. Click "Login to Dust"
4. Look for the log message in console:
   ```
   [DustZendeskAuth] Redirect URI: https://yourdomain.zendesk.com/1073173/assets/[hash]/oauth-callback.html
   ```
5. Copy this full URL

### 3. Configure WorkOS

1. Go to your WorkOS Dashboard
2. Navigate to your OAuth Application settings
3. In the "Redirect URIs" section, add the copied URL exactly as shown
4. Save the configuration

### 4. Test the OAuth Flow

1. Refresh the Zendesk app
2. Click "Login to Dust"
3. Complete the OAuth flow
4. Authentication should succeed

## After Each Deployment

After deploying a new version of the app:

1. Get the new redirect URI from browser console (see step 2 above)
2. Add the new URI to WorkOS redirect URIs
3. (Optional) Remove old URIs from previous deployments to keep the list clean

## Troubleshooting

### OAuth fails with "Invalid redirect URI"

1. Check the browser console for the redirect URI being used:
   ```
   [DustZendeskAuth] Redirect URI: https://...
   ```

2. Verify this exact URI is registered in WorkOS (check for trailing slashes, typos, etc.)

3. If the URI doesn't match, update WorkOS with the correct URI

### Callback page doesn't redirect back

1. Verify `oauth-callback.html` is accessible at the redirect URI
2. Check browser console for JavaScript errors
3. Ensure the popup window isn't being blocked by the browser

### Authentication completes but nothing happens

1. Check if the popup was blocked
2. Verify the `postMessage` communication is working (check console for messages)
3. Try refreshing the main app page
