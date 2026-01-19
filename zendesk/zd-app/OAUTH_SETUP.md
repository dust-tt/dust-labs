# OAuth Setup for Zendesk App

## Solution: Stable Redirect URI via Proxy

Zendesk deploys app assets to paths that include a hash that changes on every deployment:
```
https://1073173.apps.zdusercontent.com/1073173/assets/1765878898-ec18c462b053cebc3858fbb5e43625e1/oauth-callback.html
```

Since WorkOS doesn't support wildcard characters in URL paths, we use a stable proxy endpoint on dust.tt to handle the redirect.

### How It Works

1. The Zendesk app uses a **stable redirect URI**: `https://dust.tt/api/oauth-zendesk/callback`
2. The actual Zendesk callback URL is encoded in the OAuth `state` parameter
3. After OAuth completes, WorkOS redirects to the dust.tt proxy endpoint
4. The proxy decodes the `state` parameter and redirects to the actual Zendesk callback URL with the authorization code

This means you only need to configure the redirect URI **once** in WorkOS, and it will work across all deployments.

## One-Time Setup

### 1. Configure WorkOS

Add the following redirect URIs to your WorkOS OAuth Application:

**For US region:**
```
https://dust.tt/api/oauth-zendesk/callback
```

**For EU region:**
```
https://eu.dust.tt/api/oauth-zendesk/callback
```

### 2. Deploy the Zendesk App

Deploy your Zendesk app using the standard deployment process. No additional configuration is needed after deployment.

### 3. Test the OAuth Flow

1. Open the app in a Zendesk ticket sidebar
2. Click "Login to Dust"
3. Complete the OAuth flow
4. Authentication should succeed

## Debugging

Open browser console (F12) to see OAuth debug messages:

```
[DustZendeskAuth] OAuth redirect URI (stable): https://dust.tt/api/oauth-zendesk/callback
[DustZendeskAuth] Zendesk callback URL: https://1073173.apps.zdusercontent.com/.../oauth-callback.html
```

The first URL is what's registered in WorkOS (stable).
The second URL is where the user will be redirected after OAuth (changes per deployment, encoded in state).

## Troubleshooting

### OAuth fails with "Invalid redirect URI"

1. Verify the stable redirect URI is registered in WorkOS:
   - US: `https://dust.tt/api/oauth-zendesk/callback`
   - EU: `https://eu.dust.tt/api/oauth-zendesk/callback`

2. Check which region the app is using (console logs show which dust.tt URL is being used)

### Callback page doesn't redirect back

1. Check browser console for errors from the proxy endpoint
2. Verify `oauth-callback.html` exists in the Zendesk app assets
3. Ensure the popup window isn't being blocked

### Authentication completes but nothing happens

1. Check if the popup was blocked
2. Verify the `postMessage` communication is working (check console for messages)
3. Try refreshing the main app page
