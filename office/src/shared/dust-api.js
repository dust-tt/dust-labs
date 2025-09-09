// Dust API wrapper that uses the Vercel proxy to avoid CORS issues

// Get the proxy URL - this will be your Vercel deployment URL
function getProxyUrl() {
    // In development, use localhost with Vercel dev server
    if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
        // For local development with Vercel dev server (default port 3000)
        // Note: You'll need to run `vercel dev` instead of `npm start` to test the proxy locally
        return 'http://localhost:3000/api/dust-proxy';
    }
    // In production, use the relative path (works when deployed to Vercel)
    return '/api/dust-proxy';
}

// Helper function to make API calls through the proxy
async function callDustAPI(path, options = {}) {
    const proxyUrl = getProxyUrl();
    const region = getFromStorage('region');
    
    // Build query parameters
    const params = new URLSearchParams({ path });
    
    // Prepare headers
    const headers = {
        'Content-Type': 'application/json',
        'Authorization': options.headers?.Authorization || `Bearer ${getFromStorage('dustToken')}`,
    };
    
    if (region) {
        headers['X-Dust-Region'] = region;
    }
    
    try {
        const response = await fetch(proxyUrl + '?' + params.toString(), {
            method: options.method || 'GET',
            headers: headers,
            body: options.body ? JSON.stringify(options.body) : undefined,
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || `API call failed with status ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        console.error('Dust API call failed:', error);
        throw error;
    }
}

// Export for use in taskpane.js
window.callDustAPI = callDustAPI;