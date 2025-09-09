// Dust API wrapper that uses the Vercel proxy to avoid CORS issues

// Helper function to get from storage (compatible with both Excel and PowerPoint)
function getDustStorageValue(key) {
    // Try Excel storage first
    let value = localStorage.getItem(`dust_excel_${key}`);
    if (value) return value;
    
    // Fall back to PowerPoint storage
    return localStorage.getItem(`dust_powerpoint_${key}`);
}

// Get the proxy URL
function getProxyUrl() {
    // Always use relative path - works in both development and production
    // Vercel automatically handles the /api routes
    return '/api/dust-proxy';
}

// Helper function to make API calls through the proxy
async function callDustAPI(path, options = {}) {
    const proxyUrl = getProxyUrl();
    const region = getDustStorageValue('region');
    
    // Build query parameters
    const params = new URLSearchParams({ path });
    
    // Prepare headers
    const headers = {
        'Content-Type': 'application/json',
        'Authorization': options.headers?.Authorization || `Bearer ${getDustStorageValue('dustToken')}`,
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