// Dust API wrapper that uses the Vercel proxy to avoid CORS issues

// Detect which Office app is running
function getOfficeApp() {
    // Check if we're in an Office context
    if (typeof Office !== 'undefined' && Office.context && Office.context.host) {
        if (Office.context.host === Office.HostType.Excel) {
            return 'excel';
        } else if (Office.context.host === Office.HostType.PowerPoint) {
            return 'powerpoint';
        }
    }
    
    // Fallback: check which storage keys exist
    if (localStorage.getItem('dust_excel_workspaceId')) {
        return 'excel';
    } else if (localStorage.getItem('dust_powerpoint_workspaceId')) {
        return 'powerpoint';
    }
    
    // Default to excel if unable to determine
    return 'excel';
}

// Helper function to get from storage based on current Office app
function getDustStorageValue(key) {
    const app = getOfficeApp();
    const storageKey = `dust_${app}_${key}`;
    return localStorage.getItem(storageKey);
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
        const fetchOptions = {
            method: options.method || 'GET',
            headers: headers,
            body: options.body ? JSON.stringify(options.body) : undefined,
        };
        
        // Add abort signal if provided
        if (options.signal) {
            fetchOptions.signal = options.signal;
        }
        
        const response = await fetch(proxyUrl + '?' + params.toString(), fetchOptions);
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || `API call failed with status ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        // Don't log abort errors
        if (error.name !== 'AbortError') {
            console.error('Dust API call failed:', error);
        }
        throw error;
    }
}

// Export for use in taskpane.js
window.callDustAPI = callDustAPI;
window.getOfficeApp = getOfficeApp;