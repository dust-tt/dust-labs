// Dust API wrapper shared across Office add-ins (Excel, PowerPoint)

// Helper function to determine which Office app is running
/* global Office */

const DUST_API_URL = 'https://dust.tt';

function getOfficeApp() {
    if (typeof Office !== "undefined" && Office.context?.host) {
        if (Office.context.host === Office.HostType.Excel) {
            return "excel";
        }
        if (Office.context.host === Office.HostType.PowerPoint) {
            return "powerpoint";
        }
    }

    if (localStorage.getItem("dust_excel_workspaceId")) {
        return "excel";
    }
    if (localStorage.getItem("dust_powerpoint_workspaceId")) {
        return "powerpoint";
    }

    return "excel";
}

function getStorageKey(key) {
    return `dust_${getOfficeApp()}_${key}`;
}

function getStorageValue(key) {
    return localStorage.getItem(getStorageKey(key));
}

function setStorageValue(key, value) {
    const storageKey = getStorageKey(key);
    if (value === undefined || value === null || value === "") {
        localStorage.removeItem(storageKey);
    } else {
        localStorage.setItem(storageKey, value);
    }
}

function clearTokens() {
    setStorageValue("accessToken", null);
    setStorageValue("refreshToken", null);
}

function getDustApiBaseUrl() {
    const region = getStorageValue("region");
    if (!isDevelopmentEnvironment()) {
        if (region === "europe-west1") {
            return "https://eu.dust.tt";
        }
        if (region === "us-central1") {
            return "https://dust.tt";
        }
    }
    return DUST_API_URL;
}

function isDevelopmentEnvironment() {
    const origin = window.location.origin;
    return !origin.includes("dust.tt");
}

function getStoredAccessToken() {
    return getStorageValue("accessToken");
}

async function callDustAPI(path, options = {}) {
    const baseUrl = getDustApiBaseUrl();
    let token = getStoredAccessToken();
    let hasAttemptedRefresh = false;

    if (!token) {
        throw new Error("Missing access token. Please reconnect your Dust account.");
    }

    const buildHeaders = (accessToken) => ({
        "Content-Type": "application/json",
        Authorization: `Bearer ${accessToken}`,
        ...(isDevelopmentEnvironment() ? { "ngrok-skip-browser-warning": "true" } : {}),
        ...options.headers,
    });

    const execute = async (accessToken) => {
        const fetchOptions = {
            method: options.method || "GET",
            headers: buildHeaders(accessToken),
            body: options.body ? JSON.stringify(options.body) : undefined,
        };

        if (options.signal) {
            fetchOptions.signal = options.signal;
        }

        const normalizedPath =
            path.startsWith("http://") || path.startsWith("https://")
                ? path
                : `${baseUrl}${path.startsWith("/") ? path : `/${path}`}`;

        return fetch(normalizedPath, fetchOptions);
    };

    let response = await execute(token);

    // If we get a 401, try to refresh the token
    if (response.status === 401 && !hasAttemptedRefresh) {
        console.log("[DustAPI] Received 401, attempting to refresh token...");
        hasAttemptedRefresh = true;
        const refreshedToken = await window.DustOfficeAuth.tryRefreshAccessToken();

        if (refreshedToken) {
            token = refreshedToken;
            console.log("[DustAPI] Token refreshed successfully, retrying request...");
            response = await execute(token);

            // If we still get 401 after refresh, the refresh token might be invalid
            if (response.status === 401) {
                clearTokens();
                throw new Error("Authentication failed. Please reconnect your Dust account.");
            }
        } else {
            // Refresh failed - clear tokens and throw error
            clearTokens();
            throw new Error("Token expired and refresh failed. Please reconnect your Dust account.");
        }
    }

    if (!response.ok) {
        let error;
        try {
            error = await response.json();
        } catch (jsonError) {
            error = { error: response.statusText };
        }
        throw new Error(error.error || `API call failed with status ${response.status}`);
    }

    if (response.status === 204) {
        return null;
    }

    return response.json();
}

window.getOfficeApp = getOfficeApp;
window.getStorageValue = getStorageValue;
window.setStorageValue = setStorageValue;
window.clearDustTokens = clearTokens;
window.getDustApiBaseUrl = getDustApiBaseUrl;
window.isDustDevelopmentEnv = isDevelopmentEnvironment;
window.callDustAPI = callDustAPI;