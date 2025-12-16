(function () {
    const PKCE_VERIFIER_LENGTH = 64;
    const STORAGE_PREFIX = 'dust_zendesk_';

    let oauthCodeBeingExchanged = null;

    function getCrypto() {
        if (typeof window !== "undefined" && window.crypto) {
            return window.crypto;
        }
        throw new Error("Web Crypto API is not available in this environment.");
    }

    function base64UrlEncode(buffer) {
        const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
        let binary = "";
        for (let i = 0; i < bytes.byteLength; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    }

    function generateCodeVerifier() {
        const crypto = getCrypto();
        const randomBytes = new Uint8Array(PKCE_VERIFIER_LENGTH);
        crypto.getRandomValues(randomBytes);
        return base64UrlEncode(randomBytes);
    }

    async function generateCodeChallenge(codeVerifier) {
        const crypto = getCrypto();
        if (!crypto.subtle || !crypto.subtle.digest) {
            throw new Error("Crypto digest API is not available in this environment.");
        }
        const encoder = new TextEncoder();
        const data = encoder.encode(codeVerifier);
        const digest = await crypto.subtle.digest("SHA-256", data);
        return base64UrlEncode(new Uint8Array(digest));
    }

    function setAuthStorage(key, value) {
        const storageKey = `${STORAGE_PREFIX}${key}`;
        if (value === undefined || value === null) {
            localStorage.removeItem(storageKey);
        } else {
            localStorage.setItem(storageKey, value);
        }
    }

    function getAuthStorage(key) {
        return localStorage.getItem(`${STORAGE_PREFIX}${key}`);
    }

    function clearAuthStorage(key) {
        setAuthStorage(key, null);
    }

    function getAuthApiBaseUrl() {
        // Check if DUST_API_URL is defined globally, otherwise use default
        const baseUrl = typeof DUST_API_URL !== 'undefined' ? DUST_API_URL : 'https://dust.tt';
        return `${baseUrl}/api/v1/auth`;
    }

    function getOAuthRedirectUri() {
        // Construct redirect URI from current deployment path
        // For Zendesk apps, the path includes a hash that changes on every deploy:
        // /1073173/assets/1765878898-ec18c462b053cebc3858fbb5e43625e1/iframe.html
        const currentPath = window.location.pathname;
        const pathParts = currentPath.split('/');

        // Replace the last part (e.g., "iframe.html") with "oauth-callback.html"
        pathParts[pathParts.length - 1] = 'oauth-callback.html';

        const redirectPath = pathParts.join('/');
        const redirectUri = `${window.location.origin}${redirectPath}`;

        console.log('[DustZendeskAuth] Redirect URI:', redirectUri);

        return redirectUri;
    }

    async function tryRefreshAccessToken(requestFunction) {
        const refreshToken = getAuthStorage("refreshToken");
        if (!refreshToken) {
            return null;
        }

        try {
            const tokenEndpoint = `${getAuthApiBaseUrl()}/authenticate`;
            let data;

            if (requestFunction) {
                // Use Zendesk proxy to avoid CORS
                data = await requestFunction({
                    url: tokenEndpoint,
                    type: "POST",
                    contentType: "application/x-www-form-urlencoded",
                    data: new URLSearchParams({
                        grant_type: "refresh_token",
                        refresh_token: refreshToken,
                    }).toString(),
                    secure: true,
                });
            } else {
                // Direct fetch (fallback)
                const response = await fetch(tokenEndpoint, {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/x-www-form-urlencoded",
                    },
                    body: new URLSearchParams({
                        grant_type: "refresh_token",
                        refresh_token: refreshToken,
                    }).toString(),
                });

                if (!response.ok) {
                    console.error("[DustZendeskAuth] Token refresh failed:", response.status, response.statusText);
                    setAuthStorage("accessToken", null);
                    setAuthStorage("refreshToken", null);
                    return null;
                }

                data = await response.json();
            }

            if (data.access_token) {
                setAuthStorage("accessToken", data.access_token);
            }
            if (data.refresh_token) {
                setAuthStorage("refreshToken", data.refresh_token);
            }

            return data.access_token || getAuthStorage("accessToken");
        } catch (error) {
            console.error("[DustZendeskAuth] Token refresh error:", error);
            return null;
        }
    }

    async function prepareAuthorizationRequest() {
        const codeVerifier = generateCodeVerifier();
        setAuthStorage("oauthCodeVerifier", codeVerifier);

        const codeChallenge = await generateCodeChallenge(codeVerifier);
        const redirectUri = getOAuthRedirectUri();
        setAuthStorage("oauthRedirectUri", redirectUri);

        const params = new URLSearchParams({
            redirect_uri: redirectUri,
            response_type: "code",
            code_challenge_method: "S256",
            code_challenge: codeChallenge,
        });

        return {
            authUrl: `${getAuthApiBaseUrl()}/authorize?${params.toString()}`,
            redirectUri,
        };
    }

    function decodeJwtPayload(token) {
        try {
            const parts = token.split(".");
            if (parts.length !== 3) {
                throw new Error("Invalid token format");
            }
            const payload = parts[1];
            const paddedPayload = payload + "=".repeat((4 - (payload.length % 4)) % 4);
            const decoded = atob(paddedPayload.replace(/-/g, "+").replace(/_/g, "/"));
            return JSON.parse(decoded);
        } catch (error) {
            console.warn("[DustZendeskAuth] Failed to decode JWT payload:", error);
            return null;
        }
    }

    function decodeToken(accessToken) {
        const decoded = decodeJwtPayload(accessToken);
        const workspaceId = decoded && decoded["https://dust.tt/workspaceId"];
        const region = decoded && decoded["https://dust.tt/region"];
        return { workspaceId, region };
    }

    function normalizeOptions(options = {}) {
        return {
            errorElement: options.errorElement || null,
            loadingElement: options.loadingElement || null,
            connectButton: options.connectButton || null,
            onAuthSuccess: typeof options.onAuthSuccess === "function" ? options.onAuthSuccess : null,
            onAuthError: typeof options.onAuthError === "function" ? options.onAuthError : null,
            requestFunction: options.requestFunction || null, // For proxying requests through Zendesk
        };
    }

    function showError(errorElement, message) {
        if (!errorElement) {
            return;
        }
        errorElement.textContent = message.startsWith("❌") ? message : `❌ ${message}`;
        errorElement.style.display = "block";
    }

    function hideError(errorElement) {
        if (errorElement) {
            errorElement.style.display = "none";
        }
    }

    function showLoading(loadingElement) {
        if (loadingElement) {
            loadingElement.style.display = "block";
        }
    }

    function hideLoading(loadingElement) {
        if (loadingElement) {
            loadingElement.style.display = "none";
        }
    }

    function showConnect(connectButton) {
        if (connectButton) {
            connectButton.style.display = "block";
        }
    }

    function hideConnect(connectButton) {
        if (connectButton) {
            connectButton.style.display = "none";
        }
    }

    function resetOAuthState() {
        oauthCodeBeingExchanged = null;
        clearAuthStorage("oauthCodeVerifier");
        clearAuthStorage("oauthRedirectUri");
    }

    async function exchangeCodeForToken(code, options = {}) {
        const ui = normalizeOptions(options);

        if (oauthCodeBeingExchanged === code) {
            console.log("[DustZendeskAuth] Code already being exchanged, ignoring duplicate request");
            return;
        }

        if (oauthCodeBeingExchanged !== null) {
            console.warn("[DustZendeskAuth] Another code is already being exchanged, ignoring this request");
            return;
        }

        oauthCodeBeingExchanged = code;

        try {
            const codeVerifier = getAuthStorage("oauthCodeVerifier");
            if (!codeVerifier) {
                throw new Error("Missing code_verifier. Please restart the OAuth flow.");
            }

            const redirectUri = getOAuthRedirectUri();
            const tokenEndpoint = `${getAuthApiBaseUrl()}/authenticate`;

            const tokenParams = new URLSearchParams({
                code,
                code_verifier: codeVerifier,
                redirect_uri: redirectUri,
                grant_type: "authorization_code",
            });

            console.log("[DustZendeskAuth] Exchanging code for token...", {
                endpoint: tokenEndpoint,
                hasRequestFunction: !!ui.requestFunction,
            });

            let data;

            if (ui.requestFunction) {
                // Use Zendesk proxy to avoid CORS
                console.log("[DustZendeskAuth] Using Zendesk proxy for token exchange");
                try {
                    data = await ui.requestFunction({
                        url: tokenEndpoint,
                        type: "POST",
                        contentType: "application/x-www-form-urlencoded",
                        data: tokenParams.toString(),
                        secure: true,
                    });
                } catch (requestError) {
                    console.error("[DustZendeskAuth] Token exchange error via proxy:", requestError);
                    throw new Error(requestError.responseJSON?.error || requestError.responseJSON?.details || "Failed to exchange token");
                }
            } else {
                // Direct fetch (fallback)
                console.log("[DustZendeskAuth] Using direct fetch for token exchange");
                const response = await fetch(tokenEndpoint, {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/x-www-form-urlencoded",
                    },
                    body: tokenParams.toString(),
                });

                if (!response.ok) {
                    let errorData;
                    try {
                        errorData = await response.json();
                    } catch (jsonError) {
                        const errorText = await response.text();
                        errorData = { error: errorText || "Unknown error" };
                    }
                    console.error("[DustZendeskAuth] Token exchange error:", errorData);
                    throw new Error(errorData.details || errorData.error || `HTTP ${response.status}: Failed to exchange token`);
                }

                data = await response.json();
            }

            console.log("[DustZendeskAuth] Token exchange successful");
            resetOAuthState();

            if (ui.onAuthSuccess) {
                await ui.onAuthSuccess(data);
            }
        } catch (error) {
            console.error("[DustZendeskAuth] OAuth callback error:", error);
            oauthCodeBeingExchanged = null;

            showError(ui.errorElement, error.message);
            hideLoading(ui.loadingElement);
            showConnect(ui.connectButton);

            if (ui.onAuthError) {
                ui.onAuthError(error);
            }

            throw error;
        }
    }

    async function initiateOAuth(options = {}) {
        const ui = normalizeOptions(options);

        hideError(ui.errorElement);
        showLoading(ui.loadingElement);
        hideConnect(ui.connectButton);

        try {
            const { authUrl, redirectUri } = await prepareAuthorizationRequest();

            console.log("[DustZendeskAuth] Authorization parameters prepared:", {
                redirectUri,
                authUrlPreview: authUrl.substring(0, 150) + "...",
            });

            // Open OAuth window
            const width = 600;
            const height = 700;
            const left = window.screen.width / 2 - width / 2;
            const top = window.screen.height / 2 - height / 2;

            window.open(
                authUrl,
                "dust-oauth",
                `width=${width},height=${height},left=${left},top=${top}`
            );

            // Poll for OAuth callback
            pollForOAuthCallback(options);
        } catch (error) {
            console.error("[DustZendeskAuth] OAuth initiation error:", error);
            showError(ui.errorElement, error.message);
            hideLoading(ui.loadingElement);
            showConnect(ui.connectButton);
            if (ui.onAuthError) {
                ui.onAuthError(error);
            }
        }
    }

    async function checkOAuthCallback(options = {}) {
        const ui = normalizeOptions(options);
        const urlParams = new URLSearchParams(window.location.search);
        const code = urlParams.get("code");
        const error = urlParams.get("error");

        if (error) {
            showError(ui.errorElement, `OAuth error: ${error}`);
            hideLoading(ui.loadingElement);
            showConnect(ui.connectButton);
            window.history.replaceState({}, document.title, window.location.pathname);
            if (ui.onAuthError) {
                ui.onAuthError(new Error(error));
            }
            return;
        }

        if (code) {
            try {
                await exchangeCodeForToken(code, ui);
                window.history.replaceState({}, document.title, window.location.pathname);
            } catch (exchangeError) {
                // Error already handled inside exchangeCodeForToken
            }
        }

        const oauthCode = sessionStorage.getItem("oauth_code");
        if (oauthCode) {
            sessionStorage.removeItem("oauth_code");
            try {
                await exchangeCodeForToken(oauthCode, ui);
            } catch (exchangeError) {
                // Error already handled inside exchangeCodeForToken
            }
        }
    }

    function pollForOAuthCallback(options = {}) {
        const ui = normalizeOptions(options);
        const interval = setInterval(() => {
            const accessToken = getAuthStorage("accessToken");
            if (accessToken) {
                clearInterval(interval);
                resetOAuthState();
                hideLoading(ui.loadingElement);
                showConnect(ui.connectButton);
                if (ui.onAuthSuccess) {
                    const user = getAuthStorage("user");
                    const refreshToken = getAuthStorage("refreshToken");
                    void ui.onAuthSuccess({
                        access_token: accessToken,
                        refresh_token: refreshToken,
                        user: user ? JSON.parse(user) : null
                    });
                }
            }
        }, 1000);

        setTimeout(() => {
            clearInterval(interval);
            hideLoading(ui.loadingElement);
            showConnect(ui.connectButton);
        }, 300000); // 5 minutes timeout
    }

    // Clear all auth storage
    function clearAuth() {
        clearAuthStorage("accessToken");
        clearAuthStorage("refreshToken");
        clearAuthStorage("workspaceId");
        clearAuthStorage("region");
        clearAuthStorage("user");
        clearAuthStorage("oauthCodeVerifier");
        clearAuthStorage("oauthRedirectUri");
    }

    window.DustZendeskAuth = {
        initiateOAuth,
        exchangeCodeForToken,
        checkOAuthCallback,
        pollForOAuthCallback,
        prepareAuthorizationRequest,
        getAuthApiBaseUrl,
        getOAuthRedirectUri,
        decodeJwtPayload,
        decodeToken,
        tryRefreshAccessToken,
        getAuthStorage,
        setAuthStorage,
        clearAuth,
    };
})();
