/* global Office */
(function () {
    const PKCE_VERIFIER_LENGTH = 64;

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

    function getFallbackPrefix() {
        try {
            if (typeof Office !== "undefined" && Office.context?.host && Office.HostType) {
                if (Office.context.host === Office.HostType.PowerPoint) {
                    return "dust_powerpoint_";
                }
                if (Office.context.host === Office.HostType.Excel) {
                    return "dust_excel_";
                }
            }
        } catch (error) {
            // Ignore host detection errors and fall back to storage heuristics.
        }

        if (localStorage.getItem("dust_powerpoint_workspaceId")) {
            return "dust_powerpoint_";
        }

        return "dust_excel_";
    }

    function setAuthStorage(key, value) {
        if (typeof window.setStorageValue === "function") {
            window.setStorageValue(key, value);
            return;
        }

        const prefix = getFallbackPrefix();
        if (value === undefined || value === null) {
            localStorage.removeItem(`${prefix}${key}`);
        } else {
            localStorage.setItem(`${prefix}${key}`, value);
        }
    }

    function getAuthStorage(key) {
        if (typeof window.getStorageValue === "function") {
            return window.getStorageValue(key);
        }

        return localStorage.getItem(`${getFallbackPrefix()}${key}`);
    }

    function clearAuthStorage(key) {
        setAuthStorage(key, null);
    }

    function getAuthApiBaseUrl() {
        return `${DUST_API_URL}/api/v1/auth`;
    }

    function getOAuthRedirectUri() {
        return `${window.location.origin}/shared/oauth-callback.html`;
    }

    async function tryRefreshAccessToken() {
        const refreshToken = getAuthStorage("refreshToken");
        if (!refreshToken) {
            return null;
        }

        try {
            const tokenEndpoint = `${getAuthApiBaseUrl()}/authenticate`;
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
                console.error("[DustOfficeAuth] Token refresh failed:", response.status, response.statusText);
                setAuthStorage("accessToken", null);
                setAuthStorage("refreshToken", null);
                if (typeof window.clearDustTokens === "function") {
                    window.clearDustTokens();
                }
                return null;
            }

            const data = await response.json();
            if (data.access_token) {
                setAuthStorage("accessToken", data.access_token);
            }
            if (data.refresh_token) {
                setAuthStorage("refreshToken", data.refresh_token);
            }

            // Return the new access token or get it from storage
            return data.access_token || getAuthStorage("accessToken");
        } catch (error) {
            console.error("[DustOfficeAuth] Token refresh error:", error);
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
            console.warn("[DustOfficeAuth] Failed to decode JWT payload:", error);
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
            console.log("[DustOfficeAuth] Code already being exchanged, ignoring duplicate request");
            return;
        }

        if (oauthCodeBeingExchanged !== null) {
            console.warn("[DustOfficeAuth] Another code is already being exchanged, ignoring this request");
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
                console.error("[DustOfficeAuth] Token exchange error:", errorData);
                throw new Error(errorData.details || errorData.error || `HTTP ${response.status}: Failed to exchange token`);
            }

            const data = await response.json();
            resetOAuthState();

            if (ui.onAuthSuccess) {
                await ui.onAuthSuccess(data);
            }
        } catch (error) {
            console.error("[DustOfficeAuth] OAuth callback error:", error);
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

            console.log("[DustOfficeAuth] Authorization parameters prepared:", {
                redirectUri,
                authUrlPreview: authUrl.substring(0, 150) + "...",
            });

            if (typeof Office !== "undefined" && Office.context?.ui) {
                Office.context.ui.displayDialogAsync(
                    authUrl,
                    { height: 60, width: 30 },
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                            console.error("[DustOfficeAuth] Dialog failed to open:", asyncResult.error);
                            showError(ui.errorElement, "Failed to open authentication window");
                            hideLoading(ui.loadingElement);
                            showConnect(ui.connectButton);
                            if (ui.onAuthError) {
                                ui.onAuthError(asyncResult.error || new Error("Failed to open authentication window"));
                            }
                            return;
                        }

                        const dialog = asyncResult.value;
                        let dialogMessageProcessed = false;

                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                            console.log("[DustOfficeAuth] Dialog message received");

                            if (dialogMessageProcessed) {
                                console.log("[DustOfficeAuth] Dialog message already processed, ignoring");
                                return;
                            }

                            dialogMessageProcessed = true;

                            try {
                                dialog.close();
                            } catch (e) {
                                console.error("[DustOfficeAuth] Error closing dialog:", e);
                            }

                            try {
                                const result = JSON.parse(arg.message);

                                if (result.success && result.action === "exchange_token" && result.code) {
                                    void exchangeCodeForToken(result.code, ui);
                                } else if (result.success && result.access_token) {
                                    if (ui.onAuthSuccess) {
                                        void ui.onAuthSuccess(result);
                                    }
                                } else {
                                    const err = new Error(result.error || "Authentication failed");
                                    showError(ui.errorElement, err.message);
                                    hideLoading(ui.loadingElement);
                                    showConnect(ui.connectButton);
                                    if (ui.onAuthError) {
                                        ui.onAuthError(err);
                                    }
                                }
                            } catch (error) {
                                console.error("[DustOfficeAuth] Failed to parse dialog message:", error);
                                const err = new Error("Failed to process authentication response");
                                showError(ui.errorElement, err.message);
                                hideLoading(ui.loadingElement);
                                showConnect(ui.connectButton);
                                if (ui.onAuthError) {
                                    ui.onAuthError(err);
                                }
                            }
                        });

                        dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
                            console.log("[DustOfficeAuth] Dialog event received:", arg);
                            if (arg.error === 12006) {
                                hideLoading(ui.loadingElement);
                                showConnect(ui.connectButton);
                            }
                        });
                    }
                );
            } else {
                window.open(authUrl, "workos-oauth", "width=600,height=700");
                pollForOAuthCallback(options);
            }
        } catch (error) {
            console.error("[DustOfficeAuth] OAuth initiation error:", error);
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
                if (ui.onAuthSuccess) {
                    void ui.onAuthSuccess({ access_token: accessToken });
                }
            }
        }, 1000);

        setTimeout(() => clearInterval(interval), 300000);
    }

    window.DustOfficeAuth = {
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
    };
})();
