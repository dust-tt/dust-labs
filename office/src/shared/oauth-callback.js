/* global Office */

const urlParams = new URLSearchParams(window.location.search);
const code = urlParams.get("code");
const error = urlParams.get("error");
const errorDescription = urlParams.get("error_description");

const baseUrl = window.location.origin;

async function handleCallback() {
    console.log("[OAuth Callback] Handling callback:", {
        code: !!code,
        error,
        hasOffice: typeof Office !== "undefined",
    });

    if (error) {
        const errorMsg = errorDescription || error;
        const errorEl = document.getElementById("error");
        if (errorEl) {
            errorEl.textContent = "Authentication failed: " + errorMsg;
            errorEl.style.display = "block";
        }
        const spinner = document.querySelector(".spinner");
        if (spinner) {
            spinner.style.display = "none";
        }

        if (window.parent && window.parent !== window) {
            try {
                Office.context.ui.messageParent(
                    JSON.stringify({
                        success: false,
                        error: errorMsg,
                    })
                );
            } catch (e) {
                console.error("Failed to send message to parent:", e);
            }
        }
        return;
    }

    if (!code) {
        const errorMsg = "Missing authorization code";
        const errorEl = document.getElementById("error");
        if (errorEl) {
            errorEl.textContent = errorMsg;
            errorEl.style.display = "block";
        }
        const spinner = document.querySelector(".spinner");
        if (spinner) {
            spinner.style.display = "none";
        }

        if (window.parent && window.parent !== window) {
            try {
                Office.context.ui.messageParent(
                    JSON.stringify({
                        success: false,
                        error: errorMsg,
                    })
                );
            } catch (e) {
                console.error("Failed to send message to parent:", e);
            }
        }
        return;
    }

    const message = {
        success: true,
        code,
        action: "exchange_token",
    };

    console.log("[OAuth Callback] Attempting to send message:", {
        hasCode: !!code,
        codeLength: code ? code.length : 0,
        hasOffice: typeof Office !== "undefined",
        hasMessageParent: typeof Office !== "undefined" && Office.context?.ui?.messageParent,
        hasParent: window.parent && window.parent !== window,
    });

    let messageSent = false;

    if (typeof Office !== "undefined" && Office.context?.ui?.messageParent) {
        try {
            console.log("[OAuth Callback] Trying Office.context.ui.messageParent...");
            Office.context.ui.messageParent(JSON.stringify(message));
            console.log("[OAuth Callback] Message sent via Office dialog API");
            messageSent = true;
        } catch (e) {
            console.warn("[OAuth Callback] Office.messageParent failed:", e);
        }
    }

    if (!messageSent && window.parent && window.parent !== window) {
        try {
            console.log("[OAuth Callback] Trying window.postMessage...");
            window.parent.postMessage(JSON.stringify(message), "*");
            console.log("[OAuth Callback] Message sent via postMessage");
            messageSent = true;
        } catch (e) {
            console.error("[OAuth Callback] postMessage failed:", e);
        }
    }

    if (messageSent) {
        const spinner = document.querySelector(".spinner");
        if (spinner) {
            spinner.style.display = "none";
        }
        const status = document.querySelector(".status-message");
        if (status) {
            status.textContent = "Authentication successful! Closing...";
        }
    }
}

if (typeof Office !== "undefined" && Office.onReady) {
    console.log("[OAuth Callback] Office.js detected, waiting for onReady...");
    Office.onReady((info) => {
        console.log("[OAuth Callback] Office.onReady called:", info);
        setTimeout(() => {
            handleCallback();
        }, 100);
    });
} else {
    console.log("[OAuth Callback] Office.js not available or not ready, handling callback immediately");
    setTimeout(() => {
        handleCallback();
    }, 500);
}

