const urlParams = new URLSearchParams(window.location.search);
const code = urlParams.get("code");
const error = urlParams.get("error");
const errorDescription = urlParams.get("error_description");

async function handleCallback() {
    console.log("[OAuth Callback] Handling callback:", {
        code: !!code,
        error,
        hasOpener: !!window.opener,
    });

    const statusEl = document.querySelector(".status-message");
    const errorEl = document.getElementById("error");
    const successEl = document.getElementById("success");
    const spinner = document.querySelector(".spinner");

    if (error) {
        const errorMsg = errorDescription || error;
        if (errorEl) {
            errorEl.textContent = "Authentication failed: " + errorMsg;
            errorEl.style.display = "block";
        }
        if (spinner) {
            spinner.style.display = "none";
        }
        if (statusEl) {
            statusEl.style.display = "none";
        }

        // Try to notify parent window
        if (window.opener && !window.opener.closed) {
            try {
                window.opener.postMessage(
                    JSON.stringify({
                        success: false,
                        error: errorMsg,
                    }),
                    window.location.origin
                );
            } catch (e) {
                console.error("Failed to send message to opener:", e);
            }
        }

        // Close window after a delay
        setTimeout(() => {
            window.close();
        }, 3000);
        return;
    }

    if (!code) {
        const errorMsg = "Missing authorization code";
        if (errorEl) {
            errorEl.textContent = errorMsg;
            errorEl.style.display = "block";
        }
        if (spinner) {
            spinner.style.display = "none";
        }
        if (statusEl) {
            statusEl.style.display = "none";
        }

        // Try to notify parent window
        if (window.opener && !window.opener.closed) {
            try {
                window.opener.postMessage(
                    JSON.stringify({
                        success: false,
                        error: errorMsg,
                    }),
                    window.location.origin
                );
            } catch (e) {
                console.error("Failed to send message to opener:", e);
            }
        }

        // Close window after a delay
        setTimeout(() => {
            window.close();
        }, 3000);
        return;
    }

    // We have a code - send it back to the parent window for token exchange
    // The parent has access to the code_verifier in its localStorage
    console.log("[OAuth Callback] Sending code to parent window for token exchange");

    if (window.opener && !window.opener.closed) {
        try {
            window.opener.postMessage(
                JSON.stringify({
                    success: true,
                    code: code,
                    action: 'exchange_token',
                }),
                window.location.origin
            );

            // Show success message
            if (spinner) {
                spinner.style.display = "none";
            }
            if (statusEl) {
                statusEl.textContent = "Authentication successful! Completing setup...";
            }
            if (successEl) {
                successEl.textContent = "You can close this window.";
                successEl.style.display = "block";
            }

            // Close window after a delay
            setTimeout(() => {
                window.close();
            }, 2000);

        } catch (e) {
            console.error("Failed to send message to opener:", e);
            if (errorEl) {
                errorEl.textContent = "Failed to communicate with main window. Please close this window and try again.";
                errorEl.style.display = "block";
            }
            if (spinner) {
                spinner.style.display = "none";
            }
        }
    } else {
        console.error("[OAuth Callback] No opener window found");
        if (errorEl) {
            errorEl.textContent = "Parent window not found. Please close this window and try again.";
            errorEl.style.display = "block";
        }
        if (spinner) {
            spinner.style.display = "none";
        }
    }
}

// Run callback handler after a brief delay
setTimeout(() => {
    handleCallback();
}, 100);
