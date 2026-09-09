const urlParams = new URLSearchParams(window.location.search);
const code = urlParams.get("code");
const error = urlParams.get("error");
const errorDescription = urlParams.get("error_description");

// Track timeouts for cleanup
const pendingTimeouts = [];
function safeTimeout(fn, delay) {
    const id = setTimeout(fn, delay);
    pendingTimeouts.push(id);
    return id;
}
window.addEventListener('beforeunload', function () {
    pendingTimeouts.forEach(clearTimeout);
});

async function handleCallback() {
    // Handle OAuth callback

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
        safeTimeout(() => {
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
        safeTimeout(() => {
            window.close();
        }, 3000);
        return;
    }

    // We have a code - send it back to the parent window for token exchange
    // The parent has access to the code_verifier in its localStorage
    // Send code to parent window for token exchange

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
            safeTimeout(() => {
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
safeTimeout(() => {
    handleCallback();
}, 100);
