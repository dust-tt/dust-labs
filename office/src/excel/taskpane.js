/* global Office, Excel, DustOfficeAuth, $, callDustAPI */

let processingProgress = { current: 0, total: 0, status: "idle" };
let cancelRequested = false;
let currentAbortController = null;
let currentProcessingId = null;

function buildAuthOptions() {
    return {
        errorElement: document.getElementById("credentialError"),
        loadingElement: document.getElementById("oauthLoading"),
        connectButton: document.getElementById("connectWorkOS"),
        onAuthSuccess: handleOAuthSuccess,
        onAuthError: (error) => {
            console.error("[Taskpane] OAuth error:", error);
        },
    };
}

// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("selectCellsBtn").onclick = useSelection;
        document.getElementById("selectTargetBtn").onclick = useTargetSelection;
        document.getElementById("cellRange").addEventListener("input", function () {
            void updateRangeInfo(this.value);
        });
        document.getElementById("myForm").addEventListener("submit", handleSubmit);
        document.getElementById("connectWorkOS").onclick = () => {
            DustOfficeAuth.initiateOAuth(buildAuthOptions());
        };
        const logoutBtn = document.getElementById("logoutBtn");
        logoutBtn.onclick = () => {
            showCredentialSetup();
            removeCredentials();
        };
        document.getElementById("removeCredentialsBtn").onclick = showRemoveConfirmation;
        document.getElementById("cancelBtn").onclick = cancelProcessing;
        document.getElementById("confirmRemove").onclick = removeCredentials;
        document.getElementById("cancelRemove").onclick = hideRemoveConfirmation;

        // Listen for postMessage from OAuth callback (fallback if Office dialog API doesn't work)
        // Use once: true to prevent multiple handlers, or check if already processed
        let postMessageProcessed = false;
        window.addEventListener('message', function (event) {
            console.log('[Taskpane] postMessage received:', event.data, event.origin);

            // Prevent processing the same message multiple times
            if (postMessageProcessed) {
                console.log('[Taskpane] postMessage already processed, ignoring');
                return;
            }

            try {
                const result = JSON.parse(event.data);
                if (result.success && result.action === 'exchange_token' && result.code) {
                    console.log('[Taskpane] Received code via postMessage, exchanging...');
                    postMessageProcessed = true; // Mark as processed
                    void DustOfficeAuth.exchangeCodeForToken(result.code, buildAuthOptions());
                }
            } catch (e) {
                console.warn('[Taskpane] Failed to parse postMessage:', e);
            }
        });

        // Check for OAuth callback
        DustOfficeAuth.checkOAuthCallback(buildAuthOptions());

        // Initialize
        checkCredentialsAndInitialize();
    }
});

// Initialize Select2
function initializeSelect2() {
    $("#assistant").select2({
        placeholder: "Loading agents...",
        allowClear: true,
        width: "100%",
        language: {
            noResults: function () {
                return "No agents found";
            }
        }
    });
}

// Storage functions
function saveToStorage(key, value) {
    const storageKey = `dust_excel_${key}`;
    if (value === undefined || value === null) {
        localStorage.removeItem(storageKey);
        return;
    }
    localStorage.setItem(storageKey, value);
}

function getFromStorage(key) {
    return localStorage.getItem(`dust_excel_${key}`);
}

// Check credentials and initialize the appropriate view
function checkCredentialsAndInitialize() {
    const accessToken = getFromStorage("accessToken");
    const workspaceId = getFromStorage("workspaceId");

    if (accessToken && workspaceId) {
        // Credentials exist, show the main form
        showMainForm();
        loadAssistants();
        initializeSelect2();
    } else {
        // No credentials, show setup panel
        showCredentialSetup();
    }
}

// Show credential setup panel
function showCredentialSetup() {
    document.getElementById("credentialSetup").style.display = "block";
    document.getElementById("myForm").style.display = "none";
    const logoutBtn = document.getElementById("logoutBtn");
    if (logoutBtn) {
        logoutBtn.style.display = "none";
    }

    // Hide error message and loading initially
    const errorDiv = document.getElementById("credentialError");
    const loadingDiv = document.getElementById("oauthLoading");
    if (errorDiv) {
        errorDiv.style.display = "none";
    }
    if (loadingDiv) {
        loadingDiv.style.display = "none";
    }

    // Show/hide remove button based on whether credentials exist
    const accessToken = getFromStorage("accessToken");
    const workspaceId = getFromStorage("workspaceId");
    const removeBtn = document.getElementById("removeCredentialsBtn");
    if (removeBtn) {
        removeBtn.style.display = (accessToken && workspaceId) ? "block" : "none";
    }
}

// Show main form
function showMainForm() {
    document.getElementById("credentialSetup").style.display = "none";
    document.getElementById("myForm").style.display = "block";
    const logoutBtn = document.getElementById("logoutBtn");
    if (logoutBtn) {
        logoutBtn.style.display = "inline-block";
    }
}

// Show remove confirmation
function showRemoveConfirmation() {
    document.getElementById("removeCredentialsBtn").style.display = "none";
    document.getElementById("removeConfirmation").style.display = "block";
}

// Hide remove confirmation
function hideRemoveConfirmation() {
    document.getElementById("removeConfirmation").style.display = "none";
    document.getElementById("removeCredentialsBtn").style.display = "block";
}

// Remove credentials
function removeCredentials() {
    // Clear all stored credentials
    localStorage.removeItem("dust_excel_workspaceId");
    localStorage.removeItem("dust_excel_accessToken");
    localStorage.removeItem("dust_excel_refreshToken");
    localStorage.removeItem("dust_excel_region");
    localStorage.removeItem("dust_excel_credentialsConfigured");
    localStorage.removeItem("dust_excel_user");
    localStorage.removeItem("dust_excel_oauthCodeVerifier");
    localStorage.removeItem("dust_excel_oauthRedirectUri");

    // Hide remove button and confirmation
    document.getElementById("removeCredentialsBtn").style.display = "none";
    document.getElementById("removeConfirmation").style.display = "none";

    const connectBtn = document.getElementById("connectWorkOS");
    if (connectBtn) {
        connectBtn.style.display = "block";
    }

    // Hide error message
    const errorDiv = document.getElementById("credentialError");
    if (errorDiv) {
        errorDiv.style.display = "none";
    }

    // Clear the assistant dropdown
    const select = document.getElementById("assistant");
    select.innerHTML = '<option value=""></option>';
    select.disabled = true;

    // Reset Select2
    $("#assistant").select2({
        placeholder: "Loading agents...",
        allowClear: true,
        width: "100%"
    });

    // Clear any error messages in the main form
    const loadError = document.getElementById("loadError");
    if (loadError) {
        loadError.style.display = "none";
    }

    // Show credential setup as if starting fresh
    document.getElementById("credentialSetup").style.display = "block";
    document.getElementById("myForm").style.display = "none";
    const logoutBtn = document.getElementById("logoutBtn");
    if (logoutBtn) {
        logoutBtn.style.display = "none";
    }

    showCredentialSetup();
}

// Handle successful OAuth
async function handleOAuthSuccess(data) {
    const { access_token, user, refresh_token } = data;

    if (!access_token) {
        throw new Error('No access token received');
    }

    // Store credentials
    saveToStorage("accessToken", access_token);
    saveToStorage("refreshToken", refresh_token);

    const { workspaceId, region } = DustOfficeAuth.decodeToken(access_token);

    saveToStorage("workspaceId", workspaceId);
    saveToStorage("region", region);
    saveToStorage("user", JSON.stringify(user));

    // Hide loading and error
    const loadingDiv = document.getElementById("oauthLoading");
    const errorDiv = document.getElementById("credentialError");
    if (loadingDiv) {
        loadingDiv.style.display = "none";
    }
    if (errorDiv) {
        errorDiv.style.display = "none";
    }

    // If we don't have workspace_id from OAuth, we might need to fetch it
    // For now, we'll try to use the token to get workspace info
    try {
        if (!workspaceId) {
            // Workspace ID not available from OAuth, show error
            throw new Error('Workspace ID not found. Please ensure your WorkOS integration is configured correctly.');
        }

        saveToStorage("workspaceId", workspaceId);

        // Test credentials by fetching agents
        const apiPath = `/api/v1/w/${workspaceId}/assistant/agent_configurations`;
        await callDustAPI(apiPath);

        // Credentials are valid
        saveToStorage("credentialsConfigured", "true");

        // Switch to main form
        showMainForm();
        loadAssistants();
        initializeSelect2();
    } catch (error) {
        console.error('Failed to validate token:', error);
        // Still show error but keep token stored
        if (errorDiv) {
            errorDiv.textContent = "❌ " + error.message;
            errorDiv.style.display = "block";
        }
        const connectBtn = document.getElementById("connectWorkOS");
        if (connectBtn) {
            connectBtn.style.display = "block";
        }
        if (loadingDiv) {
            loadingDiv.style.display = "none";
        }
    }
}

async function loadAssistants() {
    const token = getFromStorage("accessToken");
    const workspaceId = getFromStorage("workspaceId");

    if (!token || !workspaceId) {
        const errorDiv = document.getElementById("loadError");
        errorDiv.textContent = "❌ Please connect your Dust account first";
        errorDiv.style.display = "block";
        $("#assistant").select2({
            placeholder: "Failed to load agents",
            allowClear: true,
            width: "100%"
        });
        return;
    }

    try {
        const apiPath = `/api/v1/w/${workspaceId}/assistant/agent_configurations`;
        const data = await callDustAPI(apiPath);
        const assistants = data.agentConfigurations;

        const sortedAssistants = assistants.sort((a, b) => a.name.localeCompare(b.name));

        const select = document.getElementById("assistant");
        select.innerHTML = "";

        const emptyOption = document.createElement("option");
        emptyOption.value = "";
        select.appendChild(emptyOption);

        sortedAssistants.forEach(a => {
            const option = document.createElement("option");
            option.value = a.sId;
            option.textContent = a.name;
            select.appendChild(option);
        });

        select.disabled = false;
        document.getElementById("loadError").style.display = "none";

        $("#assistant").select2({
            placeholder: "Select an agent",
            allowClear: true,
            width: "100%",
            language: {
                noResults: function () {
                    return "No agents found";
                }
            }
        });

        if (assistants.length === 0) {
            $("#assistant").select2({
                placeholder: "No agents available",
                allowClear: true,
                width: "100%"
            });
        }
    } catch (error) {
        const errorDiv = document.getElementById("loadError");
        errorDiv.textContent = "❌ " + error.message;
        errorDiv.style.display = "block";
        $("#assistant").select2({
            placeholder: "Failed to load agents",
            allowClear: true,
            width: "100%"
        });
    }
}

// Excel range selection functions
async function useSelection() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("address");
            await context.sync();

            const address = range.address.split("!")[1]; // Remove sheet name if present
            document.getElementById("cellRange").value = address;
            updateRangeInfo(address);
        });
    } catch (error) {
        console.error("Selection error:", error);
        alert("Please select some cells first");
    }
}

async function useTargetSelection() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load("address");
            await context.sync();

            const address = range.address.split("!")[1]; // Remove sheet name if present
            // Extract just the column letter(s)
            const columnMatch = address.match(/^([A-Z]+)/);
            if (columnMatch) {
                document.getElementById("targetColumn").value = columnMatch[1];
            }
        });
    } catch (error) {
        console.error("Selection error:", error);
        alert("Please select a column first");
    }
}

async function updateRangeInfo(rangeNotation) {
    if (!rangeNotation) {
        document.getElementById("rangeInfo").textContent = "";
        document.getElementById("headerRowSection").style.display = "none";
        return;
    }

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(rangeNotation);
            range.load(["rowCount", "columnCount"]);
            await context.sync();

            const infoDiv = document.getElementById("rangeInfo");
            const headerSection = document.getElementById("headerRowSection");

            if (range.columnCount > 1) {
                infoDiv.textContent = `Selected: ${range.rowCount} rows × ${range.columnCount} columns`;
                headerSection.style.display = "block";
            } else {
                infoDiv.textContent = `Selected: ${range.rowCount} rows × ${range.columnCount} column`;
                headerSection.style.display = "none";
            }
        });
    } catch (error) {
        document.getElementById("rangeInfo").textContent = "Invalid range";
        document.getElementById("headerRowSection").style.display = "none";
    }
}

// Cancel processing
function cancelProcessing() {
    // Prevent multiple cancel clicks
    if (cancelRequested || processingProgress.status === "cancelled") {
        return;
    }

    cancelRequested = true;
    processingProgress.status = "cancelled";

    // Abort all ongoing requests
    if (currentAbortController) {
        currentAbortController.abort();
        currentAbortController = null;
    }

    // Clear the processing ID to prevent UI updates from cancelled requests
    currentProcessingId = null;

    document.getElementById("cancelBtn").style.display = "none";
    document.getElementById("cancelBtn").disabled = true;
    document.getElementById("submitBtn").disabled = false;
    document.getElementById("status").innerHTML = '⏹️ Processing cancelled';
    setTimeout(() => {
        document.getElementById("status").innerHTML = '';
        cancelRequested = false; // Reset for next run
    }, 3000);
}

// Process functions
async function handleSubmit(e) {
    e.preventDefault();

    const assistantSelect = document.getElementById("assistant");
    const cellRange = document.getElementById("cellRange");
    const targetColumn = document.getElementById("targetColumn").value;

    if (!assistantSelect.value) {
        alert("Please select an agent");
        return;
    }

    if (!cellRange.value) {
        alert("Please select input cells");
        return;
    }

    if (!/^[A-Za-z]+$/.test(targetColumn)) {
        alert("Please enter a valid target column letter (e.g., A, B, C)");
        return;
    }

    // Check actual processable row count and warn if over 100
    let actualRowsToProcess = 0;
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const range = sheet.getRange(cellRange.value);
            range.load(["values", "rowCount", "columnCount"]);
            await context.sync();

            const numColumns = range.columnCount;
            const headerRow = parseInt(document.getElementById("headerRow").value) || 1;

            // Count non-empty rows that will actually be processed
            for (let i = 0; i < range.rowCount; i++) {
                const actualRowNumber = range.rowIndex + i + 1; // 1-based row number

                // Skip header row if processing multiple columns
                if (numColumns > 1 && actualRowNumber === headerRow) {
                    continue;
                }

                // Check if row has any content
                const rowValues = range.values[i];
                let hasContent = false;

                if (numColumns === 1) {
                    hasContent = rowValues[0] && rowValues[0].toString().trim() !== "";
                } else {
                    // For multiple columns, check if any cell has content
                    for (let j = 0; j < numColumns; j++) {
                        if (rowValues[j] && rowValues[j].toString().trim() !== "") {
                            hasContent = true;
                            break;
                        }
                    }
                }

                if (hasContent) {
                    actualRowsToProcess++;
                }
            }

            if (actualRowsToProcess > 100) {
                const message = `You're about to process ${actualRowsToProcess} rows with content. Processing this many rows may take a while and could hit rate limits.\n\nAre you sure you want to continue?`;
                if (!confirm(message)) {
                    // Exit the entire function if user cancels
                    throw new Error("User cancelled processing");
                }
            }
        });
    } catch (error) {
        if (error.message === "User cancelled processing") {
            return; // Exit silently if user cancelled
        }
        console.error("Error checking row count:", error);
    }

    // Generate a unique ID for this processing run
    const processingId = Date.now() + '_' + Math.random();
    currentProcessingId = processingId;

    // Create new abort controller for this run
    currentAbortController = new AbortController();

    cancelRequested = false;
    isRateLimited = false;
    processingProgress.status = "idle";
    document.getElementById("submitBtn").disabled = true;
    document.getElementById("cancelBtn").style.display = "block";
    document.getElementById("cancelBtn").disabled = false;
    document.getElementById("status").innerHTML = '<div class="spinner"></div> Analyzing selection...';

    try {
        await processWithAssistant(
            assistantSelect.value,
            document.getElementById("instructions").value,
            cellRange.value,
            targetColumn,
            parseInt(document.getElementById("headerRow").value) || 1,
            processingId,
            currentAbortController.signal
        );

        // Only update UI if this is the current processing run
        if (processingId === currentProcessingId) {
            document.getElementById("submitBtn").disabled = false;
            document.getElementById("cancelBtn").style.display = "none";
            if (cancelRequested) {
                document.getElementById("status").innerHTML = '⏹️ Processing cancelled';
            } else {
                document.getElementById("status").innerHTML = '✅ Processing complete';
            }
            setTimeout(() => {
                if (processingId === currentProcessingId) {
                    document.getElementById("status").innerHTML = '';
                }
            }, 3000);
        }
    } catch (error) {
        // Only update UI if this is the current processing run
        if (processingId === currentProcessingId) {
            document.getElementById("submitBtn").disabled = false;
            document.getElementById("cancelBtn").style.display = "none";
            if (error.name !== 'AbortError') {
                document.getElementById("status").textContent = '❌ Error: ' + error.message;
            }
        }
    }
}

async function processWithAssistant(assistantId, instructions, rangeA1Notation, targetColumn, headerRow, processingId, abortSignal) {
    let BATCH_SIZE = 10;
    let BATCH_DELAY = 1000;
    let retryDelay = 1000; // Initial retry delay for rate limits

    const token = getFromStorage("accessToken");
    const workspaceId = getFromStorage("workspaceId");

    if (!token || !workspaceId) {
        throw new Error("Please connect your Dust account first");
    }

    try {
        await Excel.run(async (context) => {
            console.log("Starting Excel.run with range:", rangeA1Notation, "targetColumn:", targetColumn, "headerRow:", headerRow);

            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const selectedRange = sheet.getRange(rangeA1Notation);
            selectedRange.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
            await context.sync();

            console.log("Range loaded - rows:", selectedRange.rowCount, "cols:", selectedRange.columnCount,
                "startRow:", selectedRange.rowIndex, "startCol:", selectedRange.columnIndex);

            const selectedValues = selectedRange.values;
            const numColumns = selectedRange.columnCount;
            const numRows = selectedRange.rowCount;
            const startRow = selectedRange.rowIndex;
            const startCol = selectedRange.columnIndex;
            const targetColIndex = columnToIndex(targetColumn) - 1; // Convert to 0-based

            console.log("Target column index:", targetColIndex);

            // Get headers if multiple columns
            let headers = [];
            if (numColumns > 1) {
                // Calculate the header row index relative to the selection
                const headerRowIndex = headerRow - 1 - startRow;

                console.log("Header row index calculation: headerRow=", headerRow, "startRow=", startRow, "headerRowIndex=", headerRowIndex);

                if (headerRowIndex >= 0 && headerRowIndex < numRows) {
                    // Header is within the selected range
                    headers = selectedValues[headerRowIndex];
                    console.log("Headers from selection:", headers);
                } else if (headerRow > 0) {
                    // Header is outside the selected range, fetch it separately
                    console.log("Fetching headers from row", headerRow - 1, "cols", startCol, "to", startCol + numColumns - 1);
                    const headerRange = sheet.getRangeByIndexes(headerRow - 1, startCol, 1, numColumns);
                    headerRange.load("values");
                    await context.sync();
                    headers = headerRange.values[0];
                    console.log("Headers fetched separately:", headers);
                }
            }

            const cellsToProcess = [];

            for (let i = 0; i < numRows; i++) {
                const currentRow = startRow + i;
                const actualRowNumber = currentRow + 1; // 1-based row number for display

                // Skip header row if processing multiple columns
                if (numColumns > 1 && actualRowNumber === headerRow) {
                    console.log("Skipping header row at", actualRowNumber);
                    continue;
                }

                let inputContent = "";

                if (numColumns === 1) {
                    const inputValue = selectedValues[i][0];
                    if (!inputValue) {
                        const targetCell = sheet.getRangeByIndexes(currentRow, targetColIndex, 1, 1);
                        targetCell.values = [["No input value"]];
                        continue;
                    }
                    inputContent = inputValue.toString();
                } else {
                    const rowValues = selectedValues[i];
                    const contentParts = [];

                    for (let j = 0; j < numColumns; j++) {
                        const header = headers[j] || `Column ${j + 1}`;
                        const value = rowValues[j] || "";
                        contentParts.push(`${header}: ${value}`);
                    }

                    inputContent = contentParts.join("\n");

                    if (!inputContent.trim()) {
                        const targetCell = sheet.getRangeByIndexes(currentRow, targetColIndex, 1, 1);
                        targetCell.values = [["No input value"]];
                        continue;
                    }
                }

                cellsToProcess.push({
                    row: currentRow,
                    col: targetColIndex,
                    inputContent: inputContent
                });
            }

            const totalCells = cellsToProcess.length;
            // Reset and set processing status
            processingProgress = { current: 0, total: totalCells, status: "processing" };
            cancelRequested = false;
            updateProgressDisplay(processingId);

            console.log("Processing", totalCells, "cells");

            // Process in batches
            for (let i = 0; i < cellsToProcess.length; i += BATCH_SIZE) {
                // Check if cancellation was requested
                if (cancelRequested) {
                    console.log("Processing cancelled by user");
                    break;
                }

                const batch = cellsToProcess.slice(i, i + BATCH_SIZE);

                const promises = batch.map(async (item) => {
                    // Check if this processing was cancelled or replaced
                    if (processingId !== currentProcessingId) {
                        return;
                    }

                    // Retry logic for rate limits
                    let retries = 0;
                    const maxRetries = 3;
                    let lastError = null;

                    while (retries <= maxRetries && processingId === currentProcessingId && !cancelRequested) {
                        const payload = {
                            message: {
                                content: (instructions || "") + "\n\nInput:\n" + item.inputContent,
                                mentions: [{ configurationId: assistantId }],
                                context: {
                                    username: "excel",
                                    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
                                    fullName: "Excel User",
                                    email: "excel@dust.tt",
                                    profilePictureUrl: "",
                                    origin: "excel"
                                }
                            },
                            blocking: true,
                            title: "Excel Conversation",
                            visibility: "unlisted",
                            skipToolsValidation: true
                        };

                        try {
                            const apiPath = `/api/v1/w/${workspaceId}/assistant/conversations`;
                            const result = await callDustAPI(apiPath, {
                                method: "POST",
                                body: payload,
                                headers: {
                                    "Authorization": "Bearer " + token
                                },
                                signal: abortSignal
                            });
                            const content = result.conversation.content;

                            const lastAgentMessage = content.flat().reverse().find(msg => msg.type === "agent_message");

                            // Only update cell if this is still the current processing
                            if (processingId === currentProcessingId && !cancelRequested) {
                                const targetCell = sheet.getRangeByIndexes(item.row, item.col, 1, 1);
                                // Trim extra lines at the beginning and end of the agent answer
                                const agentContent = lastAgentMessage ? lastAgentMessage.content.trim() : "No response";
                                targetCell.values = [[agentContent]];
                                targetCell.format.fill.color = "#f0f9ff"; // Light blue background

                                // Sync immediately to update the cell in Excel
                                await context.sync();
                            }

                            // Success - exit retry loop
                            break;

                        } catch (error) {
                            lastError = error;

                            // Check if it's an abort error
                            if (error.name === 'AbortError') {
                                break;
                            }

                            // Check if it's a rate limit error (429) or contains rate limit message
                            const isRateLimit = error.message.includes('429') ||
                                error.message.toLowerCase().includes('rate limit') ||
                                error.message.toLowerCase().includes('too many requests');

                            if (isRateLimit && retries < maxRetries) {
                                retries++;
                                console.log(`Rate limit hit, retry ${retries}/${maxRetries} after ${retryDelay}ms`);

                                // Update status to show rate limiting
                                if (processingId === currentProcessingId) {
                                    isRateLimited = true;
                                    // Clear any existing timeout
                                    if (rateLimitTimeout) {
                                        clearTimeout(rateLimitTimeout);
                                    }
                                    // Keep the rate limit message for at least 5 seconds
                                    rateLimitTimeout = setTimeout(() => {
                                        if (processingId === currentProcessingId) {
                                            isRateLimited = false;
                                            updateProgressDisplay(processingId);
                                        }
                                    }, 5000);
                                    updateProgressDisplay(processingId);
                                }

                                // Wait with exponential backoff
                                await new Promise(resolve => setTimeout(resolve, retryDelay));
                                retryDelay = Math.min(retryDelay * 2, 30000); // Max 30 seconds

                                // Reduce batch size and increase delay for future batches
                                if (retries === 1) {
                                    BATCH_SIZE = Math.max(Math.floor(BATCH_SIZE / 2), 1);
                                    BATCH_DELAY = BATCH_DELAY * 2;
                                    console.log(`Adjusted batch size to ${BATCH_SIZE}, delay to ${BATCH_DELAY}ms`);
                                }
                            } else {
                                // Not a rate limit error or max retries reached
                                console.error("Error processing cell:", error);

                                // Only update cell with error if this is still the current processing
                                if (processingId === currentProcessingId && !cancelRequested) {
                                    const targetCell = sheet.getRangeByIndexes(item.row, item.col, 1, 1);
                                    targetCell.values = [["Error: " + error.message]];
                                    targetCell.format.fill.color = "#fee2e2"; // Light red background

                                    // Sync immediately to update the cell in Excel
                                    await context.sync();
                                }
                                break;
                            }
                        }
                    }


                    // Only increment progress if this is still the current processing
                    if (processingId === currentProcessingId && !cancelRequested && processingProgress.status !== "cancelled") {
                        processingProgress.current++;
                        updateProgressDisplay(processingId);
                    }
                });

                await Promise.all(promises);

                // Use the potentially adjusted BATCH_DELAY
                if (i + BATCH_SIZE < cellsToProcess.length) {
                    await new Promise(resolve => setTimeout(resolve, BATCH_DELAY));
                }
            }
        });
    } catch (error) {
        console.error("Excel.run error:", error);
        throw error;
    }
}

let isRateLimited = false;
let rateLimitTimeout = null;

function updateProgressDisplay(processingId) {
    // Only update if this is the current processing
    if (processingId !== currentProcessingId) {
        return;
    }

    const statusDiv = document.getElementById("status");
    if (processingProgress.status === "processing" && !cancelRequested) {
        if (isRateLimited) {
            statusDiv.innerHTML = `<div class="spinner"></div> Rate limited, slowing down... (${processingProgress.current}/${processingProgress.total})`;
        } else {
            statusDiv.innerHTML = `<div class="spinner"></div> Processing (${processingProgress.current}/${processingProgress.total})`;
        }
    } else if (processingProgress.status === "cancelled") {
        // Don't update the display if cancelled
        return;
    }
}

// Helper function to convert column letter to index (1-based)
function columnToIndex(column) {
    if (!column || typeof column !== "string") return null;

    column = column.toUpperCase();
    let sum = 0;

    for (let i = 0; i < column.length; i++) {
        sum *= 26;
        sum += column.charCodeAt(i) - "A".charCodeAt(0) + 1;
    }

    return sum;
}