/* global Office, Excel */

const DUST_VERSION = "0.1";
let processingProgress = { current: 0, total: 0, status: "idle" };

// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("selectCellsBtn").onclick = useSelection;
        document.getElementById("selectTargetBtn").onclick = useTargetSelection;
        document.getElementById("cellRange").addEventListener("input", function() {
            updateRangeInfo(this.value);
        });
        document.getElementById("myForm").addEventListener("submit", handleSubmit);
        document.getElementById("saveSetup").onclick = saveCredentials;
        document.getElementById("updateCredentialsBtn").onclick = showCredentialSetup;
        document.getElementById("removeCredentialsBtn").onclick = showRemoveConfirmation;
        document.getElementById("confirmRemove").onclick = removeCredentials;
        document.getElementById("cancelRemove").onclick = hideRemoveConfirmation;
        
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
            noResults: function() {
                return "No agents found";
            }
        }
    });
}

// Storage functions
function saveToStorage(key, value) {
    localStorage.setItem(`dust_excel_${key}`, value);
}

function getFromStorage(key) {
    return localStorage.getItem(`dust_excel_${key}`);
}

// Check credentials and initialize the appropriate view
function checkCredentialsAndInitialize() {
    const workspaceId = getFromStorage("workspaceId");
    const dustToken = getFromStorage("dustToken");
    
    if (workspaceId && dustToken) {
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
    document.getElementById("updateCredentialsBtn").style.display = "none";
    
    // Hide error message initially
    const errorDiv = document.getElementById("credentialError");
    if (errorDiv) {
        errorDiv.style.display = "none";
    }
    
    // Load existing credentials if any
    const existingWorkspace = getFromStorage("workspaceId");
    const existingToken = getFromStorage("dustToken");
    
    document.getElementById("workspaceId").value = existingWorkspace || "";
    document.getElementById("dustToken").value = existingToken || "";
    document.getElementById("region").value = getFromStorage("region") || "";
    
    // Show/hide remove button based on whether credentials exist
    const removeBtn = document.getElementById("removeCredentialsBtn");
    if (removeBtn) {
        removeBtn.style.display = (existingWorkspace || existingToken) ? "block" : "none";
    }
}

// Show main form
function showMainForm() {
    document.getElementById("credentialSetup").style.display = "none";
    document.getElementById("myForm").style.display = "block";
    document.getElementById("updateCredentialsBtn").style.display = "block";
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
    localStorage.removeItem("dust_excel_dustToken");
    localStorage.removeItem("dust_excel_region");
    localStorage.removeItem("dust_excel_credentialsConfigured");
    
    // Clear input fields
    document.getElementById("workspaceId").value = "";
    document.getElementById("dustToken").value = "";
    document.getElementById("region").value = "";
    
    // Hide remove button and confirmation
    document.getElementById("removeCredentialsBtn").style.display = "none";
    document.getElementById("removeConfirmation").style.display = "none";
    
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
    document.getElementById("updateCredentialsBtn").style.display = "none";
}

async function saveCredentials() {
    const workspaceId = document.getElementById("workspaceId").value;
    const dustToken = document.getElementById("dustToken").value;
    const region = document.getElementById("region").value;
    
    // Hide any previous error
    const errorDiv = document.getElementById("credentialError");
    if (errorDiv) {
        errorDiv.style.display = "none";
    }
    
    if (!workspaceId || !dustToken) {
        if (errorDiv) {
            errorDiv.textContent = "Please enter both Workspace ID and API Key";
            errorDiv.style.display = "block";
        }
        return;
    }
    
    // Show loading state
    const saveBtn = document.getElementById("saveSetup");
    const originalText = saveBtn.textContent;
    saveBtn.disabled = true;
    saveBtn.innerHTML = '<span class="spinner"></span> Validating...';
    
    try {
        // Temporarily save credentials to test them
        saveToStorage("workspaceId", workspaceId);
        saveToStorage("dustToken", dustToken);
        saveToStorage("region", region);
        
        // Test credentials by fetching agents using the proxy
        const apiPath = `/api/v1/w/${workspaceId}/assistant/agent_configurations`;
        const data = await callDustAPI(apiPath);
        
        // Credentials are valid, save the configured flag
        saveToStorage("credentialsConfigured", "true");
        
        // Hide error if it was showing
        if (errorDiv) {
            errorDiv.style.display = "none";
        }
        
        // Switch to main form
        showMainForm();
        loadAssistants();
        initializeSelect2();
    } catch (error) {
        // Remove invalid credentials
        localStorage.removeItem("dust_excel_workspaceId");
        localStorage.removeItem("dust_excel_dustToken");
        localStorage.removeItem("dust_excel_region");
        localStorage.removeItem("dust_excel_credentialsConfigured");
        
        // Show error message
        if (errorDiv) {
            errorDiv.textContent = "❌ Invalid credentials. Please check your Workspace ID and API Key.";
            errorDiv.style.display = "block";
        }
    } finally {
        saveBtn.disabled = false;
        saveBtn.textContent = originalText;
    }
}


// Dust API functions
function getDustBaseUrl() {
    const region = getFromStorage("region");
    if (region && region.toLowerCase() === "eu") {
        return "https://eu.dust.tt";
    }
    return "https://dust.tt";
}

async function loadAssistants() {
    const token = getFromStorage("dustToken");
    const workspaceId = getFromStorage("workspaceId");
    
    if (!token || !workspaceId) {
        const errorDiv = document.getElementById("loadError");
        errorDiv.textContent = "❌ Please configure Dust credentials first";
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
                noResults: function() {
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
    
    document.getElementById("submitBtn").disabled = true;
    document.getElementById("status").innerHTML = '<div class="spinner"></div> Analyzing selection...';
    
    try {
        await processWithAssistant(
            assistantSelect.value,
            document.getElementById("instructions").value,
            cellRange.value,
            targetColumn,
            parseInt(document.getElementById("headerRow").value) || 1
        );
        
        document.getElementById("submitBtn").disabled = false;
        document.getElementById("status").innerHTML = '✅ Processing complete';
        setTimeout(() => {
            document.getElementById("status").innerHTML = '';
        }, 3000);
    } catch (error) {
        document.getElementById("submitBtn").disabled = false;
        document.getElementById("status").textContent = '❌ Error: ' + error.message;
    }
}

async function processWithAssistant(assistantId, instructions, rangeA1Notation, targetColumn, headerRow) {
    const BATCH_SIZE = 10;
    const BATCH_DELAY = 1000;
    
    const token = getFromStorage("dustToken");
    const workspaceId = getFromStorage("workspaceId");
    
    if (!token || !workspaceId) {
        throw new Error("Please configure your Dust credentials first");
    }
    
    
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const selectedRange = sheet.getRange(rangeA1Notation);
        selectedRange.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
        await context.sync();
        
        const selectedValues = selectedRange.values;
        const numColumns = selectedRange.columnCount;
        const numRows = selectedRange.rowCount;
        const startRow = selectedRange.rowIndex;
        const targetColIndex = columnToIndex(targetColumn) - 1; // Convert to 0-based
        
        // Get headers if multiple columns
        let headers = [];
        if (numColumns > 1) {
            const headerRowIndex = headerRow - startRow - 1;
            if (headerRowIndex >= 0 && headerRowIndex < numRows) {
                headers = selectedValues[headerRowIndex];
            } else {
                const headerRange = sheet.getRange(headerRow, selectedRange.columnIndex + 1, 1, numColumns);
                headerRange.load("values");
                await context.sync();
                headers = headerRange.values[0];
            }
        }
        
        const cellsToProcess = [];
        
        for (let i = 0; i < numRows; i++) {
            const currentRow = startRow + i + 1; // Excel is 1-based
            
            if (numColumns > 1 && currentRow === headerRow) {
                continue;
            }
            
            let inputContent = "";
            
            if (numColumns === 1) {
                const inputValue = selectedValues[i][0];
                if (!inputValue) {
                    const targetCell = sheet.getCell(currentRow - 1, targetColIndex);
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
                    const targetCell = sheet.getCell(currentRow - 1, targetColIndex);
                    targetCell.values = [["No input value"]];
                    continue;
                }
            }
            
            cellsToProcess.push({
                row: currentRow - 1,
                col: targetColIndex,
                inputContent: inputContent
            });
        }
        
        const totalCells = cellsToProcess.length;
        processingProgress = { current: 0, total: totalCells, status: "processing" };
        updateProgressDisplay();
        
        // Process in batches
        for (let i = 0; i < cellsToProcess.length; i += BATCH_SIZE) {
            const batch = cellsToProcess.slice(i, i + BATCH_SIZE);
            
            const promises = batch.map(async (item) => {
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
                        }
                    });
                    const content = result.conversation.content;
                    
                    const lastAgentMessage = content.flat().reverse().find(msg => msg.type === "agent_message");
                    
                    const targetCell = sheet.getCell(item.row, item.col);
                    targetCell.values = [[lastAgentMessage ? lastAgentMessage.content : "No response"]];
                    
                    // Add note with conversation URL
                    const appUrl = `https://dust.tt/w/${workspaceId}/assistant/${result.conversation.sId}`;
                    targetCell.format.fill.color = "#f0f9ff"; // Light blue background
                    
                } catch (error) {
                    const targetCell = sheet.getCell(item.row, item.col);
                    targetCell.values = [["Error: " + error.message]];
                    targetCell.format.fill.color = "#fee2e2"; // Light red background
                }
                
                processingProgress.current++;
                updateProgressDisplay();
            });
            
            await Promise.all(promises);
            
            if (i + BATCH_SIZE < cellsToProcess.length) {
                await new Promise(resolve => setTimeout(resolve, BATCH_DELAY));
            }
        }
        
        await context.sync();
    });
}

function updateProgressDisplay() {
    const statusDiv = document.getElementById("status");
    if (processingProgress.status === "processing") {
        statusDiv.innerHTML = `<div class="spinner"></div> Processing (${processingProgress.current}/${processingProgress.total})`;
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