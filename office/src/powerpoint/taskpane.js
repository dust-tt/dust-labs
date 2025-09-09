/* global Office, PowerPoint */

const DUST_VERSION = "0.1";
let processingProgress = { current: 0, total: 0, status: "idle" };

// Initialize Office
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
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
    localStorage.setItem(`dust_powerpoint_${key}`, value);
}

function getFromStorage(key) {
    return localStorage.getItem(`dust_powerpoint_${key}`);
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
    localStorage.removeItem("dust_powerpoint_workspaceId");
    localStorage.removeItem("dust_powerpoint_dustToken");
    localStorage.removeItem("dust_powerpoint_region");
    localStorage.removeItem("dust_powerpoint_credentialsConfigured");
    
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
        localStorage.removeItem("dust_powerpoint_workspaceId");
        localStorage.removeItem("dust_powerpoint_dustToken");
        localStorage.removeItem("dust_powerpoint_region");
        localStorage.removeItem("dust_powerpoint_credentialsConfigured");
        
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

// Process functions
async function handleSubmit(e) {
    e.preventDefault();
    
    const assistantSelect = document.getElementById("assistant");
    const scope = document.querySelector('input[name="scope"]:checked').value;
    
    if (!assistantSelect.value) {
        alert("Please select an agent");
        return;
    }
    
    document.getElementById("submitBtn").disabled = true;
    document.getElementById("status").innerHTML = '<div class="spinner"></div> Extracting content...';
    
    try {
        await processWithAssistant(
            assistantSelect.value,
            document.getElementById("instructions").value,
            scope
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

async function processWithAssistant(assistantId, instructions, scope) {
    const BATCH_SIZE = 10;
    const BATCH_DELAY = 1000;
    
    const token = getFromStorage("dustToken");
    const workspaceId = getFromStorage("workspaceId");
    
    if (!token || !workspaceId) {
        throw new Error("Please configure your Dust credentials first");
    }
    
    await PowerPoint.run(async (context) => {
        // Extract content based on scope
        let content = "";
        let targetSlides = [];
        
        if (scope === "presentation") {
            const presentation = context.presentation;
            presentation.slides.load("items");
            await context.sync();
            targetSlides = presentation.slides.items;
        } else if (scope === "slide") {
            const selectedSlides = context.presentation.getSelectedSlides();
            selectedSlides.load("items");
            await context.sync();
            targetSlides = selectedSlides.items;
        } else if (scope === "selection") {
            // Handle selection differently - just extract and process
            const selectedShapes = context.presentation.getSelectedShapes();
            selectedShapes.load("items");
            await context.sync();
            
            if (selectedShapes.items.length > 0) {
                content = "--- Selected Content ---\n";
                for (let shape of selectedShapes.items) {
                    shape.textFrame.load("textRange");
                    await context.sync();
                    
                    if (shape.textFrame && shape.textFrame.textRange) {
                        const text = shape.textFrame.textRange.text;
                        if (text && text.trim()) {
                            content += text + "\n";
                        }
                    }
                }
            } else {
                const selectedTextRange = context.presentation.getSelectedTextRange();
                selectedTextRange.load("text");
                await context.sync();
                
                if (selectedTextRange.text) {
                    content = "--- Selected Text ---\n" + selectedTextRange.text;
                }
            }
            
            // Process the selection content
            if (!content || content.trim() === "") {
                throw new Error("No content found in selection");
            }
            
            const inputContent = (instructions || "") + "\n\nContent:\n" + content;
            
            // Call Dust API
            const payload = {
                message: {
                    content: inputContent,
                    mentions: [{ configurationId: assistantId }],
                    context: {
                        username: "powerpoint",
                        timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
                        fullName: "PowerPoint User",
                        email: "powerpoint@dust.tt",
                        profilePictureUrl: "",
                        origin: "powerpoint"
                    }
                },
                blocking: true,
                title: "PowerPoint Conversation",
                visibility: "unlisted",
                skipToolsValidation: true
            };
            
            document.getElementById("status").innerHTML = '<div class="spinner"></div> Processing with agent...';
            
            const apiPath = `/api/v1/w/${workspaceId}/assistant/conversations`;
            const result = await callDustAPI(apiPath, {
                method: "POST",
                body: payload,
                headers: {
                    "Authorization": "Bearer " + token
                }
            });
            
            const messages = result.conversation.content;
            const lastAgentMessage = messages.flat().reverse().find(msg => msg.type === "agent_message");
            
            if (lastAgentMessage && lastAgentMessage.content) {
                // For selection, we'll replace the selected content
                if (selectedShapes.items.length > 0) {
                    const shape = selectedShapes.items[0];
                    shape.textFrame.textRange.text = lastAgentMessage.content;
                } else {
                    selectedTextRange.text = lastAgentMessage.content;
                }
                await context.sync();
            }
            
            return; // Exit early for selection processing
        }
        
        // Process slides for presentation and slide scopes
        if (targetSlides.length > 0) {
            for (let i = 0; i < targetSlides.length; i++) {
                const slide = targetSlides[i];
                slide.shapes.load("items");
                await context.sync();
                
                content += `\n--- Slide ${i + 1} ---\n`;
                
                for (let shape of slide.shapes.items) {
                    if (shape.type === "GeometricShape" || shape.type === "TextBox") {
                        shape.textFrame.load("textRange");
                        await context.sync();
                        
                        if (shape.textFrame && shape.textFrame.textRange) {
                            const text = shape.textFrame.textRange.text;
                            if (text && text.trim()) {
                                content += text + "\n";
                            }
                        }
                    }
                }
            }
            
            if (!content || content.trim() === "") {
                throw new Error("No content found in slides");
            }
            
            const inputContent = (instructions || "") + "\n\nPresentation content:\n" + content;
            
            // Call Dust API
            const payload = {
                message: {
                    content: inputContent,
                    mentions: [{ configurationId: assistantId }],
                    context: {
                        username: "powerpoint",
                        timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
                        fullName: "PowerPoint User",
                        email: "powerpoint@dust.tt",
                        profilePictureUrl: "",
                        origin: "powerpoint"
                    }
                },
                blocking: true,
                title: "PowerPoint Conversation",
                visibility: "unlisted",
                skipToolsValidation: true
            };
            
            document.getElementById("status").innerHTML = '<div class="spinner"></div> Processing with agent...';
            
            const apiPath = `/api/v1/w/${workspaceId}/assistant/conversations`;
            const result = await callDustAPI(apiPath, {
                method: "POST",
                body: payload,
                headers: {
                    "Authorization": "Bearer " + token
                }
            });
            
            const messages = result.conversation.content;
            const lastAgentMessage = messages.flat().reverse().find(msg => msg.type === "agent_message");
            
            if (lastAgentMessage && lastAgentMessage.content) {
                // For now, add the result as a new slide at the end
                const newSlide = context.presentation.slides.add();
                newSlide.load("shapes");
                await context.sync();
                
                // Add a text box with the result
                const textBox = newSlide.shapes.addTextBox(lastAgentMessage.content, {
                    left: 50,
                    top: 50,
                    height: 400,
                    width: 600
                });
                
                await context.sync();
            }
        }
    });
}

function updateProgressDisplay() {
    const statusDiv = document.getElementById("status");
    if (processingProgress.status === "processing") {
        statusDiv.innerHTML = `<div class="spinner"></div> Processing (${processingProgress.current}/${processingProgress.total})`;
    }
}