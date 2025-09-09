Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("saveCredentials").onclick = saveCredentials;
        document.getElementById("changeUser").onclick = changeUser;
        document.getElementById("runAgent").onclick = runAgent;
        document.getElementById("applyResults").onclick = applyResults;
        document.getElementById("discardResults").onclick = discardResults;
        document.getElementById("cancelRun").onclick = cancelRun;
        
        checkCredentials();
    }
});

let currentConversationId = null;
let currentResults = null;
let isRunning = false;

async function checkCredentials() {
    const settings = Office.context.document.settings;
    const workspaceId = settings.get("dustWorkspaceId");
    const token = settings.get("dustApiKey");
    const region = settings.get("dustRegion") || "";
    
    if (workspaceId && token) {
        document.getElementById("credentialSetup").style.display = "none";
        document.getElementById("mainPanel").style.display = "block";
        document.getElementById("userWorkspace").textContent = workspaceId;
        
        await loadAgents(workspaceId, token, region);
    } else {
        document.getElementById("credentialSetup").style.display = "block";
        document.getElementById("mainPanel").style.display = "none";
    }
}

async function saveCredentials() {
    const workspaceId = document.getElementById("workspaceId").value.trim();
    const token = document.getElementById("dustToken").value.trim();
    const region = document.getElementById("regionSelect").value;
    
    if (!workspaceId || !token) {
        showCredentialError("Please enter both Workspace ID and API Key");
        return;
    }
    
    try {
        const settings = Office.context.document.settings;
        settings.set("dustWorkspaceId", workspaceId);
        settings.set("dustApiKey", token);
        settings.set("dustRegion", region);
        
        await settings.saveAsync();
        
        document.getElementById("credentialSetup").style.display = "none";
        document.getElementById("mainPanel").style.display = "block";
        document.getElementById("userWorkspace").textContent = workspaceId;
        
        await loadAgents(workspaceId, token, region);
    } catch (error) {
        showCredentialError("Failed to save credentials: " + error.message);
    }
}

function changeUser() {
    const settings = Office.context.document.settings;
    settings.remove("dustWorkspaceId");
    settings.remove("dustApiKey");
    settings.remove("dustRegion");
    
    settings.saveAsync(() => {
        document.getElementById("workspaceId").value = "";
        document.getElementById("dustToken").value = "";
        document.getElementById("regionSelect").value = "";
        
        document.getElementById("credentialSetup").style.display = "block";
        document.getElementById("mainPanel").style.display = "none";
        document.getElementById("agentSelect").innerHTML = '<option value="">Loading agents...</option>';
    });
}

async function loadAgents(workspaceId, token, region) {
    const agentSelect = document.getElementById("agentSelect");
    
    try {
        const baseUrl = region === "eu" ? "https://eu.dust.tt" : "https://dust.tt";
        const response = await fetch(`${baseUrl}/api/v1/w/${workspaceId}/assistant/agent_configurations`, {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            throw new Error(`Failed to load agents: ${response.statusText}`);
        }
        
        const data = await response.json();
        const agents = data.agentConfigurations || [];
        
        $(agentSelect).empty();
        
        if (agents.length === 0) {
            $(agentSelect).append('<option value="">No agents available</option>');
        } else {
            $(agentSelect).append('<option value="">Select an agent...</option>');
            agents.forEach(agent => {
                $(agentSelect).append(`<option value="${agent.sId}">${agent.name}</option>`);
            });
        }
        
        $(agentSelect).select2({
            placeholder: "Select an agent",
            allowClear: false,
            minimumResultsForSearch: 5
        });
        
    } catch (error) {
        showError("Failed to load agents: " + error.message);
        $(agentSelect).html('<option value="">Failed to load agents</option>');
    }
}

async function runAgent() {
    const agentId = document.getElementById("agentSelect").value;
    if (!agentId) {
        showError("Please select an agent");
        return;
    }
    
    const scope = document.querySelector('input[name="scope"]:checked').value;
    const instructions = document.getElementById("instructionsInput").value.trim();
    
    try {
        isRunning = true;
        showProgress("Extracting content from presentation...");
        document.getElementById("runAgent").style.display = "none";
        document.getElementById("cancelRun").style.display = "inline-block";
        
        const content = await extractContent(scope);
        
        if (!content || content.trim() === "") {
            throw new Error("No content found to process");
        }
        
        showProgress("Sending to Dust agent...");
        
        const settings = Office.context.document.settings;
        const workspaceId = settings.get("dustWorkspaceId");
        const token = settings.get("dustApiKey");
        const region = settings.get("dustRegion") || "";
        
        const message = instructions 
            ? `${instructions}\n\nPresentation content:\n${content}` 
            : `Process this presentation content:\n${content}`;
        
        const result = await callDustAgent(workspaceId, token, region, agentId, message);
        
        if (!isRunning) {
            showProgress("");
            return;
        }
        
        currentResults = result;
        displayResults(result);
        
    } catch (error) {
        showError("Error: " + error.message);
    } finally {
        isRunning = false;
        document.getElementById("runAgent").style.display = "inline-block";
        document.getElementById("cancelRun").style.display = "none";
        hideProgress();
    }
}

async function extractContent(scope) {
    return new Promise((resolve, reject) => {
        PowerPoint.run(async (context) => {
            let content = "";
            
            try {
                if (scope === "presentation") {
                    const presentation = context.presentation;
                    presentation.slides.load("items");
                    await context.sync();
                    
                    for (let i = 0; i < presentation.slides.items.length; i++) {
                        const slide = presentation.slides.items[i];
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
                    
                } else if (scope === "slide") {
                    const selectedSlides = context.presentation.getSelectedSlides();
                    selectedSlides.load("items");
                    await context.sync();
                    
                    if (selectedSlides.items.length > 0) {
                        const slide = selectedSlides.items[0];
                        slide.shapes.load("items");
                        await context.sync();
                        
                        content = "--- Current Slide ---\n";
                        
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
                    
                } else if (scope === "selection") {
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
                }
                
                resolve(content);
            } catch (error) {
                reject(error);
            }
        });
    });
}

async function callDustAgent(workspaceId, token, region, agentId, message) {
    const baseUrl = region === "eu" ? "https://eu.dust.tt" : "https://dust.tt";
    
    const response = await fetch(`${baseUrl}/api/v1/w/${workspaceId}/assistant/agent_configurations/${agentId}/conversations`, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            message: {
                content: message,
                mentions: [],
                context: {
                    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
                    profilePictureUrl: null
                }
            },
            title: null,
            visibility: "unlisted"
        })
    });
    
    if (!response.ok) {
        throw new Error(`Failed to create conversation: ${response.statusText}`);
    }
    
    const data = await response.json();
    currentConversationId = data.conversation.sId;
    
    showProgress("Processing with agent...");
    
    let attempts = 0;
    const maxAttempts = 60;
    
    while (attempts < maxAttempts && isRunning) {
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        const convResponse = await fetch(`${baseUrl}/api/v1/w/${workspaceId}/assistant/conversations/${currentConversationId}`, {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!convResponse.ok) {
            throw new Error(`Failed to get conversation: ${convResponse.statusText}`);
        }
        
        const convData = await convResponse.json();
        const messages = convData.conversation.content;
        
        if (messages && messages.length > 1) {
            const lastMessage = messages[messages.length - 1];
            if (lastMessage.type === "agent_message" && lastMessage.status === "succeeded") {
                return lastMessage.content;
            }
        }
        
        attempts++;
    }
    
    throw new Error("Agent did not respond in time");
}

async function applyResults() {
    if (!currentResults) {
        showError("No results to apply");
        return;
    }
    
    const scope = document.querySelector('input[name="scope"]:checked').value;
    
    try {
        await PowerPoint.run(async (context) => {
            if (scope === "selection") {
                const selectedShapes = context.presentation.getSelectedShapes();
                selectedShapes.load("items");
                await context.sync();
                
                if (selectedShapes.items.length > 0) {
                    const shape = selectedShapes.items[0];
                    shape.textFrame.textRange.text = currentResults;
                    await context.sync();
                } else {
                    const selectedTextRange = context.presentation.getSelectedTextRange();
                    selectedTextRange.text = currentResults;
                    await context.sync();
                }
                
            } else if (scope === "slide") {
                const selectedSlides = context.presentation.getSelectedSlides();
                selectedSlides.load("items");
                await context.sync();
                
                if (selectedSlides.items.length > 0) {
                    const slide = selectedSlides.items[0];
                    
                    slide.shapes.load("items");
                    await context.sync();
                    
                    for (let shape of slide.shapes.items) {
                        shape.delete();
                    }
                    
                    const textBox = slide.shapes.addTextBox(currentResults, {
                        left: 50,
                        top: 50,
                        height: 400,
                        width: 600
                    });
                    
                    await context.sync();
                }
                
            } else if (scope === "presentation") {
                showError("Cannot automatically apply results to entire presentation. Please apply to specific slides.");
                return;
            }
            
            currentResults = null;
            document.getElementById("resultsSection").style.display = "none";
            showSuccess("Results applied successfully!");
            
        });
    } catch (error) {
        showError("Failed to apply results: " + error.message);
    }
}

function displayResults(content) {
    document.getElementById("resultsContent").innerHTML = `<pre>${escapeHtml(content)}</pre>`;
    document.getElementById("resultsSection").style.display = "block";
    hideProgress();
}

function discardResults() {
    currentResults = null;
    document.getElementById("resultsSection").style.display = "none";
    document.getElementById("resultsContent").innerHTML = "";
}

function cancelRun() {
    isRunning = false;
    currentConversationId = null;
    hideProgress();
    document.getElementById("runAgent").style.display = "inline-block";
    document.getElementById("cancelRun").style.display = "none";
}

function showProgress(message) {
    document.getElementById("progressSection").style.display = "block";
    document.getElementById("progressMessage").textContent = message;
    document.getElementById("errorSection").style.display = "none";
}

function hideProgress() {
    document.getElementById("progressSection").style.display = "none";
    document.getElementById("progressMessage").textContent = "";
}

function showError(message) {
    document.getElementById("errorSection").style.display = "block";
    document.getElementById("errorMessage").textContent = message;
    setTimeout(() => {
        document.getElementById("errorSection").style.display = "none";
    }, 5000);
}

function showSuccess(message) {
    document.getElementById("errorSection").style.display = "block";
    document.getElementById("errorMessage").textContent = message;
    document.getElementById("errorMessage").style.backgroundColor = "#d4edda";
    document.getElementById("errorMessage").style.color = "#155724";
    setTimeout(() => {
        document.getElementById("errorSection").style.display = "none";
        document.getElementById("errorMessage").style.backgroundColor = "";
        document.getElementById("errorMessage").style.color = "";
    }, 3000);
}

function showCredentialError(message) {
    document.getElementById("credentialError").textContent = message;
    document.getElementById("credentialError").style.display = "block";
    setTimeout(() => {
        document.getElementById("credentialError").style.display = "none";
    }, 5000);
}

function escapeHtml(text) {
    const map = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
}