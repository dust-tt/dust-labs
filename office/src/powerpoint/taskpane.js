/* global Office, PowerPoint */

const DUST_VERSION = "0.1";
let processingProgress = { current: 0, total: 0, status: "idle" };
let processingCancelled = false;

// Initialize Office
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("myForm").addEventListener("submit", handleSubmit);
    document.getElementById("saveSetup").onclick = saveCredentials;
    document.getElementById("updateCredentialsBtn").onclick =
      showCredentialSetup;
    document.getElementById("removeCredentialsBtn").onclick =
      showRemoveConfirmation;
    document.getElementById("confirmRemove").onclick = removeCredentials;
    document.getElementById("cancelRemove").onclick = hideRemoveConfirmation;
    document.getElementById("cancelBtn").onclick = cancelProcessing;

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
      },
    },
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
    removeBtn.style.display =
      existingWorkspace || existingToken ? "block" : "none";
  }
}

// Show main form
function showMainForm() {
  document.getElementById("credentialSetup").style.display = "none";
  document.getElementById("myForm").style.display = "block";
  document.getElementById("updateCredentialsBtn").style.display = "inline-block";
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
    width: "100%",
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
      errorDiv.textContent =
        "❌ Invalid credentials. Please check your Workspace ID and API Key.";
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
      width: "100%",
    });
    return;
  }

  try {
    const apiPath = `/api/v1/w/${workspaceId}/assistant/agent_configurations`;
    const data = await callDustAPI(apiPath);
    const assistants = data.agentConfigurations;

    const sortedAssistants = assistants.sort((a, b) =>
      a.name.localeCompare(b.name)
    );

    const select = document.getElementById("assistant");
    select.innerHTML = "";

    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    select.appendChild(emptyOption);

    sortedAssistants.forEach((a) => {
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
        },
      },
    });

    if (assistants.length === 0) {
      $("#assistant").select2({
        placeholder: "No agents available",
        allowClear: true,
        width: "100%",
      });
    }
  } catch (error) {
    const errorDiv = document.getElementById("loadError");
    errorDiv.textContent = "❌ " + error.message;
    errorDiv.style.display = "block";
    $("#assistant").select2({
      placeholder: "Failed to load agents",
      allowClear: true,
      width: "100%",
    });
  }
}

// Helper function to safely extract text from shapes
// Returns an array of objects with slideIndex, shapeId, and originalText
async function safeExtractTextFromShapes(context, shapes, slideIndexOffset = 0) {
  const textBlocks = [];
  
  if (!shapes || shapes.length === 0) {
    return textBlocks;
  }
  
  // First, load shape types to identify which shapes can have text
  for (let item of shapes) {
    item.shape.load("type,id");
  }
  await context.sync();
  
  // Filter to only shapes that can have text
  // Shape types that support text: GeometricShape, TextBox, Placeholder
  const textCapableShapes = shapes.filter(item => {
    const shapeType = item.shape.type;
    // These are the shape types that typically have textFrame
    return shapeType === "GeometricShape" || 
           shapeType === "TextBox" || 
           shapeType === "Placeholder" ||
           shapeType === "Group" ||
           !shapeType; // Sometimes type is undefined but shape has text
  });
  
  console.log(`[safeExtractText] Found ${textCapableShapes.length} text-capable shapes out of ${shapes.length} total shapes`);
  
  if (textCapableShapes.length === 0) {
    return textBlocks;
  }
  
  // Try to load textFrame.hasText for text-capable shapes
  const shapesWithTextFrame = [];
  for (let item of textCapableShapes) {
    try {
      item.shape.textFrame.load("hasText");
      shapesWithTextFrame.push(item);
    } catch (e) {
      // Even text-capable shapes might not have textFrame in some cases
      console.log(`Shape ${item.shape.id} doesn't have accessible textFrame`);
    }
  }
  
  if (shapesWithTextFrame.length === 0) {
    return textBlocks;
  }
  
  // Sync to get hasText property
  await context.sync();
  
  // Now load text content for shapes that have text
  const shapesWithText = [];
  for (let item of shapesWithTextFrame) {
    try {
      if (item.shape.textFrame && item.shape.textFrame.hasText) {
        item.shape.textFrame.load("textRange/text");
        shapesWithText.push(item);
      }
    } catch (e) {
      console.log(`Shape ${item.shape.id} hasText check failed`);
    }
  }
  
  if (shapesWithText.length === 0) {
    return textBlocks;
  }
  
  console.log(`[safeExtractText] Found ${shapesWithText.length} shapes with actual text`);
  await context.sync();
  
  // Collect text blocks
  for (let item of shapesWithText) {
    try {
      const text = item.shape.textFrame.textRange.text;
      if (text && text.trim()) {
        textBlocks.push({
          slideIndex: item.slideIndex + slideIndexOffset,
          shapeId: item.shape.id,
          originalText: text
        });
      }
    } catch (e) {
      console.log(`Error getting text from shape ${item.shape.id}:`, e.message);
    }
  }
  
  return textBlocks;
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

  // If processing entire presentation, count text blocks first
  if (scope === "presentation") {
    document.getElementById("submitBtn").disabled = true;
    
    try {
      const textBlockCount = await countPresentationTextBlocks();
      
      if (textBlockCount === 0) {
        document.getElementById("submitBtn").disabled = false;
        document.getElementById("status").textContent = "❌ No text content found to process";
        return;
      }
      
      // Only show confirmation dialog if more than 100 text blocks
      if (textBlockCount > 100) {
        const confirmed = await showConfirmationDialog(textBlockCount);
        
        if (!confirmed) {
          document.getElementById("submitBtn").disabled = false;
          document.getElementById("status").innerHTML = "";
          return;
        }
      }
    } catch (error) {
      document.getElementById("submitBtn").disabled = false;
      document.getElementById("status").textContent = "❌ Error: " + error.message;
      return;
    }
  } else {
    document.getElementById("submitBtn").disabled = true;
  }

  document.getElementById("status").innerHTML =
    '<div class="spinner"></div> Extracting content...';

  // Show cancel button, hide submit button
  document.getElementById("submitBtn").style.display = "none";
  document.getElementById("cancelBtn").style.display = "block";
  processingCancelled = false;

  try {
    await processWithAssistant(
      assistantSelect.value,
      document.getElementById("instructions").value,
      scope
    );

    if (!processingCancelled) {
      document.getElementById("status").innerHTML = "✅ Processing complete";
      setTimeout(() => {
        document.getElementById("status").innerHTML = "";
      }, 3000);
    }
  } catch (error) {
    if (!processingCancelled) {
      // Show detailed error information
      const errorDetails = `
        <div style="color: red; font-size: 12px;">
          <strong>❌ Error: ${error.message}</strong><br>
          <div style="font-size: 10px; margin-top: 5px; padding: 5px; background: #f5f5f5; border-radius: 3px; color: #333;">
            <strong>Type:</strong> ${error.name || 'Unknown'}<br>
            <strong>Debug Info:</strong> ${error.debugInfo ? JSON.stringify(error.debugInfo) : 'None'}<br>
            <strong>Stack:</strong><br>
            <pre style="margin: 0; font-size: 9px; overflow-x: auto; white-space: pre-wrap;">${error.stack || 'No stack trace'}</pre>
          </div>
        </div>
      `;
      document.getElementById("status").innerHTML = errorDetails;
    }
  } finally {
    // Always reset buttons and cancellation flag
    processingCancelled = false;
    document.getElementById("submitBtn").disabled = false;
    document.getElementById("submitBtn").style.display = "block";
    document.getElementById("cancelBtn").style.display = "none";
  }
}

async function processWithAssistant(assistantId, instructions, scope) {
  const BATCH_SIZE = 5; // Process 5 text blocks at a time
  const BATCH_DELAY = 2000; // 2 second delay between batches

  const token = getFromStorage("dustToken");
  const workspaceId = getFromStorage("workspaceId");

  if (!token || !workspaceId) {
    throw new Error("Please configure your Dust credentials first");
  }

  let processedResults = [];
  
  console.log(`[ProcessWithAssistant] Starting processing with scope: ${scope}`);
  
  try {
    await PowerPoint.run(async (context) => {
      let textBlocksToProcess = [];
    // Store only metadata, not PowerPoint objects
    // Structure: { slideIndex, shapeId, originalText }
    console.log('[ProcessWithAssistant] Initializing text blocks collection');

    if (scope === "presentation") {
      console.log('[ProcessWithAssistant] Processing entire presentation');
      // Get all slides in the presentation
      const presentation = context.presentation;
      presentation.slides.load("items");
      await context.sync();
      
      const totalSlides = presentation.slides.items.length;
      console.log(`[ProcessWithAssistant] Found ${totalSlides} slides`);
      document.getElementById("status").innerHTML = `<div class="spinner"></div> Loading ${totalSlides} slides...`;

      // Load all shapes from all slides at once
      for (let slide of presentation.slides.items) {
        slide.shapes.load("items");
      }
      await context.sync();
      console.log('[ProcessWithAssistant] Loaded all slide shapes');
      
      // Count total shapes
      let totalShapes = 0;
      for (let slide of presentation.slides.items) {
        totalShapes += slide.shapes.items.length;
      }
      
      document.getElementById("status").innerHTML = `<div class="spinner"></div> Checking ${totalShapes} shapes across ${totalSlides} slides...`;

      // Prepare shapes for safe extraction
      const shapesToCheck = [];
      for (let slideIndex = 0; slideIndex < presentation.slides.items.length; slideIndex++) {
        const slide = presentation.slides.items[slideIndex];
        for (let shape of slide.shapes.items) {
          shapesToCheck.push({ shape, slideIndex });
        }
      }
      
      console.log(`[ProcessWithAssistant] Checking ${shapesToCheck.length} shapes for text`);
      
      // Use the safe extraction function
      const extractedBlocks = await safeExtractTextFromShapes(context, shapesToCheck);
      textBlocksToProcess = extractedBlocks;
      
      console.log(`[ProcessWithAssistant] Extracted ${textBlocksToProcess.length} text blocks from presentation`);
      
      if (textBlocksToProcess.length > 0) {
        document.getElementById("status").innerHTML = `<div class="spinner"></div> Processing ${textBlocksToProcess.length} text blocks...`;
      }
    } else if (scope === "slide") {
      console.log('[ProcessWithAssistant] Processing current slide');
      document.getElementById("status").innerHTML = `<div class="spinner"></div> Scanning current slide...`;
      
      // Get the currently active slide
      let targetSlides = [];

      // For PowerPoint Online, getSelectedSlides API is not available
      // We'll just use the first slide as the "current" slide
      console.log("[ProcessWithAssistant] Using first slide for 'current slide' mode");

      // If no selected slides, use the first slide as fallback
      if (targetSlides.length === 0) {
        console.log('[ProcessWithAssistant] No selected slides, using first slide as fallback');
        const presentation = context.presentation;
        presentation.slides.load("items");
        await context.sync();
        if (presentation.slides.items.length > 0) {
          targetSlides = [presentation.slides.items[0]];
          console.log('[ProcessWithAssistant] Using first slide as target');
        }
      }

      // Load all shapes from target slides
      console.log('[ProcessWithAssistant] Loading shapes from target slides');
      for (let slide of targetSlides) {
        slide.shapes.load("items");
      }
      await context.sync();
      console.log('[ProcessWithAssistant] Shapes loaded successfully');

      // Prepare shapes for safe extraction
      const shapesToCheck = [];
      let slideIdx = 0;
      for (let slide of targetSlides) {
        for (let shape of slide.shapes.items) {
          shapesToCheck.push({ shape, slideIndex: slideIdx });
        }
        slideIdx++;
      }
      
      console.log(`[ProcessWithAssistant] Checking ${shapesToCheck.length} shapes for text`);
      
      // Use the safe extraction function
      const extractedBlocks = await safeExtractTextFromShapes(context, shapesToCheck);
      
      // Mark as slide scope for later processing
      for (let block of extractedBlocks) {
        block.isSlideScope = true;
      }
      
      textBlocksToProcess = extractedBlocks;
      console.log(`[ProcessWithAssistant] Extracted ${textBlocksToProcess.length} text blocks from current slide`);
      
      if (textBlocksToProcess.length > 0) {
        document.getElementById("status").innerHTML = `<div class="spinner"></div> Processing ${textBlocksToProcess.length} text blocks...`;
      }
    } else if (scope === "selection") {
      console.log('[ProcessWithAssistant] Processing selection');
      document.getElementById("status").innerHTML = `<div class="spinner"></div> Scanning selected text...`;
      
      // Handle selected shapes/text
      console.log('[ProcessWithAssistant] Getting selected shapes');
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();
      console.log(`[ProcessWithAssistant] Found ${selectedShapes.items.length} selected shapes`);

      if (selectedShapes.items.length > 0) {
        // Prepare shapes for safe extraction
        const shapesToCheck = [];
        for (let i = 0; i < selectedShapes.items.length; i++) {
          shapesToCheck.push({ shape: selectedShapes.items[i], slideIndex: 0 });
        }
        
        // Use the safe extraction function
        const extractedBlocks = await safeExtractTextFromShapes(context, shapesToCheck);
        
        // Mark as selection for later processing
        for (let block of extractedBlocks) {
          block.isSelection = true;
          block.slideIndex = null; // Clear slideIndex for selections
        }
        
        textBlocksToProcess = extractedBlocks;
      } else {
        // Check for selected text range
        const selectedTextRange = context.presentation.getSelectedTextRange();
        selectedTextRange.load("text");
        await context.sync();

        if (selectedTextRange.text && selectedTextRange.text.trim()) {
          // For selected text, we can't update it later, so just note it
          textBlocksToProcess.push({
            slideIndex: null,
            shapeId: null,
            originalText: selectedTextRange.text,
            isSelectedText: true
          });
        }
      }
      
      console.log(`[ProcessWithAssistant] Extracted ${textBlocksToProcess.length} text blocks from selection`);
    }

    // Check if we have text blocks to process
    if (textBlocksToProcess.length === 0) {
      console.log('[ProcessWithAssistant] No text blocks found to process');
      throw new Error("No text content found to process");
    }
    console.log(`[ProcessWithAssistant] Total text blocks to process: ${textBlocksToProcess.length}`);

    // Warn if processing more than 100 text blocks
    if (textBlocksToProcess.length > 100) {
      const message = `You're about to process ${textBlocksToProcess.length} text blocks. Processing this many blocks may take a while and could hit rate limits.\n\nAre you sure you want to continue?`;
      if (!confirm(message)) {
        throw new Error("Processing cancelled by user");
      }
    }

    // First, process all text blocks with the API and collect results
    const totalBlocks = textBlocksToProcess.length;
    let processedCount = 0;

    document.getElementById(
      "status"
    ).innerHTML = `<div class="spinner"></div> Processing ${totalBlocks} text block(s)...`;

    // Process in batches to avoid rate limiting
    for (let i = 0; i < textBlocksToProcess.length; i += BATCH_SIZE) {
      // Check if processing was cancelled
      if (processingCancelled) {
        document.getElementById("status").innerHTML = "Processing cancelled";
        return; // Exit the entire function immediately
      }
      
      const batch = textBlocksToProcess.slice(
        i,
        Math.min(i + BATCH_SIZE, textBlocksToProcess.length)
      );

      // Process each text block in the batch
      for (let textBlock of batch) {
        // Check cancellation for each item
        if (processingCancelled) {
          // Clean up and exit immediately
          document.getElementById("status").innerHTML = "Processing cancelled";
          return; // Exit the entire function
        }
        
        try {
          // Double-check cancellation before making API call
          if (processingCancelled) {
            console.log("Processing cancelled before API call");
            document.getElementById("status").innerHTML = "Processing cancelled";
            return;
          }
          
          // Prepare the content for the API
          const inputContent =
            (instructions || "Process this text:") +
            "\n\n" +
            textBlock.originalText;

          // Call Dust API for this text block
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
                origin: "powerpoint",
              },
            },
            blocking: true,
            title: "PowerPoint Conversation",
            visibility: "unlisted",
            skipToolsValidation: true,
          };

          const apiPath = `/api/v1/w/${workspaceId}/assistant/conversations`;
          const result = await callDustAPI(apiPath, {
            method: "POST",
            body: payload,
            headers: {
              Authorization: "Bearer " + token,
            },
          });

          // Check if cancelled after API call
          if (processingCancelled) {
            console.log("Processing cancelled after API call");
            document.getElementById("status").innerHTML = "Processing cancelled";
            return;
          }
          
          const messages = result.conversation.content;
          const lastAgentMessage = messages
            .flat()
            .reverse()
            .find((msg) => msg.type === "agent_message");

          if (lastAgentMessage && lastAgentMessage.content) {
            // Store the result to apply later
            processedResults.push({
              ...textBlock,
              newText: lastAgentMessage.content
            });
          }

          processedCount++;
          document.getElementById(
            "status"
          ).innerHTML = `<div class="spinner"></div> Processing (${processedCount}/${totalBlocks})...`;
        } catch (error) {
          console.error(`Error processing text block: ${error.message}`);
          // Continue processing other blocks even if one fails
        }
      }

      // Add delay between batches if not the last batch
      if (i + BATCH_SIZE < textBlocksToProcess.length) {
        await new Promise((resolve) => setTimeout(resolve, BATCH_DELAY));
      }
    }
    
      // Clear the text blocks list to free memory
      textBlocksToProcess = null;
    });
  } catch (contextError) {
    // Catch errors from PowerPoint.run context
    console.error('[ProcessWithAssistant] PowerPoint.run error:', contextError);
    const errorHtml = `
      <div style="color: red; font-size: 12px;">
        <strong>❌ PowerPoint Context Error</strong><br>
        <div style="font-size: 10px; margin-top: 5px; padding: 5px; background: #fee; border-radius: 3px;">
          <strong>Message:</strong> ${contextError.message}<br>
          <strong>Name:</strong> ${contextError.name || 'Unknown'}<br>
          <strong>Code:</strong> ${contextError.code || 'None'}<br>
          <strong>Trace:</strong> ${contextError.traceMessages ? contextError.traceMessages.join(', ') : 'None'}<br>
        </div>
      </div>
    `;
    document.getElementById("status").innerHTML = errorHtml;
    throw contextError;
  }
  
  // Now apply all the results back to PowerPoint in a NEW context
  // But ONLY if not cancelled
  if (processedResults.length > 0 && !processingCancelled) {
    console.log(`[ProcessWithAssistant] Applying ${processedResults.length} processed results`);
    document.getElementById("status").innerHTML = `<div class="spinner"></div> Updating presentation...`;
    
    await PowerPoint.run(async (context) => {
      console.log('[ProcessWithAssistant] Starting new PowerPoint context for updates');
      // Load all slides and shapes fresh
      const presentation = context.presentation;
      presentation.slides.load("items");
      await context.sync();
      console.log(`[ProcessWithAssistant] Loaded ${presentation.slides.items.length} slides for update`);
      
      // Load all shapes for all slides with their IDs and types
      for (let slide of presentation.slides.items) {
        slide.shapes.load("items/id,items/type");
      }
      await context.sync();
      console.log('[ProcessWithAssistant] Loaded all shape IDs and types');
      
      // Now load textFrame properties ONLY for text-capable shapes
      const textCapableShapes = [];
      for (let slide of presentation.slides.items) {
        for (let shape of slide.shapes.items) {
          const shapeType = shape.type;
          // Only try to load textFrame for shapes that can have text
          if (shapeType === "GeometricShape" || 
              shapeType === "TextBox" || 
              shapeType === "Placeholder" ||
              shapeType === "Group" ||
              !shapeType) {
            try {
              shape.load("textFrame/hasText,textFrame/textRange");
              textCapableShapes.push(shape);
            } catch (e) {
              // Shape doesn't have textFrame
              console.log(`Shape ${shape.id} of type ${shapeType} doesn't have textFrame`);
            }
          }
        }
      }
      
      if (textCapableShapes.length > 0) {
        await context.sync();
        console.log(`[ProcessWithAssistant] Loaded textFrames for ${textCapableShapes.length} text-capable shapes`);
      }
      
      // Apply the updates
      let updatedCount = 0;
      let failedCount = 0;
      
      for (let result of processedResults) {
        try {
          console.log(`[ProcessWithAssistant] Updating shape ${result.shapeId}`);
          
          if (result.isSelectedText) {
            // Can't update selected text automatically
            console.log("[ProcessWithAssistant] Selected text can't be updated automatically");
            continue;
          }
          
          let shapeFound = false;
          
          // Shape IDs should be unique across the presentation
          // So we can always search by shape ID regardless of scope
          if (result.shapeId) {
            // For slide scope and selection, or when we don't have reliable slideIndex
            if (result.isSlideScope || result.isSelection || result.slideIndex === null) {
              // Find shape by ID across all slides
              for (let slide of presentation.slides.items) {
                for (let shape of slide.shapes.items) {
                  if (shape.id === result.shapeId) {
                    // Check if this shape is text-capable before trying to access textFrame
                    const shapeType = shape.type;
                    if (shapeType === "GeometricShape" || 
                        shapeType === "TextBox" || 
                        shapeType === "Placeholder" ||
                        shapeType === "Group" ||
                        !shapeType) {
                      try {
                        if (shape.textFrame && shape.textFrame.hasText && shape.textFrame.textRange) {
                          // Trim the new text before setting it
                          shape.textFrame.textRange.text = result.newText.trim();
                          shapeFound = true;
                          console.log(`[ProcessWithAssistant] Updated shape ${result.shapeId} in slide scope`);
                        } else {
                          console.log(`[ProcessWithAssistant] Shape ${result.shapeId} has no text content`);
                        }
                      } catch (e) {
                        console.log(`[ProcessWithAssistant] Error accessing textFrame for shape ${result.shapeId}: ${e.message}`);
                      }
                    } else {
                      console.log(`[ProcessWithAssistant] Shape ${result.shapeId} of type ${shapeType} cannot have text`);
                    }
                    break;
                  }
                }
                if (shapeFound) break;
              }
            } else if (result.slideIndex >= 0 && result.slideIndex < presentation.slides.items.length) {
              // For presentation scope with valid slide index, optimize by checking specific slide first
              const slide = presentation.slides.items[result.slideIndex];
              if (slide) {
                for (let shape of slide.shapes.items) {
                  if (shape.id === result.shapeId) {
                    // Check if this shape is text-capable before trying to access textFrame
                    const shapeType = shape.type;
                    if (shapeType === "GeometricShape" || 
                        shapeType === "TextBox" || 
                        shapeType === "Placeholder" ||
                        shapeType === "Group" ||
                        !shapeType) {
                      try {
                        if (shape.textFrame && shape.textFrame.hasText && shape.textFrame.textRange) {
                          // Trim the new text before setting it
                          shape.textFrame.textRange.text = result.newText.trim();
                          shapeFound = true;
                          console.log(`[ProcessWithAssistant] Updated shape ${result.shapeId} in slide ${result.slideIndex}`);
                        } else {
                          console.log(`[ProcessWithAssistant] Shape ${result.shapeId} has no text content`);
                        }
                      } catch (e) {
                        console.log(`[ProcessWithAssistant] Error accessing textFrame for shape ${result.shapeId}: ${e.message}`);
                      }
                    } else {
                      console.log(`[ProcessWithAssistant] Shape ${result.shapeId} of type ${shapeType} cannot have text`);
                    }
                    break;
                  }
                }
              }
              
              // If not found in expected slide, search all slides (fallback)
              if (!shapeFound) {
                for (let slide of presentation.slides.items) {
                  for (let shape of slide.shapes.items) {
                    if (shape.id === result.shapeId) {
                      const shapeType = shape.type;
                      if (shapeType === "GeometricShape" || 
                          shapeType === "TextBox" || 
                          shapeType === "Placeholder" ||
                          shapeType === "Group" ||
                          !shapeType) {
                        try {
                          if (shape.textFrame && shape.textFrame.hasText && shape.textFrame.textRange) {
                            shape.textFrame.textRange.text = result.newText.trim();
                            shapeFound = true;
                            console.log(`[ProcessWithAssistant] Updated shape ${result.shapeId} (fallback search)`);
                          }
                        } catch (e) {
                          // Silent fail for fallback
                        }
                      }
                      break;
                    }
                  }
                  if (shapeFound) break;
                }
              }
            }
          }
          
          if (!shapeFound) {
            console.log(`[ProcessWithAssistant] Could not find shape to update: ${result.shapeId}`);
            failedCount++;
          } else {
            updatedCount++;
          }
        } catch (e) {
          console.error(`[ProcessWithAssistant] Error updating text block:`, e.message);
          failedCount++;
        }
      }
      
      console.log(`[ProcessWithAssistant] Update complete. Updated: ${updatedCount}, Failed: ${failedCount}`);
      
      await context.sync();
      console.log('[ProcessWithAssistant] Final sync completed');
      document.getElementById("status").innerHTML = `<div class="spinner"></div> Presentation updated successfully`;
    });
  } else {
    console.log('[ProcessWithAssistant] No results to apply');
  }
  
  console.log('[ProcessWithAssistant] Process completed successfully');
}

function updateProgressDisplay() {
  const statusDiv = document.getElementById("status");
  if (processingProgress.status === "processing") {
    statusDiv.innerHTML = `<div class="spinner"></div> Processing (${processingProgress.current}/${processingProgress.total})`;
  }
}

// Count text blocks in presentation with progress updates
async function countPresentationTextBlocks() {
  // Update status before entering PowerPoint.run
  document.getElementById("status").innerHTML = '<div class="spinner"></div> Loading presentation...';
  
  return await PowerPoint.run(async (context) => {
    // Get all slides in the presentation
    const presentation = context.presentation;
    presentation.slides.load("items");
    await context.sync();
    
    const totalSlides = presentation.slides.items.length;
    document.getElementById("status").innerHTML = `<div class="spinner"></div> Loading ${totalSlides} slides...`;
    
    // Load all shapes from all slides at once (batched for performance)
    for (let slide of presentation.slides.items) {
      slide.shapes.load("items");
    }
    await context.sync();
    
    // Count total shapes
    let totalShapes = 0;
    for (let slide of presentation.slides.items) {
      totalShapes += slide.shapes.items.length;
    }
    
    document.getElementById("status").innerHTML = `<div class="spinner"></div> Checking ${totalShapes} shapes across ${totalSlides} slides...`;
    
    // Prepare shapes for safe extraction
    const shapesToCheck = [];
    for (let slideIndex = 0; slideIndex < presentation.slides.items.length; slideIndex++) {
      const slide = presentation.slides.items[slideIndex];
      for (let shape of slide.shapes.items) {
        shapesToCheck.push({ shape, slideIndex });
      }
    }
    
    // Use the safe extraction function to get text blocks
    const textBlocks = await safeExtractTextFromShapes(context, shapesToCheck);
    const textBlockCount = textBlocks.length;
    
    document.getElementById("status").innerHTML = `<div class="spinner"></div> Found ${textBlockCount} text blocks`;
    return textBlockCount;
  });
}

// Show confirmation dialog
function showConfirmationDialog(textBlockCount) {
  return new Promise((resolve) => {
    const modal = document.getElementById("confirmationModal");
    const message = document.getElementById("confirmationMessage");
    const confirmBtn = document.getElementById("confirmProcessBtn");
    const cancelBtn = document.getElementById("cancelProcessBtn");
    
    // Set message for >100 text blocks warning
    message.textContent = `You're about to process ${textBlockCount} text blocks. Processing this many blocks may take a while and could hit rate limits.\n\nAre you sure you want to continue?`;
    
    // Show modal
    modal.style.display = "flex";
    
    // Handle confirm
    const handleConfirm = () => {
      modal.style.display = "none";
      confirmBtn.removeEventListener("click", handleConfirm);
      cancelBtn.removeEventListener("click", handleCancel);
      resolve(true);
    };
    
    // Handle cancel
    const handleCancel = () => {
      modal.style.display = "none";
      confirmBtn.removeEventListener("click", handleConfirm);
      cancelBtn.removeEventListener("click", handleCancel);
      resolve(false);
    };
    
    confirmBtn.addEventListener("click", handleConfirm);
    cancelBtn.addEventListener("click", handleCancel);
  });
}

// Cancel processing function
function cancelProcessing() {
  processingCancelled = true;
  document.getElementById("status").innerHTML = "⚠️ Cancelling...";
  
  // Immediately restore UI to ready state
  document.getElementById("cancelBtn").style.display = "none";
  document.getElementById("submitBtn").style.display = "block";
  document.getElementById("submitBtn").disabled = false;
  
  // Show cancelled status briefly
  setTimeout(() => {
    document.getElementById("status").innerHTML = "Processing cancelled";
    setTimeout(() => {
      document.getElementById("status").innerHTML = "";
    }, 2000);
  }, 500);
}
