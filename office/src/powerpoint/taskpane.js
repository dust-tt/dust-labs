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
  
  console.log(`[safeExtractText] Processing ${shapes.length} shapes`);
  
  // First, load all shape types and IDs in one batch
  for (let item of shapes) {
    try {
      item.shape.load("type,id,name");
    } catch (e) {
      console.log(`[safeExtractText] Error loading shape properties: ${e.message}`);
    }
  }
  await context.sync();
  
  // Identify text-capable shapes
  const textCapableShapes = [];
  for (let item of shapes) {
    try {
      const shapeType = item.shape.type;
      const shapeName = item.shape.name || "";
      
      // Be more inclusive about what might have text
      // Some shapes report as undefined or other types but still have text
      const isLikelyTextCapable = 
        shapeType === "GeometricShape" || 
        shapeType === "TextBox" || 
        shapeType === "Placeholder" ||
        shapeType === "Group" ||
        shapeType === "Rectangle" ||
        shapeType === "RoundRectangle" ||
        shapeName.toLowerCase().includes("text") ||
        shapeName.toLowerCase().includes("title") ||
        shapeName.toLowerCase().includes("content") ||
        !shapeType; // undefined type might still have text
      
      if (isLikelyTextCapable) {
        textCapableShapes.push(item);
      }
    } catch (e) {
      // If we can't determine type, try to include it anyway
      textCapableShapes.push(item);
    }
  }
  
  console.log(`[safeExtractText] Found ${textCapableShapes.length} potentially text-capable shapes out of ${shapes.length} total shapes`);
  
  if (textCapableShapes.length === 0) {
    return textBlocks;
  }
  
  // Try to load textFrame for all text-capable shapes in batches
  const shapesWithTextFrame = [];
  
  // First batch: try to load textFrame - queue all loads first
  for (let item of textCapableShapes) {
    item.shape.load("textFrame");
  }
  
  // Sync and check which shapes actually have textFrame
  try {
    await context.sync();
    
    // After sync, check which shapes have valid textFrame
    for (let item of textCapableShapes) {
      try {
        if (item.shape.textFrame) {
          shapesWithTextFrame.push(item);
        }
      } catch (e) {
        // This shape doesn't have a textFrame property
        console.log(`[safeExtractText] Shape ${item.shape.id} doesn't support textFrame`);
      }
    }
  } catch (e) {
    // Some shapes failed to load textFrame - check them individually
    console.log(`[safeExtractText] Some shapes failed textFrame load, checking individually`);
    
    for (let item of textCapableShapes) {
      try {
        // Try to access textFrame directly
        const tf = item.shape.textFrame;
        if (tf) {
          shapesWithTextFrame.push(item);
        }
      } catch (individualError) {
        // This specific shape doesn't support textFrame
        console.log(`[safeExtractText] Shape ${item.shape.id} doesn't support textFrame`);
      }
    }
  }
  
  if (shapesWithTextFrame.length > 0) {
    
    // Second batch: for shapes with textFrame, try to load hasText
    const shapesToCheckForText = [];
    
    // Queue all hasText loads
    for (let item of shapesWithTextFrame) {
      if (item.shape.textFrame) {
        item.shape.textFrame.load("hasText");
      }
    }
    
    // Sync and check which textFrames have the hasText property
    try {
      await context.sync();
      
      // After sync, check which shapes have hasText
      for (let item of shapesWithTextFrame) {
        try {
          if (item.shape.textFrame && item.shape.textFrame.hasText !== undefined) {
            shapesToCheckForText.push(item);
          }
        } catch (e) {
          console.log(`[safeExtractText] Could not access hasText for shape ${item.shape.id}`);
        }
      }
    } catch (e) {
      console.log(`[safeExtractText] Some shapes failed hasText load, checking individually`);
      
      for (let item of shapesWithTextFrame) {
        try {
          if (item.shape.textFrame && item.shape.textFrame.hasText !== undefined) {
            shapesToCheckForText.push(item);
          }
        } catch (individualError) {
          console.log(`[safeExtractText] Shape ${item.shape.id} doesn't support hasText`);
        }
      }
    }
    
    if (shapesToCheckForText.length > 0) {
      
      // Third batch: for shapes with text, load the actual text content
      const shapesWithText = [];
      
      // Queue text loads for shapes that have text
      for (let item of shapesToCheckForText) {
        try {
          if (item.shape.textFrame && item.shape.textFrame.hasText) {
            item.shape.textFrame.load("textRange/text");
            shapesWithText.push(item);
          }
        } catch (e) {
          console.log(`[safeExtractText] Could not queue text load for shape ${item.shape.id}`);
        }
      }
      
      if (shapesWithText.length > 0) {
        try {
          await context.sync();
        } catch (e) {
          console.log(`[safeExtractText] Some shapes failed text load: ${e.message}`);
          // Continue with shapes that succeeded
        }
        
        // Extract the text
        for (let item of shapesWithText) {
          try {
            if (item.shape.textFrame && 
                item.shape.textFrame.textRange && 
                item.shape.textFrame.textRange.text) {
              const text = item.shape.textFrame.textRange.text;
              if (text.trim()) {
                textBlocks.push({
                  slideIndex: item.slideIndex + slideIndexOffset,
                  shapeId: item.shape.id,
                  originalText: text
                });
              }
            }
          } catch (e) {
            console.log(`[safeExtractText] Error extracting text from shape ${item.shape.id}: ${e.message}`);
          }
        }
      }
    }
  }
  
  console.log(`[safeExtractText] Successfully extracted ${textBlocks.length} text blocks`);
  
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
  const MAX_CONCURRENT = 10; // Process up to 10 text blocks concurrently

  const token = getFromStorage("dustToken");
  const workspaceId = getFromStorage("workspaceId");

  if (!token || !workspaceId) {
    throw new Error("Please configure your Dust credentials first");
  }

  let processedResults = [];
  let finalProcessedCount = 0;
  let finalUpdatedCount = 0;
  
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
      let targetSlide = null;
      let targetSlideIndex = -1;
      
      try {
        // Try to get the selected slides (works in desktop PowerPoint)
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();
        
        if (selectedSlides.items && selectedSlides.items.length > 0) {
          targetSlide = selectedSlides.items[0];
          console.log('[ProcessWithAssistant] Found selected slide');
          
          // Find the index of this slide
          const presentation = context.presentation;
          presentation.slides.load("items");
          await context.sync();
          
          for (let i = 0; i < presentation.slides.items.length; i++) {
            if (presentation.slides.items[i].id === targetSlide.id) {
              targetSlideIndex = i;
              break;
            }
          }
        }
      } catch (e) {
        console.log('[ProcessWithAssistant] getSelectedSlides not available, will use alternative method');
      }
      
      // If we couldn't get selected slide, try to determine the current slide
      if (!targetSlide) {
        try {
          // Try to get the active view and determine current slide from there
          const presentation = context.presentation;
          presentation.load("slideMasters");
          presentation.slides.load("items");
          await context.sync();
          
          // For PowerPoint Online or when getSelectedSlides doesn't work,
          // we need a different approach. Let's check if there's selected content
          try {
            const selectedShapes = context.presentation.getSelectedShapes();
            selectedShapes.load("items");
            await context.sync();
            
            if (selectedShapes.items && selectedShapes.items.length > 0) {
              // Find which slide contains these shapes
              for (let i = 0; i < presentation.slides.items.length; i++) {
                const slide = presentation.slides.items[i];
                slide.shapes.load("items/id");
                await context.sync();
                
                // Check if any selected shape belongs to this slide
                const shapeIds = slide.shapes.items.map(s => s.id);
                for (let selectedShape of selectedShapes.items) {
                  selectedShape.load("id");
                }
                await context.sync();
                
                for (let selectedShape of selectedShapes.items) {
                  if (shapeIds.includes(selectedShape.id)) {
                    targetSlide = slide;
                    targetSlideIndex = i;
                    console.log(`[ProcessWithAssistant] Found current slide from selected shapes: slide ${i + 1}`);
                    break;
                  }
                }
                if (targetSlide) break;
              }
            }
          } catch (e) {
            console.log('[ProcessWithAssistant] Could not determine current slide from selected shapes');
          }
          
          // If still no target slide, use the first slide as fallback
          if (!targetSlide && presentation.slides.items.length > 0) {
            targetSlide = presentation.slides.items[0];
            targetSlideIndex = 0;
            console.log('[ProcessWithAssistant] Using first slide as fallback');
          }
        } catch (e) {
          console.log('[ProcessWithAssistant] Error determining current slide:', e.message);
        }
      }
      
      if (!targetSlide) {
        throw new Error("Could not determine the current slide");
      }

      // Load all shapes from the target slide
      console.log(`[ProcessWithAssistant] Loading shapes from slide ${targetSlideIndex + 1}`);
      targetSlide.shapes.load("items");
      await context.sync();
      
      const totalShapes = targetSlide.shapes.items.length;
      console.log(`[ProcessWithAssistant] Found ${totalShapes} shapes on current slide`);

      // Prepare shapes for safe extraction
      const shapesToCheck = [];
      for (let shape of targetSlide.shapes.items) {
        shapesToCheck.push({ shape, slideIndex: targetSlideIndex });
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

    // Process all text blocks in parallel with concurrency limit
    const totalBlocks = textBlocksToProcess.length;
    let processedCount = 0;
    let updatedCount = 0;

    document.getElementById(
      "status"
    ).innerHTML = `<div class="spinner"></div> Processing ${totalBlocks} text block(s)...`;

    // Create a function to process a single text block and update immediately
    const processAndUpdateTextBlock = async (textBlock) => {
      // Check if processing was cancelled
      if (processingCancelled) {
        return null;
      }

      try {
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
          return null;
        }
        
        const messages = result.conversation.content;
        const lastAgentMessage = messages
          .flat()
          .reverse()
          .find((msg) => msg.type === "agent_message");

        if (lastAgentMessage && lastAgentMessage.content) {
          processedCount++;
          
          // Immediately update the shape in PowerPoint
          if (!textBlock.isSelectedText) {
            await PowerPoint.run(async (updateContext) => {
              try {
                // Get the specific slide
                const presentation = updateContext.presentation;
                presentation.slides.load("items");
                await updateContext.sync();
                
                // For selections or slide scope, we need to find the shape differently
                if (textBlock.isSelection || textBlock.isSlideScope) {
                  // Load all shapes to find by ID
                  for (let slide of presentation.slides.items) {
                    slide.shapes.load("items/id");
                  }
                  await updateContext.sync();
                  
                  // Find the shape by ID
                  let targetShape = null;
                  for (let slide of presentation.slides.items) {
                    for (let shape of slide.shapes.items) {
                      if (shape.id === textBlock.shapeId) {
                        targetShape = shape;
                        break;
                      }
                    }
                    if (targetShape) break;
                  }
                  
                  if (targetShape) {
                    targetShape.load("textFrame");
                    await updateContext.sync();
                    
                    if (targetShape.textFrame) {
                      targetShape.textFrame.load("textRange");
                      await updateContext.sync();
                      
                      if (targetShape.textFrame.textRange) {
                        targetShape.textFrame.textRange.text = lastAgentMessage.content.trim();
                        await updateContext.sync();
                        updatedCount++;
                        console.log(`[ProcessWithAssistant] Updated shape ${textBlock.shapeId} immediately`);
                      }
                    }
                  }
                } else if (textBlock.slideIndex !== null && textBlock.slideIndex !== undefined) {
                  // Regular presentation or slide processing
                  const targetSlide = presentation.slides.items[textBlock.slideIndex];
                  if (targetSlide) {
                    targetSlide.shapes.load("items/id");
                    await updateContext.sync();
                    
                    // Find the shape by ID
                    let targetShape = null;
                    for (let shape of targetSlide.shapes.items) {
                      if (shape.id === textBlock.shapeId) {
                        targetShape = shape;
                        break;
                      }
                    }
                    
                    if (targetShape) {
                      targetShape.load("textFrame");
                      await updateContext.sync();
                      
                      if (targetShape.textFrame) {
                        targetShape.textFrame.load("textRange");
                        await updateContext.sync();
                        
                        if (targetShape.textFrame.textRange) {
                          targetShape.textFrame.textRange.text = lastAgentMessage.content.trim();
                          await updateContext.sync();
                          updatedCount++;
                          console.log(`[ProcessWithAssistant] Updated shape ${textBlock.shapeId} immediately`);
                        }
                      }
                    }
                  }
                }
                
                // Update status
                document.getElementById(
                  "status"
                ).innerHTML = `<div class="spinner"></div> Processing (${processedCount}/${totalBlocks}) - Updated ${updatedCount}...`;
              } catch (updateError) {
                console.error(`Error updating shape ${textBlock.shapeId}:`, updateError.message);
              }
            });
          }
          
          return {
            ...textBlock,
            newText: lastAgentMessage.content
          };
        }
        return null;
      } catch (error) {
        console.error(`Error processing text block: ${error.message}`);
        return null;
      }
    };

    // Process text blocks with concurrency limit
    const processingPromises = [];
    for (let i = 0; i < textBlocksToProcess.length; i++) {
      // Check if cancelled before starting new processing
      if (processingCancelled) {
        document.getElementById("status").innerHTML = "Processing cancelled";
        return;
      }

      // Add to processing queue
      processingPromises.push(processAndUpdateTextBlock(textBlocksToProcess[i]));

      // If we've reached the concurrency limit, wait for some to complete
      if (processingPromises.length >= MAX_CONCURRENT) {
        // Wait for at least one to complete before continuing
        await Promise.race(processingPromises.filter(p => p !== null));
        
        // Clean up completed promises
        for (let j = processingPromises.length - 1; j >= 0; j--) {
          if (processingPromises[j] && await Promise.race([processingPromises[j], Promise.resolve('pending')]) !== 'pending') {
            processingPromises.splice(j, 1);
          }
        }
      }
    }

    // Wait for all remaining promises to complete
    const results = await Promise.all(processingPromises);
    processedResults = results.filter(r => r !== null);
    
    // Store final counts
    finalProcessedCount = processedCount;
    finalUpdatedCount = updatedCount;
    
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
  
  // Since we've already updated shapes immediately after processing,
  // we just need to show the final status
  if (!processingCancelled) {
    if (finalUpdatedCount > 0) {
      document.getElementById("status").innerHTML = `✅ Successfully updated ${finalUpdatedCount} text block(s)`;
    } else if (finalProcessedCount > 0) {
      document.getElementById("status").innerHTML = `✅ Processed ${finalProcessedCount} text block(s)`;
    } else {
      document.getElementById("status").innerHTML = `❌ No text blocks were processed`;
    }
  } else {
    console.log('[ProcessWithAssistant] Processing was cancelled');
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
