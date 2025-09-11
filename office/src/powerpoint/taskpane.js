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
    document.getElementById("status").innerHTML =
      '<div class="spinner"></div> Counting text blocks...';
    
    try {
      const textBlockCount = await countPresentationTextBlocks();
      
      if (textBlockCount === 0) {
        document.getElementById("submitBtn").disabled = false;
        document.getElementById("status").textContent = "❌ No text content found to process";
        return;
      }
      
      // Show confirmation dialog
      const confirmed = await showConfirmationDialog(textBlockCount);
      
      if (!confirmed) {
        document.getElementById("submitBtn").disabled = false;
        document.getElementById("status").innerHTML = "";
        return;
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
      document.getElementById("status").textContent =
        "❌ Error: " + error.message;
    }
  } finally {
    // Reset buttons
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

  await PowerPoint.run(async (context) => {
    let textBlocksToProcess = [];

    if (scope === "presentation") {
      // Get all slides in the presentation
      const presentation = context.presentation;
      presentation.slides.load("items");
      await context.sync();

      // Load all shapes from all slides at once
      for (let slide of presentation.slides.items) {
        slide.shapes.load("items");
      }
      await context.sync();

      // Load all textFrames at once
      const shapesToCheck = [];
      for (let slideIndex = 0; slideIndex < presentation.slides.items.length; slideIndex++) {
        const slide = presentation.slides.items[slideIndex];
        for (let shape of slide.shapes.items) {
          shapesToCheck.push({ shape, slide, slideIndex });
          try {
            shape.load("textFrame/hasText,type");
          } catch (e) {
            // Some shapes might not have textFrame
          }
        }
      }
      
      if (shapesToCheck.length > 0) {
        await context.sync();

        // Load text content for shapes that have text
        const shapesWithText = [];
        for (let item of shapesToCheck) {
          try {
            if (item.shape.textFrame && item.shape.textFrame.hasText) {
              item.shape.textFrame.load("textRange/text");
              shapesWithText.push(item);
            }
          } catch (e) {
            // Shape might not have a textFrame, continue
          }
        }

        if (shapesWithText.length > 0) {
          await context.sync();

          // Collect text blocks
          for (let item of shapesWithText) {
            try {
              const text = item.shape.textFrame.textRange.text;
              if (text && text.trim()) {
                textBlocksToProcess.push({
                  shape: item.shape,
                  textRange: item.shape.textFrame.textRange,
                  originalText: text,
                  slideIndex: item.slideIndex + 1,
                });
              }
            } catch (e) {
              console.log("Error getting text from shape:", e);
            }
          }
        }
      }
    } else if (scope === "slide") {
      // Get the currently active slide
      let targetSlides = [];

      try {
        // Try to get selected slides first
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items");
        await context.sync();

        if (selectedSlides.items && selectedSlides.items.length > 0) {
          targetSlides = selectedSlides.items;
        }
      } catch (e) {
        // getSelectedSlides might not be available or might fail
        console.log("Could not get selected slides, using active slide");
      }

      // If no selected slides, get the active slide
      if (targetSlides.length === 0) {
        try {
          // Get the active slide through selected shapes
          const selectedShapes = context.presentation.getSelectedShapes();
          selectedShapes.load("items/id");
          await context.sync();

          if (selectedShapes.items && selectedShapes.items.length > 0) {
            // Get the slide that contains the first selected shape
            const shape = selectedShapes.items[0];
            shape.load("parentSlide");
            await context.sync();

            if (shape.parentSlide) {
              targetSlides = [shape.parentSlide];
            }
          }
        } catch (e) {
          console.log("Could not get active slide through shapes");
        }
      }

      // Final fallback: use the first slide
      if (targetSlides.length === 0) {
        const presentation = context.presentation;
        presentation.slides.load("items");
        await context.sync();
        if (presentation.slides.items.length > 0) {
          targetSlides = [presentation.slides.items[0]];
        }
      }

      // Load all shapes from target slides at once
      for (let slide of targetSlides) {
        slide.shapes.load("items");
      }
      await context.sync();

      // Load all textFrames at once
      const shapesToCheck = [];
      for (let slide of targetSlides) {
        const slideIndex = targetSlides.indexOf(slide);
        for (let shape of slide.shapes.items) {
          shapesToCheck.push({ shape, slide, slideIndex });
          try {
            shape.load("textFrame/hasText,type");
          } catch (e) {
            // Some shapes might not have textFrame
          }
        }
      }
      
      if (shapesToCheck.length > 0) {
        await context.sync();

        // Load text content for shapes that have text
        const shapesWithText = [];
        for (let item of shapesToCheck) {
          try {
            if (item.shape.textFrame && item.shape.textFrame.hasText) {
              item.shape.textFrame.load("textRange/text");
              shapesWithText.push(item);
            }
          } catch (e) {
            // Shape might not have a textFrame, continue
          }
        }

        if (shapesWithText.length > 0) {
          await context.sync();

          // Collect text blocks
          for (let item of shapesWithText) {
            try {
              const text = item.shape.textFrame.textRange.text;
              if (text && text.trim()) {
                textBlocksToProcess.push({
                  shape: item.shape,
                  textRange: item.shape.textFrame.textRange,
                  originalText: text,
                  slideIndex: item.slideIndex + 1,
                });
              }
            } catch (e) {
              console.log("Error getting text from shape:", e);
            }
          }
        }
      }
    } else if (scope === "selection") {
      // Handle selected shapes/text
      const selectedShapes = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length > 0) {
        // Process each selected shape
        for (let shape of selectedShapes.items) {
          try {
            shape.load("textFrame/hasText");
            await context.sync();
            
            if (shape.textFrame && shape.textFrame.hasText) {
              shape.textFrame.load("textRange/text");
              await context.sync();
              
              const text = shape.textFrame.textRange.text;
              if (text && text.trim()) {
                textBlocksToProcess.push({
                  shape: shape,
                  textRange: shape.textFrame.textRange,
                  originalText: text,
                  slideIndex: null,
                });
              }
            }
          } catch (e) {
            console.log("Selected shape without textFrame, skipping");
          }
        }
      } else {
        // Check for selected text range
        const selectedTextRange = context.presentation.getSelectedTextRange();
        selectedTextRange.load("text");
        await context.sync();

        if (selectedTextRange.text && selectedTextRange.text.trim()) {
          textBlocksToProcess.push({
            shape: null,
            textRange: selectedTextRange,
            originalText: selectedTextRange.text,
            slideIndex: null,
          });
        }
      }
    }

    // Check if we have text blocks to process
    if (textBlocksToProcess.length === 0) {
      throw new Error("No text content found to process");
    }

    // Warn if processing more than 100 text blocks
    if (textBlocksToProcess.length > 100) {
      const message = `You're about to process ${textBlocksToProcess.length} text blocks. Processing this many blocks may take a while and could hit rate limits.\n\nAre you sure you want to continue?`;
      if (!confirm(message)) {
        throw new Error("Processing cancelled by user");
      }
    }

    // Process text blocks in batches
    const totalBlocks = textBlocksToProcess.length;
    let processedCount = 0;

    document.getElementById(
      "status"
    ).innerHTML = `<div class="spinner"></div> Processing ${totalBlocks} text block(s)...`;

    // Process in batches to avoid rate limiting
    for (let i = 0; i < textBlocksToProcess.length; i += BATCH_SIZE) {
      // Check if processing was cancelled
      if (processingCancelled) {
        throw new Error("Processing cancelled by user");
      }
      
      const batch = textBlocksToProcess.slice(
        i,
        Math.min(i + BATCH_SIZE, textBlocksToProcess.length)
      );

      // Process each text block in the batch
      for (let textBlock of batch) {
        // Check cancellation for each item
        if (processingCancelled) {
          throw new Error("Processing cancelled by user");
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

          const messages = result.conversation.content;
          const lastAgentMessage = messages
            .flat()
            .reverse()
            .find((msg) => msg.type === "agent_message");

          if (lastAgentMessage && lastAgentMessage.content) {
            // Replace the text in the text block
            textBlock.textRange.text = lastAgentMessage.content;
            await context.sync();
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
  });
}

function updateProgressDisplay() {
  const statusDiv = document.getElementById("status");
  if (processingProgress.status === "processing") {
    statusDiv.innerHTML = `<div class="spinner"></div> Processing (${processingProgress.current}/${processingProgress.total})`;
  }
}

// Count text blocks in presentation
async function countPresentationTextBlocks() {
  return await PowerPoint.run(async (context) => {
    let textBlockCount = 0;
    
    // Get all slides in the presentation
    const presentation = context.presentation;
    presentation.slides.load("items");
    await context.sync();
    
    console.log(`Found ${presentation.slides.items.length} slides`);
    
    // Count text blocks from all slides
    for (let slideIndex = 0; slideIndex < presentation.slides.items.length; slideIndex++) {
      const slide = presentation.slides.items[slideIndex];
      slide.shapes.load("items,type");
      await context.sync();
      
      console.log(`Slide ${slideIndex + 1} has ${slide.shapes.items.length} shapes`);
      
      for (let shape of slide.shapes.items) {
        try {
          // Load textFrame and check if it has text
          shape.load("textFrame/hasText,type");
          await context.sync();
          
          if (shape.textFrame && shape.textFrame.hasText) {
            shape.textFrame.load("textRange/text");
            await context.sync();
            
            const text = shape.textFrame.textRange.text;
            if (text && text.trim()) {
              console.log(`Found text block: "${text.substring(0, 50)}..."`);
              textBlockCount++;
            }
          }
        } catch (e) {
          // Shape might not have a textFrame, continue
          console.log("Shape without textFrame, skipping:", e.message);
        }
      }
    }
    
    console.log(`Total text blocks found: ${textBlockCount}`);
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
    
    // Set message
    const blockText = textBlockCount === 1 ? "text block" : "text blocks";
    message.textContent = `This will process ${textBlockCount} ${blockText} across the entire presentation. Each text block will be sent to the selected agent for processing. Do you want to continue?`;
    
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
  
  // Hide cancel button immediately
  document.getElementById("cancelBtn").style.display = "none";
  
  setTimeout(() => {
    document.getElementById("status").innerHTML = "Processing cancelled";
    setTimeout(() => {
      document.getElementById("status").innerHTML = "";
    }, 2000);
  }, 500);
}
