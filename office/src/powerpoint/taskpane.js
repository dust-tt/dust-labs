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

// Extract text from selected shapes
async function extractTextFromSelectedShapes(context) {
  const selectedShapes = context.presentation.getSelectedShapes();
  selectedShapes.load("items");
  await context.sync();

  if (!selectedShapes.items || selectedShapes.items.length === 0) {
    throw new Error("No shapes selected");
  }

  // Batch load id and type for all shapes
  for (let shape of selectedShapes.items) {
    shape.load("id, type");
  }
  await context.sync();

  // Get slide info from first selected shape
  const parentSlide = selectedShapes.items[0].getParentSlideOrNullObject();
  parentSlide.load("id");
  await context.sync();

  if (parentSlide.isNullObject) {
    throw new Error("Selected shape does not belong to any slide");
  }

  const slideId = parentSlide.id;
  const presentation = context.presentation;
  presentation.slides.load("items");
  await context.sync();

  for (let slide of presentation.slides.items) {
    slide.load("id");
  }
  await context.sync();

  let slideIndex = presentation.slides.items.findIndex(s => s.id === slideId);
  if (slideIndex === -1) {
    throw new Error("Could not determine slide index");
  }

  if (processingCancelled) return [];

  const textBlocks = [];
  for (let shape of selectedShapes.items) {
    if (processingCancelled) break;

    try {
      // Handle grouped shapes using ShapeGroup API (PowerPointApi 1.8+)
      if (shape.type === "Group" || shape.type === PowerPoint.ShapeType.group) {
        try {
          shape.load("group");
          await context.sync();

          if (shape.group) {
            shape.group.load("shapes");
            await context.sync();

            if (shape.group.shapes) {
              shape.group.shapes.load("items");
              await context.sync();

              // Load ids for all grouped shapes
              for (const groupedShape of shape.group.shapes.items) {
                groupedShape.load("id, type");
              }
              await context.sync();

              for (const groupedShape of shape.group.shapes.items) {
                if (processingCancelled) break;
                const extracted = await extractTextFromShape(context, groupedShape, slideIndex, slideId, true, shape.id);
                if (extracted) textBlocks.push({ ...extracted, isSelection: true });
              }
            }
          }
        } catch (e) { /* Group API not available */ }
        continue;
      }

      // Regular shape
      const extracted = await extractTextFromShape(context, shape, slideIndex, slideId, false, null);
      if (extracted) textBlocks.push({ ...extracted, isSelection: true });
    } catch (e) { /* Shape doesn't support text */ }
  }

  return textBlocks;
}

// Helper to extract text from a single shape (shape must already have id loaded)
async function extractTextFromShape(context, shape, slideIndex, slideId, isGroupedShape, parentGroupId) {
  try {
    if (shape.type === "Group" || shape.type === PowerPoint.ShapeType.group) return null;

    shape.load("textFrame");
    await context.sync();
    if (!shape.textFrame) return null;

    shape.textFrame.load("hasText, textRange");
    await context.sync();
    if (!shape.textFrame.hasText || !shape.textFrame.textRange) return null;

    shape.textFrame.textRange.load("text");
    await context.sync();

    const text = shape.textFrame.textRange.text?.trim();
    if (!text) return null;

    return {
      slideIndex,
      slideId,
      shapeId: shape.id,
      originalText: text,
      isGroupedShape,
      parentGroupId
    };
  } catch (e) {
    return null;
  }
}

// Extract text from all shapes on a specific slide
async function extractTextFromSlideShapes(context, slideIndex) {
  const presentation = context.presentation;
  presentation.slides.load("items");
  await context.sync();

  if (slideIndex < 0 || slideIndex >= presentation.slides.items.length) {
    throw new Error(`Slide index ${slideIndex} is out of bounds`);
  }

  const slide = presentation.slides.items[slideIndex];
  slide.load("id");
  slide.shapes.load("items");
  await context.sync();

  const slideId = slide.id;
  const textBlocks = [];

  // Batch load all shape ids and types first
  for (const shape of slide.shapes.items) {
    shape.load("id, type");
  }
  await context.sync();

  // Check for cancellation
  if (processingCancelled) return textBlocks;

  for (const shape of slide.shapes.items) {
    if (processingCancelled) break;

    try {
      // Handle grouped shapes using ShapeGroup API (PowerPointApi 1.8+)
      if (shape.type === "Group" || shape.type === PowerPoint.ShapeType.group) {
        try {
          shape.load("group");
          await context.sync();

          if (shape.group) {
            shape.group.load("shapes");
            await context.sync();

            if (shape.group.shapes) {
              shape.group.shapes.load("items");
              await context.sync();

              // Load ids for all grouped shapes
              for (const groupedShape of shape.group.shapes.items) {
                groupedShape.load("id, type");
              }
              await context.sync();

              for (const groupedShape of shape.group.shapes.items) {
                if (processingCancelled) break;
                const extracted = await extractTextFromShape(context, groupedShape, slideIndex, slideId, true, shape.id);
                if (extracted) textBlocks.push({ ...extracted, isSlideScope: true });
              }
            }
          }
        } catch (e) { /* Group API not available */ }
        continue;
      }

      // Regular shape
      const extracted = await extractTextFromShape(context, shape, slideIndex, slideId, false, null);
      if (extracted) textBlocks.push({ ...extracted, isSlideScope: true });
    } catch (e) { /* Shape doesn't support text */ }
  }

  return textBlocks;
}
// Update shapes with new text
async function updateShapes(context, updates) {
  const hasSelectionScope = updates.some(u => u.isSelection);

  if (hasSelectionScope) {
    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items");
    await context.sync();

    for (let shape of selectedShapes.items) {
      shape.load("id, type");
    }
    await context.sync();

    // Build a map of all selected shapes including those inside groups
    const shapesMap = new Map();
    for (let shape of selectedShapes.items) {
      shapesMap.set(shape.id, shape);

      if (shape.type === "Group" || shape.type === PowerPoint.ShapeType.group) {
        try {
          shape.load("group");
          await context.sync();
          if (shape.group) {
            shape.group.load("shapes");
            await context.sync();
            if (shape.group.shapes) {
              shape.group.shapes.load("items");
              await context.sync();
              for (let groupedShape of shape.group.shapes.items) {
                groupedShape.load("id");
                await context.sync();
                shapesMap.set(groupedShape.id, groupedShape);
              }
            }
          }
        } catch (e) { /* Group API not available */ }
      }
    }

    let updatedCount = 0;
    let failedCount = 0;

    for (let update of updates) {
      const freshShape = shapesMap.get(update.shapeId);
      if (!freshShape) {
        failedCount++;
        continue;
      }

      try {
        freshShape.load("textFrame");
        await context.sync();
        if (!freshShape.textFrame) {
          failedCount++;
          continue;
        }

        freshShape.textFrame.load("textRange");
        await context.sync();
        if (!freshShape.textFrame.textRange) {
          failedCount++;
          continue;
        }

        freshShape.textFrame.textRange.text = update.newText.trim();
        await context.sync();
        updatedCount++;
      } catch (e) {
        console.error(`Error updating shape ${update.shapeId}:`, e.message);
        failedCount++;
      }
    }

    return { updatedCount, failedCount };
  } else {
    // For slide/presentation scope, update shapes by reloading slides
    const presentation = context.presentation;
    presentation.slides.load("items");
    await context.sync();

    for (let slide of presentation.slides.items) {
      slide.load("id");
    }
    await context.sync();

    // Group updates by slide
    const updatesBySlide = new Map();
    for (let update of updates) {
      if (!updatesBySlide.has(update.slideIndex)) {
        updatesBySlide.set(update.slideIndex, []);
      }
      updatesBySlide.get(update.slideIndex).push(update);
    }

    let updatedCount = 0;
    let failedCount = 0;

    for (let [slideIndex, slideUpdates] of updatesBySlide) {
      const slide = presentation.slides.items[slideIndex];
      slide.shapes.load("items");
      await context.sync();

      for (let shape of slide.shapes.items) {
        shape.load("id, type");
      }
      await context.sync();

      // Build a map of all shapes on this slide including those inside groups
      const shapesMap = new Map();
      for (let shape of slide.shapes.items) {
        shapesMap.set(shape.id, shape);

        if (shape.type === "Group" || shape.type === PowerPoint.ShapeType.group) {
          try {
            shape.load("group");
            await context.sync();
            if (shape.group) {
              shape.group.load("shapes");
              await context.sync();
              if (shape.group.shapes) {
                shape.group.shapes.load("items");
                await context.sync();
                for (let groupedShape of shape.group.shapes.items) {
                  groupedShape.load("id");
                  await context.sync();
                  shapesMap.set(groupedShape.id, groupedShape);
                }
              }
            }
          } catch (e) { /* Group API not available */ }
        }
      }

      for (let update of slideUpdates) {
        const shape = shapesMap.get(update.shapeId);
        if (!shape) {
          failedCount++;
          continue;
        }

        try {
          shape.load("textFrame");
          await context.sync();
          if (!shape.textFrame) {
            failedCount++;
            continue;
          }

          shape.textFrame.load("textRange");
          await context.sync();
          if (!shape.textFrame.textRange) {
            failedCount++;
            continue;
          }

          shape.textFrame.textRange.text = update.newText.trim();
          await context.sync();
          updatedCount++;
        } catch (e) {
          console.error(`Error updating shape ${update.shapeId}:`, e.message);
          failedCount++;
        }
      }
    }

    return { updatedCount, failedCount };
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

  // Disable submit button and show cancel button
  document.getElementById("submitBtn").disabled = true;
  document.getElementById("submitBtn").style.display = "none";
  document.getElementById("cancelBtn").style.display = "block";
  processingCancelled = false;

  document.getElementById("status").innerHTML =
    '<div class="spinner"></div> Extracting content...';

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

// Process entire presentation slide by slide
async function processPresentationSlideBySlide(assistantId, instructions, token, workspaceId) {
  const MAX_CONCURRENT = 10;

  // First, get the total slide count
  let totalSlides = 0;
  document.getElementById("status").innerHTML = `<div class="spinner"></div> Getting slide count...`;

  await PowerPoint.run(async (context) => {
    context.presentation.slides.load("items");
    await context.sync();
    totalSlides = context.presentation.slides.items.length;
  });

  let totalUpdated = 0;
  let totalFailed = 0;
  let slidesWithContent = 0;

  // Process each slide: extract → process → update
  for (let slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
    if (processingCancelled) {
      document.getElementById("status").innerHTML = "Processing cancelled";
      return;
    }

    document.getElementById("status").innerHTML = `<div class="spinner"></div> Slide ${slideIndex + 1}/${totalSlides}: Extracting text...`;

    try {
      // Step 1: Extract text from this slide
      let slideTextBlocks = [];
      await PowerPoint.run(async (context) => {
        slideTextBlocks = await extractTextFromSlideShapes(context, slideIndex);
      });

      if (slideTextBlocks.length === 0) {
        continue;
      }

      slidesWithContent++;

      // Step 2: Process text blocks via API
      document.getElementById("status").innerHTML = `<div class="spinner"></div> Slide ${slideIndex + 1}/${totalSlides}: Processing ${slideTextBlocks.length} text blocks...`;

      const slideResults = [];
      for (let i = 0; i < slideTextBlocks.length; i += MAX_CONCURRENT) {
        if (processingCancelled) {
          document.getElementById("status").innerHTML = "Processing cancelled";
          return;
        }

        const batch = slideTextBlocks.slice(i, Math.min(i + MAX_CONCURRENT, slideTextBlocks.length));
        const batchPromises = batch.map(async (textBlock) => {
          if (processingCancelled) return null;

          try {
            const messageContent = (instructions || "Process this content:") + "\n\nInput:\n" + textBlock.originalText;
            const payload = {
              message: {
                content: messageContent,
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
              headers: { Authorization: "Bearer " + token },
            });

            if (processingCancelled) return null;

            const messages = result.conversation.content;
            const lastAgentMessage = messages.flat().reverse().find((msg) => msg.type === "agent_message");

            if (lastAgentMessage && lastAgentMessage.content) {
              return { ...textBlock, newText: lastAgentMessage.content };
            }
            return null;
          } catch (error) {
            return null;
          }
        });

        const batchResults = await Promise.all(batchPromises);
        slideResults.push(...batchResults.filter(r => r !== null));
      }

      // Step 3: Update shapes for this slide
      if (slideResults.length > 0 && !processingCancelled) {
        document.getElementById("status").innerHTML = `<div class="spinner"></div> Slide ${slideIndex + 1}/${totalSlides}: Updating ${slideResults.length} text blocks...`;

        await PowerPoint.run(async (context) => {
          const { updatedCount, failedCount } = await updateShapes(context, slideResults);
          totalUpdated += updatedCount;
          totalFailed += failedCount;
        });
      }

    } catch (slideError) {
      totalFailed++;
    }
  }

  // Show final status
  if (totalUpdated > 0) {
    document.getElementById("status").innerHTML = `✅ Updated ${totalUpdated} text blocks across ${slidesWithContent} slides${totalFailed > 0 ? ` (${totalFailed} failed)` : ''}`;
  } else if (slidesWithContent === 0) {
    document.getElementById("status").innerHTML = `⚠️ No text content found in ${totalSlides} slides`;
  } else {
    document.getElementById("status").innerHTML = `❌ No text blocks were updated`;
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

  // Handle presentation scope separately - process slide by slide
  if (scope === "presentation") {
    await processPresentationSlideBySlide(assistantId, instructions, token, workspaceId);
    return;
  }

  // Handle selection and slide scopes
  try {
    await PowerPoint.run(async (context) => {
      let textBlocksToProcess = [];

    if (scope === "selection") {
      // Extract text from selected shapes
      document.getElementById("status").innerHTML = `<div class="spinner"></div> Extracting from selected shapes...`;
      textBlocksToProcess = await extractTextFromSelectedShapes(context);

      if (textBlocksToProcess.length > 0) {
        document.getElementById("status").innerHTML = `<div class="spinner"></div> Processing ${textBlocksToProcess.length} text blocks...`;
      }

    } else if (scope === "slide") {
      // Determine which slide to process
      document.getElementById("status").innerHTML = `<div class="spinner"></div> Finding current slide...`;

      let targetSlideIndex = -1;
      const presentation = context.presentation;
      presentation.slides.load("items");
      await context.sync();
      // Determine which slide to process using various methods
      let targetSlide = null;

      // Method 1: Try to find slide from selected shape (most reliable for PowerPoint Online)
      try {
        const selectedShapes = context.presentation.getSelectedShapes();
        selectedShapes.load("items");
        await context.sync();

        if (selectedShapes.items && selectedShapes.items.length > 0) {
          // Use getParentSlide to find the slide
          const parentSlide = selectedShapes.items[0].getParentSlideOrNullObject();
          parentSlide.load("id");
          await context.sync();

          if (!parentSlide.isNullObject) {
            // Find the index of this slide
            for (let i = 0; i < presentation.slides.items.length; i++) {
              presentation.slides.items[i].load("id");
            }
            await context.sync();

            for (let i = 0; i < presentation.slides.items.length; i++) {
              if (presentation.slides.items[i].id === parentSlide.id) {
                targetSlideIndex = i;
                break;
              }
            }
          }
        }
      } catch (e) { /* Could not determine slide from selected shapes */ }

      // Method 2: Try getActiveSlide() (desktop PowerPoint)
      if (targetSlideIndex === -1) {
        try {
          const activeSlide = context.presentation.getActiveSlide();
          activeSlide.load("id");
          await context.sync();

          for (let i = 0; i < presentation.slides.items.length; i++) {
            presentation.slides.items[i].load("id");
          }
          await context.sync();

          for (let i = 0; i < presentation.slides.items.length; i++) {
            if (presentation.slides.items[i].id === activeSlide.id) {
              targetSlideIndex = i;
              break;
            }
          }
        } catch (e) { /* getActiveSlide not available */ }
      }

      // Method 3: Fallback to getSelectedSlides() (from thumbnail panel)
      if (targetSlideIndex === -1) {
        try {
          const selectedSlides = context.presentation.getSelectedSlides();
          selectedSlides.load("items");
          await context.sync();

          if (selectedSlides.items && selectedSlides.items.length > 0) {
            for (let i = 0; i < presentation.slides.items.length; i++) {
              presentation.slides.items[i].load("id");
            }
            await context.sync();

            for (let i = 0; i < presentation.slides.items.length; i++) {
              if (presentation.slides.items[i].id === selectedSlides.items[0].id) {
                targetSlideIndex = i;
                break;
              }
            }
          }
        } catch (e) { /* getSelectedSlides also failed */ }
      }

      if (targetSlideIndex === -1) {
        throw new Error("Could not determine the current slide");
      }

      // Extract text from the slide using helper function
      document.getElementById("status").innerHTML = `<div class="spinner"></div> Extracting from slide ${targetSlideIndex + 1}...`;
      textBlocksToProcess = await extractTextFromSlideShapes(context, targetSlideIndex);

      if (textBlocksToProcess.length > 0) {
        document.getElementById("status").innerHTML = `<div class="spinner"></div> Processing ${textBlocksToProcess.length} text blocks...`;
      }

    }

    // Check if we have text blocks to process (for selection and slide scopes)
    if (!textBlocksToProcess || textBlocksToProcess.length === 0) {
      document.getElementById("status").innerHTML = `
        <div style="color: #f59e0b; padding: 10px; background: #fef3c7; border-radius: 4px; font-size: 12px;">
          <strong>⚠️ No text found</strong><br>
          <span style="font-size: 11px; margin-top: 5px; display: block;">
            The selected ${scope === 'slide' ? 'slide' : scope === 'selection' ? 'shape' : 'content'}
            doesn't contain any text to process.
            ${scope === 'slide' ? '<br><br>Try selecting a slide with text content, or select a specific text box instead.' : ''}
          </span>
        </div>
      `;
      return;
    }

    // Warn if processing more than 100 text blocks
    if (textBlocksToProcess.length > 100) {
      const message = `You're about to process ${textBlocksToProcess.length} text blocks. Processing this many blocks may take a while and could hit rate limits.\n\nAre you sure you want to continue?`;
      if (!confirm(message)) {
        throw new Error("Processing cancelled by user");
      }
    }

    // Process all text blocks - API calls in parallel, updates after
    const totalBlocks = textBlocksToProcess.length;
    let processedCount = 0;

    document.getElementById(
      "status"
    ).innerHTML = `<div class="spinner"></div> Processing ${totalBlocks} text block(s)...`;

    // Create a function to process a single text block via API
    const processTextBlock = async (textBlock, index) => {
      // Check if processing was cancelled
      if (processingCancelled) {
        return null;
      }

      try {
        // Prepare the message content with instructions and slide content
        const messageContent = (instructions || "Process this content:") + "\n\nInput:\n" + textBlock.originalText;

        // Call Dust API for this text block
        const payload = {
          message: {
            content: messageContent,
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

          // Update status to show progress
          document.getElementById(
            "status"
          ).innerHTML = `<div class="spinner"></div> Processing (${processedCount}/${totalBlocks})...`;

          return {
            ...textBlock,
            newText: lastAgentMessage.content
          };
        }

        return null;
      } catch (error) {
        return null;
      }
    };

    // Process all text blocks in parallel batches
    const results = [];
    for (let i = 0; i < textBlocksToProcess.length; i += MAX_CONCURRENT) {
      // Check if cancelled
      if (processingCancelled) {
        document.getElementById("status").innerHTML = "Processing cancelled";
        return;
      }

      // Process batch of MAX_CONCURRENT items in parallel
      const batch = textBlocksToProcess.slice(i, Math.min(i + MAX_CONCURRENT, textBlocksToProcess.length));
      const batchPromises = batch.map((block, idx) => processTextBlock(block, i + idx));
      const batchResults = await Promise.all(batchPromises);

      // Add non-null results
      results.push(...batchResults.filter(r => r !== null));
    }

    processedResults = results;
      textBlocksToProcess = null;
    });
  } catch (contextError) {
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

  // Now apply all the results back to PowerPoint (for selection and slide scopes)
  if (processedResults.length > 0 && !processingCancelled) {
    document.getElementById("status").innerHTML = `<div class="spinner"></div> Updating presentation...`;

    try {
      await PowerPoint.run(async (context) => {
        const { updatedCount, failedCount } = await updateShapes(context, processedResults);

        // Show final status
        if (updatedCount > 0) {
          document.getElementById("status").innerHTML = `✅ Successfully updated ${updatedCount} text block(s)${failedCount > 0 ? ` (${failedCount} failed)` : ''}`;
        } else {
          document.getElementById("status").innerHTML = `❌ Failed to update text blocks`;
        }
      });
    } catch (updateError) {
      const errorHtml = `
        <div style="color: red; font-size: 12px;">
          <strong>❌ Update Error</strong><br>
          <div style="font-size: 10px; margin-top: 5px; padding: 5px; background: #fee; border-radius: 3px;">
            <strong>Message:</strong> ${updateError.message}<br>
            <strong>Name:</strong> ${updateError.name || 'Unknown'}<br>
            <strong>Code:</strong> ${updateError.code || 'None'}
          </div>
        </div>
      `;
      document.getElementById("status").innerHTML = errorHtml;
    }
  } else if (!processingCancelled) {
    document.getElementById("status").innerHTML = `❌ No text blocks were processed`;
  }
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
