/* global Office, PowerPoint, DustOfficeAuth, callDustAPI, $ */

const DUST_VERSION = "0.1";
let processingProgress = { current: 0, total: 0, status: "idle" };
let processingCancelled = false;

// Initialize Office
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("myForm").addEventListener("submit", handleSubmit);
    document.getElementById("connectWorkOS").onclick = () => {
      DustOfficeAuth.initiateOAuth(buildAuthOptions());
    };
    const logoutBtn = document.getElementById("logoutBtn");
    if (logoutBtn) {
      logoutBtn.onclick = () => {
        showCredentialSetup();
        removeCredentials();
      };
    }
    const removeBtn = document.getElementById("removeCredentialsBtn");
    if (removeBtn) {
      removeBtn.onclick = showRemoveConfirmation;
    }
    const confirmRemoveBtn = document.getElementById("confirmRemove");
    if (confirmRemoveBtn) {
      confirmRemoveBtn.onclick = removeCredentials;
    }
    const cancelRemoveBtn = document.getElementById("cancelRemove");
    if (cancelRemoveBtn) {
      cancelRemoveBtn.onclick = hideRemoveConfirmation;
    }
    document.getElementById("cancelBtn").onclick = cancelProcessing;

    let postMessageProcessed = false;
    window.addEventListener("message", function (event) {
      console.log("[PowerPoint Taskpane] postMessage received:", event.data, event.origin);

      if (postMessageProcessed) {
        console.log("[PowerPoint Taskpane] postMessage already processed, ignoring");
        return;
      }

      try {
        const result = JSON.parse(event.data);
        if (result.success && result.action === "exchange_token" && result.code) {
          console.log("[PowerPoint Taskpane] Received code via postMessage, exchanging...");
          postMessageProcessed = true;
          void DustOfficeAuth.exchangeCodeForToken(result.code, buildAuthOptions());
        }
      } catch (error) {
        console.warn("[PowerPoint Taskpane] Failed to parse postMessage:", error);
      }
    });

    DustOfficeAuth.checkOAuthCallback(buildAuthOptions());

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
function getStorageKey(key) {
  return `dust_powerpoint_${key}`;
}

function saveToStorage(key, value) {
  if (typeof window.setStorageValue === "function") {
    window.setStorageValue(key, value);
    return;
  }

  const storageKey = getStorageKey(key);
  if (value === undefined || value === null) {
    localStorage.removeItem(storageKey);
  } else {
    localStorage.setItem(storageKey, value);
  }
}

function getFromStorage(key) {
  if (typeof window.getStorageValue === "function") {
    return window.getStorageValue(key);
  }

  return localStorage.getItem(getStorageKey(key));
}

// Check credentials and initialize the appropriate view
function checkCredentialsAndInitialize() {
  const accessToken = getFromStorage("accessToken");
  const workspaceId = getFromStorage("workspaceId");

  if (accessToken && workspaceId) {
    showMainForm();
    loadAssistants();
    initializeSelect2();
  } else {
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

  const errorDiv = document.getElementById("credentialError");
  if (errorDiv) {
    errorDiv.style.display = "none";
  }

  const loadingDiv = document.getElementById("oauthLoading");
  if (loadingDiv) {
    loadingDiv.style.display = "none";
  }

  const removeConfirmation = document.getElementById("removeConfirmation");
  if (removeConfirmation) {
    removeConfirmation.style.display = "none";
  }

  const removeBtn = document.getElementById("removeCredentialsBtn");
  if (removeBtn) {
    const hasCredentials =
      !!getFromStorage("accessToken") && !!getFromStorage("workspaceId");
    removeBtn.style.display = hasCredentials ? "block" : "none";
  }

  const connectBtn = document.getElementById("connectWorkOS");
  if (connectBtn) {
    connectBtn.style.display = "block";
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
  [
    "workspaceId",
    "accessToken",
    "refreshToken",
    "region",
    "credentialsConfigured",
    "user",
    "oauthCodeVerifier",
    "oauthRedirectUri",
  ].forEach((key) => {
    saveToStorage(key, null);
  });

  const removeBtn = document.getElementById("removeCredentialsBtn");
  if (removeBtn) {
    removeBtn.style.display = "none";
  }
  const removeConfirmation = document.getElementById("removeConfirmation");
  if (removeConfirmation) {
    removeConfirmation.style.display = "none";
  }

  const connectBtn = document.getElementById("connectWorkOS");
  if (connectBtn) {
    connectBtn.style.display = "block";
  }

  const errorDiv = document.getElementById("credentialError");
  if (errorDiv) {
    errorDiv.style.display = "none";
  }

  const select = document.getElementById("assistant");
  if (select) {
    select.innerHTML = '<option value=""></option>';
    select.disabled = true;
  }

  $("#assistant").select2({
    placeholder: "Loading agents...",
    allowClear: true,
    width: "100%",
  });

  const loadError = document.getElementById("loadError");
  if (loadError) {
    loadError.style.display = "none";
  }

  document.getElementById("credentialSetup").style.display = "block";
  document.getElementById("myForm").style.display = "none";
  const logoutBtn = document.getElementById("logoutBtn");
  if (logoutBtn) {
    logoutBtn.style.display = "none";
  }

  showCredentialSetup();
}

// Dust API functions
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

  document.getElementById("status").innerHTML = `<div class="spinner"></div> Extracting text from all slides...`;

  // Step 1: Extract text from ALL slides in a single PowerPoint.run context
  let allTextBlocks = [];
  let totalSlides = 0;

  try {
    await PowerPoint.run(async (context) => {
      const presentation = context.presentation;
      presentation.slides.load("items");
      await context.sync();

      totalSlides = presentation.slides.items.length;

      for (let slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
        if (processingCancelled) break;

        document.getElementById("status").innerHTML = `<div class="spinner"></div> Extracting text from slide ${slideIndex + 1}/${totalSlides}...`;

        const slide = presentation.slides.items[slideIndex];
        slide.load("id");
        slide.shapes.load("items");
        await context.sync();

        const slideId = slide.id;

        // Batch load all shape ids and types
        for (const shape of slide.shapes.items) {
          shape.load("id, type");
        }
        await context.sync();

        if (processingCancelled) break;

        for (const shape of slide.shapes.items) {
          if (processingCancelled) break;

          try {
            // Handle grouped shapes
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

                    for (const groupedShape of shape.group.shapes.items) {
                      groupedShape.load("id, type");
                    }
                    await context.sync();

                    for (const groupedShape of shape.group.shapes.items) {
                      if (processingCancelled) break;
                      const extracted = await extractTextFromShape(context, groupedShape, slideIndex, slideId, true, shape.id);
                      if (extracted) allTextBlocks.push({ ...extracted, isSlideScope: true });
                    }
                  }
                }
              } catch (e) { /* Group API not available */ }
              continue;
            }

            // Regular shape
            const extracted = await extractTextFromShape(context, shape, slideIndex, slideId, false, null);
            if (extracted) allTextBlocks.push({ ...extracted, isSlideScope: true });
          } catch (e) { /* Shape doesn't support text */ }
        }
      }
    });
  } catch (extractError) {
    document.getElementById("status").innerHTML = `❌ Error extracting text: ${extractError.message}`;
    return;
  }

  if (processingCancelled) {
    document.getElementById("status").innerHTML = "Processing cancelled";
    return;
  }

  if (allTextBlocks.length === 0) {
    document.getElementById("status").innerHTML = `⚠️ No text content found in ${totalSlides} slides`;
    return;
  }

  // Step 2: Process all text blocks via API
  document.getElementById("status").innerHTML = `<div class="spinner"></div> Processing ${allTextBlocks.length} text blocks...`;

  const processedResults = [];
  let processedCount = 0;

  for (let i = 0; i < allTextBlocks.length; i += MAX_CONCURRENT) {
    if (processingCancelled) {
      document.getElementById("status").innerHTML = "Processing cancelled";
      return;
    }

    const batch = allTextBlocks.slice(i, Math.min(i + MAX_CONCURRENT, allTextBlocks.length));
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
          processedCount++;
          document.getElementById("status").innerHTML = `<div class="spinner"></div> Processing (${processedCount}/${allTextBlocks.length})...`;
          return { ...textBlock, newText: lastAgentMessage.content };
        }
        return null;
      } catch (error) {
        return null;
      }
    });

    const batchResults = await Promise.all(batchPromises);
    processedResults.push(...batchResults.filter(r => r !== null));
  }

  if (processingCancelled) {
    document.getElementById("status").innerHTML = "Processing cancelled";
    return;
  }

  if (processedResults.length === 0) {
    document.getElementById("status").innerHTML = `❌ No text blocks were processed`;
    return;
  }

  // Step 3: Update all shapes in a single PowerPoint.run context
  document.getElementById("status").innerHTML = `<div class="spinner"></div> Updating ${processedResults.length} text blocks...`;

  let totalUpdated = 0;
  let totalFailed = 0;

  try {
    await PowerPoint.run(async (context) => {
      const { updatedCount, failedCount } = await updateShapes(context, processedResults);
      totalUpdated = updatedCount;
      totalFailed = failedCount;
    });
  } catch (updateError) {
    document.getElementById("status").innerHTML = `❌ Error updating: ${updateError.message}`;
    return;
  }

  // Show final status
  const slidesWithContent = new Set(processedResults.map(r => r.slideIndex)).size;
  if (totalUpdated > 0) {
    document.getElementById("status").innerHTML = `✅ Updated ${totalUpdated} text blocks across ${slidesWithContent} slides${totalFailed > 0 ? ` (${totalFailed} failed)` : ''}`;
  } else {
    document.getElementById("status").innerHTML = `❌ No text blocks were updated`;
  }
}

async function processWithAssistant(assistantId, instructions, scope) {
  const MAX_CONCURRENT = 10; // Process up to 10 text blocks concurrently

  const token = getFromStorage("accessToken");
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

function buildAuthOptions() {
  return {
    errorElement: document.getElementById("credentialError"),
    loadingElement: document.getElementById("oauthLoading"),
    connectButton: document.getElementById("connectWorkOS"),
    onAuthSuccess: handleOAuthSuccess,
    onAuthError: (error) => {
      console.error("[PowerPoint Taskpane] OAuth error:", error);
    },
  };
}

async function handleOAuthSuccess(data) {
  const { access_token, user, refresh_token } = data;

  if (!access_token) {
    throw new Error("No access token received");
  }

  saveToStorage("accessToken", access_token);
  saveToStorage("refreshToken", refresh_token);

  const { workspaceId, region } = DustOfficeAuth.decodeToken(access_token);

  saveToStorage("workspaceId", workspaceId);
  saveToStorage("region", region);
  saveToStorage("user", JSON.stringify(user));

  const loadingDiv = document.getElementById("oauthLoading");
  const errorDiv = document.getElementById("credentialError");
  if (loadingDiv) {
    loadingDiv.style.display = "none";
  }
  if (errorDiv) {
    errorDiv.style.display = "none";
  }

  const connectBtn = document.getElementById("connectWorkOS");
  if (connectBtn) {
    connectBtn.style.display = "none";
  }

  try {
    if (!workspaceId) {
      throw new Error("Workspace ID not found. Please ensure your WorkOS integration is configured correctly.");
    }

    const apiPath = `/api/v1/w/${workspaceId}/assistant/agent_configurations`;
    await callDustAPI(apiPath);

    saveToStorage("credentialsConfigured", "true");

    showMainForm();
    loadAssistants();
    initializeSelect2();
  } catch (error) {
    console.error("[PowerPoint Taskpane] Failed to validate token:", error);
    if (errorDiv) {
      errorDiv.textContent = "❌ " + error.message;
      errorDiv.style.display = "block";
    }
    if (connectBtn) {
      connectBtn.style.display = "block";
    }
    if (loadingDiv) {
      loadingDiv.style.display = "none";
    }
  }
}
