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

// ====================================================================
// EXTRACTION HELPER FUNCTIONS
// ====================================================================

// Extract text from currently selected shapes (for "selection" scope)
async function extractTextFromSelectedShapes(context) {
  console.log('[Extract] Getting selected shapes');
  const selectedShapes = context.presentation.getSelectedShapes();
  selectedShapes.load("items");
  await context.sync();

  if (!selectedShapes.items || selectedShapes.items.length === 0) {
    throw new Error("No shapes selected");
  }

  // Load shape IDs
  for (let shape of selectedShapes.items) {
    shape.load("id");
  }
  await context.sync();

  // Use getParentSlide to get the actual slide ID
  const selectedShape = selectedShapes.items[0];
  const parentSlide = selectedShape.getParentSlideOrNullObject();
  parentSlide.load("id");
  await context.sync();

  if (parentSlide.isNullObject) {
    throw new Error("Selected shape does not belong to any slide");
  }

  const slideId = parentSlide.id;

  // Find the slide index
  const presentation = context.presentation;
  presentation.slides.load("items");
  await context.sync();

  for (let slide of presentation.slides.items) {
    slide.load("id");
  }
  await context.sync();

  let slideIndex = -1;
  for (let i = 0; i < presentation.slides.items.length; i++) {
    if (presentation.slides.items[i].id === slideId) {
      slideIndex = i;
      break;
    }
  }

  if (slideIndex === -1) {
    throw new Error("Could not determine slide index");
  }

  console.log(`[Extract] Found ${selectedShapes.items.length} selected shapes on slide ${slideIndex + 1}`);

  // Extract text from each selected shape
  const textBlocks = [];
  for (let shape of selectedShapes.items) {
    try {
      // First check if this is a group
      shape.load("type");
      await context.sync();

      const shapeType = shape.type;
      console.log(`[Extract] Checking selected shape (ID: ${shape.id}, Type: ${shapeType})`);

      // Handle grouped shapes - use ShapeGroup API (PowerPointApi 1.8+)
      if (shapeType === "Group" || shapeType === PowerPoint.ShapeType.group) {
        console.log(`[Extract]   Shape ${shape.id} is a group - accessing shapes via shape.group.shapes`);

        try {
          // Use the ShapeGroup API
          shape.load("group");
          await context.sync();

          if (shape.group) {
            shape.group.load("shapes");
            await context.sync();

            if (shape.group.shapes) {
              shape.group.shapes.load("items");
              await context.sync();

              const groupedShapes = shape.group.shapes.items;
              console.log(`[Extract]   Group has ${groupedShapes.length} shapes inside`);

              for (let j = 0; j < groupedShapes.length; j++) {
                const groupedShape = groupedShapes[j];
                try {
                  groupedShape.load("id, type");
                  await context.sync();

                  console.log(`[Extract]   Processing grouped shape ${j + 1}/${groupedShapes.length} (ID: ${groupedShape.id}, Type: ${groupedShape.type})`);

                  if (groupedShape.type === "Group" || groupedShape.type === PowerPoint.ShapeType.group) {
                    console.log(`[Extract]   Skipping nested group ${groupedShape.id}`);
                    continue;
                  }

                  groupedShape.load("textFrame");
                  await context.sync();

                  if (groupedShape.textFrame) {
                    groupedShape.textFrame.load("hasText");
                    await context.sync();

                    if (groupedShape.textFrame.hasText) {
                      groupedShape.textFrame.load("textRange");
                      await context.sync();

                      if (groupedShape.textFrame.textRange) {
                        groupedShape.textFrame.textRange.load("text");
                        await context.sync();

                        const text = groupedShape.textFrame.textRange.text?.trim();
                        if (text) {
                          console.log(`[Extract]   ✓ Grouped shape ${groupedShape.id} extracted: "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}"`);
                          textBlocks.push({
                            slideIndex,
                            slideId,
                            shapeId: groupedShape.id,
                            originalText: text,
                            isSelection: true,
                            isGroupedShape: true,
                            parentGroupId: shape.id
                          });
                        }
                      }
                    }
                  }
                } catch (groupedShapeError) {
                  console.log(`[Extract]   Grouped shape ${j + 1} error: ${groupedShapeError.message}`);
                }
              }
            }
          } else {
            console.log(`[Extract]   shape.group is null - API may not be available`);
          }
        } catch (groupError) {
          console.log(`[Extract]   Error accessing group shapes: ${groupError.message}`);
        }
        continue; // Skip to next shape
      }

      // Not a group - extract text normally
      shape.load("textFrame");
      await context.sync();

      if (!shape.textFrame) continue;

      shape.textFrame.load("hasText");
      await context.sync();

      if (!shape.textFrame.hasText) continue;

      shape.textFrame.load("textRange");
      await context.sync();

      if (!shape.textFrame.textRange) continue;

      shape.textFrame.textRange.load("text");
      await context.sync();

      const text = shape.textFrame.textRange.text?.trim();
      if (text) {
        textBlocks.push({
          slideIndex,
          slideId,
          shapeId: shape.id,
          originalText: text,
          isSelection: true
        });
      }
    } catch (e) {
      console.log(`[Extract] Shape ${shape.id} doesn't support text: ${e.message}`);
    }
  }

  console.log(`[Extract] Extracted ${textBlocks.length} text blocks from selected shapes`);
  return textBlocks;
}

// Extract text from all shapes on a specific slide (for "slide" scope)
async function extractTextFromSlideShapes(context, slideIndex) {
  console.log(`[Extract] Extracting text from slide ${slideIndex + 1}`);

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
  console.log(`[Extract] Found ${slide.shapes.items.length} shapes on slide (ID: ${slideId})`);

  // Extract text from each shape individually
  const textBlocks = [];
  for (let i = 0; i < slide.shapes.items.length; i++) {
    const shape = slide.shapes.items[i];

    try {
      shape.load("id,type");
      await context.sync();

      const shapeId = shape.id;
      const shapeType = shape.type;
      console.log(`[Extract] Checking shape ${i + 1}/${slide.shapes.items.length} (ID: ${shapeId}, Type: ${shapeType})`);

      // Check if this is a group - if so, we need to extract from grouped shapes
      if (shapeType === "Group" || shapeType === PowerPoint.ShapeType.group) {
        console.log(`[Extract]   Shape ${shapeId} is a group - accessing shapes via shape.group.shapes (API 1.8+)`);

        try {
          // Use the ShapeGroup API (PowerPointApi 1.8+)
          // Access the group property which returns a ShapeGroup object
          shape.load("group");
          await context.sync();

          if (shape.group) {
            // Load the shapes collection from the group
            shape.group.load("shapes");
            await context.sync();

            if (shape.group.shapes) {
              shape.group.shapes.load("items");
              await context.sync();

              const groupedShapes = shape.group.shapes.items;
              console.log(`[Extract]   Group has ${groupedShapes.length} shapes inside`);

              // Extract text from each shape in the group
              for (let j = 0; j < groupedShapes.length; j++) {
                const groupedShape = groupedShapes[j];
                try {
                  groupedShape.load("id, type");
                  await context.sync();

                  console.log(`[Extract]   Processing grouped shape ${j + 1}/${groupedShapes.length} (ID: ${groupedShape.id}, Type: ${groupedShape.type})`);

                  // Skip if it's a nested group (could recurse, but keeping simple for now)
                  if (groupedShape.type === "Group" || groupedShape.type === PowerPoint.ShapeType.group) {
                    console.log(`[Extract]   Skipping nested group ${groupedShape.id}`);
                    continue;
                  }

                  groupedShape.load("textFrame");
                  await context.sync();

                  if (groupedShape.textFrame) {
                    groupedShape.textFrame.load("hasText");
                    await context.sync();

                    if (groupedShape.textFrame.hasText) {
                      groupedShape.textFrame.load("textRange");
                      await context.sync();

                      if (groupedShape.textFrame.textRange) {
                        groupedShape.textFrame.textRange.load("text");
                        await context.sync();

                        const text = groupedShape.textFrame.textRange.text?.trim();
                        if (text) {
                          console.log(`[Extract]   ✓ Grouped shape ${groupedShape.id} extracted: "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}"`);
                          textBlocks.push({
                            slideIndex,
                            slideId,
                            shapeId: groupedShape.id,
                            originalText: text,
                            isSlideScope: true,
                            isGroupedShape: true,
                            parentGroupId: shapeId
                          });
                        }
                      }
                    } else {
                      console.log(`[Extract]   Grouped shape ${groupedShape.id} hasText=false`);
                    }
                  } else {
                    console.log(`[Extract]   Grouped shape ${groupedShape.id} has no textFrame`);
                  }
                } catch (groupedShapeError) {
                  console.log(`[Extract]   Grouped shape ${j + 1} error: ${groupedShapeError.message}`);
                }
              }
            } else {
              console.log(`[Extract]   shape.group.shapes is null/undefined`);
            }
          } else {
            console.log(`[Extract]   shape.group is null/undefined - API may not be available`);
          }
        } catch (groupError) {
          console.log(`[Extract]   Error accessing group shapes: ${groupError.message}`);
          // Fallback: try to extract text directly from the group shape itself
          try {
            shape.load("textFrame");
            await context.sync();

            if (shape.textFrame) {
              shape.textFrame.load("hasText");
              await context.sync();

              if (shape.textFrame.hasText) {
                shape.textFrame.load("textRange");
                await context.sync();

                if (shape.textFrame.textRange) {
                  shape.textFrame.textRange.load("text");
                  await context.sync();

                  const text = shape.textFrame.textRange.text?.trim();
                  if (text) {
                    console.log(`[Extract]   ✓ Group shape ${shapeId} fallback - direct text: "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}"`);
                    textBlocks.push({
                      slideIndex,
                      slideId,
                      shapeId,
                      originalText: text,
                      isSlideScope: true,
                      isGroupShape: true
                    });
                  }
                }
              }
            }
          } catch (fallbackError) {
            console.log(`[Extract]   Fallback also failed: ${fallbackError.message}`);
          }
        }

        continue; // Skip to next shape
      }

      // Not a group - extract text normally
      shape.load("textFrame");
      await context.sync();

      if (!shape.textFrame) {
        console.log(`[Extract]   Shape ${shapeId} has no textFrame - skipping`);
        continue;
      }

      shape.textFrame.load("hasText");
      await context.sync();

      if (!shape.textFrame.hasText) {
        console.log(`[Extract]   Shape ${shapeId} hasText=false - skipping`);
        continue;
      }

      shape.textFrame.load("textRange");
      await context.sync();

      if (!shape.textFrame.textRange) {
        console.log(`[Extract]   Shape ${shapeId} has no textRange - skipping`);
        continue;
      }

      shape.textFrame.textRange.load("text");
      await context.sync();

      const text = shape.textFrame.textRange.text?.trim();
      if (text) {
        console.log(`[Extract]   ✓ Shape ${shapeId} extracted: "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}"`);
        textBlocks.push({
          slideIndex,
          slideId,
          shapeId,
          originalText: text,
          isSlideScope: true
        });
      } else {
        console.log(`[Extract]   Shape ${shapeId} has empty text - skipping`);
      }
    } catch (e) {
      console.log(`[Extract]   Shape ${i + 1} threw error: ${e.message} - skipping`);
    }
  }

  console.log(`[Extract] Extracted ${textBlocks.length} text blocks from slide ${slideIndex + 1}`);
  return textBlocks;
}

// Extract text from all slides (for "presentation" scope)
async function extractTextFromAllSlides(context) {
  console.log('[Extract] Extracting text from entire presentation');

  const presentation = context.presentation;
  presentation.slides.load("items");
  await context.sync();

  const totalSlides = presentation.slides.items.length;
  console.log(`[Extract] Found ${totalSlides} slides`);

  // Extract text from each slide using extractTextFromSlideShapes
  const allTextBlocks = [];
  for (let slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
    const slideBlocks = await extractTextFromSlideShapes(context, slideIndex);
    allTextBlocks.push(...slideBlocks);
  }

  console.log(`[Extract] Extracted ${allTextBlocks.length} text blocks from ${totalSlides} slides`);
  return allTextBlocks;
}

// ====================================================================
// UPDATE HELPER FUNCTION
// ====================================================================

// Update shapes with new text
async function updateShapes(context, updates) {
  console.log(`[Update] Updating ${updates.length} shapes`);

  // Check if all updates are for selected shapes
  const hasSelectionScope = updates.some(u => u.isSelection);

  if (hasSelectionScope) {
    // Use getSelectedShapes() for fresh references
    console.log('[Update] Using selection scope - getting selected shapes again');

    const selectedShapes = context.presentation.getSelectedShapes();
    selectedShapes.load("items");
    await context.sync();

    for (let shape of selectedShapes.items) {
      shape.load("id, type");
    }
    await context.sync();

    console.log(`[Update] Got ${selectedShapes.items.length} selected shapes`);

    // Build a map of all selected shapes including those inside groups
    const shapesMap = new Map();
    for (let shape of selectedShapes.items) {
      shapesMap.set(shape.id, shape);

      // Check if this is a group and load shapes inside it
      if (shape.type === "Group" || shape.type === PowerPoint.ShapeType.group) {
        console.log(`[Update] Shape ${shape.id} is a group - loading shapes inside`);
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
                console.log(`[Update] Added grouped shape ${groupedShape.id} to map`);
                shapesMap.set(groupedShape.id, groupedShape);
              }
            }
          }
        } catch (e) {
          console.log(`[Update] Error loading group shapes for ${shape.id}: ${e.message}`);
        }
      }
    }

    console.log(`[Update] Total shapes in map (including grouped): ${shapesMap.size}`);

    let updatedCount = 0;
    let failedCount = 0;

    for (let update of updates) {
      const freshShape = shapesMap.get(update.shapeId);

      if (!freshShape) {
        console.log(`[Update] Shape ${update.shapeId} not in current selection or groups`);
        failedCount++;
        continue;
      }

      try {
        freshShape.load("textFrame");
        await context.sync();

        if (!freshShape.textFrame) {
          console.log(`[Update] Shape ${update.shapeId} has no textFrame`);
          failedCount++;
          continue;
        }

        freshShape.textFrame.load("textRange");
        await context.sync();

        if (!freshShape.textFrame.textRange) {
          console.log(`[Update] Shape ${update.shapeId} has no textRange`);
          failedCount++;
          continue;
        }

        const newText = update.newText.trim();
        console.log(`[Update] Updating shape ${update.shapeId}: "${newText.substring(0, 30)}..."`);
        freshShape.textFrame.textRange.text = newText;
        await context.sync();

        updatedCount++;
      } catch (e) {
        console.error(`[Update] Error updating shape ${update.shapeId}: ${e.message}`);
        failedCount++;
      }
    }

    console.log(`[Update] Complete. Updated: ${updatedCount}, Failed: ${failedCount}`);
    return { updatedCount, failedCount };
  } else {
    // For slide/presentation scope, update shapes by reloading slides
    console.log('[Update] Using slide scope - reloading slides to update shapes');

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

    console.log(`[Update] Grouped updates across ${updatesBySlide.size} slides`);

    let updatedCount = 0;
    let failedCount = 0;

    // Update each slide
    for (let [slideIndex, slideUpdates] of updatesBySlide) {
      const slide = presentation.slides.items[slideIndex];
      slide.shapes.load("items");
      await context.sync();

      // Load all shape IDs and types on this slide
      for (let shape of slide.shapes.items) {
        shape.load("id, type");
      }
      await context.sync();

      // Build a map of all shapes on this slide including those inside groups
      const shapesMap = new Map();
      for (let shape of slide.shapes.items) {
        shapesMap.set(shape.id, shape);

        // Check if this is a group and load shapes inside it
        if (shape.type === "Group" || shape.type === PowerPoint.ShapeType.group) {
          console.log(`[Update] Shape ${shape.id} on slide ${slideIndex + 1} is a group - loading shapes inside`);
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
                  console.log(`[Update] Added grouped shape ${groupedShape.id} to map`);
                  shapesMap.set(groupedShape.id, groupedShape);
                }
              }
            }
          } catch (e) {
            console.log(`[Update] Error loading group shapes for ${shape.id}: ${e.message}`);
          }
        }
      }

      console.log(`[Update] Slide ${slideIndex + 1} has ${shapesMap.size} shapes (including grouped)`);

      // Update each shape
      for (let update of slideUpdates) {
        const shape = shapesMap.get(update.shapeId);

        if (!shape) {
          console.log(`[Update] Shape ${update.shapeId} not found on slide ${slideIndex + 1}`);
          failedCount++;
          continue;
        }

        try {
          shape.load("textFrame");
          await context.sync();

          if (!shape.textFrame) {
            console.log(`[Update] Shape ${update.shapeId} has no textFrame`);
            failedCount++;
            continue;
          }

          shape.textFrame.load("textRange");
          await context.sync();

          if (!shape.textFrame.textRange) {
            console.log(`[Update] Shape ${update.shapeId} has no textRange`);
            failedCount++;
            continue;
          }

          const newText = update.newText.trim();
          console.log(`[Update] Updating shape ${update.shapeId} on slide ${slideIndex + 1}`);
          shape.textFrame.textRange.text = newText;
          await context.sync();

          updatedCount++;
        } catch (e) {
          console.error(`[Update] Error updating shape ${update.shapeId}: ${e.message}`);
          failedCount++;
        }
      }
    }

    console.log(`[Update] Complete. Updated: ${updatedCount}, Failed: ${failedCount}`);
    return { updatedCount, failedCount };
  }
}

// ====================================================================
// LEGACY HELPER FUNCTION (kept for backward compatibility)
// ====================================================================

// Helper function to safely extract text from shapes
// Returns an array of objects with slideIndex, shapeId, and originalText
async function safeExtractTextFromShapes(context, shapes, slideIndexOffset = 0) {
  const textBlocks = [];

  if (!shapes || shapes.length === 0) {
    return textBlocks;
  }

  console.log(`[safeExtractText] Processing ${shapes.length} shapes`);

  // Log what we're receiving
  console.log('[safeExtractText] Input shapes:', shapes.map(s => ({
    slideIndex: s.slideIndex,
    slideId: s.slideId,
    shapeInfo: s.shape ? 'shape object present' : 'NO SHAPE'
  })));
  
  // First, load all shape types and IDs in one batch
  for (let item of shapes) {
    try {
      item.shape.load("type,id,name");
    } catch (e) {
      console.log(`[safeExtractText] Error loading shape properties: ${e.message}`);
    }
  }
  await context.sync();

  // Log what IDs we actually loaded
  console.log('[safeExtractText] After loading, shape IDs are:', shapes.map(s => {
    try {
      return { id: s.shape.id, slideIndex: s.slideIndex, slideId: s.slideId };
    } catch (e) {
      return { error: e.message };
    }
  }));
  
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
        // Shape doesn't have textFrame - this is normal for images, icons, etc.
      }
    }
  } catch (e) {
    // Some shapes failed to load textFrame - check them individually
    console.log(`[safeExtractText] Batch textFrame load failed, checking individually`);

    for (let item of textCapableShapes) {
      try {
        // Try to access textFrame directly
        const tf = item.shape.textFrame;
        if (tf) {
          shapesWithTextFrame.push(item);
        }
      } catch (individualError) {
        // Shape doesn't support textFrame - this is normal
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
          // Shape doesn't support hasText - this is normal for some shape types
        }
      }
    } catch (e) {
      console.log(`[safeExtractText] Batch hasText load failed, checking individually`);

      for (let item of shapesWithTextFrame) {
        try {
          if (item.shape.textFrame && item.shape.textFrame.hasText !== undefined) {
            shapesToCheckForText.push(item);
          }
        } catch (individualError) {
          // Shape doesn't support hasText - this is normal for some shape types
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
                  slideId: item.slideId,  // Store slide ID for unique identification
                  shapeId: item.shape.id,
                  originalText: text
                });
              }
            }
          } catch (e) {
            // Error extracting text - skip this shape
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
  
  console.log(`[ProcessWithAssistant] Starting processing with scope: ${scope}`);
  
  try {
    await PowerPoint.run(async (context) => {
      let textBlocksToProcess = [];

    console.log('[ProcessWithAssistant] Initializing text blocks collection');

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
          console.log(`[ProcessWithAssistant] Found ${selectedShapes.items.length} selected shapes, finding their slide`);

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
                console.log(`[ProcessWithAssistant] Found slide from selected shape: slide ${i + 1}`);
                break;
              }
            }
          }
        }
      } catch (e) {
        console.log('[ProcessWithAssistant] Could not determine slide from selected shapes:', e.message);
      }

      // Method 2: Try getActiveSlide() (desktop PowerPoint)
      if (targetSlideIndex === -1) {
        try {
          console.log('[ProcessWithAssistant] Attempting to get active slide');
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
              console.log(`[ProcessWithAssistant] Active slide is at index ${i} (slide ${i + 1})`);
              break;
            }
          }
        } catch (e) {
          console.log('[ProcessWithAssistant] getActiveSlide not available:', e.message);
        }
      }

      // Method 3: Fallback to getSelectedSlides() (from thumbnail panel)
      if (targetSlideIndex === -1) {
        try {
          console.log('[ProcessWithAssistant] Trying getSelectedSlides (from thumbnail panel)');
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
                console.log(`[ProcessWithAssistant] Selected slide from thumbnail panel is at index ${i} (slide ${i + 1})`);
                break;
              }
            }
          }
        } catch (e) {
          console.log('[ProcessWithAssistant] getSelectedSlides also failed:', e.message);
        }
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

    } else if (scope === "presentation") {
      // Extract text from all slides using helper function
      document.getElementById("status").innerHTML = `<div class="spinner"></div> Extracting from entire presentation...`;
      textBlocksToProcess = await extractTextFromAllSlides(context);

      if (textBlocksToProcess.length > 0) {
        document.getElementById("status").innerHTML = `<div class="spinner"></div> Processing ${textBlocksToProcess.length} text blocks...`;
      }
    }

    // Check if we have text blocks to process
    if (textBlocksToProcess.length === 0) {
      console.log('[ProcessWithAssistant] No text blocks found to process');
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
      return; // Exit gracefully instead of throwing error
    }
    console.log(`[ProcessWithAssistant] Total text blocks to process: ${textBlocksToProcess.length}`);

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

        // Log why we're returning null
        console.log(`[ProcessWithAssistant] No agent response for text block ${index}:`,
          !lastAgentMessage ? 'No agent message found' :
          !lastAgentMessage.content ? 'Agent message has no content' :
          'Unknown reason');
        return null;
      } catch (error) {
        console.error(`Error processing text block ${index}: ${error.message}`);
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

    console.log(`[ProcessWithAssistant] Processing complete. Results: ${processedResults.length}`);
    if (processedResults.length > 0) {
      console.log('[ProcessWithAssistant] First result sample:', {
        shapeId: processedResults[0].shapeId,
        slideIndex: processedResults[0].slideIndex,
        originalTextLength: processedResults[0].originalText?.length,
        newTextLength: processedResults[0].newText?.length,
        newTextPreview: processedResults[0].newText?.substring(0, 100)
      });
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
  
  // Now apply all the results back to PowerPoint
  if (processedResults.length > 0 && !processingCancelled) {
    console.log(`[ProcessWithAssistant] Applying ${processedResults.length} processed results`);
    document.getElementById("status").innerHTML = `<div class="spinner"></div> Updating presentation...`;

    try {
      await PowerPoint.run(async (context) => {
        console.log('[ProcessWithAssistant] Starting PowerPoint context for updates');

        // Use the updateShapes helper function
        const { updatedCount, failedCount } = await updateShapes(context, processedResults);

        // Show final status
        if (updatedCount > 0) {
          document.getElementById("status").innerHTML = `✅ Successfully updated ${updatedCount} text block(s)${failedCount > 0 ? ` (${failedCount} failed)` : ''}`;
        } else {
          document.getElementById("status").innerHTML = `❌ Failed to update text blocks`;
        }
      });
    } catch (updateError) {
      console.error('[ProcessWithAssistant] Error during update phase:', updateError);
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
    // Use the reliable extraction function
    const textBlocks = await extractTextFromAllSlides(context);
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
