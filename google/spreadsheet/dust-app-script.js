function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Dust")
    .addItem("Call an Agent", "processSelected")
    .addItem("Setup", "showCredentialsDialog")
    .addToUi();
}

function showSelectionToast() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Select your input cells, then click the "Call an Agent" menu item again',
    "Select Cells",
    -1 // Show indefinitely
  );
}

function showCredentialsDialog() {
  var ui = SpreadsheetApp.getUi();
  var docProperties = PropertiesService.getDocumentProperties();

  var result = ui.prompt(
    "Setup Dust",
    "Your Dust Workspace ID:",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    docProperties.setProperty("workspaceId", result.getResponseText());
  }

  var result = ui.prompt(
    "Setup Dust",
    "Your Dust API Key:",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    docProperties.setProperty("dustToken", result.getResponseText());
  }

  var result = ui.prompt(
    "Setup Dust",
    "Region (optional, leave empty for US or enter 'eu' for EU region):",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    docProperties.setProperty("region", result.getResponseText());
  }
}

function handleCellSelection() {
  try {
    const selectedRange = SpreadsheetApp.getSelection().getActiveRange();
    return selectedRange
      ? {
          range: selectedRange.getA1Notation(),
          success: true,
        }
      : {
          success: false,
        };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
    };
  }
}

function storeFormData(formData) {
  PropertiesService.getUserProperties().setProperty(
    "tempFormData",
    JSON.stringify(formData)
  );
}

function reopenModalWithRange(result, formData) {
  if (result.success) {
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty("selectedRange", result.range);
    userProps.setProperty("tempFormData", JSON.stringify(formData));
  }
  processSelected(true);
}

function analyzeSelectedRange(rangeA1Notation) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getRange(rangeA1Notation);
    const numColumns = range.getNumColumns();
    const numRows = range.getNumRows();

    return {
      success: true,
      numColumns: numColumns,
      numRows: numRows,
      hasMultipleColumns: numColumns > 1,
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
    };
  }
}

function processSelected() {
  const docProperties = PropertiesService.getDocumentProperties();
  const token = docProperties.getProperty("dustToken");
  const workspaceId = docProperties.getProperty("workspaceId");

  if (!token || !workspaceId) {
    SpreadsheetApp.getUi().alert(
      "Please configure your Dust credentials first"
    );
    return;
  }

  var htmlContent =
    "" +
    '<link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />' +
    '<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>' +
    '<script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>' +
    "<style>" +
    "* {" +
    "font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;" +
    "}" +
    "body {" +
    "padding: 10px;" +
    "}" +
    ".spinner {" +
    "display: inline-block;" +
    "width: 12px;" +
    "height: 12px;" +
    "border: 3px solid #f3f3f3;" +
    "border-top: 3px solid #3498db;" +
    "border-radius: 50%;" +
    "animation: spin 1s linear infinite;" +
    "margin-right: 8px;" +
    "vertical-align: middle;" +
    "}" +
    "@keyframes spin {" +
    "0% { transform: rotate(0deg); }" +
    "100% { transform: rotate(360deg); }" +
    "}" +
    "#status {" +
    "text-align: center;" +
    "margin: 15px 0;" +
    "color: #666;" +
    "line-height: 1.4;" +
    "}" +
    "select, input, textarea {" +
    "width: 100%;" +
    "padding-top: 8px;" +
    "padding-bottom: 8px;" +
    "margin-bottom: 10px;" +
    "border: 1px solid #ddd;" +
    "border-radius: 4px;" +
    "font-family: inherit;" +
    "}" +
    "select {" +
    "background: white;" +
    "}" +
    'input[type="submit"], input[type="button"] {' +
    "background-color: #61A5FA;" +
    "color: white;" +
    "border: none;" +
    "padding: 10px 20px;" +
    "cursor: pointer;" +
    "transition: background-color 0.2s ease;" +
    "}" +
    'input[type="submit"]:hover, input[type="button"]:hover {' +
    "background-color: #4884d9;" +
    "}" +
    'input[type="submit"]:disabled {' +
    "background-color: #cccccc;" +
    "cursor: not-allowed;" +
    "}" +
    ".error {" +
    "color: red;" +
    "display: none;" +
    "margin-bottom: 10px;" +
    "}" +
    "label {" +
    "font-weight: 500;" +
    "color: #333;" +
    "margin-bottom: 5px;" +
    "display: inline-block;" +
    "}" +
    ".select2-container {" +
    "width: 100% !important;" +
    "margin-bottom: 10px;" +
    "}" +
    ".select2-selection {" +
    "height: 38px !important;" +
    "padding: 4px !important;" +
    "}" +
    ".select2-selection__arrow {" +
    "height: 36px !important;" +
    "}" +
    "#cellRange {" +
    "width: calc(100% - 90px);" +
    "display: inline-block;" +
    "margin-right: 5px;" +
    "}" +
    'input[type="submit"], input[type="button"], #selectCellsBtn {' +
    "background-color: #61A5FA;" +
    "color: white;" +
    "border: none;" +
    "padding: 10px 20px;" +
    "cursor: pointer;" +
    "transition: background-color 0.2s ease;" +
    "}" +
    'input[type="submit"]:hover, input[type="button"]:hover, #selectCellsBtn:hover {' +
    "background-color: #4884d9;" +
    "}" +
    "#selectCellsBtn {" +
    "width: 85px;" +
    "padding: 8px 5px;" +
    "border-radius: 4px;" +
    "display: inline-block;" +
    "}" +
    "#targetColumn {" +
    "width: 100%;" +
    "}" +
    'input[type="submit"] {' +
    "width: 100%;" +
    "margin-top: 10px;" +
    "}" +
    ".button-group {" +
    "display: block;" +
    "margin-top: 20px;" +
    "}" +
    "textarea {" +
    "width: 99%;" +
    "resize: vertical;" +
    "}" +
    ".header-row-section {" +
    "display: none;" +
    "margin-bottom: 10px;" +
    "}" +
    ".info-text {" +
    "font-size: 12px;" +
    "color: #666;" +
    "margin-top: 5px;" +
    "font-style: italic;" +
    "}" +
    "</style>" +
    '<form id="myForm">' +
    '<div style="margin-bottom: 10px;">' +
    '<label for="agent">Agent:</label><br>' +
    '<select id="agent" name="agent" required disabled>' +
    '<option value=""></option>' +
    "</select>" +
    '<div id="loadError" class="error">Failed to load agents</div>' +
    "</div>" +
    '<div style="margin-bottom: 10px;">' +
    '<label for="cellRange">Input Cells:</label><br>' +
    '<input type="text" id="cellRange" name="cellRange" required placeholder="e.g., A1:A10" style="width:120px">' +
    '<input type="button" value="Use Selection" id="selectCellsBtn" style="width:120px">' +
    '<div id="rangeInfo" class="info-text"></div>' +
    "</div>" +
    '<div id="headerRowSection" class="header-row-section">' +
    '<label for="headerRow">Header Row Number:</label><br>' +
    '<input type="number" id="headerRow" name="headerRow" value="1" min="1" style="width:120px">' +
    '<div class="info-text">Row number containing column headers (default: 1)</div>' +
    "</div>" +
    '<div style="margin-bottom: 10px;">' +
    '<label for="targetColumn">Target Column:</label><br>' +
    '<input type="text" id="targetColumn" name="targetColumn" required placeholder="e.g., B" style="width:120px">' +
    "</div>" +
    '<div style="margin-bottom: 10px;">' +
    '<label for="instructions">Instructions (optional):</label><br>' +
    '<textarea id="instructions" name="instructions" rows="4" style="width:99%"></textarea>' +
    "</div>" +
    '<div id="status"></div>' +
    '<div class="button-group">' +
    '<input type="submit" value="Process" id="submitBtn">' +
    "</div>" +
    "</form>" +
    "<script>" +
    "function debounce(func, wait) {" +
    "let timeout;" +
    "return function executedFunction(...args) {" +
    "const later = () => {" +
    "clearTimeout(timeout);" +
    "func(...args);" +
    "};" +
    "clearTimeout(timeout);" +
    "timeout = setTimeout(later, wait);" +
    "};" +
    "}" +
    "const debouncedGetSelection = debounce(getSelection, 250);" +
    "function updateRangeInfo(rangeNotation) {" +
    "if (!rangeNotation) {" +
    "document.getElementById('rangeInfo').textContent = '';" +
    "document.getElementById('headerRowSection').style.display = 'none';" +
    "return;" +
    "}" +
    "google.script.run" +
    ".withSuccessHandler(function(result) {" +
    "const infoDiv = document.getElementById('rangeInfo');" +
    "const headerSection = document.getElementById('headerRowSection');" +
    "if (result.success) {" +
    "if (result.hasMultipleColumns) {" +
    "infoDiv.textContent = 'Selected: ' + result.numRows + ' rows × ' + result.numColumns + ' columns';" +
    "headerSection.style.display = 'block';" +
    "} else {" +
    "infoDiv.textContent = 'Selected: ' + result.numRows + ' rows × ' + result.numColumns + ' column';" +
    "headerSection.style.display = 'none';" +
    "}" +
    "} else {" +
    "infoDiv.textContent = 'Invalid range';" +
    "headerSection.style.display = 'none';" +
    "}" +
    "})" +
    ".withFailureHandler(function(error) {" +
    "document.getElementById('rangeInfo').textContent = 'Error analyzing range';" +
    "document.getElementById('headerRowSection').style.display = 'none';" +
    "})" +
    ".analyzeSelectedRange(rangeNotation);" +
    "}" +
    "function getSelection() {" +
    "const selectCellsBtn = document.getElementById('selectCellsBtn');" +
    "const cellRangeInput = document.getElementById('cellRange');" +
    "selectCellsBtn.disabled = true;" +
    "selectCellsBtn.value = '...';" +
    "google.script.run" +
    ".withSuccessHandler(function(result) {" +
    "selectCellsBtn.disabled = false;" +
    "selectCellsBtn.value = 'Use Selection';" +
    "if (result.success) {" +
    "cellRangeInput.value = result.range;" +
    "updateRangeInfo(result.range);" +
    "} else {" +
    "if (result.error) {" +
    "console.error('Selection error:', result.error);" +
    "}" +
    "if (!cellRangeInput.value) {" +
    'alert("Please select some cells first");' +
    "}" +
    "}" +
    "})" +
    ".withFailureHandler(function(error) {" +
    "selectCellsBtn.disabled = false;" +
    "selectCellsBtn.value = 'Use Selection';" +
    "console.error('Selection failed:', error);" +
    "})" +
    ".handleCellSelection();" +
    "}" +
    "document.getElementById('selectCellsBtn').addEventListener('click', function() {" +
    "debouncedGetSelection();" +
    "});" +
    "document.getElementById('cellRange').addEventListener('input', function() {" +
    "updateRangeInfo(this.value);" +
    "});" +
    "function onLoad() {" +
    "google.script.run" +
    ".withSuccessHandler(function() {})" +
    ".withFailureHandler(function(error) {})" +
    ".getCurrentSelection();" +
    "}" +
    "$(document).ready(function() {" +
    "onLoad();" +
    "$('#agents').select2({" +
    "placeholder: 'Loading agents...', " +
    "allowClear: true, " +
    "width: '100%', " +
    "language: {" +
    "noResults: function() {" +
    "return 'No agents found';" +
    "}" +
    "}" +
    "});" +
    "});" +
    "google.script.run" +
    ".withSuccessHandler(function(data) {" +
    "const select = document.getElementById('agent');" +
    "if (data.error) {" +
    "const errorDiv = document.getElementById('loadError');" +
    "errorDiv.textContent = '❌ ' + data.error;" +
    "errorDiv.style.display = 'block';" +
    "$('#agent').select2({" +
    "placeholder: 'Failed to load agents', " +
    "allowClear: true, " +
    "width: '100%'" +
    "});" +
    "return;" +
    "}" +
    "select.innerHTML = '';" +
    "const emptyOption = document.createElement('option');" +
    "emptyOption.value = '';" +
    "select.appendChild(emptyOption);" +
    "data.agents.forEach(function(a) {" +
    "const option = document.createElement('option');" +
    "option.value = a.id;" +
    "option.textContent = a.name;" +
    "select.appendChild(option);" +
    "});" +
    "select.disabled = false;" +
    "$('#agent').select2({" +
    "placeholder: 'Select an agent', " +
    "allowClear: true, " +
    "width: '100%', " +
    "language: {" +
    "noResults: function() {" +
    "return 'No agents found';" +
    "}" +
    "}" +
    "});" +
    "if (data.agents.length === 0) {" +
    "$('#agents').select2({" +
    "placeholder: 'No agents available', " +
    "allowClear: true, " +
    "width: '100%'" +
    "});" +
    "}" +
    "})" +
    ".withFailureHandler(function(error) {" +
    "const errorDiv = document.getElementById('loadError');" +
    "errorDiv.textContent = '❌ ' + error;" +
    "errorDiv.style.display = 'block';" +
    "$('#agents').select2({" +
    "placeholder: 'Failed to load agents', " +
    "allowClear: true, " +
    "width: '100%'" +
    "});" +
    "})" +
    ".fetchAgents();" +
    "document.getElementById('myForm').addEventListener('submit', function(e) {" +
    "e.preventDefault();" +
    "const agentSelect = document.getElementById('agent');" +
    "const cellRange = document.getElementById('cellRange');" +
    "if (agentSelect.disabled) {" +
    "alert('Please wait for agents to load');" +
    "return;" +
    "}" +
    "if (!agentSelect.value) {" +
    "alert('Please select an agent');" +
    "return;" +
    "}" +
    "if (!cellRange.value) {" +
    "alert('Please select input cells');" +
    "return;" +
    "}" +
    "const targetColumn = document.getElementById('targetColumn').value;" +
    "if (!/^[A-Za-z]+$/.test(targetColumn)) {" +
    "alert('Please enter a valid target column letter (e.g., A, B, C)');" +
    "return;" +
    "}" +
    "const headerRow = parseInt(document.getElementById('headerRow').value) || 1;" +
    "document.getElementById('submitBtn').disabled = true;" +
    "document.getElementById('status').innerHTML = '<div id=\"spinner\" class=\"spinner\"></div> Processing...';" +
    "google.script.run" +
    ".withSuccessHandler(function(result) {" +
    "if (result.completed) {" +
    "document.getElementById('submitBtn').disabled = false;" +
    "document.getElementById('status').innerHTML = '✅ Processing complete';" +
    "setTimeout(function() {" +
    "document.getElementById('status').innerHTML = '';" +
    "}, 3000);" +
    "} else if (result.progress) {" +
    "document.getElementById('status').innerHTML = " +
    '\'<div id="spinner" class="spinner"></div> \' + result.progress;' +
    "}" +
    "})" +
    ".withFailureHandler(function(error) {" +
    "document.getElementById('submitBtn').disabled = false;" +
    "document.getElementById('status').textContent = '❌ Error: ' + error;" +
    "})" +
    ".processWithAgent(" +
    "agentSelect.value, " +
    "document.getElementById('instructions').value, " +
    "cellRange.value, " +
    "document.getElementById('targetColumn').value, " +
    "headerRow" +
    ");" +
    "});" +
    "</script>";

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setTitle("Dust");
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getCurrentSelection() {
  try {
    const selection = SpreadsheetApp.getActiveRange();
    if (selection) {
      return selection.getA1Notation();
    }
    SpreadsheetApp.getUi().alert("Please select some cells first");
    return null;
  } catch (error) {
    console.error("Error getting selection:", error);
    return null;
  }
}

function fetchAgents() {
  const docProperties = PropertiesService.getDocumentProperties();
  const token = docProperties.getProperty("dustToken");
  const workspaceId = docProperties.getProperty("workspaceId");

  if (!token || !workspaceId) {
    SpreadsheetApp.getUi().alert("Please configure Dust credentials first");
    return;
  }

  try {
    const BASE_URL = getDustBaseUrl() + "/api/v1/w/" + workspaceId;
    const response = UrlFetchApp.fetch(
      BASE_URL + "/assistant/agent_configurations",
      {
        method: "get",
        headers: {
          Authorization: "Bearer " + token,
        },
        muteHttpExceptions: true,
      }
    );

    if (response.getResponseCode() !== 200) {
      return { error: "API returned " + response.getResponseCode() };
    }

    const agents = JSON.parse(
      response.getContentText()
    ).agentConfigurations;

    const sortedAgents = agents.sort(function (a, b) {
      return a.name.localeCompare(b.name);
    });

    return {
      agents: sortedAgents.map(function (a) {
        return {
          id: a.sId,
          name: a.name,
        };
      }),
    };
  } catch (error) {
    return { error: error.toString() };
  }
}

function processWithAgent(
  agentId,
  instructions,
  rangeA1Notation,
  targetColumn,
  headerRow
) {
  headerRow = headerRow || 1;
  const BATCH_SIZE = 10;
  const BATCH_DELAY = 1000;
  const docProperties = PropertiesService.getDocumentProperties();
  const token = docProperties.getProperty("dustToken");
  const workspaceId = docProperties.getProperty("workspaceId");

  if (!token || !workspaceId) {
    throw new Error("Please configure your Dust credentials first");
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const selected = sheet.getRange(rangeA1Notation);
  const targetColIndex = columnToIndex(targetColumn);

  if (!targetColIndex) {
    throw new Error(
      "Invalid target column. Please enter a valid column letter (e.g., A, B, C)"
    );
  }

  const BASE_URL = getDustBaseUrl() + "/api/v1/w/" + workspaceId;
  const selectedValues = selected.getValues();
  const numColumns = selected.getNumColumns();
  const numRows = selected.getNumRows();
  const startRow = selected.getRow();

  // Get headers if multiple columns
  var headers = [];
  if (numColumns > 1) {
    const headerRowIndex = headerRow - startRow;
    if (headerRowIndex >= 0 && headerRowIndex < numRows) {
      headers = selectedValues[headerRowIndex];
    } else {
      // If header row is outside the selected range, get it from the sheet
      const headerRange = sheet.getRange(
        headerRow,
        selected.getColumn(),
        1,
        numColumns
      );
      headers = headerRange.getValues()[0];
    }
  }

  const cellsToProcess = [];

  for (var i = 0; i < numRows; i++) {
    const currentRow = startRow + i;

    // Skip the header row if it's within our selection
    if (numColumns > 1 && currentRow === headerRow) {
      continue;
    }

    const targetCell = sheet.getRange(currentRow, targetColIndex);

    var inputContent = "";

    if (numColumns === 1) {
      // Single column - just use the value
      const inputValue = selectedValues[i][0];
      if (!inputValue) {
        targetCell.setValue("No input value");
        continue;
      }
      inputContent = inputValue.toString();
    } else {
      // Multiple columns - combine with headers
      const rowValues = selectedValues[i];
      const contentParts = [];

      for (var j = 0; j < numColumns; j++) {
        const header = headers[j] || "Column " + (j + 1);
        const value = rowValues[j] || "";
        contentParts.push(header + ": " + value);
      }

      inputContent = contentParts.join("\n");

      if (!inputContent.trim()) {
        targetCell.setValue("No input value");
        continue;
      }
    }

    const payload = {
      message: {
        content: (instructions || "") + "\n\nInput:\n" + inputContent,
        mentions: [{ configurationId: agentId }],
        context: {
          username: "gsheet",
          timezone: Session.getScriptTimeZone(),
          fullName: "Google Sheets",
          email: "gsheet@dust.tt",
          profilePictureUrl: "",
          origin: "gsheet",
        },
      },
      blocking: true,
      title: "Google Sheets Conversation",
      visibility: "unlisted",
    };

    const request = {
      url: BASE_URL + "/assistant/conversations",
      method: "post",
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json",
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    cellsToProcess.push({
      cell: targetCell,
      request: request,
    });
  }

  const totalCells = cellsToProcess.length;
  var processedCells = 0;

  const batches = [];
  for (var i = 0; i < cellsToProcess.length; i += BATCH_SIZE) {
    batches.push(cellsToProcess.slice(i, i + BATCH_SIZE));
  }

  for (var batchIndex = 0; batchIndex < batches.length; batchIndex++) {
    const batch = batches[batchIndex];
    const batchRequests = batch.map(function (item) {
      return item.request;
    });
    const responses = UrlFetchApp.fetchAll(batchRequests);

    responses.forEach(function (response, index) {
      const targetCell = batch[index].cell;
      try {
        const result = JSON.parse(response.getContentText());
        const content = result.conversation.content;
        const lastAgentMessage = content
          .flat()
          .reverse()
          .find(function (msg) {
            return msg.type === "agent_message";
          });

        const appUrl =
          "https://dust.tt/w/" +
          workspaceId +
          "/assistant/" +
          result.conversation.sId;
        targetCell.setValue(
          lastAgentMessage ? lastAgentMessage.content : "No response"
        );
        targetCell.setNote("View conversation on Dust: " + appUrl);
      } catch (error) {
        targetCell.setValue("Error: " + error.toString());
      }
      processedCells++;
    });

    const progress = Math.round((processedCells / totalCells) * 100);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Processed " +
        processedCells +
        "/" +
        totalCells +
        " cells (" +
        progress +
        "%)",
      "Progress",
      -1
    );

    if (batchIndex < batches.length - 1) {
      Utilities.sleep(BATCH_DELAY);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Completed: " + totalCells + "/" + totalCells + " cells (100%)",
    "Progress",
    3
  );

  return {
    completed: true,
  };
}

// Helper function to convert column letter to index
function columnToIndex(column) {
  if (!column || typeof column !== "string") return null;
  column = column.toUpperCase();
  var sum = 0;
  for (var i = 0; i < column.length; i++) {
    sum *= 26;
    sum += column.charCodeAt(i) - "A".charCodeAt(0) + 1;
  }
  return sum;
}

function getDustBaseUrl() {
  const docProperties = PropertiesService.getDocumentProperties();
  const region = docProperties.getProperty("region");
  if (region && region.toLowerCase() === "eu") {
    return "https://eu.dust.tt";
  }
  return "https://dust.tt";
}
