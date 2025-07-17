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

    // Get the actual start and end positions
    const startRow = range.getRow();
    const startCol = range.getColumn();
    const lastRow = range.getLastRow();
    const lastCol = range.getLastColumn();

    // Calculate dimensions explicitly
    const numRows = lastRow - startRow + 1;
    const numColumns = lastCol - startCol + 1;

    return {
      success: true,
      numColumns: numColumns,
      numRows: numRows,
      hasMultipleColumns: numColumns > 1,
      debug: {
        startRow: startRow,
        lastRow: lastRow,
        startCol: startCol,
        lastCol: lastCol,
        rangeNotation: rangeA1Notation,
      },
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
    };
  }
}

// Helper function to get the actual value for a cell, handling merged cells
function getCellValue(sheet, row, col) {
  try {
    const cell = sheet.getRange(row, col);
    const mergedRanges = sheet.getRange(row, col, 1, 1).getMergedRanges();

    if (mergedRanges.length > 0) {
      // This cell is part of a merged range, get the value from the top-left cell
      const mergedRange = mergedRanges[0];
      return mergedRange.getCell(1, 1).getValue();
    } else {
      // Regular cell
      return cell.getValue();
    }
  } catch (error) {
    console.error("Error getting cell value:", error);
    return "";
  }
}

// Helper function to get values from a range, handling merged cells
function getValuesWithMergedCells(sheet, startRow, startCol, numRows, numCols) {
  const values = [];

  for (let row = 0; row < numRows; row++) {
    const rowValues = [];
    for (let col = 0; col < numCols; col++) {
      const cellValue = getCellValue(sheet, startRow + row, startCol + col);
      rowValues.push(cellValue);
    }
    values.push(rowValues);
  }

  return values;
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
    '<link href="https://fonts.googleapis.com/css2?family=Geist:wght@400;500;600&display=swap" rel="stylesheet">' +
    '<link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />' +
    '<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>' +
    '<script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>' +
    "<style>" +
    "* {" +
    "font-family: 'Geist', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;" +
    "box-sizing: border-box;" +
    "}" +
    "body {" +
    "margin: 0;" +
    "padding: 0;" +
    "background-color: #f8f9fa;" +
    "}" +
    ".container {" +
    "padding: 24px;" +
    "background-color: white;" +
    "min-height: 100vh;" +
    "}" +
    ".logo {" +
    "display: flex;" +
    "justify-content: start;" +
    "}" +
    ".main-title {" +
    "font-size: 18px;" +
    "font-weight: 600;" +
    "color: #1f2937;" +
    "line-height: 1.2;" +
    "}" +
    ".form-group {" +
    "margin-bottom: 24px;" +
    "}" +
    ".form-label {" +
    "display: block;" +
    "font-weight: 500;" +
    "color: #374151;" +
    "margin-bottom: 8px;" +
    "font-size: 14px;" +
    "}" +
    ".search-container {" +
    "position: relative;" +
    "}" +
    ".search-input {" +
    "width: 100%;" +
    "padding: 12px 16px;" +
    "padding-right: 48px;" +
    "border: 1px solid #d1d5db;" +
    "border-radius: 8px;" +
    "font-size: 14px;" +
    "background-color: #f9fafb;" +
    "}" +
    ".search-icon {" +
    "position: absolute;" +
    "right: 16px;" +
    "top: 50%;" +
    "transform: translateY(-50%);" +
    "color: #9ca3af;" +
    "}" +
    ".input-group {" +
    "display: flex;" +
    "gap: 8px;" +
    "align-items: flex-end;" +
    "}" +
    ".input-group input {" +
    "flex: 1;" +
    "padding: 12px 16px;" +
    "border: 1px solid #d1d5db;" +
    "border-radius: 8px;" +
    "font-size: 14px;" +
    "width:100px" +
    "}" +
    "#headerRow {" +
    "padding: 12px 16px;" +
    "border: 1px solid #d1d5db;" +
    "border-radius: 8px;" +
    "font-size: 14px;" +
    "width: 100px;" +
    "}" +
    ".btn {" +
    "padding: 12px 20px;" +
    "border: none;" +
    "border-radius: 8px;" +
    "font-size: 14px;" +
    "font-weight: 500;" +
    "cursor: pointer;" +
    "transition: all 0.2s ease;" +
    "}" +
    ".btn-secondary {" +
    "background-color: #f3f4f6;" +
    "color: #374151;" +
    "}" +
    ".btn-secondary:hover {" +
    "background-color: #e5e7eb;" +
    "}" +
    ".btn-primary {" +
    "background-color: #3b82f6;" +
    "color: white;" +
    "width: 100%;" +
    "padding: 16px;" +
    "font-size: 16px;" +
    "margin-top: 24px;" +
    "}" +
    ".btn-primary:hover {" +
    "background-color: #2563eb;" +
    "}" +
    ".btn-primary:disabled {" +
    "background-color: #9ca3af;" +
    "cursor: not-allowed;" +
    "}" +
    "textarea {" +
    "width: 100%;" +
    "padding: 12px 16px;" +
    "border: 1px solid #d1d5db;" +
    "border-radius: 8px;" +
    "font-size: 14px;" +
    "resize: vertical;" +
    "min-height: 80px;" +
    "}" +
    ".spinner {" +
    "display: inline-block;" +
    "width: 16px;" +
    "height: 16px;" +
    "border: 2px solid #f3f3f3;" +
    "border-top: 2px solid #3b82f6;" +
    "border-radius: 50%;" +
    "animation: spin 1s linear infinite;" +
    "margin-right: 8px;" +
    "}" +
    "@keyframes spin {" +
    "0% { transform: rotate(0deg); }" +
    "100% { transform: rotate(360deg); }" +
    "}" +
    "#status {" +
    "text-align: center;" +
    "margin: 16px 0;" +
    "color: #6b7280;" +
    "font-size: 14px;" +
    "}" +
    ".select2-container {" +
    "width: 100% !important;" +
    "}" +
    ".select2-selection {" +
    "height: 48px !important;" +
    "border: 1px solid #d1d5db !important;" +
    "border-radius: 8px !important;" +
    "background-color: #f9fafb !important;" +
    "}" +
    ".select2-selection__rendered {" +
    "line-height: 46px !important;" +
    "padding-left: 16px !important;" +
    "}" +
    ".select2-selection__arrow {" +
    "height: 46px !important;" +
    "}" +
    ".error {" +
    "color: #dc2626;" +
    "font-size: 14px;" +
    "margin-top: 8px;" +
    "display: none;" +
    "}" +
    ".info-text {" +
    "font-size: 12px;" +
    "color: #6b7280;" +
    "margin-top: 4px;" +
    "}" +
    ".header-row-section {" +
    "display: none;" +
    "}" +
    "</style>" +
    '<div class="container">' +
    '<div class="logo">' +
    '<svg width="96" height="24" viewBox="0 0 96 24" fill="none" xmlns="http://www.w3.org/2000/svg">' +
    '<path d="M84 0H72V24H84V0Z" fill="#FE9C1A"/>' +
    '<path d="M12 24C18.6274 24 24 18.6274 24 12C24 5.37258 18.6274 0 12 0C5.37258 0 0 5.37258 0 12C0 18.6274 5.37258 24 12 24Z" fill="#E2F78C"/>' +
    '<path d="M36 24C42.6274 24 48 18.6274 48 12C48 5.37258 42.6274 0 36 0C29.3726 0 24 5.37258 24 12C24 18.6274 29.3726 24 36 24Z" fill="#FFC3DF"/>' +
    '<path d="M12 0H0V24H12V0Z" fill="#418B5C"/>' +
    '<path d="M48 0H24V12H48V0Z" fill="#E14322"/>' +
    '<path fill-rule="evenodd" clip-rule="evenodd" d="M60 12C56.6863 12 54 9.31371 54 6C54 2.68629 56.6863 0 60 0H96V12H60Z" fill="#3B82F6"/>' +
    '<path fill-rule="evenodd" clip-rule="evenodd" d="M48 24V12H60C63.3137 12 66 14.6863 66 18C66 21.3137 63.3137 24 60 24H48Z" fill="#9FDBFF"/>' +
    "</svg>" +
    "</div>" +
    '<h2 class="main-title">Use AI Agents in your<br>spreadsheet</h2>' +
    '<form id="myForm">' +
    '<div class="form-group">' +
    '<label class="form-label">Select agent</label>' +
    '<div class="search-container">' +
    '<select id="assistant" name="assistant" required disabled>' +
    '<option value=""></option>' +
    "</select>" +
    '<div class="search-icon">🔍</div>' +
    "</div>" +
    '<div id="loadError" class="error">Failed to load agents</div>' +
    "</div>" +
    '<div class="form-group">' +
    '<label class="form-label">Select input range</label>' +
    '<div class="input-group">' +
    '<input type="text" id="cellRange" name="cellRange" required placeholder="A1:A10">' +
    '<button type="button" class="btn btn-secondary" id="selectCellsBtn">Use selection</button>' +
    "</div>" +
    '<div id="rangeInfo" class="info-text"></div>' +
    "</div>" +
    '<div id="headerRowSection" class="header-row-section form-group">' +
    '<label class="form-label">Header Row Number:</label>' +
    '<input type="number" id="headerRow" name="headerRow" value="1" min="1">' +
    '<div class="info-text">Row number containing column headers (default: 1)</div>' +
    "</div>" +
    '<div class="form-group">' +
    '<label class="form-label">Select output column</label>' +
    '<div class="input-group">' +
    '<input type="text" id="targetColumn" name="targetColumn" required placeholder="B">' +
    '<button type="button" class="btn btn-secondary">Use selection</button>' +
    "</div>" +
    "</div>" +
    '<div class="form-group">' +
    '<label class="form-label">Additional instructions (optional)</label>' +
    '<textarea id="instructions" name="instructions" placeholder="Summarize each row in 2 sentences max"></textarea>' +
    "</div>" +
    '<div id="status"></div>' +
    '<button type="submit" class="btn btn-primary" id="submitBtn">Run</button>' +
    "</form>" +
    "</div>" +
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
    "if (result.debug) {" +
    "console.log('Range debug:', result.debug);" +
    "}" +
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
    "selectCellsBtn.textContent = 'Loading...';" +
    "google.script.run" +
    ".withSuccessHandler(function(result) {" +
    "selectCellsBtn.disabled = false;" +
    "selectCellsBtn.textContent = 'Use selection';" +
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
    "selectCellsBtn.textContent = 'Use selection';" +
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
    "$('#assistant').select2({" +
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
    "const select = document.getElementById('assistant');" +
    "if (data.error) {" +
    "const errorDiv = document.getElementById('loadError');" +
    "errorDiv.textContent = '❌ ' + data.error;" +
    "errorDiv.style.display = 'block';" +
    "$('#assistant').select2({" +
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
    "data.assistants.forEach(function(a) {" +
    "const option = document.createElement('option');" +
    "option.value = a.id;" +
    "option.textContent = a.name;" +
    "select.appendChild(option);" +
    "});" +
    "select.disabled = false;" +
    "$('#assistant').select2({" +
    "placeholder: 'Select an agent', " +
    "allowClear: true, " +
    "width: '100%', " +
    "language: {" +
    "noResults: function() {" +
    "return 'No agents found';" +
    "}" +
    "}" +
    "});" +
    "if (data.assistants.length === 0) {" +
    "$('#assistant').select2({" +
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
    "$('#assistant').select2({" +
    "placeholder: 'Failed to load agents', " +
    "allowClear: true, " +
    "width: '100%'" +
    "});" +
    "})" +
    ".fetchAssistants();" +
    "document.getElementById('myForm').addEventListener('submit', function(e) {" +
    "e.preventDefault();" +
    "const assistantSelect = document.getElementById('assistant');" +
    "const cellRange = document.getElementById('cellRange');" +
    "if (assistantSelect.disabled) {" +
    "alert('Please wait for agents to load');" +
    "return;" +
    "}" +
    "if (!assistantSelect.value) {" +
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
    "document.getElementById('status').innerHTML = '<div class=\"spinner\"></div> Processing...';" +
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
    "'<div class=\"spinner\"></div> ' + result.progress;" +
    "}" +
    "})" +
    ".withFailureHandler(function(error) {" +
    "document.getElementById('submitBtn').disabled = false;" +
    "document.getElementById('status').textContent = '❌ Error: ' + error;" +
    "})" +
    ".processWithAssistant(" +
    "assistantSelect.value, " +
    "document.getElementById('instructions').value, " +
    "cellRange.value, " +
    "document.getElementById('targetColumn').value, " +
    "headerRow" +
    ");" +
    "});" +
    "</script>";

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setTitle("Dust")
    .setWidth(400);
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

function fetchAssistants() {
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

    const assistants = JSON.parse(
      response.getContentText()
    ).agentConfigurations;

    const sortedAssistants = assistants.sort(function (a, b) {
      return a.name.localeCompare(b.name);
    });

    return {
      assistants: sortedAssistants.map(function (a) {
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

function processWithAssistant(
  assistantId,
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

  // Use the new merged cell handling function
  const selectedValues = getValuesWithMergedCells(
    sheet,
    selected.getRow(),
    selected.getColumn(),
    selected.getNumRows(),
    selected.getNumColumns()
  );

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
      // If header row is outside the selected range, get it from the sheet with merged cell handling
      headers = getValuesWithMergedCells(
        sheet,
        headerRow,
        selected.getColumn(),
        1,
        numColumns
      )[0];
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
        mentions: [{ configurationId: assistantId }],
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
      skipToolsValidation: true,
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
