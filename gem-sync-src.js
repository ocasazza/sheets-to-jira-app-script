/*
                GOOGLE SHEETS TO JIRA TICKET CREATOR

WHAT THIS SCRIPT DOES:
This script automatically creates Jira tickets from rows in your Google Sheet.
It will process all rows that don't already have a Jira ticket linked to them.

HOW TO USE:
1. Configure the settings below (see detailed instructions for each section)
2. Run the function "createJiraTicketsForAllRows" from the Apps Script editor
3. Check your Google Sheet for results - successful tickets will have Jira links,
   failed rows will have error messages in the "Script Errors" column
*/

// =============================================================================
//                             JIRA CONNECTION SETTINGS
// =============================================================================
const JIRA_DOMAIN = "";
const JIRA_EMAIL = "";
const JIRA_API_TOKEN = "";
const PROJECT_KEY = "";
const ISSUE_TYPE = "Task";
const REPORTER_ID = "";
const ASSIGNEE_ID = "";
const ISSUE_LABEL = "google-sheet-automation"; // Default label for new issues


// =============================================================================
//                         GOOGLE SHEET COLUMN MAPPING
// =============================================================================
const FIELD_MAPPINGS = [{
    sheetColumnName: 'Timestamp',
    jiraFieldName: 'timestamp',
    failIfNotFound: true
  },
  {
    sheetColumnName: 'Title',
    jiraFieldName: 'summary',
    failIfNotFound: true
  },
  {
    sheetColumnName: 'Description',
    jiraFieldName: 'description',
    failIfNotFound: false
  },
  {
    sheetColumnName: 'Labels',
    jiraFieldName: 'labels',
    failIfNotFound: false
  },
  {
    sheetColumnName: 'Priority',
    jiraFieldName: 'priority',
    failIfNotFound: false
  },
  {
    sheetColumnName: 'Due Date',
    jiraFieldName: 'duedate',
    failIfNotFound: false
  },
  {
    sheetColumnName: 'Email Address',
    jiraFieldName: 'reporter',
    failIfNotFound: true
  },
];


// =============================================================================
//                             GOOGLE SHEET SETTINGS
// =============================================================================
const SPREADSHEET_ID = '1Grc8qre52YSS7U_YtMHheqIDvuhdMHIGsZuhZTc30b4';
const SHEET_NAME = 'Form Responses';
const TABLE_START_ROW = 2;
const TABLE_START_COLUMN = 1;
const IGNORED_COLUMNS = ['Jira Key', 'Script Errors'];


// --- PROGRAM ------------------------------------
// --- DO NOT EDIT BELOW HERE UNLESS DEVELOPING ---

/**
 * @typedef {Object} SheetData
 * @property {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @property {string[]} headers - An array of header values.
 * @property {any[][]} allRowsData - A 2D array of all row data.
 * @property {number} lastRowIndex - The index of the last row.
 * @property {number} lastColIndex - The index of the last column.
 * @property {number} tableStartRow - The starting row of the data table.
 * @property {number} tableStartColumn - The starting column of the data table.
 */

/**
 * Main function to create Jira tickets for all rows.
 */
function createJiraTicketsForAllRows() {
  Logger.log("Starting Jira ticket creation process...");
  try {
    const data = getSheetData();
    if (!data) return;

    const rowsToProcess = getRowsWithoutJiraTickets(data.allRowsData, data.headers)
      .map((rowData, index) => ({
        rowData,
        rowIndex: index + data.tableStartRow + 1
      }));


    if (rowsToProcess.length === 0) {
      Logger.log("No new rows to process.");
      return;
    }

    Logger.log(`Processing ${rowsToProcess.length} new rows...`);

    const results = rowsToProcess.map(rowInfo => processRow(rowInfo, data));

    const successCount = results.filter(result => result.success).length;
    const errorCount = results.length - successCount;

    Logger.log(`Process completed. Success: ${successCount}, Errors: ${errorCount}`);

  } catch (err) {
    Logger.log(`An unexpected error occurred: ${err.toString()}`);
  }
}

/**
 * Processes a single row from the sheet.
 * @param {Object} rowInfo - Information about the row to process.
 * @param {SheetData} data - The sheet data.
 * @returns {Object} An object indicating the success or failure of the operation.
 */
function processRow(rowInfo, data) {
  try {
    const {
      rowData,
      rowIndex
    } = rowInfo;
    const {
      headers,
      sheet
    } = data;
    const validation = validateRow(rowData, headers);

    if (!validation.success) {
      updateErrorColumn(sheet, rowIndex, headers, validation.error);
      return {
        success: false,
        error: validation.error
      };
    }

    const payload = buildJiraPayload(rowData, headers, rowIndex);
    const jiraResponse = createJiraIssue(payload);

    updateSheetWithJiraLink(sheet, rowIndex, headers.length, headers, jiraResponse);
    clearErrorColumn(sheet, rowIndex, headers);
    Logger.log(`Successfully created Jira ticket ${jiraResponse.key} for row ${rowIndex}`);
    return {
      success: true,
      key: jiraResponse.key
    };

  } catch (err) {
    const errorMsg = `Error processing row ${rowInfo.rowIndex}: ${err.toString()}`;
    Logger.log(errorMsg);
    updateErrorColumn(data.sheet, rowInfo.rowIndex, data.headers, errorMsg);
    return {
      success: false,
      error: errorMsg
    };
  }
}


/**
 * Validates a single row against the field mappings.
 * @param {any[]} rowData - The data for the row.
 * @param {string[]} headers - The sheet headers.
 * @returns {{success: boolean, error: string|null}}
 */
function validateRow(rowData, headers) {
  const errors = FIELD_MAPPINGS
    .filter(mapping => mapping.failIfNotFound)
    .map(mapping => {
      const colIndex = headers.indexOf(mapping.sheetColumnName);
      const hasValue = colIndex !== -1 && rowData[colIndex] && rowData[colIndex].toString().trim() !== '';
      return hasValue ? null : `Missing required field: ${mapping.sheetColumnName}`;
    })
    .filter(error => error !== null);

  return {
    success: errors.length === 0,
    error: errors.join('; ')
  };
}

/**
 * Builds the Jira issue payload for a given row.
 * @param {any[]} rowData - The data for the row.
 * @param {string[]} headers - The sheet headers.
 * @param {number} rowIndex - The index of the row.
 * @returns {Object} The Jira issue payload.
 */
function buildJiraPayload(rowData, headers, rowIndex) {
  // Extract all data from the sheet based on the mappings
  const mappedData = FIELD_MAPPINGS.reduce((acc, mapping) => {
    const colIndex = headers.indexOf(mapping.sheetColumnName);
    if (colIndex !== -1 && rowData[colIndex]) {
      acc[mapping.jiraFieldName] = rowData[colIndex].toString().trim();
    }
    return acc;
  }, {});

  // --- Explicitly define required Jira fields ---
  const summary = mappedData.summary || 'Ticket from Google Sheets';
  const priority = mappedData.priority || 'Major';
  const labels = mappedData.labels ? mappedData.labels.split(',').map(l => l.trim()) : [ISSUE_LABEL];
  const dueDateStr = mappedData.duedate ? Utilities.formatDate(new Date(mappedData.duedate), Session.getScriptTimeZone(), "yyyy-MM-dd") : null;

  // --- Build the description ---
  const mainDescription = mappedData.description || '';
  const timestampInfo = mappedData.timestamp ? `Timestamp: ${mappedData.timestamp}` : '';
  const reporterInfo = mappedData.reporter ? `Submitted by: ${mappedData.reporter}` : '';

  // Collect all other unmapped columns
  const unmappedInfo = headers
    .map((header, i) => {
      const isMapped = FIELD_MAPPINGS.some(m => m.sheetColumnName === header);
      const isIgnored = IGNORED_COLUMNS.includes(header);
      return (!isMapped && !isIgnored && rowData[i]) ? `${header}: ${rowData[i]}` : null;
    })
    .filter(Boolean)
    .join('\n');

  const fullDescriptionText = [mainDescription, timestampInfo, reporterInfo, unmappedInfo]
    .filter(Boolean) // Remove any empty parts
    .join('\n\n');

  const descriptionAdf = {
    type: "doc",
    version: 1,
    content: [{
      type: "paragraph",
      content: [{
        type: "text",
        text: fullDescriptionText || 'This ticket was created by the Google Sheets Apps Script automation.'
      }]
    }]
  };

  // Append a link back to the sheet to the description
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit#gid=0&range=${rowIndex}:${rowIndex}`;
  descriptionAdf.content.push({
    type: "rule"
  });
  descriptionAdf.content.push({
    type: "paragraph",
    content: [{
      type: "text",
      text: "Source Google Sheet Row: "
    }, {
      type: "text",
      text: "Link",
      marks: [{
        type: "link",
        attrs: {
          href: sheetUrl
        }
      }]
    }]
  });

  // --- Construct the final payload ---
  const payload = {
    fields: {
      project: {
        key: PROJECT_KEY
      },
      summary: summary,
      issuetype: {
        name: ISSUE_TYPE
      },
      assignee: {
        id: ASSIGNEE_ID
      },
      reporter: {
        id: REPORTER_ID
      },
      priority: {
        name: priority
      },
      labels: labels,
      description: descriptionAdf,
    }
  };

  if (dueDateStr) payload.fields.duedate = dueDateStr;

  return payload;
}


/**
 * Creates a Jira issue.
 * @param {Object} payload - The issue payload.
 * @returns {Object} The parsed JSON response from Jira.
 */
function createJiraIssue(payload) {
  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`)
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(`${JIRA_DOMAIN}/rest/api/3/issue`, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode !== 201) throw new Error(`Failed to create Jira issue (Status ${responseCode}). Response: ${responseBody}`);

  return JSON.parse(responseBody);
}


/**
 * Updates the sheet with the Jira ticket link.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @param {number} rowIndex - The row to update.
 * @param {number} colCount - The total number of columns.
 * @param {string[]} headers - The sheet headers.
 * @param {Object} jiraResponse - The Jira API response.
 */
function updateSheetWithJiraLink(sheet, rowIndex, colCount, headers, jiraResponse) {
  const jiraTicketUrl = `${JIRA_DOMAIN}/browse/${jiraResponse.key}`;
  const link = SpreadsheetApp.newRichTextValue()
    .setText(jiraResponse.key)
    .setLinkUrl(jiraTicketUrl)
    .build();

  let jiraKeyColIndex = headers.indexOf("Jira Key");

  if (jiraKeyColIndex === -1) {
    jiraKeyColIndex = headers.length;
    sheet.getRange(TABLE_START_ROW, TABLE_START_COLUMN + jiraKeyColIndex).setValue("Jira Key");
    headers.push("Jira Key");
  }
  sheet.getRange(rowIndex, TABLE_START_COLUMN + jiraKeyColIndex).setRichTextValue(link);
}


/**
 * Updates the Script Errors column for a specific row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @param {number} rowIndex - The row to update.
 * @param {string[]} headers - The array of sheet headers.
 * @param {string} errorMessage - The error message to write.
 */
function updateErrorColumn(sheet, rowIndex, headers, errorMessage) {
  let errorColIndex = headers.indexOf("Script Errors");

  if (errorColIndex === -1) {
    errorColIndex = headers.length;
    sheet.getRange(TABLE_START_ROW, TABLE_START_COLUMN + errorColIndex).setValue("Script Errors");
    headers.push("Script Errors");
  }
  sheet.getRange(rowIndex, TABLE_START_COLUMN + errorColIndex).setValue(errorMessage);
}


/**
 * Clears the Script Errors column for a specific row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @param {number} rowIndex - The row to clear.
 * @param {string[]} headers - The array of sheet headers.
 */
function clearErrorColumn(sheet, rowIndex, headers) {
  const errorColIndex = headers.indexOf("Script Errors");
  if (errorColIndex !== -1) sheet.getRange(rowIndex, TABLE_START_COLUMN + errorColIndex).clearContent();
}

/**
 * Gets all data from the Google Sheet.
 * @returns {SheetData|null}
 */
function getSheetData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    Logger.log(`Sheet "${SHEET_NAME}" not found.`);
    return null;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < TABLE_START_ROW || lastCol < TABLE_START_COLUMN) {
    Logger.log("No data found in the specified table range.");
    return null;
  }

  const dataRange = sheet.getRange(TABLE_START_ROW, TABLE_START_COLUMN, lastRow - TABLE_START_ROW + 1, lastCol - TABLE_START_COLUMN + 1);
  const allData = dataRange.getValues();
  const headers = allData.shift();

  return {
    sheet,
    headers,
    allRowsData: allData,
    lastRowIndex: lastRow,
    lastColIndex: lastCol,
    tableStartRow: TABLE_START_ROW,
    tableStartColumn: TABLE_START_COLUMN
  };
}


/**
 * Filters rows to get only those without a Jira ticket.
 * @param {any[][]} allRowsData - All rows from the sheet.
 * @param {string[]} headers - The sheet headers.
 * @returns {any[][]}
 */
function getRowsWithoutJiraTickets(allRowsData, headers) {
  const jiraKeyColIndex = headers.indexOf("Jira Key");
  return allRowsData.map((row, index) => ({
      rowData: row,
      originalIndex: index
    }))
    .filter(item => jiraKeyColIndex === -1 || !item.rowData[jiraKeyColIndex])
    .map(item => item.rowData);
}
