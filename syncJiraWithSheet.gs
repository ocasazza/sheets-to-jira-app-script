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

WHAT HAPPENS WHEN YOU RUN IT:
- The script looks at every row in your sheet
- For rows without a Jira ticket link, it creates a new Jira ticket
- It uses the column data from your sheet to fill in the Jira ticket details
- It adds a clickable link to the new Jira ticket in your sheet
- If there are any errors, they're recorded in a "Script Errors" column

IMPORTANT NOTES:
- The script will NOT create duplicate tickets (it skips rows that already have Jira links)
- You can run this script multiple times safely
- Any errors are logged so you can fix issues and re-run
*/

// =============================================================================
//                           JIRA CONNECTION SETTINGS
// =============================================================================
/*
These settings tell the script how to connect to your Jira instance.
You'll need to get these values from your Jira administrator.
*/

// Your company's Jira web address (the URL you use to access Jira in your browser)
// Example: "https://yourcompany.atlassian.net" or "https://jira.yourcompany.com"
const JIRA_DOMAIN = "schrodinger.atlassian.net";

// The email address associated with your Jira account
const JIRA_EMAIL = "";

// Your personal Jira API token (this is like a password - keep it secure!)
// To get this: Go to Jira → Your Profile → Personal Access Tokens → Create Token
const JIRA_API_TOKEN = "";

// The Jira project where tickets should be created (usually a short code like "PROJ")
// You can find this in Jira by looking at existing ticket numbers (e.g., if tickets are "PROJ", then use "")
const PROJECT_KEY = "";

// What type of Jira tickets to create (common options: "Task", "Bug", "Story", "Epic")
// Check with your Jira admin for what types are available in your project
const ISSUE_TYPE = "Task";

// Who should be listed as the reporter and assignee for new tickets
// These are Jira user IDs (not email addresses) - ask your Jira admin for these
const REPORTER_ID = "";

// =============================================================================
//                        GOOGLE SHEET COLUMN MAPPING
// =============================================================================
/*
This section maps Jira ticket fields to specific columns in your Google Sheet.
Each Jira field specifies which column to use for that data.

HOW IT WORKS:
- Each Jira field (like 'summary', 'description') maps to ONE specific column
- You specify which column name to use for each Jira field
- Required fields must have data (or will use default values)
- Optional fields are included if the column has data

To customize this for your sheet:
1. Look at the column headers in row 1 of your Google Sheet
2. Update the 'columnName' values below to match your sheet exactly (case-sensitive!)
3. Add new Jira fields as needed
4. Set required: true for Jira fields that must have data

EXAMPLE:
If your sheet has a column called "Product Name" and you want that to be the
Jira ticket title, change the 'summary' mapping to:
'summary': { columnName: 'Product Name', required: true, ... }
*/

const JIRA_FIELD_TO_COLUMN_MAPPING = {
  // REQUIRED JIRA FIELDS - These must be mapped to columns in your sheet
  'summary': {
    columnName: 'Title',
    required: true,
    defaultValue: 'Apps Script to Jira Automation',
    description: 'The title/summary of the Jira ticket'
  },
  'timestamp': {
    columnName: 'Timestamp',
    required: true,
    defaultValue: null,
    description: 'When the request was submitted (will be added to ticket description)'
  },
  'reporter': {
    columnName: 'Email Address',
    required: true,
    defaultValue: 'Google Sheets Apps Script Automation',
    description: 'Who made the request (will be added to ticket description)'
  },

  // OPTIONAL JIRA FIELDS - These are nice to have but not required
  'description': {
    columnName: 'Description',
    required: false,
    description: 'This ticket was created by the Google Sheets Apps Script automation'
  },
  'duedate': {
    columnName: 'Due Date',
    required: false,
    description: 'When the work should be completed'
  },
  'priority': {
    columnName: 'Priority',
    required: false,
    description: 'How urgent this is (e.g., "High", "Medium", "Low")'
  }
};

/*
ADDING MORE COLUMNS:
To add more columns from your sheet, add them to the mapping above like this:
'Your Column Name': { jiraField: 'description', required: false },

COLUMN NAMES MUST MATCH EXACTLY:
- If your sheet has "Product Name", use 'Product Name' (not 'product name' or 'Product_Name')
- Spaces, capitalization, and punctuation must match exactly

WHAT HAPPENS TO OTHER COLUMNS:
Any columns in your sheet that aren't listed above (except "Jira Key" and "Script Errors")
will automatically be included in the ticket description as additional information.
*/

// Columns that should be completely ignored (won't appear in Jira tickets at all)
// You usually don't need to change this unless you have other columns you want to exclude
const IGNORED_COLUMNS = ['Jira Key', 'Script Errors'];

// =============================================================================
//                           GOOGLE SHEET SETTINGS
// =============================================================================
/*
These settings tell the script which Google Sheet to work with.
*/

// The ID of your Google Sheet (found in the URL when you have the sheet open)
// Example: If your sheet URL is https://docs.google.com/spreadsheets/d/1ABC123.../edit
// Then your ID is the part between "/d/" and "/edit" → "1ABC123..."
const SPREADSHEET_ID = '';

// The name of the specific sheet tab to process (shown at the bottom of Google Sheets)
// This is usually something like "Sheet1", "Form Responses 1", or a custom name you've given it
const SHEET_NAME = 'Form Responses';

// =============================================================================
//                           TABLE POSITION SETTINGS
// =============================================================================
/*
These settings tell the script where your data table is located within the sheet.
This allows you to have other content above or to the left of your data table.

IMPORTANT: Your data table must have:
- Headers in the first row of the table
- Data rows immediately below the headers
- No empty rows between headers and data
*/

// The top-left corner of your data table (where the first header cell is located)
// Row and column numbers start at 1 (not 0)
// Examples:
//   - If your table starts at A1: TABLE_START_ROW = 1, TABLE_START_COLUMN = 1
//   - If your table starts at C5: TABLE_START_ROW = 5, TABLE_START_COLUMN = 3
//   - If your table starts at B2: TABLE_START_ROW = 2, TABLE_START_COLUMN = 2
const TABLE_START_ROW = 2;    // Row number where your headers are located
const TABLE_START_COLUMN = 1; // Column number where your first header is located

/*
                              TROUBLESHOOTING

COMMON ISSUES AND SOLUTIONS:

1. "Column not found" errors:
   → Check that your column names in COLUMN_TO_JIRA_MAPPING exactly match your sheet headers

2. "Authentication failed" errors:
   → Verify your JIRA_EMAIL and JIRA_API_TOKEN are correct
   → Make sure your API token hasn't expired

3. "Project not found" errors:
   → Confirm PROJECT_KEY matches an existing Jira project you have access to

4. "Sheet not found" errors:
   → Check that SPREADSHEET_ID and SHEET_NAME are correct
   → Make sure this script has permission to access your Google Sheet

5. Script runs but no tickets are created:
   → Check if all rows already have Jira links (script skips rows with existing tickets)
   → Look for error messages in the "Script Errors" column of your sheet

6. Some rows are skipped:
   → Check the "Script Errors" column for specific error messages
   → Verify that required fields (Timestamp, Requester, Title) have data in those rows

GETTING HELP:
- Check the "Script Errors" column in your sheet for specific error messages
- Look at the execution log in Apps Script (View → Logs) for detailed information
- Contact your Jira administrator for help with Jira-specific settings
*/


// --- PROGRAM ------------------------------------
// --- DO NOT EDIT BELOW HERE UNLESS DEVELOPING ---
/**
 * @typedef {Object} SheetData
 * @property {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @property {string[]} headers - An array of header values.
 * @property {any[]} latestRowData - An array of values from the last row.
 * @property {number} lastRowIndex - The index of the last row.
 * @property {number} lastColIndex - The index of the last column.
 */

/**
 * Main function to orchestrate the process of creating Jira tickets for all rows without existing tickets.
 */
function createJiraTicketsForAllRows() {
  Logger.log("Starting process to create Jira tickets for all rows without existing tickets...");
  try {
    const data = getAllSheetData();
    if (!data) return;

    const rowsWithoutTickets = getRowsWithoutJiraTickets(data.allRowsData, data.headers);

    if (rowsWithoutTickets.length === 0) {
      Logger.log("No rows found without Jira tickets. All rows are already processed.");
      return;
    }

    Logger.log(`Found ${rowsWithoutTickets.length} rows without Jira tickets. Processing...`);

    let successCount = 0;
    let errorCount = 0;

    rowsWithoutTickets.forEach((rowInfo, index) => {
      try {
        Logger.log(`Processing row ${rowInfo.rowIndex} (${index + 1}/${rowsWithoutTickets.length})...`);

        // Validate required fields first
        const validation = validateRequiredFields(rowInfo.rowData, data.headers);
        if (!validation.success) {
          Logger.log(`Validation failed for row ${rowInfo.rowIndex}: ${validation.error}`);
          updateErrorColumn(data.sheet, rowInfo.rowIndex, data.headers, validation.error);
          errorCount++;
          return;
        }

        // Build Jira payload
        const result = buildJiraPayload(rowInfo.rowData, data.headers, rowInfo.rowIndex);
        if (!result || !result.payload) {
          const errorMsg = "Could not build Jira payload";
          Logger.log(`Skipping row ${rowInfo.rowIndex}: ${errorMsg}`);
          updateErrorColumn(data.sheet, rowInfo.rowIndex, data.headers, errorMsg);
          errorCount++;
          return;
        }

        // Create Jira issue
        const jiraResponse = createJiraIssue(result.payload);

        // Update sheet with default values if any were used
        if (result.defaultsUsed && result.defaultsUsed.length > 0) {
          updateSheetWithDefaults(data.sheet, rowInfo.rowIndex, result.defaultsUsed);
        }

        // Update sheet with Jira link and clear any previous errors
        updateSheetWithJiraLink(data.sheet, rowInfo.rowIndex, data.lastColIndex, data.headers, jiraResponse);
        clearErrorColumn(data.sheet, rowInfo.rowIndex, data.headers);

        Logger.log(`Successfully created Jira ticket ${jiraResponse.key} for row ${rowInfo.rowIndex}`);
        successCount++;

      } catch (err) {
        const errorMsg = `Error creating Jira ticket: ${err.toString()}`;
        Logger.log(`Error processing row ${rowInfo.rowIndex}: ${errorMsg}`);
        updateErrorColumn(data.sheet, rowInfo.rowIndex, data.headers, errorMsg);
        errorCount++;
      }
    });

    Logger.log(`Process completed. Successfully created ${successCount} tickets. ${errorCount} errors occurred.`);

  } catch (err) {
    Logger.log("An unexpected error occurred during execution: " + err.toString());
  }
}

/**
 * Validates that all required Jira fields have data or default values available.
 * @param {any[]} rowData - An array of values from the sheet row.
 * @param {string[]} headers - An array of header values.
 * @returns {Object} Validation result with success flag and error message if any.
 */
function validateRequiredFields(rowData, headers) {
  const missingFields = [];

  // Check each required Jira field
  for (const [jiraField, config] of Object.entries(JIRA_FIELD_TO_COLUMN_MAPPING)) {
    if (config.required) {
      const colIndex = headers.indexOf(config.columnName);
      let hasValue = false;

      // Check if column exists and has data
      if (colIndex !== -1 && rowData[colIndex] && rowData[colIndex].toString().trim() !== '') {
        hasValue = true;
      }

      // If no value from column, check if we have a default
      if (!hasValue && (config.defaultValue !== null && config.defaultValue !== undefined)) {
        hasValue = true;
        Logger.log(`Will use default value for "${jiraField}": ${config.defaultValue}`);
      }

      // Only fail if we have neither column data nor a default value
      if (!hasValue) {
        if (colIndex === -1) {
          missingFields.push(`Required Jira field "${jiraField}" expects column "${config.columnName}" but column not found and no default value provided`);
        } else {
          missingFields.push(`Required Jira field "${jiraField}" has no data in column "${config.columnName}" and no default value provided`);
        }
      }
    }
  }

  return {
    success: missingFields.length === 0,
    error: missingFields.length > 0 ? missingFields.join('; ') : null
  };
}

/**
 * Updates the Script Errors column for a specific row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @param {number} rowIndex - The row to update (1-based).
 * @param {string[]} headers - The array of sheet headers.
 * @param {string} errorMessage - The error message to write.
 */
function updateErrorColumn(sheet, rowIndex, headers, errorMessage) {
  let errorColIndex = headers.indexOf("Script Errors");

  if (errorColIndex === -1) {
    // Create Script Errors column if it doesn't exist
    const newColIndex = sheet.getLastColumn() + 1;
    sheet.getRange(TABLE_START_ROW, newColIndex).setValue("Script Errors");
    errorColIndex = newColIndex - TABLE_START_COLUMN; // Convert to 0-based relative to table start
    headers.push("Script Errors"); // Update headers array
  }

  // Write error message to the cell (accounting for table origin)
  const actualColIndex = TABLE_START_COLUMN + errorColIndex;
  sheet.getRange(rowIndex, actualColIndex).setValue(errorMessage);
  Logger.log(`Updated Script Errors column at row ${rowIndex}, column ${actualColIndex}: ${errorMessage}`);
}

/**
 * Clears the Script Errors column for successful processing.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @param {number} rowIndex - The row to clear (1-based).
 * @param {string[]} headers - The array of sheet headers.
 */
function clearErrorColumn(sheet, rowIndex, headers) {
  const errorColIndex = headers.indexOf("Script Errors");
  if (errorColIndex !== -1) {
    // Account for table origin when calculating column position
    const actualColIndex = TABLE_START_COLUMN + errorColIndex;
    sheet.getRange(rowIndex, actualColIndex).setValue("");
  }
}

/**
 * Builds the JSON payload for the Jira API request using the configured Jira field mapping.
 * @param {any[]} rowData - An array of values from the sheet row.
 * @param {string[]} headers - An array of header values.
 * @param {number} rowIndex - The row number (1-based) for creating a link back to the sheet.
 * @returns {Object} Object containing the Jira payload and any default values used.
 */
function buildJiraPayload(rowData, headers, rowIndex) {
  const payload = {
    fields: {
      project: { key: PROJECT_KEY },
      issuetype: { name: ISSUE_TYPE },
      reporter: { id: REPORTER_ID },
      assignee: { id: ASSIGNEE_ID }
    }
  };

  const defaultsUsed = []; // Track which defaults were applied

  // Process each Jira field mapping
  for (const [jiraField, config] of Object.entries(JIRA_FIELD_TO_COLUMN_MAPPING)) {
    const colIndex = headers.indexOf(config.columnName);
    let value = null;
    let usedDefault = false;

    // Get value from column or use default
    if (colIndex !== -1 && rowData[colIndex] && rowData[colIndex].toString().trim() !== '') {
      value = rowData[colIndex].toString().trim();
    } else if (config.defaultValue !== null && config.defaultValue !== undefined) {
      value = config.defaultValue;
      usedDefault = true;
      Logger.log(`Using default value for "${jiraField}": ${value}`);

      // Track this default for updating the sheet later
      if (colIndex !== -1) {
        defaultsUsed.push({
          columnName: config.columnName,
          columnIndex: colIndex,
          defaultValue: value
        });
      }
    }

    // Apply the value to the appropriate Jira field
    if (value) {
      switch (jiraField) {
        case 'summary':
          payload.fields.summary = value;
          break;
        case 'description':
          payload.fields.description = {
            type: "doc",
            version: 1,
            content: [{ type: "paragraph", content: [{ type: "text", text: value }] }]
          };
          break;
        case 'duedate':
          const dueDate = new Date(value);
          if (!isNaN(dueDate.getTime())) {
            payload.fields.duedate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          break;
        case 'priority':
          payload.fields.priority = { name: value };
          break;
        case 'timestamp':
          // Add timestamp to description since Jira created field is auto-set
          if (!payload.fields.description) {
            payload.fields.description = {
              type: "doc",
              version: 1,
              content: [{ type: "paragraph", content: [{ type: "text", text: `Timestamp: ${value}\n` }] }]
            };
          } else {
            payload.fields.description.content[0].content[0].text = `Timestamp: ${value}\n` + payload.fields.description.content[0].content[0].text;
          }
          break;
        case 'reporter':
          // Note: This would require looking up the Jira user by email
          // For now, we'll add it to the description
          if (!payload.fields.description) {
            payload.fields.description = {
              type: "doc",
              version: 1,
              content: [{ type: "paragraph", content: [{ type: "text", text: `Requester: ${value}\n` }] }]
            };
          } else {
            payload.fields.description.content[0].content[0].text = `Requester: ${value}\n` + payload.fields.description.content[0].content[0].text;
          }
          break;
      }
    }
  }

  // Get list of mapped column names to exclude from additional info
  const mappedColumns = Object.values(JIRA_FIELD_TO_COLUMN_MAPPING).map(config => config.columnName);

  // Add additional columns to description (excluding ignored columns and mapped columns)
  let additionalInfo = "";
  headers.forEach((header, i) => {
    if (!IGNORED_COLUMNS.includes(header) &&
        !mappedColumns.includes(header) &&
        rowData[i]) {
      additionalInfo += `${header}: ${rowData[i]}\n`;
    }
  });

  if (additionalInfo) {
    if (!payload.fields.description) {
      payload.fields.description = {
        type: "doc",
        version: 1,
        content: [{ type: "paragraph", content: [{ type: "text", text: additionalInfo }] }]
      };
    } else {
      payload.fields.description.content[0].content[0].text += `\n${additionalInfo}`;
    }
  }

  // Add link back to the Google Sheet row at the end of the description
  // Create a range that highlights the entire data row within the table bounds
  const startCol = String.fromCharCode(64 + TABLE_START_COLUMN); // Convert column number to letter (A, B, C, etc.)
  const endCol = String.fromCharCode(64 + TABLE_START_COLUMN + headers.length - 1);
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit#gid=0&range=${startCol}${rowIndex}:${endCol}${rowIndex}`;

  // Create separator text
  const separatorText = `\n\n---\nSource: ${SHEET_NAME} Row ${rowIndex}`;

  // Create the inline card for the Google Sheets link
  const sheetLinkCard = {
    type: "paragraph",
    content: [
      {
        type: "inlineCard",
        attrs: {
          url: sheetUrl
        }
      }
    ]
  };

  if (!payload.fields.description) {
    payload.fields.description = {
      type: "doc",
      version: 1,
      content: [
        { type: "paragraph", content: [{ type: "text", text: separatorText }] },
        sheetLinkCard
      ]
    };
  } else {
    // Add separator text to existing description
    payload.fields.description.content[0].content[0].text += separatorText;
    // Add the inline card as a new paragraph
    payload.fields.description.content.push(sheetLinkCard);
  }

  return {
    payload: payload,
    defaultsUsed: defaultsUsed
  };
}

/**
 * Sends the request to Jira to create a new issue.
 * @param {Object} payload - The issue payload.
 * @returns {Object} The parsed JSON response from Jira.
 * @throws {Error} If the API call fails.
 */
function createJiraIssue(payload) {
  const jsonPayload = JSON.stringify(payload, null, 2);

  const options = {
    method: "POST",
    contentType: "application/json",
    payload: jsonPayload,
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`)
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(`${JIRA_DOMAIN}/rest/api/3/issue`, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();
  if (responseCode !== 201) {
    Logger.log(`Error creating Jira issue. Status: ${responseCode}. Response: ${responseBody}`);
    throw new Error(`Failed to create Jira issue (Status ${responseCode}). Response: ${responseBody}`);
  }
  return JSON.parse(responseBody);
}

/**
 * Updates the Google Sheet with default values that were used during Jira ticket creation.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object to update.
 * @param {number} rowIndex - The row to update (1-based).
 * @param {Object[]} defaultsUsed - Array of default values that were applied.
 */
function updateSheetWithDefaults(sheet, rowIndex, defaultsUsed) {
  defaultsUsed.forEach(defaultInfo => {
    try {
      // Account for table origin when calculating column position
      const actualColIndex = TABLE_START_COLUMN + defaultInfo.columnIndex;
      sheet.getRange(rowIndex, actualColIndex).setValue(defaultInfo.defaultValue);
      Logger.log(`Updated row ${rowIndex}, column "${defaultInfo.columnName}" with default value: ${defaultInfo.defaultValue}`);
    } catch (err) {
      Logger.log(`Error updating default value for column "${defaultInfo.columnName}": ${err.toString()}`);
    }
  });
}

/**
 * Writes a hyperlink to the newly created Jira ticket back to the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object to update.
 * @param {number} rowIndex - The row to write to (1-based).
 * @param {number} colCount - The current number of columns in the sheet.
 * @param {string[]} headers - The array of sheet headers.
 * @param {Object} jiraResponse - The parsed response from the Jira API.
 */
function updateSheetWithJiraLink(sheet, rowIndex, colCount, headers, jiraResponse) {
    const jiraTicketUrl = `${JIRA_DOMAIN}/browse/${jiraResponse.key}`;
    const link = SpreadsheetApp.newRichTextValue()
      .setText(jiraResponse.key)
      .setLinkUrl(jiraTicketUrl)
      .build();

    let jiraKeyColIndex = headers.indexOf("Jira Key");

    if (jiraKeyColIndex !== -1) {
      // Account for table origin when calculating column position
      const actualColIndex = TABLE_START_COLUMN + jiraKeyColIndex;
      sheet.getRange(rowIndex, actualColIndex).setRichTextValue(link);
      Logger.log(`Updated existing Jira Key column at row ${rowIndex}, column ${actualColIndex}`);
    } else {
      // Create new "Jira Key" column
      const newColIndex = colCount + 1;
      sheet.getRange(TABLE_START_ROW, newColIndex).setValue("Jira Key"); // Header row
      sheet.getRange(rowIndex, newColIndex).setRichTextValue(link); // Data row
      Logger.log(`Created new Jira Key column at column ${newColIndex}, updated row ${rowIndex}`);
    }
    Logger.log(`Updated sheet with Jira key: ${jiraResponse.key}`);
}


/**
 * Fetches all rows of data from the specified Google Sheet using the configured table origin.
 * @returns {Object | null} An object containing sheet data and all rows, or null if no data is found.
 */
function getAllSheetData() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log(`Error: Sheet with name "${SHEET_NAME}" was not found.`);
    return null;
  }

  const lastRowIndex = sheet.getLastRow();
  const lastColIndex = sheet.getLastColumn();

  // Check if we have enough rows for our table starting position
  if (lastRowIndex < TABLE_START_ROW) {
    Logger.log(`No data found. Sheet only has ${lastRowIndex} rows, but table starts at row ${TABLE_START_ROW}.`);
    return null;
  }

  // Check if we have enough columns for our table starting position
  if (lastColIndex < TABLE_START_COLUMN) {
    Logger.log(`No data found. Sheet only has ${lastColIndex} columns, but table starts at column ${TABLE_START_COLUMN}.`);
    return null;
  }

  // Calculate the actual data range based on table origin
  const dataRowCount = lastRowIndex - TABLE_START_ROW;
  const dataColCount = lastColIndex - TABLE_START_COLUMN + 1;

  if (dataRowCount <= 0) {
    Logger.log("No data rows found to process after accounting for table position.");
    return null;
  }

  // Get headers from the table start position
  const headers = sheet.getRange(TABLE_START_ROW, TABLE_START_COLUMN, 1, dataColCount).getValues()[0];

  // Get all data rows (starting from the row after headers)
  const allRowsData = sheet.getRange(TABLE_START_ROW + 1, TABLE_START_COLUMN, dataRowCount, dataColCount).getValues();

  Logger.log(`Found table at row ${TABLE_START_ROW}, column ${TABLE_START_COLUMN} with ${dataRowCount} data rows and ${dataColCount} columns`);

  return {
    sheet,
    headers,
    allRowsData,
    lastRowIndex,
    lastColIndex,
    tableStartRow: TABLE_START_ROW,
    tableStartColumn: TABLE_START_COLUMN
  };
}

/**
 * Filters rows to find those without Jira tickets.
 * @param {any[][]} allRowsData - All rows of data from the sheet.
 * @param {string[]} headers - An array of header values.
 * @returns {Object[]} An array of objects containing row data and row indices for rows without Jira tickets.
 */
function getRowsWithoutJiraTickets(allRowsData, headers) {
  const jiraKeyColIndex = headers.indexOf("Jira Key");
  const rowsWithoutTickets = [];

  allRowsData.forEach((rowData, index) => {
    // Calculate actual row index based on table origin
    // index is 0-based within the data array, so we add TABLE_START_ROW + 1 to get the actual sheet row
    const actualRowIndex = index + TABLE_START_ROW + 1;
    const hasJiraKey = jiraKeyColIndex !== -1 && rowData[jiraKeyColIndex] && rowData[jiraKeyColIndex].toString().trim() !== '';

    if (!hasJiraKey) {
      rowsWithoutTickets.push({
        rowData: rowData,
        rowIndex: actualRowIndex
      });
    }
  });

  return rowsWithoutTickets;
}
