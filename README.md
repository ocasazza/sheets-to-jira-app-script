# Google Sheets to Jira Automation

This automation creates Jira tickets automatically from Google Sheets data using Google Apps Script. Perfect for processing form responses, support requests, or any structured data that needs to become Jira tickets.

## Quick Start

### Step 1: Copy the Template Resources

1. **Access the template folder**: [Google Drive Template Folder](https://drive.google.com/drive/folders/1CI44hdYy9GAE_a39LyHNLLHrBzvqG8nT?usp=drive_link)

2. **Copy the Google Form** (optional):
   - Right-click on the Google Form → "Make a copy"
   - Rename it to match your use case
   - Customize the form questions as needed

3. **Copy the Google Sheet**:
   - Right-click on the Google Sheet → "Make a copy"
   - Rename it appropriately
   - Note the new Sheet ID from the URL (you'll need this later)

### Step 2: Set Up Google Apps Script

1. **Open your copied Google Sheet**
2. **Go to Extensions → Apps Script**
3. **Delete the default `Code.gs` file**
4. **Create a new file called `syncJiraWithSheet.gs`**
5. **Copy the entire script** from this repository's `syncJiraWithSheet.gs` file
6. **Paste it into your new Apps Script file**
7. **Save the project** (Ctrl+S or Cmd+S)

### Step 3: Configure Your Settings

Edit the configuration section at the top of the script:

#### 🔗 Jira Connection Settings

```javascript
// Your company's Jira web address
const JIRA_DOMAIN = "https://yourcompany.atlassian.net";

// Your Jira account email
const JIRA_EMAIL = "your.email@company.com";

// Your Jira API token (see instructions below)
const JIRA_API_TOKEN = "your_api_token_here";

// Your Jira project key (e.g., "PROJ", "ITOPS")
const PROJECT_KEY = "YOUR_PROJECT_KEY";

// Type of tickets to create
const ISSUE_TYPE = "Task";

// Your Jira user ID (see instructions below)
const REPORTER_ID = "your_jira_user_id";
```

#### 📊 Sheet Configuration

```javascript
// Your Google Sheet ID (from the URL)
const SPREADSHEET_ID = 'your_sheet_id_here';

// Name of the sheet tab
const SHEET_NAME = 'Form Responses';

// Where your data table starts (usually row 1, column 1)
const TABLE_START_ROW = 1;
const TABLE_START_COLUMN = 1;
```

#### 🗂️ Column Mapping

Configure which columns map to which Jira fields:

```javascript
const JIRA_FIELD_TO_COLUMN_MAPPING = {
  'summary': {
    columnName: 'Title',  // Change to match your column name
    required: true,
    defaultValue: 'New Request',
    description: 'The title/summary of the Jira ticket'
  },
  'timestamp': {
    columnName: 'Timestamp',  // Change to match your column name
    required: true,
    defaultValue: null,
    description: 'When the request was submitted'
  },
  'reporter': {
    columnName: 'Email Address',  // Change to match your column name
    required: true,
    defaultValue: 'Unknown requester',
    description: 'Who made the request'
  },
  // Add more mappings as needed...
};
```

### Step 4: Get Your Jira API Token

1. **Log into your Jira instance**
2. **Go to your Profile → Personal Access Tokens**
3. **Create a new token**:
   - Give it a descriptive name (e.g., "Google Sheets Automation")
   - Set appropriate permissions
   - Copy the token immediately (you won't see it again!)
4. **Paste the token** into the `JIRA_API_TOKEN` field in your script

### Step 5: Find Your Jira User ID

1. **Go to your Jira profile page**
2. **Look at the URL** - it will contain your user ID
3. **Or use the Jira API**:
   - Go to: `https://yourcompany.atlassian.net/rest/api/3/myself`
   - Look for the `accountId` field
4. **Paste the ID** into the `REPORTER_ID` field

### Step 6: Configure Your Sheet Columns

1. **Look at your Google Sheet headers** (row 1)
2. **Update the column mappings** in the script to match your exact column names
3. **Make sure column names match exactly** (case-sensitive!)

Example:
- If your sheet has "Product Name" → use `columnName: 'Product Name'`
- If your sheet has "Email" → use `columnName: 'Email'`

### Step 7: Test the Setup

1. **Save your Apps Script project**
2. **Run the function `createJiraTicketsForAllRows`**:
   - Click the play button next to the function name
   - Grant necessary permissions when prompted
3. **Check the execution log** for any errors
4. **Verify tickets were created** in your Jira project
5. **Check your Google Sheet** for Jira links in the new "Jira Key" column

## 🔧 Advanced Configuration

### Table Positioning

If your data doesn't start at A1, configure the table origin:

```javascript
const TABLE_START_ROW = 2;    // If headers are in row 2
const TABLE_START_COLUMN = 3; // If data starts in column C
```

### Custom Field Mappings

Add more Jira fields by extending the mapping:

```javascript
'priority': {
  columnName: 'Priority Level',
  required: false,
  description: 'Ticket priority (High, Medium, Low)'
},
'duedate': {
  columnName: 'Due Date',
  required: false,
  description: 'When work should be completed'
}
```

### Ignored Columns

Specify columns to exclude from tickets:

```javascript
const IGNORED_COLUMNS = ['Jira Key', 'Script Errors', 'Internal Notes'];
```

## Security Notes

- **Keep your API token secure** - don't share it or commit it to version control
- **Use minimal permissions** - only grant what the automation needs
- **Regularly rotate tokens** - create new tokens periodically
- **Monitor usage** - check Jira audit logs for automation activity

## Contributing

Found a bug or want to add a feature? Contributions are welcome!

1. Fork the repository
2. Make your changes
3. Test thoroughly
4. Submit a pull request

## License

This project is open source and available under the MIT License.
