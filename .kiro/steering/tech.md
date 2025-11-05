# Gmail Invoice Scanner - Technology Stack

## Platform & Runtime
- **Google Apps Script**: JavaScript runtime environment in Google Cloud
- **Runtime Version**: V8 (modern JavaScript support)
- **Time Zone**: Asia/Jerusalem (configured in appsscript.json)

## Google APIs & Services
- **Gmail API v1**: Email scanning and message retrieval
- **Google Drive API v2**: File upload, folder management, sharing
- **Google Sheets API**: Data storage and manipulation
- **Google Apps Script Services**: UI dialogs, properties, utilities

## Required OAuth Scopes
```javascript
"oauthScopes": [
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/gmail.readonly", 
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/drive.appdata",
  "https://www.googleapis.com/auth/script.container.ui",
  "https://www.googleapis.com/auth/script.scriptapp",
  "https://www.googleapis.com/auth/script.external_request",
  "https://www.googleapis.com/auth/userinfo.email"
]
```

## Development Tools
- **Google Apps Script Editor**: Web-based IDE
- **clasp**: Command-line tool for local development (optional)
- **Script ID**: `1tA_el37r9JKkgt9nu04KVHlXrha9Nb4ys5M3oksxCbleUhn5K5D0khcE`

## Key Libraries & Patterns
- **Native GAS Services**: GmailApp, DriveApp, SpreadsheetApp, HtmlService
- **Error Handling**: Try-catch blocks with detailed logging
- **Async Operations**: Built-in GAS async handling
- **Data Persistence**: Google Sheets + PropertiesService for settings
- **UI Framework**: Custom HTML dialogs with inline CSS/JavaScript

## Performance Considerations
- **Execution Time Limits**: 6 minutes for scripts, 30 seconds for triggers
- **API Quotas**: Gmail API (1 billion quota units/day), Drive API (20,000 requests/100 seconds)
- **Memory Management**: Batch operations, periodic SpreadsheetApp.flush()
- **Rate Limiting**: Built-in exponential backoff for API calls

## Common Commands & Operations

### Deployment
```bash
# Using clasp (if configured)
clasp push    # Upload local files to Apps Script
clasp deploy  # Create new deployment
```

### Testing & Debugging
- Use Apps Script Editor's built-in debugger
- Console.log() outputs to Execution Transcript
- Custom logging to dedicated Log sheet with _logMessage()

### Key Functions to Test
```javascript
// Setup and initialization
setupSheets()           // Create Settings and Log sheets
onOpen()               // Test menu creation

// Core scanning functionality  
processInvoices()      // Main Gmail scanning function
quickApiScanInvoices() // Fast scan variant

// Bulk operations
runBulkActions()       // Process selected invoices
syncInvoicesToDrive()  // Sync to Google Drive
```

## Error Handling Patterns
- Comprehensive try-catch blocks in all main functions
- Detailed error logging with timestamps and context
- Graceful degradation (continue processing other items if one fails)
- User-friendly error messages via SpreadsheetApp.getUi().alert()

## Security & Permissions
- Read-only Gmail access (no email modification)
- Drive file creation and sharing permissions
- Spreadsheet access limited to current spreadsheet
- No external API calls except Google services