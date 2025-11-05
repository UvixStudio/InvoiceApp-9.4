# Gmail Invoice Scanner - Project Structure

## File Organization

### Core Application Files
```
├── code.gs              # Main application logic, menu creation, sheet setup
├── gmailProcessor.gs    # Gmail API scanning and email processing
├── DriveSyncer.gs       # Google Drive integration and file management
├── runActions.gs        # Bulk operations on selected invoices
├── Helpers.gs           # Utility functions and common operations
└── utils.gs             # Additional utility functions
```

### Configuration Files
```
├── appsscript.json      # Apps Script manifest (APIs, scopes, timezone)
├── .clasp.json          # clasp configuration for local development
└── .claspignore         # Files to ignore during clasp operations
```

### UI Components
```
├── ExportFolderDialog.html  # HTML dialog for folder selection
└── [Inline HTML in .gs files] # Action dialogs embedded in JavaScript
```

### Archive & Legacy
```
└── archive/             # Deprecated or backup files
    ├── DateUtils.gs     # Legacy date utilities
    ├── EmailUtils.gs    # Legacy email processing
    ├── gmailApiAdvanced.gs # Previous Gmail implementation
    └── [other legacy files]
```

## Code Architecture Patterns

### Main Entry Points
- **onOpen()**: Creates menus when spreadsheet opens
- **processInvoices()**: Primary Gmail scanning function
- **runBulkActions()**: Handles selected invoice operations
- **syncInvoicesToDrive()**: Drive synchronization

### Data Flow
1. **Settings Sheet**: Configuration (dates, keywords, exclusions)
2. **Gmail Scan**: Email retrieval and filtering
3. **Data Processing**: Extraction and categorization
4. **Results Sheet**: Structured invoice data
5. **Drive Upload**: PDF file management
6. **Log Sheet**: Operation tracking and debugging

### Sheet Structure
```
Settings Sheet:
- A: Start Date, B: End Date
- C: EN Keywords, D: HE Keywords  
- E: Excluded Keywords, F: Exclude Emails
- G: Approved Senders, I: Export Folder ID

Results Sheets (invoices YYYY-MM-DD):
- A: Action (checkbox), B: Category
- C: Sender Name, D: Sender Email, E: Date
- F: Subject, G: PDF Link, H: Email ID
- I: Attachment Name, J-L: Amount/Currency data

Log Sheet:
- A: Timestamp, B: Message (with color coding by type)
```

## Naming Conventions

### Functions
- **camelCase**: `processInvoices()`, `extractSenderName()`
- **Prefixed helpers**: `_logMessage()`, `_parseSender()` (internal use)
- **Action handlers**: `handleExcludeAction()`, `handleApproveAction()`

### Variables
- **camelCase**: `startDate`, `emailData`, `pdfAttachment`
- **CONSTANTS**: `SETTINGS_SHEET_NAME`, `LOG_SHEET_NAME`
- **Arrays**: `checkedRows`, `excludedEmails`, `allKeywords`

### Files & Sheets
- **PascalCase**: `DriveSyncer.gs`, `Helpers.gs`
- **Sheet names**: "Settings", "Log", "invoices YYYY-MM-DD"
- **Folder structure**: Organized by date ranges and categories

## Development Workflow

### Adding New Features
1. **Core Logic**: Add main function to appropriate .gs file
2. **Menu Integration**: Update `createMainMenu()` or `createActionsMenu()`
3. **Error Handling**: Add try-catch with `_logMessage()` logging
4. **Testing**: Test with small data sets first
5. **Documentation**: Update relevant steering docs

### Debugging Approach
1. **Console Logs**: Use `console.log()` for development
2. **Sheet Logging**: Use `_logMessage(log, message, type)` for production
3. **Execution Transcript**: Check Apps Script editor for runtime errors
4. **Step-by-step**: Test individual functions before integration

### Code Organization Rules
- **Single Responsibility**: Each .gs file has a clear purpose
- **Helper Functions**: Common utilities in Helpers.gs and utils.gs
- **Error Boundaries**: Wrap main operations in try-catch blocks
- **State Management**: Use Google Sheets and PropertiesService for persistence
- **UI Separation**: Keep HTML dialogs separate from business logic

## Integration Points

### Google Services
- **SpreadsheetApp**: Data storage and UI
- **GmailApp**: Email scanning and retrieval  
- **DriveApp**: File upload and organization
- **HtmlService**: Custom dialog interfaces
- **PropertiesService**: User preferences and settings

### External Dependencies
- **None**: Pure Google Apps Script implementation
- **APIs**: Only Google's native APIs (Gmail, Drive, Sheets)
- **Libraries**: No external JavaScript libraries required