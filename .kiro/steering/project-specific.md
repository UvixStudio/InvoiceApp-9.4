# Development Guidelines - Invoice Scanner

## 🚨 Critical Rules - Read Before ANY Code Change

### Core Principles
- **Slow and careful is better than fast and broken**
- **One fix at a time**
- **Test after every change**
- **Document everything**

## Pre-Change Checklist

### 1. Preparation
- [ ] Understand the current code and architecture
- [ ] Have a backup of current working code
- [ ] Know exactly what you're changing and why
- [ ] Identify which files need modification
- [ ] Identify which functions might be affected

### 2. Planning
- [ ] Change is minimal and focused (one task only)
- [ ] Have a clear testing plan
- [ ] Checked for duplicates or conflicts
- [ ] Won't break architecture rules

## Architecture Rules (MUST FOLLOW)

### ✅ DO:
- **Create checkboxes ONLY after writing data** (never before)
- **Use `findLastRowWithData()`** to find the real last row (not `getLastRow()`)
- **Use try-catch** on all external operations (Gmail, Drive, Sheets)
- **Use unified logging** with `_logMessage(log, message, type)`
- **Write to specific rows** using `getRange(row, col).setValues()` instead of `appendRow()`

### ❌ DON'T:
- **Never create 1000 checkboxes upfront** (causes row 1000 bug)
- **Never use hardcoded column references** - use `getExportFolderId()` and `saveFolderId()` functions
- **Never create duplicate functions** with same name
- **Never delete old code before new code works**
- **Never change multiple things at once**

## Code Quality Standards

### Naming Conventions
- **Functions**: camelCase (`processInvoices`, `extractSenderName`)
- **Internal helpers**: Prefixed with underscore (`_logMessage`, `_parseSender`)
- **Constants**: UPPER_SNAKE_CASE (`SETTINGS_SHEET_NAME`, `LOG_SHEET_NAME`)
- **Variables**: camelCase (`startDate`, `emailData`, `pdfAttachment`)

### Code Structure
- Clear, descriptive names
- Comments for complex sections
- No dead or commented-out code
- Follow existing patterns

## Testing Requirements

### Basic Tests (MANDATORY)
- [ ] Code runs without JavaScript errors
- [ ] Menus display correctly (in Hebrew, no gibberish)
- [ ] New function works as expected
- [ ] Old functions still work

### Functional Tests
- [ ] Scan works (creates sheet, writes from row 2)
- [ ] Actions work (select all, exclude, export)
- [ ] Logs write correctly (to Log sheet)
- [ ] Error messages are clear and in Hebrew

### Full Flow Test (when relevant)
- [ ] Scan → Mark → Exclude → Export
- [ ] Error handling: no internet/permissions
- [ ] Performance: scan is fast enough

## Post-Change Requirements

### Documentation
- [ ] Update `Tasks_Management.txt` with task status
- [ ] Write in `Release_notes` what changed
- [ ] Update this file with test results

### Deployment
- [ ] Test locally first
- [ ] `clasp push` only after verification
- [ ] `git push` with clear commit message
- [ ] Document in Release Notes with date and deploy info

## Emergency Procedures

### If Something Breaks:
1. **STOP** - Don't continue changing
2. **READ** - Re-read the code and error
3. **BACKUP** - Ensure backup exists
4. **ROLLBACK** - Return to working version
5. **ANALYZE** - Understand what went wrong
6. **PLAN** - Re-plan before next attempt

### If Unsure:
1. **ASK** - Ask the user
2. **EXPLAIN** - Explain what you want to do
3. **WAIT** - Wait for approval before changing
4. **DOCUMENT** - Document your thinking

## Known Issues to Avoid

### 🔴 Critical Bugs (Fixed - Don't Reintroduce!)
1. **Row 1000 Bug**: Never use `insertCheckboxes("A2:A1000")` before writing data
2. **Folder ID Mismatch**: There's inconsistency - `Helpers.gs` uses H2, `DriveSyncer.gs` uses I2. Always use `getExportFolderId()` and `saveFolderId()` functions instead of direct cell access
3. **Template String Errors**: Always use proper quotes in template strings

### 🟡 Common Pitfalls
4. **Function Duplicates**: Check for existing functions before creating new ones (e.g., `_logMessage` exists in multiple files, `saveFolderId` in both Helpers.gs and code.gs)
5. **Broken Code**: Remove incomplete/broken function definitions
6. **Missing Menu Items**: Ensure all functions are accessible from menus
7. **Archive Files**: Files in `archive/` folder should not be used - they're old versions (e.g., `gmailApiAdvanced.gs` is archived, use `gmailProcessor.gs` instead)

## File-Specific Guidelines

### gmailProcessor.gs ⭐ (PRIMARY SCAN ENGINE)
- **Main scanning engine** - use this for all Gmail scanning
- Uses `findLastRowWithData()` for row detection
- Creates checkboxes AFTER data writing
- Includes `cleanupOldCheckboxes()` for cleanup
- Function: `processInvoices()` - the main entry point

### code.gs (LEGACY + MENUS)
- Menu creation: `createMainMenu()` and `createActionsMenu()`
- Setup functions: `setupSheets()`, `resetDateRange()`
- Contains legacy code and some deprecated functions
- Contains "FIXED VERSION BY CLAUDE" marker
- Has duplicate functions (e.g., `saveFolderId`) - prefer Helpers.gs versions

### Helpers.gs (UTILITIES)
- Utility functions: `_logMessage()`, `_parseSender()`, `extractSenderName()`
- Date formatters: `formatDateGmail()`, `formatDateShort()`, `formatDateInput()`
- Email categorization: `isLocalEmail()`, `getSenderCategory()`, `setSenderCategory()`
- PDF extraction: `extractInvoiceLinkFromHtml()`, `extractInvoiceLinkFromPlainText()`
- Folder management: `getExportFolderId()`, `saveFolderId()` (uses H2)

### DriveSyncer.gs (DRIVE INTEGRATION)
- Google Drive sync: `syncInvoicesToDrive()`
- File upload and organization
- Folder management (uses I2 for folder ID - inconsistent with Helpers.gs!)
- Functions: `createUniqueFileName()`, `syncInvoiceFile()`

### runActions.gs (BULK OPERATIONS)
- Bulk actions on selected invoices
- Functions: `runBulkActions()`, `processSelectedAction()`
- Handles: exclude, approve, categorize, export

### utils.gs (ADDITIONAL UTILITIES)
- Additional helper functions
- Safe operations with retry logic
- Validation functions

### archive/ folder ⚠️
- **DO NOT USE** - contains old/deprecated code
- `gmailApiAdvanced.gs` is archived - use `gmailProcessor.gs` instead

## Change Log Template

```
=== CHANGE EXECUTED ===
Date: [DD/MM/YYYY HH:MM]
Task: TASK #XXX - [Task Name]
Files Changed: [List of files]

✅ PRE-CHANGE:
- Read PRD: [Yes/No]
- Made backup: [Yes/No]
- Planned change: [Yes/No]

✅ DURING-CHANGE:
- Followed architecture rules: [Yes/No]
- Maintained code quality: [Yes/No]
- Worked safely: [Yes/No]

✅ POST-CHANGE:
- Basic tests passed: [Yes/No]
- Functional tests passed: [Yes/No]
- Updated documentation: [Yes/No]

🎯 Final Result:
- ✅ Task completed successfully
- ⚠️ Issues need follow-up
- ❌ Change failed - needs rollback

📝 Notes:
[What I learned, what was difficult, what to improve next time]
```

## Remember Always

> "The best code is code that works reliably. Speed comes second to correctness."

- Test thoroughly before deploying
- Document every significant change
- Keep changes small and focused
- Ask when uncertain
- Backup before major changes