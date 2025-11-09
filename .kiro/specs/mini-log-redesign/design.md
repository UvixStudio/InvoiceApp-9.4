# Design Document - Mini-Log Redesign

## Overview

The Mini-Log is a visual status panel that provides real-time feedback during Gmail invoice scanning. It occupies cells E1:F9 in each invoice results sheet, displaying scan progress, statistics, and download links in a compact, color-coded format.

## Architecture

### Component Structure

```
Mini-Log Component
├── MiniLogManager (Main Controller)
│   ├── initializeMiniLog()
│   ├── updateProgress()
│   ├── appendLogMessage()
│   └── finalizeMiniLog()
├── MiniLogFormatter (Styling & Layout)
│   ├── applyMiniLogStyles()
│   ├── formatProgressText()
│   └── formatSummaryText()
└── MiniLogData (Data Model)
    ├── startTime
    ├── endTime
    ├── gmailQuery
    ├── fetchedCount
    ├── progressPercent
    ├── logMessages[]
    └── zipFileUrl
```

### Integration Points

The Mini-Log integrates with the existing `gmailProcessor.gs` scanning flow:

1. **Initialization**: Called at the start of `processInvoices()`
2. **Progress Updates**: Called during email processing loop
3. **Finalization**: Called at the end of scan with summary statistics

---

## Data Models

### MiniLogData Object

```javascript
const miniLogData = {
  // Scan metadata
  startTime: Date,           // Scan start timestamp
  endTime: Date,             // Scan end timestamp
  duration: String,          // Formatted duration (M:SS)
  
  // Gmail query
  gmailQuery: String,        // Full Gmail API query string
  
  // Scan statistics
  fetchedCount: Number,      // Total emails found
  processedCount: Number,    // Emails processed so far
  addedCount: Number,        // Invoices added
  skippedCount: Number,      // Emails skipped
  duplicateCount: Number,    // Duplicate invoices found
  
  // Progress tracking
  progressPercent: Number,   // Current progress (0-100)
  progressText: String,      // Hebrew progress text
  
  // Log messages
  logMessages: Array,        // Array of status messages
  
  // ZIP file
  zipFileUrl: String         // Google Drive download link
};
```

### Cell Mapping

| Row | Column E (Label) | Column F (Content) | Data Source |
|-----|------------------|-------------------|-------------|
| 1 | "google search" | Gmail query string | `miniLogData.gmailQuery` |
| 2 | "Start time" | HH:MM:SS | `miniLogData.startTime` |
| 3 | "End Time" | HH:MM:SS - M:SS | `miniLogData.endTime` + `duration` |
| 4 | "Fetched mails" | Count | `miniLogData.fetchedCount` |
| 5 | "Progress" | XX% \| Hebrew text | `miniLogData.progressPercent` + `progressText` |
| 6 | "Log Update" | Status messages | `miniLogData.logMessages.join('\n')` |
| 7 | "final result" | Summary statistics | Calculated from counts |
| 8 | "📦 ZIP File" | Download link | `miniLogData.zipFileUrl` |

---

## Components and Interfaces

### 1. MiniLogManager

**Purpose**: Main controller for Mini-Log operations

**Key Functions**:

```javascript
/**
 * Initialize Mini-Log at the start of scan
 * @param {Sheet} sheet - The invoice results sheet
 * @param {String} gmailQuery - The Gmail API query string
 * @returns {Object} miniLogData object
 */
function initializeMiniLog(sheet, gmailQuery) {
  const miniLogData = {
    startTime: new Date(),
    gmailQuery: gmailQuery,
    fetchedCount: 0,
    processedCount: 0,
    addedCount: 0,
    skippedCount: 0,
    duplicateCount: 0,
    progressPercent: 0,
    logMessages: [],
    zipFileUrl: ''
  };
  
  // Set up Mini-Log structure
  setupMiniLogStructure(sheet);
  
  // Write initial values
  sheet.getRange('F1').setValue(gmailQuery);
  sheet.getRange('F2').setValue(Utilities.formatDate(miniLogData.startTime, 
    Session.getScriptTimeZone(), 'HH:mm:ss'));
  
  // Apply styling
  applyMiniLogStyles(sheet);
  
  return miniLogData;
}

/**
 * Update progress during scan
 * @param {Sheet} sheet - The invoice results sheet
 * @param {Object} miniLogData - Current Mini-Log data
 * @param {Number} current - Current email index
 * @param {Number} total - Total emails to process
 */
function updateMiniLogProgress(sheet, miniLogData, current, total) {
  miniLogData.processedCount = current;
  miniLogData.progressPercent = Math.round((current / total) * 100);
  miniLogData.progressText = `סרק ${current} מתוך ${total}`;
  
  const progressDisplay = `${miniLogData.progressPercent}% | ${miniLogData.progressText}`;
  sheet.getRange('F5').setValue(progressDisplay);
  
  // Flush every 10 emails or at 100%
  if (current % 10 === 0 || current === total) {
    SpreadsheetApp.flush();
  }
}

/**
 * Append a log message
 * @param {Sheet} sheet - The invoice results sheet
 * @param {Object} miniLogData - Current Mini-Log data
 * @param {String} message - Message to append
 */
function appendMiniLogMessage(sheet, miniLogData, message) {
  miniLogData.logMessages.push(message);
  
  // Update Log Update cell (F6)
  const logText = miniLogData.logMessages.join('\n');
  sheet.getRange('F6').setValue(logText);
}

/**
 * Finalize Mini-Log at end of scan
 * @param {Sheet} sheet - The invoice results sheet
 * @param {Object} miniLogData - Current Mini-Log data
 * @param {String} zipFileUrl - ZIP file download URL
 */
function finalizeMiniLog(sheet, miniLogData, zipFileUrl) {
  miniLogData.endTime = new Date();
  const durationMs = miniLogData.endTime - miniLogData.startTime;
  miniLogData.duration = formatDuration(durationMs);
  
  // Update End Time (F3)
  const endTimeText = Utilities.formatDate(miniLogData.endTime, 
    Session.getScriptTimeZone(), 'HH:mm:ss') + ' - ' + miniLogData.duration;
  sheet.getRange('F3').setValue(endTimeText);
  
  // Update final result (F7)
  const summaryText = formatSummaryText(miniLogData);
  sheet.getRange('F7').setValue(summaryText);
  
  // Update ZIP file link (F8)
  if (zipFileUrl) {
    sheet.getRange('F8').setValue(zipFileUrl);
    sheet.getRange('F8').setFontColor('#0000ff').setFontLine('underline');
  }
  
  SpreadsheetApp.flush();
}
```

### 2. MiniLogFormatter

**Purpose**: Handle styling and text formatting

**Key Functions**:

```javascript
/**
 * Apply visual styling to Mini-Log cells
 * @param {Sheet} sheet - The invoice results sheet
 */
function applyMiniLogStyles(sheet) {
  // Column E (Labels) - Yellow background, bold text
  const labelRange = sheet.getRange('E1:E8');
  labelRange.setBackground('#ffe599')
           .setFontWeight('bold')
           .setFontFamily('Arial')
           .setFontSize(11);
  
  // Column F (Content) - White background
  const contentRange = sheet.getRange('F1:F8');
  contentRange.setBackground('#ffffff')
             .setFontFamily('Arial')
             .setFontSize(11)
             .setWrap(true);
  
  // F6 (Log Update) - Light blue background
  sheet.getRange('F6').setBackground('#cfe2f3');
  
  // Apply borders
  const miniLogRange = sheet.getRange('E1:F8');
  miniLogRange.setBorder(true, true, true, true, true, true, 
                         '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Format progress text in Hebrew
 * @param {Number} current - Current count
 * @param {Number} total - Total count
 * @returns {String} Formatted Hebrew text
 */
function formatProgressText(current, total) {
  return `סרק ${current} מתוך ${total}`;
}

/**
 * Format summary statistics text
 * @param {Object} miniLogData - Mini-Log data
 * @returns {String} Formatted summary in Hebrew
 */
function formatSummaryText(miniLogData) {
  const parts = [];
  
  parts.push(`✅ ${miniLogData.addedCount} חשבוניות נוספו`);
  parts.push(`📊 ${miniLogData.skippedCount} מיילים דולגו`);
  parts.push(`${miniLogData.duplicateCount} כפילויות נמצאו`);
  parts.push(`${miniLogData.fetchedCount} מיילים נבדקו בסך הכל`);
  parts.push(`🎉 הסריקה הושלמה!`);
  
  return parts.join(' • ');
}

/**
 * Format duration from milliseconds
 * @param {Number} durationMs - Duration in milliseconds
 * @returns {String} Formatted duration (M:SS)
 */
function formatDuration(durationMs) {
  const totalSeconds = Math.floor(durationMs / 1000);
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}:${seconds.toString().padStart(2, '0')}`;
}
```

### 3. Setup Helper

```javascript
/**
 * Set up Mini-Log structure with labels
 * @param {Sheet} sheet - The invoice results sheet
 */
function setupMiniLogStructure(sheet) {
  const labels = [
    ['google search', ''],
    ['Start time', ''],
    ['End Time', ''],
    ['Fetched mails', ''],
    ['Progress', ''],
    ['Log Update', ''],
    ['final result', ''],
    ['📦 ZIP File', '']
  ];
  
  sheet.getRange('E1:F8').setValues(labels);
}
```

---

## Integration with gmailProcessor.gs

### Modified Scan Flow

```javascript
function processInvoices() {
  try {
    // ... existing setup code ...
    
    // Create results sheet
    const resultsSheet = createResultsSheet(startDate, endDate);
    
    // Initialize Mini-Log
    const gmailQuery = buildGmailQuery(startDate, endDate, keywords);
    const miniLogData = initializeMiniLog(resultsSheet, gmailQuery);
    
    // Fetch emails
    const threads = GmailApp.search(gmailQuery);
    miniLogData.fetchedCount = threads.length;
    resultsSheet.getRange('F4').setValue(threads.length);
    
    // Process emails
    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      
      // Update progress
      updateMiniLogProgress(resultsSheet, miniLogData, i + 1, threads.length);
      
      // Process email
      const result = processEmail(thread);
      
      // Log result
      if (result.added) {
        miniLogData.addedCount++;
        appendMiniLogMessage(resultsSheet, miniLogData, 
          `✅ חשבונית ${result.invoiceNumber} - ${result.sender}`);
      } else if (result.skipped) {
        miniLogData.skippedCount++;
      } else if (result.duplicate) {
        miniLogData.duplicateCount++;
      }
    }
    
    // Create ZIP file
    const zipFileUrl = createZipFile(resultsSheet);
    
    // Finalize Mini-Log
    finalizeMiniLog(resultsSheet, miniLogData, zipFileUrl);
    
  } catch (error) {
    // Error handling
    appendMiniLogMessage(resultsSheet, miniLogData, 
      `❌ שגיאה: ${error.message}`);
    throw error;
  }
}
```

---

## Error Handling

### Error Display Strategy

1. **Scan Errors**: Display in Log Update cell (F6) with ❌ emoji
2. **No Results**: Show "0 מיילים נמצאו" in Fetched mails cell (F4)
3. **Interrupted Scan**: Preserve last known state in all cells
4. **Long Text**: Use text wrapping to prevent overflow

### Error Message Format

```javascript
function handleMiniLogError(sheet, miniLogData, error) {
  const errorMessage = `❌ שגיאה: ${error.message}`;
  appendMiniLogMessage(sheet, miniLogData, errorMessage);
  
  // Mark scan as incomplete
  sheet.getRange('F7').setValue('⚠️ הסריקה לא הושלמה');
  
  SpreadsheetApp.flush();
}
```

---

## Performance Optimization

### Batching Strategy

1. **Progress Updates**: Every 10 emails or every 5 seconds
2. **Log Messages**: Batch append every 5 messages
3. **Flush Calls**: Minimize to critical points only

### Optimization Techniques

```javascript
// Batch progress updates
let lastFlushTime = Date.now();
const FLUSH_INTERVAL_MS = 5000;

function updateMiniLogProgressOptimized(sheet, miniLogData, current, total) {
  miniLogData.processedCount = current;
  miniLogData.progressPercent = Math.round((current / total) * 100);
  
  const progressDisplay = `${miniLogData.progressPercent}% | סרק ${current} מתוך ${total}`;
  sheet.getRange('F5').setValue(progressDisplay);
  
  // Flush based on interval or completion
  const now = Date.now();
  if (now - lastFlushTime > FLUSH_INTERVAL_MS || current === total) {
    SpreadsheetApp.flush();
    lastFlushTime = now;
  }
}
```

---

## Testing Strategy

### Unit Tests

1. **MiniLogManager Functions**:
   - Test `initializeMiniLog()` creates correct structure
   - Test `updateMiniLogProgress()` calculates percentages correctly
   - Test `finalizeMiniLog()` formats summary correctly

2. **MiniLogFormatter Functions**:
   - Test `formatProgressText()` with various counts
   - Test `formatSummaryText()` with different statistics
   - Test `formatDuration()` with various time spans

### Integration Tests

1. **Full Scan Flow**:
   - Run scan with 10 test emails
   - Verify Mini-Log updates at each stage
   - Confirm final summary matches actual results

2. **Error Scenarios**:
   - Test with no emails found
   - Test with interrupted scan
   - Test with API errors

### Visual Tests

1. **Styling Verification**:
   - Verify colors match specification (#ffe599, #cfe2f3)
   - Verify fonts and sizes (Arial 11pt)
   - Verify borders and wrapping

---

## Migration Plan

### Phase 1: Create New Functions
- Implement MiniLogManager functions
- Implement MiniLogFormatter functions
- Keep existing code intact

### Phase 2: Integration
- Modify `processInvoices()` to use new Mini-Log
- Test with small data sets
- Compare output with existing implementation

### Phase 3: Cleanup
- Remove old Mini-Log code
- Update documentation
- Deploy to production

---

## Notes

- All Hebrew text uses right-to-left (RTL) formatting
- Emojis are used consistently: ✅ (success), ❌ (error), 📊 (stats), 🎉 (completion), 📦 (ZIP)
- Color codes match Google Sheets standard palette
- Performance target: <5% overhead on scan time
