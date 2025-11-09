# Requirements Document - Mini-Log Redesign

## Introduction

This specification defines the requirements for redesigning the Mini-Log feature in the Gmail Invoice Scanner. The Mini-Log is a visual status panel displayed at the top of each invoice results sheet, providing real-time feedback during the scanning process and a summary upon completion.

## Glossary

- **Mini-Log**: A compact status panel (rows 1-9, columns E-F) that displays scan progress and results
- **Invoice Results Sheet**: A Google Sheet tab created for each scan, named "invoices DD.MM.YY-DD.MM.YY"
- **Gmail Scanner**: The system that scans Gmail for invoice emails using the Gmail API
- **Progress Bar**: A visual indicator showing scan completion percentage
- **Log Update**: Real-time messages displayed during the scanning process

---

## Requirements

### Requirement 1: Mini-Log Structure and Layout

**User Story:** As a user running an invoice scan, I want to see a clear, organized status panel at the top of my results sheet, so that I can quickly understand what's happening during the scan.

#### Acceptance Criteria

1. WHEN the Gmail Scanner creates a new invoice results sheet, THE Mini-Log SHALL be positioned in cells E1:F9
2. THE Mini-Log SHALL display exactly 8 labeled rows with the following structure:
   - Row 1: "google search" label with the Gmail query string
   - Row 2: "Start time" label with timestamp (format: HH:MM:SS)
   - Row 3: "End Time" label with timestamp and duration (format: HH:MM:SS - M:SS)
   - Row 4: "Fetched mails" label with count of emails found
   - Row 5: "Progress" label with percentage and Hebrew progress text
   - Row 6: "Log Update" label with real-time status messages
   - Row 7: "final result" label with summary statistics
   - Row 8: "ZIP File" label with download link
3. THE Mini-Log SHALL use column E for labels and column F for content values
4. THE Mini-Log SHALL maintain consistent row heights for readability
5. THE Mini-Log SHALL use text wrapping in column F to accommodate long content

### Requirement 2: Visual Styling and Colors

**User Story:** As a user viewing the Mini-Log, I want it to be visually distinct and easy to read, so that I can quickly identify different types of information.

#### Acceptance Criteria

1. THE Mini-Log labels (column E) SHALL have a yellow background color (#ffe599)
2. THE Mini-Log content cells (column F) SHALL have a white or light blue background (#cfe2f3 for "Log Update" row)
3. THE Mini-Log text SHALL use Arial font at 11pt size
4. THE Mini-Log SHALL use bold text for labels in column E
5. THE Mini-Log SHALL apply borders (0.74pt solid #d9d9d9) around content cells when appropriate
6. THE "Log Update" row (F6) SHALL have a light blue background (#cfe2f3) to distinguish it from other rows

### Requirement 3: Real-Time Progress Updates

**User Story:** As a user watching a scan in progress, I want to see live updates of what's happening, so that I know the system is working and can estimate completion time.

#### Acceptance Criteria

1. WHEN the scan starts, THE Mini-Log SHALL display the start time in cell F2
2. WHILE the scan is running, THE Mini-Log SHALL update the "Progress" cell (F5) with current percentage and Hebrew text (e.g., "68% | סרק 1 מתוך 32")
3. WHILE processing each email, THE Mini-Log SHALL append status messages to the "Log Update" cell (F6)
4. THE "Log Update" messages SHALL include:
   - Invoice detection confirmations with ✅ emoji
   - Sender information and invoice numbers
   - Running counts of processed, skipped, and added invoices
   - Color-coded emojis (📊 for statistics, 🎉 for completion)
5. WHEN the scan completes, THE Mini-Log SHALL display the end time and duration in cell F3

### Requirement 4: Gmail Query Display

**User Story:** As a user reviewing scan results, I want to see the exact Gmail query that was used, so that I can verify the search parameters and troubleshoot if needed.

#### Acceptance Criteria

1. THE Mini-Log SHALL display the complete Gmail API query string in cell F1
2. THE query string SHALL include all search parameters: date range, attachment filters, exclusions, and keywords
3. THE query string SHALL be formatted as a single line with proper Gmail search syntax
4. THE query string SHALL match the actual query executed by the Gmail Scanner

### Requirement 5: Summary Statistics

**User Story:** As a user who has completed a scan, I want to see a clear summary of what was found and processed, so that I can quickly assess the results.

#### Acceptance Criteria

1. WHEN the scan completes, THE Mini-Log SHALL display summary statistics in cell F7
2. THE summary statistics SHALL include:
   - Total number of invoices added
   - Number of emails skipped
   - Number of duplicate invoices found
   - Total emails checked
   - Confirmation message with 🎉 emoji
3. THE summary text SHALL be in Hebrew for user-facing messages
4. THE summary SHALL use emojis (✅, 📊, 🎉) to enhance readability

### Requirement 6: ZIP File Download Link

**User Story:** As a user who wants to download all invoice PDFs, I want a convenient download link in the Mini-Log, so that I can easily access the ZIP file.

#### Acceptance Criteria

1. WHEN PDF files are available for download, THE Mini-Log SHALL display a ZIP file link in cell F8
2. THE ZIP file link SHALL be a clickable Google Drive URL
3. THE ZIP file link SHALL be formatted as a hyperlink with blue text (#0000ff) and underline
4. THE ZIP file link SHALL include the 📦 emoji in the label (cell E8)
5. IF no ZIP file is available, THE cell F8 SHALL remain empty

### Requirement 7: Error Handling and Edge Cases

**User Story:** As a user encountering errors during a scan, I want to see clear error messages in the Mini-Log, so that I understand what went wrong and can take corrective action.

#### Acceptance Criteria

1. IF the scan encounters an error, THE Mini-Log SHALL display an error message in the "Log Update" cell (F6)
2. THE error messages SHALL be in Hebrew and include relevant context
3. IF no emails are found, THE Mini-Log SHALL display "0 מיילים נמצאו" in the "Fetched mails" cell
4. IF the scan is interrupted, THE Mini-Log SHALL preserve the last known state
5. THE Mini-Log SHALL handle long text content by wrapping within the cell boundaries

### Requirement 8: Performance and Responsiveness

**User Story:** As a user running a scan with many emails, I want the Mini-Log to update smoothly without slowing down the scan, so that I get timely results.

#### Acceptance Criteria

1. THE Mini-Log updates SHALL NOT cause the scan process to slow down by more than 5%
2. THE Mini-Log SHALL batch updates to minimize API calls to Google Sheets
3. THE Mini-Log SHALL use SpreadsheetApp.flush() only when necessary to ensure visibility
4. THE Progress updates SHALL occur at reasonable intervals (e.g., every 10 emails or every 5 seconds)
5. THE Mini-Log SHALL complete all updates within 2 seconds of scan completion

---

## Notes

- The Mini-Log design is based on the existing implementation in `docs/invoiceApp_30_9.4_Qodo.xml`
- Color codes and styling match the Google Sheets color palette for consistency
- Hebrew text is used for user-facing messages to match the application's primary language
- Emojis are used strategically to enhance visual communication and user experience
