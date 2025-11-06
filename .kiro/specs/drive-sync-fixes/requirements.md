# Requirements Document - Drive Sync Fixes

## Introduction

This specification addresses critical issues in the Google Drive synchronization functionality of the Gmail Invoice Scanner. The system currently has two main problems: files are being copied instead of moved, creating duplicates, and the sync process operates on all invoice sheets instead of just the currently active sheet.

## Glossary

- **Drive_Sync_System**: The Google Apps Script module responsible for synchronizing invoice PDFs to Google Drive
- **Active_Sheet**: The currently selected sheet tab that the user is viewing
- **Invoice_Sheet**: Any sheet with a name matching the pattern "invoices YYYY-MM-DD"
- **Source_File**: The original PDF file in Google Drive (usually in My Drive root)
- **Target_Folder**: The organized folder structure where files should be moved
- **File_Move_Operation**: Moving a file from one location to another (not copying)

## Requirements

### Requirement 1

**User Story:** As a user managing invoices, I want files to be moved (not copied) during sync, so that I don't have duplicate files cluttering my Google Drive.

#### Acceptance Criteria

1. WHEN the Drive_Sync_System processes a Google Drive PDF link, THE Drive_Sync_System SHALL move the file from its current location to the target folder
2. WHEN a file is moved to the target folder, THE Drive_Sync_System SHALL ensure the file no longer exists in its original location
3. WHEN the sync operation completes, THE Drive_Sync_System SHALL verify that no duplicate files were created
4. IF a file move operation fails, THEN THE Drive_Sync_System SHALL log the error and continue with other files
5. THE Drive_Sync_System SHALL rename the moved file to follow the organized naming convention

### Requirement 2

**User Story:** As a user working with a specific invoice sheet, I want the sync operation to only process the current sheet I'm viewing, so that I can control which invoices are synchronized.

#### Acceptance Criteria

1. WHEN a user initiates sync from an Invoice_Sheet, THE Drive_Sync_System SHALL process only the Active_Sheet
2. THE Drive_Sync_System SHALL NOT process other Invoice_Sheet instances during a single sync operation
3. WHEN the sync operation starts, THE Drive_Sync_System SHALL display which sheet is being processed
4. IF the Active_Sheet is not an Invoice_Sheet, THEN THE Drive_Sync_System SHALL display an appropriate error message
5. THE Drive_Sync_System SHALL provide progress updates specific to the Active_Sheet being processed

### Requirement 3

**User Story:** As a user, I want clear feedback about the sync operation results, so that I understand what was moved, what was skipped, and if any errors occurred.

#### Acceptance Criteria

1. WHEN the sync operation completes, THE Drive_Sync_System SHALL display a summary showing files moved, skipped, and errors
2. THE Drive_Sync_System SHALL log each file operation with appropriate status (moved, skipped, error)
3. WHEN a file is skipped because it already exists, THE Drive_Sync_System SHALL indicate this in the summary
4. IF errors occur during sync, THEN THE Drive_Sync_System SHALL provide specific error details in the log
5. THE Drive_Sync_System SHALL show the name of the processed sheet in the final summary

### Requirement 4

**User Story:** As a user, I want the sync operation to handle different types of PDF links correctly, so that all invoice files are properly organized regardless of their source.

#### Acceptance Criteria

1. WHEN processing a Google Drive PDF link, THE Drive_Sync_System SHALL extract the file ID and move the existing file
2. WHEN processing an HTTP PDF link, THE Drive_Sync_System SHALL download the file and save it to the target folder
3. THE Drive_Sync_System SHALL maintain the same organized folder structure for all file types
4. IF a file cannot be accessed or downloaded, THEN THE Drive_Sync_System SHALL create a text file with the link as backup
5. THE Drive_Sync_System SHALL apply consistent naming conventions regardless of the file source type