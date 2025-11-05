# Gmail Invoice Scanner - Product Overview

## Purpose
A Google Apps Script-based invoice processing system that automatically scans Gmail for invoices, extracts key information, and organizes them in Google Sheets with Google Drive integration.

## Key Features
- **Automated Gmail Scanning**: Searches Gmail for emails containing invoices based on configurable keywords
- **Smart Invoice Detection**: Identifies invoices using keywords in English and Hebrew, PDF attachments, and content analysis
- **Data Extraction**: Extracts sender information, amounts, currencies, and PDF links
- **Categorization**: Automatically categorizes invoices as Local (Israeli) or International
- **Google Drive Integration**: Uploads PDF attachments to Google Drive with organized folder structure
- **Bulk Actions**: Exclude senders, approve whitelist, mark categories, export invoices
- **Advanced Filtering**: Configurable exclusion lists, approved senders, and keyword management

## Target Users
- Small businesses and freelancers managing invoices from Gmail
- Accountants processing client invoices
- Anyone needing to organize and track invoices from email

## Core Workflow
1. **Setup**: Configure date ranges, keywords, exclusions in Settings sheet
2. **Scan**: Run Gmail API scan to find invoice emails
3. **Process**: Extract data, upload PDFs to Drive, categorize invoices
4. **Manage**: Use bulk actions to exclude, approve, or export invoices
5. **Sync**: Synchronize processed invoices to organized Drive folders

## Technical Architecture
- **Platform**: Google Apps Script (JavaScript runtime)
- **Storage**: Google Sheets for data, Google Drive for PDFs
- **APIs**: Gmail API, Drive API, Sheets API
- **UI**: Custom HTML dialogs and Google Sheets interface