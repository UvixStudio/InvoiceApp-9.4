/**
 * Utility functions for the Gmail Invoice Processor Add-on
 */

/**
 * Adds an email address to the list of approved senders.
 * @param {string} email - The email address to add as an approved sender
 * @returns {boolean} True if the email was added successfully, false if it already exists
 */
function addApprovedSender(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const approvedSenders = sheet.getRange('B2').getValue().split(',').map(s => s.trim());
  
  if (approvedSenders.includes(email)) {
    return false;
  }
  
  approvedSenders.push(email);
  sheet.getRange('B2').setValue(approvedSenders.join(', '));
  return true;
}

/**
 * Sanitizes a filename by removing or replacing invalid characters.
 * @param {string} filename - The filename to sanitize
 * @returns {string} The sanitized filename
 */
function sanitizeFileName(filename) {
  return filename.replace(/[^a-zA-Z0-9-_. ]/g, '_')
                 .replace(/\s+/g, '_')
                 .trim();
}

/**
 * Extracts the file ID from a Google Drive URL.
 * @param {string} url - The Google Drive URL
 * @returns {string|null} The file ID if found, null otherwise
 */
function extractFileIdFromUrl(url) {
  if (!url) return null;
  
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * Creates a timestamp in the format YYYY-MM-DD HH:mm:ss.
 * @returns {string} Formatted timestamp
 */
function getFormattedTimestamp() {
  const now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Logs an error message with timestamp to the specified sheet.
 * @param {string} message - The error message to log
 * @param {string} [sheetName='Logs'] - The name of the sheet to log to
 */
function logError(message, sheetName = 'Logs') {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  
  const timestamp = getFormattedTimestamp();
  sheet.appendRow([timestamp, 'ERROR', message]);
}

/**
 * Validates an email address format.
 * @param {string} email - The email address to validate
 * @returns {boolean} True if the email format is valid
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Creates a safe wrapper for sheet operations with retry logic
 * @param {Function} operation - The operation to perform
 * @param {number} maxRetries - Maximum number of retries
 * @returns {*} Result of the operation
 */
function safeSheetOperation(operation, maxRetries = 3) {
  let lastError = null;
  
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return operation();
    } catch (error) {
      lastError = error;
      if (attempt < maxRetries - 1) {
        Utilities.sleep(1000 * Math.pow(2, attempt)); // Exponential backoff
      }
    }
  }
  
  throw lastError;
}

/**
 * Safely appends a row to a sheet with error handling
 * @param {Sheet} sheet - The sheet to append to
 * @param {Array} values - The values to append
 * @returns {boolean} True if successful
 */
function safeAppendRow(sheet, values) {
  try {
    return safeSheetOperation(() => {
      sheet.appendRow(values);
      return true;
    });
  } catch (error) {
    console.error(`Error appending row: ${error.message}`);
    return false;
  }
}

/**
 * Checks if a URL is valid
 * @param {string} url - The URL to validate
 * @returns {boolean} True if the URL is valid
 */
function isValidUrl(url) {
  try {
    new URL(url);
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Truncates text to a specified length
 * @param {string} text - The text to truncate
 * @param {number} maxLength - Maximum length
 * @returns {string} Truncated text
 */
function truncateText(text, maxLength = 100) {
  if (!text || typeof text !== 'string') return '';
  return text.length > maxLength ? text.substring(0, maxLength) + '...' : text;
}

/**
 * Converts a date to Gmail search format
 * @param {Date} date - The date to format
 * @returns {string} Formatted date string
 */
function formatDateGmail(date) {
  if (!date || !(date instanceof Date)) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

/**
 * Parses sender information from email header
 * @param {string} from - The from header
 * @returns {string} Clean email address
 */
function parseSenderEmail(from) {
  if (!from || typeof from !== 'string') return 'unknown';
  const match = from.match(/<(.+?)>/);
  return match ? match[1] : from;
}

/**
 * Checks if an email domain is private (personal)
 * @param {string} email - Email address to check
 * @returns {boolean} True if private domain
 */
function isPrivateDomain(email) {
  const privateDomains = [
    'gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'walla.co.il',
    'aol.com', 'mail.ru', 'yandex.ru', 'protonmail.com', 'icloud.com',
    'zoho.com', 'gmx.com', 'live.com', 'msn.com', 'me.com'
  ];
  
  const domain = email.toLowerCase().split('@')[1];
  return privateDomains.includes(domain);
}