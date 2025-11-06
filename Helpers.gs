// === Helpers.gs ===

// === GLOBALS ===
const PRIVATE_DOMAINS = [
  "gmail.com", "walla.co.il", "yahoo.com", "hotmail.com", "outlook.com", 
  "aol.com", "mail.ru", "yandex.ru", "protonmail.com", "icloud.com",
  "zoho.com", "gmx.com", "live.com", "msn.com", "me.com"
]; // הרחבת רשימת הדומיינים הפרטיים

function _logMessage(sheet, msg, type = 'INFO') {
  // בדיקת תקינות הגיליון
  if (!sheet || typeof sheet.appendRow !== 'function') {
    console.error("אובייקט גיליון לא תקין נשלח ל-_logMessage");
    console.log(`[${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")}] ${msg}`);
    return;
  }

  try {
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const row = sheet.getLastRow() + 1;
    
    // הוספת השורה עם חותמת זמן והודעה
    sheet.appendRow([timestamp, msg]);
    
    // החלת עיצוב לפי סוג ההודעה
    const range = sheet.getRange(row, 1, 1, 2);
    
    switch (type.toUpperCase()) {
      case 'CRITICAL':
        range.setBackground('#f4cccc').setFontWeight('bold');
        if (!msg.startsWith('❌')) {
          sheet.getRange(row, 2).setValue('❌ ' + msg);
        }
        break;
      case 'ERROR':
        range.setBackground('#fce5cd');
        if (!msg.startsWith('❌')) {
          sheet.getRange(row, 2).setValue('❌ ' + msg);
        }
        break;
      case 'WARNING':
        range.setBackground('#fff2cc');
        if (!msg.startsWith('⚠️')) {
          sheet.getRange(row, 2).setValue('⚠️ ' + msg);
        }
        break;
      case 'SUCCESS':
        range.setBackground('#d9ead3');
        if (!msg.startsWith('✅')) {
          sheet.getRange(row, 2).setValue('✅ ' + msg);
        }
        break;
      case 'SKIPPED':
        range.setBackground('#f3f3f3');
        if (!msg.startsWith('⏩')) {
          sheet.getRange(row, 2).setValue('⏩ ' + msg);
        }
        break;
      case 'DEBUG':
        range.setBackground('#f0f0f0');
        if (!msg.startsWith('🐛')) {
          sheet.getRange(row, 2).setValue('🐛 ' + msg);
        }
        break;
      case 'PERFORMANCE':
        range.setBackground('#e1f5fe');
        if (!msg.startsWith('⚡')) {
          sheet.getRange(row, 2).setValue('⚡ ' + msg);
        }
        break;
      case 'INFO':
      default:
        range.setBackground('#e6f3ff');
        if (!msg.startsWith('ℹ️') && !msg.startsWith('📨') && !msg.startsWith('✅') && 
            !msg.startsWith('❌') && !msg.startsWith('⚠️') && !msg.startsWith('🔍') && 
            !msg.startsWith('📧') && !msg.startsWith('📋') && !msg.startsWith('🚀') && 
            !msg.startsWith('💾') && !msg.startsWith('🔗') && !msg.startsWith('🎉')) {
          sheet.getRange(row, 2).setValue('ℹ️ ' + msg);
        }
        break;
    }
    range.setBorder(true, true, true, true, false, false, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);
    sheet.autoResizeColumn(2);
  } catch (e) {
    console.error(`Error writing to log sheet: ${e.message}`);
    console.log(`[${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")}] ${msg}`);
  }
}

function _parseSender(from) {
  if (!from || typeof from !== 'string') return 'unknown';
  const match = from.match(/<(.+?)>/);
  return match ? match[1] : from;
}

function extractSenderName(email) {
  if (!email || typeof email !== 'string') return "unknown";
  if (email.includes('@')) {
    const parts = email.split("@");
    if (parts.length < 2) return "unknown";
    const username = parts[0];
    const domain = parts[1];
    const cleanUsername = username
      .replace(/[._-]/g, ' ')
      .replace(/no[- ]?reply/gi, '')
      .replace(/admin/gi, '')
      .replace(/support/gi, '')
      .replace(/info/gi, '')
      .trim();
    if (cleanUsername && cleanUsername.length > 2) {
      return cleanUsername;
    }
    const domainParts = domain.split(".");
    return domainParts.length > 0 ? domainParts[0] : "unknown";
  }
  return email;
}

function formatDateGmail(date) {
  if (!date || !(date instanceof Date)) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

function formatDateShort(date) {
  if (!date || !(date instanceof Date)) return "";
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${dd}.${mm}.${String(yyyy).slice(2)}`;
}

function formatDateInput(date) {
  if (!date || !(date instanceof Date)) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// === LOCAL EMAIL CHECKER ===
// ✅ V32: לוגיקה פשוטה ומהירה - אינדיקטור ויזואלי בלבד
function isLocalEmail(email, subject, body) {
  const localDomains = [".co.il", ".org.il", ".gov.il", ".muni.il", ".ac.il"];
  const isLocalDomain = localDomains.some(domain => email.toLowerCase().endsWith(domain));
  
  // Check for Hebrew characters in subject or body
  const hasHebrew = /[\u0590-\u05FF]/.test(subject) || /[\u0590-\u05FF]/.test(body);
  
  return isLocalDomain || hasHebrew;
}

function hasPdfAttachment(message) {
  if (!message || typeof message.getAttachments !== 'function') return null;
  const attachments = message.getAttachments();
  for (const att of attachments) {
    const fileName = att.getName().toLowerCase();
    const contentType = att.getContentType().toLowerCase();

    if (
      fileName.endsWith(".pdf") ||
      contentType === "application/pdf" ||
      contentType.includes("pdf") ||
      (contentType.includes("octet-stream") && /pdf/i.test(fileName)) || // Check for common misidentified PDFs
      (contentType.includes("application/") && fileName.endsWith(".pdf"))
    ) {
      return att;
    }
  }
  return null;
}

function extractInvoiceLinkFromHtml(html) {
  if (!html || typeof html !== 'string') {
    console.log('❌ HTML input is invalid or empty.');
    return null;
  }

  const patterns = [
    /<a[^>]+href\s*=\s*["']([^"']*?\.pdf)["']/gi,
    /<a[^>]+href\s*=\s*["']([^"']*?(?:download|invoice|receipt|bill|payment|view)[^"']*?)["']/gi,
    /<button[^>]*onclick\s*=\s*["'][^"']*?window\.open\(['"](https?:\/\/[^'"]+)['"]/gi,
    /<iframe[^>]+src\s*=\s*["']([^"']*?\.pdf)["']/gi
  ];

  for (const regex of patterns) {
    regex.lastIndex = 0;
    const match = regex.exec(html);
    if (match && match[1]) {
      const originalLink = match[1];
      console.log(`✅ נמצא לינק ב-HTML: ${originalLink.substring(0, 80)}...`);
      // פשוט מחזיר את הלינק המקורי - לא מעלה ל-Drive!
      return originalLink;
    }
  }

  console.log('⚠️ לא נמצא לינק ב-HTML');
  return null;
}

function extractInvoiceLinkFromPlainText(body) {
  if (!body || typeof body !== 'string') {
    console.log('❌ Plain text input is invalid or empty.');
    return null;
  }

  const urlRegexPatterns = [
    /(https?:\/\/[^\s"'<>]+\.pdf)/gi,
    /(https?:\/\/[^\s"'<>]+(?:download|invoice|receipt|bill|payment|view)[^\s"'<>]*)/gi,
    /(?:(?:invoice|receipt|bill|download|view).{0,30})(https?:\/\/[^\s"'<>]+)/gi
  ];

  for (const regex of urlRegexPatterns) {
    const matches = body.match(regex);
    if (matches) {
      for (let url of matches) {
        // ניקוי URL מתווים מיותרים בסוף
        url = url.replace(/[.,;:)\]}>]+$/, '');
        
        console.log(`✅ נמצא לינק בטקסט: ${url.substring(0, 80)}...`);
        // פשוט מחזיר את הלינק המקורי - לא מעלה ל-Drive!
        return url;
      }
    }
  }

  console.log('⚠️ לא נמצא לינק בטקסט');
  return null;
}

// פונקציה בטוחה ללוגים - מונעת קריסות
function safeLogMessage(logSheet, message) {
  if (!logSheet) {
    console.log(message);
    return;
  }
  try {
    const timestamp = new Date().toLocaleString('he-IL');
    logSheet.appendRow([timestamp, message]);
  } catch (e) {
    console.error("שגיאה בכתיבה ללוג:", e);
    console.log(message);
  }
}

function setExportFolder(folderUrl) {
  const folderId = extractFolderIdFromUrl(folderUrl);
  if (folderId) {
    saveFolderId(folderId);
    SpreadsheetApp.getUi().alert("✅ מזהה התיקייה נשמר בהצלחה.");
  } else {
    SpreadsheetApp.getUi().alert("❌ כתובת התיקייה אינה תקינה.");
  }
}

function extractFolderIdFromUrl(url) {
  if (!url || typeof url !== 'string') return null;
  const folderRegex = /folders\/([a-zA-Z0-9_-]+)(\?|$)/i;
  const match = url.match(folderRegex);
  return match ? match[1] : null;
}

function saveFolderId(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName("Settings");
  if (!settings) return;

  settings.getRange("I2").setValue(id);
}

function getExportFolderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName("Settings");
  if (!settings) return null;

  return settings.getRange("I2").getValue();
}