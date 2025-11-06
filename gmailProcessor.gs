// === GmailProcessor.gs - סריקה חדשה ופועלת ===

/**
 * יצירת progress bar ויזואלי לטוסט
 */
function createScanProgressBar(current, total, width = 20) {
  const percentage = Math.round((current / total) * 100);
  const filled = Math.round((current / total) * width);
  const empty = width - filled;
  
  // רק בלוקים מלאים ורווחים - נקי וברור
  const bar = '█'.repeat(filled) + ' '.repeat(empty);
  return `|${bar}| ${percentage}%`;
}

/**
 * המרת שניות לפורמט זמן קריא
 */
function formatScanDuration(seconds) {
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const secs = seconds % 60;
  
  if (hours > 0) {
    return `${hours}:${minutes.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
  } else if (minutes > 0) {
    return `${minutes}:${secs.toString().padStart(2, '0')}`;
  } else {
    return `${secs}s`;
  }
}

// === HELPER: Check Drive Permissions ===
function checkDrivePermissions() {
  try {
    // ניסיון פשוט לגשת ל-Drive
    DriveApp.getRootFolder();
    return { success: true, message: "הרשאות Drive תקינות" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// === MAIN SCAN FUNCTION ===
function processInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!settings) throw new Error('Settings sheet not found.');
  if (!log) throw new Error('Log sheet not found.');
  
  const startTime = new Date();
  _logMessage(log, `🔍 התחלת סריקת חשבוניות - ${startTime.toLocaleString('he-IL')}`, 'INFO');
  
  try {
    // בדיקת הרשאות
    const driveCheck = checkDrivePermissions();
    if (!driveCheck.success) {
      SpreadsheetApp.getUi().alert(
        "❌ בעיית הרשאות Drive\n\n" + driveCheck.message + 
        "\n\nאנא רענן את הדף ואשר הרשאות מחדש"
      );
      return;
    }

    // קריאת הגדרות
    const config = readScanSettings(settings, log);
    if (!config) return;
    
    // הודעת לוג מפורטת על טווח התאריכים
    _logMessage(log, `📅 סורק מיילים מ-${formatDateInput(config.startDate)} עד ${formatDateInput(config.endDate)}`, 'INFO');

    // יצירת/עדכון גיליון תוצאות
    const sheet = createOrUpdateResultsSheet(ss, config.tabName, log);
    if (!sheet) return;

    // ✅ V32: איפוס Mini-Log (F1-G11)
    const statusSheet = sheet;
    const statusRange = statusSheet.getRange("F1:G11");
    statusRange.clearContent();

    // ✅ V32: כותרות Mini-Log
    statusSheet.getRange("F1").setValue("status");
    statusSheet.getRange("G1").setValue("שאילתה");

    // בניית שאילתה
    const query = buildGmailQuery(config);
    
    // ✅ V32: שאילתה ב-Mini-Log
    statusSheet.getRange("G2").setValue(query);
    
    // ✅ V32: זמן התחלה ב-Mini-Log
    statusSheet.getRange("F3").setValue(`זמן התחלה: ${startTime.toLocaleTimeString()}`);
    
    // ✅ V32: מצב ראשוני
    statusSheet.getRange("F6").setValue("מתחיל סריקה...");
    statusSheet.getRange("F7").setValue("0 % |                    |");
    
    _logMessage(log, `⏱️ התחלת סריקה: ${startTime.toLocaleTimeString()}`, 'INFO');
    
    const scanResults = performGmailScan(config, log, startTime, statusSheet);
    const scanEndTime = new Date();
    const scanDuration = Math.round((scanEndTime - startTime) / 1000);
    _logMessage(log, `⏱️ סיום סריקת Gmail: ${scanEndTime.toLocaleTimeString()} - ${formatScanDuration(scanDuration)}`, 'INFO');
    
    // ✅ V32: עדכון Mini-Log - נמצאו מיילים
    statusSheet.getRange("F5").setValue(`נמצאו ${scanResults.emails.length} מיילים בסינון הראשוני`);
    
    if (scanResults.emails.length === 0) {
      statusSheet.getRange("F6").setValue("ℹ️ לא נמצאו מיילים");
      statusSheet.getRange("F9").setValue("לא נמצאו מיילים בטווח התאריכים שהוגדר");
      SpreadsheetApp.getUi().alert("לא נמצאו מיילים בטווח התאריכים שהוגדר.\n\nבדוק:\n1. טווח התאריכים\n2. מילות המפתח\n3. שיש מיילים עם קבצים מצורפים");
      _logMessage(log, "❌ לא נמצאו מיילים לעיבוד", 'WARNING');
      return;
    }

    // עיבוד המיילים
    const processStartTime = new Date();
    _logMessage(log, `⏱️ התחלת עיבוד מיילים: ${processStartTime.toLocaleTimeString()}`, 'INFO');
    const processResults = processEmails(
      scanResults.emails, 
      config, 
      sheet, 
      log, 
      startTime
    );
    const processEndTime = new Date();
    const processDuration = Math.round((processEndTime - processStartTime) / 1000);
    _logMessage(log, `⏱️ סיום עיבוד מיילים: ${processEndTime.toLocaleTimeString()} - ${formatScanDuration(processDuration)}`, 'INFO');
    
    // סיכום וסיום
    showScanSummary(processResults, log);
    
    _logMessage(log, "✅ סריקה הושלמה בהצלחה", 'SUCCESS');
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בסריקה: ${error.message}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`שגיאה בסריקה: ${error.message}`);
  }
}

// === QUICK API SCAN ===
function quickApiScanInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!settings || !log) {
    SpreadsheetApp.getUi().alert("❌ גיליונות Settings או Log חסרים.");
    return;
  }
  
  _logMessage(log, "⚡ התחלת סריקה מהירה", 'INFO');
  
  // קריאה לפונקציה הרגילה - זה יהיה מהיר יותר
  processInvoices();
}

// === SETTINGS READER ===
function readScanSettings(settings, log) {
  try {
    const startStr = settings.getRange("A2").getValue();
    const endStr = settings.getRange("B2").getValue();
    
    if (!startStr || !endStr) {
      SpreadsheetApp.getUi().alert("❌ יש למלא תאריכים תקפים בגיליון Settings.");
      return null;
    }

    const startDate = new Date(startStr);
    const endDate = new Date(endStr);
    endDate.setHours(23, 59, 59, 999);

    if (startDate.getTime() >= endDate.getTime()) {
      SpreadsheetApp.getUi().alert("❌ תאריך התחלה חייב להיות לפני תאריך הסיום.");
      return null;
    }

    const enKeywords = settings.getRange("C2:C").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
    const heKeywords = settings.getRange("D2:D").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
    const excludedKeywords = settings.getRange("E2:E").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
    const excludedEmails = settings.getRange("F2:F").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
    const approvedSenders = settings.getRange("G2:G").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());

    const tabName = `invoices ${formatDateShort(startDate)}-${formatDateShort(endDate)}`;
    
    _logMessage(log, `📅 טווח תאריכים: ${formatDateInput(startDate)} עד ${formatDateInput(endDate)}`, 'INFO');
    _logMessage(log, `🔤 מילות מפתח: EN(${enKeywords.length}), HE(${heKeywords.length})`, 'INFO');
    _logMessage(log, `🚫 מוחרגים: ${excludedEmails.length} מיילים, ${excludedKeywords.length} מילים`, 'INFO');
    _logMessage(log, `✅ מאושרים: ${approvedSenders.length} שולחים`, 'INFO');

    return {
      startDate,
      endDate,
      enKeywords,
      heKeywords,
      excludedKeywords,
      excludedEmails,
      approvedSenders,
      tabName
    };
  } catch (error) {
    _logMessage(log, `❌ שגיאה בקריאת הגדרות: ${error.message}`, 'ERROR');
    return null;
  }
}

// === SHEET CREATOR ===
function createOrUpdateResultsSheet(ss, tabName, log) {
  try {
    let sheet = ss.getSheetByName(tabName);
    
    if (!sheet) {
      _logMessage(log, `📑 יוצר גיליון חדש: ${tabName}`, 'INFO');
      sheet = ss.insertSheet(tabName);
      
      // הגדרת רוחב עמודות
      sheet.setColumnWidth(1, 50);   // Action
      sheet.setColumnWidth(2, 120);  // Category
      sheet.setColumnWidth(3, 150);  // Sender Name
      sheet.setColumnWidth(4, 200);  // Sender Email
      sheet.setColumnWidth(5, 100);  // Date
      sheet.setColumnWidth(6, 300);  // Subject
      sheet.setColumnWidth(7, 250);  // PDF Link
      sheet.setColumnWidth(8, 150);  // Email ID
      sheet.setColumnWidth(9, 150);  // Attachment Name
      sheet.setColumnWidth(10, 100); // Sum Local
      sheet.setColumnWidth(11, 100); // Sum Intl
      sheet.setColumnWidth(12, 80);  // Currency
      sheet.setColumnWidth(13, 250); // Mail Link
      sheet.setColumnWidth(14, 250); // Download Link

      const headers = [
        "Action", "Category", "Sender Name", "Sender Email", "Date", 
        "Subject", "PDF Link", "Email ID", "Attachment Name", 
        "Sum (Local)", "Sum (Int'l)", "Currency",
        "Mail Link", "Download Link"
      ];
      
      // ✅ V32: כותרות בשורה 13 (לא 1!)
      sheet.getRange(13, 1, 1, headers.length).setValues([headers]);
      
      // עיצוב כותרות
      const headerRange = sheet.getRange(13, 1, 1, headers.length);
      headerRange.setFontWeight("bold")
                 .setBackground("#4285f4")
                 .setFontColor("white")
                 .setHorizontalAlignment("center");
      
      // ✅ V32: הקפאת 13 שורות ראשונות (Mini-Log 1-11, רווח 12, כותרות 13)
      sheet.setFrozenRows(13);
      
      // הסתרת עמודת Action
      sheet.hideColumns(1, 1);
      
      SpreadsheetApp.flush();
      
      _logMessage(log, "✅ גיליון תוצאות נוצר בהצלחה", 'SUCCESS');
    } else {
      _logMessage(log, `📋 משתמש בגיליון קיים: ${tabName}`, 'INFO');
      
      // ניקוי צ'קבוקסים ישנים
      try {
        cleanupOldCheckboxes(sheet, log);
      } catch (e) {
        _logMessage(log, `⚠️ שגיאה בניקוי צ'קבוקסים ישנים: ${e.message}`, 'WARNING');
      }
    }
    
    return sheet;
  } catch (error) {
    _logMessage(log, `❌ שגיאה ביצירת גיליון: ${error.message}`, 'ERROR');
    return null;
  }
}

// === GMAIL SCANNER ===
// ✅ V32: הסרת Mini-Log מכאן - הוא מנוהל ב-processInvoices
function performGmailScan(config, log, startTime, statusSheet) {
  try {
    // בניית שאילתת חיפוש
    const query = buildGmailQuery(config);
    _logMessage(log, `שאילתת Gmail:`, 'INFO');
    _logMessage(log, `${query}`, 'INFO');
    
    // חיפוש מיילים
    const threads = GmailApp.search(query, 0, 500);
    _logMessage(log, `📧 נמצאו ${threads.length} שרשורי מייל שעברו סינון ראשוני`, 'INFO');
    
    if (threads.length === 0) {
      return { emails: [], skipped: [] };
    }
    
    // איסוף כל המיילים מהשרשורים
    const allEmails = [];
    let totalMessages = 0;
    
    for (const thread of threads) {
      const messages = thread.getMessages();
      totalMessages += messages.length;
      
      for (const message of messages) {
        const email = extractEmailData(message);
        if (email) {
          allEmails.push(email);
        }
      }
    }
    
    _logMessage(log, `📨 נמצאו ${totalMessages} הודעות ב-${threads.length} שרשורים`, 'INFO');
    _logMessage(log, `📋 חולצו ${allEmails.length} מיילים לעיבוד`, 'INFO');
    
    return { emails: allEmails, skipped: [] };
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בסריקת Gmail: ${error.message}`, 'ERROR');
    throw error;
  }
}

// === GMAIL QUERY BUILDER ===
// Based on KiroAI Invoice Scanner Methodology (04.11.2025)
function buildGmailQuery(config) {
  const parts = [];
  
  // טווח תאריכים - חובה
  parts.push(`after:${formatDateGmail(config.startDate)}`);
  parts.push(`before:${formatDateGmail(config.endDate)}`);
  
  // רק קבצי PDF מצורפים - חובה
  parts.push("has:attachment");
  parts.push("filename:pdf");
  
  // חריגת צ'אטים ושרשורים
  parts.push("-is:chat");
  parts.push("-in:chats");
  parts.push("-is:muted");
  parts.push('-subject:(re: OR fwd: OR fw: OR תשובה OR העברה)');
  
  // מילות חיפוש רק בנושא (subject) - OR ביניהן
  const allKeywords = [...config.enKeywords, ...config.heKeywords].filter(k => k && k.trim());
  
  if (allKeywords.length > 0) {
    // הוספת גרשיים לכל מילת חיפוש - עוזר ל-Gmail לזהות נכון (במיוחד עברית)
    const keywordQuery = allKeywords.map(k => `"${k.trim()}"`).join(" OR ");
    parts.push(`subject:(${keywordQuery})`);
    console.log(`🔍 חיפוש לפי ${allKeywords.length} מילות מפתח (עם גרשיים)`);
  } else {
    throw new Error('❌ לא הוגדרו מילות חיפוש בטאב Settings! אנא הוסף מילות חיפוש בעמודות C ו-D');
  }
  
  // הערה: Approved Senders ו-Excluded Emails יטופלו בשלב הסינון הפנימי (Post-Processing)
  // לא בשאילתת Gmail - כך השאילתה נשארת פשוטה ומהירה
  
  return parts.join(" ");
}

// === EMAIL DATA EXTRACTOR ===
function extractEmailData(message) {
  try {
    const from = message.getFrom();
    const email = _parseSender(from);
    const subject = message.getSubject() || "";
    const date = message.getDate();
    const body = message.getPlainBody() || "";
    const htmlBody = message.getBody() || "";
    const messageId = message.getId();
    
    // בדיקת קבצים מצורפים
    const attachments = message.getAttachments();
    const pdfAttachment = attachments.find(att => 
      att.getName().toLowerCase().endsWith('.pdf') || 
      att.getContentType().toLowerCase().includes('pdf')
    );
    
    return {
      messageId,
      email,
      senderName: extractSenderName(email),
      subject,
      date,
      body,
      htmlBody,
      attachments,
      pdfAttachment,
      from
    };
  } catch (error) {
    console.error(`Error extracting email data: ${error.message}`);
    return null;
  }
}

// === EMAIL PROCESSOR ===
function processEmails(emails, config, sheet, log, startTime) {
  let inserted = 0;
  let skipped = 0;
  let duplicates = 0;
  
  // קבלת רשומות קיימות
  const existingRecords = getExistingRecords(sheet);
  
  _logMessage(log, `🔄 מתחיל עיבוד ${emails.length} מיילים`, 'INFO');
  
  // ✅ V32: Batch Write - בניית מערך
  const allRecords = [];
  
  // ⚡ V32 PERFORMANCE: Batch Logging - אוסף לוגים למערך
  const logBatch = [];
  
  // ✅ V32 CRITICAL: startingRow דינאמי!
  const startingRow = findLastRowWithData(sheet) + 1;
  _logMessage(log, `📝 רשומות חדשות יתחילו משורה ${startingRow}`, 'INFO');
  
  // ✅ V32: הפניה ל-Mini-Log
  const statusSheet = sheet;
  
  // ⚡ V32 FIXED: Batch Write אמיתי - כתיבה אחת בסוף!
  for (let i = 0; i < emails.length; i++) {
    const emailData = emails[i];
    
    try {
      // בדיקת כפילויות
      if (existingRecords.has(emailData.messageId)) {
        duplicates++;
        continue;
      }
      
      // יצירת רשומה ודחיפה למערך
      try {
        const record = createInvoiceRecord(emailData, config);
        allRecords.push(record);  // ← רק דוחף למערך, לא כותב!
        
        existingRecords.set(emailData.messageId, true);
        inserted++;
        
        // לוג בגיליון Log (לא Mini-Log!)
        _logMessage(log, `✅ מתקבל: ${emailData.email} - "${emailData.subject}"`, 'SUCCESS');
      } catch (recordError) {
        _logMessage(log, `❌ שגיאה ביצירת רשומה: ${emailData.email} - ${recordError.message}`, 'ERROR');
        skipped++;
      }
      
    } catch (error) {
      _logMessage(log, `❌ שגיאה בעיבוד מייל ${emailData.email}: ${error.message}`, 'ERROR');
      skipped++;
    }
  }
  
  // ⚡ כתיבה אחת של כל הרשומות!
  if (allRecords.length > 0) {
    sheet.getRange(startingRow, 1, allRecords.length, allRecords[0].length)
         .setValues(allRecords);
    sheet.getRange(startingRow, 1, allRecords.length, 1)
         .insertCheckboxes();
    _logMessage(log, `✅ נכתבו ${allRecords.length} רשומות בבת אחת`, 'SUCCESS');
  }
  
  // ✅ Flush אחד בסוף
  SpreadsheetApp.flush();
  
  // ✅ עדכון Mini-Log - סיכום בלבד (F9)
  statusSheet.getRange("F9").setValue(
    `תוצאות: ${inserted} נוספו, ${skipped} דולגו, ${duplicates} כפילויות`
  );
  
  _logMessage(log, `📊 תוצאות עיבוד: ${inserted} נוספו, ${skipped} דולגו, ${duplicates} כפילויות`, 'SUCCESS');
  
  return { inserted, skipped, duplicates, total: emails.length };
}

// === EXISTING RECORDS GETTER ===
function getExistingRecords(sheet) {
  const existingRecords = new Map();
  
  try {
    // מציאת השורה האחרונה עם נתונים אמיתיים (לא רק צ'קבוקסים)
    const lastRowWithData = findLastRowWithData(sheet);
    
    if (lastRowWithData > 1) {
      const data = sheet.getRange(2, 1, lastRowWithData - 1, sheet.getLastColumn()).getValues();
      
      for (const row of data) {
        const messageId = row[7]; // Email ID column
        if (messageId) {
          existingRecords.set(messageId, true);
        }
      }
    }
  } catch (error) {
    console.error(`Error getting existing records: ${error.message}`);
  }
  
  return existingRecords;
}

// ✅ V32: פונקציה למציאת השורה האחרונה עם נתונים אמיתיים
function findLastRowWithData(sheet) {
  const maxRows = sheet.getLastRow();
  
  // בדיקה מלמטה למעלה לשורה עם נתונים בעמודת Email ID (עמודה H)
  for (let row = maxRows; row >= 14; row--) {  // ✅ מתחיל מ-14 (אחרי כותרות בשורה 13)
    const emailId = sheet.getRange(row, 8).getValue(); // עמודה H = Email ID
    if (emailId && emailId.toString().trim() !== '') {
      return row;
    }
  }
  
  return 13; // ✅ אם אין נתונים, החזר 13 (הכותרות)
}

// פונקציה לניקוי צ'קבוקסים ישנים
function cleanupOldCheckboxes(sheet, log) {
  const lastRowWithData = findLastRowWithData(sheet);
  const maxRows = sheet.getLastRow();
  
  _logMessage(log, `🔍 בדיקת ניקוי: נתונים עד שורה ${lastRowWithData}, גיליון עד שורה ${maxRows}`, 'INFO');
  
  // אם יש צ'קבוקסים מיותרים (מעבר לנתונים האמיתיים)
  if (maxRows > lastRowWithData + 5) { // השארת מרווח של 5 שורות בלבד
    _logMessage(log, `🧹 מנקה צ'קבוקסים ישנים משורה ${lastRowWithData + 1} עד ${maxRows}`, 'INFO');
    
    try {
      // הסרת צ'קבוקסים מיותרים
      const rangeToClean = sheet.getRange(lastRowWithData + 1, 1, maxRows - lastRowWithData, 1);
      rangeToClean.removeCheckboxes();
      
      // ניקוי תוכן השורות המיותרות
      rangeToClean.clearContent();
      
      _logMessage(log, `✅ צ'קבוקסים ישנים נוקו בהצלחה`, 'SUCCESS');
    } catch (e) {
      _logMessage(log, `⚠️ שגיאה בניקוי צ'קבוקסים: ${e.message}`, 'WARNING');
    }
  } else {
    _logMessage(log, `✅ אין צורך בניקוי - הגיליון נקי`, 'SUCCESS');
  }
}

// === EMAIL FILTER WITH DETAILED REASONS ===
function getSkipReason(emailData, config) {
  const email = emailData.email.toLowerCase();
  const subject = emailData.subject.toLowerCase();
  
  // בדיקה 1: האם זה שולח מאושר? אם כן - אישור מיידי!
  if (config.approvedSenders && config.approvedSenders.length > 0) {
    const isApproved = config.approvedSenders.some(approved => 
      email.includes(approved.toLowerCase())
    );
    if (isApproved) {
      console.log(`✅ שולח מאושר: ${email}`);
      return null; // מאושר - לא דולג
    }
  }
  
  // בדיקה 2: מיילים מוחרגים מפורשים
  const excludedEmail = config.excludedEmails.find(excluded => 
    email.includes(excluded.toLowerCase())
  );
  if (excludedEmail) {
    return `מייל מוחרג: ${excludedEmail}`;
  }
  
  // בדיקה 3: מילים מוחרגות בנושא
  const excludedWord = config.excludedKeywords.find(word => 
    subject.includes(word.toLowerCase())
  );
  if (excludedWord) {
    return `מילה מוחרגת: "${excludedWord}"`;
  }
  
  // בדיקה 4: הזמנות קלנדר
  if (subject.includes('invitation:') || subject.includes('invite') || 
      subject.includes('meeting') || subject.includes('calendar')) {
    return 'הזמנת קלנדר';
  }
  
  // בדיקה 5: האם יש מילת מפתח בנושא?
  const allKeywords = [...config.enKeywords, ...config.heKeywords].filter(k => k && k.trim());
  const hasKeyword = allKeywords.some(keyword => 
    subject.includes(keyword.toLowerCase())
  );
  
  if (!hasKeyword) {
    return 'אין מילת מפתח בנושא';
  }
  
  return null; // עבר את כל הבדיקות - מאושר
}

// === EMAIL FILTER (LEGACY COMPATIBILITY) ===
function shouldSkipEmail(emailData, config) {
  return getSkipReason(emailData, config) !== null;
}

// === PRIVATE EMAIL CHECKER ===
function isPrivateEmail(email) {
  const privateDomains = [
    "gmail.com", "walla.co.il", "yahoo.com", "hotmail.com", "outlook.com", 
    "aol.com", "mail.ru", "yandex.ru", "protonmail.com", "icloud.com",
    "zoho.com", "gmx.com", "live.com", "msn.com", "me.com"
  ];
  
  return privateDomains.some(domain => email.toLowerCase().includes(domain));
}

// === INVOICE LINK CHECKER ===
function hasInvoiceLink(emailData) {
  const htmlLink = extractInvoiceLinkFromHtml(emailData.htmlBody);
  const textLink = extractInvoiceLinkFromPlainText(emailData.body);
  
  return htmlLink || textLink;
}

// ⚡ V32B: פונקציה ליצירת Gmail Message Link
function buildGmailMessageLink(messageId) {
  return `https://mail.google.com/mail/u/0/#all/${messageId}`;
}

// ⚡ V32C: פונקציה ליצירת Gmail Attachment Link
function buildGmailAttachmentLink(messageId, attachmentIndex) {
  const attId = `0.${attachmentIndex + 1}`;
  return `https://mail.google.com/mail/u/0/?ui=2&view=att&th=${messageId}&attid=${attId}&disp=safe&zw`;
}

// === LOCAL EMAIL CHECKER ===
// ✅ V32 PERFORMANCE: לוגיקה פשוטה ומהירה (בלי PropertiesService!)
function isLocalEmail(email, subject, body) {
  const localDomains = [".co.il", ".org.il", ".gov.il", ".muni.il", ".ac.il"];
  const isLocalDomain = localDomains.some(domain => email.toLowerCase().endsWith(domain));
  
  // בדיקת תווים עבריים
  const hasHebrew = /[\u0590-\u05FF]/.test(subject) || /[\u0590-\u05FF]/.test(body);
  
  return isLocalDomain || hasHebrew;
}

// === RECORD CREATOR ===
function createInvoiceRecord(emailData, config) {
  // ⚡ V32: סריקה מהירה - רק נתוני מייל, ללא Drive!
  const isLocal = isLocalEmail(emailData.email, emailData.subject, emailData.body);
  const category = isLocal ? "🔷 Local" : "🌐 International";

  let attachmentName = "";
  let mailLink = "";
  
  // ⚡ V32B: יצירת Gmail Message Link (לפתיחת המייל)
  if (emailData.messageId) {
    mailLink = buildGmailMessageLink(emailData.messageId);
  }

  // שמירת שם הקובץ המצורף (אם יש)
  if (emailData.pdfAttachment) {
    attachmentName = emailData.pdfAttachment.getName();
  }

  // ⚡ V32B: PDF Link נשאר ריק - לא יוצרים Drive בזמן הסריקה!
  // ⚡ V32C: Download Link נשאר ריק - יתמלא רק ב-"הכן קבצים להורדה"
  
  return [
    false,                                    // Action (checkbox)
    category,                                 // Category
    emailData.senderName || "",               // Sender Name
    emailData.email || "",                    // Sender Email
    formatDateInput(emailData.date),          // Date
    emailData.subject || "",                  // Subject
    "",                                       // PDF Link (ריק - V32B)
    emailData.messageId || "",                // Email ID
    attachmentName,                           // Attachment Name
    "",                                       // Sum (Local) - ריק
    "",                                       // Sum (Int'l) - ריק
    "",                                       // Currency - ריק
    mailLink,                                 // Mail Link (V32B)
    ""                                        // Download Link (ריק - V32C)
  ];
}

// === ROW FORMATTER ===
function formatInvoiceRow(sheet, rowNum, category) {
  // DISABLED FOR DEBUGGING - זמנית מבוטל לבדיקה
  return;
  
  try {
    const range = sheet.getRange(rowNum, 1, 1, 12);
    
    // צבע לפי קטגוריה
    if (category === "🔷 Local") {
      range.setBackground("#e8f5e8"); // ירוק בהיר
    } else {
      range.setBackground("#e8f0ff"); // כחול בהיר
    }
    
    // גבולות
    range.setBorder(true, true, true, true, false, false, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
    
  } catch (error) {
    console.error(`Error formatting row: ${error.message}`);
  }
}

// === SCAN SUMMARY ===
function showScanSummary(results, log) {
  const message = `✅ הסריקה הושלמה!

📊 תוצאות:
• ${results.inserted} חשבוניות חדשות נוספו
• ${results.skipped} מיילים דולגו
• ${results.duplicates} כפילויות נמצאו
• ${results.total} מיילים נבדקו בסך הכל

${results.inserted > 0 ? '🎉 החשבוניות זמינות בגיליון החדש!' : '⚠️ לא נמצאו חשבוניות חדשות.'}

💡 עכשיו אפשר:
1. לצפייה/הורדה ישירה → לחץ על Download Link
2. לסנכרון ל-Drive → תפריט "פעולות" → "הכן קבצים להורדה"`;

  SpreadsheetApp.getUi().alert("סיכום סריקה", message, SpreadsheetApp.getUi().ButtonSet.OK);
  
  // ✅ V32: אין טוסטים! כל ההתקדמות ב-Mini-Log
  
  _logMessage(log, message.replace(/\n/g, ' '), 'SUCCESS');
}

// === DATE FORMATTERS ===
function formatDateShort(date) {
  if (!date || !(date instanceof Date)) return "";
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${dd}.${mm}.${String(yyyy).slice(2)}`;
}

// === ACTIONS MENU FUNCTIONS ===
function runActions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();

  // Check if the active sheet is an 'invoices' sheet
  if (!/^invoices /i.test(sheetName)) {
    SpreadsheetApp.getUi().alert("❌ אנא עבור לגיליון 'invoices' המתאים לפני הפעלת פעולות.");
    return;
  }

  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  if (!settings || !log) {
    SpreadsheetApp.getUi().alert("❌ גיליונות Settings או Log חסרים. אנא הפעל את 'Setup Sheets' תחילה.");
    return;
  }

  const dataRange = activeSheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  const actionColumnIndex = headers.indexOf("Action");
  const emailColumnIndex = headers.indexOf("Sender Email");
  const categoryColumnIndex = headers.indexOf("Category");

  if (actionColumnIndex === -1 || emailColumnIndex === -1) {
    SpreadsheetApp.getUi().alert("❌ לא נמצאו עמודות 'Action' או 'Sender Email' בגיליון הנוכחי.");
    return;
  }

  // Get all checked rows
  const checkedRows = [];
  const actionRange = activeSheet.getRange(2, actionColumnIndex + 1, data.length - 1, 1);
  const checkboxValues = actionRange.getValues();
  for (let i = 0; i < checkboxValues.length; i++) {
    if (checkboxValues[i][0] === true) {
      checkedRows.push({
        rowIndex: i + 2, // +2 because we start from row 2 and i is 0-based
        email: data[i + 1][emailColumnIndex],
        senderName: data[i + 1][headers.indexOf("Sender Name")],
        emailId: data[i + 1][headers.indexOf("Email ID")],
        subject: data[i + 1][headers.indexOf("Subject")],
      });
    }
  }

  if (checkedRows.length === 0) {
    SpreadsheetApp.getUi().alert("אין שורות מסומנות לביצוע פעולות.");
    return;
  }

  // בדיקת סטטוס מוחרג/מאושר
  const excludedList = settings.getRange("F2:F").getValues().flat().filter(Boolean);
  const approvedList = settings.getRange("G2:G").getValues().flat().filter(Boolean);
  checkedRows.forEach(row => {
    row.isExcluded = excludedList.includes(row.email);
    row.isApproved = approvedList.includes(row.email);
  });

  // Create and show the actions dialog
  const html = HtmlService.createHtmlOutput(`
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; direction: rtl; }
        h3 { margin-top: 0; color: #4285f4; }
        .option { margin: 15px 0; padding: 12px; border: 1px solid #ddd; border-radius: 5px; cursor: pointer; transition: background-color 0.2s; }
        .option:hover { background-color: #f8f9fa; }
        .option h4 { margin-top: 0; }
        .option p { color: #666; margin-bottom: 0; font-size: 0.9em; }
        .status { font-size: 0.85em; color: #b00; margin-right: 10px; }
        .approved { color: #0a0; }
        .buttons { margin-top: 20px; text-align: center; }
        button { padding: 8px 16px; margin: 0 5px; border: none; border-radius: 4px; cursor: pointer; }
        .primary { background-color: #4285f4; color: white; }
        .secondary { background-color: #f1f3f4; color: #202124; }
      </style>
    </head>
    <body>
      <h3>פעולות עבור ${checkedRows.length} שורות מסומנות</h3>
      <ul>
        ${checkedRows.map(row => `<li>${row.email} <span class="status">${row.isExcluded ? 'מוחרג' : row.isApproved ? '<span class="approved">מאושר</span>' : ''}</span></li>`).join('')}
      </ul>
      <div class="option" onclick="selectAction('exclude')">
        <h4>החרג חשבוניות מסומנות</h4>
        <p>הוסף את המיילים המסומנים לרשימת ההחרגות ומחק את כל הרשומות עם אותם מיילים (כולל כפילויות).</p>
      </div>
      <div class="option" onclick="selectAction('approve')">
        <h4>הוסף למיילים מאושרים</h4>
        <p>הוסף את המיילים המסומנים לרשימת המאושרים (whitelist) – ייכנסו אוטומטית בעתיד.</p>
      </div>
      <div class="option" onclick="selectAction('local')">
        <h4>סמן כמקומיות</h4>
        <p>סמן את החשבוניות המסומנות כ-🔷 Local ושמור את ההגדרה לסריקות הבאות</p>
      </div>
      <div class="option" onclick="selectAction('international')">
        <h4>סמן כבינלאומיות</h4>
        <p>סמן את החשבוניות המסומנות כ-🌐 International ושמור את ההגדרה לסריקות הבאות</p>
      </div>
      <div class="option" onclick="selectAction('export')">
        <h4>ייצא חשבוניות מסומנות</h4>
        <p>ייצא כקובץ ZIP רק את החשבוניות המסומנות</p>
      </div>
      <div class="buttons">
        <button class="secondary" onclick="google.script.host.close()">ביטול</button>
      </div>
      <script>
        function selectAction(action) {
          google.script.run
            .withSuccessHandler(function(result) {
              if (result && result.message) {
                alert(result.message);
              }
              google.script.host.close();
            })
            .processActionSelection('${activeSheet.getName()}', action, ${JSON.stringify(checkedRows)});
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(420)
  .setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, "בחר פעולה");
}

// === ACTION PROCESSOR ===
function processActionSelection(sheetName, action, checkedRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet || !settings) {
    return { success: false, message: "❌ שגיאה בגישה לגיליונות." };
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const emailColumnIndex = headers.indexOf("Sender Email") + 1;
  const categoryColumnIndex = headers.indexOf("Category") + 1;

  switch (action) {
    case 'exclude': {
      // איסוף כל השורות עם אותם מיילים (כפולים)
      const allData = sheet.getDataRange().getValues();
      const emailsToExclude = checkedRows.map(r => r.email);
      const rowsToDelete = [];
      for (let i = 1; i < allData.length; i++) {
        if (emailsToExclude.includes(allData[i][emailColumnIndex - 1])) {
          rowsToDelete.push(i + 1); // +1 כי יש כותרת
        }
      }
      // בדיקה אם יש כבר מוחרגים
      const excludeRange = settings.getRange("F2:F");
      const excludeValues = excludeRange.getValues().flat().filter(Boolean);
      let newlyExcluded = 0;
      emailsToExclude.forEach(email => {
        if (!excludeValues.includes(email)) {
          settings.getRange(excludeValues.length + 2 + newlyExcluded, 6).setValue(email);
          newlyExcluded++;
        }
      });
      // מחיקת כל השורות הרלוונטיות (מלמטה למעלה)
      rowsToDelete.sort((a, b) => b - a).forEach(rowIdx => sheet.deleteRow(rowIdx));
      _logMessage(log, `הוחרגו ומחקו ${rowsToDelete.length} שורות (${emailsToExclude.length} מיילים)`, 'SUCCESS');
      return { success: true, message: `✅ ${rowsToDelete.length} שורות הוחרגו ונמחקו. (${emailsToExclude.length} מיילים הוספו להחרגות)` };
    }
    case 'approve': {
      // הוספה ל-whitelist
      const approvedRange = settings.getRange("G2:G");
      const approvedValues = approvedRange.getValues().flat().filter(Boolean);
      let newlyApproved = 0;
      checkedRows.forEach(row => {
        if (!approvedValues.includes(row.email)) {
          settings.getRange(approvedValues.length + 2 + newlyApproved, 7).setValue(row.email);
          newlyApproved++;
        }
      });
      _logMessage(log, `הוספו ${newlyApproved} מיילים ל-whitelist`, 'SUCCESS');
      return { success: true, message: `✅ ${newlyApproved} מיילים הוספו לרשימת המאושרים.` };
    }
    case 'local':
    case 'international': {
      const category = action === 'local' ? "🔷 Local" : "🌐 International";
      for (const row of checkedRows) {
        const rowRange = sheet.getRange(row.rowIndex, categoryColumnIndex);
        rowRange.setValue(category);
      }
      return { 
        success: true, 
        message: `✅ ${checkedRows.length} חשבוניות סומנו כ-${category}.` 
      };
    }
    case 'export':
      return exportMarkedInvoicesAsZip(sheetName, checkedRows.map(row => row.rowIndex));
    default:
      return { success: false, message: "❌ פעולה לא מוכרת." };
  }
}

// === HELPER FUNCTIONS FOR ACTIONS ===
function getSelectedRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const actionColumnIndex = headers.indexOf("Action");
  
  // חיפוש גמיש של עמודת המייל
  let emailColumnIndex = headers.indexOf("Email Sender");
  if (emailColumnIndex === -1) {
    emailColumnIndex = headers.indexOf("Sender Email"); // fallback לגיליון הישן
  }
  
  if (actionColumnIndex === -1 || emailColumnIndex === -1) {
    SpreadsheetApp.getUi().alert(`❌ לא נמצאו עמודות נדרשות.\nעמודות זמינות: ${headers.join(', ')}`);
    return [];
  }
  
  const selectedRows = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][actionColumnIndex] === true) {
      selectedRows.push({
        rowIndex: i + 1, // +1 for 1-based indexing
        email: data[i][emailColumnIndex],
        senderName: data[i][headers.indexOf("Company Name")] || data[i][headers.indexOf("Sender Name")] || "",
        emailId: data[i][headers.indexOf("Email ID")] || "",
        subject: data[i][headers.indexOf("Subject")] || ""
      });
    }
  }
  
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("אין שורות מסומנות לביצוע פעולות.");
  }
  
  return selectedRows;
}

function excludeSelected() {
  const selectedRows = getSelectedRows();
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("אין שורות מסומנות.");
    return;
  }
  runActions();
}

function markAsLocal() {
  const selectedRows = getSelectedRows();
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("אין שורות מסומנות.");
    return;
  }
  runActions();
}

function markAsInternational() {
  const selectedRows = getSelectedRows();
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("אין שורות מסומנות.");
    return;
  }
  runActions();
}

function exportSelected() {
  const selectedRows = getSelectedRows();
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("אין שורות מסומנות.");
    return;
  }
  runActions();
}

function deleteSelected() {
  const selectedRows = getSelectedRows();
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("אין שורות מסומנות.");
    return;
  }
  runActions();
}

function approveSelected() {
  const selectedRows = getSelectedRows();
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("אין שורות מסומנות.");
    return;
  }
  runActions();
}

// === EXPORT FUNCTION PLACEHOLDER ===
function exportMarkedInvoicesAsZip(sheetName, rowIndices) {
  // This will be implemented by the zipExporter.gs
  SpreadsheetApp.getUi().alert("פונקציית ייצוא תמומש בהמשך");
  return { success: true, message: "ייצוא יבוצע בהמשך" };
}

// === DIAGNOSTIC FUNCTIONS ===
function diagnosticScan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!settings || !log) {
    SpreadsheetApp.getUi().alert("❌ גיליונות Settings או Log חסרים.");
    return;
  }
  
  _logMessage(log, "🔍 התחלת אבחון מתקדם", 'INFO');
  
  // בדיקת הרשאות Drive
  const driveCheck = checkDrivePermissions();
  _logMessage(log, `🔐 הרשאות Drive: ${driveCheck.success ? '✅ תקינות' : '❌ בעיה - ' + driveCheck.message}`, driveCheck.success ? 'SUCCESS' : 'ERROR');

  // בדיקת הרשאות כתיבה לגיליון
  const sheetPermCheck = checkSpreadsheetWritePermissions();
  _logMessage(log, `📝 הרשאות כתיבה לגיליון: ${sheetPermCheck.success ? '✅ תקינות' : '❌ בעיה - ' + sheetPermCheck.message}`, sheetPermCheck.success ? 'SUCCESS' : 'ERROR');
  
  // בדיקת הגדרות
  const startStr = settings.getRange("A2").getValue();
  const endStr = settings.getRange("B2").getValue();
  const enKeywords = settings.getRange("C2:C").getValues().flat().filter(Boolean);
  const heKeywords = settings.getRange("D2:D").getValues().flat().filter(Boolean);
  const excludedKeywords = settings.getRange("E2:E").getValues().flat().filter(Boolean);
  const excludedEmails = settings.getRange("F2:F").getValues().flat().filter(Boolean);
  const approved = settings.getRange("G2:G").getValues().flat().filter(Boolean);
  
  _logMessage(log, `📅 טווח תאריכים: ${startStr} עד ${endStr}`, 'INFO');
  _logMessage(log, `🔤 מילות מפתח אנגלית (${enKeywords.length}): ${enKeywords.join(', ')}`, 'INFO');
  _logMessage(log, `🔤 מילות מפתח עברית (${heKeywords.length}): ${heKeywords.join(', ')}`, 'INFO');
  _logMessage(log, `🚫 מילים מוחרגות (${excludedKeywords.length}): ${excludedKeywords.join(', ')}`, 'INFO');
  _logMessage(log, `🚫 מיילים מוחרגים (${excludedEmails.length}): ${excludedEmails.join(', ')}`, 'INFO');
  _logMessage(log, `✅ שולחים מאושרים (${approved.length}): ${approved.join(', ')}`, 'INFO');
  
  if (!startStr || !endStr) {
    _logMessage(log, "❌ תאריכים לא מוגדרים", 'ERROR');
    SpreadsheetApp.getUi().alert("❌ יש להגדיר תאריכים בגיליון Settings לפני הסריקה");
    return;
  }
  
  const startDate = new Date(startStr);
  const endDate = new Date(endStr);
  
  // בדיקת תקינות התאריכים
  if (startDate.getTime() >= endDate.getTime()) {
    _logMessage(log, "❌ תאריך התחלה חייב להיות לפני תאריך הסיום", 'ERROR');
    SpreadsheetApp.getUi().alert("❌ תאריך התחלה חייב להיות לפני תאריך הסיום");
    return;
  }
  
  // בדיקת שאילתה Gmail
  let query = `after:${formatDateGmail(startDate)} before:${formatDateGmail(endDate)} has:attachment`;
  _logMessage(log, `🔍 שאילתה בסיסית: ${query}`, 'INFO');
  
  // הוספת מילות מפתח לשאילתה
  const allKeywords = [...enKeywords, ...heKeywords].filter(k => k && k.trim());
  if (allKeywords.length > 0) {
    const keywordQuery = allKeywords.map(k => `"${k.trim()}"`).join(" OR ");
    query += ` (${keywordQuery})`;
    _logMessage(log, `🔍 שאילתה עם מילות מפתח: ${query}`, 'INFO');
  }
  
  // הוספת שולחים מאושרים
  if (approved.length > 0) {
    const approvedQuery = approved.map(sender => `from:${sender.trim()}`).join(" OR ");
    if (allKeywords.length > 0) {
      query += ` OR (${approvedQuery})`;
    } else {
      query += ` (${approvedQuery})`;
    }
    _logMessage(log, `🔍 שאילתה מלאה: ${query}`, 'INFO');
  }
  
  try {
    const threads = GmailApp.search(query, 0, 20); // 20 ראשונים לבדיקה מפורטת
    _logMessage(log, `📧 נמצאו ${threads.length} מיילים עם קבצים מצורפים בטווח`, 'SUCCESS');
    
    if (threads.length === 0) {
      _logMessage(log, "⚠️ לא נמצאו מיילים בטווח התאריכים. בדוק:", 'WARNING');
      _logMessage(log, "   1. טווח התאריכים נכון", 'WARNING');
      _logMessage(log, "   2. יש מיילים עם קבצים מצורפים בטווח זה", 'WARNING');
      _logMessage(log, "   3. מילות המפתח מתאימות", 'WARNING');
      SpreadsheetApp.getUi().alert("⚠️ לא נמצאו מיילים בטווח התאריכים.\n\nבדוק:\n1. טווח התאריכים נכון\n2. יש מיילים עם קבצים מצורפים בטווח זה\n3. מילות המפתח מתאימות");
      return;
    }
    
    let validInvoices = 0;
    let skippedCount = 0;
    
    // בדיקת המיילים הראשונים
    for (let i = 0; i < Math.min(10, threads.length); i++) {
      const thread = threads[i];
      const messages = thread.getMessages();
      const msg = messages[messages.length - 1]; // המייל האחרון בשרשור
      
      const email = _parseSender(msg.getFrom());
      const subject = msg.getSubject();
      const body = msg.getPlainBody();
      const date = msg.getDate();
      
      _logMessage(log, `📧 מייל ${i+1}: מאת ${email}, נושא: "${subject.substring(0, 50)}..."`, 'INFO');
      _logMessage(log, `📅 תאריך: ${formatDateInput(date)}`, 'INFO');
      
      // בדיקת סינונים
      const isPrivate = isPrivateEmail(email);
      const isExcluded = excludedEmails.map(e => e.toLowerCase()).includes(email.toLowerCase());
      const foundKeyword = [...enKeywords, ...heKeywords].some(k => subject.toLowerCase().includes(k.toLowerCase()));
      const foundExcludedKeyword = excludedKeywords.some(k => subject.toLowerCase().includes(k.toLowerCase()));
      const hasPdf = hasPdfAttachment(msg);
      const hasLink = extractInvoiceLinkFromHtml(msg.getBody()) || extractInvoiceLinkFromPlainText(body);
      
      _logMessage(log, `   🔍 דומיין פרטי: ${isPrivate ? 'כן' : 'לא'}`, isPrivate ? 'WARNING' : 'SUCCESS');
      _logMessage(log, `   🚫 מוחרג: ${isExcluded ? 'כן' : 'לא'}`, isExcluded ? 'WARNING' : 'SUCCESS');
      _logMessage(log, `   🔤 מילות מפתח: ${foundKeyword ? 'כן' : 'לא'}`, foundKeyword ? 'SUCCESS' : 'WARNING');
      _logMessage(log, `   ✅ שולח מאושר: ${approved.includes(email) ? 'כן' : 'לא'}`, approved.includes(email) ? 'SUCCESS' : 'INFO');
      _logMessage(log, `   📎 קובץ PDF: ${hasPdf ? 'כן' : 'לא'}`, hasPdf ? 'SUCCESS' : 'INFO');
      _logMessage(log, `   🔗 קישור בגוף: ${hasLink ? 'כן' : 'לא'}`, hasLink ? 'SUCCESS' : 'INFO');
      
      const shouldInclude = !isPrivate && !isExcluded && foundKeyword && (hasPdf || hasLink);
      _logMessage(log, `   ✅ יכלל בתוצאות: ${shouldInclude ? 'כן' : 'לא'}`, shouldInclude ? 'SUCCESS' : 'WARNING');
      
      if (shouldInclude) {
        validInvoices++;
      } else {
        skippedCount++;
        let skipReason = [];
        if (isPrivate) skipReason.push("דומיין פרטי");
        if (isExcluded) skipReason.push("מוחרג");
        if (!foundKeyword) skipReason.push("אין מילות מפתח");
        if (!hasPdf && !hasLink) skipReason.push("אין PDF או קישור");
        _logMessage(log, `   ❌ סיבת דחייה: ${skipReason.join(', ')}`, 'WARNING');
      }
    }
    
    _logMessage(log, `📊 סיכום אבחון: ${validInvoices} חשבוניות תקינות, ${skippedCount} נדחו`, 'INFO');
    
    if (validInvoices === 0) {
      _logMessage(log, "⚠️ לא נמצאו חשבוניות תקינות! בדוק הגדרות", 'WARNING');
      SpreadsheetApp.getUi().alert("⚠️ לא נמצאו חשבוניות תקינות!\n\nבדוק:\n1. מילות המפתח מתאימות\n2. רשימת השולחים המאושרים\n3. רשימת ההחרגות");
    } else {
      _logMessage(log, `✅ נמצאו ${validInvoices} חשבוניות תקינות מתוך ${Math.min(10, threads.length)} שנבדקו`, 'SUCCESS');
    }
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בחיפוש Gmail: ${error.message}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ שגיאה בחיפוש Gmail:\n${error.message}`);
  }
  
  _logMessage(log, "🔍 אבחון הושלם", 'SUCCESS');
  SpreadsheetApp.getUi().alert("אבחון הושלם. בדוק את גיליון Log לפרטים מלאים.");
}

function checkSpreadsheetWritePermissions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let testSheet = ss.getSheetByName('WriteTestSheet');
    if (!testSheet) testSheet = ss.insertSheet('WriteTestSheet');
    const testValue = 'permission_test_' + new Date().getTime();
    testSheet.getRange(1, 1).setValue(testValue);
    const readBack = testSheet.getRange(1, 1).getValue();
    // ניקוי
    testSheet.clear();
    ss.deleteSheet(testSheet);
    if (readBack === testValue) {
      return { success: true, message: '✅ הרשאות כתיבה לגיליון תקינות' };
    } else {
      return { success: false, message: '❌ בעיית הרשאות כתיבה: לא ניתן לאמת כתיבה/קריאה' };
    }
  } catch (e) {
    return { success: false, message: `❌ בעיית הרשאות כתיבה: ${e.message}` };
  }
}

// ============================================================================
// ⚡ V32C: "הכן קבצים להורדה" - יצירת Download Links לשורות מסומנות
// ============================================================================

/**
 * ⚡ V32C FIX: פונקציה ראשית - הכן קבצים להורדה
 * עובדת לפי בחירת שורות פיזית (selection) או על כל השורות
 * ✅ V32 Phase 1: Batch Write מלא + gmailCache + דילוג על שורות 1-13
 */
function prepareDownloadLinksForSelected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  // בדיקה שזה גיליון invoices
  if (!/^invoices /i.test(activeSheet.getName())) {
    SpreadsheetApp.getUi().alert("❌ אנא עבור לגיליון 'invoices' המתאים.");
    return;
  }
  
  _logMessage(log, "📦 התחלת הכנת לינקי הורדה", 'INFO');
  
  try {
    // קריאת כל הנתונים (Batch Read)
    const dataRange = activeSheet.getDataRange();
    const data = dataRange.getValues();
    const headers = data[12];  // ✅ שורה 13 = אינדקס 12
    
    // מציאת אינדקסים של עמודות
    const emailIdColIdx = headers.indexOf("Email ID");
    const downloadLinkColIdx = headers.indexOf("Download Link");
    
    if (emailIdColIdx === -1 || downloadLinkColIdx === -1) {
      SpreadsheetApp.getUi().alert("❌ לא נמצאו עמודות נדרשות (Email ID, Download Link)");
      return;
    }
    
    // ⚡ V32C FIX: בדיקת selection פיזית
    const selection = activeSheet.getSelection();
    const activeRange = selection.getActiveRange();
    let selectedRows = [];
    
    if (activeRange && activeRange.getNumRows() > 1) {
      // יש selection - עבוד רק על השורות המסומנות
      const startRow = activeRange.getRow();
      const numRows = activeRange.getNumRows();
      
      _logMessage(log, `📦 נבחרו ${numRows} שורות (${startRow} עד ${startRow + numRows - 1})`, 'INFO');
      
      for (let i = startRow; i < startRow + numRows; i++) {
        if (i > 13 && i <= data.length && data[i - 1][emailIdColIdx]) {  // ✅ דילוג על שורות 1-13
          selectedRows.push({
            dataIndex: i - 1,  // אינדקס במערך data
            messageId: data[i - 1][emailIdColIdx]
          });
        }
      }
    } else {
      // אין selection - עבוד על כל השורות עם נתונים
      _logMessage(log, `📦 אין selection - מכין לינקים לכל השורות`, 'INFO');
      
      for (let i = 13; i < data.length; i++) {  // ✅ מתחיל מ-13 (שורה 14 = אינדקס 13)
        if (data[i][emailIdColIdx]) {
          selectedRows.push({
            dataIndex: i,
            messageId: data[i][emailIdColIdx]
          });
        }
      }
    }
    
    if (selectedRows.length === 0) {
      SpreadsheetApp.getUi().alert("אין שורות לעיבוד.\n\nסמן שורות פיזית או השאר ריק לעיבוד כל השורות.");
      return;
    }
    
    _logMessage(log, `📦 מכין לינקי הורדה ל-${selectedRows.length} חשבוניות`, 'INFO');
    
    // ⚡ V32 Phase 1: gmailCache למניעת קריאות כפולות
    const gmailCache = {};
    
    // עיבוד כל שורה מסומנת
    let preparedCount = 0;
    let skippedCount = 0;
    
    for (let i = 0; i < selectedRows.length; i++) {
      const row = selectedRows[i];
      
      try {
        if (!row.messageId) {
          skippedCount++;
          continue;
        }
        
        // ✅ קבלת המייל מ-Cache או מ-Gmail
        let message = gmailCache[row.messageId];
        if (!message) {
          message = GmailApp.getMessageById(row.messageId);
          if (message) {
            gmailCache[row.messageId] = message;
          }
        }
        
        if (!message) {
          _logMessage(log, `⚠️ לא נמצא מייל עם ID: ${row.messageId}`, 'WARNING');
          skippedCount++;
          continue;
        }
        
        // חיפוש PDF בקבצים המצורפים
        const attachments = message.getAttachments();
        const pdfAttachmentIndex = attachments.findIndex(att => 
          att.getName().toLowerCase().endsWith('.pdf') || 
          att.getContentType().toLowerCase().includes('pdf')
        );
        
        if (pdfAttachmentIndex === -1) {
          _logMessage(log, `⚠️ לא נמצא PDF במייל: ${row.messageId}`, 'WARNING');
          skippedCount++;
          continue;
        }
        
        // יצירת Download Link
        const downloadLink = buildGmailAttachmentLink(row.messageId, pdfAttachmentIndex);
        
        // ✅ עדכון ישיר ב-array (לא push לרשימה נפרדת)
        data[row.dataIndex][downloadLinkColIdx] = downloadLink;
        
        preparedCount++;
        
        // עדכון Progress ב-Mini-Log
        if ((i + 1) % 5 === 0) {
          const progress = Math.round(((i + 1) / selectedRows.length) * 100);
          activeSheet.getRange("G7").setValue(`📦 ${i + 1} מתוך ${selectedRows.length} (${progress}%)`);
        }
        
      } catch (error) {
        _logMessage(log, `❌ שגיאה בעיבוד שורה ${row.dataIndex + 1}: ${error.message}`, 'ERROR');
        skippedCount++;
      }
    }
    
    // ✅ V32 Phase 1: Batch Write אמיתי - כתיבה אחת של כל ה-data
    if (preparedCount > 0) {
      activeSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      SpreadsheetApp.flush();
    }
    
    // עדכון Mini-Log - סיכום
    activeSheet.getRange("G6").setValue("✅ הכנת לינקים הושלמה!");
    activeSheet.getRange("G7").setValue("");
    activeSheet.getRange("G9").setValue(
      `📦 ${preparedCount} חשבוניות קיבלו Download Link, ${skippedCount} דולגו`
    );
    
    _logMessage(log, `✅ הכנת לינקים הושלמה: ${preparedCount} הוכנו, ${skippedCount} דולגו`, 'SUCCESS');
    
    // פופאפ סיכום
    SpreadsheetApp.getUi().alert(
      "✅ הכנת לינקי הורדה הסתיימה!\n\n" +
      `• ${preparedCount} חשבוניות קיבלו Download Link\n` +
      `• ${skippedCount} דולגו (אין PDF או שגיאה)\n\n` +
      "💡 עכשיו אפשר:\n" +
      "1. ללחוץ על Download Link לצפייה/הורדה ישירה\n" +
      "2. להשתמש ב-'יצירת ZIP' או 'סנכרון לDrive' (בעתיד)"
    );
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בהכנת לינקי הורדה: ${error.message}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ שגיאה: ${error.message}`);
  }
}
