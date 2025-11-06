// === GmailProcessor.gs - V31 BACKUP ===
// גיבוי נוצר: 06/11/2025
// לפני שינויי V32A
// כל הפונקציות עם סיומת _BACKUP למניעת כפילויות

/**
 * יצירת progress bar ויזואלי לטוסט
 */
function createScanProgressBar_BACKUP(current, total, width = 20) {
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
function formatScanDuration_BACKUP(seconds) {
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
function checkDrivePermissions_BACKUP() {
  try {
    // ניסיון פשוט לגשת ל-Drive
    DriveApp.getRootFolder();
    return { success: true, message: "הרשאות Drive תקינות" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// === MAIN SCAN FUNCTION ===
function processInvoices_BACKUP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!settings) throw new Error('Settings sheet not found.');
  if (!log) throw new Error('Log sheet not found.');
  
  const startTime = new Date();
  _logMessage(log, `🔍 התחלת סריקת חשבוניות - ${startTime.toLocaleString('he-IL')}`, 'INFO');
  
  try {
    // בדיקת הרשאות
    const driveCheck = checkDrivePermissions_BACKUP();
    if (!driveCheck.success) {
      SpreadsheetApp.getUi().alert(
        "❌ בעיית הרשאות Drive\n\n" + driveCheck.message + 
        "\n\nאנא רענן את הדף ואשר הרשאות מחדש"
      );
      return;
    }

    // קריאת הגדרות
    const config = readScanSettings_BACKUP(settings, log);
    if (!config) return;
    
    // הודעת לוג מפורטת על טווח התאריכים
    _logMessage(log, `📅 סורק מיילים מ-${formatDateInput(config.startDate)} עד ${formatDateInput(config.endDate)}`, 'INFO');

    // יצירת/עדכון גיליון תוצאות
    const sheet = createOrUpdateResultsSheet_BACKUP(ss, config.tabName, log);
    if (!sheet) return;

    // התחלת סריקה עם מדידת זמן
    const startTime = new Date();
    SpreadsheetApp.getActiveSpreadsheet().toast("🔍 מתחיל סריקת Gmail...", "סריקה", 5);
    
    _logMessage(log, `⏱️ התחלת סריקה: ${startTime.toLocaleTimeString()}`, 'INFO');
    const scanResults = performGmailScan_BACKUP(config, log, startTime);
    const scanEndTime = new Date();
    const scanDuration = Math.round((scanEndTime - startTime) / 1000);
    _logMessage(log, `⏱️ סיום סריקת Gmail: ${scanEndTime.toLocaleTimeString()} - ${formatScanDuration_BACKUP(scanDuration)}`, 'INFO');
    
    if (scanResults.emails.length === 0) {
      SpreadsheetApp.getUi().alert("לא נמצאו מיילים בטווח התאריכים שהוגדר.\n\nבדוק:\n1. טווח התאריכים\n2. מילות המפתח\n3. שיש מיילים עם קבצים מצורפים");
      _logMessage(log, "❌ לא נמצאו מיילים לעיבוד", 'WARNING');
      return;
    }

    // עיבוד המיילים עם progress bar
    SpreadsheetApp.getActiveSpreadsheet().toast(`📧 נמצאו ${scanResults.emails.length} מיילים - מתחיל עיבוד...`, "עיבוד", 5);
    
    const processStartTime = new Date();
    _logMessage(log, `⏱️ התחלת עיבוד מיילים: ${processStartTime.toLocaleTimeString()}`, 'INFO');
    const processResults = processEmails_BACKUP(scanResults.emails, config, sheet, log, startTime);
    const processEndTime = new Date();
    const processDuration = Math.round((processEndTime - processStartTime) / 1000);
    _logMessage(log, `⏱️ סיום עיבוד מיילים: ${processEndTime.toLocaleTimeString()} - ${formatScanDuration_BACKUP(processDuration)}`, 'INFO');
    
    // סיכום וסיום
    showScanSummary_BACKUP(processResults, log);
    
    _logMessage(log, "✅ סריקה הושלמה בהצלחה", 'SUCCESS');
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בסריקה: ${error.message}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`שגיאה בסריקה: ${error.message}`);
  }
}

// === QUICK API SCAN ===
function quickApiScanInvoices_BACKUP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!settings || !log) {
    SpreadsheetApp.getUi().alert("❌ גיליונות Settings או Log חסרים.");
    return;
  }
  
  _logMessage(log, "⚡ התחלת סריקה מהירה", 'INFO');
  SpreadsheetApp.getActiveSpreadsheet().toast("⚡ מתחיל סריקה מהירה...", "סריקה מהירה", 3);
  
  // קריאה לפונקציה הרגילה - זה יהיה מהיר יותר
  processInvoices_BACKUP();
}

// === SETTINGS READER ===
function readScanSettings_BACKUP(settings, log) {
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
function createOrUpdateResultsSheet_BACKUP(ss, tabName, log) {
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

      // יצירת כותרות
      const headers = [
        "Action", "Category", "Sender Name", "Sender Email", "Date", 
        "Subject", "PDF Link", "Email ID", "Attachment Name", 
        "Sum (Local)", "Sum (Int'l)", "Currency"
      ];
      
      sheet.appendRow(headers);
      
      // עיצוב כותרות
      const headerRange = sheet.getRange("A1:L1");
      headerRange.setFontWeight("bold")
                 .setBackground("#4285f4")
                 .setFontColor("white")
                 .setHorizontalAlignment("center");
      
      // הוספת validation לקטגוריה
      const categoryRange = sheet.getRange("B2:B1000");
      categoryRange.setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['🔷 Local', '🌐 International'])
          .build()
      );
      
      sheet.setFrozenRows(1);
      
      // הסתרת עמודת Action (נשלט דרך התפריט)
      sheet.hideColumns(1, 1); // מסתיר עמודה A
      
      SpreadsheetApp.flush();
      
      _logMessage(log, "✅ גיליון תוצאות נוצר בהצלחה", 'SUCCESS');
    } else {
      _logMessage(log, `📋 משתמש בגיליון קיים: ${tabName}`, 'INFO');
      
      // ניקוי צ'קבוקסים ישנים שגורמים לבעיית שורה 1000
      try {
        cleanupOldCheckboxes_BACKUP(sheet, log);
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
function performGmailScan_BACKUP(config, log, startTime) {
  try {
    // בניית שאילתת חיפוש
    const query = buildGmailQuery_BACKUP(config);
    // לוג השאילתה בשורות נפרדות לקריאות טובה יותר
    _logMessage(log, `שאילתת Gmail:`, 'INFO');
    _logMessage(log, `${query}`, 'INFO');
    
    // חיפוש מיילים
    const threads = GmailApp.search(query, 0, 500); // מקסימום 500 מיילים
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
        const email = extractEmailData_BACKUP(message);
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
function buildGmailQuery_BACKUP(config) {
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
  
  return parts.join(" ");
}

// === EMAIL DATA EXTRACTOR ===
function extractEmailData_BACKUP(message) {
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
function processEmails_BACKUP(emails, config, sheet, log, startTime) {
  let skipped = 0;
  let duplicates = 0;
  
  // קבלת רשומות קיימות
  const existingRecords = getExistingRecords_BACKUP(sheet);
  
  _logMessage(log, `🔄 מתחיל עיבוד ${emails.length} מיילים`, 'INFO');
  
  // ⚡ V32 OPTIMIZATION: Batch Write - בניית מערך במקום כתיבה שורה-שורה
  const allRecords = [];
  const startingRow = findLastRowWithData_BACKUP(sheet) + 1;
  _logMessage(log, `📝 רשומות חדשות יתחילו משורה ${startingRow}`, 'INFO');
  
  // טוסט התחלה
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `⚡ מעבד ${emails.length} מיילים...`, 
    "Phase 1: עיבוד", 3
  );
  
  for (let i = 0; i < emails.length; i++) {
    const emailData = emails[i];
    
    try {
      // בדיקת כפילויות
      if (existingRecords.has(emailData.messageId)) {
        duplicates++;
        continue;
      }
      
      // יצירת רשומה
      try {
        const record = createInvoiceRecord_BACKUP(emailData, config);
        allRecords.push(record);
        existingRecords.set(emailData.messageId, true);
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
  
  // ⚡ בדיקה: אם אין רשומות, לא לכתוב כלום
  if (allRecords.length === 0) {
    _logMessage(log, `⚠️ אין רשומות חדשות לכתיבה`, 'WARNING');
    return { inserted: 0, skipped, duplicates, total: emails.length };
  }
  
  // טוסט כתיבה
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `📝 כותב ${allRecords.length} רשומות לגיליון...`, 
    "Phase 2: כתיבה", 3
  );
  
  // ⚡ V32 OPTIMIZATION: כתיבה אחת לכל הרשומות!
  try {
    sheet.getRange(startingRow, 1, allRecords.length, allRecords[0].length)
         .setValues(allRecords);
    
    _logMessage(log, `✅ נכתבו ${allRecords.length} רשומות בבת אחת (שורות ${startingRow}-${startingRow + allRecords.length - 1})`, 'SUCCESS');
  } catch (writeError) {
    _logMessage(log, `❌ שגיאה בכתיבה: ${writeError.message}`, 'ERROR');
    return { inserted: 0, skipped, duplicates, total: emails.length };
  }
  
  // Flush אחד בלבד
  SpreadsheetApp.flush();
  
  // הוספת צ'קבוקסים
  try {
    sheet.getRange(startingRow, 1, allRecords.length, 1).insertCheckboxes();
    _logMessage(log, `✅ צ'קבוקסים נוצרו לשורות ${startingRow}-${startingRow + allRecords.length - 1}`, 'SUCCESS');
  } catch (e) {
    _logMessage(log, `⚠️ שגיאה ביצירת צ'קבוקסים: ${e.message}`, 'WARNING');
  }
  
  _logMessage(log, `📊 תוצאות עיבוד: ${allRecords.length} נוספו, ${skipped} דולגו, ${duplicates} כפילויות`, 'SUCCESS');
  
  return { inserted: allRecords.length, skipped, duplicates, total: emails.length };
}

// === EXISTING RECORDS GETTER ===
function getExistingRecords_BACKUP(sheet) {
  const existingRecords = new Map();
  
  try {
    // מציאת השורה האחרונה עם נתונים אמיתיים (לא רק צ'קבוקסים)
    const lastRowWithData = findLastRowWithData_BACKUP(sheet);
    
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

// פונקציה למציאת השורה האחרונה עם נתונים אמיתיים
function findLastRowWithData_BACKUP(sheet) {
  const maxRows = sheet.getLastRow();
  
  // בדיקה מלמטה למעלה לשורה עם נתונים בעמודת Email ID (עמודה H)
  for (let row = maxRows; row >= 2; row--) {
    const emailId = sheet.getRange(row, 8).getValue(); // עמודה H = Email ID
    if (emailId && emailId.toString().trim() !== '') {
      return row;
    }
  }
  
  return 1; // אם לא נמצאו נתונים, החזר 1 (רק כותרות)
}

// פונקציה לניקוי צ'קבוקסים ישנים
function cleanupOldCheckboxes_BACKUP(sheet, log) {
  const lastRowWithData = findLastRowWithData_BACKUP(sheet);
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

// === RECORD CREATOR ===
function createInvoiceRecord_BACKUP(emailData, config) {
  // יוצר רשומה + מעלה PDF ל-Drive
  const isLocal = isLocalEmail(emailData.email, emailData.subject, emailData.body);
  const category = isLocal ? "🔷 Local" : "🌐 International";

  let pdfLink = "";
  let attachmentName = "";

  // אם יש PDF מצורף - העלאה ל-Drive (כמו בקוד הישן שעבד!)
  if (emailData.pdfAttachment) {
    attachmentName = emailData.pdfAttachment.getName();
    console.log(`📎 נמצא PDF מצורף: ${attachmentName}`);
    
    try {
      // העלאה ל-Drive - זה מה שעבד בגרסה הקודמת!
      const driveFile = DriveApp.createFile(emailData.pdfAttachment);
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfLink = driveFile.getUrl();
      console.log(`✅ הועלה ל-Drive: ${pdfLink}`);
    } catch (driveError) {
      console.error(`❌ שגיאה בהעלאה ל-Drive: ${driveError.message}`);
      // אם נכשל - ננסה למצוא לינק בגוף המייל
      pdfLink = "";
    }
  }
  
  // אם אין PDF מצורף או ההעלאה נכשלה - חיפוש לינק בגוף המייל
  if (!pdfLink || pdfLink === "") {
    console.log(`🔍 מחפש לינק לחשבונית בגוף המייל...`);
    pdfLink = extractInvoiceLinkFromHtml(emailData.htmlBody) || 
              extractInvoiceLinkFromPlainText(emailData.body) || "";
    
    if (pdfLink) {
      console.log(`✅ נמצא לינק בגוף המייל: ${pdfLink.substring(0, 80)}...`);
    } else {
      console.log(`⚠️ לא נמצא לינק`);
    }
  }

  return [
    false,                                    // Action (checkbox)
    category,                                 // Category
    emailData.senderName || "",               // Sender Name
    emailData.email || "",                    // Sender Email
    formatDateInput(emailData.date),          // Date
    emailData.subject || "",                  // Subject
    pdfLink,                                  // PDF Link (Drive או מהמייל)
    emailData.messageId || "",                // Email ID
    attachmentName                            // Attachment Name
  ];
}

// === SCAN SUMMARY ===
function showScanSummary_BACKUP(results, log) {
  const message = `✅ הסריקה הושלמה!

📊 תוצאות:
• ${results.inserted} חשבוניות חדשות נוספו
• ${results.skipped} מיילים דולגו
• ${results.duplicates} כפילויות נמצאו
• ${results.total} מיילים נבדקו בסך הכל

${results.inserted > 0 ? '🎉 החשבוניות זמינות בגיליון החדש!' : '⚠️ לא נמצאו חשבוניות חדשות.'}`;

  SpreadsheetApp.getUi().alert("סיכום סריקה", message, SpreadsheetApp.getUi().ButtonSet.OK);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ הסריקה הושלמה: ${results.inserted} נוספו, ${results.skipped} דולגו`,
    "הושלם", 10
  );
  
  _logMessage(log, message.replace(/\n/g, ' '), 'SUCCESS');
}

// === END OF BACKUP ===
