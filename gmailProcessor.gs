// === GmailProcessor.gs - סריקה חדשה ופועלת ===

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

    // התחלת סריקה
    SpreadsheetApp.getActiveSpreadsheet().toast("🔍 מתחיל סריקת Gmail...", "סריקה", 5);
    
    const scanResults = performGmailScan(config, log);
    
    if (scanResults.emails.length === 0) {
      SpreadsheetApp.getUi().alert("לא נמצאו מיילים בטווח התאריכים שהוגדר.\n\nבדוק:\n1. טווח התאריכים\n2. מילות המפתח\n3. שיש מיילים עם קבצים מצורפים");
      _logMessage(log, "❌ לא נמצאו מיילים לעיבוד", 'WARNING');
      return;
    }

    // עיבוד המיילים
    SpreadsheetApp.getActiveSpreadsheet().toast(`📧 מעבד ${scanResults.emails.length} מיילים...`, "עיבוד", 5);
    
    const processResults = processEmails(scanResults.emails, config, sheet, log);
    
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
  SpreadsheetApp.getActiveSpreadsheet().toast("⚡ מתחיל סריקה מהירה...", "סריקה מהירה", 3);
  
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
      
      // צ'קבוקסים יתווספו אחרי כתיבת הנתונים
      // const checkboxRange = sheet.getRange("A2:A1000"); // הוסר - גורם לבעיית שורה 1000
      // checkboxRange.insertCheckboxes();
      
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
function performGmailScan(config, log) {
  try {
    // בניית שאילתת חיפוש
    const query = buildGmailQuery(config);
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
  parts.push('-subject:(re: OR fwd: OR fw: OR "תשובה" OR "העברה")');
  
  // מילות חיפוש רק בנושא (subject) - OR ביניהן
  const allKeywords = [...config.enKeywords, ...config.heKeywords].filter(k => k && k.trim());
  if (allKeywords.length > 0) {
    const keywordQuery = allKeywords.map(k => k.trim()).join(" OR ");
    parts.push(`subject:(${keywordQuery})`);
  } else {
    // אם אין מילות מפתח בסטינגס - תן שגיאה ברורה
    throw new Error('❌ לא הוגדרו מילות חיפוש בטאב Settings! אנא הוסף מילות חיפוש בעמודות C ו-D');
  }
  
  // חריגת מיילים מוחרגים
  if (config.excludedEmails && config.excludedEmails.length > 0) {
    const excludedQuery = config.excludedEmails.map(email => `-from:${email.trim()}`).join(" ");
    parts.push(excludedQuery);
  }
  
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
function processEmails(emails, config, sheet, log) {
  let inserted = 0;
  let skipped = 0;
  let duplicates = 0;
  
  // קבלת רשומות קיימות
  const existingRecords = getExistingRecords(sheet);
  
  _logMessage(log, `🔄 מתחיל עיבוד ${emails.length} מיילים`, 'INFO');
  
  // הודעה על מיקום כתיבת הרשומות
  const startingRow = findLastRowWithData(sheet) + 1;
  _logMessage(log, `📝 רשומות חדשות יתחילו משורה ${startingRow}`, 'INFO');
  
  for (let i = 0; i < emails.length; i++) {
    const emailData = emails[i];
    
    // עדכון התקדמות
    if (i % 10 === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `מעבד מייל ${i + 1}/${emails.length}...`, 
        "עיבוד", 2
      );
    }
    
    try {
      // בדיקת כפילויות
      if (existingRecords.has(emailData.messageId)) {
        duplicates++;
        continue;
      }
      
      // השאילתה כבר עשתה את כל הסינון - אנחנו רק כותבים
      // אין סינון נוסף בשלב זה
      _logMessage(log, `✅ מתקבל: ${emailData.email} - "${emailData.subject}"`, 'SUCCESS');
      
      // יצירת רשומה - בלי בדיקות, הכל נכנס
      try {
        const record = createInvoiceRecord(emailData, config);
        
        // מציאת השורה הבאה לכתיבה
        const nextRow = findLastRowWithData(sheet) + 1;
        
        // כתיבה בשורה ספציפית
        sheet.getRange(nextRow, 1, 1, record.length).setValues([record]);
        
        SpreadsheetApp.flush();
        
        _logMessage(log, `✅ נוסף: ${emailData.email} - שורה ${nextRow}`, 'SUCCESS');
        
        inserted++;
        existingRecords.set(emailData.messageId, true);
        
        // Flush כל 20 רשומות
        if (inserted % 20 === 0) {
          SpreadsheetApp.flush();
        }
      } catch (recordError) {
        _logMessage(log, `❌ שגיאה ביצירת רשומה: ${emailData.email} - ${recordError.message}`, 'ERROR');
        skipped++;
      }
      
    } catch (error) {
      _logMessage(log, `❌ שגיאה בעיבוד מייל ${emailData.email}: ${error.message}`, 'ERROR');
      skipped++;
    }
  }
  
  // Flush סופי
  SpreadsheetApp.flush();
  
  // הוספת צ'קבוקסים רק לשורות שנכתבו בפועל
  if (inserted > 0) {
    try {
      // מציאת השורה האחרונה עם נתונים אמיתיים
      const lastRowWithData = findLastRowWithData(sheet);
      const startRow = lastRowWithData - inserted + 1; // השורה הראשונה של הרשומות החדשות
      
      sheet.getRange(startRow, 1, inserted, 1).insertCheckboxes();
      _logMessage(log, `✅ צ'קבוקסים נוצרו לשורות ${startRow}-${lastRowWithData} (${inserted} רשומות)`, 'SUCCESS');
    } catch (e) {
      _logMessage(log, `⚠️ שגיאה ביצירת צ'קבוקסים: ${e.message}`, 'WARNING');
    }
  }
  
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

// פונקציה למציאת השורה האחרונה עם נתונים אמיתיים
function findLastRowWithData(sheet) {
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
  
  // רק בדיקות בסיסיות - השאילתה כבר עשתה את העבודה הכבדה
  
  // בדיקת מיילים מוחרגים מפורשים
  const excludedEmail = config.excludedEmails.find(excluded => email.includes(excluded));
  if (excludedEmail) {
    return `מייל מוחרג: ${excludedEmail}`;
  }
  
  // בדיקת מילים מוחרגות בנושא בלבד
  const excludedWord = config.excludedKeywords.find(word => subject.includes(word));
  if (excludedWord) {
    return `מילה מוחרגת: "${excludedWord}"`;
  }
  
  // דילוג על הזמנות קלנדר ברורות
  if (subject.includes('invitation:') || subject.includes('invite') || 
      subject.includes('meeting') || subject.includes('calendar')) {
    return 'הזמנת קלנדר';
  }
  
  // השאילתה כבר סיננה subject: ו-filename:pdf
  // אז אנחנו מאשרים כמעט הכל שהגיע עד כאן
  
  return null; // לא דולג - מאשר את המייל
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

// === RECORD CREATOR ===
function createInvoiceRecord(emailData, config) {
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

${results.inserted > 0 ? '🎉 החשבוניות זמינות בגיליון החדש!' : '⚠️ לא נמצאו חשבוניות חדשות.'}`;

  SpreadsheetApp.getUi().alert("סיכום סריקה", message, SpreadsheetApp.getUi().ButtonSet.OK);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ הסריקה הושלמה: ${results.inserted} נוספו, ${results.skipped} דולגו`,
    "הושלם", 10
  );
  
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
        setSenderCategory(row.email, category);
        const rowRange = sheet.getRange(row.rowIndex, categoryColumnIndex);
        rowRange.setValue(category);
      }
      return { 
        success: true, 
        message: `✅ ${checkedRows.length} חשבוניות סומנו כ-${category} ונשמרו להמשך.` 
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