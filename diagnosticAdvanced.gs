// === diagnosticAdvanced.gs - אבחון מפורט עם Gmail API ===

/**
 * אבחון מפורט עם Gmail API - מציג כל מייל בנפרד
 * מבצע סריקה מלאה ומציג פירוט מדויק של כל רשומה
 */
function diagnosticAdvanced() {
  const startTime = new Date();
  _logMessage("🚀 התחלת אבחון מפורט עם Gmail API");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!settings) {
    SpreadsheetApp.getUi().alert("❌ גיליון Settings לא נמצא. אנא הפעל Setup Sheets תחילה.");
    return;
  }

  // קריאת הגדרות
  const config = readDiagnosticSettings(settings);
  if (!config.isValid) {
    SpreadsheetApp.getUi().alert("❌ הגדרות לא תקינות. אנא מלא תאריכים ומילות מפתח.");
    return;
  }

  // יצירת גיליון אבחון מפורט
  const diagnosticSheetName = `Detailed Diagnostic ${new Date().toLocaleDateString('he-IL')} ${new Date().toLocaleTimeString('he-IL', {hour: '2-digit', minute: '2-digit'})}`;
  let diagnosticSheet = ss.getSheetByName(diagnosticSheetName);
  if (diagnosticSheet) {
    ss.deleteSheet(diagnosticSheet);
  }
  diagnosticSheet = ss.insertSheet(diagnosticSheetName);
  
  // הגדרת כותרות מפורטות
  const headers = [
    "מס'", "שולח", "מייל", "תאריך", "נושא", "מילות מפתח נמצאו", 
    "PDF?", "סטטוס לינק", "מוחרג?", "מאושר?", "יכנס לרשומות?", "סיבה", "פרטים נוספים"
  ];
  diagnosticSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  diagnosticSheet.getRange(`A1:${String.fromCharCode(64 + headers.length)}1`)
    .setFontWeight("bold")
    .setBackground("#4285f4")
    .setFontColor("white");
  diagnosticSheet.setFrozenRows(1);
  
  // הודעת התחלה
  SpreadsheetApp.getActiveSpreadsheet().toast("🔍 מתחיל סריקה מפורטת...", "אבחון מתקדם", 5);
  
  // ביצוע סריקה מפורטת
  const detailedResults = performDetailedEmailScan(config, diagnosticSheet);
  
  // הוספת סיכום
  addDetailedSummary(diagnosticSheet, detailedResults, startTime);
  
  // עיצוב הגיליון
  formatDetailedDiagnosticSheet(diagnosticSheet, detailedResults.totalEmails);
  
  SpreadsheetApp.setActiveSheet(diagnosticSheet);
  
  const duration = (new Date() - startTime) / 1000;
  const summaryMessage = `🎯 אבחון מפורט הושלם!\n\n` +
    `📊 נסרקו ${detailedResults.totalEmails} מיילים\n` +
    `✅ יכנסו לרשומות: ${detailedResults.willBeIncluded}\n` +
    `❌ יידחו: ${detailedResults.willBeRejected}\n` +
    `⏱️ זמן ביצוע: ${duration.toFixed(1)} שניות\n\n` +
    `📋 פירוט מלא זמין בגיליון "${diagnosticSheetName}"`;
  
  SpreadsheetApp.getUi().alert(summaryMessage);
  _logMessage(`✅ אבחון מפורט הושלם בתוך ${duration.toFixed(1)} שניות`);
}

// === פונקציות עזר ===

function readDiagnosticSettings(settings) {
  try {
    const startDate = settings.getRange('A2').getValue();
    const endDate = settings.getRange('B2').getValue();
    const enKeywords = settings.getRange('C2:C').getValues().flat().filter(Boolean);
    const heKeywords = settings.getRange('D2:D').getValues().flat().filter(Boolean);
    const excludedKeywords = settings.getRange('E2:E').getValues().flat().filter(Boolean);
    const excludedEmails = settings.getRange('F2:F').getValues().flat().filter(Boolean);
    const approved = settings.getRange('G2:G').getValues().flat().filter(Boolean);
    
    return {
      isValid: startDate && endDate && (enKeywords.length > 0 || heKeywords.length > 0),
      startDate,
      endDate,
      enKeywords,
      heKeywords,
      excludedKeywords,
      excludedEmails,
      approved
    };
  } catch (e) {
    return { isValid: false };
  }
}

function performDetailedEmailScan(config, diagnosticSheet) {
  _logMessage("🔍 מתחיל סריקה מפורטת של מיילים...");
  
  // בניית שאילתת Gmail
  let query = `after:${formatDateGmail(config.startDate)} before:${formatDateGmail(config.endDate)} has:attachment`;
  
  const allKeywords = [...config.enKeywords, ...config.heKeywords].filter(k => k && k.trim());
  if (allKeywords.length > 0) {
    const keywordQuery = allKeywords.map(k => `"${k.trim()}"`).join(" OR ");
    query += ` (${keywordQuery})`;
  }
  
  // הוספת שולחים מאושרים
  if (config.approved.length > 0) {
    const approvedQuery = config.approved.map(sender => `from:${sender.trim()}`).join(" OR ");
    if (allKeywords.length > 0) {
      query += ` OR (${approvedQuery})`;
    } else {
      query += ` (${approvedQuery})`;
    }
  }
  
  _logMessage(`🔍 שאילתת Gmail: ${query}`);
  
  let threads = [];
  try {
    threads = GmailApp.search(query, 0, 50); // עד 50 מיילים לאבחון מפורט
    _logMessage(`📧 נמצאו ${threads.length} מיילים לבדיקה`);
  } catch (e) {
    _logMessage(`❌ שגיאה בחיפוש Gmail: ${e.message}`);
    return { totalEmails: 0, willBeIncluded: 0, willBeRejected: 0 };
  }
  
  if (threads.length === 0) {
    return { totalEmails: 0, willBeIncluded: 0, willBeRejected: 0 };
  }
  
  let currentRow = 2;
  let willBeIncluded = 0;
  let willBeRejected = 0;
  
  // בדיקת כל מייל בפירוט
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();
    const msg = messages[messages.length - 1]; // המייל האחרון בשרשור
    
    try {
      const email = _parseSender(msg.getFrom());
      const senderName = extractSenderName(msg.getFrom());
      const subject = msg.getSubject();
      const body = msg.getPlainBody();
      const date = msg.getDate();
      
      // בדיקת כל הקריטריונים
      const isPrivate = isPrivateEmail(email);
      const isExcluded = config.excludedEmails.map(e => e.toLowerCase()).includes(email.toLowerCase());
      const isApproved = config.approved.map(e => e.toLowerCase()).includes(email.toLowerCase());
      
      // בדיקת מילות מפתח
      const foundKeywords = [];
      for (const keyword of allKeywords) {
        if (subject.toLowerCase().includes(keyword.toLowerCase()) || 
            body.toLowerCase().includes(keyword.toLowerCase())) {
          foundKeywords.push(keyword);
        }
      }
      const hasKeywords = foundKeywords.length > 0;
      
      // בדיקת מילות מפתח מוחרגות
      const foundExcludedKeywords = [];
      for (const keyword of config.excludedKeywords) {
        if (subject.toLowerCase().includes(keyword.toLowerCase())) {
          foundExcludedKeywords.push(keyword);
        }
      }
      const hasExcludedKeywords = foundExcludedKeywords.length > 0;
      
      // בדיקת PDF
      const hasPdf = hasPdfAttachment(msg);
      
      // בדיקת קישור מפורטת - שימוש בפונקציות האמיתיות
      let htmlLink = null;
      let plainLink = null;
      let driveLink = null;
      
      // ניסיון חילוץ לינק מ-HTML
      try {
        htmlLink = extractInvoiceLinkFromHtml(msg.getBody());
      } catch (e) {
        console.log(`שגיאה בחילוץ HTML: ${e.message}`);
      }
      
      // ניסיון חילוץ לינק מטקסט רגיל
      try {
        plainLink = extractInvoiceLinkFromPlainText(body);
      } catch (e) {
        console.log(`שגיאה בחילוץ טקסט: ${e.message}`);
      }
      
      // ניסיון יצירת לינק Drive מ-attachment
      try {
        driveLink = extractPdfLink(msg);
      } catch (e) {
        console.log(`שגיאה ביצירת Drive: ${e.message}`);
      }
      
      const hasLink = htmlLink || plainLink || driveLink;
      
      // ניתוח סטטוס הלינק
      let linkStatus = "❌ אין לינק";
      let linkSource = "לא זמין";
      
      if (driveLink) {
        linkStatus = "✅ Drive (נוצר)";
        linkSource = "Drive - נוצר מ-attachment";
      } else if (htmlLink) {
        if (htmlLink.includes('drive.google.com')) {
          linkStatus = "✅ Drive (HTML)";
          linkSource = "Drive - מ-HTML";
        } else if (htmlLink.includes('mail.google.com')) {
          linkStatus = "⚠️ Gmail (HTML)";
          linkSource = "Gmail - מ-HTML";
        } else {
          linkStatus = "🔗 חיצוני (HTML)";
          linkSource = "חיצוני - מ-HTML";
        }
      } else if (plainLink) {
        if (plainLink.includes('drive.google.com')) {
          linkStatus = "✅ Drive (טקסט)";
          linkSource = "Drive - מטקסט";
        } else if (plainLink.includes('mail.google.com')) {
          linkStatus = "⚠️ Gmail (טקסט)";
          linkSource = "Gmail - מטקסט";
        } else {
          linkStatus = "🔗 חיצוני (טקסט)";
          linkSource = "חיצוני - מטקסט";
        }
      }
      
      // החלטה סופית
      const shouldInclude = !isPrivate && !isExcluded && !hasExcludedKeywords && 
                           (hasKeywords || isApproved) && (hasPdf || hasLink);
      
      // קביעת סיבת דחייה
      let rejectionReason = "";
      if (!shouldInclude) {
        const reasons = [];
        if (isPrivate) reasons.push("דומיין פרטי");
        if (isExcluded) reasons.push("מוחרג");
        if (hasExcludedKeywords) reasons.push(`מילות מפתח מוחרגות: ${foundExcludedKeywords.join(', ')}`);
        if (!hasKeywords && !isApproved) reasons.push("אין מילות מפתח");
        if (!hasPdf && !hasLink) reasons.push("אין PDF או קישור");
        rejectionReason = reasons.join("; ");
      }
      
      // הכנת נתונים לגיליון
      const rowData = [
        i + 1, // מספר סידורי
        senderName || email.split('@')[0], // שולח
        email, // מייל
        formatDateInput(date), // תאריך
        subject.substring(0, 80) + (subject.length > 80 ? "..." : ""), // נושא (מקוצר)
        foundKeywords.length > 0 ? foundKeywords.join(", ") : "לא נמצאו", // מילות מפתח
        hasPdf ? "✅ כן" : "❌ לא", // PDF
        linkStatus, // סטטוס לינק עם אייקונים
        isExcluded ? "כן" : "לא", // מוחרג
        isApproved ? "כן" : "לא", // מאושר
        shouldInclude ? "✅ כן" : "❌ לא", // יכנס לרשומות
        shouldInclude ? "✅ תקין" : rejectionReason, // סיבה
        `דומיין פרטי: ${isPrivate ? 'כן' : 'לא'}, מקור לינק: ${linkSource}` // פרטים נוספים
      ];
      
      // כתיבה לגיליון
      diagnosticSheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
      
      // צביעת השורה לפי תוצאה
      const rowRange = diagnosticSheet.getRange(currentRow, 1, 1, rowData.length);
      if (shouldInclude) {
        rowRange.setBackground("#d9ead3"); // ירוק בהיר
        willBeIncluded++;
      } else {
        rowRange.setBackground("#f4cccc"); // אדום בהיר
        willBeRejected++;
      }
      
      currentRow++;
      
      // הצגת התקדמות
      if (i % 10 === 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`📧 בודק מייל ${i + 1}/${threads.length}`, "אבחון מפורט", 1);
      }
      
    } catch (e) {
      _logMessage(`⚠️ שגיאה בעיבוד מייל ${i + 1}: ${e.message}`);
      willBeRejected++;
    }
  }
  
  return {
    totalEmails: threads.length,
    willBeIncluded: willBeIncluded,
    willBeRejected: willBeRejected
  };
}

function addDetailedSummary(diagnosticSheet, results, startTime) {
  const duration = (new Date() - startTime) / 1000;
  
  // הוספת הסיכום בתחתית הטבלה
  const summaryStartCol = 1; // עמודה A
  const summaryRow = results.totalEmails + 3; // 3 שורות אחרי הנתונים
  
  // יצירת תיבת סיכום
  const summaryData = [
    ["=== סיכום אבחון ===", "", ""],
    [`📊 סה"כ נבדקו: ${results.totalEmails}`, "", ""],
    [`✅ יכנסו: ${results.willBeIncluded}`, "", ""],
    [`❌ יידחו: ${results.willBeRejected}`, "", ""],
    [`⏱️ זמן: ${duration.toFixed(1)}s`, "", ""],
    [`📈 הצלחה: ${results.totalEmails > 0 ? ((results.willBeIncluded / results.totalEmails) * 100).toFixed(1) : 0}%`, "", ""]
  ];
  
  // כתיבת הסיכום בתחתית הטבלה
  diagnosticSheet.getRange(summaryRow, summaryStartCol, summaryData.length, 3).setValues(summaryData);
  
  // עיצוב תיבת הסיכום
  const summaryRange = diagnosticSheet.getRange(summaryRow, summaryStartCol, summaryData.length, 3);
  summaryRange.setFontWeight("bold")
             .setBackground("#e8f4fd")
             .setBorder(true, true, true, true, true, true)
             .setHorizontalAlignment("center");
  
  // עיצוב הכותרת
  diagnosticSheet.getRange(summaryRow, summaryStartCol, 1, 3)
                .setBackground("#4285f4")
                .setFontColor("white");
  
  // הסיכום מועבר לראש הטבלה - לא צריך סיכום נוסף בתחתית
}

function formatDetailedDiagnosticSheet(diagnosticSheet, totalEmails) {
  // הגדרת רוחב עמודות
  diagnosticSheet.setColumnWidth(1, 50);  // מס'
  diagnosticSheet.setColumnWidth(2, 150); // שולח
  diagnosticSheet.setColumnWidth(3, 200); // מייל
  diagnosticSheet.setColumnWidth(4, 100); // תאריך
  diagnosticSheet.setColumnWidth(5, 300); // נושא
  diagnosticSheet.setColumnWidth(6, 200); // מילות מפתח
  diagnosticSheet.setColumnWidth(7, 80);  // PDF
  diagnosticSheet.setColumnWidth(8, 120); // סטטוס לינק
  diagnosticSheet.setColumnWidth(9, 60);  // מוחרג
  diagnosticSheet.setColumnWidth(10, 60); // מאושר
  diagnosticSheet.setColumnWidth(11, 80); // יכנס
  diagnosticSheet.setColumnWidth(12, 250); // סיבה
  diagnosticSheet.setColumnWidth(13, 200); // פרטים נוספים
  
  // הוספת גבולות
  if (totalEmails > 0) {
    diagnosticSheet.getRange(1, 1, totalEmails + 1, 13).setBorder(true, true, true, true, true, true);
  }
  
  // הקפאת שורת כותרות
  diagnosticSheet.setFrozenRows(1);
}

// === פונקציות עזר נוספות ===

function isPrivateEmail(email) {
  const privateDomains = ['gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'walla.co.il', 'walla.com'];
  const domain = email.split('@')[1];
  return privateDomains.includes(domain);
}

function hasPdfAttachment(message) {
  try {
    const attachments = message.getAttachments();
    return attachments.some(att => att.getContentType().includes('pdf'));
  } catch (e) {
    return false;
  }
}

// הפונקציות extractInvoiceLinkFromHtml ו-extractInvoiceLinkFromPlainText 
// מוגדרות ב-Helpers.gs - לא צריך כפילות כאן

function formatDateGmail(date) {
  if (!date || !(date instanceof Date)) return '';
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}