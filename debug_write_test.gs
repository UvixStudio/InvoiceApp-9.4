// === Debug Write Test ===

function testWritePermissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName("Log");
  
  try {
    // בדיקה 1: יצירת גיליון חדש
    const testSheetName = "DEBUG_TEST_" + new Date().getTime();
    const testSheet = ss.insertSheet(testSheetName);
    _logMessage(log, `✅ הצלחתי ליצור גיליון: ${testSheetName}`, 'SUCCESS');
    
    // בדיקה 2: כתיבת כותרות
    const headers = ["Test1", "Test2", "Test3"];
    testSheet.appendRow(headers);
    _logMessage(log, `✅ הצלחתי לכתוב כותרות: ${headers.join(', ')}`, 'SUCCESS');
    
    // בדיקה 3: כתיבת נתונים
    const testData = ["Data1", "Data2", "Data3"];
    testSheet.appendRow(testData);
    _logMessage(log, `✅ הצלחתי לכתוב נתונים: ${testData.join(', ')}`, 'SUCCESS');
    
    // בדיקה 4: flush וקריאה חזרה
    SpreadsheetApp.flush();
    const writtenData = testSheet.getRange(2, 1, 1, 3).getValues()[0];
    _logMessage(log, `✅ קראתי חזרה: ${writtenData.join(', ')}`, 'SUCCESS');
    
    // בדיקה 5: מחיקת הגיליון
    ss.deleteSheet(testSheet);
    _logMessage(log, `✅ מחקתי את גיליון הבדיקה`, 'SUCCESS');
    
    SpreadsheetApp.getUi().alert("✅ כל בדיקות הכתיבה עברו בהצלחה!");
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בבדיקת כתיבה: ${error.message}`, 'ERROR');
    _logMessage(log, `📊 Stack trace: ${error.stack}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ שגיאה בבדיקת כתיבה:\n${error.message}`);
  }
}

function debugProcessEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName("Log");
  
  // בדיקה מפורטת של פונקציית processEmails
  _logMessage(log, "🔍 מתחיל דיבוג של processEmails", 'INFO');
  
  try {
    // חיפוש הגיליון האחרון שנוצר
    const sheets = ss.getSheets();
    let invoiceSheet = null;
    
    for (const sheet of sheets) {
      if (sheet.getName().startsWith("invoices ")) {
        invoiceSheet = sheet;
        break;
      }
    }
    
    if (!invoiceSheet) {
      _logMessage(log, "❌ לא נמצא גיליון חשבוניות", 'ERROR');
      return;
    }
    
    _logMessage(log, `📋 נמצא גיליון: ${invoiceSheet.getName()}`, 'INFO');
    
    // בדיקת תוכן הגיליון
    const lastRow = invoiceSheet.getLastRow();
    const lastCol = invoiceSheet.getLastColumn();
    
    _logMessage(log, `📊 גיליון: ${lastRow} שורות, ${lastCol} עמודות`, 'INFO');
    
    if (lastRow > 1) {
      const data = invoiceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      _logMessage(log, `📋 נתונים בגיליון: ${data.length} שורות`, 'INFO');
      
      for (let i = 0; i < Math.min(data.length, 3); i++) {
        _logMessage(log, `📝 שורה ${i + 2}: ${JSON.stringify(data[i]).substring(0, 100)}...`, 'INFO');
      }
    } else {
      _logMessage(log, "❌ אין נתונים בגיליון (רק כותרות)", 'ERROR');
    }
    
    // בדיקת הרשאות כתיבה
    try {
      invoiceSheet.appendRow(["DEBUG", "TEST", "ROW", new Date().toISOString()]);
      SpreadsheetApp.flush();
      _logMessage(log, "✅ הצלחתי לכתוב שורת בדיקה", 'SUCCESS');
      
      // מחיקת שורת הבדיקה
      const newLastRow = invoiceSheet.getLastRow();
      if (newLastRow > 1) {
        invoiceSheet.deleteRow(newLastRow);
        _logMessage(log, "✅ מחקתי את שורת הבדיקה", 'SUCCESS');
      }
    } catch (writeError) {
      _logMessage(log, `❌ שגיאה בכתיבה: ${writeError.message}`, 'ERROR');
    }
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בדיבוג: ${error.message}`, 'ERROR');
  }
}

function debugGmailData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName("Log");
  const settings = ss.getSheetByName("Settings");
  
  _logMessage(log, "🔍 מתחיל דיבוג נתוני Gmail", 'INFO');
  
  try {
    // קריאת הגדרות
    const startStr = settings.getRange('A2').getValue();
    const endStr = settings.getRange('B2').getValue();
    const startDate = new Date(startStr);
    const endDate = new Date(endStr);
    
    // בניית query
    const query = `after:${formatDateGmail(startDate)} before:${formatDateGmail(endDate)} has:attachment`;
    _logMessage(log, `🔍 Query: ${query}`, 'INFO');
    
    // חיפוש מיילים
    const threads = GmailApp.search(query, 0, 5); // רק 5 ראשונים
    _logMessage(log, `📧 נמצאו ${threads.length} threads`, 'INFO');
    
    if (threads.length === 0) {
      _logMessage(log, "❌ אין threads לבדיקה", 'ERROR');
      return;
    }
    
    // בדיקת thread ראשון
    const firstThread = threads[0];
    const messages = firstThread.getMessages();
    _logMessage(log, `📨 Thread ראשון: ${messages.length} הודעות`, 'INFO');
    
    if (messages.length === 0) {
      _logMessage(log, "❌ אין הודעות ב-thread", 'ERROR');
      return;
    }
    
    // בדיקת הודעה ראשונה
    const firstMessage = messages[0];
    _logMessage(log, `📧 בודק הודעה ראשונה...`, 'INFO');
    
    // חילוץ נתונים בסיסיים
    const from = firstMessage.getFrom();
    const subject = firstMessage.getSubject();
    const date = firstMessage.getDate();
    const messageId = firstMessage.getId();
    
    _logMessage(log, `📧 From: ${from}`, 'INFO');
    _logMessage(log, `📧 Subject: ${subject}`, 'INFO');
    _logMessage(log, `📧 Date: ${date}`, 'INFO');
    _logMessage(log, `📧 ID: ${messageId}`, 'INFO');
    
    // בדיקת פונקציות עזר
    const email = _parseSender(from);
    const senderName = extractSenderName(email);
    
    _logMessage(log, `👤 Parsed email: ${email}`, 'INFO');
    _logMessage(log, `👤 Sender name: ${senderName}`, 'INFO');
    
    // בדיקת קבצים מצורפים
    const attachments = firstMessage.getAttachments();
    _logMessage(log, `📎 Attachments: ${attachments.length}`, 'INFO');
    
    if (attachments.length > 0) {
      for (let i = 0; i < Math.min(attachments.length, 3); i++) {
        const att = attachments[i];
        _logMessage(log, `📎 Attachment ${i+1}: ${att.getName()} (${att.getContentType()})`, 'INFO');
      }
    }
    
    // בדיקת extractEmailData
    const emailData = extractEmailData(firstMessage);
    if (emailData) {
      _logMessage(log, `✅ extractEmailData החזיר נתונים`, 'SUCCESS');
      _logMessage(log, `📧 Email: ${emailData.email}`, 'INFO');
      _logMessage(log, `👤 Sender: ${emailData.senderName}`, 'INFO');
      _logMessage(log, `📧 Subject: ${emailData.subject}`, 'INFO');
      _logMessage(log, `📅 Date: ${emailData.date}`, 'INFO');
    } else {
      _logMessage(log, `❌ extractEmailData החזיר null`, 'ERROR');
    }
    
    // בדיקת createInvoiceRecord
    if (emailData) {
      const config = { excludedEmails: [], excludedKeywords: [], approvedSenders: [] };
      const record = createInvoiceRecord(emailData, config);
      
      if (record) {
        _logMessage(log, `✅ createInvoiceRecord החזיר רשומה`, 'SUCCESS');
        _logMessage(log, `📋 Record: ${JSON.stringify(record).substring(0, 200)}...`, 'INFO');
      } else {
        _logMessage(log, `❌ createInvoiceRecord החזיר null`, 'ERROR');
      }
    }
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בדיבוג Gmail: ${error.message}`, 'ERROR');
    _logMessage(log, `📊 Stack: ${error.stack}`, 'ERROR');
  }
}

function debugEmailFiltering() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName("Log");
  const settings = ss.getSheetByName("Settings");
  
  _logMessage(log, "🔍 מתחיל דיבוג סינון מיילים", 'INFO');
  
  try {
    // קריאת הגדרות
    const startStr = settings.getRange('A2').getValue();
    const endStr = settings.getRange('B2').getValue();
    const startDate = new Date(startStr);
    const endDate = new Date(endStr);
    
    const enKeywords = settings.getRange('C2:C').getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
    const heKeywords = settings.getRange('D2:D').getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
    const excludedKeywords = settings.getRange('E2:E').getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
    const excludedEmails = settings.getRange('F2:F').getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
    const approvedSenders = settings.getRange('G2:G').getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
    
    const config = {
      enKeywords,
      heKeywords,
      excludedKeywords,
      excludedEmails,
      approvedSenders
    };
    
    // בניית query
    const query = `after:${formatDateGmail(startDate)} before:${formatDateGmail(endDate)} has:attachment`;
    _logMessage(log, `🔍 Query: ${query}`, 'INFO');
    
    // חיפוש מיילים
    const threads = GmailApp.search(query, 0, 3); // רק 3 ראשונים
    _logMessage(log, `📧 נמצאו ${threads.length} threads`, 'INFO');
    
    if (threads.length === 0) {
      _logMessage(log, "❌ אין threads לבדיקה", 'ERROR');
      return;
    }
    
    // בדיקת כל המיילים
    for (let t = 0; t < threads.length; t++) {
      const thread = threads[t];
      const messages = thread.getMessages();
      
      for (let m = 0; m < messages.length; m++) {
        const message = messages[m];
        const emailData = extractEmailData(message);
        
        if (!emailData) continue;
        
        _logMessage(log, `📧 בודק מייל ${t+1}.${m+1}: ${emailData.subject}`, 'INFO');
        _logMessage(log, `👤 שולח: ${emailData.email}`, 'INFO');
        
        // בדיקת סינון צעד אחר צעד
        const email = emailData.email.toLowerCase();
        const subject = emailData.subject.toLowerCase();
        const body = emailData.body.toLowerCase();
        
        // בדיקה 1: מיילים מוחרגים
        const isExcludedEmail = config.excludedEmails.some(excluded => email.includes(excluded));
        _logMessage(log, `🚫 מייל מוחרג: ${isExcludedEmail}`, isExcludedEmail ? 'WARNING' : 'INFO');
        
        // בדיקה 2: מילים מוחרגות
        const hasExcludedWords = config.excludedKeywords.some(word => subject.includes(word) || body.includes(word));
        _logMessage(log, `🚫 מילים מוחרגות: ${hasExcludedWords}`, hasExcludedWords ? 'WARNING' : 'INFO');
        
        // בדיקה 3: הזמנות קלנדר
        const isCalendarInvite = subject.includes('invitation:') || subject.includes('invite') || 
                                subject.includes('meeting') || subject.includes('calendar');
        _logMessage(log, `📅 הזמנת קלנדר: ${isCalendarInvite}`, isCalendarInvite ? 'WARNING' : 'INFO');
        
        // בדיקה 4: קבצי .ics
        const hasIcsFiles = emailData.attachments && emailData.attachments.some(att => 
            att.getName().toLowerCase().endsWith('.ics') || 
            att.getContentType().toLowerCase().includes('calendar'));
        _logMessage(log, `📎 קבצי ICS: ${hasIcsFiles}`, hasIcsFiles ? 'WARNING' : 'INFO');
        
        // בדיקה 5: דומיינים פרטיים
        const isPrivate = isPrivateEmail(email);
        const isApproved = config.approvedSenders.some(approved => email.includes(approved));
        const shouldSkipPrivate = isPrivate && !isApproved;
        _logMessage(log, `🏠 דומיין פרטי: ${isPrivate}, מאושר: ${isApproved}, לדלג: ${shouldSkipPrivate}`, shouldSkipPrivate ? 'WARNING' : 'INFO');
        
        // בדיקה 6: מילות מפתח
        const allKeywords = [...config.enKeywords, ...config.heKeywords];
        const hasRelevantKeywords = allKeywords.some(keyword => 
          subject.includes(keyword.toLowerCase()) || body.includes(keyword.toLowerCase())
        );
        const shouldSkipKeywords = !hasRelevantKeywords && !isApproved;
        _logMessage(log, `🔤 מילות מפתח: ${hasRelevantKeywords}, לדלג: ${shouldSkipKeywords}`, shouldSkipKeywords ? 'WARNING' : 'INFO');
        
        // בדיקה 7: PDF או קישור
        const hasPdfOrLink = emailData.pdfAttachment || hasInvoiceLink(emailData);
        _logMessage(log, `📎 PDF/קישור: ${hasPdfOrLink}`, !hasPdfOrLink ? 'WARNING' : 'INFO');
        
        // תוצאה סופית
        const shouldSkip = shouldSkipEmail(emailData, config);
        _logMessage(log, `🎯 תוצאה סופית - לדלג: ${shouldSkip}`, shouldSkip ? 'WARNING' : 'SUCCESS');
        
        _logMessage(log, `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`, 'INFO');
      }
    }
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בדיבוג סינון: ${error.message}`, 'ERROR');
  }
}

function debugSheetOperations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName("Log");
  
  _logMessage(log, "🔍 מתחיל דיבוג פעולות גיליון", 'INFO');
  
  try {
    // מציאת גיליון החשבוניות האחרון
    const sheets = ss.getSheets();
    let invoiceSheet = null;
    
    for (const sheet of sheets) {
      if (sheet.getName().startsWith("invoices ")) {
        invoiceSheet = sheet;
        break;
      }
    }
    
    if (!invoiceSheet) {
      _logMessage(log, "❌ לא נמצא גיליון חשבוניות", 'ERROR');
      return;
    }
    
    const sheetName = invoiceSheet.getName();
    _logMessage(log, `📋 בודק גיליון: ${sheetName}`, 'INFO');
    
    // בדיקת מבנה הגיליון
    const lastRow = invoiceSheet.getLastRow();
    const lastCol = invoiceSheet.getLastColumn();
    _logMessage(log, `📊 מבנה גיליון: ${lastRow} שורות, ${lastCol} עמודות`, 'INFO');
    
    // בדיקת כותרות
    if (lastRow >= 1) {
      const headers = invoiceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      _logMessage(log, `📋 כותרות: ${JSON.stringify(headers)}`, 'INFO');
    }
    
    // בדיקת נתונים
    if (lastRow > 1) {
      const dataRange = invoiceSheet.getRange(2, 1, lastRow - 1, lastCol);
      const data = dataRange.getValues();
      
      _logMessage(log, `📊 ${data.length} שורות נתונים נמצאו`, 'INFO');
      
      // בדיקת כמה שורות ריקות
      let emptyRows = 0;
      let nonEmptyRows = 0;
      
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const isEmpty = row.every(cell => cell === "" || cell === false);
        
        if (isEmpty) {
          emptyRows++;
        } else {
          nonEmptyRows++;
          if (nonEmptyRows <= 3) { // הצג רק 3 ראשונות
            _logMessage(log, `📝 שורה ${i+2} (לא ריקה): ${JSON.stringify(row).substring(0, 150)}...`, 'INFO');
          }
        }
      }
      
      _logMessage(log, `📊 סיכום: ${nonEmptyRows} שורות עם נתונים, ${emptyRows} שורות ריקות`, nonEmptyRows > 0 ? 'SUCCESS' : 'ERROR');
    } else {
      _logMessage(log, "❌ אין שורות נתונים בגיליון (רק כותרות)", 'ERROR');
    }
    
    // בדיקת הרשאות הגיליון
    try {
      const protection = invoiceSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      if (protection.length > 0) {
        _logMessage(log, `🔒 הגיליון מוגן! ${protection.length} הגנות פעילות`, 'WARNING');
        for (let i = 0; i < protection.length; i++) {
          const prot = protection[i];
          _logMessage(log, `🔒 הגנה ${i+1}: ${prot.getDescription()}`, 'WARNING');
        }
      } else {
        _logMessage(log, "🔓 הגיליון לא מוגן", 'SUCCESS');
      }
    } catch (protError) {
      _logMessage(log, `⚠️ לא ניתן לבדוק הגנות: ${protError.message}`, 'WARNING');
    }
    
    // בדיקת triggers אוטומטיים
    try {
      const triggers = ScriptApp.getProjectTriggers();
      _logMessage(log, `⚡ נמצאו ${triggers.length} triggers`, 'INFO');
      
      for (let i = 0; i < triggers.length; i++) {
        const trigger = triggers[i];
        _logMessage(log, `⚡ Trigger ${i+1}: ${trigger.getHandlerFunction()} - ${trigger.getEventType()}`, 'INFO');
      }
    } catch (triggerError) {
      _logMessage(log, `⚠️ לא ניתן לבדוק triggers: ${triggerError.message}`, 'WARNING');
    }
    
    // בדיקת היסטוריית שינויים (אם אפשר)
    try {
      const revisions = DriveApp.getFileById(ss.getId()).getRevisions();
      _logMessage(log, `📚 נמצאו ${revisions.length} גרסאות של הקובץ`, 'INFO');
    } catch (revError) {
      _logMessage(log, `⚠️ לא ניתן לבדוק היסטוריה: ${revError.message}`, 'WARNING');
    }
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בדיבוג גיליון: ${error.message}`, 'ERROR');
    _logMessage(log, `📊 Stack: ${error.stack}`, 'ERROR');
  }
}

function debugMenuActions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName("Log");
  
  _logMessage(log, "🔍 מתחיל דיבוג תפריט פעולות", 'INFO');
  
  try {
    // בדיקה אם יש פונקציות שרצות אוטומטית
    const functionNames = [
      'excludeSelected', 'markAsLocal', 'markAsInternational', 
      'exportSelected', 'deleteSelected', 'approveSelected',
      'runActions', 'processActionSelection'
    ];
    
    for (const funcName of functionNames) {
      try {
        const func = eval(funcName);
        if (typeof func === 'function') {
          _logMessage(log, `✅ פונקציה ${funcName} קיימת`, 'INFO');
        }
      } catch (e) {
        _logMessage(log, `❌ פונקציה ${funcName} לא קיימת או שגיאה: ${e.message}`, 'WARNING');
      }
    }
    
    // בדיקת הגדרות תפריט
    _logMessage(log, "🔍 בודק אם יש פעולות אוטומטיות שרצות", 'INFO');
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בדיבוג תפריט: ${error.message}`, 'ERROR');
  }
}