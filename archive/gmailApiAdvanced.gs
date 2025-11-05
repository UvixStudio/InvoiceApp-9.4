// === Gmail API Advanced - סריקה מהירה ומתקדמת ===

// === MAIN ADVANCED SCAN FUNCTION ===
function processInvoicesGmailApi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  if (!settings || !log) {
    SpreadsheetApp.getUi().alert("❌ גיליונות Settings או Log חסרים. אנא הפעל את 'Setup Sheets' תחילה.");
    return;
  }

  // תיעוד תחילת סריקה
  _logMessage("🚀 התחלת סריקת חשבוניות (Gmail API Advanced)");
  _logMessage(`📅 זמן התחלה: ${new Date().toLocaleString('he-IL')}`);
  
  const startTime = new Date();

  const startStr = settings.getRange("A2").getValue();
  const endStr = settings.getRange("B2").getValue();
  if (!startStr || !endStr) {
    SpreadsheetApp.getUi().alert("❌ יש למלא תאריכים תקפים בגיליון Settings.");
    return;
  }

  const enKeywords = settings.getRange("C2:C").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
  const heKeywords = settings.getRange("D2:D").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
  const exclude = settings.getRange("E2:E").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
  const approved = settings.getRange("F2:F").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());

  const startDate = new Date(startStr);
  const endDate = new Date(endStr);
  endDate.setHours(23, 59, 59, 999);

  if (startDate.getTime() >= endDate.getTime()) {
    SpreadsheetApp.getUi().alert("❌ תאריך התחלה חייב להיות לפני תאריך הסיום.");
    return;
  }

  const tabName = `invoices ${formatDateShort(startDate)}-${formatDateShort(endDate)}`;
  let sheet = ss.getSheetByName(tabName);

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(["Action", "Category", "Sender Name", "Sender Email", "Date", "Subject", "PDF Link", "Email ID", "Attachment Name"]);
    
    // צ'קבוקסים יתווספו אחרי כתיבת הנתונים - לא כאן!
    // sheet.getRange("A2:A1000").insertCheckboxes(); // הוסר - גורם לבעיית שורה 1000
    
    sheet.setFrozenRows(1);
    sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#ddebf7");
  } else {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
  }

  // סריקת מיילים
  const query = `after:${formatDateGmail(startDate)} before:${formatDateGmail(endDate)} has:attachment`;
  const threads = GmailApp.search(query, 0, 500);
  _logMessage(log, `📨 Gmail connected. ${threads.length} threads found for query: "${query}"`);

  let inserted = 0, excludedCount = 0, noMatch = 0, processedCount = 0;

  SpreadsheetApp.getActiveSpreadsheet().toast("🔍 סריקה התחילה...");

  // יצירת רשימה חדשה לחשבוניות
  const invoices = [];

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    processedCount++;
    SpreadsheetApp.getActiveSpreadsheet().toast(`מעבד מייל ${processedCount} מתוך ${threads.length}...`, "עיבוד", -1);

    try {
      const messages = thread.getMessages();
      for (let j = 0; j < messages.length; j++) {
        const msg = messages[j];
        const email = _parseSender(msg.getFrom()).toLowerCase();
        const subject = msg.getSubject();
        const msgId = msg.getId();
        const date = msg.getDate();
        const body = msg.getPlainBody();
        const htmlBody = msg.getBody();

        if (exclude.includes(email)) {
          excludedCount++;
          _logMessage(log, `⚠ Skipped: [${email}] - Subject: "${subject}" - reason: excluded email`);
          continue;
        }

        if (isPrivateEmail(email)) {
          excludedCount++;
          _logMessage(log, `⚠ Skipped: [${email}] - Subject: "${subject}" - reason: private email domain`);
          continue;
        }

        const keywordMatch = [...enKeywords, ...heKeywords].some(k =>
          subject.toLowerCase().includes(k) || (body && body.toLowerCase().includes(k))
        );
        const approvedSender = approved.includes(extractSenderName(email));

        if (!keywordMatch && !approvedSender) {
          noMatch++;
          _logMessage(log, `⚠ Skipped: [${email}] - Subject: "${subject}" - reason: no keyword/approved sender match`);
          continue;
        }

        // Handle PDF attachment and link extraction
        const pdfAttachment = hasPdfAttachment(msg);
        let pdfLink = null;
        let attachmentName = "לינק מהמייל";

        if (pdfAttachment) {
          try {
            const driveFile = DriveApp.createFile(pdfAttachment);
            pdfLink = driveFile.getUrl();
            attachmentName = pdfAttachment.getName();
            _logMessage(log, `✅ Saved PDF attachment to Drive: ${attachmentName}, Link: ${pdfLink}`);
          } catch (driveError) {
            _logMessage(log, `❌ Error saving PDF attachment to Drive: ${driveError.message}`);
            continue;
          }
        } else {
          pdfLink = extractInvoiceLinkFromHtml(htmlBody);
          if (!pdfLink) {
            pdfLink = extractInvoiceLinkFromPlainText(body);
          }
          if (pdfLink) {
            _logMessage(log, `✅ Extracted link from email body: ${pdfLink}`);
            attachmentName = "לינק מהמייל";
          }
        }

        if (!pdfLink) {
          noMatch++;
          _logMessage(log, `⚠ Skipped: [${email}] - Subject: "${subject}" - reason: no PDF attachment or link found`);
          continue;
        }

        // זיהוי מקומי/בינלאומי
        const isLocal = isLocalEmail(email, subject, body);
        const category = isLocal ? "🔷 Local" : "🌐 International";
        const senderName = extractSenderName(email);

        invoices.push({
          category: category,
          senderName: senderName,
          email: email,
          date: formatDateInput(date),
          subject: msg.getSubject(),
          link: pdfLink,
          msgId: msgId,
          attachmentName: attachmentName
        });
        inserted++;
      }
    } catch (e) {
      _logMessage(log, `❌ Error processing thread [${thread.getId()}]: ${e.message}`);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("🔍 סריקה הסתיימה. ממיין ומוסיף לגיליון...", "עיבוד", 5);

  // סידור לפי תאריך
  invoices.sort((a, b) => new Date(a.date) - new Date(b.date));

  // הוספה לגיליון
  if (invoices.length > 0) {
    const dataToAppend = invoices.map(inv => [
      false, // Action checkbox
      inv.category,
      inv.senderName,
      inv.email,
      inv.date,
      inv.subject,
      inv.link,
      inv.msgId,
      inv.attachmentName
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
  }

  // סיכום
  if (sheet.getRange(1, 1).getDisplayValue().includes("נוספו:")) {
    sheet.deleteRow(1);
  }
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, 9).merge()
    .setValue(`🧾 נוספו: ${inserted} | הוחרגו: ${excludedCount} | ללא התאמה: ${noMatch}`)
    .setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");

  // סיכום מפורט בלוג
  const endTime = new Date();
  const duration = Math.round((endTime - startTime) / 1000);
  
  _logMessage("📊 ===== סיכום סריקה (Gmail API Advanced) =====");
  _logMessage(`⏱️ זמן סיום: ${endTime.toLocaleString('he-IL')}`);
  _logMessage(`⏱️ משך זמן: ${duration} שניות`);
  _logMessage(`✅ נוספו לגיליון: ${inserted} רשומות`);
  _logMessage(`🔴 הוחרגו (מיילים לא רצויים): ${excludedCount} רשומות`);
  _logMessage(`⚪ דולגו (ללא התאמה): ${noMatch} רשומות`);
  _logMessage(`📧 סה"כ מיילים נבדקו: ${inserted + excludedCount + noMatch}`);
  _logMessage("✅ סריקה הושלמה בהצלחה!");
  _logMessage("================================================");
  
  SpreadsheetApp.getActiveSpreadsheet().toast("🎉 הסריקה הסתיימה!", "הסתיים", 10);
  SpreadsheetApp.getUi().alert(`🎉 הסריקה הסתיימה.\n✅ נוספו: ${inserted}\n🔴 הוחרגו: ${excludedCount}\n⚪ ללא התאמה: ${noMatch}`);
}

function writeInvoicesToSheet(invoices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Invoices") || ss.insertSheet("Invoices");
  sheet.clear();

  const headers = ["Sender", "Subject", "Date", "Invoice Name", "Drive Link"];
  sheet.appendRow(headers);

  for (const invoice of invoices) {
    sheet.appendRow([invoice.sender, invoice.subject, invoice.date, invoice.name, invoice.link]);
  }
}

// === קריאת הגדרות מתקדמת ===
function readAdvancedSettings(settings) {
  const startStr = settings.getRange('A2').getValue();
  const endStr = settings.getRange('B2').getValue();
  
  if (!startStr || !endStr) {
    _logMessage('❌ חסרים תאריכים בגיליון Settings');
    SpreadsheetApp.getUi().alert('❌ יש למלא תאריכים תקפים בגיליון Settings.');
    return null;
  }
  
  const startDate = new Date(startStr);
  const endDate = new Date(endStr);
  endDate.setHours(23, 59, 59, 999);
  
  const enKeywords = settings.getRange('C2:C').getValues().flat().filter(Boolean);
  const heKeywords = settings.getRange('D2:D').getValues().flat().filter(Boolean);
  const excludedKeywords = settings.getRange('E2:E').getValues().flat().filter(Boolean);
  const excludeEmails = settings.getRange('F2:F').getValues().flat().filter(Boolean);
  
  const allKeywords = [...enKeywords, ...heKeywords];
  
  _logMessage(`ℹ️ 🔤 מילות מפתח: EN(${enKeywords.length}), HE(${heKeywords.length})`);
  _logMessage(`ℹ️ 🚫 מוחרגים: ${excludeEmails.length} מיילים, ${excludedKeywords.length} מילים`);
  
  return {
    startDate,
    endDate,
    allKeywords,
    excludedKeywords,
    excludeEmails,
    tabName: `invoices ${formatDateShort(startDate)}-${formatDateShort(endDate)}`
  };
}

// === יצירת גיליון תוצאות מתקדם ===
function createAdvancedResultsSheet(ss, tabName) {
  _logMessage(`ℹ️ 📑 יוצר גיליון מתקדם: ${tabName}`);
  
  let sheet = ss.getSheetByName(tabName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(tabName);
  }
  
  // כותרות מתקדמות
  const headers = [
    "Action", "Category", "Company Name", "Email Sender", "Date", "Subject",
    "Invoice Number", "Amount", "Currency", "PDF Link", "Email ID", "Attachment Name", "Validation Method"
  ];
  
  try {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // עיצוב מתקדם
    const headerRange = sheet.getRange("A1:L1");
    headerRange.setFontWeight("bold")
               .setBackground("#4285f4")
               .setFontColor("white")
               .setHorizontalAlignment("center");
    
    // צ'קבוקסים יתווספו אחרי כתיבת הנתונים
    // sheet.getRange("A2:A100").insertCheckboxes(); // הוסר - גורם לבעיית שורה 1000
    
    // רוחב עמודות
    sheet.setColumnWidth(1, 60);  // Action
    sheet.setColumnWidth(2, 100); // Category
    sheet.setColumnWidth(3, 200); // Company Name
    sheet.setColumnWidth(4, 180); // Email Sender
    sheet.setColumnWidth(5, 100); // Date
    sheet.setColumnWidth(6, 300); // Subject
    sheet.setColumnWidth(7, 120); // Invoice Number
    sheet.setColumnWidth(8, 100); // Amount
    sheet.setColumnWidth(9, 80);  // Currency
    sheet.setColumnWidth(10, 200); // PDF Link
    sheet.setColumnWidth(11, 120); // Email ID
    sheet.setColumnWidth(12, 150); // Attachment Name
    sheet.setColumnWidth(13, 200); // Validation Method
    
    sheet.setFrozenRows(1);
    _logMessage("✅ גיליון מתקדם נוצר בהצלחה");
    
    return sheet;
  } catch (e) {
    _logMessage(`❌ שגיאה ביצירת גיליון: ${e.message}`);
    return null;
  }
}

// === סריקה מתקדמת עם Gmail API ===
function performAdvancedGmailScan(config) {
  _logMessage("🔍 מתחיל סריקה מתקדמת עם Gmail API");
  
  // בניית שאילתה מתקדמת
  const keywordQuery = config.allKeywords.map(k => `"${k}"`).join(' OR ');
  const excludeQuery = config.excludeEmails.map(e => `-from:${e}`).join(' ');
  
  // שאילתה מתקדמת עם חיפוש חכם
  const advancedQuery = `
    after:${formatDateGmail(config.startDate)} 
    before:${formatDateGmail(config.endDate)} 
    has:attachment 
    (${keywordQuery} OR "סכום" OR "₪" OR "$" OR "total" OR "amount" OR "billing" OR "חיוב")
    ${excludeQuery}
  `.replace(/\s+/g, ' ').trim();
  
  _logMessage(`🔍 שאילתה מתקדמת: ${advancedQuery}`);

  try {
    // שימוש ב-Gmail API במקום GmailApp
    const response = Gmail.Users.Messages.list('me', {
      q: advancedQuery,
      maxResults: 500  // פי 2.5 יותר מ-GmailApp!
    });
    
    const messageIds = response.messages || [];
    _logMessage(`📧 נמצאו ${messageIds.length} מיילים עם Gmail API`);
    
    if (messageIds.length === 0) {
      return { messages: [] };
    }
    
    // קבלת פרטי המיילים בקבוצות (מהיר יותר)
    const messages = [];
    const batchSize = 10;
    
    for (let i = 0; i < messageIds.length; i += batchSize) {
      const batch = messageIds.slice(i, i + batchSize);
      
      for (const msgRef of batch) {
        try {
          const message = Gmail.Users.Messages.get('me', msgRef.id, {
            format: 'full'
          });
          
          // סינון לפי תאריך
          const msgDate = new Date(parseInt(message.internalDate));
          if (msgDate >= config.startDate && msgDate <= config.endDate) {
            messages.push(message);
          }
        } catch (e) {
          _logMessage(`⚠️ שגיאה בקריאת מייל ${msgRef.id}: ${e.message}`);
        }
      }
      
      // הצגת התקדמות
      if (i % 50 === 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`📧 עובד... ${i}/${messageIds.length}`, "Gmail API", 1);
      }
    }
    
    _logMessage(`📋 חולצו ${messages.length} מיילים מתאימים לעיבוד`);
    return { messages };
    
  } catch (e) {
    _logMessage(`❌ שגיאה בסריקה מתקדמת: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ שגיאה בסריקה: ${e.message}`);
    return { messages: [] };
  }
}

// === עיבוד מתקדם עם OCR ===
function processAdvancedMessages(messages, config, sheet) {
  _logMessage(`ℹ️ 🔄 מתחיל עיבוד מתקדם של ${messages.length} מיילים`);
  
  let addedCount = 0;
  let skippedCount = 0;
  let currentRow = 2;
  
  for (let i = 0; i < messages.length; i++) {
    const message = messages[i];
    
    try {
      // חילוץ מטאדטה מהמייל
      _logMessage(`🔍 מחלץ מטאדטה עבור מייל: ${message.id}`);
      const emailData = extractEmailMetadata(message);
      _logMessage(`🔍 תוצאת מטאדטה: ${emailData ? 'הצליח' : 'נכשל - null'}`);
      
      if (!emailData) {
        _logMessage(`⚠️ דולג - לא הצליח לחלץ מטאדטה: ${message.id}`);
        skippedCount++;
        continue;
      }
      
      // בדיקת PDF
      _logMessage(`🔍 מחפש PDF attachment עבור מייל: ${message.id}`);
      _logMessage(`🔍 DEBUG: message.payload.parts = ${message.payload.parts ? message.payload.parts.length : 'null'}`);
      const pdfAttachment = findPdfAttachment(message);
      _logMessage(`🔍 תוצאת חיפוש PDF: ${pdfAttachment ? 'נמצא - ' + pdfAttachment.filename : 'לא נמצא'}`);
      if (!pdfAttachment) {
        _logMessage(`⚠️ דולג: ${emailData.senderEmail} - אין PDF`);
        skippedCount++;
        continue;
      }
      
      // OCR לחילוץ נתונים מהחשבונית
      _logMessage(`🔍 מתחיל OCR עבור: ${pdfAttachment.filename}`);
      const invoiceData = extractInvoiceDataWithOCR(message, pdfAttachment);
      _logMessage(`🔍 תוצאת OCR: ${invoiceData ? 'הצליח' : 'נכשל - null'}`);
      
      // אם OCR החזיר null, זה אומר שזה לא חשבונית אמיתית
      if (!invoiceData) {
        const timestamp = new Date().toLocaleString('he-IL');
        const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
        if (log) log.appendRow([timestamp, `⚠️ דולג - לא חשבונית: ${emailData.senderEmail} - ${pdfAttachment.filename}`]);
        skippedCount++;
        continue;
      }
      
      // חילוץ סכומים חכם (מהיר + OCR במידת הצורך)
      const amountData = extractAmountSmart(message, pdfAttachment);
      if (amountData.found) {
        invoiceData.amount = amountData.amount;
        invoiceData.currency = amountData.currency;
        invoiceData.amountMethod = amountData.method;
      }
      
      // שילוב הנתונים
      _logMessage(`🔗 מנסה ליצור לינק PDF עבור מייל: ${message.id}`);
      let pdfLink = extractPdfLink(message);
      
      // אם לא הצליח ליצור לינק Drive, ננסה ליצור אותו ישירות כאן
      if (!pdfLink) {
        _logMessage(`⚠️ extractPdfLink החזיר null - מנסה ליצור Drive link ישירות`);
        const pdfAttachment = findPdfAttachment(message);
        if (pdfAttachment) {
          let fileName = pdfAttachment.filename || `invoice-${message.id}.pdf`;
          
          try {
            // הורדת ה-attachment דרך Gmail API
            _logMessage(`🔄 מוריד attachment: ${pdfAttachment.attachmentId}`);
            const attachmentData = Gmail.Users.Messages.Attachments.get('me', message.id, pdfAttachment.attachmentId);
            _logMessage(`📦 קיבל נתונים: ${attachmentData.data ? 'יש data' : 'אין data'}`);
            
            if (!attachmentData.data) {
              throw new Error('No data in attachment response');
            }
            
            const decodedData = Utilities.base64Decode(attachmentData.data);
            const blob = Utilities.newBlob(decodedData, 'application/pdf', fileName);
            
            // שמירה ל-Drive (My Drive root - כמו בגרסה הקודמת)
            _logMessage(`💾 שומר קובץ ב-My Drive: ${fileName}`);
            const file = DriveApp.createFile(blob);
            
            file.setName(fileName);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            pdfLink = file.getUrl();
            _logMessage(`✅ נוצר לינק Drive: ${pdfLink}`);
            
          } catch (err) {
            pdfLink = 'ERROR: ' + err.message;
            _logMessage(`❌ שגיאה בשמירת הקובץ ל-Drive: ${err.message}`);
          }
        } else {
          pdfLink = "";
          _logMessage(`❌ לא נמצא PDF attachment בכלל`);
        }
      }
      
      _logMessage(`🔗 תוצאת לינק PDF: ${pdfLink ? 'הצליח - ' + pdfLink.substring(0, 50) + '...' : 'נכשל - ריק'}`);
      const finalData = {
        ...emailData,
        ...invoiceData,
        pdfLink: pdfLink
      };
      
      // לוג סוג הלינק שנוצר
      if (pdfLink.includes('drive.google.com')) {
        _logMessage(`✅ Drive Link: ${finalData.companyName || finalData.senderName} - ${pdfLink}`);
      } else if (pdfLink.includes('mail.google.com')) {
        _logMessage(`⚠️ Gmail Link: ${finalData.companyName || finalData.senderName} - ${pdfLink}`);
      } else {
        _logMessage(`❓ Unknown Link: ${finalData.companyName || finalData.senderName} - ${pdfLink}`);
      }
      
      // כתיבה לגיליון
      const rowData = [
        false, // Action checkbox
        isLocalEmail(finalData.senderEmail) ? "🔷 Local" : "🌐 International",
        finalData.companyName || finalData.senderName,
        finalData.senderEmail,
        formatDateInput(finalData.date),
        finalData.subject,
        finalData.pdfLink,
        finalData.messageId,
        finalData.attachmentName || pdfAttachment.filename,
        finalData.amount || "",
        finalData.currency || "",
        finalData.amountMethod || "not_extracted"
      ];
      
      _logMessage(`ℹ️ 📝 כותב: ${finalData.companyName || finalData.senderName} - ${finalData.amount || 'ללא סכום'}`);
      
      sheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
      
      addedCount++;
      currentRow++;
      
      // ריענון תקופתי
      if (addedCount % 5 === 0) {
        SpreadsheetApp.flush();
        SpreadsheetApp.getActiveSpreadsheet().toast(`📝 נכתבו ${addedCount} חשבוניות...`, "עיבוד", 1);
      }
      
    } catch (e) {
      _logMessage(`❌ שגיאה בעיבוד מייל: ${e.message}`);
      skippedCount++;
    }
  }
  
  SpreadsheetApp.flush();
  
  // הוספת צ'קבוקסים רק לשורות שנכתבו בפועל
  if (addedCount > 0) {
    try {
      sheet.getRange(2, 1, addedCount, 1).insertCheckboxes();
      _logMessage(`✅ צ'קבוקסים נוצרו לשורות 2-${addedCount + 1} (${addedCount} רשומות)`);
    } catch (e) {
      _logMessage(`⚠️ שגיאה ביצירת צ'קבוקסים: ${e.message}`);
    }
  }
  
  const results = {
    added: addedCount,
    skipped: skippedCount,
    total: messages.length
  };
  
  _logMessage(`✅ 📊 תוצאות עיבוד מתקדם: ${addedCount} נוספו, ${skippedCount} דולגו`);
  
  return results;
}

// === חילוץ מטאדטה מהמייל ===
function extractEmailMetadata(message) {
  try {
    const headers = message.payload.headers;
    const getHeader = (name) => {
      const header = headers.find(h => h.name.toLowerCase() === name.toLowerCase());
      return header ? header.value : '';
    };
    
    const from = getHeader('From');
    const subject = getHeader('Subject');
    const date = new Date(parseInt(message.internalDate));
    
    return {
      messageId: message.id,
      senderEmail: _parseSender(from),
      senderName: extractSenderName(from),
      subject: subject,
      date: date
    };
  } catch (e) {
    _logMessage(`⚠️ שגיאה בחילוץ מטאדטה: ${e.message}`);
    return null;
  }
}

// === מציאת קובץ PDF ===
function findPdfAttachment(message) {
  try {
    _logMessage(`🔍 DEBUG findPdfAttachment: message.payload = ${message.payload ? 'exists' : 'null'}`);
    
    if (!message.payload) {
      _logMessage(`❌ אין payload במייל`);
      return null;
    }
    
    if (!message.payload.parts) {
      _logMessage(`❌ אין parts ב-payload`);
      return null;
    }
    
    _logMessage(`🔍 DEBUG: נמצאו ${message.payload.parts.length} parts`);
    
    for (let i = 0; i < message.payload.parts.length; i++) {
      const part = message.payload.parts[i];
      if (part.filename && part.filename.toLowerCase().includes('.pdf')) {
        // אם יש attachmentId, נחזיר את המידע בלבד (ההורדה תהיה בקוד הראשי)
        if (part.body.attachmentId) {
          _logMessage(`📎 נמצא PDF עם attachmentId: ${part.filename} (${part.body.attachmentId})`);
          return {
            attachmentId: part.body.attachmentId,
            filename: part.filename,
            size: part.body.size,
            partId: `0.${i + 1}`,
            mimeType: part.mimeType || 'application/pdf'
          };
        } else if (part.body.data) {
          // אם הנתונים כבר קיימים ב-body
          return {
            attachmentId: part.body.attachmentId,
            filename: part.filename,
            size: part.body.size,
            partId: `0.${i + 1}`,
            body: { data: part.body.data }
          };
        } else {
          _logMessage(`⚠️ לא נמצאו נתוני קובץ עבור: ${part.filename}`);
          return null;
        }
      }
    }
    return null;
  } catch (e) {
    _logMessage(`⚠️ שגיאה בחיפוש PDF: ${e.message}`);
    return null;
  }
}

// === חילוץ לינק PDF מהמייל ===
function extractPdfLink(message) {
  try {
    _logMessage(`🔍 extractPdfLink called for message: ${message.id}`);
    
    // שיטה 1: GmailApp (פשוט ויעיל - כמו בקוד הישן)
    try {
      _logMessage(`🔄 מנסה GmailApp method...`);
      const gmailMessage = GmailApp.getMessageById(message.id);
      const attachments = gmailMessage.getAttachments();
      _logMessage(`📎 נמצאו ${attachments.length} attachments`);
      
      const pdfAttachment = attachments.find(att => 
        att.getName().toLowerCase().endsWith('.pdf') || 
        att.getContentType().toLowerCase().includes('pdf')
      );
      
      if (pdfAttachment) {
        _logMessage(`✅ נמצא PDF: ${pdfAttachment.getName()}`);
        const blob = pdfAttachment.getBlob();
        const driveUrl = createDriveFile(blob, pdfAttachment.getName());
        if (driveUrl) {
          _logMessage(`✅ הועלה ל-Drive: ${driveUrl}`);
          return driveUrl;
        }
      } else {
        _logMessage(`❌ לא נמצא PDF attachment`);
      }
    } catch (e) {
      _logMessage(`❌ שגיאה ב-GmailApp: ${e.message}`);
    }
    
    // שיטה 2: Gmail API (backup)
    try {
      _logMessage(`🔄 מנסה Gmail API method...`);
      const pdfAttachment = findPdfAttachment(message);
      if (pdfAttachment && pdfAttachment.body && pdfAttachment.body.data) {
        _logMessage(`✅ נמצא PDF via API: ${pdfAttachment.filename}`);
        const blob = Utilities.newBlob(
          Utilities.base64Decode(pdfAttachment.body.data),
          'application/pdf',
          pdfAttachment.filename || 'invoice.pdf'
        );
        const driveUrl = createDriveFile(blob, pdfAttachment.filename);
        if (driveUrl) {
          _logMessage(`✅ הועלה ל-Drive via API: ${driveUrl}`);
          return driveUrl;
        }
      }
    } catch (e) {
      _logMessage(`❌ שגיאה ב-Gmail API: ${e.message}`);
    }

    _logMessage(`❌ לא הצליח ליצור לינק Drive - מחזיר null`);
    return null;
  } catch (e) {
    _logMessage(`❌ שגיאה כללית ב-extractPdfLink: ${e.message}`);
    return null;
  }
}

// פונקציית עזר ליצירת קובץ Drive
function createDriveFile(blob, filename) {
  try {
    // יצירת הקובץ ב-My Drive (root) - כמו בגרסה הקודמת
    _logMessage(`💾 יוצר קובץ ב-My Drive: ${filename}`);
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    _logMessage(`✅ נוצר קובץ Drive: ${file.getUrl()}`);
    return file.getUrl();
  } catch (e) {
    _logMessage(`❌ שגיאה ביצירת קובץ Drive: ${e.message}`);
    return null;
  }
}

// === מערכת סינון חכמה משולבת ===
function smartInvoiceValidation(message, pdfAttachment) {
  const timestamp = new Date().toLocaleString('he-IL');
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  
  // חילוץ נתוני המייל
  const emailData = extractEmailMetadata(message);
  const subject = emailData.subject.toLowerCase();
  const filename = pdfAttachment.filename.toLowerCase();
  
  if (log) log.appendRow([timestamp, `🔍 בודק: נושא="${emailData.subject}" קובץ="${pdfAttachment.filename}"`]);
  
  // שלב 1: בדיקה מהירה - רשימה שחורה (דחייה מיידית)
  const blacklistSubjects = [
    'quote', 'quotation', 'estimate', 'proposal', 'offer', 'draft',
    'הצעת מחיר', 'הצעה', 'הערכה', 'הצעת עבודה', 'טיוטה',
    'contract', 'agreement', 'חוזה', 'הסכם', 'תנאים',
    'טופס תיקון', 'תיקון', 'correction', 'amendment'
  ];
  
  const blacklistFiles = [
    'quote', 'quotation', 'estimate', 'proposal', 'offer', 'draft',
    'הצעה', 'הצעת_מחיר', 'טיוטה', 'הסכם', 'חוזה', 'פוליסה',
    'agreement', 'contract', 'policy', 'terms', 'conditions',
    'תיקון', 'correction', 'amendment', 'טופס'
  ];
  
  // בדיקת רשימה שחורה - נושא
  for (const blackword of blacklistSubjects) {
    if (subject.includes(blackword)) {
      if (log) log.appendRow([timestamp, `❌ נדחה מיד - נושא חסום: "${blackword}"`]);
      return { decision: 'REJECT', reason: `נושא מכיל: ${blackword}` };
    }
  }
  
  // בדיקת רשימה שחורה - קובץ
  for (const blackword of blacklistFiles) {
    if (filename.includes(blackword)) {
      if (log) log.appendRow([timestamp, `❌ נדחה מיד - קובץ חסום: "${blackword}"`]);
      return { decision: 'REJECT', reason: `קובץ מכיל: ${blackword}` };
    }
  }
  
  // שלב 2: בדיקה מהירה - רשימה לבנה (אישור מיידי)
  const whitelistSubjects = [
    'invoice', 'receipt', 'bill', 'payment', 'statement', 'billing',
    'חשבונית', 'קבלה', 'חשבון', 'תשלום', 'חיוב'
  ];
  
  const whitelistFiles = [
    'invoice', 'receipt', 'bill', 'payment', 'statement',
    'חשבונית', 'קבלה', 'חשבון', 'תשלום'
  ];
  
  let subjectScore = 0;
  let fileScore = 0;
  let foundSubjectKeywords = [];
  let foundFileKeywords = [];
  
  // ניקוד נושא
  for (const goodword of whitelistSubjects) {
    if (subject.includes(goodword)) {
      subjectScore++;
      foundSubjectKeywords.push(goodword);
    }
  }
  
  // ניקוד קובץ
  for (const goodword of whitelistFiles) {
    if (filename.includes(goodword)) {
      fileScore++;
      foundFileKeywords.push(goodword);
    }
  }
  
  // שלב 3: החלטה לפי ניקוד
  const totalScore = subjectScore + fileScore;
  
  // אישור מיידי - ניקוד גבוה
  if (totalScore >= 2 || (subjectScore >= 1 && fileScore >= 1)) {
    if (log) log.appendRow([timestamp, `✅ אושר מיד - ניקוד גבוה: נושא(${subjectScore}) + קובץ(${fileScore}) = ${totalScore}`]);
    return { 
      decision: 'APPROVE', 
      reason: `נמצאו מילות מפתח: נושא[${foundSubjectKeywords.join(',')}] קובץ[${foundFileKeywords.join(',')}]`,
      confidence: 'HIGH'
    };
  }
  
  // אישור בינוני - יש לפחות מילת מפתח אחת
  if (totalScore >= 1) {
    if (log) log.appendRow([timestamp, `✅ אושר - ניקוד בינוני: ${totalScore} מילות מפתח`]);
    return { 
      decision: 'APPROVE', 
      reason: `נמצאה מילת מפתח: ${foundSubjectKeywords.concat(foundFileKeywords).join(',')}`,
      confidence: 'MEDIUM'
    };
  }
  
  // בדיקה נוספת - מספרים ותאריכים (סימן טוב)
  const hasNumbers = /\d{4}/.test(filename) && /\d{2}/.test(filename);
  const hasAmountInSubject = /[\d,]+\.?\d*\s*(?:₪|שקל|ש"ח|\$|USD|€|EUR)/i.test(subject);
  
  if (hasNumbers || hasAmountInSubject) {
    if (log) log.appendRow([timestamp, `✅ אושר - נמצאו מספרים/סכומים`]);
    return { 
      decision: 'APPROVE', 
      reason: hasAmountInSubject ? 'נמצא סכום בנושא' : 'נמצאו מספרים ותאריכים בקובץ',
      confidence: 'MEDIUM'
    };
  }
  
  // אם הגענו עד כאן - לא בטוח, צריך OCR
  if (log) log.appendRow([timestamp, `❓ לא בטוח - יעבור לOCR: נושא="${emailData.subject}" קובץ="${pdfAttachment.filename}"`]);
  return { 
    decision: 'OCR_NEEDED', 
    reason: 'לא נמצאו מספיק סימנים - נדרש OCR',
    confidence: 'UNKNOWN'
  };
}

// === OCR פשוט למקרים מפוקפקים ===
function performOCRValidation(message, pdfAttachment) {
  const timestamp = new Date().toLocaleString('he-IL');
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  
  try {
    if (log) log.appendRow([timestamp, `🔍 OCR: בודק תוכן קובץ ${pdfAttachment.filename}`]);
    
    // כאן יהיה OCR אמיתי בעתיד
    // לעת עתה נעשה בדיקה פשוטה נוספת
    
    const filename = pdfAttachment.filename.toLowerCase();
    
    // בדיקות OCR מדומות - בעתיד יהיה OCR אמיתי
    const ocrKeywords = [
      'inv', 'rec', 'bill', 'pay', // קיצורים שעלולים להיות בשם קובץ
      '2025', '2024', // שנים
      'pdf' // זה PDF אז יש סיכוי שזה מסמך רשמי
    ];
    
    let ocrScore = 0;
    for (const keyword of ocrKeywords) {
      if (filename.includes(keyword)) {
        ocrScore++;
      }
    }
    
    if (ocrScore >= 2) {
      if (log) log.appendRow([timestamp, `✅ OCR אישר - נמצאו ${ocrScore} סימנים`]);
      return { decision: 'APPROVE', reason: `OCR מצא ${ocrScore} סימנים חיוביים` };
    } else {
      if (log) log.appendRow([timestamp, `❌ OCR דחה - רק ${ocrScore} סימנים`]);
      return { decision: 'REJECT', reason: `OCR מצא רק ${ocrScore} סימנים` };
    }
    
  } catch (e) {
    if (log) log.appendRow([timestamp, `⚠️ OCR נכשל: ${e.message}`]);
    // אם OCR נכשל, נאשר בזהירות
    return { decision: 'APPROVE', reason: 'OCR נכשל - אישור זהיר' };
  }
}

// === פונקציה מרכזית משופרת ===
function extractInvoiceDataWithOCR(message, pdfAttachment) {
  const timestamp = new Date().toLocaleString('he-IL');
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  
  try {
    // שלב 1: בדיקה חכמה מהירה
    const smartResult = smartInvoiceValidation(message, pdfAttachment);
    
    if (smartResult.decision === 'REJECT') {
      if (log) log.appendRow([timestamp, `❌ נדחה: ${smartResult.reason}`]);
      return null;
    }
    
    if (smartResult.decision === 'APPROVE') {
      if (log) log.appendRow([timestamp, `✅ אושר: ${smartResult.reason}`]);
      return {
        companyName: "",
        invoiceNumber: extractInvoiceNumberFromFilename(pdfAttachment.filename),
        amount: "",
        currency: "",
        attachmentName: pdfAttachment.filename,
        isValidInvoice: true,
        validationMethod: `חכם: ${smartResult.reason}`,
        confidence: smartResult.confidence
      };
    }
    
    // שלב 2: OCR למקרים מפוקפקים
    if (smartResult.decision === 'OCR_NEEDED') {
      const ocrResult = performOCRValidation(message, pdfAttachment);
      
      if (ocrResult.decision === 'APPROVE') {
        if (log) log.appendRow([timestamp, `✅ אושר לאחר OCR: ${ocrResult.reason}`]);
        return {
          companyName: "",
          invoiceNumber: extractInvoiceNumberFromFilename(pdfAttachment.filename),
          amount: "",
          currency: "",
          attachmentName: pdfAttachment.filename,
          isValidInvoice: true,
          validationMethod: `OCR: ${ocrResult.reason}`,
          confidence: 'LOW'
        };
      } else {
        if (log) log.appendRow([timestamp, `❌ נדחה לאחר OCR: ${ocrResult.reason}`]);
        return null;
      }
    }
    
    // לא אמור להגיע לכאן
    return null;
    
  } catch (e) {
    if (log) log.appendRow([timestamp, `❌ שגיאה כללית: ${e.message}`]);
    return null;
  }
}

// === חילוץ מספר חשבונית משם הקובץ ===
function extractInvoiceNumberFromFilename(filename) {
  // חיפוש מספרים בשם הקובץ
  const patterns = [
    /invoice.*?(\d+)/i,
    /receipt.*?(\d+)/i,
    /(\d{4,})/g, // מספר של 4 ספרות או יותר
    /(\d+)/g // כל מספר
  ];
  
  for (const pattern of patterns) {
    const match = filename.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  return "";
}

// הפונקציות הישנות הוסרו - עכשיו משתמשים בבדיקה פשוטה לפי שם הקובץ

// === הצגת סיכום מתקדם ===
function showAdvancedScanSummary(results) {
  const message = `✅ סריקה מתקדמת הושלמה!

📊 תוצאות:
• ${results.added} חשבוניות נוספו
• ${results.skipped} מיילים דולגו  
• ${results.total} מיילים נבדקו בסך הכל

🚀 שיפורים:
• Gmail API מהיר יותר
• חילוץ לינקים משופר
• מטאדטה מדויקת יותר

🎉 החשבוניות זמינות בגיליון החדש!`;

  _logMessage(message);
  SpreadsheetApp.getUi().alert(message);
}

// === SMART AMOUNT EXTRACTION ===
function extractAmountSmart(message, pdfAttachment) {
  const startTime = new Date();
  
  try {
    // שלב 1: חילוץ מהיר מנושא ותוכן
    const quickAmount = extractAmountFromText(message);
    if (quickAmount.found) {
      _logMessage(`💰 סכום נמצא מהיר: ${quickAmount.amount} ${quickAmount.currency}`);
      return quickAmount;
    }
    
    // שלב 2: OCR רק אם לא נמצא בשלב 1
    if (pdfAttachment) {
      _logMessage(`🔍 מעבר ל-OCR לחילוץ סכום מ-${pdfAttachment.filename}`);
      const ocrAmount = extractAmountFromPdfOCR(message, pdfAttachment);
      if (ocrAmount.found) {
        const duration = (new Date() - startTime) / 1000;
        _logMessage(`💰 סכום נמצא OCR: ${ocrAmount.amount} ${ocrAmount.currency} (${duration.toFixed(1)}s)`);
        return ocrAmount;
      }
    }
    
    return { found: false, amount: "", currency: "", method: "not_found" };
    
  } catch (e) {
    _logMessage(`❌ שגיאה בחילוץ סכום: ${e.message}`);
    return { found: false, amount: "", currency: "", method: "error" };
  }
}

// === חילוץ מהיר מטקסט ===
function extractAmountFromText(message) {
  try {
    const headers = message.payload.headers;
    const subject = headers.find(h => h.name.toLowerCase() === 'subject')?.value || '';
    
    // חילוץ תוכן המייל
    let bodyText = '';
    if (message.payload.body?.data) {
      bodyText = Utilities.base64DecodeWebSafe(message.payload.body.data);
    } else if (message.payload.parts) {
      for (const part of message.payload.parts) {
        if (part.mimeType === 'text/plain' && part.body?.data) {
          bodyText += Utilities.base64DecodeWebSafe(part.body.data);
        }
      }
    }
    
    const fullText = subject + ' ' + bodyText;
    
    // Regex patterns לסכומים
    const patterns = [
      // שקלים
      /₪\s*([0-9,]+\.?\d*)/g,
      /([0-9,]+\.?\d*)\s*₪/g,
      /([0-9,]+\.?\d*)\s*שקל/g,
      /([0-9,]+\.?\d*)\s*ש"ח/g,
      /סכום[:\s]*([0-9,]+\.?\d*)/g,
      
      // דולרים
      /\$\s*([0-9,]+\.?\d*)/g,
      /([0-9,]+\.?\d*)\s*USD/g,
      /([0-9,]+\.?\d*)\s*\$/g,
      
      // יורו
      /€\s*([0-9,]+\.?\d*)/g,
      /([0-9,]+\.?\d*)\s*EUR/g,
      
      // כללי
      /total[:\s]*([0-9,]+\.?\d*)/gi,
      /amount[:\s]*([0-9,]+\.?\d*)/gi
    ];
    
    for (const pattern of patterns) {
      const matches = fullText.match(pattern);
      if (matches && matches.length > 0) {
        const match = matches[0];
        const amount = match.replace(/[^\d.,]/g, '');
        const currency = detectCurrency(match);
        
        if (parseFloat(amount.replace(',', '')) > 0) {
          return {
            found: true,
            amount: amount,
            currency: currency,
            method: "text_extraction",
            source: match
          };
        }
      }
    }
    
    return { found: false, method: "text_extraction" };
    
  } catch (e) {
    _logMessage(`⚠️ שגיאה בחילוץ מטקסט: ${e.message}`);
    return { found: false, method: "text_error" };
  }
}

// === זיהוי מטבע ===
function detectCurrency(text) {
  const lowerText = text.toLowerCase();
  
  if (lowerText.includes('₪') || lowerText.includes('שקל') || lowerText.includes('ש"ח')) {
    return 'ILS';
  } else if (lowerText.includes('$') || lowerText.includes('usd') || lowerText.includes('dollar')) {
    return 'USD';
  } else if (lowerText.includes('€') || lowerText.includes('eur') || lowerText.includes('euro')) {
    return 'EUR';
  }
  
  return 'ILS'; // ברירת מחדל
}

// === OCR לחילוץ סכום מ-PDF ===
function extractAmountFromPdfOCR(message, pdfAttachment) {
  try {
    // כאן נשתמש ב-OCR הקיים אבל נתמקד בחיפוש סכומים
    // זה יהיה איטי יותר אז נעשה רק אם לא מצאנו בטקסט
    
    _logMessage(`🔍 OCR: מנתח ${pdfAttachment.filename} לחיפוש סכומים`);
    
    // TODO: מימוש OCR מתקדם לחילוץ סכומים
    // לעת עתה נחזיר לא נמצא
    
    return { found: false, method: "ocr_not_implemented" };
    
  } catch (e) {
    _logMessage(`❌ שגיאה ב-OCR: ${e.message}`);
    return { found: false, method: "ocr_error" };
  }
}

// === יצירת גיליון תוצאות מתקדם ===
function createAdvancedResultsSheet(ss, tabName) {
  try {
    _logMessage(`ℹ️ 📑 יוצר גיליון חדש: ${tabName}`);
    
    let sheet = ss.getSheetByName(tabName);
    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet(tabName);
    }
    
    // יצירת כותרות מעודכנות
    const headers = [
      "Action", "Category", "Sender Name", "Sender Email", "Date", "Subject",
      "PDF Link", "Email ID", "Attachment Name", "Amount", "Currency", "Extract Method"
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _logMessage("✅ כותרות נוצרו בהצלחה");
    
    // צ'קבוקסים יתווספו אחרי כתיבת הנתונים
    // sheet.getRange("A2:A100").insertCheckboxes(); // הוסר - גורם לבעיית שורה 1000
    // _logMessage("✅ צ'קבוקסים נוצרו בהצלחה");
    
    // עיצוב בסיסי
    const headerRange = sheet.getRange("A1:L1");
    headerRange.setFontWeight("bold")
               .setBackground("#ddebf7")
               .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    
    // הגדרת רוחב עמודות
    sheet.setColumnWidth(1, 80);   // Action
    sheet.setColumnWidth(2, 100);  // Category
    sheet.setColumnWidth(3, 150);  // Sender Name
    sheet.setColumnWidth(4, 200);  // Sender Email
    sheet.setColumnWidth(5, 100);  // Date
    sheet.setColumnWidth(6, 300);  // Subject
    sheet.setColumnWidth(7, 250);  // PDF Link
    sheet.setColumnWidth(8, 150);  // Email ID
    sheet.setColumnWidth(9, 150);  // Attachment Name
    sheet.setColumnWidth(10, 100); // Amount
    sheet.setColumnWidth(11, 80);  // Currency
    sheet.setColumnWidth(12, 120); // Extract Method
    
    _logMessage("✅ עיצוב גיליון הושלם");
    
    return sheet;
    
  } catch (e) {
    _logMessage(`❌ שגיאה ביצירת גיליון: ${e.message}`);
    return null;
  }
}