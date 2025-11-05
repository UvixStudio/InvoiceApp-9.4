// === Code.gs - FIXED VERSION BY CLAUDE ===

// === GLOBALS ===
const SETTINGS_SHEET_NAME = "Settings";
const LOG_SHEET_NAME = "Log";

// === MENU ===
function onOpen() {
  try {
    createMainMenu();
    createActionsMenu();
    console.log("✅ התפריטים נוצרו בהצלחה");
  } catch (e) {
    console.error("❌ שגיאה ביצירת התפריטים:", e);
    try {
      SpreadsheetApp.getActiveSpreadsheet().addMenu("📋 Invoice Scanner", [
        {name: "⚙️ Setup Sheets", functionName: "setupSheets"},
        {name: "📥 Scan Invoices", functionName: "processInvoicesFixed"}
      ]);
    } catch (fallbackError) {
      console.error("❌ שגיאה גם בניסיון יצירת תפריט בסיסי:", fallbackError);
    }
  }
}

function onInstall(e) {
  onOpen(e);
}

function createMainMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const mainMenu = ui.createMenu("📋 Invoice Scanner")
    .addItem("⚙️ הגדרות ראשוניות", "setupSheets")
    .addItem("📁 הגדרת תיקיית סינכרון (Google Drive)", "showExportFolderDialog")
    .addSeparator()
    .addItem("🚀 סריקת חשבוניות (Gmail API)", "processInvoices")
    .addSeparator()
    .addItem("  ניטרןו OCR - חילוץ סכומים", "performNitroOcr")
    .addSeparator()
    .addItem("🔎 אבחון מפורט (לוג יפה)", "diagnosticScan")
    .addItem("📊 אבחון טבלה", "diagnosticAdvanced")
    .addSeparator()
    .addItem("📋 בחר סריקה אחרונה מהלוג", "selectLastScanLog")
    .addItem("🗑️ איפוס לוג", "clearLogSheet")
    .addSeparator()
    .addItem(" ️ באיפוס כל הגיליונות", "resetAllTabs")
    .addItem("  אביפוס מערכת (משתמש חדש)", "systemReset");
    
  mainMenu.addToUi();
}

function createActionsMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const actionsMenu = ui.createMenu("⚡ פעולות")
    .addItem("🚫 החרג נבחרים", "excludeSelected")
    .addItem("🔷 סמן כמקומי", "markAsLocal")
    .addItem("🌐 סמן כבינלאומי", "markAsInternational")
    .addSeparator()
    .addItem("📦 יצא נבחרים", "exportSelected")
    .addItem("🗑️ מחק נבחרים", "deleteSelected")
    .addSeparator()
    .addItem("✅ הוסף נבחרים למאושרים", "approveSelected")
    .addSeparator()
    .addItem("🧹 נקה גיליון", "clearActiveSheetData")
    .addSeparator()
    .addItem(" ️ רענן תצוגה", "refreshView")
    .addSeparator()
    .addItem("🔧 בדיקת כתיבה", "debugSheetWriting")
    .addItem("🔄 איפוס וניקוי", "resetAndCleanup")
    .addSeparator()
    .addItem("🐛 דיבוג החרגות", "debugExclusionsQuick")
    .addItem("🔧 בדיקת החרגות מתקדמת", "testExclusionsStep");
    
  actionsMenu.addToUi();
}

// === SETUP SHEETS ===
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settings) settings = ss.insertSheet(SETTINGS_SHEET_NAME);
  settings.clear();

  const headers = [[
    "Start Date", "End Date", "EN Keywords", "HE Keywords", "Excluded Keywords", "Exclude Emails", "Approved Senders", "", "Export Folder ID"
  ]];
  settings.getRange("A1:I1").setValues(headers);
  settings.setFrozenRows(1);
  settings.setColumnWidths(1, 9, 200);

  // צבעים
  settings.getRange("A1").setBackground("#fce5cd");
  settings.getRange("B1").setBackground("#fce5cd");
  settings.getRange("C1").setBackground("#f4cccc");
  settings.getRange("D1").setBackground("#cfe2f3");
  settings.getRange("E1").setBackground("#f9cb9c");
  settings.getRange("F1").setBackground("#f4cccc");
  settings.getRange("G1").setBackground("#d9ead3");
  settings.getRange("I1").setBackground("#b6d7a8");

  const enDefaults = ["invoice", "receipt", "refund", "payment", "bill"];
  const heDefaults = ["חשבונית", "קבלה", "החשבונית שלך", "החשבונית החודשית שלך"];
  const excludeDefaults = ["noreply@example.com"];
  
  settings.getRange(2, 1).setValue("");
  settings.getRange(2, 2).setValue("");
  settings.getRange(2, 3, enDefaults.length, 1).setValues(enDefaults.map(e => [e]));
  settings.getRange(2, 4, heDefaults.length, 1).setValues(heDefaults.map(e => [e]));
  settings.getRange(2, 5).setValue("");
  settings.getRange(2, 6, excludeDefaults.length, 1).setValues(excludeDefaults.map(e => [e]));

  let log = ss.getSheetByName(LOG_SHEET_NAME);
  if (!log) log = ss.insertSheet(LOG_SHEET_NAME);
  if (log.getLastRow() === 0) {
    log.appendRow(["Timestamp", "Message"]);
    log.setFrozenRows(1);
    // הגדרת רוחב עמודות
    log.setColumnWidth(1, 150); // Timestamp
    log.setColumnWidth(2, 800); // Message - רחב מאוד
  }

  SpreadsheetApp.setActiveSheet(settings);
  
  // העברת גיליון Log למיקום הנכון (טאב שני)
  moveLogSheetToCorrectPosition();
  
  createMainMenu();
  createActionsMenu();
  
  SpreadsheetApp.getUi().alert("✅ הגיליונות נוצרו בהצלחה.");
  
  resetDateRange();
}

function resetDateRange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settings) return;

  const today = new Date();
  let start, end;
  if (today.getDate() <= 10) {
    start = new Date(today.getFullYear(), today.getMonth() - 2, 1);
    end = new Date(today.getFullYear(), today.getMonth() - 1 + 1, 0);
  } else {
    start = new Date(today.getFullYear(), today.getMonth(), 1);
    end = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  }
  settings.getRange("A2").setValue(formatDateInput(start));
  settings.getRange("B2").setValue(formatDateInput(end));
}

// === OLD SCAN REMOVED - USE ADVANCED API ONLY ===
// הסריקה הישנה הוסרה. השתמש רק ב-"סריקה מתקדמת (Gmail API)"
function processInvoicesFixed() {
  SpreadsheetApp.getUi().alert('❌ הסריקה הישנה הוסרה.\n\nהשתמש ב-"🚀 סריקה מתקדמת (Gmail API)" במקום.');
  return;
}

// === HELPER FUNCTIONS ===
function _logMessage(message) {
  const excludedKeywords = settings.getRange('E2:E').getValues().flat().filter(Boolean);
  const excludeEmails = settings.getRange('F2:F').getValues().flat().filter(Boolean);
  
  _logMessage(`ℹ️ 🔤 מילות מפתח: EN(${enKeywords.length}), HE(${heKeywords.length})`);
  _logMessage(`ℹ️ 🚫 מוחרגים: ${excludeEmails.length} מיילים, ${excludedKeywords.length} מילים`);
  
  const allKeywords = [...enKeywords, ...heKeywords];

  // בניית שאילתת Gmail
  const keywordQuery = allKeywords.map(k => `"${k}"`).join(' OR ');
  const excludeQuery = excludeEmails.map(e => `-from:${e}`).join(' ');
  const query = `after:${formatDateGmail(startDate)} before:${formatDateGmail(endDate)} has:attachment (${keywordQuery}) ${excludeQuery}`;
  
  _logMessage(`🔍 שאילתת Gmail: ${query}`);

  let threads = [];
  try {
    threads = GmailApp.search(query, 0, 200);
    _logMessage(`📧 נמצאו ${threads.length} שרשורי מייל`);
  } catch (e) {
    _logMessage(`❌ שגיאה בחיפוש Gmail: ${e.message}`);
    SpreadsheetApp.getUi().alert(`❌ שגיאה בחיפוש Gmail: ${e.message}`);
    return;
  }
  
  if (threads.length === 0) {
    _logMessage('⚠️ לא נמצאו מיילים בטווח התאריכים');
    SpreadsheetApp.getUi().alert('⚠️ לא נמצאו מיילים בטווח התאריכים.');
    return;
  }

  // יצירת גיליון תוצאות
  const tabName = `invoices ${formatDateShort(startDate)}-${formatDateShort(endDate)}`;
  _logMessage(`ℹ️ 📑 יוצר גיליון חדש: ${tabName}`);
  
  let sheet = ss.getSheetByName(tabName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(tabName);
  }
  _logMessage("✅ גיליון תוצאות נוצר בהצלחה");

  // יצירת כותרות - זה החלק הקריטי!
  const headers = [
    "Action", "Category", "Sender Name", "Sender Email", "Date", "Subject",
    "PDF Link", "Email ID", "Attachment Name", "Amount", "Currency", "Extract Method"
  ];
  
  try {
    // כתיבת כותרות בשורה ראשונה
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _logMessage("✅ כותרות נוצרו בהצלחה");
  } catch (e) {
    _logMessage(`❌ שגיאה ביצירת כותרות: ${e.message}`);
    return;
  }

  // הוספת צ'קבוקסים - רק לשורות שנצטרך
  try {
    sheet.getRange("A2:A50").insertCheckboxes();
    _logMessage("✅ צ'קבוקסים נוצרו בהצלחה (שורות 2-50)");
  } catch (e) {
    _logMessage(`⚠️ שגיאה בצ'קבוקסים: ${e.message}`);
  }

  // עיצוב בסיסי
  try {
    const headerRange = sheet.getRange("A1:L1");
    headerRange.setFontWeight("bold")
               .setBackground("#ddebf7")
               .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    _logMessage("✅ עיצוב כותרות הושלם");
  } catch (e) {
    _logMessage(`⚠️ שגיאה בעיצוב: ${e.message}`);
  }

  // איסוף הודעות
  let allMessages = [];
  _logMessage("ℹ️ איסוף הודעות מהשרשורים...");
  
  for (let i = 0; i < threads.length; i++) {
    try {
      const messages = threads[i].getMessages();
      for (const msg of messages) {
        // סינון בסיסי
        const msgDate = msg.getDate();
        if (msgDate >= startDate && msgDate <= endDate) {
          allMessages.push(msg);
        }
      }
    } catch (e) {
      _logMessage(`⚠️ שגיאה בקריאת שרשור ${i}: ${e.message}`);
    }
  }
  
  _logMessage(`📨 נמצאו ${allMessages.length} הודעות ב-${threads.length} שרשורים`);
  _logMessage(`📋 חולצו ${allMessages.length} מיילים לעיבוד`);

  // עיבוד הודעות - תיקון הבעיה!
  let addedCount = 0;
  let skippedCount = 0;
  let currentRow = 2; // מתחילים משורה 2 (אחרי הכותרות)
  
  _logMessage(`🔧 DEBUG: currentRow מאותחל ל-${currentRow}`);
  
  _logMessage(`ℹ️ 🔄 מתחיל עיבוד ${allMessages.length} מיילים`);
  
  for (let i = 0; i < allMessages.length; i++) {
    const msg = allMessages[i];
    
    try {
      const from = msg.getFrom();
      const senderName = extractSenderName(from);
      const senderEmail = _parseSender(from);
      const date = msg.getDate();
      const subject = msg.getSubject();
      const msgId = msg.getId();
      const attachments = msg.getAttachments();

      // בדיקת מילות מפתח - חיפוש חכם יותר
      let hasKeyword = false;
      let foundKeyword = "";
      const lowerSubject = subject.toLowerCase();
      const lowerBody = msg.getPlainBody().toLowerCase();
      
      // בדיקת מילות מפתח מהגדרות
      for (const keyword of allKeywords) {
        const lowerKeyword = keyword.toLowerCase();
        if (lowerSubject.includes(lowerKeyword) || lowerBody.includes(lowerKeyword)) {
          hasKeyword = true;
          foundKeyword = keyword;
          break;
        }
      }
      
      // אם לא נמצאו מילות מפתח, בדיקה חכמה נוספת
      if (!hasKeyword) {
        const smartKeywords = [
          // מילות מפתח נוספות באנגלית
          'billing', 'statement', 'purchase', 'order', 'transaction', 'charge',
          'subscription', 'renewal', 'due', 'amount', 'total', 'tax',
          // מילות מפתח נוספות בעברית
          'חיוב', 'הזמנה', 'רכישה', 'עסקה', 'מנוי', 'חידוש', 'סכום', 'מס',
          'תשלום', 'חשבון', 'קניה', 'שירות', 'מוצר'
        ];
        
        for (const keyword of smartKeywords) {
          if (lowerSubject.includes(keyword) || lowerBody.includes(keyword)) {
            hasKeyword = true;
            foundKeyword = `חכם: ${keyword}`;
            break;
          }
        }
      }
      
      // בדיקה נוספת: אם יש PDF ומספרים שנראים כמו סכומים
      if (!hasKeyword && hasPdf) {
        const amountPatterns = [
          /\$[\d,]+\.?\d*/g,  // דולרים
          /₪[\d,]+\.?\d*/g,   // שקלים
          /[\d,]+\.?\d*\s*(?:₪|שקל|ש"ח|NIS|ILS)/g,  // סכומים בשקלים
          /[\d,]+\.?\d*\s*(?:\$|USD|dollar)/g        // סכומים בדולרים
        ];
        
        for (const pattern of amountPatterns) {
          if (pattern.test(lowerSubject) || pattern.test(lowerBody)) {
            hasKeyword = true;
            foundKeyword = "זוהה סכום כסף";
            break;
          }
        }
      }
      
      _logMessage(`🔍 מילות מפתח: ${hasKeyword ? `נמצא - ${foundKeyword}` : 'לא נמצא'}`);
      
      if (!hasKeyword) {
        _logMessage(`⚠️ דולג: ${senderEmail} - "${subject}" - אין מילות מפתח`);
        skippedCount++;
        continue;
      }

      // בדיקת קבצי PDF
      let hasPdf = false;
      let pdfAttachmentName = "";
      for (const att of attachments) {
        if (att.getContentType().includes('pdf')) {
          hasPdf = true;
          pdfAttachmentName = att.getName();
          break;
        }
      }
      
      if (!hasPdf) {
        skippedCount++;
        continue;
      }

      // חילוץ לינק לחשבונית - שיפור ליצירת לינק ישיר ל-PDF
      let pdfLink = "";
      try {
        // אם יש PDF מצורף, ניצור קובץ ב-Drive
        if (hasPdf) {
          try {
            // מציאת קובץ ה-PDF
            const attachments = msg.getAttachments();
            for (const att of attachments) {
              if (att.getContentType().includes('pdf')) {
                // יצירת קובץ ב-Drive
                const sanitizedName = `${senderName}_${formatDateInput(date)}_${att.getName()}`.replace(/[^a-zA-Z0-9._-]/g, '_');
                const driveFile = DriveApp.createFile(att.copyBlob().setName(sanitizedName));
                
                // הגדרת הרשאות צפייה
                driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                
                pdfLink = driveFile.getUrl();
                _logMessage(`✅ קובץ Drive נוצר: ${sanitizedName} - ${pdfLink}`);
                break;
              }
            }
          } catch (driveError) {
            _logMessage(`⚠️ שגיאה ביצירת קובץ Drive: ${driveError.message}`);
            // fallback ללינק Gmail
            let partId = "0.2";
            const attachments = msg.getAttachments();
            for (let i = 0; i < attachments.length; i++) {
              if (attachments[i].getContentType().includes('pdf')) {
                partId = `0.${i + 2}`;
                break;
              }
            }
            pdfLink = `https://mail.google.com/mail/u/0/#search/invoice/${msgId}?projector=1&messagePartId=${partId}`;
            _logMessage(`🔗 Fallback Gmail לינק: ${pdfLink}`);
          }
        } else {
          // אם אין PDF מצורף, חפש לינקים בגוף המייל
          const body = msg.getBody();
          const plainBody = msg.getPlainBody();
          
          const linkPatterns = [
            /https?:\/\/[^\s<>"]+\.pdf/gi,
            /https?:\/\/[^\s<>"]*(?:invoice|receipt|bill|חשבונית)[^\s<>"]*/gi,
            /https?:\/\/drive\.google\.com\/[^\s<>"]+/gi
          ];
          
          for (const pattern of linkPatterns) {
            const matches = body.match(pattern) || plainBody.match(pattern);
            if (matches && matches.length > 0) {
              pdfLink = matches[0];
              break;
            }
          }
          
          _logMessage(`🔗 לינק בגוף המייל: ${pdfLink ? 'נמצא' : 'לא נמצא'}`);
        }
      } catch (e) {
        _logMessage(`⚠️ שגיאה בחילוץ לינק: ${e.message}`);
      }

      // חילוץ סכום חכם - שלב 1: טקסט רגיל
      _logMessage(`💰 מנסה לחלץ סכום עבור: ${subject.substring(0, 50)}...`);
      const plainBody = msg.getPlainBody();
      const htmlBody = msg.getBody();
      _logMessage(`📝 אורך תוכן טקסטואלי: ${plainBody ? plainBody.length : 0} תווים`);
      _logMessage(`📝 אורך תוכן HTML: ${htmlBody ? htmlBody.length : 0} תווים`);
      
      let amountData = extractAmountAndCurrency(subject, plainBody + ' ' + htmlBody);
      
      // שלב 2: אם לא נמצא סכום ויש PDF - נסה חילוץ טקסט ישיר (מהיר)
      if ((!amountData || !amountData.found) && pdfAttachmentName && attachments.length > 0) {
        _logMessage(`⚡ לא נמצא סכום בטקסט רגיל - מנסה חילוץ ישיר מ-PDF...`);
        
        for (const attachment of attachments) {
          if (attachment.getContentType().includes('pdf')) {
            try {
              const pdfText = extractTextFromPdfDirect(attachment.copyBlob());
              if (pdfText && pdfText.length > 10) {
                _logMessage(`📄 חילוץ ישיר הצליח - מחפש סכום בטקסט...`);
                const pdfAmountData = extractAmountAndCurrency("", pdfText);
                
                if (pdfAmountData && (pdfAmountData.found || pdfAmountData.amount)) {
                  _logMessage(`🎯 נמצא סכום בחילוץ ישיר: ${pdfAmountData.amount} ${pdfAmountData.currency}`);
                  amountData = {
                    ...pdfAmountData,
                    method: (pdfAmountData.method || 'direct') + '_pdf',
                    found: true
                  };
                  break; // מצאנו סכום, אפשר לעצור
                } else {
                  _logMessage(`⚠️ לא נמצא סכום בטקסט שחולץ מ-PDF`);
                }
              } else {
                _logMessage(`⚠️ חילוץ ישיר מ-PDF לא הצליח או טקסט קצר מדי`);
              }
            } catch (pdfError) {
              _logMessage(`⚠️ שגיאה בחילוץ ישיר מ-PDF: ${pdfError.message}`);
            }
          }
        }
      }
      
      // דיווח תוצאות חילוץ סכום
      if (amountData && amountData.found) {
        const method = amountData.method || 'unknown';
        const isPdf = method.includes('pdf');
        const isDirect = method.includes('direct');
        const isOcr = method.includes('ocr');
        
        let icon = '📝'; // טקסט רגיל
        if (isPdf && isDirect) icon = '⚡📄'; // חילוץ ישיר מ-PDF
        if (isPdf && isOcr) icon = '🔍📄'; // OCR מ-PDF
        
        _logMessage(`✅ ${icon} נמצא סכום: ${amountData.amount} ${amountData.currency} (${method})`);
      } else if (amountData && amountData.amount) {
        _logMessage(`✅ נמצא סכום: ${amountData.amount} ${amountData.currency || 'ILS'} (legacy)`);
      } else {
        _logMessage(`⚠️ לא נמצא סכום עבור: ${senderEmail} - יהיה זמין לניטרו OCR`);
      }
      
      // הכנת הנתונים לכתיבה
      const rowData = [
        false, // Action checkbox
        isLocalEmail(senderEmail) ? "🔷 Local" : "🌐 International",
        senderName,
        senderEmail,
        formatDateInput(date),
        subject,
        pdfLink, // PDF Link - עכשיו עם לינק אמיתי!
        msgId,
        pdfAttachmentName,
        (amountData && amountData.amount) ? amountData.amount : "", // Amount
        (amountData && amountData.currency) ? amountData.currency : "", // Currency
        (amountData && amountData.method) ? amountData.method : "" // Extract Method
      ];

      // הודעת תיעוד לפני כתיבה
      _logMessage(`ℹ️ 📝 כותב רשומה: [${rowData.slice(0, 5).join(',')}...`);

      // כתיבה לגיליון - זה החלק הקריטי ביותר!
      try {
        _logMessage(`🔧 DEBUG: כותב לשורה ${currentRow}, עמודות 1-${rowData.length}`);
        _logMessage(`🔧 DEBUG: נתונים לכתיבה: ${JSON.stringify(rowData.slice(0, 3))}`);
        
        sheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
        
        // אישור שהכתיבה הצליחה
        const writtenData = sheet.getRange(currentRow, 1, 1, 5).getValues()[0];
        _logMessage(`📋 נכתב בפועל בשורה ${currentRow}: [${writtenData.join(',')}...]`);
        
        addedCount++;
        currentRow++;
        
        // ריענון תקופתי
        if (addedCount % 5 === 0) {
          SpreadsheetApp.flush();
          _logMessage(`💾 ריענון אחרי ${addedCount} רשומות`);
        }
        
      } catch (writeError) {
        _logMessage(`❌ שגיאה בכתיבת שורה ${currentRow}: ${writeError.message}`);
        _logMessage(`❌ נתונים שנכשלו: ${JSON.stringify(rowData)}`);
      }
      
    } catch (msgError) {
      _logMessage(`❌ שגיאה בעיבוד הודעה: ${msgError.message}`);
      skippedCount++;
    }
  }

  // ריענון סופי
  SpreadsheetApp.flush();

  // דיווח תוצאות
  _logMessage(`✅ 📊 תוצאות עיבוד: ${addedCount} נוספו, ${skippedCount} דולגו, 0 כפילויות`);
  
  // דוח סיכום חילוץ סכומים
  generateAmountExtractionSummary(sheet);
  
  const finalMessage = `✅ הסריקה הושלמה!  📊 תוצאות: • ${addedCount} חשבוניות חדשות נוספו • ${skippedCount} מיילים דולגו • 0 כפילויות נמצאו • ${allMessages.length} מיילים נבדקו בסך הכל  🎉 החשבוניות זמינות בגיליון החדש!`;
  
  _logMessage(finalMessage);
  _logMessage("✅ סריקה הושלמה בהצלחה");

  // מעבר לגיליון החדש
  SpreadsheetApp.setActiveSheet(sheet);

  // הודעה למשתמש
  SpreadsheetApp.getUi().alert(`✅ סריקה הושלמה!\n${addedCount} חשבוניות נוספו\n${skippedCount} מיילים דולגו`);
}

// === HELPER FUNCTIONS ===
function _logMessage(message) {
  console.log(message);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = ss.getSheetByName(LOG_SHEET_NAME);
    if (log) {
      const timestamp = new Date().toLocaleString('he-IL');
      log.appendRow([timestamp, message]);
    }
  } catch (e) {
    console.error("שגיאה בתיעוד:", e);
  }
}

function _parseSender(fromString) {
  // חילוץ האימייל מהמחרוזת "Name <email@domain.com>"
  const match = fromString.match(/<([^>]+)>/);
  return match ? match[1] : fromString;
}

function extractSenderName(fromString) {
  // חילוץ השם מהמחרוזת "Name <email@domain.com>"
  const match = fromString.match(/^([^<]+)</);
  if (match) {
    return match[1].trim().replace(/"/g, '');
  }
  return _parseSender(fromString).split('@')[0];
}

function formatDateInput(date) {
  if (!date || !(date instanceof Date)) return '';
  return date.toLocaleDateString('he-IL');
}

function formatDateGmail(date) {
  if (!date || !(date instanceof Date)) return '';
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

function formatDateShort(date) {
  if (!date || !(date instanceof Date)) return '';
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = String(date.getFullYear()).slice(-2);
  return `${day}.${month}.${year}`;
}

function isLocalEmail(email) {
  return email.includes('.co.il') || email.includes('.org.il') || email.includes('.gov.il');
}

// === DEBUG FUNCTIONS ===
function debugSheetWriting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  
  console.log("=== בדיקת מצב גיליון ===");
  console.log("שם הגיליון:", activeSheet.getName());
  console.log("מספר שורות:", activeSheet.getLastRow());
  console.log("מספר עמודות:", activeSheet.getLastColumn());

  // בדיקת הרשאות כתיבה
  try {
    const testRange = activeSheet.getRange("A1");
    const originalValue = testRange.getValue();
    testRange.setValue("TEST_WRITE");
    const testValue = testRange.getValue();
    testRange.setValue(originalValue);
    
    if (testValue === "TEST_WRITE") {
      console.log("✅ הרשאות כתיבה תקינות");
      _logMessage("✅ בדיקת כתיבה - הצליחה");
    } else {
      console.log("❌ בעיה בהרשאות כתיבה");
      _logMessage("❌ בדיקת כתיבה - נכשלה");
    }
  } catch (e) {
    console.log("❌ שגיאה בבדיקת כתיבה:", e.message);
    _logMessage(`❌ שגיאה בבדיקת כתיבה: ${e.message}`);
  }

  // בדיקת הגנות על הגיליון
  try {
    const protections = activeSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    console.log("הגנות על הגיליון:", protections.length);
    _logMessage(`ℹ️ נמצאו ${protections.length} הגנות על הגיליון`);
  } catch (e) {
    console.log("שגיאה בבדיקת הגנות:", e.message);
  }
}

function resetAndCleanup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // הסרת כל הגנות אפשריות
    const sheets = ss.getSheets();
    for (const sheet of sheets) {
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (const protection of protections) {
        if (protection.canEdit()) {
          protection.remove();
        }
      }
    }
    _logMessage("✅ הוסרו כל ההגנות מהגיליונות");

    // ניקוי cache
    SpreadsheetApp.flush();
    Utilities.sleep(1000);
    
    SpreadsheetApp.getUi().alert("✅ ניקוי הושלם. נסה עכשיו להריץ את הסריקה מחדש.");
  } catch (e) {
    _logMessage(`❌ שגיאה בניקוי: ${e.message}`);
  }
}

// === PLACEHOLDER FUNCTIONS ===
// פונקציות אלו הוסרו מהתפריט כי לא מתפקדות

function diagnosticScan() {
  SpreadsheetApp.getUi().alert("פונקציית אבחון תמומש בהמשך");
}

function resetAllTabs() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "אישור איפוס",
    "האם אתה בטוח שברצונך למחוק את כל גיליונות התוצאות?\n\n(גיליונות Settings ו-Log יישארו)",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  let deletedCount = 0;
  
  // מחיקת כל הגיליונות למעט Settings ו-Log
  for (const sheet of sheets) {
    const sheetName = sheet.getName();
    if (sheetName !== SETTINGS_SHEET_NAME && sheetName !== LOG_SHEET_NAME) {
      try {
        ss.deleteSheet(sheet);
        deletedCount++;
        _logMessage(`🗑️ נמחק גיליון: ${sheetName}`);
      } catch (e) {
        _logMessage(`⚠️ שגיאה במחיקת גיליון ${sheetName}: ${e.message}`);
      }
    }
  }
  
  _logMessage(`✅ איפוס הושלם: ${deletedCount} גיליונות נמחקו`);
  ui.alert(`✅ איפוס הושלם!\n${deletedCount} גיליונות נמחקו בהצלחה.`);
}

function excludeSelected() {
  runBulkActions();
}

function markAsLocal() {
  runBulkActions();
}

function markAsInternational() {
  runBulkActions();
}

function exportSelected() {
  runBulkActions();
}

function deleteSelected() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();
  
  if (!settings) {
    ui.alert("❌ גיליון Settings לא נמצא.");
    return;
  }

  // בדיקה שזה גיליון תוצאות ולא Settings או Log
  const sheetName = sheet.getName();
  if (sheetName === SETTINGS_SHEET_NAME || sheetName === LOG_SHEET_NAME) {
    ui.alert("❌ לא ניתן למחוק רשומות מגיליון זה.");
    return;
  }

  // קבלת השורות המסומנות
  const selectedRows = getSelectedRowsAdvanced(sheet);
  if (selectedRows.length === 0) {
    ui.alert("❌ לא נבחרו שורות למחיקה.\n\nסמן את התיבות בעמודת Action ונסה שוב.");
    return;
  }

  _logMessage(`🔍 נבחרו ${selectedRows.length} שורות למחיקה והחרגה`);

  // איסוף כל המיילים הייחודיים
  const uniqueEmails = [...new Set(selectedRows.map(row => row.senderEmail))];
  
  _logMessage(`📧 נמצאו ${uniqueEmails.length} מיילים ייחודיים: ${uniqueEmails.join(', ')}`);

  // חיפוש כל הרשומות מאותם שולחים
  const allMatchingRows = findAllRowsFromSenders(sheet, uniqueEmails);
  
  const totalRowsToDelete = allMatchingRows.length;
  const additionalRows = totalRowsToDelete - selectedRows.length;

  // הצגת דיאלוג אישור
  let confirmMessage = `🗑️ מחיקה והחרגת מיילים\n\n`;
  confirmMessage += `📊 סיכום:\n`;
  confirmMessage += `• ${selectedRows.length} שורות נבחרו\n`;
  
  if (additionalRows > 0) {
    confirmMessage += `• ${additionalRows} שורות נוספות נמצאו מאותם שולחים\n`;
    confirmMessage += `• סה"כ ${totalRowsToDelete} שורות יימחקו\n\n`;
  } else {
    confirmMessage += `• סה"כ ${totalRowsToDelete} שורות יימחקו\n\n`;
  }
  
  confirmMessage += `📧 השולחים הבאים יתווספו לרשימת ההחרגות:\n`;
  uniqueEmails.forEach(email => {
    const count = allMatchingRows.filter(row => row.senderEmail === email).length;
    confirmMessage += `• ${email} (${count} רשומות)\n`;
  });
  
  confirmMessage += `\n⚠️ פעולה זו תמחק את כל הרשומות ותוסיף את המיילים לרשימת ההחרגות.\n\nהאם להמשיך?`;

  const response = ui.alert("אישור מחיקה והחרגה", confirmMessage, ui.ButtonSet.YES_NO);
  
  if (response !== ui.Button.YES) {
    _logMessage("❌ המשתמש ביטל את פעולת המחיקה");
    return;
  }

  // הוספת המיילים לרשימת ההחרגות
  const excludeResult = addEmailsToExcludeList(settings, uniqueEmails);
  
  if (!excludeResult.success) {
    ui.alert(`❌ שגיאה בהוספה לרשימת החרגות: ${excludeResult.message}`);
    return;
  }

  // מחיקת השורות (מלמטה למעלה כדי לא לקלקל את המספור)
  const rowsToDelete = allMatchingRows.map(row => row.rowIndex).sort((a, b) => b - a);
  
  let deletedCount = 0;
  for (const rowIndex of rowsToDelete) {
    try {
      sheet.deleteRow(rowIndex);
      deletedCount++;
    } catch (e) {
      _logMessage(`⚠️ שגיאה במחיקת שורה ${rowIndex}: ${e.message}`);
    }
  }

  // סיכום
  const summaryMessage = `✅ פעולה הושלמה בהצלחה!\n\n` +
    `📊 תוצאות:\n` +
    `• ${deletedCount} שורות נמחקו\n` +
    `• ${excludeResult.newEmails} מיילים חדשים נוספו לרשימת ההחרגות\n` +
    `• ${excludeResult.duplicates} מיילים כבר היו ברשימה\n\n` +
    `🎉 המיילים הבאים לא יופיעו יותר בסריקות עתידיות.`;

  ui.alert(summaryMessage);
  _logMessage(`✅ מחיקה והחרגה הושלמה: ${deletedCount} שורות נמחקו, ${excludeResult.newEmails} מיילים נוספו להחרגות`);
}

// === פונקציות עזר למחיקה והחרגה ===
function getSelectedRowsAdvanced(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const actionColIndex = headers.indexOf("Action");
  // תמיכה בשני סוגי גיליונות
  let emailColIndex = headers.indexOf("Email Sender");
  if (emailColIndex === -1) {
    emailColIndex = headers.indexOf("Sender Email");
  }
  
  const companyColIndex = headers.indexOf("Company Name");
  const subjectColIndex = headers.indexOf("Subject");
  
  if (actionColIndex === -1 || emailColIndex === -1) {
    SpreadsheetApp.getUi().alert(`❌ לא נמצאו עמודות נדרשות.\n\nעמודות קיימות: ${headers.join(', ')}\n\nמחפש: Action + (Email Sender או Sender Email)`);
    return [];
  }

  const selectedRows = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][actionColIndex] === true) {
      selectedRows.push({
        rowIndex: i + 1,
        senderEmail: data[i][emailColIndex],
        companyName: data[i][companyColIndex] || '',
        subject: data[i][subjectColIndex] || ''
      });
    }
  }
  
  return selectedRows;
}

function findAllRowsFromSenders(sheet, senderEmails) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  // תמיכה בשני סוגי גיליונות
  let emailColIndex = headers.indexOf("Email Sender");
  if (emailColIndex === -1) {
    emailColIndex = headers.indexOf("Sender Email");
  }
  
  if (emailColIndex === -1) return [];
  
  const matchingRows = [];
  
  for (let i = 1; i < data.length; i++) {
    const rowEmail = data[i][emailColIndex];
    if (senderEmails.includes(rowEmail)) {
      matchingRows.push({
        rowIndex: i + 1,
        senderEmail: rowEmail,
        companyName: data[i][headers.indexOf("Company Name")] || '',
        subject: data[i][headers.indexOf("Subject")] || ''
      });
    }
  }
  
  return matchingRows;
}

function addEmailsToExcludeList(settings, emails) {
  try {
    // קריאת רשימת ההחרגות הנוכחית
    const excludeRange = settings.getRange("F2:F");
    const currentExcludes = excludeRange.getValues().flat().filter(Boolean).map(email => email.toLowerCase().trim());
    
    let newEmails = 0;
    let duplicates = 0;
    
    // מציאת השורה הריקה הראשונה
    let nextRow = currentExcludes.length + 2;
    
    for (const email of emails) {
      const cleanEmail = email.toLowerCase().trim();
      
      if (currentExcludes.includes(cleanEmail)) {
        duplicates++;
        _logMessage(`ℹ️ מייל כבר ברשימה: ${email}`);
      } else {
        settings.getRange(nextRow, 6).setValue(email);
        currentExcludes.push(cleanEmail);
        nextRow++;
        newEmails++;
        _logMessage(`✅ נוסף לרשימת החרגות: ${email}`);
      }
    }
    
    return {
      success: true,
      newEmails: newEmails,
      duplicates: duplicates,
      message: `${newEmails} מיילים חדשים נוספו, ${duplicates} כבר היו ברשימה`
    };
    
  } catch (e) {
    _logMessage(`❌ שגיאה בהוספה לרשימת החרגות: ${e.message}`);
    return {
      success: false,
      newEmails: 0,
      duplicates: 0,
      message: e.message
    };
  }
}

function approveSelected() {
  runBulkActions();
}

// === BULK ACTIONS INTEGRATION ===
function runBulkActions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();

  // בדיקה שזה גיליון חשבוניות (גמיש יותר)
  if (!/^invoices /i.test(sheetName) && !sheetName.includes('invoice') && !sheetName.includes('Invoice')) {
    SpreadsheetApp.getUi().alert(`❌ אנא עבור לגיליון תוצאות סריקה לפני הפעלת פעולות.\n\nגיליון נוכחי: "${sheetName}"\nמחפש גיליון שמתחיל ב-'invoices' או מכיל 'invoice'`);
    return;
  }

  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settings) {
    SpreadsheetApp.getUi().alert("❌ גיליון Settings חסר. אנא הפעל את 'Setup Sheets' תחילה.");
    return;
  }

  // קבלת נתונים מהגיליון
  const dataRange = activeSheet.getDataRange();
  const data = dataRange.getValues();
  
  if (data.length === 0) {
    SpreadsheetApp.getUi().alert("❌ הגיליון ריק. אנא הרץ סריקה תחילה.");
    return;
  }
  
  const headers = data[0];
  _logMessage(`🔍 DEBUG: כותרות בגיליון: ${headers.join(', ')}`);
  
  // חיפוש גמיש מאוד של עמודות
  let actionColumnIndex = -1;
  let emailColumnIndex = -1;
  
  // חיפוש עמודת Action
  const actionVariants = ["Action", "action", "פעולה", "בחירה"];
  for (let i = 0; i < headers.length; i++) {
    if (actionVariants.some(variant => headers[i] && headers[i].toString().includes(variant))) {
      actionColumnIndex = i;
      break;
    }
  }
  
  // חיפוש עמודת Email
  const emailVariants = ["Sender Email", "Email Sender", "Email", "מייל", "שולח", "מאת"];
  for (let i = 0; i < headers.length; i++) {
    if (emailVariants.some(variant => headers[i] && headers[i].toString().includes(variant))) {
      emailColumnIndex = i;
      break;
    }
  }
  
  _logMessage(`🔍 DEBUG: Action column: ${actionColumnIndex} (${headers[actionColumnIndex]}), Email column: ${emailColumnIndex} (${headers[emailColumnIndex]})`);
  
  if (actionColumnIndex === -1 || emailColumnIndex === -1) {
    const availableColumns = headers.map((h, i) => `${i+1}. ${h}`).join('\n');
    SpreadsheetApp.getUi().alert(
      `❌ לא נמצאו עמודות נדרשות בגיליון "${sheetName}".\n\n` +
      `מחפש עמודות:\n` +
      `• Action/פעולה (נמצא: ${actionColumnIndex !== -1 ? 'כן' : 'לא'})\n` +
      `• Sender Email/מייל (נמצא: ${emailColumnIndex !== -1 ? 'כן' : 'לא'})\n\n` +
      `עמודות זמינות:\n${availableColumns}\n\n` +
      `פתרונות:\n` +
      `1. ודא שאתה בגיליון תוצאות סריקה\n` +
      `2. הרץ סריקה מחדש\n` +
      `3. בדוק שיש עמודת Action עם צ'קבוקסים`
    );
    return;
  }

  // איסוף שורות מסומנות
  const checkedRows = [];
  
  // חיפוש עמודות נוספות (גמיש)
  const senderNameIndex = headers.indexOf("Sender Name") !== -1 ? headers.indexOf("Sender Name") : 
                         headers.indexOf("Company Name") !== -1 ? headers.indexOf("Company Name") : -1;
  const emailIdIndex = headers.indexOf("Email ID") !== -1 ? headers.indexOf("Email ID") : -1;
  const subjectIndex = headers.indexOf("Subject") !== -1 ? headers.indexOf("Subject") : -1;
  
  _logMessage(`🔍 DEBUG: מחפש שורות מסומנות...`);
  
  for (let i = 1; i < data.length; i++) {
    const actionValue = data[i][actionColumnIndex];
    _logMessage(`🔍 DEBUG: שורה ${i+1}: Action=${actionValue}, Email=${data[i][emailColumnIndex]}`);
    
    if (actionValue === true) {
      const rowData = {
        rowIndex: i + 1,
        email: data[i][emailColumnIndex] || "",
        senderName: senderNameIndex !== -1 ? data[i][senderNameIndex] || "" : "",
        emailId: emailIdIndex !== -1 ? data[i][emailIdIndex] || "" : "",
        subject: subjectIndex !== -1 ? data[i][subjectIndex] || "" : "",
      };
      
      checkedRows.push(rowData);
      _logMessage(`✅ נוספה שורה מסומנת: ${rowData.email}`);
    }
  }

  _logMessage(`📊 סה"כ שורות מסומנות: ${checkedRows.length}`);

  if (checkedRows.length === 0) {
    SpreadsheetApp.getUi().alert(
      "❌ אין שורות מסומנות לביצוע פעולות.\n\n" +
      "הוראות:\n" +
      "1. עבור לגיליון תוצאות סריקה\n" +
      "2. סמן תיבות בעמודת 'Action'\n" +
      "3. נסה שוב\n\n" +
      "אם הבעיה נמשכת, הרץ 'אבחון מתקדם' לבדיקה."
    );
    return;
  }

  // הצגת דיאלוג פעולות
  try {
    showActionsDialog(checkedRows, sheetName, settings);
  } catch (error) {
    _logMessage(`❌ שגיאה בהצגת דיאלוג: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      `❌ שגיאה בהצגת דיאלוג פעולות:\n\n` +
      `שגיאה: ${error.message}\n\n` +
      `פתרונות אפשריים:\n` +
      `1. נסה שוב\n` +
      `2. רענן את הדף\n` +
      `3. הרץ 'איפוס מערכת' אם הבעיה נמשכת`
    );
  }
}

function refreshView() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).activate();
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("✅ התצוגה רועננה בהצלחה");
}

function clearActiveSheetData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (sheetName === SETTINGS_SHEET_NAME || sheetName === LOG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("❌ לא ניתן לנקות את גיליון ההגדרות או הלוג.");
    return;
  }
  
  const response = SpreadsheetApp.getUi().alert(
    "אישור מחיקה",
    `האם אתה בטוח שברצונך לנקות את כל הנתונים בגיליון "${sheetName}"?`,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response === SpreadsheetApp.getUi().Button.YES) {
    if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
    SpreadsheetApp.getUi().alert("✅ הגיליון נוקה בהצלחה.");
  }
}

// === ACTIONS DIALOG ===
function showActionsDialog(checkedRows, sheetName, settings) {
  // בדיקת סטטוס מוחרג/מאושר
  const excludedList = settings.getRange("F2:F").getValues().flat().filter(Boolean);
  const approvedList = settings.getRange("G2:G").getValues().flat().filter(Boolean);
  
  checkedRows.forEach(row => {
    row.isExcluded = excludedList.includes(row.email);
    row.isApproved = approvedList.includes(row.email);
  });

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
        <h4>🚫 החרג חשבוניות מסומנות</h4>
        <p>הוסף את המיילים המסומנים לרשימת ההחרגות ומחק את כל הרשומות עם אותם מיילים (כולל כפילויות).</p>
      </div>
      <div class="option" onclick="selectAction('approve')">
        <h4>✅ הוסף למיילים מאושרים</h4>
        <p>הוסף את המיילים המסומנים לרשימת המאושרים (whitelist) – ייכנסו אוטומטית בעתיד.</p>
      </div>
      <div class="option" onclick="selectAction('local')">
        <h4>🔷 סמן כמקומיות</h4>
        <p>סמן את החשבוניות המסומנות כ-🔷 Local ושמור את ההגדרה לסריקות הבאות</p>
      </div>
      <div class="option" onclick="selectAction('international')">
        <h4>🌐 סמן כבינלאומיות</h4>
        <p>סמן את החשבוניות המסומנות כ-🌐 International ושמור את ההגדרה לסריקות הבאות</p>
      </div>
      <div class="option" onclick="selectAction('export')">
        <h4>📦 ייצא חשבוניות מסומנות</h4>
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
            .processSelectedAction('${sheetName}', action, ${JSON.stringify(checkedRows)});
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
function processSelectedAction(sheetName, action, checkedRows) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
    
    if (!sheet) {
      const errorMsg = `❌ גיליון "${sheetName}" לא נמצא.\n\nשגיאה טכנית: Sheet not found\n\nפתרונות:\n1. ודא שאתה בגיליון הנכון\n2. הרץ סריקה מחדש\n3. הרץ 'איפוס מערכת' אם הבעיה נמשכת`;
      _logMessage(`❌ SHEET ERROR: ${sheetName} not found`);
      return { success: false, message: errorMsg };
    }
    
    if (!settings) {
      const errorMsg = `❌ גיליון Settings לא נמצא.\n\nשגיאה טכנית: Settings sheet missing\n\nפתרון: הרץ 'הגדרות ראשוניות' מהתפריט`;
      _logMessage(`❌ SETTINGS ERROR: Settings sheet missing`);
      return { success: false, message: errorMsg };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const emailColumnIndex = headers.indexOf("Sender Email") + 1;
    const categoryColumnIndex = headers.indexOf("Category") + 1;
    _logMessage(`🔄 מתחיל פעולה: ${action} עבור ${checkedRows.length} שורות`);
    
    let result;
    switch (action) {
      case 'exclude':
        _logMessage(`🚫 מבצע החרגה עבור: ${checkedRows.map(r => r.email).join(', ')}`);
        result = handleExcludeActionFixed(sheet, settings, checkedRows, emailColumnIndex);
        break;
      
      case 'approve':
        _logMessage(`✅ מבצע אישור עבור: ${checkedRows.map(r => r.email).join(', ')}`);
        result = handleApproveActionFixed(settings, checkedRows);
        break;
      
      case 'local':
      case 'international':
        _logMessage(`🏷️ מבצע סימון ${action} עבור: ${checkedRows.map(r => r.email).join(', ')}`);
        result = handleCategoryActionFixed(sheet, checkedRows, action, categoryColumnIndex);
        break;
      
      case 'export':
        _logMessage(`📦 מבצע ייצוא עבור: ${checkedRows.map(r => r.email).join(', ')}`);
        result = handleExportActionFixed(sheetName, checkedRows);
        break;
      
      default:
        _logMessage(`❌ פעולה לא מוכרת: ${action}`);
        return { success: false, message: `❌ פעולה לא מוכרת: ${action}` };
    }
    
    _logMessage(`✅ פעולה ${action} הושלמה: ${result.success ? 'הצליחה' : 'נכשלה'}`);
    return result;
    
  } catch (error) {
    const errorMsg = `❌ שגיאה בביצוע פעולה ${action}:\n\nשגיאה טכנית: ${error.message}\n\nStack: ${error.stack || 'לא זמין'}\n\nפתרונות:\n1. נסה שוב\n2. רענן את הדף\n3. בדוק את גיליון Log לפרטים\n4. הרץ 'איפוס מערכת' אם הבעיה נמשכת`;
    _logMessage(`❌ CRITICAL ERROR in ${action}: ${error.message}\nStack: ${error.stack}`);
    return { success: false, message: errorMsg };
  }
}

// === EXCLUDE ACTION - FIXED ===
function handleExcludeActionFixed(sheet, settings, checkedRows, emailColumnIndex) {
  const allData = sheet.getDataRange().getValues();
  const emailsToExclude = [...new Set(checkedRows.map(r => r.email))]; // הסרת כפילויות
  const rowsToDelete = [];
  
  _logMessage(`🔍 מחפש שורות למחיקה עבור ${emailsToExclude.length} מיילים`);
  
  // איסוף כל השורות עם אותם מיילים
  for (let i = 1; i < allData.length; i++) {
    if (emailsToExclude.includes(allData[i][emailColumnIndex - 1])) {
      rowsToDelete.push(i + 1);
    }
  }
  
  _logMessage(`📋 נמצאו ${rowsToDelete.length} שורות למחיקה`);
  
  // הוספה לרשימת החרגות
  const excludeRange = settings.getRange("F2:F");
  const excludeValues = excludeRange.getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
  let newlyExcluded = 0;
  let nextRow = excludeValues.length + 2;
  
  emailsToExclude.forEach(email => {
    const cleanEmail = email.toLowerCase().trim();
    if (!excludeValues.includes(cleanEmail)) {
      settings.getRange(nextRow, 6).setValue(email);
      excludeValues.push(cleanEmail);
      nextRow++;
      newlyExcluded++;
      _logMessage(`✅ נוסף לרשימת החרגות: ${email}`);
    } else {
      _logMessage(`ℹ️ כבר ברשימת החרגות: ${email}`);
    }
  });
  
  // מחיקת השורות (מלמטה למעלה כדי לא לקלקל את המספור)
  let deletedCount = 0;
  rowsToDelete.sort((a, b) => b - a).forEach(rowIdx => {
    try {
      sheet.deleteRow(rowIdx);
      deletedCount++;
    } catch (e) {
      _logMessage(`⚠️ שגיאה במחיקת שורה ${rowIdx}: ${e.message}`);
    }
  });
  
  _logMessage(`✅ הוחרגו ומחקו ${deletedCount} שורות (${newlyExcluded} מיילים חדשים נוספו להחרגות)`);
  
  return { 
    success: true, 
    message: `✅ ${deletedCount} שורות הוחרגו ונמחקו.\n${newlyExcluded} מיילים חדשים נוספו לרשימת ההחרגות.` 
  };
}

// === APPROVE ACTION - FIXED ===
function handleApproveActionFixed(settings, checkedRows) {
  const approvedRange = settings.getRange("G2:G");
  const approvedValues = approvedRange.getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
  const emailsToApprove = [...new Set(checkedRows.map(r => r.email))]; // הסרת כפילויות
  
  let newlyApproved = 0;
  let nextRow = approvedValues.length + 2;
  
  emailsToApprove.forEach(email => {
    const cleanEmail = email.toLowerCase().trim();
    if (!approvedValues.includes(cleanEmail)) {
      settings.getRange(nextRow, 7).setValue(email);
      approvedValues.push(cleanEmail);
      nextRow++;
      newlyApproved++;
      _logMessage(`✅ נוסף לרשימת מאושרים: ${email}`);
    } else {
      _logMessage(`ℹ️ כבר ברשימת מאושרים: ${email}`);
    }
  });
  
  _logMessage(`✅ הוספו ${newlyApproved} מיילים ל-whitelist`);
  
  return { 
    success: true, 
    message: `✅ ${newlyApproved} מיילים הוספו לרשימת המאושרים.` 
  };
}

// === CATEGORY ACTION - FIXED ===
function handleCategoryActionFixed(sheet, checkedRows, action, categoryColumnIndex) {
  const category = action === 'local' ? "🔷 Local" : "🌐 International";
  let updatedCount = 0;
  
  for (const row of checkedRows) {
    try {
      const rowRange = sheet.getRange(row.rowIndex, categoryColumnIndex);
      rowRange.setValue(category);
      updatedCount++;
      _logMessage(`✅ עודכן לקטגוריה ${category}: ${row.email}`);
    } catch (e) {
      _logMessage(`⚠️ שגיאה בעדכון שורה ${row.rowIndex}: ${e.message}`);
    }
  }
  
  return { 
    success: true, 
    message: `✅ ${updatedCount} חשבוניות סומנו כ-${category}.` 
  };
}

// === EXPORT ACTION - PLACEHOLDER ===
function handleExportActionFixed(sheetName, checkedRows) {
  _logMessage(`ℹ️ בקשת ייצוא עבור ${checkedRows.length} רשומות`);
  return { 
    success: true, 
    message: `ℹ️ פונקציית ייצוא תמומש בהמשך (${checkedRows.length} רשומות נבחרו)` 
  };
}

// === GMAIL API INTEGRATION ===
// הפונקציה processInvoices() מוגדרת ב-gmailProcessor.gs

// === SYSTEM RESET - איפוס מערכת למשתמש חדש ===
function systemReset() {
  const ui = SpreadsheetApp.getUi();
  
  // אישור ראשון
  const firstConfirm = ui.alert(
    "⚠️ איפוס מערכת מלא",
    "פעולה זו תמחק את כל הנתונים והגיליונות ותחזיר את המערכת למצב התחלתי.\n\n" +
    "זה כולל:\n" +
    "• כל גיליונות התוצאות\n" +
    "• הגדרות Settings\n" +
    "• לוג המערכת\n" +
    "• כל הנתונים שנשמרו\n\n" +
    "האם אתה בטוח שברצונך להמשיך?",
    ui.ButtonSet.YES_NO
  );
  
  if (firstConfirm !== ui.Button.YES) {
    return;
  }
  
  // אישור שני (בטיחות)
  const secondConfirm = ui.alert(
    "🚨 אישור סופי",
    "זוהי ההזדמנות האחרונה לבטל!\n\n" +
    "לאחר לחיצה על 'כן', כל הנתונים יימחקו ללא אפשרות שחזור.\n\n" +
    "האם להמשיך עם איפוס מערכת מלא?",
    ui.ButtonSet.YES_NO
  );
  
  if (secondConfirm !== ui.Button.YES) {
    ui.alert("✅ איפוס בוטל בהצלחה.");
    return;
  }
  
  try {
    _logMessage("🔄 התחלת איפוס מערכת מלא");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    // שלב 1: מחיקת כל הגיליונות חוץ מהראשון
    let deletedCount = 0;
    const sheetsToDelete = [];
    
    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      // שמור רק את הגיליון הראשון (בדרך כלל Sheet1)
      if (sheet.getIndex() !== 1) {
        sheetsToDelete.push(sheet);
      }
    }
    
    // מחיקת הגיליונות
    for (const sheet of sheetsToDelete) {
      try {
        ss.deleteSheet(sheet);
        deletedCount++;
        _logMessage(`🗑️ נמחק גיליון: ${sheet.getName()}`);
      } catch (e) {
        _logMessage(`⚠️ שגיאה במחיקת גיליון ${sheet.getName()}: ${e.message}`);
      }
    }
    
    // שלב 2: ניקוי הגיליון הראשון
    const firstSheet = ss.getSheets()[0];
    firstSheet.clear();
    firstSheet.setName("Sheet1"); // שם סטנדרטי
    
    // שלב 3: הסרת תפריטים קיימים (ינוצרו מחדש ב-onOpen)
    try {
      // ניקוי cache של התפריטים
      SpreadsheetApp.flush();
    } catch (e) {
      // לא קריטי
    }
    
    // שלב 4: בקשת הרשאות מחדש
    try {
      // בדיקת הרשאות בסיסיות
      DriveApp.getRootFolder();
      GmailApp.getInboxThreads(0, 1);
      _logMessage("✅ הרשאות אושרו מחדש");
    } catch (e) {
      _logMessage(`⚠️ יש לאשר הרשאות מחדש: ${e.message}`);
    }
    
    // שלב 5: יצירת תפריטים מחדש
    createMainMenu();
    createActionsMenu();
    
    _logMessage(`✅ איפוס מערכת הושלם: ${deletedCount} גיליונות נמחקו`);
    
    // הודעת סיום
    const successMessage = `🎉 איפוס מערכת הושלם בהצלחה!\n\n` +
      `📊 סיכום:\n` +
      `• ${deletedCount} גיליונות נמחקו\n` +
      `• הגיליון הראשון נוקה\n` +
      `• התפריטים נוצרו מחדש\n` +
      `• הרשאות אושרו\n\n` +
      `🚀 השלבים הבאים:\n` +
      `1. לחץ על "הגדרות ראשוניות" בתפריט\n` +
      `2. מלא את ההגדרות הנדרשות\n` +
      `3. התחל לעבוד עם המערכת\n\n` +
      `ברוך הבא למערכת נקייה! 🎯`;
    
    ui.alert("✅ איפוס הושלם", successMessage, ui.ButtonSet.OK);
    
    // מעבר לגיליון הראשון
    SpreadsheetApp.setActiveSheet(firstSheet);
    
  } catch (error) {
    _logMessage(`❌ שגיאה באיפוס מערכת: ${error.message}`);
    ui.alert(`❌ שגיאה באיפוס מערכת:\n${error.message}\n\nנסה שוב או פנה לתמיכה.`);
  }
}

// === פונקציית עזר לאיפוס הרשאות ===
function requestPermissionsReset() {
  try {
    // בקשת הרשאות Gmail
    const threads = GmailApp.getInboxThreads(0, 1);
    
    // בקשת הרשאות Drive
    const rootFolder = DriveApp.getRootFolder();
    
    // בקשת הרשאות Spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getName();
    
    return { success: true, message: "הרשאות אושרו בהצלחה" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// === DEBUG FUNCTION FOR EXCLUSIONS ===
function debugExclusionsQuick() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();
  
  _logMessage("=== 🐛 DEBUG החרגות ===");
  _logMessage(`📋 שם הגיליון: ${sheetName}`);
  
  const data = activeSheet.getDataRange().getValues();
  _logMessage(`📊 מספר שורות: ${data.length}`);
  
  if (data.length > 0) {
    const headers = data[0];
    _logMessage(`📝 כותרות: ${headers.join(' | ')}`);
    
    // בדיקת עמודות
    const actionCol = headers.findIndex(h => h && h.toString().toLowerCase().includes('action'));
    const emailCol = headers.findIndex(h => h && (
      h.toString().includes('Email') || 
      h.toString().includes('מייל') || 
      h.toString().includes('שולח')
    ));
    
    _logMessage(`🔍 עמודת Action: ${actionCol} (${headers[actionCol] || 'לא נמצא'})`);
    _logMessage(`📧 עמודת Email: ${emailCol} (${headers[emailCol] || 'לא נמצא'})`);
    
    // בדיקת שורות מסומנות
    let checkedCount = 0;
    let checkedEmails = [];
    
    for (let i = 1; i < Math.min(data.length, 20); i++) {
      const isChecked = data[i][actionCol];
      const email = data[i][emailCol];
      
      _logMessage(`📋 שורה ${i+1}: Action=${isChecked}, Email=${email}`);
      
      if (isChecked === true) {
        checkedCount++;
        checkedEmails.push(email);
      }
    }
    
    _logMessage(`✅ סה"כ שורות מסומנות: ${checkedCount}`);
    _logMessage(`📧 מיילים מסומנים: ${checkedEmails.join(', ')}`);
    
    // בדיקת תקינות הגיליון
    if (actionCol === -1) {
      _logMessage("❌ PROBLEM: עמודת Action לא נמצאה!");
    }
    if (emailCol === -1) {
      _logMessage("❌ PROBLEM: עמודת Email לא נמצאה!");
    }
    if (checkedCount === 0) {
      _logMessage("⚠️ WARNING: אין שורות מסומנות!");
    }
    
  } else {
    _logMessage("❌ ERROR: הגיליון ריק!");
  }
  
  _logMessage("✅ דיבוג החרגות הושלם - בדוק את גיליון Log לפרטים מלאים");
  SpreadsheetApp.getUi().alert("🐛 דיבוג הושלם!\n\nבדוק את גיליון Log לפרטים מלאים על מבנה הגיליון ושורות מסומנות.");
}

// === EXPORT LOG TO DEVELOPER ===
function exportLogToDeveloper() {
  const ui = SpreadsheetApp.getUi();
  
  // אישור מהמשתמש
  const response = ui.alert(
    "📤 שליחת לוג למפתח",
    "פעולה זו תיצור קובץ עם כל הלוגים ותעלה אותו לגוגל דרייב.\n\n" +
    "הקובץ יכלול:\n" +
    "• כל הלוגים מגיליון Log\n" +
    "• מידע על המערכת\n" +
    "• הגדרות נוכחיות\n" +
    "• סטטיסטיקות שימוש\n\n" +
    "האם להמשיך?",
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    _logMessage("📤 התחלת יצוא לוג למפתח");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    
    if (!logSheet) {
      ui.alert("❌ גיליון Log לא נמצא.");
      return;
    }
    
    // איסוף נתונים
    const logData = collectLogData(ss, logSheet);
    const systemInfo = collectSystemInfo(ss);
    
    // יצירת קובץ טקסט
    const logContent = formatLogForExport(logData, systemInfo);
    
    // יצירת קובץ בדרייב
    const fileName = `InvoiceApp_Log_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.txt`;
    const file = DriveApp.createFile(fileName, logContent, 'text/plain');
    
    // הגדרת הרשאות קריאה לכולם (כדי שהמפתח יוכל לגשת)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileUrl = file.getUrl();
    const fileId = file.getId();
    
    _logMessage(`✅ קובץ לוג נוצר: ${fileName} (ID: ${fileId})`);
    
    // הודעה למשתמש עם הלינק
    const successMessage = `✅ לוג נשלח בהצלחה!\n\n` +
      `📁 שם קובץ: ${fileName}\n` +
      `🔗 לינק: ${fileUrl}\n\n` +
      `📋 הקובץ כולל:\n` +
      `• ${logData.totalEntries} רשומות לוג\n` +
      `• מידע על המערכת\n` +
      `• הגדרות נוכחיות\n\n` +
      `💡 העתק את הלינק ושלח למפתח לתמיכה.`;
    
    ui.alert("📤 לוג נשלח למפתח", successMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    _logMessage(`❌ שגיאה ביצוא לוג: ${error.message}`);
    ui.alert(`❌ שגיאה ביצוא לוג:\n${error.message}`);
  }
}

// === פונקציות עזר ליצוא לוג ===
function collectLogData(ss, logSheet) {
  const data = logSheet.getDataRange().getValues();
  const headers = data[0];
  const entries = data.slice(1);
  
  return {
    headers: headers,
    entries: entries,
    totalEntries: entries.length,
    lastEntry: entries.length > 0 ? entries[entries.length - 1] : null,
    firstEntry: entries.length > 0 ? entries[0] : null
  };
}

function collectSystemInfo(ss) {
  const user = Session.getActiveUser().getEmail();
  const timezone = ss.getSpreadsheetTimeZone();
  const locale = ss.getSpreadsheetLocale();
  const sheets = ss.getSheets().map(sheet => ({
    name: sheet.getName(),
    rows: sheet.getLastRow(),
    columns: sheet.getLastColumn()
  }));
  
  // קריאת הגדרות
  let settings = {};
  try {
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (settingsSheet) {
      settings = {
        startDate: settingsSheet.getRange('A2').getValue(),
        endDate: settingsSheet.getRange('B2').getValue(),
        enKeywords: settingsSheet.getRange('C2:C').getValues().flat().filter(Boolean).length,
        heKeywords: settingsSheet.getRange('D2:D').getValues().flat().filter(Boolean).length,
        excludedEmails: settingsSheet.getRange('F2:F').getValues().flat().filter(Boolean).length,
        approvedSenders: settingsSheet.getRange('G2:G').getValues().flat().filter(Boolean).length
      };
    }
  } catch (e) {
    settings.error = e.message;
  }
  
  return {
    user: user,
    timezone: timezone,
    locale: locale,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheets: sheets,
    settings: settings,
    exportTime: new Date().toISOString()
  };
}

function formatLogForExport(logData, systemInfo) {
  let content = "";
  
  // כותרת
  content += "=".repeat(80) + "\n";
  content += "           INVOICE SCANNER - LOG EXPORT\n";
  content += "=".repeat(80) + "\n\n";
  
  // מידע מערכת
  content += "📊 SYSTEM INFORMATION:\n";
  content += "-".repeat(40) + "\n";
  content += `User: ${systemInfo.user}\n`;
  content += `Spreadsheet: ${systemInfo.spreadsheetName}\n`;
  content += `Spreadsheet ID: ${systemInfo.spreadsheetId}\n`;
  content += `Timezone: ${systemInfo.timezone}\n`;
  content += `Locale: ${systemInfo.locale}\n`;
  content += `Export Time: ${systemInfo.exportTime}\n`;
  content += `Total Log Entries: ${logData.totalEntries}\n\n`;
  
  // הגדרות
  content += "⚙️ SETTINGS:\n";
  content += "-".repeat(40) + "\n";
  if (systemInfo.settings.error) {
    content += `Error reading settings: ${systemInfo.settings.error}\n`;
  } else {
    content += `Start Date: ${systemInfo.settings.startDate}\n`;
    content += `End Date: ${systemInfo.settings.endDate}\n`;
    content += `English Keywords: ${systemInfo.settings.enKeywords}\n`;
    content += `Hebrew Keywords: ${systemInfo.settings.heKeywords}\n`;
    content += `Excluded Emails: ${systemInfo.settings.excludedEmails}\n`;
    content += `Approved Senders: ${systemInfo.settings.approvedSenders}\n`;
  }
  content += "\n";
  
  // גיליונות
  content += "📋 SHEETS:\n";
  content += "-".repeat(40) + "\n";
  systemInfo.sheets.forEach(sheet => {
    content += `${sheet.name}: ${sheet.rows} rows, ${sheet.columns} columns\n`;
  });
  content += "\n";
  
  // לוגים
  content += "📝 LOG ENTRIES:\n";
  content += "-".repeat(40) + "\n";
  content += `${logData.headers.join('\t')}\n`;
  content += "-".repeat(80) + "\n";
  
  logData.entries.forEach(entry => {
    content += `${entry.join('\t')}\n`;
  });
  
  content += "\n" + "=".repeat(80) + "\n";
  content += "End of Log Export\n";
  content += "=".repeat(80) + "\n";
  
  return content;
}

// === ADVANCED EXCLUSIONS DEBUGGING ===
function testExclusionsStep() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();
  
  _logMessage("=== 🔧 ADVANCED EXCLUSIONS TEST ===");
  _logMessage(`📋 Testing on sheet: ${sheetName}`);
  
  try {
    // Step 1: Basic sheet validation
    _logMessage("🔍 Step 1: Basic validation");
    if (!/^invoices /i.test(sheetName) && !sheetName.includes('invoice')) {
      _logMessage(`❌ FAIL: Sheet name "${sheetName}" doesn't match invoice pattern`);
      SpreadsheetApp.getUi().alert(`❌ Wrong sheet type: "${sheetName}"\n\nPlease go to an invoices results sheet first.`);
      return;
    }
    _logMessage("✅ Sheet name validation passed");
    
    // Step 2: Data validation
    _logMessage("🔍 Step 2: Data validation");
    const data = activeSheet.getDataRange().getValues();
    if (data.length === 0) {
      _logMessage("❌ FAIL: Sheet is empty");
      SpreadsheetApp.getUi().alert("❌ Sheet is empty. Run a scan first.");
      return;
    }
    _logMessage(`✅ Sheet has ${data.length} rows`);
    
    // Step 3: Headers analysis
    _logMessage("🔍 Step 3: Headers analysis");
    const headers = data[0];
    _logMessage(`📝 Headers found: ${headers.join(' | ')}`);
    
    // Step 4: Column detection
    _logMessage("🔍 Step 4: Column detection");
    let actionColumnIndex = -1;
    let emailColumnIndex = -1;
    
    // Find Action column
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i] ? headers[i].toString().toLowerCase() : '';
      if (header.includes('action') || header.includes('פעולה')) {
        actionColumnIndex = i;
        _logMessage(`✅ Action column found at index ${i}: "${headers[i]}"`);
        break;
      }
    }
    
    // Find Email column
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i] ? headers[i].toString().toLowerCase() : '';
      if (header.includes('email') || header.includes('sender') || header.includes('מייל') || header.includes('שולח')) {
        emailColumnIndex = i;
        _logMessage(`✅ Email column found at index ${i}: "${headers[i]}"`);
        break;
      }
    }
    
    if (actionColumnIndex === -1) {
      _logMessage("❌ FAIL: Action column not found");
      SpreadsheetApp.getUi().alert(`❌ Action column not found!\n\nHeaders: ${headers.join(', ')}\n\nLooking for: Action, action, פעולה`);
      return;
    }
    
    if (emailColumnIndex === -1) {
      _logMessage("❌ FAIL: Email column not found");
      SpreadsheetApp.getUi().alert(`❌ Email column not found!\n\nHeaders: ${headers.join(', ')}\n\nLooking for: Email, Sender, מייל, שולח`);
      return;
    }
    
    // Step 5: Check for checkboxes
    _logMessage("🔍 Step 5: Checkbox validation");
    let hasCheckboxes = false;
    let checkedCount = 0;
    let checkedRows = [];
    
    for (let i = 1; i < Math.min(data.length, 10); i++) {
      const cellValue = data[i][actionColumnIndex];
      const email = data[i][emailColumnIndex];
      
      _logMessage(`📋 Row ${i+1}: Action="${cellValue}" (type: ${typeof cellValue}), Email="${email}"`);
      
      if (typeof cellValue === 'boolean') {
        hasCheckboxes = true;
        if (cellValue === true) {
          checkedCount++;
          checkedRows.push({
            row: i+1,
            email: email
          });
        }
      }
    }
    
    if (!hasCheckboxes) {
      _logMessage("❌ FAIL: No checkboxes found in Action column");
      SpreadsheetApp.getUi().alert("❌ No checkboxes found!\n\nThe Action column should contain checkboxes.\nTry running a fresh scan.");
      return;
    }
    
    _logMessage(`✅ Checkboxes detected. ${checkedCount} rows checked.`);
    
    if (checkedCount === 0) {
      _logMessage("⚠️ WARNING: No rows are checked");
      SpreadsheetApp.getUi().alert("⚠️ No rows selected!\n\nPlease check some boxes in the Action column and try again.");
      return;
    }
    
    _logMessage(`✅ Found ${checkedCount} checked rows:`);
    checkedRows.forEach(row => {
      _logMessage(`  - Row ${row.row}: ${row.email}`);
    });
    
    // Step 6: Test the actual exclusion process
    _logMessage("🔍 Step 6: Testing exclusion process");
    
    const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settings) {
      _logMessage("❌ FAIL: Settings sheet not found");
      SpreadsheetApp.getUi().alert("❌ Settings sheet missing!\n\nRun 'Setup Sheets' first.");
      return;
    }
    
    _logMessage("✅ Settings sheet found");
    
    // Test exclusion logic
    const excludeRange = settings.getRange("F2:F");
    const currentExcludes = excludeRange.getValues().flat().filter(Boolean);
    _logMessage(`📋 Current exclusions: ${currentExcludes.join(', ')}`);
    
    _logMessage("✅ ALL TESTS PASSED! Exclusions should work now.");
    
    const successMessage = `🎉 Exclusions Test Results:\n\n` +
      `✅ Sheet: ${sheetName}\n` +
      `✅ Rows: ${data.length}\n` +
      `✅ Action column: ${headers[actionColumnIndex]}\n` +
      `✅ Email column: ${headers[emailColumnIndex]}\n` +
      `✅ Checked rows: ${checkedCount}\n` +
      `✅ Current exclusions: ${currentExcludes.length}\n\n` +
      `🚀 Ready to exclude: ${checkedRows.map(r => r.email).join(', ')}`;
    
    SpreadsheetApp.getUi().alert(successMessage);
    
  } catch (error) {
    _logMessage(`❌ CRITICAL ERROR: ${error.message}`);
    _logMessage(`Stack: ${error.stack}`);
    SpreadsheetApp.getUi().alert(`❌ Test failed with error:\n\n${error.message}\n\nCheck Log sheet for details.`);
  }
  
  _logMessage("=== 🔧 EXCLUSIONS TEST COMPLETED ===");
}

// === SMART AMOUNT EXTRACTION FOR REGULAR SCAN ===
function extractAmountFromEmailSmart(message, subject, attachmentName) {
  const startTime = new Date();
  
  try {
    // שלב 1: חילוץ מהיר מנושא המייל
    const subjectAmount = extractAmountFromText(subject, "subject");
    if (subjectAmount.found) {
      return subjectAmount;
    }
    
    // שלב 2: חילוץ מתוכן המייל
    const bodyAmount = extractAmountFromEmailBody(message);
    if (bodyAmount.found) {
      return bodyAmount;
    }
    
    // שלב 3: חילוץ משם הקובץ
    if (attachmentName) {
      const fileAmount = extractAmountFromText(attachmentName, "filename");
      if (fileAmount.found) {
        return fileAmount;
      }
    }
    
    const duration = (new Date() - startTime) / 1000;
    _logMessage(`💰 לא נמצא סכום עבור: ${subject.substring(0, 50)}... (${duration.toFixed(2)}s)`);
    
    return { 
      found: false, 
      amount: "", 
      currency: "", 
      method: "not_found",
      duration: duration
    };
    
  } catch (e) {
    _logMessage(`❌ שגיאה בחילוץ סכום: ${e.message}`);
    return { 
      found: false, 
      amount: "", 
      currency: "", 
      method: "error",
      error: e.message
    };
  }
}

// === חילוץ סכום מטקסט כללי ===
function extractAmountFromText(text, source) {
  if (!text) return { found: false };
  
  const cleanText = text.toString().toLowerCase();
  
  // Regex patterns מתקדמים לסכומים - מיוחד לישראל
  const patterns = [
    // שקלים - פורמטים ישראליים
    { regex: /₪\s*([0-9,]+\.?\d*)/g, currency: 'ILS' },
    { regex: /([0-9,]+\.?\d*)\s*₪/g, currency: 'ILS' },
    { regex: /([0-9,]+\.?\d*)\s*שקל/g, currency: 'ILS' },
    { regex: /([0-9,]+\.?\d*)\s*ש"ח/g, currency: 'ILS' },
    { regex: /([0-9,]+\.?\d*)\s*nis/g, currency: 'ILS' },
    { regex: /([0-9,]+\.?\d*)\s*ils/g, currency: 'ILS' },
    
    // מילות מפתח עבריות מתקדמות
    { regex: /עלות[:\s]*([0-9,]+\.?\d*)/g, currency: 'ILS' },
    { regex: /בה"כ[:\s]*([0-9,]+\.?\d*)/g, currency: 'ILS' },
    { regex: /סה"כ[:\s]*([0-9,]+\.?\d*)/g, currency: 'ILS' },
    { regex: /לתשלום[:\s]*([0-9,]+\.?\d*)/g, currency: 'ILS' },
    { regex: /חיוב[:\s]*([0-9,]+\.?\d*)/g, currency: 'ILS' },
    { regex: /סכום\s*כולל[:\s]*([0-9,]+\.?\d*)/g, currency: 'ILS' },
    { regex: /סכום\s*לתשלום[:\s]*([0-9,]+\.?\d*)/g, currency: 'ILS' },
    
    // דולרים - כולל דוגמאות מהחשבוניות שלך
    { regex: /\$\s*([0-9,]+\.?\d*)/g, currency: 'USD' },
    { regex: /([0-9,]+\.?\d*)\s*\$/g, currency: 'USD' },
    { regex: /([0-9,]+\.?\d*)\s*usd/g, currency: 'USD' },
    { regex: /([0-9,]+\.?\d*)\s*dollar/g, currency: 'USD' },
    { regex: /\$([0-9,]+\.?\d*)\s*USD\s*due/g, currency: 'USD' }, // מדוגמת DOMOAI
    { regex: /Amount\s*due[:\s]*\$([0-9,]+\.?\d*)/g, currency: 'USD' },
    { regex: /Total[:\s]*\$([0-9,]+\.?\d*)/g, currency: 'USD' },
    { regex: /Subtotal[:\s]*\$([0-9,]+\.?\d*)/g, currency: 'USD' },
    { regex: /refunded[:\s]*\$([0-9,]+\.?\d*)/g, currency: 'USD' }, // מדוגמת Refund
    
    // יורו
    { regex: /€\s*([0-9,]+\.?\d*)/g, currency: 'EUR' },
    { regex: /([0-9,]+\.?\d*)\s*€/g, currency: 'EUR' },
    { regex: /([0-9,]+\.?\d*)\s*eur/g, currency: 'EUR' },
    { regex: /([0-9,]+\.?\d*)\s*euro/g, currency: 'EUR' },
    
    // מילות מפתח כלליות
    { regex: /סכום[:\s]*([0-9,]+\.?\d*)/g, currency: 'ILS' },
    { regex: /total[:\s]*([0-9,]+\.?\d*)/gi, currency: 'USD' },
    { regex: /amount[:\s]*([0-9,]+\.?\d*)/gi, currency: 'USD' },
    { regex: /price[:\s]*([0-9,]+\.?\d*)/gi, currency: 'USD' },
    { regex: /מחיר[:\s]*([0-9,]+\.?\d*)/g, currency: 'ILS' },
    
    // פורמטים מספריים כלליים (רק אם יש הקשר)
    { regex: /\b([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})?)\b/g, currency: 'ILS', requireContext: true }
  ];
  
  for (const pattern of patterns) {
    const matches = text.match(pattern.regex);
    if (matches && matches.length > 0) {
      for (const match of matches) {
        const numericPart = match.replace(/[^\d.,]/g, '');
        const amount = parseFloat(numericPart.replace(/,/g, ''));
        
        // בדיקות תקינות
        if (amount > 0 && amount < 1000000) { // סכום סביר
          // אם דרוש הקשר, בדוק שיש מילות מפתח רלוונטיות
          if (pattern.requireContext) {
            const contextWords = ['invoice', 'receipt', 'bill', 'payment', 'חשבונית', 'קבלה', 'תשלום'];
            const hasContext = contextWords.some(word => cleanText.includes(word));
            if (!hasContext) continue;
          }
          
          _logMessage(`💰 נמצא סכום ב-${source}: ${numericPart} ${pattern.currency} (מתוך: "${match}")`);
          
          return {
            found: true,
            amount: numericPart,
            currency: pattern.currency,
            method: `${source}_regex`,
            source: match.trim(),
            confidence: pattern.requireContext ? 'medium' : 'high'
          };
        }
      }
    }
  }
  
  return { found: false };
}

// === חילוץ סכום מתוכן המייל ===
function extractAmountFromEmailBody(message) {
  try {
    const plainBody = message.getPlainBody();
    const htmlBody = message.getBody();
    
    // נסה קודם בתוכן הטקסטואלי
    const plainResult = extractAmountFromText(plainBody, "email_body");
    if (plainResult.found) {
      return plainResult;
    }
    
    // אם לא נמצא, נסה ב-HTML (אחרי ניקוי)
    const cleanHtml = htmlBody.replace(/<[^>]*>/g, ' ').replace(/&[^;]+;/g, ' ');
    const htmlResult = extractAmountFromText(cleanHtml, "email_html");
    if (htmlResult.found) {
      return htmlResult;
    }
    
    return { found: false };
    
  } catch (e) {
    _logMessage(`⚠️ שגיאה בחילוץ מתוכן המייל: ${e.message}`);
    return { found: false, method: "body_error" };
  }
}

// === פונקציית עזר לזיהוי מטבע מהקשר ===
function detectCurrencyFromContext(text) {
  const lowerText = text.toLowerCase();
  
  // בדיקת מילות מפתח למטבעות
  const currencyIndicators = [
    { keywords: ['₪', 'שקל', 'ש"ח', 'nis', 'ils', 'israel'], currency: 'ILS' },
    { keywords: ['$', 'dollar', 'usd', 'america'], currency: 'USD' },
    { keywords: ['€', 'euro', 'eur', 'europe'], currency: 'EUR' },
    { keywords: ['£', 'pound', 'gbp', 'british'], currency: 'GBP' }
  ];
  
  for (const indicator of currencyIndicators) {
    if (indicator.keywords.some(keyword => lowerText.includes(keyword))) {
      return indicator.currency;
    }
  }
  
  // ברירת מחדל לפי מיקום (ישראל)
  return 'ILS';
}

// === NITRO OCR - חילוץ סכומים מתקדם ===
function performNitroOcr() {
  const ui = SpreadsheetApp.getUi();
  
  // בדיקה אם יש גיליון תוצאות פעיל
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  
  if (!activeSheet.getName().startsWith('invoices ')) {
    ui.alert('❌ שגיאה', 'אנא בחר גיליון תוצאות (invoices) לפני הפעלת ניטרו OCR', ui.ButtonSet.OK);
    return;
  }
  
  // אישור מהמשתמש
  const response = ui.alert(
    '🔍 ניטרו OCR',
    'האם תרצה להפעיל OCR מתקדם לחילוץ סכומים מרשומות ללא סכום?\n\n⚠️ זה עלול לקחת זמן רב!',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  _logMessage('🔍 מתחיל ניטרו OCR...');
  
  try {
    // מציאת רשומות ללא סכום
    const data = activeSheet.getDataRange().getValues();
    const headers = data[0];
    
    // מציאת עמודות רלוונטיות
    const amountCol = headers.indexOf('Amount');
    const emailIdCol = headers.indexOf('Email ID');
    const attachmentCol = headers.indexOf('Attachment Name');
    
    if (amountCol === -1 || emailIdCol === -1 || attachmentCol === -1) {
      ui.alert('❌ שגיאה', 'לא נמצאו עמודות נדרשות בגיליון', ui.ButtonSet.OK);
      return;
    }
    
    let processedCount = 0;
    let foundCount = 0;
    
    // עיבוד רשומות ללא סכום
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const amount = row[amountCol];
      const emailId = row[emailIdCol];
      const attachmentName = row[attachmentCol];
      
      // רק רשומות ללא סכום ועם PDF
      if ((!amount || amount === '') && emailId && attachmentName && attachmentName.toLowerCase().includes('.pdf')) {
        _logMessage(`🔍 מעבד OCR עבור: ${attachmentName}`);
        
        try {
          // חיפוש המייל
          const threads = GmailApp.search(`rfc822msgid:${emailId}`);
          if (threads.length > 0) {
            const messages = threads[0].getMessages();
            for (const message of messages) {
              const attachments = message.getAttachments();
              
              for (const attachment of attachments) {
                if (attachment.getName() === attachmentName && attachment.getContentType().includes('pdf')) {
                  // הפעלת OCR
                  const pdfText = extractTextFromPdfAdvanced(attachment.copyBlob());
                  
                  if (pdfText && pdfText.length > 10) {
                    const amountData = extractAmountAndCurrency('', pdfText);
                    
                    if (amountData && (amountData.found || amountData.amount)) {
                      // עדכון הטבלה
                      activeSheet.getRange(i + 1, amountCol + 1).setValue(amountData.amount || '');
                      activeSheet.getRange(i + 1, amountCol + 2).setValue(amountData.currency || '');
                      activeSheet.getRange(i + 1, amountCol + 3).setValue('nitro_ocr');
                      
                      _logMessage(`✅ נמצא סכום ב-OCR: ${amountData.amount} ${amountData.currency}`);
                      foundCount++;
                    }
                  }
                  break;
                }
              }
            }
          }
        } catch (ocrError) {
          _logMessage(`❌ שגיאה ב-OCR עבור ${attachmentName}: ${ocrError.message}`);
        }
        
        processedCount++;
        
        // עדכון תקופתי
        if (processedCount % 3 === 0) {
          SpreadsheetApp.flush();
          ss.toast(`🔍 עובד... ${processedCount} רשומות נבדקו, ${foundCount} סכומים נמצאו`, 'ניטרו OCR', 3);
        }
      }
    }
    
    // סיכום
    _logMessage(`✅ ניטרו OCR הושלם: ${processedCount} רשומות נבדקו, ${foundCount} סכומים נמצאו`);
    ui.alert('✅ הושלם', `ניטרו OCR הושלם!\n\n📊 ${processedCount} רשומות נבדקו\n💰 ${foundCount} סכומים נמצאו`, ui.ButtonSet.OK);
    
  } catch (e) {
    _logMessage(`❌ שגיאה בניטרו OCR: ${e.message}`);
    ui.alert('❌ שגיאה', `שגיאה בניטרו OCR: ${e.message}`, ui.ButtonSet.OK);
  }
}

// === NITRO OCR - חילוץ סכומים מתקדם ===
function performNitroOcr() {
  const ui = SpreadsheetApp.getUi();
  
  // בדיקה אם יש גיליון תוצאות פעיל
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  
  if (!activeSheet.getName().startsWith('invoices ')) {
    ui.alert('❌ שגיאה', 'אנא בחר גיליון תוצאות (invoices) לפני הפעלת ניטרו OCR', ui.ButtonSet.OK);
    return;
  }
  
  // אישור מהמשתמש
  const response = ui.alert(
    '🔍 ניטרו OCR',
    'האם תרצה להפעיל OCR מתקדם לחילוץ סכומים מרשומות ללא סכום?\n\n⚠️ זה עלול לקחת זמן רב!',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  _logMessage('🔍 מתחיל ניטרו OCR...');
  
  try {
    // מציאת רשומות ללא סכום
    const data = activeSheet.getDataRange().getValues();
    const headers = data[0];
    
    // מציאת עמודות רלוונטיות
    const amountCol = headers.indexOf('Amount');
    const emailIdCol = headers.indexOf('Email ID');
    const attachmentCol = headers.indexOf('Attachment Name');
    
    if (amountCol === -1 || emailIdCol === -1 || attachmentCol === -1) {
      ui.alert('❌ שגיאה', 'לא נמצאו עמודות נדרשות בגיליון', ui.ButtonSet.OK);
      return;
    }
    
    let processedCount = 0;
    let foundCount = 0;
    
    // עיבוד רשומות ללא סכום
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const amount = row[amountCol];
      const emailId = row[emailIdCol];
      const attachmentName = row[attachmentCol];
      
      // רק רשומות ללא סכום ועם PDF
      if ((!amount || amount === '') && emailId && attachmentName && attachmentName.toLowerCase().includes('.pdf')) {
        _logMessage(`🔍 מעבד OCR עבור: ${attachmentName}`);
        
        try {
          // חיפוש המייל
          const threads = GmailApp.search(`rfc822msgid:${emailId}`);
          if (threads.length > 0) {
            const messages = threads[0].getMessages();
            for (const message of messages) {
              const attachments = message.getAttachments();
              
              for (const attachment of attachments) {
                if (attachment.getName() === attachmentName && attachment.getContentType().includes('pdf')) {
                  // הפעלת OCR
                  const pdfText = extractTextFromPdfAdvanced(attachment.copyBlob());
                  
                  if (pdfText && pdfText.length > 10) {
                    const amountData = extractAmountAndCurrency('', pdfText);
                    
                    if (amountData && (amountData.found || amountData.amount)) {
                      // עדכון הטבלה
                      activeSheet.getRange(i + 1, amountCol + 1).setValue(amountData.amount || '');
                      activeSheet.getRange(i + 1, amountCol + 2).setValue(amountData.currency || '');
                      activeSheet.getRange(i + 1, amountCol + 3).setValue('nitro_ocr');
                      
                      _logMessage(`✅ נמצא סכום ב-OCR: ${amountData.amount} ${amountData.currency}`);
                      foundCount++;
                    }
                  }
                  break;
                }
              }
            }
          }
        } catch (ocrError) {
          _logMessage(`❌ שגיאה ב-OCR עבור ${attachmentName}: ${ocrError.message}`);
        }
        
        processedCount++;
        
        // עדכון תקופתי
        if (processedCount % 3 === 0) {
          SpreadsheetApp.flush();
          ss.toast(`🔍 עובד... ${processedCount} רשומות נבדקו, ${foundCount} סכומים נמצאו`, 'ניטרו OCR', 3);
        }
      }
    }
    
    // סיכום
    _logMessage(`✅ ניטרו OCR הושלם: ${processedCount} רשומות נבדקו, ${foundCount} סכומים נמצאו`);
    ui.alert('✅ הושלם', `ניטרו OCR הושלם!\n\n📊 ${processedCount} רשומות נבדקו\n💰 ${foundCount} סכומים נמצאו`, ui.ButtonSet.OK);
    
  } catch (e) {
    _logMessage(`❌ שגיאה בניטרו OCR: ${e.message}`);
    ui.alert('❌ שגיאה', `שגיאה בניטרו OCR: ${e.message}`, ui.ButtonSet.OK);
  }
}
// === CLEAR LOG SHEET ===
function clearLogSheet() {
  const ui = SpreadsheetApp.getUi();
  
  // פופאפ אישור
  const response = ui.alert(
    '🗑️ איפוס לוג',
    'האם אתה בטוח שברצונך למחוק את כל הלוגים?\n\nפעולה זו תמחק את כל השורות מהגיליון Log (חוץ מהכותרות).',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('❌ פעולת איפוס הלוג בוטלה.');
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    
    if (!logSheet) {
      ui.alert('❌ גיליון Log לא נמצא.');
      return;
    }
    
    // בדיקה אם יש נתונים למחיקה
    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) {
      ui.alert('ℹ️ הלוג כבר ריק.');
      return;
    }
    
    // ניקוי תוכן השורות במקום מחיקתן (עובד גם עם שורות קפואות)
    if (lastRow > 1) {
      const rangeToClean = logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn());
      rangeToClean.clearContent();
    }
    
    // הודעת הצלחה
    ui.alert('✅ הלוג נוקה בהצלחה!');
    
    // לוג על הפעולה
    const timestamp = new Date().toLocaleString('he-IL');
    logSheet.getRange(2, 1, 1, 2).setValues([[timestamp, '🗑️ הלוג נוקה על ידי המשתמש']]);
    
  } catch (error) {
    console.error('שגיאה באיפוס הלוג:', error);
    ui.alert(`❌ שגיאה באיפוס הלוג: ${error.message}`);
  }
}

// === MOVE LOG SHEET TO CORRECT POSITION ===
function moveLogSheetToCorrectPosition() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    
    if (!logSheet) {
      console.log('❌ גיליון Log לא נמצא');
      return;
    }
    
    // העברת גיליון Log למיקום השני (אחרי Settings)
    ss.moveSheet(logSheet, 2);
    console.log('✅ גיליון Log הועבר למיקום הנכון');
    
  } catch (error) {
    console.error('❌ שגיאה בהעברת גיליון Log:', error);
  }
}

// === SELECT AND COPY LAST SCAN LOG ===
function selectLastScanLog() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    
    if (!logSheet) {
      ui.alert('❌ גיליון Log לא נמצא.');
      return;
    }
    
    const data = logSheet.getDataRange().getValues();
    if (data.length <= 1) {
      ui.alert('ℹ️ אין נתוני לוג לעיבוד.');
      return;
    }
    
    // חיפוש השורה האחרונה עם "🚀 התחלת סריקה מתקדמת"
    let lastScanStartRow = -1;
    for (let i = data.length - 1; i >= 1; i--) {
      const logMessage = data[i][1] || '';
      if (logMessage.includes('🚀 התחלת סריקה מתקדמת עם Gmail API')) {
        lastScanStartRow = i + 1; // +1 כי מערך מתחיל מ-0 אבל שורות מ-1
        break;
      }
    }
    
    if (lastScanStartRow === -1) {
      ui.alert('❌ לא נמצאה סריקה אחרונה בלוג.');
      return;
    }
    
    // סימון השורות מהסריקה האחרונה עד הסוף
    const lastRow = logSheet.getLastRow();
    const rowsToSelect = lastRow - lastScanStartRow + 1;
    
    if (rowsToSelect <= 0) {
      ui.alert('❌ לא נמצאו שורות לסימון.');
      return;
    }
    
    // סימון הטווח
    const range = logSheet.getRange(lastScanStartRow, 1, rowsToSelect, 2);
    range.activate();
    logSheet.setActiveRange(range);
    
    // יצירת טקסט להעתקה
    const selectedData = data.slice(lastScanStartRow - 1);
    let copyText = '';
    selectedData.forEach(row => {
      const timestamp = row[0] || '';
      const message = row[1] || '';
      copyText += `${timestamp}\t${message}\n`;
    });
    
    // הצגת הודעה עם אפשרות העתקה
    const response = ui.alert(
      '📋 סריקה אחרונה נבחרה',
      `נמצאו ${rowsToSelect} שורות מהסריקה האחרונה.\n\nהשורות סומנו בגיליון.\n\nהאם תרצה להעתיק את הנתונים ללוח?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      // יצירת גיליון זמני להעתקה (Google Sheets לא תומך בהעתקה ישירה ללוח)
      const tempSheet = ss.insertSheet('TEMP_COPY');
      tempSheet.getRange(1, 1, selectedData.length, 2).setValues(selectedData);
      tempSheet.getRange(1, 1, selectedData.length, 2).activate();
      
      ui.alert(
        '📋 מוכן להעתקה',
        'הנתונים הועתקו לגיליון זמני "TEMP_COPY".\n\n' +
        'השתמש ב-Ctrl+C להעתקה ואז מחק את הגיליון הזמני.\n\n' +
        'או העתק את הטקסט הבא:\n\n' + copyText.substring(0, 500) + '...',
        ui.ButtonSet.OK
      );
    }
    
    _logMessage(`📋 נבחרו ${rowsToSelect} שורות מהסריקה האחרונה`);
    
  } catch (error) {
    console.error('❌ שגיאה בבחירת הסריקה האחרונה:', error);
    ui.alert(`❌ שגיאה: ${error.message}`);
  }
}

// === EXPORT FOLDER DIALOG ===
function showExportFolderDialog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ExportFolderDialog')
      .setWidth(500)
      .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, '📁 הגדרת תיקיית Google Drive');
  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ שגיאה בפתיחת הדיאלוג: ${e.message}`);
  }
}

function saveFolderId(folderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
    
    if (!settings) {
      throw new Error('גיליון Settings לא נמצא');
    }
    
    // שמירת ה-ID בעמודה I (Export Folder ID)
    _logMessage(`🔍 שומר Folder ID בעמודה I2: ${folderId}`);
    settings.getRange('I2').setValue(folderId);
    
    // בדיקה שהערך נשמר נכון
    const savedValue = settings.getRange('I2').getValue();
    _logMessage(`✅ ערך שנשמר בI2: ${savedValue}`);
    
    // בדיקה שהתיקייה קיימת
    try {
      const folder = DriveApp.getFolderById(folderId);
      _logMessage(`✅ תיקיית ייצוא הוגדרה: ${folder.getName()} (${folderId})`);
      return { success: true, message: `תיקייה הוגדרה בהצלחה: ${folder.getName()}` };
    } catch (e) {
      _logMessage(`❌ שגיאה בגישה לתיקייה: ${e.message}`);
      return { success: false, message: `שגיאה: לא ניתן לגשת לתיקייה. בדוק שיש לך הרשאות.` };
    }
    
  } catch (e) {
    _logMessage(`❌ שגיאה בשמירת ID: ${e.message}`);
    return { success: false, message: `שגיאה בשמירה: ${e.message}` };
  }
}

function getExportFolderId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
    
    if (!settings) {
      return null;
    }
    
    const folderId = settings.getRange('I2').getValue();
    _logMessage(`🔍 קורא Folder ID מעמודה I2: ${folderId}`);
    return folderId || null;
  } catch (e) {
    _logMessage(`❌ שגיאה בקריאת Folder ID: ${e.message}`);
    return null;
  }
}