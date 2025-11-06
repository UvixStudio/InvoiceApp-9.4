// === DriveSyncer.gs - סנכרון עם Google Drive ===

/**
 * סנכרון חכם של חשבוניות ל-Google Drive
 */
function syncInvoicesToDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!settings || !log) {
    SpreadsheetApp.getUi().alert("❌ גיליונות Settings או Log חסרים.");
    return;
  }
  
  // מדידת זמן התחלה
  const startTime = new Date();
  _logMessage(log, `🔄 התחלת סנכרון חכם ל-Drive [${startTime.toLocaleTimeString()}]`, 'INFO');
  
  // הודעת התחלה
  SpreadsheetApp.getActiveSpreadsheet().toast("🔄 מתחיל סנכרון לדרייב...", "סנכרון", 3);
  
  try {
    // קבלת מזהה תיקיית היעד
    const folderId = settings.getRange("I2").getValue();
    
    if (!folderId) {
      SpreadsheetApp.getUi().alert("❌ לא הוגדרה תיקיית יעד. אנא הגדר תיקיית יצוא תחילה.");
      return;
    }
    
    // בדיקת גישה לתיקייה
    let targetFolder;
    try {
      targetFolder = DriveApp.getFolderById(folderId);
    } catch (e) {
      SpreadsheetApp.getUi().alert("❌ לא ניתן לגשת לתיקיית היעד. בדוק את מזהה התיקייה.");
      _logMessage(log, `❌ שגיאה בגישה לתיקייה: ${e.message}`, 'ERROR');
      return;
    }
    
    _logMessage(log, `📁 מסנכרן לתיקייה: ${targetFolder.getName()}`, 'INFO');
    
    // איסוף חשבוניות מכל הגיליונות
    const syncResults = performSmartSync(targetFolder, log, startTime);
    
    // חישוב זמן ביצוע
    const endTime = new Date();
    const durationMs = endTime - startTime;
    const durationSec = Math.round(durationMs / 1000);
    
    // הוספת זמן לתוצאות
    syncResults.startTime = startTime;
    syncResults.endTime = endTime;
    syncResults.durationSec = durationSec;
    
    // הצגת תוצאות
    showSyncResults(syncResults);
    
    _logMessage(log, `✅ סנכרון הושלם בהצלחה [${endTime.toLocaleTimeString()}] - זמן ביצוע: ${durationSec} שניות`, 'SUCCESS');
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בסנכרון: ${error.message}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ שגיאה בסנכרון: ${error.message}`);
  }
}

/**
 * ביצוע סנכרון חכם
 */
function performSmartSync(targetFolder, log, startTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pattern = /^invoices /i;
  
  // עיבוד רק הגיליון הפעיל במקום כל הגיליונות
  const activeSheet = ss.getActiveSheet();
  
  // בדיקה שהגיליון הפעיל הוא גיליון חשבוניות
  if (!pattern.test(activeSheet.getName())) {
    throw new Error(`❌ הגיליון הנוכחי "${activeSheet.getName()}" אינו גיליון חשבוניות.\nאנא עבור לגיליון חשבוניות (invoices YYYY-MM-DD) ונסה שוב.`);
  }
  
  const sheets = [activeSheet]; // רק הגיליון הפעיל
  
  let totalProcessed = 0;
  let totalSynced = 0;
  let totalSkipped = 0;
  let totalErrors = 0;
  
  // קבלת קבצים קיימים בתיקייה
  const existingFiles = getExistingFiles(targetFolder);
  
  // ✅ דיווח על מצב התיקייה
  if (existingFiles.size === 0) {
    _logMessage(log, `📁 התיקייה ריקה - זו הפעם הראשונה שמסנכרנים`, 'INFO');
  } else {
    _logMessage(log, `📁 נמצאו ${existingFiles.size} קבצים קיימים בתיקייה`, 'INFO');
  }
  
  for (const sheet of sheets) {
    _logMessage(log, `🔄 מסנכרן גיליון פעיל: ${sheet.getName()}`, 'INFO');
    
    const sheetResults = syncSheetToFolder(sheet, targetFolder, existingFiles, log, startTime);
    
    totalProcessed += sheetResults.processed;
    totalSynced += sheetResults.synced;
    totalSkipped += sheetResults.skipped;
    totalErrors += sheetResults.errors;
  }
  
  return {
    totalProcessed,
    totalSynced,
    totalSkipped,
    totalErrors,
    sheetsProcessed: sheets.length,
    sheetName: activeSheet.getName() // שם הגיליון שעובד
  };
}

/**
 * סנכרון גיליון בודד לתיקייה
 */
function syncSheetToFolder(sheet, targetFolder, existingFiles, log, startTime) {
  const data = sheet.getDataRange().getValues();
  const headers = data[12];  // ✅ V32 FIX: כותרות בשורה 13 (אינדקס 12)
  
  const pdfLinkIndex = headers.indexOf("PDF Link");
  const senderNameIndex = headers.indexOf("Sender Name");
  const senderEmailIndex = headers.indexOf("Sender Email");
  const subjectIndex = headers.indexOf("Subject");
  const dateIndex = headers.indexOf("Date");
  const categoryIndex = headers.indexOf("Category");
  
  let processed = 0;
  let synced = 0;
  let skipped = 0;
  let errors = 0;
  
  // יצירת תת-תיקייה לפי טווח תאריכים (לא שם הגיליון)
  // חילוץ טווח התאריכים משם הגיליון (למשל: "invoices 01.09.25-30.09.25" → "01.09.25-30.09.25")
  const sheetName = sheet.getName();
  const dateRangeMatch = sheetName.match(/(\d{2}\.\d{2}\.\d{2}-\d{2}\.\d{2}\.\d{2})/);
  const sheetFolderName = dateRangeMatch ? dateRangeMatch[1] : sanitizeFileName(sheetName);
  let sheetFolder = getOrCreateSubfolder(targetFolder, sheetFolderName);
  
  // הודעת התחלת עיבוד
  const totalRows = data.length - 13; // ✅ V32 FIX: מינוס Mini-Log (11) + רווח (1) + כותרות (1) = 13
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `📋 מעבד גיליון: ${sheet.getName()} (${totalRows} רשומות)`, 
    "סנכרון", 3
  );
  
  // ✅ V32 FIX: עבור רק על שורות הנתונים (מ-14 ואילך, אחרי Mini-Log + כותרות)
  for (let i = 13; i < data.length; i++) {  // ← שורה 13 = כותרות (אינדקס 12), נתונים מ-14 (אינדקס 13)
    const row = data[i];
    processed++;
    
    // הצגת התקדמות כל 5 רשומות - טוסט קבוע יותר
    if (processed % 5 === 0) {
      const progressBar = createProgressBar(processed, totalRows);
      const elapsedTime = Math.round((new Date() - startTime) / 1000);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `🔍 בודק רשומה ${processed}/${totalRows} | ⏱️ ${formatDuration(elapsedTime)}\n${progressBar}`, 
        "סנכרון פעיל", 8
      );
    }
    
    try {
      const pdfLink = row[pdfLinkIndex];
      const senderName = row[senderNameIndex] || 'Unknown';
      const senderEmail = row[senderEmailIndex] || '';
      const subject = row[subjectIndex] || 'No Subject';
      const date = row[dateIndex];
      const category = row[categoryIndex];
      const attachmentName = row[headers.indexOf("Attachment Name")] || '';
      
      if (!pdfLink || !pdfLink.trim()) {
        skipped++;
        continue;
      }
      
      // יצירת שם קובץ ייחודי וחכם
      const fileName = createUniqueFileName(senderName, subject, date, category, senderEmail, attachmentName);
      
      // בדיקה אם הקובץ כבר קיים
      if (existingFiles.has(fileName)) {
        skipped++;
        continue;
      }
      
      // ניסיון סנכרון הקובץ - עם מדידת זמן
      const fileStartTime = new Date();
      if (syncInvoiceFile(pdfLink, fileName, sheetFolder, log)) {
        const fileEndTime = new Date();
        const fileDuration = Math.round((fileEndTime - fileStartTime) / 1000);
        
        synced++;
        existingFiles.add(fileName);
        
        // לוג מפורט לכל קובץ
        _logMessage(log, `📤 קובץ ${synced}: ${fileName} - ${formatDuration(fileDuration)}`, 'SUCCESS');
        
        // הצגת התקדמות כל קובץ - טוסט קבוע
        const progressBar = createProgressBar(synced, totalRows);
        const elapsedTime = Math.round((new Date() - startTime) / 1000);
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `📤 מסנכרן לדרייב... ${synced}/${totalRows} קבצים הועברו | ⏱️ ${formatDuration(elapsedTime)}\n${progressBar}`, 
          "סנכרון פעיל", 10
        );
      } else {
        const fileEndTime = new Date();
        const fileDuration = Math.round((fileEndTime - fileStartTime) / 1000);
        _logMessage(log, `❌ שגיאה בקובץ: ${fileName} - ${formatDuration(fileDuration)}`, 'ERROR');
        errors++;
      }
      
    } catch (error) {
      errors++;
      console.error(`שגיאה בשורה ${i + 1}: ${error.message}`);
    }
  }
  
  return { processed, synced, skipped, errors };
}

/**
 * קבלת רשימת קבצים קיימים בתיקייה
 */
function getExistingFiles(folder) {
  const existingFiles = new Set();
  
  try {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      existingFiles.add(file.getName());
    }
    
    // גם מתת-תיקיות
    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      const subFiles = subfolder.getFiles();
      while (subFiles.hasNext()) {
        const file = subFiles.next();
        existingFiles.add(file.getName());
      }
    }
  } catch (error) {
    console.error(`Error getting existing files: ${error.message}`);
  }
  
  return existingFiles;
}

/**
 * יצירה או קבלת תת-תיקייה
 */
function getOrCreateSubfolder(parentFolder, folderName) {
  try {
    const subfolders = parentFolder.getFoldersByName(folderName);
    if (subfolders.hasNext()) {
      return subfolders.next();
    } else {
      return parentFolder.createFolder(folderName);
    }
  } catch (error) {
    console.error(`Error creating subfolder: ${error.message}`);
    return parentFolder; // fallback to parent folder
  }
}

/**
 * יצירת שם קובץ ייחודי
 */
function createUniqueFileName(senderName, subject, date, category, senderEmail, originalFileName) {
  // חילוץ שם חברה חכם
  const companyName = extractCompanyName(senderName, senderEmail, subject);
  
  // פורמט תאריך
  let dateStr = '';
  if (date && date instanceof Date) {
    dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else if (date && typeof date === 'string') {
    dateStr = date.substring(0, 10);
  } else {
    dateStr = new Date().toISOString().split('T')[0];
  }
  
  // חילוץ מספר חשבונית מהנושא
  const invoiceNumber = extractInvoiceNumber(subject);
  
  // חילוץ סוג המסמך
  const docType = extractDocumentType(subject);
  
  // בניית השם בפורמט יפה כמו שלך: תאריך - חברה - מספר - סוג
  const parts = [
    dateStr,
    companyName,
    invoiceNumber,
    docType
  ].filter(part => part && part.length > 0);
  
  return `${parts.join(' - ')}.pdf`;
}

/**
 * חילוץ שם חברה חכם מהמידע הזמין
 */
function extractCompanyName(senderName, senderEmail, subject) {
  // 1. חיפוש בנושא - מילות מפתח של חברות
  const subjectCompanies = {
    'כינרת': 'Kineret',
    'גילת': 'Gilat', 
    'טלפקום': 'Telecom',
    'בזק': 'Bezeq',
    'חברת חשמל': 'Electric',
    'מי אביבים': 'Water',
    'DigitalOcean': 'DigitalOcean',
    'Google': 'Google',
    'GoTo': 'GoTo'
  };
  
  for (const [hebrew, english] of Object.entries(subjectCompanies)) {
    if (subject && (subject.includes(hebrew) || subject.includes(english))) {
      return english;
    }
  }
  
  // 2. חילוץ מכתובת מייל (החלק לפני @)
  if (senderEmail && senderEmail.includes('@')) {
    const domain = senderEmail.split('@')[1];
    const domainParts = domain.split('.');
    
    // מיפוי דומיינים מוכרים
    const domainMapping = {
      'morning.co': 'Morning',
      'gilat.net': 'Gilat',
      'digitalocean.com': 'DigitalOcean',
      'google.com': 'Google',
      'gotoglobal.com': 'GoTo',
      'finbox.co.il': 'Finbox',
      'comax.co.il': 'Comax'
    };
    
    if (domainMapping[domain]) {
      return domainMapping[domain];
    }
    
    // אם לא נמצא מיפוי מדויק, נסה לחלץ שם חברה מהדומיין
    const mainDomain = domainParts[0];
    
    // ניקוי ושיפור שמות דומיין
    const cleanDomain = mainDomain
      .replace(/noreply|no-reply|donot|notify|support|info/gi, '') // הסרת מילים טכניות
      .replace(/[^a-zA-Z]/g, '') // רק אותיות
      .toLowerCase();
    
    if (cleanDomain.length >= 3) {
      // הפיכה לאות ראשונה גדולה
      return cleanDomain.charAt(0).toUpperCase() + cleanDomain.slice(1);
    }
    
    // אם הכל נמחק, קח את הדומיין המקורי
    return sanitizeFileName(mainDomain).substring(0, 10);
  }
  
  // 3. נסיון מ-sender name
  if (senderName && senderName !== 'Unknown') {
    return sanitizeFileName(senderName).substring(0, 10);
  }
  
  return 'Unknown';
}

/**
 * חילוץ מספר חשבונית מהנושא
 */
function extractInvoiceNumber(subject) {
  if (!subject) return '';
  
  // חיפוש מספרים בנושא (5-6 ספרות)
  const numberMatch = subject.match(/\b\d{4,6}\b/);
  if (numberMatch) {
    return numberMatch[0];
  }
  
  return '';
}

/**
 * חילוץ סוג המסמך מהנושא
 */
function extractDocumentType(subject) {
  if (!subject) return 'Document';
  
  const docTypes = {
    'הוראה מתקנת': 'הוראה מתקנת',
    'חשבונית חודשית': 'חשבונית חודשית', 
    'חשבונית': 'חשבונית',
    'קבלה': 'קבלה',
    'receipt': 'Receipt',
    'invoice': 'Invoice',
    'Your invoice': 'Invoice',
    'GoTo': 'GoTo Invoice'
  };
  
  // חיפוש סוג המסמך
  for (const [pattern, type] of Object.entries(docTypes)) {
    if (subject.includes(pattern)) {
      return type;
    }
  }
  
  // אם לא נמצא, נסה לקחת חלק מהנושא
  const cleanSubject = subject.replace(/\d+/g, '').trim();
  if (cleanSubject.length > 0) {
    return sanitizeFileName(cleanSubject).substring(0, 15);
  }
  
  return 'Document';
}

/**
 * סנכרון קובץ חשבונית בודד
 */
function syncInvoiceFile(pdfLink, fileName, targetFolder, log) {
  try {
    if (!pdfLink || !pdfLink.trim()) {
      return false;
    }
    
    // בדיקה אם זה קישור Google Drive - מהיר!
    if (pdfLink.includes('drive.google.com')) {
      try {
        // חילוץ מזהה הקובץ מהקישור
        const fileId = extractDriveFileId(pdfLink);
        if (fileId) {
          const sourceFile = DriveApp.getFileById(fileId);
          
          // שלב 1: קבלת רשימת תיקיות הורה לפני השינוי
          const originalParents = [];
          const parents = sourceFile.getParents();
          while (parents.hasNext()) {
            originalParents.push(parents.next());
          }
          
          // שלב 2: הוספה לתיקייה החדשה
          targetFolder.addFile(sourceFile);
          
          // שלב 3: הסרה רק מהתיקיות המקוריות (לא מהתיקייה החדשה)
          for (const parent of originalParents) {
            if (parent.getId() !== targetFolder.getId()) {
              parent.removeFile(sourceFile);
            }
          }
          
          // שלב 4: שינוי שם לפורמט מאורגן
          if (sourceFile.getName() !== fileName) {
            sourceFile.setName(fileName);
          }
          
          // שלב 5: לוג הצלחה (ללא אימות איטי)
          _logMessage(log, `✅ קובץ הועבר: ${fileName}`, 'SUCCESS');
          return true;
        }
      } catch (driveError) {
        console.error(`שגיאה בסנכרון: ${driveError.message}`);
      }
    }
    
    // אם זה קישור HTTP רגיל - ננסה להוריד
    if (pdfLink.startsWith('http')) {
      try {
        const response = UrlFetchApp.fetch(pdfLink, {
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
          },
          muteHttpExceptions: true
        });
        
        if (response.getResponseCode() === 200) {
          const blob = response.getBlob();
          blob.setName(fileName);
          targetFolder.createFile(blob);
          _logMessage(log, `✅ הורד מהאינטרנט: ${fileName}`, 'SUCCESS');
          return true;
        } else {
          throw new Error(`HTTP ${response.getResponseCode()}`);
        }
      } catch (fetchError) {
        _logMessage(log, `⚠️ שגיאה בהורדה: ${fetchError.message}`, 'WARNING');
        
        // יצירת קובץ טקסט עם הקישור כגיבוי
        const textContent = `Invoice Link: ${pdfLink}\nDownload failed: ${fetchError.message}\nDate: ${new Date().toISOString()}`;
        const textBlob = Utilities.newBlob(textContent, 'text/plain', fileName.replace('.pdf', '_link.txt'));
        targetFolder.createFile(textBlob);
        return true;
      }
    }
    
    return false;
  } catch (error) {
    _logMessage(log, `❌ שגיאה בסנכרון קובץ ${fileName}: ${error.message}`, 'ERROR');
    return false;
  }
}

/**
 * יצירת progress bar ויזואלי לטוסט
 */
function createProgressBar(current, total, width = 20) {
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
function formatDuration(seconds) {
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

// פונקציה לחילוץ מזהה קובץ מקישור Google Drive
function extractDriveFileId(driveUrl) {
  try {
    // דפוסים שונים של קישורי Google Drive
    const patterns = [
      /\/file\/d\/([a-zA-Z0-9-_]+)/,
      /id=([a-zA-Z0-9-_]+)/,
      /\/d\/([a-zA-Z0-9-_]+)/
    ];
    
    for (const pattern of patterns) {
      const match = driveUrl.match(pattern);
      if (match && match[1]) {
        return match[1];
      }
    }
    
    return null;
  } catch (error) {
    console.error(`Error extracting Drive file ID: ${error.message}`);
    return null;
  }
}

/**
 * ניקוי שם קובץ
 */
function sanitizeFileName(fileName) {
  if (!fileName || typeof fileName !== 'string') return 'unknown';
  
  return fileName
    .replace(/[<>:"\/\\|?*\x00-\x1F]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .trim();
}

/**
 * בדיקה מהירה של הגדרות הסנכרון
 */
function checkSyncSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!settings) {
    SpreadsheetApp.getUi().alert("❌ גיליון Settings לא נמצא");
    return;
  }
  
  const folderId = settings.getRange("I2").getValue();
  
  if (!folderId) {
    SpreadsheetApp.getUi().alert("❌ לא הוגדרה תיקיית יעד\n\nאנא הכנס מזהה תיקיית Google Drive בתא I2 בגיליון Settings");
    return;
  }
  
  try {
    const targetFolder = DriveApp.getFolderById(folderId);
    const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;
    
    SpreadsheetApp.getUi().alert(
      "✅ הגדרות סנכרון תקינות",
      `תיקיית היעד: ${targetFolder.getName()}\n\nקישור: ${folderUrl}\n\nהסנכרון מוכן לשימוש!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ שגיאה בגישה לתיקייה\n\nבדוק שמזהה התיקייה נכון ושיש לך הרשאות גישה");
  }
}

/**
 * הצגת תוצאות הסנכרון
 */
function showSyncResults(results) {
  const timeInfo = results.durationSec ? `⏱️ זמן ביצוע: ${formatDuration(results.durationSec)}` : '';
  
  const message = `✅ סנכרון הושלם!

📋 גיליון: ${results.sheetName || 'לא זוהה'}
${timeInfo}

📊 תוצאות:
• ${results.totalProcessed} רשומות נבדקו
• ${results.totalSynced} קבצים הועברו
• ${results.totalSkipped} קבצים דולגו (כבר קיימים)
• ${results.totalErrors} שגיאות

${results.totalSynced > 0 ? '🎉 הקבצים הועברו לתיקיית Google Drive!' : '⚠️ לא הועברו קבצים חדשים.'}`;

  SpreadsheetApp.getUi().alert("סיכום סנכרון", message, SpreadsheetApp.getUi().ButtonSet.OK);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ סנכרון ${results.sheetName}: ${results.totalSynced} הועברו (${formatDuration(results.durationSec)})`,
    "סנכרון הושלם", 10
  );
}

/**
 * ניקוי תיקיית Drive (מחיקת קבצים ישנים)
 */
function cleanupDriveFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!settings) {
    SpreadsheetApp.getUi().alert("❌ לא נמצא גיליון Settings.");
    return;
  }
  
  const folderId = settings.getRange("I2").getValue();
  if (!folderId) {
    SpreadsheetApp.getUi().alert("❌ לא הוגדרה תיקיית יעד.");
    return;
  }
  
  const response = SpreadsheetApp.getUi().alert(
    "אישור מחיקה",
    "האם אתה בטוח שברצונך למחוק את כל הקבצים בתיקיית הסנכרון?",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (response !== SpreadsheetApp.getUi().Button.YES) {
    return;
  }
  
  try {
    const targetFolder = DriveApp.getFolderById(folderId);
    let deletedCount = 0;
    
    // מחיקת קבצים
    const files = targetFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      file.setTrashed(true);
      deletedCount++;
    }
    
    // מחיקת תת-תיקיות
    const subfolders = targetFolder.getFolders();
    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      subfolder.setTrashed(true);
      deletedCount++;
    }
    
    _logMessage(log, `🗑️ נמחקו ${deletedCount} פריטים מתיקיית הסנכרון`, 'SUCCESS');
    SpreadsheetApp.getUi().alert(`✅ נמחקו ${deletedCount} פריטים מתיקיית הסנכרון.`);
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בניקוי תיקייה: ${error.message}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ שגיאה בניקוי תיקייה: ${error.message}`);
  }
}

/**
 * בדיקת מצב תיקיית הסנכרון
 */
function checkSyncFolderStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!settings) {
    SpreadsheetApp.getUi().alert("❌ לא נמצא גיליון Settings.");
    return;
  }
  
  const folderId = settings.getRange("I2").getValue();
  if (!folderId) {
    SpreadsheetApp.getUi().alert("❌ לא הוגדרה תיקיית יעד.");
    return;
  }
  
  try {
    const targetFolder = DriveApp.getFolderById(folderId);
    const files = targetFolder.getFiles();
    const subfolders = targetFolder.getFolders();
    
    let fileCount = 0;
    let folderCount = 0;
    
    while (files.hasNext()) {
      files.next();
      fileCount++;
    }
    
    while (subfolders.hasNext()) {
      subfolders.next();
      folderCount++;
    }
    
    SpreadsheetApp.getUi().alert(
      `📁 מצב תיקיית הסנכרון\n\n` +
      `שם התיקייה: ${targetFolder.getName()}\n` +
      `מספר קבצים: ${fileCount}\n` +
      `מספר תת-תיקיות: ${folderCount}\n` +
      `קישור: ${targetFolder.getUrl()}`
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ שגיאה בבדיקת תיקייה: ${error.message}`);
  }
}