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
  
  _logMessage(log, "🔄 התחלת סנכרון חכם ל-Drive", 'INFO');
  
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
    const syncResults = performSmartSync(targetFolder, log);
    
    // הצגת תוצאות
    showSyncResults(syncResults);
    
    _logMessage(log, "✅ סנכרון הושלם בהצלחה", 'SUCCESS');
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בסנכרון: ${error.message}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ שגיאה בסנכרון: ${error.message}`);
  }
}

/**
 * ביצוע סנכרון חכם
 */
function performSmartSync(targetFolder, log) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pattern = /^invoices /i;
  const sheets = ss.getSheets().filter(sheet => pattern.test(sheet.getName()));
  
  let totalProcessed = 0;
  let totalSynced = 0;
  let totalSkipped = 0;
  let totalErrors = 0;
  
  // קבלת קבצים קיימים בתיקייה
  const existingFiles = getExistingFiles(targetFolder);
  
  for (const sheet of sheets) {
    _logMessage(log, `🔄 מעבד גיליון: ${sheet.getName()}`, 'INFO');
    
    const sheetResults = syncSheetToFolder(sheet, targetFolder, existingFiles, log);
    
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
    sheetsProcessed: sheets.length
  };
}

/**
 * סנכרון גיליון בודד לתיקייה
 */
function syncSheetToFolder(sheet, targetFolder, existingFiles, log) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
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
  
  // יצירת תת-תיקייה לגיליון
  const sheetFolderName = sanitizeFileName(sheet.getName());
  let sheetFolder = getOrCreateSubfolder(targetFolder, sheetFolderName);
  
  // עבור על כל השורות
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    processed++;
    
    try {
      const pdfLink = row[pdfLinkIndex];
      const senderName = row[senderNameIndex] || 'Unknown';
      const subject = row[subjectIndex] || 'No Subject';
      const date = row[dateIndex];
      const category = row[categoryIndex];
      
      if (!pdfLink || !pdfLink.trim()) {
        skipped++;
        continue;
      }
      
      // יצירת שם קובץ ייחודי
      const fileName = createUniqueFileName(senderName, subject, date, category);
      
      // בדיקה אם הקובץ כבר קיים
      if (existingFiles.has(fileName)) {
        skipped++;
        continue;
      }
      
      // ניסיון סנכרון הקובץ
      if (syncInvoiceFile(pdfLink, fileName, sheetFolder, log)) {
        synced++;
        existingFiles.add(fileName);
      } else {
        errors++;
      }
      
    } catch (error) {
      errors++;
      _logMessage(log, `❌ שגיאה בעיבוד שורה ${i + 1}: ${error.message}`, 'ERROR');
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
function createUniqueFileName(senderName, subject, date, category) {
  // ניקוי שמות
  const cleanSender = sanitizeFileName(senderName || 'Unknown').substring(0, 20);
  const cleanSubject = sanitizeFileName(subject || 'No_Subject').substring(0, 30);
  
  // פורמט תאריך
  let dateStr = '';
  if (date && date instanceof Date) {
    dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else if (date && typeof date === 'string') {
    dateStr = date.substring(0, 10);
  } else {
    dateStr = new Date().toISOString().split('T')[0];
  }
  
  // קטגוריה
  const categoryPrefix = category === '🔷 Local' ? 'IL' : 'INT';
  
  // שם הקובץ הסופי
  return `${categoryPrefix}_${dateStr}_${cleanSender}_${cleanSubject}.pdf`.replace(/_{2,}/g, '_');
}

/**
 * סנכרון קובץ חשבונית בודד
 */
function syncInvoiceFile(pdfLink, fileName, targetFolder, log) {
  try {
    if (pdfLink.startsWith('PDF:')) {
      // זה PDF מצורף - כרגע לא נוכל להוריד אותו ישירות
      // נוצר קובץ טקסט עם המידע
      const textContent = `Invoice PDF: ${pdfLink}\nFile would be: ${fileName}`;
      const textBlob = Utilities.newBlob(textContent, 'text/plain', fileName.replace('.pdf', '.txt'));
      targetFolder.createFile(textBlob);
      return true;
    } else if (pdfLink.startsWith('http')) {
      // זה קישור - ננסה להוריד
      try {
        const response = UrlFetchApp.fetch(pdfLink);
        const blob = response.getBlob();
        blob.setName(fileName);
        targetFolder.createFile(blob);
        return true;
      } catch (fetchError) {
        // אם ההורדה נכשלה, נוצר קובץ טקסט עם הקישור
        const textContent = `Invoice Link: ${pdfLink}\nFailed to download: ${fetchError.message}`;
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
 * הצגת תוצאות הסנכרון
 */
function showSyncResults(results) {
  const message = `✅ סנכרון הושלם!

📊 תוצאות:
• ${results.sheetsProcessed} גיליונות עובדו
• ${results.totalProcessed} רשומות נבדקו
• ${results.totalSynced} קבצים חדשים נוספו
• ${results.totalSkipped} קבצים דולגו (כבר קיימים)
• ${results.totalErrors} שגיאות

${results.totalSynced > 0 ? '🎉 הקבצים זמינים בתיקיית Google Drive!' : ⚠️ לא נוספו קבצים חדשים.'}`;

  SpreadsheetApp.getUi().alert("סיכום סנכרון", message, SpreadsheetApp.getUi().ButtonSet.OK);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ סנכרון הושלם: ${results.totalSynced} נוספו, ${results.totalSkipped} דולגו`,
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