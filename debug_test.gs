// === פונקציית דיבאגינג מיוחדת לבדיקת כתיבה לגיליון ===

function testSheetWriting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    SpreadsheetApp.getUi().alert("❌ גיליון Log לא נמצא");
    return;
  }
  
  logMessage(log, "🧪 התחלת בדיקת כתיבה לגיליון", 'INFO');
  
  try {
    // יצירת גיליון בדיקה
    let testSheet = ss.getSheetByName("TEST_WRITE");
    if (testSheet) {
      ss.deleteSheet(testSheet);
    }
    
    testSheet = ss.insertSheet("TEST_WRITE");
    logMessage(log, "✅ גיליון בדיקה נוצר", 'SUCCESS');
    
    // הוספת כותרות
    const headers = ["Action", "Category", "Sender Name", "Sender Email", "Date", "Subject"];
    testSheet.appendRow(headers);
    logMessage(log, `✅ כותרות נוספו: ${headers.join(', ')}`, 'SUCCESS');
    
    // בדיקת מספר שורות
    let rowCount = testSheet.getLastRow();
    logMessage(log, `📊 מספר שורות אחרי כותרות: ${rowCount}`, 'INFO');
    
    // הוספת רשומת בדיקה
    const testRecord = [false, "🔷 Local", "Test Sender", "test@example.com", "2025-07-16", "Test Subject"];
    
    logMessage(log, `🔧 מנסה להוסיף רשומת בדיקה: [${testRecord.join(', ')}]`, 'INFO');
    
    const beforeRows = testSheet.getLastRow();
    logMessage(log, `📊 שורות לפני הוספה: ${beforeRows}`, 'INFO');
    
    testSheet.appendRow(testRecord);
    SpreadsheetApp.flush();
    
    const afterRows = testSheet.getLastRow();
    logMessage(log, `📊 שורות אחרי הוספה: ${afterRows}`, 'INFO');
    
    if (afterRows > beforeRows) {
      logMessage(log, "✅ הוספת רשומה הצליחה!", 'SUCCESS');
      
      // קריאת הרשומה שנוספה
      const addedRow = testSheet.getRange(afterRows, 1, 1, testRecord.length).getValues()[0];
      logMessage(log, `📋 רשומה שנוספה: [${addedRow.join(', ')}]`, 'SUCCESS');
      
      // בדיקת הרשאות כתיבה נוספות
      try {
        testSheet.getRange(afterRows, 1).setBackground("#e8f5e8");
        logMessage(log, "✅ עיצוב תא הצליח", 'SUCCESS');
      } catch (formatError) {
        logMessage(log, `❌ שגיאה בעיצוב: ${formatError.message}`, 'ERROR');
      }
      
    } else {
      logMessage(log, "❌ הוספת רשומה נכשלה - מספר השורות לא השתנה", 'ERROR');
    }
    
    // בדיקת הוספה מרובה
    logMessage(log, "🔧 בודק הוספה מרובה...", 'INFO');
    
    const multipleRecords = [
      [false, "🌐 International", "Test Sender 2", "test2@example.com", "2025-07-16", "Test Subject 2"],
      [false, "🔷 Local", "Test Sender 3", "test3@example.com", "2025-07-16", "Test Subject 3"]
    ];
    
    const beforeMultiple = testSheet.getLastRow();
    
    for (let i = 0; i < multipleRecords.length; i++) {
      testSheet.appendRow(multipleRecords[i]);
      if (i % 1 === 0) { // Flush כל רשומה
        SpreadsheetApp.flush();
      }
    }
    
    const afterMultiple = testSheet.getLastRow();
    logMessage(log, `📊 הוספה מרובה: ${beforeMultiple} -> ${afterMultiple} (הוספו ${afterMultiple - beforeMultiple})`, 'INFO');
    
    if (afterMultiple - beforeMultiple === multipleRecords.length) {
      logMessage(log, "✅ הוספה מרובה הצליחה!", 'SUCCESS');
    } else {
      logMessage(log, `❌ הוספה מרובה נכשלה - צפוי ${multipleRecords.length}, התקבל ${afterMultiple - beforeMultiple}`, 'ERROR');
    }
    
    // סיכום
    const finalRowCount = testSheet.getLastRow();
    logMessage(log, `📊 סיכום: גיליון מכיל ${finalRowCount} שורות (כולל כותרת)`, 'INFO');
    
    SpreadsheetApp.getUi().alert(`בדיקת כתיבה הושלמה!\n\nגיליון TEST_WRITE נוצר עם ${finalRowCount} שורות.\nבדוק את גיליון Log לפרטים מלאים.`);
    
  } catch (error) {
    logMessage(log, `❌ שגיאה כללית בבדיקת כתיבה: ${error.message}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`שגיאה בבדיקת כתיבה: ${error.message}`);
  }
}

// פונקציה לבדיקת הרשאות מפורטת
function checkDetailedPermissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    SpreadsheetApp.getUi().alert("❌ גיליון Log לא נמצא");
    return;
  }
  
  logMessage(log, "🔐 התחלת בדיקת הרשאות מפורטת", 'INFO');
  
  // בדיקת הרשאות בסיסיות
  try {
    const spreadsheetId = ss.getId();
    logMessage(log, `📋 מזהה גיליון: ${spreadsheetId}`, 'INFO');
    
    const sheets = ss.getSheets();
    logMessage(log, `📊 מספר גיליונות: ${sheets.length}`, 'INFO');
    
    const activeSheet = ss.getActiveSheet();
    logMessage(log, `📄 גיליון פעיל: ${activeSheet.getName()}`, 'INFO');
    
  } catch (error) {
    logMessage(log, `❌ שגיאה בבדיקת הרשאות בסיסיות: ${error.message}`, 'ERROR');
  }
  
  // בדיקת הרשאות Gmail
  try {
    const threads = GmailApp.getInboxThreads(0, 1);
    logMessage(log, `📧 הרשאות Gmail: ✅ תקינות (${threads.length} שרשורים)`, 'SUCCESS');
  } catch (error) {
    logMessage(log, `❌ שגיאה בהרשאות Gmail: ${error.message}`, 'ERROR');
  }
  
  // בדיקת הרשאות Drive
  try {
    const folders = DriveApp.getFolders();
    let folderCount = 0;
    while (folders.hasNext() && folderCount < 5) {
      folders.next();
      folderCount++;
    }
    logMessage(log, `📁 הרשאות Drive: ✅ תקינות (${folderCount}+ תיקיות)`, 'SUCCESS');
  } catch (error) {
    logMessage(log, `❌ שגיאה בהרשאות Drive: ${error.message}`, 'ERROR');
  }
  
  logMessage(log, "🔐 בדיקת הרשאות הושלמה", 'SUCCESS');
  SpreadsheetApp.getUi().alert("בדיקת הרשאות הושלמה. בדוק את גיליון Log לפרטים.");
}