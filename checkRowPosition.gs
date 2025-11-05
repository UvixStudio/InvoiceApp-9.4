// === בדיקת מיקום כתיבת רשומות ===

function checkRowPosition() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    SpreadsheetApp.getUi().alert("❌ גיליון Log לא נמצא");
    return;
  }
  
  try {
    // בדיקת השורה האחרונה עם נתונים אמיתיים
    const lastRowWithData = findLastRowWithData(activeSheet);
    const totalRows = activeSheet.getLastRow();
    const nextWriteRow = lastRowWithData + 1;
    
    const message = `🔍 בדיקת מיקום כתיבה:\n\n` +
      `📊 גיליון: ${activeSheet.getName()}\n` +
      `📝 שורה אחרונה עם נתונים: ${lastRowWithData}\n` +
      `📏 סה"כ שורות בגיליון: ${totalRows}\n` +
      `➡️ רשומה חדשה תיכתב בשורה: ${nextWriteRow}\n\n` +
      `${totalRows > lastRowWithData + 10 ? '⚠️ יש צ\'קבוקסים מיותרים שיש לנקות' : '✅ הגיליון נקי'}`;
    
    SpreadsheetApp.getUi().alert("בדיקת מיקום", message, SpreadsheetApp.getUi().ButtonSet.OK);
    
    // לוג הבדיקה
    _logMessage(log, `🔍 בדיקת מיקום: ${activeSheet.getName()} - נתונים עד שורה ${lastRowWithData}, הבא: ${nextWriteRow}`, 'INFO');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ שגיאה בבדיקה: ${error.message}`);
    _logMessage(log, `❌ שגיאה בבדיקת מיקום: ${error.message}`, 'ERROR');
  }
}

// פונקציה למציאת השורה האחרונה עם נתונים אמיתיים (העתק מ-gmailProcessor)
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