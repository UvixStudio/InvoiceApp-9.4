// === LogManager.gs - ניהול לוגים ===

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

function clearLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  if (log) {
    log.clear();
    log.appendRow(["Timestamp", "Message"]);
    log.setFrozenRows(1);
    log.setColumnWidth(1, 150);
    log.setColumnWidth(2, 800);
    SpreadsheetApp.getUi().alert("✅ לוג נוקה בהצלחה");
  }
}

function selectLastScanLog() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = ss.getSheetByName(LOG_SHEET_NAME);
    
    if (!log) {
      ui.alert("❌ גיליון Log לא נמצא");
      return;
    }
    
    const data = log.getDataRange().getValues();
    if (data.length <= 1) {
      ui.alert("❌ אין נתונים בלוג");
      return;
    }
    
    // חיפוש הסריקה האחרונה
    let lastScanStart = -1;
    let lastScanEnd = -1;
    
    for (let i = data.length - 1; i >= 1; i--) {
      const message = data[i][1];
      if (message && message.includes("✅ סריקה הושלמה")) {
        lastScanEnd = i + 1;
        break;
      }
    }
    
    if (lastScanEnd === -1) {
      ui.alert("❌ לא נמצאה סריקה שהושלמה בלוג");
      return;
    }
    
    // חיפוש תחילת הסריקה
    for (let i = lastScanEnd - 1; i >= 1; i--) {
      const message = data[i][1];
      if (message && (message.includes("🚀 התחלת סריקה") || message.includes("🔍 מתחיל סריקה"))) {
        lastScanStart = i + 1;
        break;
      }
    }
    
    if (lastScanStart === -1) {
      lastScanStart = Math.max(1, lastScanEnd - 50); // ברירת מחדל - 50 שורות אחורה
    }
    
    // בחירת הטווח
    const range = log.getRange(lastScanStart, 1, lastScanEnd - lastScanStart + 1, 2);
    range.activate();
    log.setActiveRange(range);
    
    // העתקה ללוח
    range.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    ui.alert(`✅ נבחרו ${lastScanEnd - lastScanStart + 1} שורות מהסריקה האחרונה\n\nהטווח נבחר ומועתק ללוח`);
    
  } catch (e) {
    ui.alert(`❌ שגיאה בבחירת לוג: ${e.message}`);
  }
}