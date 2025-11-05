// === SheetSetup.gs - הגדרת גיליונות ===

// === SETUP FUNCTIONS ===
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
    log.setColumnWidth(1, 150); // Timestamp
    log.setColumnWidth(2, 800); // Message - רחב מאוד
  }

  SpreadsheetApp.setActiveSheet(settings);
  
  // העברת גיליון Log למיקום הנכון (טאב שני)
  moveLogSheetToCorrectPosition();
  
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

function moveLogSheetToCorrectPosition() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (logSheet) {
    ss.moveSheet(logSheet, 2); // מעביר לעמדה השנייה
  }
}