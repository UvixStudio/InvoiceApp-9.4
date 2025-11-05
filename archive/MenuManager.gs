// === MenuManager.gs - ניהול תפריטים ===

// === GLOBALS ===
const SETTINGS_SHEET_NAME = "Settings";
const LOG_SHEET_NAME = "Log";

// === MENU FUNCTIONS ===
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
        {name: "🚀 Scan Invoices (API)", functionName: "processInvoicesGmailApi"}
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
    .addItem("🚀 סריקת חשבוניות (Gmail API)", "processInvoicesGmailApi")
    .addSeparator()
    .addItem("🔎 אבחון מפורט (לוג יפה)", "diagnosticScan")
    .addItem("📊 אבחון טבלה", "diagnosticAdvanced")
    .addSeparator()
    .addItem("📋 בחר סריקה אחרונה מהלוג", "selectLastScanLog")
    .addItem("🗑️ איפוס לוג", "clearLogSheet")
    .addSeparator()
    .addItem("🔄 איפוס כל הגיליונות", "resetAllTabs")
    .addItem("🆘 איפוס מערכת (משתמש חדש)", "systemReset");
    
  mainMenu.addToUi();
}

function createActionsMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const actionsMenu = ui.createMenu("⚡ פעולות")
    .addItem("✔️ סמן הכל", "selectAllRows")
    .addItem("❌ נקה סימון", "clearAllSelections")
    .addSeparator()
    .addItem("🚫 החרג נבחרים", "excludeSelected")
    .addItem("🔷 סמן כמקומי", "markAsLocal")
    .addItem("🌐 סמן כבינלאומי", "markAsInternational")
    .addSeparator()
    .addItem("📦 יצא נבחרים", "exportSelected")
    .addItem("📤 חלץ הכל", "syncInvoicesToDrive")
    .addItem("🗑️ מחק נבחרים", "deleteSelected")
    .addSeparator()
    .addItem("✅ הוסף נבחרים למאושרים", "approveSelected")
    .addSeparator()
    .addItem("🧹 נקה גיליון", "clearActiveSheetData")
    .addItem("🔄 רענן תצוגה", "refreshView");
    
  actionsMenu.addToUi();
}