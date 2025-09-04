// === runActions.gs - פעולות על רשומות ===

// === BULK ACTIONS MAIN FUNCTION ===
function runBulkActions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();

  // בדיקה שזה גיליון חשבוניות
  if (!/^invoices /i.test(sheetName)) {
    SpreadsheetApp.getUi().alert("❌ אנא עבור לגיליון 'invoices' המתאים לפני הפעלת פעולות.");
    return;
  }

  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  if (!settings || !log) {
    SpreadsheetApp.getUi().alert("❌ גיליונות Settings או Log חסרים. אנא הפעל את 'Setup Sheets' תחילה.");
    return;
  }

  // קבלת נתונים מהגיליון
  const dataRange = activeSheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  const actionColumnIndex = headers.indexOf("Action");
  const emailColumnIndex = headers.indexOf("Sender Email");

  if (actionColumnIndex === -1 || emailColumnIndex === -1) {
    SpreadsheetApp.getUi().alert("❌ לא נמצאו עמודות 'Action' או 'Sender Email' בגיליון הנוכחי.");
    return;
  }

  // איסוף שורות מסומנות
  const checkedRows = [];
  const actionRange = activeSheet.getRange(2, actionColumnIndex + 1, data.length - 1, 1);
  const checkboxValues = actionRange.getValues();
  
  for (let i = 0; i < checkboxValues.length; i++) {
    if (checkboxValues[i][0] === true) {
      checkedRows.push({
        rowIndex: i + 2,
        email: data[i + 1][emailColumnIndex],
        senderName: data[i + 1][headers.indexOf("Sender Name")],
        emailId: data[i + 1][headers.indexOf("Email ID")],
        subject: data[i + 1][headers.indexOf("Subject")],
      });
    }
  }

  if (checkedRows.length === 0) {
    SpreadsheetApp.getUi().alert("אין שורות מסומנות לביצוע פעולות.");
    return;
  }

  // הצגת דיאלוג פעולות
  showActionsDialog(checkedRows, sheetName, settings, log);
}

// === ACTIONS DIALOG ===
function showActionsDialog(checkedRows, sheetName, settings, log) {
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
        <h4>החרג חשבוניות מסומנות</h4>
        <p>הוסף את המיילים המסומנים לרשימת ההחרגות ומחק את כל הרשומות עם אותם מיילים (כולל כפילויות).</p>
      </div>
      <div class="option" onclick="selectAction('approve')">
        <h4>הוסף למיילים מאושרים</h4>
        <p>הוסף את המיילים המסומנים לרשימת המאושרים (whitelist) – ייכנסו אוטומטית בעתיד.</p>
      </div>
      <div class="option" onclick="selectAction('local')">
        <h4>סמן כמקומיות</h4>
        <p>סמן את החשבוניות המסומנות כ-🔷 Local ושמור את ההגדרה לסריקות הבאות</p>
      </div>
      <div class="option" onclick="selectAction('international')">
        <h4>סמן כבינלאומיות</h4>
        <p>סמן את החשבוניות המסומנות כ-🌐 International ושמור את ההגדרה לסריקות הבאות</p>
      </div>
      <div class="option" onclick="selectAction('export')">
        <h4>ייצא חשבוניות מסומנות</h4>
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!sheet || !settings) {
    return { success: false, message: "❌ שגיאה בגישה לגיליונות." };
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const emailColumnIndex = headers.indexOf("Sender Email") + 1;
  const categoryColumnIndex = headers.indexOf("Category") + 1;

  try {
    switch (action) {
      case 'exclude':
        return handleExcludeAction(sheet, settings, log, checkedRows, emailColumnIndex);
      
      case 'approve':
        return handleApproveAction(settings, log, checkedRows);
      
      case 'local':
      case 'international':
        return handleCategoryAction(sheet, checkedRows, action, categoryColumnIndex);
      
      case 'export':
        return handleExportAction(sheetName, checkedRows);
      
      default:
        return { success: false, message: "❌ פעולה לא מוכרת." };
    }
  } catch (error) {
    _logMessage(log, `❌ שגיאה בביצוע פעולה ${action}: ${error.message}`, 'ERROR');
    return { success: false, message: `❌ שגיאה: ${error.message}` };
  }
}

// === EXCLUDE ACTION ===
function handleExcludeAction(sheet, settings, log, checkedRows, emailColumnIndex) {
  const allData = sheet.getDataRange().getValues();
  const emailsToExclude = checkedRows.map(r => r.email);
  const rowsToDelete = [];
  
  // איסוף כל השורות עם אותם מיילים
  for (let i = 1; i < allData.length; i++) {
    if (emailsToExclude.includes(allData[i][emailColumnIndex - 1])) {
      rowsToDelete.push(i + 1);
    }
  }
  
  // הוספה לרשימת החרגות
  const excludeRange = settings.getRange("F2:F");
  const excludeValues = excludeRange.getValues().flat().filter(Boolean);
  let newlyExcluded = 0;
  
  emailsToExclude.forEach(email => {
    if (!excludeValues.includes(email)) {
      settings.getRange(excludeValues.length + 2 + newlyExcluded, 6).setValue(email);
      newlyExcluded++;
    }
  });
  
  // מחיקת השורות (מלמטה למעלה)
  rowsToDelete.sort((a, b) => b - a).forEach(rowIdx => sheet.deleteRow(rowIdx));
  
  _logMessage(log, `הוחרגו ומחקו ${rowsToDelete.length} שורות (${emailsToExclude.length} מיילים)`, 'SUCCESS');
  
  return { 
    success: true, 
    message: `✅ ${rowsToDelete.length} שורות הוחרגו ונמחקו. (${emailsToExclude.length} מיילים הוספו להחרגות)` 
  };
}

// === APPROVE ACTION ===
function handleApproveAction(settings, log, checkedRows) {
  const approvedRange = settings.getRange("G2:G");
  const approvedValues = approvedRange.getValues().flat().filter(Boolean);
  let newlyApproved = 0;
  
  checkedRows.forEach(row => {
    if (!approvedValues.includes(row.email)) {
      settings.getRange(approvedValues.length + 2 + newlyApproved, 7).setValue(row.email);
      newlyApproved++;
    }
  });
  
  _logMessage(log, `הוספו ${newlyApproved} מיילים ל-whitelist`, 'SUCCESS');
  
  return { 
    success: true, 
    message: `✅ ${newlyApproved} מיילים הוספו לרשימת המאושרים.` 
  };
}

// === CATEGORY ACTION ===
function handleCategoryAction(sheet, checkedRows, action, categoryColumnIndex) {
  const category = action === 'local' ? "🔷 Local" : "🌐 International";
  
  for (const row of checkedRows) {
    setSenderCategory(row.email, category);
    const rowRange = sheet.getRange(row.rowIndex, categoryColumnIndex);
    rowRange.setValue(category);
  }
  
  return { 
    success: true, 
    message: `✅ ${checkedRows.length} חשבוניות סומנו כ-${category} ונשמרו להמשך.` 
  };
}

// === EXPORT ACTION ===
function handleExportAction(sheetName, checkedRows) {
  // זה יחובר לפונקציית הייצוא
  SpreadsheetApp.getUi().alert("פונקציית ייצוא תמומש בהמשך");
  return { success: true, message: "ייצוא יבוצע בהמשך" };
}

// === HELPER FUNCTIONS ===
function clearAllActions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pattern = /^invoices /i;
  const sheets = ss.getSheets().filter(sheet => pattern.test(sheet.getName()));
  
  sheets.forEach(sheet => {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const actionColIndex = headers.indexOf("Action") + 1;
    if (actionColIndex > 0) {
      const range = sheet.getRange(2, actionColIndex, sheet.getLastRow() - 1, 1);
      range.setValue(false); // Clear checkboxes
    }
  });
  
  SpreadsheetApp.getUi().alert("✅ כל הצ'קבוקסים נוקו.");
}

function selectAllActions() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (!/^invoices /i.test(sheetName)) {
    SpreadsheetApp.getUi().alert("❌ אנא עבור לגיליון 'invoices' המתאים.");
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const actionColIndex = headers.indexOf("Action") + 1;
  
  if (actionColIndex > 0 && sheet.getLastRow() > 1) {
    const range = sheet.getRange(2, actionColIndex, sheet.getLastRow() - 1, 1);
    range.setValue(true); // Check all checkboxes
    SpreadsheetApp.getUi().alert("✅ כל השורות נבחרו.");
  }
}