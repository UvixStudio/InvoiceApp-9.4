// === testAppend.gs - פונקציות בדיקה ובדיקות איכות ===

/**
 * בדיקה כללית של המערכת
 */
function runSystemTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    SpreadsheetApp.getUi().alert("❌ לא נמצא גיליון Log. אנא הפעל Setup Sheets תחילה.");
    return;
  }
  
  _logMessage(log, "🧪 התחלת בדיקת מערכת כוללת", 'INFO');
  
  const testResults = {
    passed: 0,
    failed: 0,
    total: 0,
    details: []
  };
  
  // בדיקת גיליונות בסיסיים
  runTest("בדיקת גיליון Settings", testSettingsSheet, testResults, log);
  runTest("בדיקת גיליון Log", testLogSheet, testResults, log);
  
  // בדיקת הרשאות
  runTest("בדיקת הרשאות Gmail", testGmailPermissions, testResults, log);
  runTest("בדיקת הרשאות Drive", testDrivePermissions, testResults, log);
  
  // בדיקת פונקציות עזר
  runTest("בדיקת פונקציות תאריך", testDateFunctions, testResults, log);
  runTest("בדיקת פונקציות מייל", testEmailFunctions, testResults, log);
  runTest("בדיקת פונקציות חילוץ", testExtractionFunctions, testResults, log);
  
  // בדיקת תפריטים
  runTest("בדיקת יצירת תפריטים", testMenuCreation, testResults, log);
  
  // הצגת תוצאות
  showTestResults(testResults, log);
}

/**
 * הרצת בדיקה בודדת
 */
function runTest(testName, testFunction, results, log) {
  results.total++;
  
  try {
    const result = testFunction();
    if (result.success) {
      results.passed++;
      _logMessage(log, `✅ ${testName}: ${result.message}`, 'SUCCESS');
      results.details.push(`✅ ${testName}: PASSED`);
    } else {
      results.failed++;
      _logMessage(log, `❌ ${testName}: ${result.message}`, 'ERROR');
      results.details.push(`❌ ${testName}: FAILED - ${result.message}`);
    }
  } catch (error) {
    results.failed++;
    _logMessage(log, `❌ ${testName}: Exception - ${error.message}`, 'ERROR');
    results.details.push(`❌ ${testName}: EXCEPTION - ${error.message}`);
  }
}

/**
 * בדיקת גיליון Settings
 */
function testSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!settings) {
    return { success: false, message: "גיליון Settings לא קיים" };
  }
  
  // בדיקת כותרות
  const headers = settings.getRange("A1:I1").getValues()[0];
  const expectedHeaders = [
    "Start Date", "End Date", "EN Keywords", "HE Keywords", 
    "Excluded Keywords", "Exclude Emails", "Approved Senders", "", "Export Folder ID"
  ];
  
  for (let i = 0; i < expectedHeaders.length; i++) {
    if (headers[i] !== expectedHeaders[i]) {
      return { success: false, message: `כותרת שגויה בעמודה ${i + 1}: ${headers[i]}` };
    }
  }
  
  return { success: true, message: "גיליון Settings תקין" };
}

/**
 * בדיקת גיליון Log
 */
function testLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    return { success: false, message: "גיליון Log לא קיים" };
  }
  
  // בדיקת כותרות
  const headers = log.getRange("A1:B1").getValues()[0];
  if (headers[0] !== "Timestamp" || headers[1] !== "Message") {
    return { success: false, message: "כותרות Log שגויות" };
  }
  
  // בדיקת יכולת כתיבה
  try {
    const testMessage = "Test message " + new Date().getTime();
    _logMessage(log, testMessage, 'INFO');
    return { success: true, message: "גיליון Log תקין וניתן לכתיבה" };
  } catch (error) {
    return { success: false, message: `שגיאה בכתיבה ל-Log: ${error.message}` };
  }
}

/**
 * בדיקת הרשאות Gmail
 */
function testGmailPermissions() {
  try {
    const threads = GmailApp.getInboxThreads(0, 1);
    return { success: true, message: "הרשאות Gmail תקינות" };
  } catch (error) {
    return { success: false, message: `בעיית הרשאות Gmail: ${error.message}` };
  }
}

/**
 * בדיקת הרשאות Drive
 */
function testDrivePermissions() {
  try {
    const rootFolder = DriveApp.getRootFolder();
    const testFolderName = "test_folder_" + new Date().getTime();
    const testFolder = rootFolder.createFolder(testFolderName);
    testFolder.setTrashed(true);
    return { success: true, message: "הרשאות Drive תקינות" };
  } catch (error) {
    return { success: false, message: `בעיית הרשאות Drive: ${error.message}` };
  }
}

/**
 * בדיקת פונקציות תאריך
 */
function testDateFunctions() {
  try {
    const testDate = new Date(2024, 0, 15); // January 15, 2024
    
    const gmailFormat = formatDateGmail(testDate);
    if (gmailFormat !== "2024/01/15") {
      return { success: false, message: `formatDateGmail שגוי: ${gmailFormat}` };
    }
    
    const inputFormat = formatDateInput(testDate);
    if (inputFormat !== "2024-01-15") {
      return { success: false, message: `formatDateInput שגוי: ${inputFormat}` };
    }
    
    const shortFormat = formatDateShort(testDate);
    if (shortFormat !== "15.01.24") {
      return { success: false, message: `formatDateShort שגוי: ${shortFormat}` };
    }
    
    return { success: true, message: "פונקציות תאריך תקינות" };
  } catch (error) {
    return { success: false, message: `שגיאה בפונקציות תאריך: ${error.message}` };
  }
}

/**
 * בדיקת פונקציות מייל
 */
function testEmailFunctions() {
  try {
    // בדיקת חילוץ מייל
    const testFrom = "Test User <test@example.com>";
    const email = _parseSender(testFrom);
    if (email !== "test@example.com") {
      return { success: false, message: `_parseSender שגוי: ${email}` };
    }
    
    // בדיקת שם שולח
    const senderName = extractSenderName("test@company.com");
    if (!senderName || senderName === "unknown") {
      return { success: false, message: `extractSenderName שגוי: ${senderName}` };
    }
    
    // בדיקת דומיין פרטי
    const isPrivate = isPrivateEmail("test@gmail.com");
    if (!isPrivate) {
      return { success: false, message: "isPrivateEmail לא זיהה gmail.com כפרטי" };
    }
    
    const isNotPrivate = isPrivateEmail("test@company.com");
    if (isNotPrivate) {
      return { success: false, message: "isPrivateEmail זיהה company.com כפרטי" };
    }
    
    return { success: true, message: "פונקציות מייל תקינות" };
  } catch (error) {
    return { success: false, message: `שגיאה בפונקציות מייל: ${error.message}` };
  }
}

/**
 * בדיקת פונקציות חילוץ
 */
function testExtractionFunctions() {
  try {
    // בדיקת חילוץ סכום
    const testSubject = "Invoice #123 - Amount: $150.00";
    const testBody = "Total amount due: $150.00";
    
    const amountData = extractAmountAndCurrency(testSubject, testBody);
    if (!amountData || !amountData.amount) {
      return { success: false, message: "לא הצליח לחלץ סכום" };
    }
    
    // בדיקת חילוץ קישור
    const testHtml = '<a href="https://example.com/invoice.pdf">Download Invoice</a>';
    const link = extractInvoiceLinkFromHtml(testHtml);
    if (!link || !link.includes("invoice.pdf")) {
      return { success: false, message: "לא הצליח לחלץ קישור מ-HTML" };
    }
    
    // בדיקת חילוץ מטקסט רגיל
    const testText = "Please download your invoice from: https://company.com/download/invoice123.pdf";
    const textLink = extractInvoiceLinkFromPlainText(testText);
    if (!textLink || !textLink.includes("invoice123.pdf")) {
      return { success: false, message: "לא הצליח לחלץ קישור מטקסט רגיל" };
    }
    
    return { success: true, message: "פונקציות חילוץ תקינות" };
  } catch (error) {
    return { success: false, message: `שגיאה בפונקציות חילוץ: ${error.message}` };
  }
}

/**
 * בדיקת יצירת תפריטים
 */
function testMenuCreation() {
  try {
    createMainMenu();
    createActionsMenu();
    return { success: true, message: "תפריטים נוצרו בהצלחה" };
  } catch (error) {
    return { success: false, message: `שגיאה ביצירת תפריטים: ${error.message}` };
  }
}

/**
 * הצגת תוצאות הבדיקות
 */
function showTestResults(results, log) {
  const successRate = Math.round((results.passed / results.total) * 100);
  
  const message = `🧪 תוצאות בדיקת המערכת

📊 סיכום:
• ${results.total} בדיקות בוצעו
• ${results.passed} עברו בהצלחה
• ${results.failed} נכשלו
• שיעור הצלחה: ${successRate}%

${successRate >= 80 ? '✅ המערכת תקינה ומוכנה לשימוש!' : '⚠️ יש בעיות שדורשות תיקון.'}`;

  _logMessage(log, message.replace(/\n/g, ' '), 'INFO');
  
  SpreadsheetApp.getUi().alert("תוצאות בדיקת מערכת", message, SpreadsheetApp.getUi().ButtonSet.OK);
  
  // הצגת פרטים בלוג
  results.details.forEach(detail => {
    _logMessage(log, detail, 'INFO');
  });
}

/**
 * בדיקה מהירה של פונקציית הסריקה
 */
function testScanFunction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!settings || !log) {
    SpreadsheetApp.getUi().alert("❌ גיליונות Settings או Log חסרים.");
    return;
  }
  
  _logMessage(log, "🧪 בדיקה מהירה של פונקציית הסריקה", 'INFO');
  
  try {
    // בדיקת קריאת הגדרות
    const config = readScanSettings(settings, log);
    if (!config) {
      _logMessage(log, "❌ שגיאה בקריאת הגדרות", 'ERROR');
      return;
    }
    
    _logMessage(log, "✅ קריאת הגדרות תקינה", 'SUCCESS');
    
    // בדיקת בניית שאילתה
    const query = buildGmailQuery(config);
    if (!query || query.length < 10) {
      _logMessage(log, "❌ שגיאה בבניית שאילתת Gmail", 'ERROR');
      return;
    }
    
    _logMessage(log, `✅ שאילתת Gmail נבנתה: ${query}`, 'SUCCESS');
    
    // בדיקת חיפוש Gmail (ללא עיבוד)
    const threads = GmailApp.search(query, 0, 5);
    _logMessage(log, `✅ חיפוש Gmail הצליח: נמצאו ${threads.length} שרשורים`, 'SUCCESS');
    
    SpreadsheetApp.getUi().alert(`✅ בדיקה מהירה הצליחה!\n\nנמצאו ${threads.length} שרשורי מייל עם השאילתה:\n${query}`);
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בבדיקה מהירה: ${error.message}`, 'ERROR');
    SpreadsheetApp.getUi().alert(`❌ שגיאה בבדיקה מהירה: ${error.message}`);
  }
}

/**
 * בדיקת ביצועים
 */
function performanceTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    SpreadsheetApp.getUi().alert("❌ לא נמצא גיליון Log.");
    return;
  }
  
  _logMessage(log, "⚡ התחלת בדיקת ביצועים", 'INFO');
  
  const startTime = new Date().getTime();
  
  // בדיקת מהירות כתיבה לגיליון
  const testSheet = ss.insertSheet("PerformanceTest");
  const writeStartTime = new Date().getTime();
  
  for (let i = 0; i < 100; i++) {
    testSheet.appendRow([`Test ${i}`, new Date(), Math.random()]);
  }
  
  const writeEndTime = new Date().getTime();
  const writeTime = writeEndTime - writeStartTime;
  
  // בדיקת מהירות קריאה
  const readStartTime = new Date().getTime();
  const data = testSheet.getDataRange().getValues();
  const readEndTime = new Date().getTime();
  const readTime = readEndTime - readStartTime;
  
  // ניקוי
  ss.deleteSheet(testSheet);
  
  const totalTime = new Date().getTime() - startTime;
  
  const results = `⚡ תוצאות בדיקת ביצועים:

• כתיבת 100 שורות: ${writeTime}ms
• קריאת ${data.length} שורות: ${readTime}ms
• זמן כולל: ${totalTime}ms

${totalTime < 5000 ? '✅ ביצועים טובים' : '⚠️ ביצועים איטיים'}`;

  _logMessage(log, results.replace(/\n/g, ' '), 'INFO');
  SpreadsheetApp.getUi().alert("תוצאות בדיקת ביצועים", results, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * בדיקת זיכרון ומגבלות
 */
function memoryTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    SpreadsheetApp.getUi().alert("❌ לא נמצא גיליון Log.");
    return;
  }
  
  _logMessage(log, "💾 בדיקת זיכרון ומגבלות", 'INFO');
  
  try {
    // בדיקת מספר גיליונות
    const sheets = ss.getSheets();
    _logMessage(log, `📊 מספר גיליונות: ${sheets.length}`, 'INFO');
    
    // בדיקת גודל נתונים
    let totalCells = 0;
    sheets.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      totalCells += lastRow * lastCol;
    });
    
    _logMessage(log, `📊 סך תאים: ${totalCells}`, 'INFO');
    
    // בדיקת זמן ביצוע
    const startTime = new Date().getTime();
    
    // פעולה כבדה לבדיקה
    for (let i = 0; i < 1000; i++) {
      Math.random() * Math.random();
    }
    
    const executionTime = new Date().getTime() - startTime;
    _logMessage(log, `⏱️ זמן ביצוע לולאה: ${executionTime}ms`, 'INFO');
    
    const status = totalCells < 100000 && executionTime < 100 ? 
      '✅ מצב זיכרון טוב' : '⚠️ מצב זיכרון דורש מעקב';
    
    _logMessage(log, status, totalCells < 100000 ? 'SUCCESS' : 'WARNING');
    
    SpreadsheetApp.getUi().alert(`💾 בדיקת זיכרון\n\n${status}\n\nגיליונות: ${sheets.length}\nתאים: ${totalCells}\nזמן ביצוע: ${executionTime}ms`);
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בבדיקת זיכרון: ${error.message}`, 'ERROR');
  }
}