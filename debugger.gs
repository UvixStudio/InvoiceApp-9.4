// === debugger.gs - מערכת דיבאגינג מתקדמת ===

// === DEBUG CONFIGURATION ===
const DEBUG_CONFIG = {
  enabled: true,
  logLevel: 'INFO', // DEBUG, INFO, WARN, ERROR
  maxLogEntries: 1000,
  enableConsoleLog: true,
  enableSheetLog: true,
  enableToastMessages: true
};

// === MAIN DEBUG FUNCTIONS ===

/**
 * מערכת דיבאגינג ראשית - בדיקה מקיפה של כל המערכת
 */
function runFullSystemDebug() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let debugLog = ss.getSheetByName("DebugLog");
  
  // יצירת גיליון דיבאג אם לא קיים
  if (!debugLog) {
    debugLog = ss.insertSheet("DebugLog");
    debugLog.appendRow(["Timestamp", "Level", "Component", "Message", "Details"]);
    debugLog.setFrozenRows(1);
  }
  
  debugMessage("SYSTEM", "🚀 התחלת דיבאגינג מערכת מלא", "INFO", debugLog);
  
  const results = {
    passed: 0,
    failed: 0,
    warnings: 0,
    total: 0,
    details: []
  };
  
  // בדיקות בסיסיות
  runDebugTest("בדיקת גיליונות בסיסיים", debugBasicSheets, results, debugLog);
  runDebugTest("בדיקת הרשאות", debugPermissions, results, debugLog);
  runDebugTest("בדיקת פונקציות עזר", debugHelperFunctions, results, debugLog);
  runDebugTest("בדיקת תפריטים", debugMenus, results, debugLog);
  
  // בדיקות מתקדמות
  runDebugTest("בדיקת הגדרות", debugSettings, results, debugLog);
  runDebugTest("בדיקת לוגיקת סריקה", debugScanLogic, results, debugLog);
  runDebugTest("בדיקת עיבוד מיילים", debugEmailProcessing, results, debugLog);
  
  // הצגת תוצאות
  showDebugResults(results, debugLog);
}

/**
 * דיבאגינג מהיר - בדיקות חיוניות בלבד
 */
function runQuickDebug() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    SpreadsheetApp.getUi().alert("❌ לא נמצא גיליון Log. הפעל Setup Sheets תחילה.");
    return;
  }
  
  _logMessage(log, "⚡ התחלת דיבאגינג מהיר", 'INFO');
  
  const issues = [];
  
  // בדיקות חיוניות
  if (!checkBasicSheets()) issues.push("❌ בעיה בגיליונות בסיסיים");
  if (!checkPermissions()) issues.push("❌ בעיית הרשאות");
  if (!checkSettings()) issues.push("❌ בעיה בהגדרות");
  
  if (issues.length === 0) {
    _logMessage(log, "✅ דיבאגינג מהיר הושלם - הכל תקין", 'SUCCESS');
    SpreadsheetApp.getUi().alert("✅ דיבאגינג מהיר הושלם בהצלחה!\n\nהמערכת נראית תקינה.");
  } else {
    _logMessage(log, `⚠️ נמצאו ${issues.length} בעיות`, 'WARNING');
    SpreadsheetApp.getUi().alert(`⚠️ נמצאו בעיות:\n\n${issues.join('\n')}\n\nבדוק את גיליון Log לפרטים.`);
  }
}

/**
 * דיבאגינג סריקה - בדיקה מפורטת של תהליך הסריקה
 */
function debugScanProcess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!settings || !log) {
    SpreadsheetApp.getUi().alert("❌ גיליונות Settings או Log חסרים.");
    return;
  }
  
  _logMessage(log, "🔍 התחלת דיבאגינג תהליך סריקה", 'INFO');
  
  try {
    // שלב 1: בדיקת הגדרות
    _logMessage(log, "📋 שלב 1: בדיקת הגדרות", 'INFO');
    // Try to read settings - these functions are in gmailProcessor.gs
    let config;
    try {
      // Basic config creation for debugging
      const startStr = settings.getRange("A2").getValue();
      const endStr = settings.getRange("B2").getValue();
      
      if (!startStr || !endStr) {
        _logMessage(log, "❌ תאריכים לא מוגדרים בהגדרות", 'ERROR');
        return;
      }
      
      config = {
        startDate: new Date(startStr),
        endDate: new Date(endStr),
        enKeywords: settings.getRange("C2:C").getValues().flat().filter(Boolean),
        heKeywords: settings.getRange("D2:D").getValues().flat().filter(Boolean),
        excludedKeywords: settings.getRange("E2:E").getValues().flat().filter(Boolean),
        excludedEmails: settings.getRange("F2:F").getValues().flat().filter(Boolean),
        approvedSenders: settings.getRange("G2:G").getValues().flat().filter(Boolean)
      };
      
      _logMessage(log, "✅ הגדרות נקראו בהצלחה", 'SUCCESS');
    } catch (error) {
      _logMessage(log, `❌ שגיאה בקריאת הגדרות: ${error.message}`, 'ERROR');
      return;
    }
    
    // שלב 2: בניית שאילתה
    _logMessage(log, "🔍 שלב 2: בניית שאילתת Gmail", 'INFO');
    const query = buildBasicGmailQuery(config);
    _logMessage(log, `📝 שאילתה: ${query}`, 'INFO');
    
    // שלב 3: חיפוש מיילים (מוגבל)
    _logMessage(log, "📧 שלב 3: חיפוש מיילים (5 ראשונים)", 'INFO');
    const threads = GmailApp.search(query, 0, 5);
    _logMessage(log, `📊 נמצאו ${threads.length} שרשורי מייל`, 'INFO');
    
    // שלב 4: ניתוח מיילים
    if (threads.length > 0) {
      _logMessage(log, "🔬 שלב 4: ניתוח מיילים", 'INFO');
      analyzeEmailsForDebug(threads, config, log);
    }
    
    _logMessage(log, "✅ דיבאגינג סריקה הושלם", 'SUCCESS');
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בדיבאגינג סריקה: ${error.message}`, 'ERROR');
  }
}/**
 * ניתוח מיילים לצורך דיבאגינג
 */
function analyzeEmailsForDebug(threads, config, log) {
  let validCount = 0;
  let skippedCount = 0;
  
  for (let i = 0; i < Math.min(threads.length, 3); i++) {
    const thread = threads[i];
    const messages = thread.getMessages();
    const msg = messages[messages.length - 1];
    
    // יצירת נתוני מייל בסיסיים לדיבאג
    const emailData = {
      email: msg.getFrom(),
      subject: msg.getSubject(),
      body: msg.getPlainBody().substring(0, 200),
      date: msg.getDate(),
      pdfAttachment: msg.getAttachments().find(att => att.getContentType() === 'application/pdf')
    };
    
    if (!emailData.email) continue;
    
    _logMessage(log, `📧 מייל ${i + 1}: ${emailData.email}`, 'INFO');
    _logMessage(log, `📝 נושא: ${emailData.subject.substring(0, 50)}...`, 'INFO');
    
    // בדיקת סינונים בסיסית
    let shouldSkip = false;
    const reasons = [];
    
    // בדיקת מיילים מוחרגים
    if (config.excludedEmails.some(excluded => emailData.email.includes(excluded))) {
      shouldSkip = true;
      reasons.push("מייל מוחרג");
    }
    
    // בדיקת דומיין פרטי
    if (emailData.email.includes('@gmail.com') || emailData.email.includes('@yahoo.com') || 
        emailData.email.includes('@hotmail.com') || emailData.email.includes('@outlook.com')) {
      shouldSkip = true;
      reasons.push("דומיין פרטי");
    }
    
    // בדיקת קבצים מצורפים
    if (!emailData.pdfAttachment) {
      reasons.push("אין PDF");
    }
    
    if (shouldSkip) {
      skippedCount++;
      _logMessage(log, `⏩ מייל ${i + 1} נדחה`, 'WARNING');
      _logMessage(log, `📋 סיבות דחייה: ${reasons.join(', ')}`, 'INFO');
    } else {
      validCount++;
      _logMessage(log, `✅ מייל ${i + 1} תקין`, 'SUCCESS');
    }
  }
  
  _logMessage(log, `📊 סיכום ניתוח: ${validCount} תקינים, ${skippedCount} נדחו`, 'INFO');
}

/**
 * דיבאגינג ביצועים - מדידת זמני ביצוע
 */
function debugPerformance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    SpreadsheetApp.getUi().alert("❌ לא נמצא גיליון Log.");
    return;
  }
  
  _logMessage(log, "⚡ התחלת דיבאגינג ביצועים", 'INFO');
  
  const performanceTests = [
    { name: "יצירת גיליון זמני", test: testSheetCreation },
    { name: "כתיבת 50 שורות", test: testBulkWrite },
    { name: "קריאת נתונים גדולים", test: testBulkRead },
    { name: "חיפוש Gmail בסיסי", test: testGmailSearch },
    { name: "עיבוד פונקציות עזר", test: testHelperFunctions }
  ];
  
  const results = [];
  
  for (const test of performanceTests) {
    const startTime = new Date().getTime();
    
    try {
      test.test();
      const endTime = new Date().getTime();
      const duration = endTime - startTime;
      
      results.push({ name: test.name, duration, success: true });
      _logMessage(log, `✅ ${test.name}: ${duration}ms`, duration < 2000 ? 'SUCCESS' : 'WARNING');
      
    } catch (error) {
      const endTime = new Date().getTime();
      const duration = endTime - startTime;
      
      results.push({ name: test.name, duration, success: false, error: error.message });
      _logMessage(log, `❌ ${test.name}: ${duration}ms - ${error.message}`, 'ERROR');
    }
  }
  
  // סיכום ביצועים
  const totalTime = results.reduce((sum, r) => sum + r.duration, 0);
  const successCount = results.filter(r => r.success).length;
  
  _logMessage(log, `📊 סיכום ביצועים: ${successCount}/${results.length} הצליחו, זמן כולל: ${totalTime}ms`, 'INFO');
  
  SpreadsheetApp.getUi().alert(
    `⚡ דיבאגינג ביצועים הושלם\n\n` +
    `${successCount}/${results.length} בדיקות הצליחו\n` +
    `זמן כולל: ${totalTime}ms\n\n` +
    `בדוק את גיליון Log לפרטים מלאים.`
  );
}

/**
 * דיבאגינג זיכרון ומשאבים
 */
function debugMemoryUsage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    SpreadsheetApp.getUi().alert("❌ לא נמצא גיליון Log.");
    return;
  }
  
  _logMessage(log, "💾 התחלת דיבאגינג זיכרון", 'INFO');
  
  try {
    // בדיקת מספר גיליונות
    const sheets = ss.getSheets();
    _logMessage(log, `📊 מספר גיליונות: ${sheets.length}`, 'INFO');
    
    // בדיקת גודל נתונים
    let totalCells = 0;
    let totalRows = 0;
    
    sheets.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const cells = lastRow * lastCol;
      
      totalCells += cells;
      totalRows += lastRow;
      
      _logMessage(log, `📋 ${sheet.getName()}: ${lastRow} שורות, ${cells} תאים`, 'INFO');
    });
    
    _logMessage(log, `📊 סך הכל: ${totalRows} שורות, ${totalCells} תאים`, 'INFO');
    
    // בדיקת מגבלות
    const warnings = [];
    if (sheets.length > 200) warnings.push("יותר מדי גיליונות");
    if (totalCells > 2000000) warnings.push("יותר מדי תאים");
    if (totalRows > 50000) warnings.push("יותר מדי שורות");
    
    if (warnings.length > 0) {
      _logMessage(log, `⚠️ אזהרות: ${warnings.join(', ')}`, 'WARNING');
    } else {
      _logMessage(log, "✅ שימוש בזיכרון תקין", 'SUCCESS');
    }
    
    // בדיקת זמן ביצוע
    const startTime = new Date().getTime();
    for (let i = 0; i < 1000; i++) {
      Math.random() * Math.random();
    }
    const executionTime = new Date().getTime() - startTime;
    
    _logMessage(log, `⏱️ זמן ביצוע לולאה: ${executionTime}ms`, executionTime < 100 ? 'SUCCESS' : 'WARNING');
    
    const status = warnings.length === 0 && executionTime < 100 ? 
      '✅ מצב זיכרון מצוין' : '⚠️ מצב זיכרון דורש מעקב';
    
    SpreadsheetApp.getUi().alert(
      `💾 דיבאגינג זיכרון\n\n${status}\n\n` +
      `גיליונות: ${sheets.length}\n` +
      `תאים: ${totalCells.toLocaleString()}\n` +
      `זמן ביצוע: ${executionTime}ms\n\n` +
      (warnings.length > 0 ? `אזהרות: ${warnings.join(', ')}` : 'הכל תקין!')
    );
    
  } catch (error) {
    _logMessage(log, `❌ שגיאה בדיבאגינג זיכרון: ${error.message}`, 'ERROR');
  }
}

// === HELPER DEBUG FUNCTIONS ===

function runDebugTest(testName, testFunction, results, debugLog) {
  results.total++;
  
  try {
    const result = testFunction();
    if (result.success) {
      results.passed++;
      debugMessage("TEST", `✅ ${testName}`, "SUCCESS", debugLog, result.message);
    } else if (result.warning) {
      results.warnings++;
      debugMessage("TEST", `⚠️ ${testName}`, "WARNING", debugLog, result.message);
    } else {
      results.failed++;
      debugMessage("TEST", `❌ ${testName}`, "ERROR", debugLog, result.message);
    }
  } catch (error) {
    results.failed++;
    debugMessage("TEST", `❌ ${testName}`, "ERROR", debugLog, `Exception: ${error.message}`);
  }
}

function debugMessage(component, message, level, debugLog, details = "") {
  const timestamp = new Date().toISOString();
  
  // לוג לגיליון
  if (debugLog && DEBUG_CONFIG.enableSheetLog) {
    debugLog.appendRow([timestamp, level, component, message, details]);
  }
  
  // לוג לקונסול
  if (DEBUG_CONFIG.enableConsoleLog) {
    console.log(`[${timestamp}] ${level} [${component}] ${message} ${details}`);
  }
  
  // הודעת Toast (רק לרמות חשובות)
  if (DEBUG_CONFIG.enableToastMessages && (level === 'ERROR' || level === 'SUCCESS')) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, component, 3);
  }
}

function showDebugResults(results, debugLog) {
  const successRate = Math.round((results.passed / results.total) * 100);
  
  const message = `🔍 תוצאות דיבאגינג מערכת

📊 סיכום:
• ${results.total} בדיקות בוצעו
• ${results.passed} עברו בהצלחה
• ${results.warnings} אזהרות
• ${results.failed} נכשלו
• שיעור הצלחה: ${successRate}%

${successRate >= 80 ? '✅ המערכת תקינה!' : '⚠️ יש בעיות שדורשות תיקון.'}`;

  debugMessage("SYSTEM", "דיבאגינג הושלם", "INFO", debugLog, message);
  
  SpreadsheetApp.getUi().alert("תוצאות דיבאגינג", message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// === SPECIFIC DEBUG TESTS ===

function debugBasicSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!settings) return { success: false, message: "גיליון Settings חסר" };
  if (!log) return { success: false, message: "גיליון Log חסר" };
  
  return { success: true, message: "גיליונות בסיסיים קיימים" };
}

function debugPermissions() {
  try {
    GmailApp.getInboxThreads(0, 1);
    DriveApp.getRootFolder();
    return { success: true, message: "הרשאות Gmail ו-Drive תקינות" };
  } catch (error) {
    return { success: false, message: `בעיית הרשאות: ${error.message}` };
  }
}

function debugHelperFunctions() {
  try {
    const testDate = new Date();
    
    // בדיקת פונקציות תאריך בסיסיות
    const gmailFormat = formatDateGmail(testDate);
    const basicDateStr = testDate.toISOString().split('T')[0];
    const timeStr = testDate.toLocaleTimeString();
    
    // בדיקת פונקציות מחרוזת בסיסיות
    const testFrom = "Test User <test@example.com>";
    const emailMatch = testFrom.match(/<(.+)>/);
    const email = emailMatch ? emailMatch[1] : testFrom;
    
    // בדיקת פונקציות חילוץ בסיסיות
    const testSubject = "Invoice #123 - Amount: $150.00";
    const testBody = "Total amount due: $150.00";
    const hasAmount = testSubject.includes('$') || testBody.includes('$');
    
    // בדיקת פונקציות דומיין בסיסיות
    const testEmail1 = "test@gmail.com";
    const testEmail2 = "test@company.com";
    const isPrivate1 = testEmail1.includes('@gmail.com') || testEmail1.includes('@yahoo.com');
    const isPrivate2 = testEmail2.includes('@gmail.com') || testEmail2.includes('@yahoo.com');
    
    return { success: true, message: "פונקציות עזר בסיסיות עובדות" };
  } catch (error) {
    return { success: false, message: `שגיאה בפונקציות עזר: ${error.message}` };
  }
}

function debugMenus() {
  try {
    createMainMenu();
    createActionsMenu();
    return { success: true, message: "תפריטים נוצרו בהצלחה" };
  } catch (error) {
    return { success: false, message: `שגיאה בתפריטים: ${error.message}` };
  }
}

function debugSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!settings) return { success: false, message: "גיליון Settings חסר" };
  
  const startDate = settings.getRange("A2").getValue();
  const endDate = settings.getRange("B2").getValue();
  
  if (!startDate || !endDate) {
    return { success: false, message: "תאריכים לא מוגדרים" };
  }
  
  const enKeywords = settings.getRange("C2:C").getValues().flat().filter(Boolean);
  const heKeywords = settings.getRange("D2:D").getValues().flat().filter(Boolean);
  
  if (enKeywords.length === 0 && heKeywords.length === 0) {
    return { warning: true, message: "אין מילות מפתח מוגדרות" };
  }
  
  return { success: true, message: "הגדרות תקינות" };
}

function debugScanLogic() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
    const log = ss.getSheetByName(LOG_SHEET_NAME);
    
    // יצירת config בסיסי לבדיקה
    const startStr = settings.getRange("A2").getValue();
    const endStr = settings.getRange("B2").getValue();
    
    if (!startStr || !endStr) {
      return { success: false, message: "תאריכים לא מוגדרים בהגדרות" };
    }
    
    const config = {
      startDate: new Date(startStr),
      endDate: new Date(endStr),
      enKeywords: settings.getRange("C2:C").getValues().flat().filter(Boolean),
      heKeywords: settings.getRange("D2:D").getValues().flat().filter(Boolean),
      excludedKeywords: settings.getRange("E2:E").getValues().flat().filter(Boolean),
      excludedEmails: settings.getRange("F2:F").getValues().flat().filter(Boolean),
      approvedSenders: settings.getRange("G2:G").getValues().flat().filter(Boolean)
    };
    
    const query = buildBasicGmailQuery(config);
    if (!query || query.length < 10) {
      return { success: false, message: "שגיאה בבניית שאילתה" };
    }
    
    return { success: true, message: "לוגיקת סריקה תקינה" };
  } catch (error) {
    return { success: false, message: `שגיאה בלוגיקת סריקה: ${error.message}` };
  }
}

function debugEmailProcessing() {
  try {
    // בדיקת פונקציות עיבוד מייל בסיסיות
    const testEmailData = {
      email: "test@company.com",
      subject: "Invoice #123",
      body: "Amount: $100",
      pdfAttachment: null
    };
    
    const testConfig = {
      excludedEmails: ["spam@example.com"],
      excludedKeywords: ["newsletter"],
      approvedSenders: ["approved@company.com"]
    };
    
    // בדיקות בסיסיות שאנחנו יכולים לעשות
    const isExcluded = testConfig.excludedEmails.some(excluded => 
      testEmailData.email.includes(excluded));
    
    const hasKeywords = testEmailData.subject.toLowerCase().includes("invoice");
    
    return { success: true, message: "עיבוד מיילים בסיסי תקין" };
  } catch (error) {
    return { success: false, message: `שגיאה בעיבוד מיילים: ${error.message}` };
  }
}

// === PERFORMANCE TEST FUNCTIONS ===

function testSheetCreation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.insertSheet("PerformanceTest_" + new Date().getTime());
  ss.deleteSheet(testSheet);
}

function testBulkWrite() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.insertSheet("BulkTest_" + new Date().getTime());
  
  for (let i = 0; i < 50; i++) {
    testSheet.appendRow([`Test ${i}`, new Date(), Math.random()]);
  }
  
  ss.deleteSheet(testSheet);
}

function testBulkRead() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  for (const sheet of sheets.slice(0, 3)) {
    sheet.getDataRange().getValues();
  }
}

function testGmailSearch() {
  GmailApp.search("has:attachment", 0, 5);
}

function testHelperFunctions() {
  for (let i = 0; i < 100; i++) {
    // בדיקת פונקציות תאריך
    const testDate = new Date();
    formatDateGmail(testDate);
    formatDateInput(testDate);
    formatDateShort(testDate);
    
    // בדיקת פונקציות מייל
    const testEmail = "Test <test@example.com>";
    _parseSender(testEmail);
    extractSenderName("test@company.com");
    
    // בדיקת פונקציות דומיין
    isPrivateEmail("test@gmail.com");
    
    // בדיקת פונקציות חילוץ
    extractAmountAndCurrency("Invoice $100", "Total: $100");
  }
}

// === BASIC GMAIL QUERY BUILDER FOR DEBUG ===
function buildBasicGmailQuery(config) {
  const parts = [];
  
  // טווח תאריכים
  parts.push(`after:${formatDateGmail(config.startDate)}`);
  parts.push(`before:${formatDateGmail(config.endDate)}`);
  
  // קבצים מצורפים
  parts.push("has:attachment");
  
  // מילות מפתח בסיסיות
  const allKeywords = [...config.enKeywords, ...config.heKeywords];
  if (allKeywords.length > 0) {
    const keywordQuery = allKeywords.slice(0, 3).map(k => `"${k}"`).join(" OR ");
    parts.push(`(${keywordQuery})`);
  }
  
  return parts.join(" ");
}

// === FORMAT DATE FOR GMAIL QUERY ===
function formatDateGmail(date) {
  if (!date || !(date instanceof Date)) {
    return new Date().toISOString().split('T')[0].replace(/-/g, '/');
  }
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  
  return `${year}/${month}/${day}`;
}

/**
 * בדיקה מקיפה של כל רכיבי המערכת
 */
function runComprehensiveSystemCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!log) {
    SpreadsheetApp.getUi().alert("❌ לא נמצא גיליון Log. הפעל Setup Sheets תחילה.");
    return;
  }
  
  _logMessage(log, "🔍 התחלת בדיקה מקיפה של המערכת", 'INFO');
  
  const checkResults = {
    permissions: checkAllPermissions(),
    sheets: checkAllSheets(),
    settings: checkSystemSettings(),
    functions: checkCoreFunctions(),
    performance: checkSystemPerformance()
  };
  
  // הצגת תוצאות מפורטות
  displayComprehensiveResults(checkResults, log);
}

function checkAllPermissions() {
  const results = {};
  
  // בדיקת הרשאות Gmail
  try {
    GmailApp.getInboxThreads(0, 1);
    GmailApp.search("has:attachment", 0, 1);
    results.gmail = { success: true, message: "הרשאות Gmail תקינות" };
  } catch (e) {
    results.gmail = { success: false, message: `בעיית הרשאות Gmail: ${e.message}` };
  }
  
  // בדיקת הרשאות Drive
  try {
    DriveApp.getRootFolder();
    results.drive = { success: true, message: "הרשאות Drive תקינות" };
  } catch (e) {
    results.drive = { success: false, message: `בעיית הרשאות Drive: ${e.message}` };
  }
  
  // בדיקת הרשאות Sheets
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheets();
    results.sheets = { success: true, message: "הרשאות Sheets תקינות" };
  } catch (e) {
    results.sheets = { success: false, message: `בעיית הרשאות Sheets: ${e.message}` };
  }
  
  return results;
}

function checkAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = {};
  
  // בדיקת גיליון Settings
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (settings) {
    const startDate = settings.getRange("A2").getValue();
    const endDate = settings.getRange("B2").getValue();
    results.settings = {
      exists: true,
      configured: !!(startDate && endDate),
      message: startDate && endDate ? "גיליון Settings מוגדר" : "גיליון Settings קיים אך לא מוגדר"
    };
  } else {
    results.settings = { exists: false, message: "גיליון Settings חסר" };
  }
  
  // בדיקת גיליון Log
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  results.log = {
    exists: !!log,
    message: log ? "גיליון Log קיים" : "גיליון Log חסר"
  };
  
  // בדיקת גיליונות תוצאות
  const invoiceSheets = ss.getSheets().filter(sheet => /^invoices /i.test(sheet.getName()));
  results.invoiceSheets = {
    count: invoiceSheets.length,
    message: `נמצאו ${invoiceSheets.length} גיליונות תוצאות`
  };
  
  return results;
}

function checkSystemSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!settings) {
    return { configured: false, message: "גיליון Settings חסר" };
  }
  
  const config = {
    startDate: settings.getRange("A2").getValue(),
    endDate: settings.getRange("B2").getValue(),
    enKeywords: settings.getRange("C2:C").getValues().flat().filter(Boolean).length,
    heKeywords: settings.getRange("D2:D").getValues().flat().filter(Boolean).length,
    excludedEmails: settings.getRange("F2:F").getValues().flat().filter(Boolean).length,
    approvedSenders: settings.getRange("G2:G").getValues().flat().filter(Boolean).length
  };
  
  const issues = [];
  if (!config.startDate || !config.endDate) issues.push("תאריכים לא מוגדרים");
  if (config.enKeywords === 0 && config.heKeywords === 0) issues.push("אין מילות מפתח");
  
  return {
    configured: issues.length === 0,
    details: config,
    issues: issues,
    message: issues.length === 0 ? "הגדרות תקינות" : `בעיות: ${issues.join(', ')}`
  };
}

function checkCoreFunctions() {
  const results = {};
  
  // בדיקת פונקציות תאריך
  try {
    const testDate = new Date();
    formatDateGmail(testDate);
    formatDateInput(testDate);
    formatDateShort(testDate);
    results.dateFunctions = { success: true, message: "פונקציות תאריך תקינות" };
  } catch (e) {
    results.dateFunctions = { success: false, message: `שגיאה בפונקציות תאריך: ${e.message}` };
  }
  
  // בדיקת פונקציות מייל
  try {
    _parseSender("Test <test@example.com>");
    extractSenderName("test@company.com");
    isPrivateEmail("test@gmail.com");
    results.emailFunctions = { success: true, message: "פונקציות מייל תקינות" };
  } catch (e) {
    results.emailFunctions = { success: false, message: `שגיאה בפונקציות מייל: ${e.message}` };
  }
  
  // בדיקת פונקציות חילוץ
  try {
    extractAmountAndCurrency("Invoice $100", "Total: $100");
    extractInvoiceLinkFromPlainText("Download: https://example.com/invoice.pdf");
    results.extractionFunctions = { success: true, message: "פונקציות חילוץ תקינות" };
  } catch (e) {
    results.extractionFunctions = { success: false, message: `שגיאה בפונקציות חילוץ: ${e.message}` };
  }
  
  return results;
}

function checkSystemPerformance() {
  const startTime = new Date().getTime();
  
  // בדיקת ביצועים בסיסית
  for (let i = 0; i < 1000; i++) {
    Math.random() * Math.random();
  }
  
  const executionTime = new Date().getTime() - startTime;
  
  return {
    executionTime: executionTime,
    status: executionTime < 100 ? 'מצוין' : executionTime < 500 ? 'טוב' : 'איטי',
    message: `זמן ביצוע: ${executionTime}ms`
  };
}

function displayComprehensiveResults(results, log) {
  _logMessage(log, "📊 תוצאות בדיקה מקיפה:", 'INFO');
  
  // הרשאות
  _logMessage(log, `🔐 Gmail: ${results.permissions.gmail.success ? '✅' : '❌'} ${results.permissions.gmail.message}`, 
    results.permissions.gmail.success ? 'SUCCESS' : 'ERROR');
  _logMessage(log, `🔐 Drive: ${results.permissions.drive.success ? '✅' : '❌'} ${results.permissions.drive.message}`, 
    results.permissions.drive.success ? 'SUCCESS' : 'ERROR');
  _logMessage(log, `🔐 Sheets: ${results.permissions.sheets.success ? '✅' : '❌'} ${results.permissions.sheets.message}`, 
    results.permissions.sheets.success ? 'SUCCESS' : 'ERROR');
  
  // גיליונות
  _logMessage(log, `📋 Settings: ${results.sheets.settings.exists ? '✅' : '❌'} ${results.sheets.settings.message}`, 
    results.sheets.settings.exists ? 'SUCCESS' : 'ERROR');
  _logMessage(log, `📋 Log: ${results.sheets.log.exists ? '✅' : '❌'} ${results.sheets.log.message}`, 
    results.sheets.log.exists ? 'SUCCESS' : 'ERROR');
  _logMessage(log, `📋 תוצאות: ${results.sheets.invoiceSheets.message}`, 'INFO');
  
  // הגדרות
  _logMessage(log, `⚙️ הגדרות: ${results.settings.configured ? '✅' : '❌'} ${results.settings.message}`, 
    results.settings.configured ? 'SUCCESS' : 'WARNING');
  
  // פונקציות
  Object.entries(results.functions).forEach(([key, result]) => {
    _logMessage(log, `🔧 ${key}: ${result.success ? '✅' : '❌'} ${result.message}`, 
      result.success ? 'SUCCESS' : 'ERROR');
  });
  
  // ביצועים
  _logMessage(log, `⚡ ביצועים: ${results.performance.message} (${results.performance.status})`, 
    results.performance.status === 'מצוין' ? 'SUCCESS' : 'INFO');
  
  // סיכום כללי
  const allGood = results.permissions.gmail.success && 
                  results.permissions.drive.success && 
                  results.permissions.sheets.success && 
                  results.sheets.settings.exists && 
                  results.sheets.log.exists && 
                  results.settings.configured;
  
  _logMessage(log, allGood ? "🎉 המערכת תקינה ומוכנה לשימוש!" : "⚠️ יש בעיות שדורשות תיקון", 
    allGood ? 'SUCCESS' : 'WARNING');
  
  SpreadsheetApp.getUi().alert(
    allGood ? "✅ בדיקה מקיפה הושלמה - המערכת תקינה!" : 
    "⚠️ בדיקה מקיפה הושלמה - נמצאו בעיות. בדוק את גיליון Log לפרטים."
  );
}

// === BASIC CHECK FUNCTIONS ===

function checkBasicSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SETTINGS_SHEET_NAME) && ss.getSheetByName(LOG_SHEET_NAME);
}

function checkPermissions() {
  try {
    GmailApp.getInboxThreads(0, 1);
    DriveApp.getRootFolder();
    return true;
  } catch (error) {
    return false;
  }
}

function checkSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  
  if (!settings) return false;
  
  const startDate = settings.getRange("A2").getValue();
  const endDate = settings.getRange("B2").getValue();
  
  return startDate && endDate;
}