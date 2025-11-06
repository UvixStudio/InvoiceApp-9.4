# V32 – הנחיות ממוקדות לקירו (Mini-Log, Toasts, Batch, Download Links)

מסמך זה מרכז את ההנחיות המדויקות עבור Kiro לתיקון V32, כולל דוגמאות קוד.

---

## 1️⃣ Mini-Log – מיקום, מבנה ולוגיקה

### מבנה הגיליון בפועל

- שורות **1–11** = אזור סטטוס (mini-log)
- שורה **12** = שורת רווח
- שורה **13** = כותרות הטבלה (header הכחול)
- שורה **14+** = רשומות החשבוניות (data)

בעיקר משתמשים בעמודות **F** ו־**G**:

```text
F1: "status"          (כותרת)
G1: "שאילתה"         (כותרת)

G2: טקסט השאילתה המלא
F3: "זמן התחלה: HH:MM:SS"
F4: "משך: 00:00:17"  (אופציונלי)
F5: "נמצאו X מיילים בסינון הראשוני"
F6: "סורק i מתוך N"
F7: "0% |████      | 40%  (progress text)"
F8: "מתקבל: שולח – נושא" (מייל נוכחי)
F9: "תוצאות: X נוספו, Y דולגו, Z כפילויות"
F10–F11: רזרבה עתידית
```

### איפוס והכנה בתחילת הסריקה (`gmailProcessor.gs / processInvoices`)

```js
// בתוך processInvoices, אחרי שנקבע sheet והשאילתה query

const statusSheet = sheet; // אותו גיליון של התוצאות

// איפוס תכנים בלבד – לא מוחקים עיצוב
const statusRange = statusSheet.getRange("F1:G11");
statusRange.clearContent();

// כותרות
statusSheet.getRange("F1").setValue("status");
statusSheet.getRange("G1").setValue("שאילתה");

// שאילתה (אחרי שנבנתה)
statusSheet.getRange("G2").setValue(query);

// זמן התחלה
statusSheet.getRange("F3").setValue(
  `זמן התחלה: ${startTime.toLocaleTimeString()}`
);

// מצב ראשוני
statusSheet.getRange("F6").setValue("מתחיל סריקה...");
statusSheet.getRange("F7").setValue("0 % |                    |");
```

### עדכון תוך כדי ריצה (`gmailProcessor.gs / processEmails`)

```js
function processEmails(emails, config, sheet, log, startTime) {
  let inserted = 0;
  let skipped = 0;
  let duplicates = 0;

  const existingRecords = getExistingRecords(sheet);
  const allRecords = [];

  const startingRow = findLastRowWithData(sheet) + 1;
  _logMessage(log, `📝 רשומות חדשות יתחילו משורה ${startingRow}`, 'INFO');

  const statusSheet = sheet;

  for (let i = 0; i < emails.length; i++) {
    const emailData = emails[i];

    try {
      if (existingRecords.has(emailData.messageId)) {
        duplicates++;
        continue;
      }

      const record = createInvoiceRecord(emailData, config);
      allRecords.push(record);
      inserted++;

    } catch (error) {
      _logMessage(log, `❌ שגיאה בעיבוד מייל ${emailData.email}: ${error.message}`, 'ERROR');
      skipped++;
    }

    // עדכון Mini-Log כל כמה מיילים
    if (i === 0 || i === emails.length - 1 || i % 3 === 0) {
      const current = i + 1;
      const total = emails.length;
      const bar = createScanProgressBar(current, total);

      statusSheet.getRange("F6").setValue(`סורק ${current} מתוך ${total}`);
      statusSheet.getRange("F7").setValue(bar);
      statusSheet.getRange("F8").setValue(
        `מתקבל: ${emailData.email} – ${emailData.subject}`
      );
    }
  }

  if (allRecords.length > 0) {
    const range = sheet.getRange(startingRow, 1, allRecords.length, allRecords[0].length);
    range.setValues(allRecords);

    // צ'קבוקסים לכל השורות החדשות
    sheet.getRange(startingRow, 1, allRecords.length, 1).insertCheckboxes();
  }

  SpreadsheetApp.flush();

  statusSheet.getRange("F9").setValue(
    `תוצאות: ${inserted} נוספו, ${skipped} דולגו, ${duplicates} כפילויות`
  );

  _logMessage(log,
    `📊 תוצאות עיבוד: ${inserted} נוספו, ${skipped} דולגו, ${duplicates} כפילויות`,
    'SUCCESS'
  );

  return { inserted, skipped, duplicates, total: emails.length };
}
```

⚠️ חשוב: Mini-Log מנוהל מ־`processInvoices` ו־`processEmails`, **לא** מתוך `createOrUpdateResultsSheet`.

---

## 2️⃣ Toasts – לנקות מהזרימה

### המצב הרצוי

- במהלך הסריקה: **אין** טוסטים.  
- אם רוצים – טוסט יחיד קצר בסוף, אחרי הסיכום.

### דוגמאות

```js
// ❌ להסיר מתוך processEmails:
// SpreadsheetApp.getActiveSpreadsheet().toast(...);

// ❌ להסיר מתוך תחילת processInvoices:
// SpreadsheetApp.getActiveSpreadsheet().toast("🔍 מתחיל סריקת Gmail...", "סריקה", 5);

// ✅ אם משאירים – טוסט אחד בסוף (אופציונלי בלבד)
SpreadsheetApp.getActiveSpreadsheet().toast(
  `✅ הסריקה הושלמה: ${results.inserted} נוספו, ${results.skipped} דולגו`,
  "הושלם",
  5
);
```

כל התקדמות בזמן אמת עוברת דרך ה־Mini-Log ב־F1–G11, לא דרך toast חוזרים.

---

## 3️⃣ Batch Write ב־processEmails

העיקרון: **אין setValues / flush בתוך הלולאה**.

```js
function processEmails(emails, config, sheet, log, startTime) {
  let inserted = 0;
  let skipped = 0;
  let duplicates = 0;

  const existingRecords = getExistingRecords(sheet);
  const allRecords = [];

  // הנתונים מתחילים מתחת לכותרות
  const startingRow = findLastRowWithData(sheet) + 1;
  _logMessage(log, `📝 רשומות חדשות יתחילו משורה ${startingRow}`, 'INFO');

  const statusSheet = sheet;

  for (let i = 0; i < emails.length; i++) {
    const emailData = emails[i];

    try {
      if (existingRecords.has(emailData.messageId)) {
        duplicates++;
        continue;
      }

      const record = createInvoiceRecord(emailData, config);
      allRecords.push(record);
      inserted++;

    } catch (error) {
      _logMessage(log, `❌ שגיאה בעיבוד מייל ${emailData.email}: ${error.message}`, 'ERROR');
      skipped++;
    }

    // כאן רק עדכון mini-log (כמו בסעיף 1), בלי setValues ובלי flush
  }

  if (allRecords.length > 0) {
    const range = sheet.getRange(startingRow, 1, allRecords.length, allRecords[0].length);
    range.setValues(allRecords);
    sheet.getRange(startingRow, 1, allRecords.length, 1).insertCheckboxes();
  }

  SpreadsheetApp.flush(); // פעם אחת בסוף בלבד

  statusSheet.getRange("F9").setValue(
    `תוצאות: ${inserted} נוספו, ${skipped} דולגו, ${duplicates} כפילויות`
  );

  _logMessage(log,
    `📊 תוצאות עיבוד: ${inserted} נוספו, ${skipped} דולגו, ${duplicates} כפילויות`,
    'SUCCESS'
  );

  return { inserted, skipped, duplicates, total: emails.length };
}
```

---

## 4️⃣ Gmail-First – createInvoiceRecord + כותרות

### הרחבת כותרות (`gmailProcessor.gs / createOrUpdateResultsSheet`)

```js
const headers = [
  "Action", "Category", "Sender Name", "Sender Email", "Date",
  "Subject", "PDF Link", "Email ID", "Attachment Name",
  "Sum (Local)", "Sum (Int'l)", "Currency",
  "Mail Link", "Download Link" // עמודות חדשות
];

// כותרות בשורה 13
sheet.getRange(13, 1, 1, headers.length).setValues([headers]);
sheet.setFrozenRows(13); // אופציונלי – הקפאת 13 שורות ראשונות

// התאמת רוחב עמודות Gmail
sheet.setColumnWidth(13, 250); // Mail Link
sheet.setColumnWidth(14, 250); // Download Link
```

### יצירת רשומה ללא Drive (`gmailProcessor.gs / createInvoiceRecord`)

```js
function buildGmailMessageLink(messageId) {
  return `https://mail.google.com/mail/u/0/#all/${messageId}`;
}

function createInvoiceRecord(emailData, config) {
  const isLocal = isLocalEmail(emailData.email, emailData.subject, emailData.body);
  const category = isLocal ? "🔷 Local" : "🌐 International";
  const mailLink = buildGmailMessageLink(emailData.messageId);

  const attachmentName = emailData.pdfAttachment
    ? emailData.pdfAttachment.getName()
    : "";

  return [
    false,                                   // Action
    category,                                // Category
    emailData.senderName || "",              // Sender Name
    emailData.email || "",                   // Sender Email
    formatDateInput(emailData.date),         // Date
    emailData.subject || "",                 // Subject
    "",                                      // PDF Link (ריק בסריקה)
    emailData.messageId || "",               // Email ID
    attachmentName,                          // Attachment Name
    "",                                      // Sum (Local)
    "",                                      // Sum (Int'l)
    "",                                      // Currency
    mailLink,                                // Mail Link
    ""                                       // Download Link (ימולא בהמשך)
  ];
}
```

אין `DriveApp.createFile` בשלב הסריקה.

---

## 5️⃣ prepareDownloadLinksForSelected – לוגיקת selection ו-Batch Write

### בניית קישור להורדת הקובץ

```js
function buildGmailAttachmentLink(messageId, attachmentIndex) {
  const attId = `0.${attachmentIndex + 1}`; // 0.1 לקובץ הראשון, 0.2 לשני וכו'
  return `https://mail.google.com/mail/u/0/?ui=2&view=att&th=${messageId}&attid=${attId}&disp=safe&zw`;
}
```

### לוגיקת בחירת שורות + Batch Write

```js
function prepareDownloadLinksForSelected() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];

  const emailIdColIdx = headers.indexOf("Email ID");
  const downloadColIdx =	headers.indexOf("Download Link");

  if (emailIdColIdx === -1 || downloadColIdx === -1) {
    SpreadsheetApp.getUi().alert("❌ חסרות עמודות 'Email ID' או 'Download Link'.");
    return;
  }

  const selection = sheet.getSelection();
  const activeRange = selection ? selection.getActiveRange() : null;

  const rowDescriptors = [];

  if (activeRange && activeRange.getNumRows() > 1) {
    // יש selection → משתמשים רק בו
    const startRow = activeRange.getRow();        // 1-based
    const numRows = activeRange.getNumRows();

    for (let r = startRow; r < startRow + numRows; r++) {
      if (r <= 1 || r > data.length) continue;   // דילוג על כותרות ושורות מעבר לטווח
      const rowData = data[r - 1];               // data[0] = כותרות
      const messageId = rowData[emailIdColIdx];
      if (messageId) {
        rowDescriptors.push({ sheetRow: r, dataIndex: r - 1, messageId });
      }
    }
  } else {
    // אין selection → כל השורות עם דאטה
    for (let i = 1; i < data.length; i++) {      // i=1 -> שורה 2
      const messageId = data[i][emailIdColIdx];
      if (messageId) {
        rowDescriptors.push({ sheetRow: i + 1, dataIndex: i, messageId });
      }
    }
  }

  if (rowDescriptors.length === 0) {
    SpreadsheetApp.getUi().alert("אין שורות עם Email ID לעיבוד.");
    return;
  }

  const gmailCache = {};
  let updated = 0;

  for (const row of rowDescriptors) {
    try {
      const message = gmailCache[row.messageId] ||
        GmailApp.getMessageById(row.messageId);
      gmailCache[row.messageId] = message;

      const attachments = message.getAttachments();
      const pdfAttachmentIndex = attachments.findIndex(att =>
        att.getName().toLowerCase().endsWith('.pdf')
      );
      if (pdfAttachmentIndex === -1) continue;

      const link = buildGmailAttachmentLink(row.messageId, pdfAttachmentIndex);
      data[row.dataIndex][downloadColIdx] = link;
      updated++;
    } catch (e) {
      // אפשר לרשום ללוג אם רוצים
    }
  }

  // כתיבה חזרה בבת אחת (כולל כותרות)
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  SpreadsheetApp.getUi().alert(`✅ הוכנו קישורי הורדה ל-${updated} שורות.`);
}
```

התנהגות:

- אם יש selection של כמה שורות – רק עליהן עוברים.  
- אם אין selection – עוברים על כל השורות שיש בהן Email ID.  
- אין תלות בצ'קבוקסים לפיצ'ר הזה.  
- כל העדכונים נעשים על המערך `data`, ורק בסוף נכתב לגיליון פעם אחת.

---

## 6️⃣ חיבור לפעולות בתפריט (`code.gs`)

```js
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ פעולות')
    // פעולות קיימות...
    .addItem('📦 הכן קבצים להורדה (נבחרים / הכל)', 'prepareDownloadLinksForSelected')
    .addToUi();
}
```

---

## 7️⃣ מה לבדוק אחרי התיקון

1. להריץ סריקה על חודש אחד:
   - לראות שה־Mini-Log ב־F1–G11 מתעדכן בזמן ריצה.
   - לראות שהשורות בטבלה "קופצות" בבת אחת בסוף, ולא אחת־אחת.
   - לוודא שלא נוצרים קבצי Drive חדשים בזמן הסריקה.

2. לבחור כמה שורות ולהריץ **"הכן קבצים להורדה"**:
   - עם selection – רק השורות המסומנות מקבלות Download Link.
   - בלי selection – כל השורות עם Email ID מקבלות Download Link.
   - לחיצה על Download Link פותחת את ה־PDF ב־Gmail Viewer.

אם כל זה עובד – V32 עומדת בדיוק בדרישות:  
סריקה מהירה, Mini-Log חי מעל הטבלה, וזרימת Gmail-First נקייה.
