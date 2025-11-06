# 📋 V32 – מסמך תיקון מעודכן לקירו (על בסיס הקוד הנוכחי)

## 0. מטרה

1. להפוך את הסריקה למהירה ויציבה (בלי העלאות Drive מיותרות, בלי כתיבה שורה-שורה).  
2. לתת ליוזר חיווי לייב **רק** דרך שדות סטטוס בגיליון (mini-log), לא דרך טוסטים שרצים בלולאה.  
3. לעבוד בגישת **Gmail-First**:
   - סריקה: רק מטא־דאטה + Gmail Message Link  
   - הורדה: רק כשמשתמש מבקש, רק לשורות שנבחרו.

---

## 1️⃣ Mini-Log – מיקום נכון ומיפוי תאים

### מבנה בפועל
```
F1:G11  →  סטטוס, זמנים, תוצאות, פרוגרס
שורה 12 →  רווח
שורה 13 →  כותרות (Header)
שורות 14+ → נתוני חשבוניות
```

### קוד להטמעה
```js
const statusSheet = sheet;
const statusRange = statusSheet.getRange("F1:G11");
statusRange.clearContent().clearFormat();

statusSheet.getRange("F1").setValue("status");
statusSheet.getRange("G1").setValue("שאילתה");
statusSheet.getRange("G2").setValue(query);
statusSheet.getRange("F3").setValue(`זמן התחלה: ${startTime.toLocaleTimeString()}`);
statusSheet.getRange("F6").setValue("מתחיל סריקה...");
statusSheet.getRange("F7").setValue("0 % |                    |");
```

### עדכון תוך כדי ריצה
```js
if (i === 0 || i === emails.length - 1 || i % 3 === 0) {
  const current = i + 1;
  const total = emails.length;
  const percentage = Math.round(current * 100 / total);
  const bar = createScanProgressBar(current, total);

  statusSheet.getRange("F6").setValue(`סורק ${current} מתוך ${total}`);
  statusSheet.getRange("F7").setValue(bar);
  statusSheet.getRange("F8").setValue(`מתקבל: ${emailData.email} – ${emailData.subject}`);
}
```

### בסיום
```js
statusSheet.getRange("F9").setValue(
  `תוצאות: ${inserted} נוספו, ${skipped} דולגו, ${duplicates} כפילויות`
);
```

---

## 2️⃣ Toasts – להעלים כמעט לגמרי

### רק טוסט אחד בסוף
```js
SpreadsheetApp.getActiveSpreadsheet().toast(
  `✅ הסריקה הושלמה: ${results.inserted} נוספו, ${results.skipped} דולגו`,
  "הושלם",
  5
);
```

---

## 3️⃣ Batch Write – כתיבה מרוכזת

```js
function processEmails(emails, config, sheet, log, startTime) {
  let inserted = 0;
  let skipped = 0;
  let duplicates = 0;

  const existingRecords = getExistingRecords(sheet);
  const allRecords = [];

  const startingRow = findLastRowWithData(sheet) + 1;

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
  }

  if (allRecords.length > 0) {
    const range = sheet.getRange(startingRow, 1, allRecords.length, allRecords[0].length);
    range.setValues(allRecords);
    sheet.getRange(startingRow, 1, allRecords.length, 1).insertCheckboxes();
  }

  SpreadsheetApp.flush();
  return { inserted, skipped, duplicates, total: emails.length };
}
```

---

## 4️⃣ Gmail-First – ללא Drive בסריקה

### createOrUpdateResultsSheet
```js
const headers = [
  "Action", "Category", "Sender Name", "Sender Email", "Date", 
  "Subject", "PDF Link", "Email ID", "Attachment Name", 
  "Sum (Local)", "Sum (Int'l)", "Currency",
  "Mail Link", "Download Link"
];
```

### createInvoiceRecord
```js
function createInvoiceRecord(emailData, config) {
  const isLocal = isLocalEmail(emailData.email, emailData.subject, emailData.body);
  const category = isLocal ? "🔷 Local" : "🌐 International";
  const mailLink = buildGmailMessageLink(emailData.messageId);

  return [
    false, category, emailData.senderName || "", emailData.email || "",
    formatDateInput(emailData.date), emailData.subject || "", "",
    emailData.messageId || "", emailData.pdfAttachment?.getName() || "",
    "", "", "", mailLink, ""
  ];
}
```

---

## 5️⃣ הכנת קישורי הורדה

```js
function buildGmailAttachmentLink(messageId, attachmentIndex) {
  const attId = `0.${attachmentIndex + 1}`;
  return `https://mail.google.com/mail/u/0/?ui=2&view=att&th=${messageId}&attid=${attId}&disp=safe&zw`;
}

function prepareDownloadLinksForSelected() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailIdColIdx = headers.indexOf("Email ID");
  const downloadColIdx = headers.indexOf("Download Link");

  const selection = sheet.getSelection();
  const activeRange = selection ? selection.getActiveRange() : null;
  const rows = [];

  if (activeRange && activeRange.getNumRows() > 1) {
    const startRow = activeRange.getRow();
    const numRows = activeRange.getNumRows();
    for (let r = startRow; r < startRow + numRows; r++) {
      if (r > 1 && r <= data.length && data[r-1][emailIdColIdx]) {
        rows.push({ sheetRow: r, dataIndex: r - 1, messageId: data[r - 1][emailIdColIdx] });
      }
    }
  } else {
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIdColIdx]) {
        rows.push({ sheetRow: i + 1, dataIndex: i, messageId: data[i][emailIdColIdx] });
      }
    }
  }

  const gmailCache = {};
  let updated = 0;

  for (const row of rows) {
    try {
      const message = gmailCache[row.messageId] || GmailApp.getMessageById(row.messageId);
      gmailCache[row.messageId] = message;
      const attachments = message.getAttachments();
      const pdfIndex = attachments.findIndex(a => a.getName().toLowerCase().endsWith('.pdf'));
      if (pdfIndex === -1) continue;
      const link = buildGmailAttachmentLink(row.messageId, pdfIndex);
      data[row.dataIndex][downloadColIdx] = link;
      updated++;
    } catch (e) {}
  }

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  SpreadsheetApp.getUi().alert(`✅ הוכנו קישורי הורדה ל-${updated} שורות.`);
}
```

---

## 6️⃣ חיבור לתפריט

```js
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ פעולות')
    .addItem('📦 הכן קבצים להורדה (נבחרים / הכל)', 'prepareDownloadLinksForSelected')
    .addToUi();
}
```

---

## 7️⃣ סיכום

✅ סריקה מהירה (Batch Write בלבד)  
✅ חיווי בזמן אמת מעל הטבלה (F1–G11)  
✅ ללא Drive בזמן סריקה  
✅ הורדה רק לפי בחירת המשתמש  
✅ קישורים ל־Gmail ול־Download לכל רשומה
