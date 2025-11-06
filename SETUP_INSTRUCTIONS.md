# 📥 Gmail Invoice Scanner - Local Download Setup

## הוראות התקנה והפעלה

### שלב 1: התקנת Python
ודא ש-Python 3.8 ומעלה מותקן במחשב:
```bash
python --version
```

### שלב 2: התקנת ספריות נדרשות
```bash
pip install -r requirements.txt
```

### שלב 3: הגדרת Google Service Account

#### 3.1 יצירת Service Account
1. עבור ל-[Google Cloud Console](https://console.cloud.google.com/)
2. בחר את הפרויקט שלך (או צור חדש)
3. עבור ל-**APIs & Services** > **Credentials**
4. לחץ על **Create Credentials** > **Service Account**
5. תן שם ל-Service Account (למשל: "invoice-downloader")
6. לחץ **Create and Continue**
7. תן הרשאה: **Editor** (או לפחות Viewer)
8. לחץ **Done**

#### 3.2 הורדת Credentials
1. לחץ על ה-Service Account שיצרת
2. עבור ל-**Keys** tab
3. לחץ **Add Key** > **Create new key**
4. בחר **JSON**
5. הקובץ יורד אוטומטית
6. **שנה את שם הקובץ ל-`credentials.json`**
7. **העתק אותו לאותה תיקייה של הסקריפט**

#### 3.3 שיתוף הגיליון עם Service Account
1. פתח את קובץ `credentials.json`
2. חפש את השדה `"client_email"` (נראה כמו: `xxx@xxx.iam.gserviceaccount.com`)
3. העתק את המייל הזה
4. פתח את הגיליון שלך ב-Google Sheets
5. לחץ **Share** (שיתוף)
6. הדבק את המייל של ה-Service Account
7. תן הרשאת **Viewer** (מספיק לקריאה)
8. לחץ **Send**

### שלב 4: הגדרת Spreadsheet ID

#### 4.1 מציאת ה-Spreadsheet ID
פתח את הגיליון שלך ב-Google Sheets.
ה-URL נראה כך:
```
https://docs.google.com/spreadsheets/d/1abc123XYZ456/edit
                                      ^^^^^^^^^^^^^^^^
                                      זה ה-Spreadsheet ID
```

#### 4.2 עדכון הסקריפט
פתח את `download_invoices.py` וערוך את השורה:
```python
SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE"  # Replace with your actual spreadsheet ID
```

החלף ל:
```python
SPREADSHEET_ID = "1abc123XYZ456"  # השם שלך
```

### שלב 5: הגדרת נתיב התיקייה בגיליון

1. פתח את הגיליון שלך
2. עבור לגיליון **Settings**
3. בתא **I2** (Export Folder ID), שים את הנתיב המלא:
   ```
   G:\Other computers\#Mystuff\#myStuff\uvix_biz\הוצ 2025\Bucket
   ```

### שלב 6: הרצת הסקריפט

```bash
python download_invoices.py
```

## מה הסקריפט עושה?

1. ✅ מתחבר לגיליון Google Sheets שלך
2. ✅ קורא את טווח התאריכים מ-Settings (B1, B2)
3. ✅ קורא את נתיב התיקייה מ-Settings (I2)
4. ✅ יוצר תיקיית משנה בשם: `YYYY-MM-DD_to_YYYY-MM-DD`
5. ✅ בודק אילו קבצים כבר קיימים בתיקייה
6. ✅ מוריד רק קבצים חדשים (מדלג על קיימים)
7. ✅ מציג סיכום: כמה הורדו, כמה דולגו, כמה נכשלו

## דוגמת פלט

```
============================================================
📥 Gmail Invoice Scanner - Local Download
============================================================

🔗 Connecting to Google Sheets...
✅ Connected successfully!

⚙️ Reading settings...
📅 Date Range: 2025-01-01 to 2025-01-31
📁 Base Path: G:\Other computers\#Mystuff\#myStuff\uvix_biz\הוצ 2025\Bucket

📁 Creating target folder...
✅ Target folder: G:\Other computers\#Mystuff\#myStuff\uvix_biz\הוצ 2025\Bucket\2025-01-01_to_2025-01-31

🔍 Checking existing files...
Found 5 existing files

📊 Finding invoice data...
📊 Using sheet: invoices 2025-01-15

📥 Reading invoice data...
Found 31 invoices to process

⬇️ Starting downloads...
------------------------------------------------------------
⏭️ [1/31] Skipping (exists): invoice_001.pdf
⬇️ [2/31] Downloading: invoice_002.pdf
   ✅ Saved to: G:\...\invoice_002.pdf
⬇️ [3/31] Downloading: invoice_003.pdf
   ✅ Saved to: G:\...\invoice_003.pdf
...

============================================================
📊 Download Summary
============================================================
✅ Downloaded: 26 files
⏭️ Skipped (already exist): 5 files
❌ Failed: 0 files
📁 Location: G:\Other computers\#Mystuff\#myStuff\uvix_biz\הוצ 2025\Bucket\2025-01-01_to_2025-01-31

🎉 Done!
```

## פתרון בעיות נפוצות

### שגיאה: "credentials.json file not found"
**פתרון:** ודא שקובץ `credentials.json` נמצא באותה תיקייה של הסקריפט.

### שגיאה: "Permission denied"
**פתרון:** ודא ששיתפת את הגיליון עם המייל של ה-Service Account.

### שגיאה: "No invoice sheets found"
**פתרון:** ודא שיש גיליון שמתחיל ב-`invoices ` (עם רווח).

### הקבצים לא מורדים
**פתרון:** ודא שעמודת "Download Link" מכילה לינקים תקינים של Google Drive.

### נתיב התיקייה לא עובד
**פתרון:** ודא שהנתיב בתא I2 קיים ושיש לך הרשאות כתיבה.

## שימוש מתקדם

### הרצה אוטומטית (Windows Task Scheduler)
אפשר להגדיר את הסקריפט לרוץ אוטומטית כל יום:
1. פתח **Task Scheduler**
2. צור **New Task**
3. הגדר Trigger (למשל: כל יום ב-9:00)
4. הגדר Action: `python C:\path\to\download_invoices.py`

### הרצה עם פרמטרים
ניתן להוסיף פרמטרים לסקריפט (דורש שינוי קוד).

## תמיכה
אם יש בעיות, בדוק:
1. ✅ Python מותקן ועובד
2. ✅ כל הספריות מותקנות (`pip install -r requirements.txt`)
3. ✅ קובץ `credentials.json` קיים
4. ✅ הגיליון משותף עם Service Account
5. ✅ Spreadsheet ID נכון בסקריפט
6. ✅ נתיב התיקייה בתא I2 קיים ותקין
