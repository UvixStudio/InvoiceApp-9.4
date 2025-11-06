// === SmartFileNaming.gs - שמות קבצים חכמים ===

/**
 * ⚡ V32D: בניית שם קובץ חכם ומובנה
 * פורמט: YYYY-MM-DD_Company_InvoiceNumber_Description.pdf
 */
function buildSmartFileName(messageId, dateStr, senderEmail, senderName, subject, originalFileName) {
  try {
    // 1. תאריך בפורמט YYYY-MM-DD
    const date = new Date(dateStr);
    const formattedDate = formatDateForFileName(date);
    
    // 2. שם חברה חכם (מהנושא קודם!)
    const companyName = extractCompanyName(subject, senderName, senderEmail);
    
    // 3. מספר חשבונית (אם יש)
    const invoiceNumber = extractInvoiceNumber(subject, originalFileName);
    
    // 4. תיאור קצר
    const description = extractDescription(subject, originalFileName);
    
    // 5. בניית השם
    let fileName = `${formattedDate}_${companyName}`;
    
    if (invoiceNumber) {
      fileName += `_${invoiceNumber}`;
    }
    
    if (description) {
      fileName += `_${description}`;
    }
    
    fileName += '.pdf';
    
    // 6. ניקוי תווים לא חוקיים
    fileName = sanitizeFileName(fileName);
    
    // 7. הגבלת אורך (מקסימום 100 תווים)
    if (fileName.length > 100) {
      const ext = '.pdf';
      fileName = fileName.substring(0, 100 - ext.length) + ext;
    }
    
    return fileName;
    
  } catch (error) {
    console.error(`Error building smart filename: ${error.message}`);
    // fallback לשם מקורי
    return originalFileName || `invoice_${messageId}.pdf`;
  }
}

/**
 * פורמט תאריך לשם קובץ: YYYY-MM-DD
 */
function formatDateForFileName(date) {
  if (!date || !(date instanceof Date)) {
    date = new Date();
  }
  
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  
  return `${yyyy}-${mm}-${dd}`;
}

/**
 * חילוץ שם חברה חכם - קודם מהנושא, אחר כך מהשולח
 */
function extractCompanyName(subject, senderName, senderEmail) {
  // 1. חיפוש בנושא - זה הכי חשוב!
  const subjectCompany = findCompanyInSubject(subject);
  if (subjectCompany) {
    return subjectCompany;
  }
  
  // 2. מ-Sender Name (אם לא generic)
  if (senderName && !isGenericSenderName(senderName)) {
    const cleaned = cleanCompanyName(senderName);
    if (cleaned && cleaned.length >= 3) {
      return cleaned;
    }
  }
  
  // 3. מהדומיין (fallback)
  return extractCompanyFromDomain(senderEmail);
}

/**
 * חיפוש שם חברה בנושא המייל
 */
function findCompanyInSubject(subject) {
  if (!subject) return null;
  
  const subjectLower = subject.toLowerCase();
  
  // דפוסים לחיפוש שם חברה
  const patterns = [
    // עברית
    { regex: /מאת\s+([א-ת\w\s]{3,20})/i, group: 1 },           // "מאת יוביקס סטודיו"
    { regex: /-\s*([א-ת]{3,15})\s*-/i, group: 1 },             // "- כינרת -"
    { regex: /ב([א-ת]{3,15})(?:\s|$)/i, group: 1 },            // "בבזק", "בכינרת"
    { regex: /מבית\s+([א-ת\w]{3,15})/i, group: 1 },           // "מבית קרפור"
    { regex: /של\s+([א-ת]{3,15})/i, group: 1 },               // "של בזק"
    
    // אנגלית
    { regex: /\b([A-Z][a-z]{2,}(?:\s+[A-Z][a-z]+)?)\b/, group: 1 },  // "DigitalOcean", "Google Workspace"
    { regex: /from\s+([A-Z][a-zA-Z\s]{2,20})/i, group: 1 },    // "from Company Name"
  ];
  
  for (const pattern of patterns) {
    const match = subject.match(pattern.regex);
    if (match && match[pattern.group]) {
      const company = cleanCompanyName(match[pattern.group]);
      if (company && company.length >= 3) {
        return company;
      }
    }
  }
  
  // חיפוש שמות חברות ידועים (case-insensitive)
  const knownCompanies = {
    'כינרת': 'Kinneret',
    'kinneret': 'Kinneret',
    'בזק': 'Bezeq',
    'bezeq': 'Bezeq',
    'פרטנר': 'Partner',
    'partner': 'Partner',
    'סלקום': 'Cellcom',
    'cellcom': 'Cellcom',
    'גילת': 'Gilat',
    'gilat': 'Gilat',
    'קרפור': 'Carrefour',
    'carrefour': 'Carrefour',
    'יוביקס': 'Uvix',
    'uvix': 'Uvix',
    'digitalocean': 'DigitalOcean',
    'google': 'Google',
    'gotoglobal': 'GoTo',
    'goto': 'GoTo'
  };
  
  for (const [key, value] of Object.entries(knownCompanies)) {
    if (subjectLower.includes(key.toLowerCase())) {
      return value;
    }
  }
  
  return null;
}

/**
 * בדיקה אם שם השולח הוא generic (לא מועיל)
 */
function isGenericSenderName(name) {
  if (!name) return true;
  
  const genericNames = [
    'noreply', 'no-reply', 'do not reply', 'donot', 'donotreply',
    'notify', 'notification', 'support', 'info', 'admin',
    'accounting', 'billing', 'invoices', 'payments',
    'thankyou', 'thank you', 'web customer', 'customer'
  ];
  
  const nameLower = name.toLowerCase().replace(/[^a-z]/g, '');
  
  return genericNames.some(generic => nameLower.includes(generic));
}

/**
 * ניקוי שם חברה
 */
function cleanCompanyName(name) {
  if (!name) return '';
  
  // הסרת תווים מיוחדים ורווחים מיותרים
  let cleaned = name
    .trim()
    .replace(/[<>:"\/\\|?*]/g, '')  // תווים לא חוקיים
    .replace(/\s+/g, ' ')            // רווחים כפולים
    .substring(0, 20);               // הגבלת אורך
  
  // אות ראשונה גדולה
  if (cleaned.length > 0) {
    cleaned = cleaned.charAt(0).toUpperCase() + cleaned.slice(1);
  }
  
  return cleaned;
}

/**
 * חילוץ שם חברה מהדומיין
 */
function extractCompanyFromDomain(email) {
  if (!email || !email.includes('@')) {
    return 'Unknown';
  }
  
  const domain = email.split('@')[1];
  
  // מיפוי דומיינים ידועים
  const domainMapping = {
    'morning.co': 'Morning',
    'gilat.net': 'Gilat',
    'digitalocean.com': 'DigitalOcean',
    'google.com': 'Google',
    'gotoglobal.com': 'GoTo',
    'finbot.co.il': 'Finbot',
    'comax.co.il': 'Comax',
    'bezeq.co.il': 'Bezeq',
    'partner.net.il': 'Partner',
    'we-com.co.il': 'Wecom'
  };
  
  if (domainMapping[domain]) {
    return domainMapping[domain];
  }
  
  // חילוץ מהדומיין
  const domainParts = domain.split('.');
  const mainDomain = domainParts[0];
  
  // ניקוי
  const cleaned = mainDomain
    .replace(/noreply|no-reply|notify|support|info/gi, '')
    .replace(/[^a-zA-Z]/g, '')
    .toLowerCase();
  
  if (cleaned.length >= 3) {
    return cleaned.charAt(0).toUpperCase() + cleaned.slice(1);
  }
  
  return 'Unknown';
}

/**
 * חילוץ מספר חשבונית
 */
function extractInvoiceNumber(subject, originalFileName) {
  const text = `${subject} ${originalFileName}`;
  
  // דפוסים למספרי חשבונית
  const patterns = [
    /(?:חשבונית|קבלה|receipt|invoice)[\s#:]*(\d{4,})/i,  // "חשבונית 80549", "Invoice 12345"
    /\b(\d{5,})\b/,                                         // מספר של 5+ ספרות
    /_(\d{4,})(?:_|\.|$)/,                                  // "_80549_" או "_80549.pdf"
  ];
  
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  return null;
}

/**
 * חילוץ תיאור קצר
 */
function extractDescription(subject, originalFileName) {
  if (!subject) return null;
  
  // מילות מפתח לתיאור
  const keywords = [
    'קבלה', 'חשבונית', 'invoice', 'receipt', 
    'הוראה', 'תשלום', 'payment', 'bill'
  ];
  
  for (const keyword of keywords) {
    if (subject.toLowerCase().includes(keyword.toLowerCase())) {
      return keyword;
    }
  }
  
  // אם אין מילת מפתח, קח מילה ראשונה מהנושא
  const words = subject.split(/\s+/).filter(w => w.length >= 3);
  if (words.length > 0) {
    return sanitizeFileName(words[0]).substring(0, 15);
  }
  
  return null;
}

/**
 * ניקוי שם קובץ מתווים לא חוקיים
 */
function sanitizeFileName(name) {
  if (!name) return '';
  
  return name
    .replace(/[<>:"\/\\|?*]/g, '_')  // תווים לא חוקיים → _
    .replace(/\s+/g, '_')             // רווחים → _
    .replace(/_+/g, '_')              // _ כפולים → _
    .replace(/^_|_$/g, '');           // הסרת _ מההתחלה/סוף
}
