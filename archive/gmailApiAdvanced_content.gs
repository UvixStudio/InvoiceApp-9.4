// === Gmail API Advanced - ×¡×¨×™×§×” ×ž×”×™×¨×” ×•×ž×ª×§×“×ž×ª ===

// === MAIN ADVANCED SCAN FUNCTION ===
function processInvoicesGmailApi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const log = ss.getSheetByName(LOG_SHEET_NAME);
  if (!settings || !log) {
    SpreadsheetApp.getUi().alert("âŒ ×’×™×œ×™×•× ×•×ª Settings ××• Log ×—×¡×¨×™×. ×× × ×”×¤×¢×œ ××ª 'Setup Sheets' ×ª×—×™×œ×”.");
    return;
  }

  // ×ª×™×¢×•×“ ×ª×—×™×œ×ª ×¡×¨×™×§×”
  _logMessage("ðŸš€ ×”×ª×—×œ×ª ×¡×¨×™×§×ª ×—×©×‘×•× ×™×•×ª (Gmail API Advanced)");
  _logMessage(`ðŸ“… ×–×ž×Ÿ ×”×ª×—×œ×”: ${new Date().toLocaleString('he-IL')}`);
  
  const startTime = new Date();

  const startStr = settings.getRange("A2").getValue();
  const endStr = settings.getRange("B2").getValue();
  if (!startStr || !endStr) {
    SpreadsheetApp.getUi().alert("âŒ ×™×© ×œ×ž×œ× ×ª××¨×™×›×™× ×ª×§×¤×™× ×‘×’×™×œ×™×•×Ÿ Settings.");
    return;
  }

  const enKeywords = settings.getRange("C2:C").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
  const heKeywords = settings.getRange("D2:D").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
  const exclude = settings.getRange("E2:E").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());
  const approved = settings.getRange("F2:F").getValues().flat().filter(Boolean).map(e => e.toLowerCase().trim());

  const startDate = new Date(startStr);
  const endDate = new Date(endStr);
  endDate.setHours(23, 59, 59, 999);

  if (startDate.getTime() >= endDate.getTime()) {
    SpreadsheetApp.getUi().alert("âŒ ×ª××¨×™×š ×”×ª×—×œ×” ×—×™×™×‘ ×œ×”×™×•×ª ×œ×¤× ×™ ×ª××¨×™×š ×”×¡×™×•×.");
    return;
  }

  const tabName = `invoices ${formatDateShort(startDate)}-${formatDateShort(endDate)}`;
  let sheet = ss.getSheetByName(tabName);

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(["Action", "Category", "Sender Name", "Sender Email", "Date", "Subject", "PDF Link", "Email ID", "Attachment Name"]);
    
    // ×¦'×§×‘×•×§×¡×™× ×™×ª×•×•×¡×¤×• ××—×¨×™ ×›×ª×™×‘×ª ×”× ×ª×•× ×™× - ×œ× ×›××Ÿ!
    // sheet.getRange("A2:A1000").insertCheckboxes(); // ×”×•×¡×¨ - ×’×•×¨× ×œ×‘×¢×™×™×ª ×©×•×¨×” 1000
    
    sheet.setFrozenRows(1);
    sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#ddebf7");
  } else {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
  }

  // ×¡×¨×™×§×ª ×ž×™×™×œ×™×
  const query = `after:${formatDateGmail(startDate)} before:${formatDateGmail(endDate)} has:attachment`;
  const threads = GmailApp.search(query, 0, 500);
  _logMessage(log, `ðŸ“¨ Gmail connected. ${threads.length} threads found for query: "${query}"`);

  let inserted = 0, excludedCount = 0, noMatch = 0, processedCount = 0;

  SpreadsheetApp.getActiveSpreadsheet().toast("ðŸ” ×¡×¨×™×§×” ×”×ª×—×™×œ×”...");

  // ×™×¦×™×¨×ª ×¨×©×™×ž×” ×—×“×©×” ×œ×—×©×‘×•× ×™×•×ª
  const invoices = [];

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    processedCount++;
    SpreadsheetApp.getActiveSpreadsheet().toast(`×ž×¢×‘×“ ×ž×™×™×œ ${processedCount} ×ž×ª×•×š ${threads.length}...`, "×¢×™×‘×•×“", -1);

    try {
      const messages = thread.getMessages();
      for (let j = 0; j < messages.length; j++) {
        const msg = messages[j];
        const email = _parseSender(msg.getFrom()).toLowerCase();
        const subject = msg.getSubject();
        const msgId = msg.getId();
        const date = msg.getDate();
        const body = msg.getPlainBody();
        const htmlBody = msg.getBody();

        if (exclude.includes(email)) {
          excludedCount++;
          _logMessage(log, `âš  Skipped: [${email}] - Subject: "${subject}" - reason: excluded email`);
          continue;
        }

        if (isPrivateEmail(email)) {
          excludedCount++;
          _logMessage(log, `âš  Skipped: [${email}] - Subject: "${subject}" - reason: private email domain`);
          continue;
        }

        const keywordMatch = [...enKeywords, ...heKeywords].some(k =>
          subject.toLowerCase().includes(k) || (body && body.toLowerCase().includes(k))
        );
        const approvedSender = approved.includes(extractSenderName(email));

        if (!keywordMatch && !approvedSender) {
          noMatch++;
          _logMessage(log, `âš  Skipped: [${email}] - Subject: "${subject}" - reason: no keyword/approved sender match`);
          continue;
        }

        // Handle PDF attachment and link extraction
        const pdfAttachment = hasPdfAttachment(msg);
        let pdfLink = null;
        let attachmentName = "×œ×™× ×§ ×ž×”×ž×™×™×œ";

        if (pdfAttachment) {
          try {
            const driveFile = DriveApp.createFile(pdfAttachment);
            pdfLink = driveFile.getUrl();
            attachmentName = pdfAttachment.getName();
            _logMessage(log, `âœ… Saved PDF attachment to Drive: ${attachmentName}, Link: ${pdfLink}`);
          } catch (driveError) {
            _logMessage(log, `âŒ Error saving PDF attachment to Drive: ${driveError.message}`);
            continue;
          }
        } else {
          pdfLink = extractInvoiceLinkFromHtml(htmlBody);
          if (!pdfLink) {
            pdfLink = extractInvoiceLinkFromPlainText(body);
          }
          if (pdfLink) {
            _logMessage(log, `âœ… Extracted link from email body: ${pdfLink}`);
            attachmentName = "×œ×™× ×§ ×ž×”×ž×™×™×œ";
          }
        }

        if (!pdfLink) {
          noMatch++;
          _logMessage(log, `âš  Skipped: [${email}] - Subject: "${subject}" - reason: no PDF attachment or link found`);
          continue;
        }

        // ×–×™×”×•×™ ×ž×§×•×ž×™/×‘×™× ×œ××•×ž×™
        const isLocal = isLocalEmail(email, subject, body);
        const category = isLocal ? "ðŸ”· Local" : "ðŸŒ International";
        const senderName = extractSenderName(email);

        invoices.push({
          category: category,
          senderName: senderName,
          email: email,
          date: formatDateInput(date),
          subject: msg.getSubject(),
          link: pdfLink,
          msgId: msgId,
          attachmentName: attachmentName
        });
        inserted++;
      }
    } catch (e) {
      _logMessage(log, `âŒ Error processing thread [${thread.getId()}]: ${e.message}`);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("ðŸ” ×¡×¨×™×§×” ×”×¡×ª×™×™×ž×”. ×ž×ž×™×™×Ÿ ×•×ž×•×¡×™×£ ×œ×’×™×œ×™×•×Ÿ...", "×¢×™×‘×•×“", 5);

  // ×¡×™×“×•×¨ ×œ×¤×™ ×ª××¨×™×š
  invoices.sort((a, b) => new Date(a.date) - new Date(b.date));

  // ×”×•×¡×¤×” ×œ×’×™×œ×™×•×Ÿ
  if (invoices.length > 0) {
    const dataToAppend = invoices.map(inv => [
      false, // Action checkbox
      inv.category,
      inv.senderName,
      inv.email,
      inv.date,
      inv.subject,
      inv.link,
      inv.msgId,
      inv.attachmentName
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
  }

  // ×¡×™×›×•×
  if (sheet.getRange(1, 1).getDisplayValue().includes("× ×•×¡×¤×•:")) {
    sheet.deleteRow(1);
  }
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1, 1, 9).merge()
    .setValue(`ðŸ§¾ × ×•×¡×¤×•: ${inserted} | ×”×•×—×¨×’×•: ${excludedCount} | ×œ×œ× ×”×ª××ž×”: ${noMatch}`)
    .setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");

  // ×¡×™×›×•× ×ž×¤×•×¨×˜ ×‘×œ×•×’
  const endTime = new Date();
  const duration = Math.round((endTime - startTime) / 1000);
  
  _logMessage("ðŸ“Š ===== ×¡×™×›×•× ×¡×¨×™×§×” (Gmail API Advanced) =====");
  _logMessage(`â±ï¸ ×–×ž×Ÿ ×¡×™×•×: ${endTime.toLocaleString('he-IL')}`);
  _logMessage(`â±ï¸ ×ž×©×š ×–×ž×Ÿ: ${duration} ×©× ×™×•×ª`);
  _logMessage(`âœ… × ×•×¡×¤×• ×œ×’×™×œ×™×•×Ÿ: ${inserted} ×¨×©×•×ž×•×ª`);
  _logMessage(`ðŸ”´ ×”×•×—×¨×’×• (×ž×™×™×œ×™× ×œ× ×¨×¦×•×™×™×): ${excludedCount} ×¨×©×•×ž×•×ª`);
  _logMessage(`âšª ×“×•×œ×’×• (×œ×œ× ×”×ª××ž×”): ${noMatch} ×¨×©×•×ž×•×ª`);
  _logMessage(`ðŸ“§ ×¡×”"×› ×ž×™×™×œ×™× × ×‘×“×§×•: ${inserted + excludedCount + noMatch}`);
  _logMessage("âœ… ×¡×¨×™×§×” ×”×•×©×œ×ž×” ×‘×”×¦×œ×—×”!");
  _logMessage("================================================");
  
  SpreadsheetApp.getActiveSpreadsheet().toast("ðŸŽ‰ ×”×¡×¨×™×§×” ×”×¡×ª×™×™×ž×”!", "×”×¡×ª×™×™×", 10);
  SpreadsheetApp.getUi().alert(`ðŸŽ‰ ×”×¡×¨×™×§×” ×”×¡×ª×™×™×ž×”.\nâœ… × ×•×¡×¤×•: ${inserted}\nðŸ”´ ×”×•×—×¨×’×•: ${excludedCount}\nâšª ×œ×œ× ×”×ª××ž×”: ${noMatch}`);
}

function writeInvoicesToSheet(invoices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Invoices") || ss.insertSheet("Invoices");
  sheet.clear();

  const headers = ["Sender", "Subject", "Date", "Invoice Name", "Drive Link"];
  sheet.appendRow(headers);

  for (const invoice of invoices) {
    sheet.appendRow([invoice.sender, invoice.subject, invoice.date, invoice.name, invoice.link]);
  }
}

// === ×§×¨×™××ª ×”×’×“×¨×•×ª ×ž×ª×§×“×ž×ª ===
function readAdvancedSettings(settings) {
  const startStr = settings.getRange('A2').getValue();
  const endStr = settings.getRange('B2').getValue();
  
  if (!startStr || !endStr) {
    _logMessage('âŒ ×—×¡×¨×™× ×ª××¨×™×›×™× ×‘×’×™×œ×™×•×Ÿ Settings');
    SpreadsheetApp.getUi().alert('âŒ ×™×© ×œ×ž×œ× ×ª××¨×™×›×™× ×ª×§×¤×™× ×‘×’×™×œ×™×•×Ÿ Settings.');
    return null;
  }
  
  const startDate = new Date(startStr);
  const endDate = new Date(endStr);
  endDate.setHours(23, 59, 59, 999);
  
  const enKeywords = settings.getRange('C2:C').getValues().flat().filter(Boolean);
  const heKeywords = settings.getRange('D2:D').getValues().flat().filter(Boolean);
  const excludedKeywords = settings.getRange('E2:E').getValues().flat().filter(Boolean);
  const excludeEmails = settings.getRange('F2:F').getValues().flat().filter(Boolean);
  
  const allKeywords = [...enKeywords, ...heKeywords];
  
  _logMessage(`â„¹ï¸ ðŸ”¤ ×ž×™×œ×•×ª ×ž×¤×ª×—: EN(${enKeywords.length}), HE(${heKeywords.length})`);
  _logMessage(`â„¹ï¸ ðŸš« ×ž×•×—×¨×’×™×: ${excludeEmails.length} ×ž×™×™×œ×™×, ${excludedKeywords.length} ×ž×™×œ×™×`);
  
  return {
    startDate,
    endDate,
    allKeywords,
    excludedKeywords,
    excludeEmails,
    tabName: `invoices ${formatDateShort(startDate)}-${formatDateShort(endDate)}`
  };
}

// === ×™×¦×™×¨×ª ×’×™×œ×™×•×Ÿ ×ª×•×¦××•×ª ×ž×ª×§×“× ===
function createAdvancedResultsSheet(ss, tabName) {
  _logMessage(`â„¹ï¸ ðŸ“‘ ×™×•×¦×¨ ×’×™×œ×™×•×Ÿ ×ž×ª×§×“×: ${tabName}`);
  
  let sheet = ss.getSheetByName(tabName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(tabName);
  }
  
  // ×›×•×ª×¨×•×ª ×ž×ª×§×“×ž×•×ª
  const headers = [
    "Action", "Category", "Company Name", "Email Sender", "Date", "Subject",
    "Invoice Number", "Amount", "Currency", "PDF Link", "Email ID", "Attachment Name", "Validation Method"
  ];
  
  try {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ×¢×™×¦×•×‘ ×ž×ª×§×“×
    const headerRange = sheet.getRange("A1:L1");
    headerRange.setFontWeight("bold")
               .setBackground("#4285f4")
               .setFontColor("white")
               .setHorizontalAlignment("center");
    
    // ×¦'×§×‘×•×§×¡×™× ×™×ª×•×•×¡×¤×• ××—×¨×™ ×›×ª×™×‘×ª ×”× ×ª×•× ×™×
    // sheet.getRange("A2:A100").insertCheckboxes(); // ×”×•×¡×¨ - ×’×•×¨× ×œ×‘×¢×™×™×ª ×©×•×¨×” 1000
    
    // ×¨×•×—×‘ ×¢×ž×•×“×•×ª
    sheet.setColumnWidth(1, 60);  // Action
    sheet.setColumnWidth(2, 100); // Category
    sheet.setColumnWidth(3, 200); // Company Name
    sheet.setColumnWidth(4, 180); // Email Sender
    sheet.setColumnWidth(5, 100); // Date
    sheet.setColumnWidth(6, 300); // Subject
    sheet.setColumnWidth(7, 120); // Invoice Number
    sheet.setColumnWidth(8, 100); // Amount
    sheet.setColumnWidth(9, 80);  // Currency
    sheet.setColumnWidth(10, 200); // PDF Link
    sheet.setColumnWidth(11, 120); // Email ID
    sheet.setColumnWidth(12, 150); // Attachment Name
    sheet.setColumnWidth(13, 200); // Validation Method
    
    sheet.setFrozenRows(1);
    _logMessage("âœ… ×’×™×œ×™×•×Ÿ ×ž×ª×§×“× × ×•×¦×¨ ×‘×”×¦×œ×—×”");
    
    return sheet;
  } catch (e) {
    _logMessage(`âŒ ×©×’×™××” ×‘×™×¦×™×¨×ª ×’×™×œ×™×•×Ÿ: ${e.message}`);
    return null;
  }
}

// === ×¡×¨×™×§×” ×ž×ª×§×“×ž×ª ×¢× Gmail API ===
function performAdvancedGmailScan(config) {
  _logMessage("ðŸ” ×ž×ª×—×™×œ ×¡×¨×™×§×” ×ž×ª×§×“×ž×ª ×¢× Gmail API");
  
  // ×‘× ×™×™×ª ×©××™×œ×ª×” ×ž×ª×§×“×ž×ª
  const keywordQuery = config.allKeywords.map(k => `"${k}"`).join(' OR ');
  const excludeQuery = config.excludeEmails.map(e => `-from:${e}`).join(' ');
  
  // ×©××™×œ×ª×” ×ž×ª×§×“×ž×ª ×¢× ×—×™×¤×•×© ×—×›×
  const advancedQuery = `
    after:${formatDateGmail(config.startDate)} 
    before:${formatDateGmail(config.endDate)} 
    has:attachment 
    (${keywordQuery} OR "×¡×›×•×" OR "â‚ª" OR "$" OR "total" OR "amount" OR "billing" OR "×—×™×•×‘")
    ${excludeQuery}
  `.replace(/\s+/g, ' ').trim();
  
  _logMessage(`ðŸ” ×©××™×œ×ª×” ×ž×ª×§×“×ž×ª: ${advancedQuery}`);

  try {
    // ×©×™×ž×•×© ×‘-Gmail API ×‘×ž×§×•× GmailApp
    const response = Gmail.Users.Messages.list('me', {
      q: advancedQuery,
      maxResults: 500  // ×¤×™ 2.5 ×™×•×ª×¨ ×ž-GmailApp!
    });
    
    const messageIds = response.messages || [];
    _logMessage(`ðŸ“§ × ×ž×¦××• ${messageIds.length} ×ž×™×™×œ×™× ×¢× Gmail API`);
    
    if (messageIds.length === 0) {
      return { messages: [] };
    }
    
    // ×§×‘×œ×ª ×¤×¨×˜×™ ×”×ž×™×™×œ×™× ×‘×§×‘×•×¦×•×ª (×ž×”×™×¨ ×™×•×ª×¨)
    const messages = [];
    const batchSize = 10;
    
    for (let i = 0; i < messageIds.length; i += batchSize) {
      const batch = messageIds.slice(i, i + batchSize);
      
      for (const msgRef of batch) {
        try {
          const message = Gmail.Users.Messages.get('me', msgRef.id, {
            format: 'full'
          });
          
          // ×¡×™× ×•×Ÿ ×œ×¤×™ ×ª××¨×™×š
          const msgDate = new Date(parseInt(message.internalDate));
          if (msgDate >= config.startDate && msgDate <= config.endDate) {
            messages.push(message);
          }
        } catch (e) {
          _logMessage(`âš ï¸ ×©×’×™××” ×‘×§×¨×™××ª ×ž×™×™×œ ${msgRef.id}: ${e.message}`);
        }
      }
      
      // ×”×¦×’×ª ×”×ª×§×“×ž×•×ª
      if (i % 50 === 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast(`ðŸ“§ ×¢×•×‘×“... ${i}/${messageIds.length}`, "Gmail API", 1);
      }
    }
    
    _logMessage(`ðŸ“‹ ×—×•×œ×¦×• ${messages.length} ×ž×™×™×œ×™× ×ž×ª××™×ž×™× ×œ×¢×™×‘×•×“`);
    return { messages };
    
  } catch (e) {
    _logMessage(`âŒ ×©×’×™××” ×‘×¡×¨×™×§×” ×ž×ª×§×“×ž×ª: ${e.message}`);
    SpreadsheetApp.getUi().alert(`âŒ ×©×’×™××” ×‘×¡×¨×™×§×”: ${e.message}`);
    return { messages: [] };
  }
}

// === ×¢×™×‘×•×“ ×ž×ª×§×“× ×¢× OCR ===
function processAdvancedMessages(messages, config, sheet) {
  _logMessage(`â„¹ï¸ ðŸ”„ ×ž×ª×—×™×œ ×¢×™×‘×•×“ ×ž×ª×§×“× ×©×œ ${messages.length} ×ž×™×™×œ×™×`);
  
  let addedCount = 0;
  let skippedCount = 0;
  let currentRow = 2;
  
  for (let i = 0; i < messages.length; i++) {
    const message = messages[i];
    
    try {
      // ×—×™×œ×•×¥ ×ž×˜××“×˜×” ×ž×”×ž×™×™×œ
      _logMessage(`ðŸ” ×ž×—×œ×¥ ×ž×˜××“×˜×” ×¢×‘×•×¨ ×ž×™×™×œ: ${message.id}`);
      const emailData = extractEmailMetadata(message);
      _logMessage(`ðŸ” ×ª×•×¦××ª ×ž×˜××“×˜×”: ${emailData ? '×”×¦×œ×™×—' : '× ×›×©×œ - null'}`);
      
      if (!emailData) {
        _logMessage(`âš ï¸ ×“×•×œ×’ - ×œ× ×”×¦×œ×™×— ×œ×—×œ×¥ ×ž×˜××“×˜×”: ${message.id}`);
        skippedCount++;
        continue;
      }
      
      // ×‘×“×™×§×ª PDF
      _logMessage(`ðŸ” ×ž×—×¤×© PDF attachment ×¢×‘×•×¨ ×ž×™×™×œ: ${message.id}`);
      _logMessage(`ðŸ” DEBUG: message.payload.parts = ${message.payload.parts ? message.payload.parts.length : 'null'}`);
      const pdfAttachment = findPdfAttachment(message);
      _logMessage(`ðŸ” ×ª×•×¦××ª ×—×™×¤×•×© PDF: ${pdfAttachment ? '× ×ž×¦× - ' + pdfAttachment.filename : '×œ× × ×ž×¦×'}`);
      if (!pdfAttachment) {
        _logMessage(`âš ï¸ ×“×•×œ×’: ${emailData.senderEmail} - ××™×Ÿ PDF`);
        skippedCount++;
        continue;
      }
      
      // OCR ×œ×—×™×œ×•×¥ × ×ª×•× ×™× ×ž×”×—×©×‘×•× ×™×ª
      _logMessage(`ðŸ” ×ž×ª×—×™×œ OCR ×¢×‘×•×¨: ${pdfAttachment.filename}`);
      const invoiceData = extractInvoiceDataWithOCR(message, pdfAttachment);
      _logMessage(`ðŸ” ×ª×•×¦××ª OCR: ${invoiceData ? '×”×¦×œ×™×—' : '× ×›×©×œ - null'}`);
      
      // ×× OCR ×”×—×–×™×¨ null, ×–×” ××•×ž×¨ ×©×–×” ×œ× ×—×©×‘×•× ×™×ª ××ž×™×ª×™×ª
      if (!invoiceData) {
        const timestamp = new Date().toLocaleString('he-IL');
        const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
        if (log) log.appendRow([timestamp, `âš ï¸ ×“×•×œ×’ - ×œ× ×—×©×‘×•× ×™×ª: ${emailData.senderEmail} - ${pdfAttachment.filename}`]);
        skippedCount++;
        continue;
      }
      
      // ×—×™×œ×•×¥ ×¡×›×•×ž×™× ×—×›× (×ž×”×™×¨ + OCR ×‘×ž×™×“×ª ×”×¦×•×¨×š)
      const amountData = extractAmountSmart(message, pdfAttachment);
      if (amountData.found) {
        invoiceData.amount = amountData.amount;
        invoiceData.currency = amountData.currency;
        invoiceData.amountMethod = amountData.method;
      }
      
      // ×©×™×œ×•×‘ ×”× ×ª×•× ×™×
      _logMessage(`ðŸ”— ×ž× ×¡×” ×œ×™×¦×•×¨ ×œ×™× ×§ PDF ×¢×‘×•×¨ ×ž×™×™×œ: ${message.id}`);
      let pdfLink = extractPdfLink(message);
      
      // ×× ×œ× ×”×¦×œ×™×— ×œ×™×¦×•×¨ ×œ×™× ×§ Drive, × × ×¡×” ×œ×™×¦×•×¨ ××•×ª×• ×™×©×™×¨×•×ª ×›××Ÿ
      if (!pdfLink) {
        _logMessage(`âš ï¸ extractPdfLink ×”×—×–×™×¨ null - ×ž× ×¡×” ×œ×™×¦×•×¨ Drive link ×™×©×™×¨×•×ª`);
        const pdfAttachment = findPdfAttachment(message);
        if (pdfAttachment) {
          let fileName = pdfAttachment.filename || `invoice-${message.id}.pdf`;
          
          try {
            // ×”×•×¨×“×ª ×”-attachment ×“×¨×š Gmail API
            _logMessage(`ðŸ”„ ×ž×•×¨×™×“ attachment: ${pdfAttachment.attachmentId}`);
            const attachmentData = Gmail.Users.Messages.Attachments.get('me', message.id, pdfAttachment.attachmentId);
            _logMessage(`ðŸ“¦ ×§×™×‘×œ × ×ª×•× ×™×: ${attachmentData.data ? '×™×© data' : '××™×Ÿ data'}`);
            
            if (!attachmentData.data) {
              throw new Error('No data in attachment response');
            }
            
            const decodedData = Utilities.base64Decode(attachmentData.data);
            const blob = Utilities.newBlob(decodedData, 'application/pdf', fileName);
            
            // ×©×ž×™×¨×” ×œ-Drive (My Drive root - ×›×ž×• ×‘×’×¨×¡×” ×”×§×•×“×ž×ª)
            _logMessage(`ðŸ’¾ ×©×•×ž×¨ ×§×•×‘×¥ ×‘-My Drive: ${fileName}`);
            const file = DriveApp.createFile(blob);
            
            file.setName(fileName);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            pdfLink = file.getUrl();
            _logMessage(`âœ… × ×•×¦×¨ ×œ×™× ×§ Drive: ${pdfLink}`);
            
          } catch (err) {
            pdfLink = 'ERROR: ' + err.message;
            _logMessage(`âŒ ×©×’×™××” ×‘×©×ž×™×¨×ª ×”×§×•×‘×¥ ×œ-Drive: ${err.message}`);
          }
        } else {
          pdfLink = "";
          _logMessage(`âŒ ×œ× × ×ž×¦× PDF attachment ×‘×›×œ×œ`);
        }
      }
      
      _logMessage(`ðŸ”— ×ª×•×¦××ª ×œ×™× ×§ PDF: ${pdfLink ? '×”×¦×œ×™×— - ' + pdfLink.substring(0, 50) + '...' : '× ×›×©×œ - ×¨×™×§'}`);
      const finalData = {
        ...emailData,
        ...invoiceData,
        pdfLink: pdfLink
      };
      
      // ×œ×•×’ ×¡×•×’ ×”×œ×™× ×§ ×©× ×•×¦×¨
      if (pdfLink.includes('drive.google.com')) {
        _logMessage(`âœ… Drive Link: ${finalData.companyName || finalData.senderName} - ${pdfLink}`);
      } else if (pdfLink.includes('mail.google.com')) {
        _logMessage(`âš ï¸ Gmail Link: ${finalData.companyName || finalData.senderName} - ${pdfLink}`);
      } else {
        _logMessage(`â“ Unknown Link: ${finalData.companyName || finalData.senderName} - ${pdfLink}`);
      }
      
      // ×›×ª×™×‘×” ×œ×’×™×œ×™×•×Ÿ
      const rowData = [
        false, // Action checkbox
        isLocalEmail(finalData.senderEmail) ? "ðŸ”· Local" : "ðŸŒ International",
        finalData.companyName || finalData.senderName,
        finalData.senderEmail,
        formatDateInput(finalData.date),
        finalData.subject,
        finalData.pdfLink,
        finalData.messageId,
        finalData.attachmentName || pdfAttachment.filename,
        finalData.amount || "",
        finalData.currency || "",
        finalData.amountMethod || "not_extracted"
      ];
      
      _logMessage(`â„¹ï¸ ðŸ“ ×›×•×ª×‘: ${finalData.companyName || finalData.senderName} - ${finalData.amount || '×œ×œ× ×¡×›×•×'}`);
      
      sheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
      
      addedCount++;
      currentRow++;
      
      // ×¨×™×¢× ×•×Ÿ ×ª×§×•×¤×ª×™
      if (addedCount % 5 === 0) {
        SpreadsheetApp.flush();
        SpreadsheetApp.getActiveSpreadsheet().toast(`ðŸ“ × ×›×ª×‘×• ${addedCount} ×—×©×‘×•× ×™×•×ª...`, "×¢×™×‘×•×“", 1);
      }
      
    } catch (e) {
      _logMessage(`âŒ ×©×’×™××” ×‘×¢×™×‘×•×“ ×ž×™×™×œ: ${e.message}`);
      skippedCount++;
    }
  }
  
  SpreadsheetApp.flush();
  
  // ×”×•×¡×¤×ª ×¦'×§×‘×•×§×¡×™× ×¨×§ ×œ×©×•×¨×•×ª ×©× ×›×ª×‘×• ×‘×¤×•×¢×œ
  if (addedCount > 0) {
    try {
      sheet.getRange(2, 1, addedCount, 1).insertCheckboxes();
      _logMessage(`âœ… ×¦'×§×‘×•×§×¡×™× × ×•×¦×¨×• ×œ×©×•×¨×•×ª 2-${addedCount + 1} (${addedCount} ×¨×©×•×ž×•×ª)`);
    } catch (e) {
      _logMessage(`âš ï¸ ×©×’×™××” ×‘×™×¦×™×¨×ª ×¦'×§×‘×•×§×¡×™×: ${e.message}`);
    }
  }
  
  const results = {
    added: addedCount,
    skipped: skippedCount,
    total: messages.length
  };
  
  _logMessage(`âœ… ðŸ“Š ×ª×•×¦××•×ª ×¢×™×‘×•×“ ×ž×ª×§×“×: ${addedCount} × ×•×¡×¤×•, ${skippedCount} ×“×•×œ×’×•`);
  
  return results;
}

// === ×—×™×œ×•×¥ ×ž×˜××“×˜×” ×ž×”×ž×™×™×œ ===
function extractEmailMetadata(message) {
  try {
    const headers = message.payload.headers;
    const getHeader = (name) => {
      const header = headers.find(h => h.name.toLowerCase() === name.toLowerCase());
      return header ? header.value : '';
    };
    
    const from = getHeader('From');
    const subject = getHeader('Subject');
    const date = new Date(parseInt(message.internalDate));
    
    return {
      messageId: message.id,
      senderEmail: _parseSender(from),
      senderName: extractSenderName(from),
      subject: subject,
      date: date
    };
  } catch (e) {
    _logMessage(`âš ï¸ ×©×’×™××” ×‘×—×™×œ×•×¥ ×ž×˜××“×˜×”: ${e.message}`);
    return null;
  }
}

// === ×ž×¦×™××ª ×§×•×‘×¥ PDF ===
function findPdfAttachment(message) {
  try {
    _logMessage(`ðŸ” DEBUG findPdfAttachment: message.payload = ${message.payload ? 'exists' : 'null'}`);
    
    if (!message.payload) {
      _logMessage(`âŒ ××™×Ÿ payload ×‘×ž×™×™×œ`);
      return null;
    }
    
    if (!message.payload.parts) {
      _logMessage(`âŒ ××™×Ÿ parts ×‘-payload`);
      return null;
    }
    
    _logMessage(`ðŸ” DEBUG: × ×ž×¦××• ${message.payload.parts.length} parts`);
    
    for (let i = 0; i < message.payload.parts.length; i++) {
      const part = message.payload.parts[i];
      if (part.filename && part.filename.toLowerCase().includes('.pdf')) {
        // ×× ×™×© attachmentId, × ×—×–×™×¨ ××ª ×”×ž×™×“×¢ ×‘×œ×‘×“ (×”×”×•×¨×“×” ×ª×”×™×” ×‘×§×•×“ ×”×¨××©×™)
        if (part.body.attachmentId) {
          _logMessage(`ðŸ“Ž × ×ž×¦× PDF ×¢× attachmentId: ${part.filename} (${part.body.attachmentId})`);
          return {
            attachmentId: part.body.attachmentId,
            filename: part.filename,
            size: part.body.size,
            partId: `0.${i + 1}`,
            mimeType: part.mimeType || 'application/pdf'
          };
        } else if (part.body.data) {
          // ×× ×”× ×ª×•× ×™× ×›×‘×¨ ×§×™×™×ž×™× ×‘-body
          return {
            attachmentId: part.body.attachmentId,
            filename: part.filename,
            size: part.body.size,
            partId: `0.${i + 1}`,
            body: { data: part.body.data }
          };
        } else {
          _logMessage(`âš ï¸ ×œ× × ×ž×¦××• × ×ª×•× ×™ ×§×•×‘×¥ ×¢×‘×•×¨: ${part.filename}`);
          return null;
        }
      }
    }
    return null;
  } catch (e) {
    _logMessage(`âš ï¸ ×©×’×™××” ×‘×—×™×¤×•×© PDF: ${e.message}`);
    return null;
  }
}

// === ×—×™×œ×•×¥ ×œ×™× ×§ PDF ×ž×”×ž×™×™×œ ===
function extractPdfLink(message) {
  try {
    _logMessage(`ðŸ” extractPdfLink called for message: ${message.id}`);
    
    // ×©×™×˜×” 1: GmailApp (×¤×©×•×˜ ×•×™×¢×™×œ - ×›×ž×• ×‘×§×•×“ ×”×™×©×Ÿ)
    try {
      _logMessage(`ðŸ”„ ×ž× ×¡×” GmailApp method...`);
      const gmailMessage = GmailApp.getMessageById(message.id);
      const attachments = gmailMessage.getAttachments();
      _logMessage(`ðŸ“Ž × ×ž×¦××• ${attachments.length} attachments`);
      
      const pdfAttachment = attachments.find(att => 
        att.getName().toLowerCase().endsWith('.pdf') || 
        att.getContentType().toLowerCase().includes('pdf')
      );
      
      if (pdfAttachment) {
        _logMessage(`âœ… × ×ž×¦× PDF: ${pdfAttachment.getName()}`);
        const blob = pdfAttachment.getBlob();
        const driveUrl = createDriveFile(blob, pdfAttachment.getName());
        if (driveUrl) {
          _logMessage(`âœ… ×”×•×¢×œ×” ×œ-Drive: ${driveUrl}`);
          return driveUrl;
        }
      } else {
        _logMessage(`âŒ ×œ× × ×ž×¦× PDF attachment`);
      }
    } catch (e) {
      _logMessage(`âŒ ×©×’×™××” ×‘-GmailApp: ${e.message}`);
    }
    
    // ×©×™×˜×” 2: Gmail API (backup)
    try {
      _logMessage(`ðŸ”„ ×ž× ×¡×” Gmail API method...`);
      const pdfAttachment = findPdfAttachment(message);
      if (pdfAttachment && pdfAttachment.body && pdfAttachment.body.data) {
        _logMessage(`âœ… × ×ž×¦× PDF via API: ${pdfAttachment.filename}`);
        const blob = Utilities.newBlob(
          Utilities.base64Decode(pdfAttachment.body.data),
          'application/pdf',
          pdfAttachment.filename || 'invoice.pdf'
        );
        const driveUrl = createDriveFile(blob, pdfAttachment.filename);
        if (driveUrl) {
          _logMessage(`âœ… ×”×•×¢×œ×” ×œ-Drive via API: ${driveUrl}`);
          return driveUrl;
        }
      }
    } catch (e) {
      _logMessage(`âŒ ×©×’×™××” ×‘-Gmail API: ${e.message}`);
    }

    _logMessage(`âŒ ×œ× ×”×¦×œ×™×— ×œ×™×¦×•×¨ ×œ×™× ×§ Drive - ×ž×—×–×™×¨ null`);
    return null;
  } catch (e) {
    _logMessage(`âŒ ×©×’×™××” ×›×œ×œ×™×ª ×‘-extractPdfLink: ${e.message}`);
    return null;
  }
}

// ×¤×•× ×§×¦×™×™×ª ×¢×–×¨ ×œ×™×¦×™×¨×ª ×§×•×‘×¥ Drive
function createDriveFile(blob, filename) {
  try {
    // ×™×¦×™×¨×ª ×”×§×•×‘×¥ ×‘-My Drive (root) - ×›×ž×• ×‘×’×¨×¡×” ×”×§×•×“×ž×ª
    _logMessage(`ðŸ’¾ ×™×•×¦×¨ ×§×•×‘×¥ ×‘-My Drive: ${filename}`);
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    _logMessage(`âœ… × ×•×¦×¨ ×§×•×‘×¥ Drive: ${file.getUrl()}`);
    return file.getUrl();
  } catch (e) {
    _logMessage(`âŒ ×©×’×™××” ×‘×™×¦×™×¨×ª ×§×•×‘×¥ Drive: ${e.message}`);
    return null;
  }
}

// === ×ž×¢×¨×›×ª ×¡×™× ×•×Ÿ ×—×›×ž×” ×ž×©×•×œ×‘×ª ===
function smartInvoiceValidation(message, pdfAttachment) {
  const timestamp = new Date().toLocaleString('he-IL');
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  
  // ×—×™×œ×•×¥ × ×ª×•× ×™ ×”×ž×™×™×œ
  const emailData = extractEmailMetadata(message);
  const subject = emailData.subject.toLowerCase();
  const filename = pdfAttachment.filename.toLowerCase();
  
  if (log) log.appendRow([timestamp, `ðŸ” ×‘×•×“×§: × ×•×©×="${emailData.subject}" ×§×•×‘×¥="${pdfAttachment.filename}"`]);
  
  // ×©×œ×‘ 1: ×‘×“×™×§×” ×ž×”×™×¨×” - ×¨×©×™×ž×” ×©×—×•×¨×” (×“×—×™×™×” ×ž×™×™×“×™×ª)
  const blacklistSubjects = [
    'quote', 'quotation', 'estimate', 'proposal', 'offer', 'draft',
    '×”×¦×¢×ª ×ž×—×™×¨', '×”×¦×¢×”', '×”×¢×¨×›×”', '×”×¦×¢×ª ×¢×‘×•×“×”', '×˜×™×•×˜×”',
    'contract', 'agreement', '×—×•×–×”', '×”×¡×›×', '×ª× ××™×',
    '×˜×•×¤×¡ ×ª×™×§×•×Ÿ', '×ª×™×§×•×Ÿ', 'correction', 'amendment'
  ];
  
  const blacklistFiles = [
    'quote', 'quotation', 'estimate', 'proposal', 'offer', 'draft',
    '×”×¦×¢×”', '×”×¦×¢×ª_×ž×—×™×¨', '×˜×™×•×˜×”', '×”×¡×›×', '×—×•×–×”', '×¤×•×œ×™×¡×”',
    'agreement', 'contract', 'policy', 'terms', 'conditions',
    '×ª×™×§×•×Ÿ', 'correction', 'amendment', '×˜×•×¤×¡'
  ];
  
  // ×‘×“×™×§×ª ×¨×©×™×ž×” ×©×—×•×¨×” - × ×•×©×
  for (const blackword of blacklistSubjects) {
    if (subject.includes(blackword)) {
      if (log) log.appendRow([timestamp, `âŒ × ×“×—×” ×ž×™×“ - × ×•×©× ×—×¡×•×: "${blackword}"`]);
      return { decision: 'REJECT', reason: `× ×•×©× ×ž×›×™×œ: ${blackword}` };
    }
  }
  
  // ×‘×“×™×§×ª ×¨×©×™×ž×” ×©×—×•×¨×” - ×§×•×‘×¥
  for (const blackword of blacklistFiles) {
    if (filename.includes(blackword)) {
      if (log) log.appendRow([timestamp, `âŒ × ×“×—×” ×ž×™×“ - ×§×•×‘×¥ ×—×¡×•×: "${blackword}"`]);
      return { decision: 'REJECT', reason: `×§×•×‘×¥ ×ž×›×™×œ: ${blackword}` };
    }
  }
  
  // ×©×œ×‘ 2: ×‘×“×™×§×” ×ž×”×™×¨×” - ×¨×©×™×ž×” ×œ×‘× ×” (××™×©×•×¨ ×ž×™×™×“×™)
  const whitelistSubjects = [
    'invoice', 'receipt', 'bill', 'payment', 'statement', 'billing',
    '×—×©×‘×•× ×™×ª', '×§×‘×œ×”', '×—×©×‘×•×Ÿ', '×ª×©×œ×•×', '×—×™×•×‘'
  ];
  
  const whitelistFiles = [
    'invoice', 'receipt', 'bill', 'payment', 'statement',
    '×—×©×‘×•× ×™×ª', '×§×‘×œ×”', '×—×©×‘×•×Ÿ', '×ª×©×œ×•×'
  ];
  
  let subjectScore = 0;
  let fileScore = 0;
  let foundSubjectKeywords = [];
  let foundFileKeywords = [];
  
  // × ×™×§×•×“ × ×•×©×
  for (const goodword of whitelistSubjects) {
    if (subject.includes(goodword)) {
      subjectScore++;
      foundSubjectKeywords.push(goodword);
    }
  }
  
  // × ×™×§×•×“ ×§×•×‘×¥
  for (const goodword of whitelistFiles) {
    if (filename.includes(goodword)) {
      fileScore++;
      foundFileKeywords.push(goodword);
    }
  }
  
  // ×©×œ×‘ 3: ×”×—×œ×˜×” ×œ×¤×™ × ×™×§×•×“
  const totalScore = subjectScore + fileScore;
  
  // ××™×©×•×¨ ×ž×™×™×“×™ - × ×™×§×•×“ ×’×‘×•×”
  if (totalScore >= 2 || (subjectScore >= 1 && fileScore >= 1)) {
    if (log) log.appendRow([timestamp, `âœ… ××•×©×¨ ×ž×™×“ - × ×™×§×•×“ ×’×‘×•×”: × ×•×©×(${subjectScore}) + ×§×•×‘×¥(${fileScore}) = ${totalScore}`]);
    return { 
      decision: 'APPROVE', 
      reason: `× ×ž×¦××• ×ž×™×œ×•×ª ×ž×¤×ª×—: × ×•×©×[${foundSubjectKeywords.join(',')}] ×§×•×‘×¥[${foundFileKeywords.join(',')}]`,
      confidence: 'HIGH'
    };
  }
  
  // ××™×©×•×¨ ×‘×™× ×•× ×™ - ×™×© ×œ×¤×—×•×ª ×ž×™×œ×ª ×ž×¤×ª×— ××—×ª
  if (totalScore >= 1) {
    if (log) log.appendRow([timestamp, `âœ… ××•×©×¨ - × ×™×§×•×“ ×‘×™× ×•× ×™: ${totalScore} ×ž×™×œ×•×ª ×ž×¤×ª×—`]);
    return { 
      decision: 'APPROVE', 
      reason: `× ×ž×¦××” ×ž×™×œ×ª ×ž×¤×ª×—: ${foundSubjectKeywords.concat(foundFileKeywords).join(',')}`,
      confidence: 'MEDIUM'
    };
  }
  
  // ×‘×“×™×§×” × ×•×¡×¤×ª - ×ž×¡×¤×¨×™× ×•×ª××¨×™×›×™× (×¡×™×ž×Ÿ ×˜×•×‘)
  const hasNumbers = /\d{4}/.test(filename) && /\d{2}/.test(filename);
  const hasAmountInSubject = /[\d,]+\.?\d*\s*(?:â‚ª|×©×§×œ|×©"×—|\$|USD|â‚¬|EUR)/i.test(subject);
  
  if (hasNumbers || hasAmountInSubject) {
    if (log) log.appendRow([timestamp, `âœ… ××•×©×¨ - × ×ž×¦××• ×ž×¡×¤×¨×™×/×¡×›×•×ž×™×`]);
    return { 
      decision: 'APPROVE', 
      reason: hasAmountInSubject ? '× ×ž×¦× ×¡×›×•× ×‘× ×•×©×' : '× ×ž×¦××• ×ž×¡×¤×¨×™× ×•×ª××¨×™×›×™× ×‘×§×•×‘×¥',
      confidence: 'MEDIUM'
    };
  }
  
  // ×× ×”×’×¢× ×• ×¢×“ ×›××Ÿ - ×œ× ×‘×˜×•×—, ×¦×¨×™×š OCR
  if (log) log.appendRow([timestamp, `â“ ×œ× ×‘×˜×•×— - ×™×¢×‘×•×¨ ×œOCR: × ×•×©×="${emailData.subject}" ×§×•×‘×¥="${pdfAttachment.filename}"`]);
  return { 
    decision: 'OCR_NEEDED', 
    reason: '×œ× × ×ž×¦××• ×ž×¡×¤×™×§ ×¡×™×ž× ×™× - × ×“×¨×© OCR',
    confidence: 'UNKNOWN'
  };
}

// === OCR ×¤×©×•×˜ ×œ×ž×§×¨×™× ×ž×¤×•×§×¤×§×™× ===
function performOCRValidation(message, pdfAttachment) {
  const timestamp = new Date().toLocaleString('he-IL');
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  
  try {
    if (log) log.appendRow([timestamp, `ðŸ” OCR: ×‘×•×“×§ ×ª×•×›×Ÿ ×§×•×‘×¥ ${pdfAttachment.filename}`]);
    
    // ×›××Ÿ ×™×”×™×” OCR ××ž×™×ª×™ ×‘×¢×ª×™×“
    // ×œ×¢×ª ×¢×ª×” × ×¢×©×” ×‘×“×™×§×” ×¤×©×•×˜×” × ×•×¡×¤×ª
    
    const filename = pdfAttachment.filename.toLowerCase();
    
    // ×‘×“×™×§×•×ª OCR ×ž×“×•×ž×•×ª - ×‘×¢×ª×™×“ ×™×”×™×” OCR ××ž×™×ª×™
    const ocrKeywords = [
      'inv', 'rec', 'bill', 'pay', // ×§×™×¦×•×¨×™× ×©×¢×œ×•×œ×™× ×œ×”×™×•×ª ×‘×©× ×§×•×‘×¥
      '2025', '2024', // ×©× ×™×
      'pdf' // ×–×” PDF ××– ×™×© ×¡×™×›×•×™ ×©×–×” ×ž×¡×ž×š ×¨×©×ž×™
    ];
    
    let ocrScore = 0;
    for (const keyword of ocrKeywords) {
      if (filename.includes(keyword)) {
        ocrScore++;
      }
    }
    
    if (ocrScore >= 2) {
      if (log) log.appendRow([timestamp, `âœ… OCR ××™×©×¨ - × ×ž×¦××• ${ocrScore} ×¡×™×ž× ×™×`]);
      return { decision: 'APPROVE', reason: `OCR ×ž×¦× ${ocrScore} ×¡×™×ž× ×™× ×—×™×•×‘×™×™×` };
    } else {
      if (log) log.appendRow([timestamp, `âŒ OCR ×“×—×” - ×¨×§ ${ocrScore} ×¡×™×ž× ×™×`]);
      return { decision: 'REJECT', reason: `OCR ×ž×¦× ×¨×§ ${ocrScore} ×¡×™×ž× ×™×` };
    }
    
  } catch (e) {
    if (log) log.appendRow([timestamp, `âš ï¸ OCR × ×›×©×œ: ${e.message}`]);
    // ×× OCR × ×›×©×œ, × ××©×¨ ×‘×–×”×™×¨×•×ª
    return { decision: 'APPROVE', reason: 'OCR × ×›×©×œ - ××™×©×•×¨ ×–×”×™×¨' };
  }
}

// === ×¤×•× ×§×¦×™×” ×ž×¨×›×–×™×ª ×ž×©×•×¤×¨×ª ===
function extractInvoiceDataWithOCR(message, pdfAttachment) {
  const timestamp = new Date().toLocaleString('he-IL');
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  
  try {
    // ×©×œ×‘ 1: ×‘×“×™×§×” ×—×›×ž×” ×ž×”×™×¨×”
    const smartResult = smartInvoiceValidation(message, pdfAttachment);
    
    if (smartResult.decision === 'REJECT') {
      if (log) log.appendRow([timestamp, `âŒ × ×“×—×”: ${smartResult.reason}`]);
      return null;
    }
    
    if (smartResult.decision === 'APPROVE') {
      if (log) log.appendRow([timestamp, `âœ… ××•×©×¨: ${smartResult.reason}`]);
      return {
        companyName: "",
        invoiceNumber: extractInvoiceNumberFromFilename(pdfAttachment.filename),
        amount: "",
        currency: "",
        attachmentName: pdfAttachment.filename,
        isValidInvoice: true,
        validationMethod: `×—×›×: ${smartResult.reason}`,
        confidence: smartResult.confidence
      };
    }
    
    // ×©×œ×‘ 2: OCR ×œ×ž×§×¨×™× ×ž×¤×•×§×¤×§×™×
    if (smartResult.decision === 'OCR_NEEDED') {
      const ocrResult = performOCRValidation(message, pdfAttachment);
      
      if (ocrResult.decision === 'APPROVE') {
        if (log) log.appendRow([timestamp, `âœ… ××•×©×¨ ×œ××—×¨ OCR: ${ocrResult.reason}`]);
        return {
          companyName: "",
          invoiceNumber: extractInvoiceNumberFromFilename(pdfAttachment.filename),
          amount: "",
          currency: "",
          attachmentName: pdfAttachment.filename,
          isValidInvoice: true,
          validationMethod: `OCR: ${ocrResult.reason}`,
          confidence: 'LOW'
        };
      } else {
        if (log) log.appendRow([timestamp, `âŒ × ×“×—×” ×œ××—×¨ OCR: ${ocrResult.reason}`]);
        return null;
      }
    }
    
    // ×œ× ××ž×•×¨ ×œ×”×’×™×¢ ×œ×›××Ÿ
    return null;
    
  } catch (e) {
    if (log) log.appendRow([timestamp, `âŒ ×©×’×™××” ×›×œ×œ×™×ª: ${e.message}`]);
    return null;
  }
}

// === ×—×™×œ×•×¥ ×ž×¡×¤×¨ ×—×©×‘×•× ×™×ª ×ž×©× ×”×§×•×‘×¥ ===
function extractInvoiceNumberFromFilename(filename) {
  // ×—×™×¤×•×© ×ž×¡×¤×¨×™× ×‘×©× ×”×§×•×‘×¥
  const patterns = [
    /invoice.*?(\d+)/i,
    /receipt.*?(\d+)/i,
    /(\d{4,})/g, // ×ž×¡×¤×¨ ×©×œ 4 ×¡×¤×¨×•×ª ××• ×™×•×ª×¨
    /(\d+)/g // ×›×œ ×ž×¡×¤×¨
  ];
  
  for (const pattern of patterns) {
    const match = filename.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  return "";
}

// ×”×¤×•× ×§×¦×™×•×ª ×”×™×©× ×•×ª ×”×•×¡×¨×• - ×¢×›×©×™×• ×ž×©×ª×ž×©×™× ×‘×‘×“×™×§×” ×¤×©×•×˜×” ×œ×¤×™ ×©× ×”×§×•×‘×¥

// === ×”×¦×’×ª ×¡×™×›×•× ×ž×ª×§×“× ===
function showAdvancedScanSummary(results) {
  const message = `âœ… ×¡×¨×™×§×” ×ž×ª×§×“×ž×ª ×”×•×©×œ×ž×”!

ðŸ“Š ×ª×•×¦××•×ª:
â€¢ ${results.added} ×—×©×‘×•× ×™×•×ª × ×•×¡×¤×•
â€¢ ${results.skipped} ×ž×™×™×œ×™× ×“×•×œ×’×•  
â€¢ ${results.total} ×ž×™×™×œ×™× × ×‘×“×§×• ×‘×¡×š ×”×›×œ

ðŸš€ ×©×™×¤×•×¨×™×:
â€¢ Gmail API ×ž×”×™×¨ ×™×•×ª×¨
â€¢ ×—×™×œ×•×¥ ×œ×™× ×§×™× ×ž×©×•×¤×¨
â€¢ ×ž×˜××“×˜×” ×ž×“×•×™×§×ª ×™×•×ª×¨

ðŸŽ‰ ×”×—×©×‘×•× ×™×•×ª ×–×ž×™× ×•×ª ×‘×’×™×œ×™×•×Ÿ ×”×—×“×©!`;

  _logMessage(message);
  SpreadsheetApp.getUi().alert(message);
}

// === SMART AMOUNT EXTRACTION ===
function extractAmountSmart(message, pdfAttachment) {
  const startTime = new Date();
  
  try {
    // ×©×œ×‘ 1: ×—×™×œ×•×¥ ×ž×”×™×¨ ×ž× ×•×©× ×•×ª×•×›×Ÿ
    const quickAmount = extractAmountFromText(message);
    if (quickAmount.found) {
      _logMessage(`ðŸ’° ×¡×›×•× × ×ž×¦× ×ž×”×™×¨: ${quickAmount.amount} ${quickAmount.currency}`);
      return quickAmount;
    }
    
    // ×©×œ×‘ 2: OCR ×¨×§ ×× ×œ× × ×ž×¦× ×‘×©×œ×‘ 1
    if (pdfAttachment) {
      _logMessage(`ðŸ” ×ž×¢×‘×¨ ×œ-OCR ×œ×—×™×œ×•×¥ ×¡×›×•× ×ž-${pdfAttachment.filename}`);
      const ocrAmount = extractAmountFromPdfOCR(message, pdfAttachment);
      if (ocrAmount.found) {
        const duration = (new Date() - startTime) / 1000;
        _logMessage(`ðŸ’° ×¡×›×•× × ×ž×¦× OCR: ${ocrAmount.amount} ${ocrAmount.currency} (${duration.toFixed(1)}s)`);
        return ocrAmount;
      }
    }
    
    return { found: false, amount: "", currency: "", method: "not_found" };
    
  } catch (e) {
    _logMessage(`âŒ ×©×’×™××” ×‘×—×™×œ×•×¥ ×¡×›×•×: ${e.message}`);
    return { found: false, amount: "", currency: "", method: "error" };
  }
}

// === ×—×™×œ×•×¥ ×ž×”×™×¨ ×ž×˜×§×¡×˜ ===
function extractAmountFromText(message) {
  try {
    const headers = message.payload.headers;
    const subject = headers.find(h => h.name.toLowerCase() === 'subject')?.value || '';
    
    // ×—×™×œ×•×¥ ×ª×•×›×Ÿ ×”×ž×™×™×œ
    let bodyText = '';
    if (message.payload.body?.data) {
      bodyText = Utilities.base64DecodeWebSafe(message.payload.body.data);
    } else if (message.payload.parts) {
      for (const part of message.payload.parts) {
        if (part.mimeType === 'text/plain' && part.body?.data) {
          bodyText += Utilities.base64DecodeWebSafe(part.body.data);
        }
      }
    }
    
    const fullText = subject + ' ' + bodyText;
    
    // Regex patterns ×œ×¡×›×•×ž×™×
    const patterns = [
      // ×©×§×œ×™×
      /â‚ª\s*([0-9,]+\.?\d*)/g,
      /([0-9,]+\.?\d*)\s*â‚ª/g,
      /([0-9,]+\.?\d*)\s*×©×§×œ/g,
      /([0-9,]+\.?\d*)\s*×©"×—/g,
      /×¡×›×•×[:\s]*([0-9,]+\.?\d*)/g,
      
      // ×“×•×œ×¨×™×
      /\$\s*([0-9,]+\.?\d*)/g,
      /([0-9,]+\.?\d*)\s*USD/g,
      /([0-9,]+\.?\d*)\s*\$/g,
      
      // ×™×•×¨×•
      /â‚¬\s*([0-9,]+\.?\d*)/g,
      /([0-9,]+\.?\d*)\s*EUR/g,
      
      // ×›×œ×œ×™
      /total[:\s]*([0-9,]+\.?\d*)/gi,
      /amount[:\s]*([0-9,]+\.?\d*)/gi
    ];
    
    for (const pattern of patterns) {
      const matches = fullText.match(pattern);
      if (matches && matches.length > 0) {
        const match = matches[0];
        const amount = match.replace(/[^\d.,]/g, '');
        const currency = detectCurrency(match);
        
        if (parseFloat(amount.replace(',', '')) > 0) {
          return {
            found: true,
            amount: amount,
            currency: currency,
            method: "text_extraction",
            source: match
          };
        }
      }
    }
    
    return { found: false, method: "text_extraction" };
    
  } catch (e) {
    _logMessage(`âš ï¸ ×©×’×™××” ×‘×—×™×œ×•×¥ ×ž×˜×§×¡×˜: ${e.message}`);
    return { found: false, method: "text_error" };
  }
}

// === ×–×™×”×•×™ ×ž×˜×‘×¢ ===
function detectCurrency(text) {
  const lowerText = text.toLowerCase();
  
  if (lowerText.includes('â‚ª') || lowerText.includes('×©×§×œ') || lowerText.includes('×©"×—')) {
    return 'ILS';
  } else if (lowerText.includes('$') || lowerText.includes('usd') || lowerText.includes('dollar')) {
    return 'USD';
  } else if (lowerText.includes('â‚¬') || lowerText.includes('eur') || lowerText.includes('euro')) {
    return 'EUR';
  }
  
  return 'ILS'; // ×‘×¨×™×¨×ª ×ž×—×“×œ
}

// === OCR ×œ×—×™×œ×•×¥ ×¡×›×•× ×ž-PDF ===
function extractAmountFromPdfOCR(message, pdfAttachment) {
  try {
    // ×›××Ÿ × ×©×ª×ž×© ×‘-OCR ×”×§×™×™× ××‘×œ × ×ª×ž×§×“ ×‘×—×™×¤×•×© ×¡×›×•×ž×™×
    // ×–×” ×™×”×™×” ××™×˜×™ ×™×•×ª×¨ ××– × ×¢×©×” ×¨×§ ×× ×œ× ×ž×¦×× ×• ×‘×˜×§×¡×˜
    
    _logMessage(`ðŸ” OCR: ×ž× ×ª×— ${pdfAttachment.filename} ×œ×—×™×¤×•×© ×¡×›×•×ž×™×`);
    
    // TODO: ×ž×™×ž×•×© OCR ×ž×ª×§×“× ×œ×—×™×œ×•×¥ ×¡×›×•×ž×™×
    // ×œ×¢×ª ×¢×ª×” × ×—×–×™×¨ ×œ× × ×ž×¦×
    
    return { found: false, method: "ocr_not_implemented" };
    
  } catch (e) {
    _logMessage(`âŒ ×©×’×™××” ×‘-OCR: ${e.message}`);
    return { found: false, method: "ocr_error" };
  }
}

// === ×™×¦×™×¨×ª ×’×™×œ×™×•×Ÿ ×ª×•×¦××•×ª ×ž×ª×§×“× ===
function createAdvancedResultsSheet(ss, tabName) {
  try {
    _logMessage(`â„¹ï¸ ðŸ“‘ ×™×•×¦×¨ ×’×™×œ×™×•×Ÿ ×—×“×©: ${tabName}`);
    
    let sheet = ss.getSheetByName(tabName);
    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet(tabName);
    }
    
    // ×™×¦×™×¨×ª ×›×•×ª×¨×•×ª ×ž×¢×•×“×›× ×•×ª
    const headers = [
      "Action", "Category", "Sender Name", "Sender Email", "Date", "Subject",
      "PDF Link", "Email ID", "Attachment Name", "Amount", "Currency", "Extract Method"
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    _logMessage("âœ… ×›×•×ª×¨×•×ª × ×•×¦×¨×• ×‘×”×¦×œ×—×”");
    
    // ×¦'×§×‘×•×§×¡×™× ×™×ª×•×•×¡×¤×• ××—×¨×™ ×›×ª×™×‘×ª ×”× ×ª×•× ×™×
    // sheet.getRange("A2:A100").insertCheckboxes(); // ×”×•×¡×¨ - ×’×•×¨× ×œ×‘×¢×™×™×ª ×©×•×¨×” 1000
    // _logMessage("âœ… ×¦'×§×‘×•×§×¡×™× × ×•×¦×¨×• ×‘×”×¦×œ×—×”");
    
    // ×¢×™×¦×•×‘ ×‘×¡×™×¡×™
    const headerRange = sheet.getRange("A1:L1");
    headerRange.setFontWeight("bold")
               .setBackground("#ddebf7")
               .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    
    // ×”×’×“×¨×ª ×¨×•×—×‘ ×¢×ž×•×“×•×ª
    sheet.setColumnWidth(1, 80);   // Action
    sheet.setColumnWidth(2, 100);  // Category
    sheet.setColumnWidth(3, 150);  // Sender Name
    sheet.setColumnWidth(4, 200);  // Sender Email
    sheet.setColumnWidth(5, 100);  // Date
    sheet.setColumnWidth(6, 300);  // Subject
    sheet.setColumnWidth(7, 250);  // PDF Link
    sheet.setColumnWidth(8, 150);  // Email ID
    sheet.setColumnWidth(9, 150);  // Attachment Name
    sheet.setColumnWidth(10, 100); // Amount
    sheet.setColumnWidth(11, 80);  // Currency
    sheet.setColumnWidth(12, 120); // Extract Method
    
    _logMessage("âœ… ×¢×™×¦×•×‘ ×’×™×œ×™×•×Ÿ ×”×•×©×œ×");
    
    return sheet;
    
  } catch (e) {
    _logMessage(`âŒ ×©×’×™××” ×‘×™×¦×™×¨×ª ×’×™×œ×™×•×Ÿ: ${e.message}`);
    return null;
  }
}
