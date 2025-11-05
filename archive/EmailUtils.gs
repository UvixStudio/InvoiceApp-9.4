// === EmailUtils.gs - פונקציות מייל ===

function _parseSender(fromString) {
  // חילוץ האימייל מהמחרוזת "Name <email@domain.com>"
  const match = fromString.match(/<([^>]+)>/);
  return match ? match[1] : fromString;
}

function extractSenderName(fromString) {
  // חילוץ השם מהמחרוזת "Name <email@domain.com>"
  const match = fromString.match(/^([^<]+)</);
  if (match) {
    return match[1].trim().replace(/"/g, '');
  }
  return _parseSender(fromString).split('@')[0];
}

function isLocalEmail(email) {
  return email.includes('.co.il') || email.includes('.org.il') || email.includes('.gov.il');
}