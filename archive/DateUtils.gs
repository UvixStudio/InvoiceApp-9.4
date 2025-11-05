// === DateUtils.gs - פונקציות תאריכים ===

function formatDateInput(date) {
  if (!date || !(date instanceof Date)) return '';
  return date.toLocaleDateString('he-IL');
}

function formatDateGmail(date) {
  if (!date || !(date instanceof Date)) return '';
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

function formatDateShort(date) {
  if (!date || !(date instanceof Date)) return '';
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = String(date.getFullYear()).slice(-2);
  return `${day}.${month}.${year}`;
}