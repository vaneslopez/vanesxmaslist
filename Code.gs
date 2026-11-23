// Paste this into Extensions > Apps Script in a new Google Sheet, then deploy
// it as a Web app with access set to "Anyone".
const SHEET_NAME = "claims";

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["item_id", "name", "claimed_at"]);
  }
  return sheet;
}

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const claims = Object.fromEntries(rows.filter(row => row[0]).map(row => [row[0], row[1]]));
  return ContentService.createTextOutput(JSON.stringify(claims))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const id = (e.parameter.id || "").trim();
  const name = (e.parameter.name || "").trim().slice(0, 40);
  if (!id || !name) return reply_({ ok: false });
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = sheet_();
    const ids = sheet.getDataRange().getValues().slice(1).map(row => row[0]);
    if (ids.includes(id)) return reply_({ ok: false, claimed: true });
    sheet.appendRow([id, name, new Date().toISOString()]);
    return reply_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function reply_(body) {
  return ContentService.createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}
