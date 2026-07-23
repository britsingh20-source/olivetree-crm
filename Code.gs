/**
 * Olivetree Interiors — Lead CRM backend
 * Deploy: Extensions > Apps Script > paste this > Deploy > New deployment >
 * Web app · Execute as: Me · Who has access: Anyone > copy the /exec URL.
 */
const SHEET_NAME = "Leads";
const HEADERS = ["id","dateAdded","name","phone","source","service","budget","stage","nextFollowUp","lastContact","notes"];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }
  return sh;
}

function doGet(e) {
  const sh = getSheet_();
  const rows = sh.getDataRange().getValues();
  const leads = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    const o = {};
    HEADERS.forEach((h, j) => {
      let v = rows[i][j];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
      o[h] = v;
    });
    leads.push(o);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, leads }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sh = getSheet_();
    const data = sh.getDataRange().getValues();
    const id = body.lead && body.lead.id;
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { rowIndex = i + 1; break; }
    }
    if (body.action === "add" || (body.action === "update" && rowIndex === -1)) {
      sh.appendRow(HEADERS.map(h => body.lead[h] != null ? body.lead[h] : ""));
    } else if (body.action === "update") {
      sh.getRange(rowIndex, 1, 1, HEADERS.length)
        .setValues([HEADERS.map(h => body.lead[h] != null ? body.lead[h] : "")]);
    } else if (body.action === "delete" && rowIndex > -1) {
      sh.deleteRow(rowIndex);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
