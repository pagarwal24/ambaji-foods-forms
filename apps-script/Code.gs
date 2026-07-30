const SPREADSHEET_ID = "1aaAciTT0-EQKHJU1-WSbumRK-7v5CfVor55agUHvEFE";
const DRIVE_FOLDER_ID = "1qw3s0QLr6eM190eBcCjQFlcGwmQ5AgYG";
const DATA_SHEET = "Form Records";

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || "list");
  const callback = String((e && e.parameter && e.parameter.callback) || "callback").replace(/[^\w$.]/g, "");
  const data = action === "list" ? listRecords_() : [];
  return ContentService
    .createTextOutput(callback + "(" + JSON.stringify(data) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || "");
  if (action === "upsert") upsertRecord_(JSON.parse(params.record || "{}"));
  if (action === "delete") deleteRecord_(String(params.id || ""));
  if (action === "upload") uploadImage_(params);
  return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
}

function sheet_() {
  const book = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = book.getSheetByName(DATA_SHEET);
  if (!sheet) {
    sheet = book.insertSheet(DATA_SHEET);
    sheet.appendRow(["Record ID", "Form ID", "Document ID", "Form Data", "Created At", "Updated At", "Image URL"]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 7).setBackground("#174B3A").setFontColor("#FFFFFF").setFontWeight("bold");
  }
  return sheet;
}

function listRecords_() {
  const sheet = sheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 7).getValues().filter(row => row[0]).map(row => ({
    id: String(row[0]),
    formId: String(row[1]),
    docId: String(row[2]),
    data: JSON.parse(String(row[3] || "{}")),
    createdAt: String(row[4] || ""),
    updatedAt: String(row[5] || ""),
    imageUrl: String(row[6] || "")
  }));
}

function upsertRecord_(record) {
  if (!record || !record.id) throw new Error("Record ID is required.");
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = sheet_();
    const lastRow = sheet.getLastRow();
    let targetRow = lastRow + 1;
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
      const index = ids.indexOf(String(record.id));
      if (index >= 0) targetRow = index + 2;
    }
    const existingImage = targetRow <= lastRow ? sheet.getRange(targetRow, 7).getValue() : "";
    sheet.getRange(targetRow, 1, 1, 7).setValues([[
      record.id,
      record.formId || "",
      record.docId || "",
      JSON.stringify(record.data || {}),
      record.createdAt || new Date().toISOString(),
      record.updatedAt || new Date().toISOString(),
      existingImage
    ]]);
  } finally {
    lock.releaseLock();
  }
}

function deleteRecord_(id) {
  if (!id) return;
  const sheet = sheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
  const index = ids.indexOf(id);
  if (index >= 0) sheet.deleteRow(index + 2);
}

function uploadImage_(params) {
  if (!params.base64 || !params.fileName) throw new Error("Image data is required.");
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const bytes = Utilities.base64Decode(params.base64);
  const blob = Utilities.newBlob(bytes, params.mimeType || "image/png", params.fileName);
  const file = folder.createFile(blob);
  const id = String(params.recordId || "");
  if (id) {
    const sheet = sheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
      const index = ids.indexOf(id);
      if (index >= 0) sheet.getRange(index + 2, 7).setValue(file.getUrl());
    }
  }
  return file.getUrl();
}
