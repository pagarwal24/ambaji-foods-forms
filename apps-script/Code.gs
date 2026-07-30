const SPREADSHEET_ID = "1aaAciTT0-EQKHJU1-WSbumRK-7v5CfVor55agUHvEFE";
const DRIVE_FOLDER_ID = "1qw3s0QLr6eM190eBcCjQFlcGwmQ5AgYG";
const DATA_SHEET = "Form Records";

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.a) {
    try {
      const decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(String(params.a))).getDataAsString();
      const approval = JSON.parse(decoded);
      return handleApproval_(
        String(approval.id || ""),
        String(approval.role || ""),
        String(approval.decision || ""),
        String(approval.token || "")
      );
    } catch (error) {
      return handleApproval_("", "", "", "");
    }
  }
  const action = String(params.action || "list");
  if (action === "approval") {
    return handleApproval_(
      String(e.parameter.id || ""),
      String(e.parameter.role || ""),
      String(e.parameter.decision || ""),
      String(e.parameter.token || "")
    );
  }
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
  if (action === "approvalCode") {
    const decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(String(params.code || ""))).getDataAsString();
    const approval = JSON.parse(decoded);
    recordApproval_(
      String(approval.id || ""),
      String(approval.role || ""),
      String(approval.decision || ""),
      String(approval.token || "")
    );
  }
  return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
}

function sheet_() {
  const book = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = book.getSheetByName(DATA_SHEET);
  if (!sheet) {
    sheet = book.insertSheet(DATA_SHEET);
    sheet.appendRow(["Record ID", "Form ID", "Document ID", "Form Data", "Created At", "Updated At", "Image URL", "PDF URL", "Approval Status"]);
    sheet.setFrozenRows(1);
  }
  if (sheet.getMaxColumns() < 9) sheet.insertColumnsAfter(sheet.getMaxColumns(), 9 - sheet.getMaxColumns());
  sheet.getRange(1, 1, 1, 9).setValues([["Record ID", "Form ID", "Document ID", "Form Data", "Created At", "Updated At", "Image URL", "PDF URL", "Approval Status"]]);
  sheet.getRange(1, 1, 1, 9).setBackground("#174B3A").setFontColor("#FFFFFF").setFontWeight("bold");
  return sheet;
}

function listRecords_() {
  const sheet = sheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 9).getValues().filter(row => row[0]).map(row => ({
    id: String(row[0]),
    formId: String(row[1]),
    docId: String(row[2]),
    data: JSON.parse(String(row[3] || "{}")),
    createdAt: String(row[4] || ""),
    updatedAt: String(row[5] || ""),
    imageUrl: String(row[6] || ""),
    pdfUrl: String(row[7] || ""),
    approvalStatus: String(row[8] || "pending")
  }));
}

function upsertRecord_(record) {
  if (!record || !record.id) throw new Error("Record ID is required.");
  record.data = record.data || {};
  record.data.approvalStatus = String(record.data.approvalStatus || "pending").toLowerCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let targetRow;
  let existingPdf = "";
  try {
    const sheet = sheet_();
    const lastRow = sheet.getLastRow();
    targetRow = lastRow + 1;
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
      const index = ids.indexOf(String(record.id));
      if (index >= 0) targetRow = index + 2;
    }
    const existingImage = targetRow <= lastRow ? sheet.getRange(targetRow, 7).getValue() : "";
    existingPdf = targetRow <= lastRow ? sheet.getRange(targetRow, 8).getValue() : "";
    sheet.getRange(targetRow, 1, 1, 9).setValues([[
      record.id,
      record.formId || "",
      record.docId || "",
      JSON.stringify(record.data),
      record.createdAt || new Date().toISOString(),
      record.updatedAt || new Date().toISOString(),
      existingImage,
      existingPdf,
      record.data.approvalStatus
    ]]);
  } finally {
    lock.releaseLock();
  }
  const pdfUrl = createRecordPdf_(record, record.data.approvalStatus);
  const sheet = sheet_();
  sheet.getRange(targetRow, 8).setValue(pdfUrl);
  sheet.getRange(targetRow, 9).setValue(record.data.approvalStatus);
  if (existingPdf && existingPdf !== pdfUrl) trashFileByUrl_(existingPdf);
}

function statusLabel_(status) {
  const value = String(status || "pending").toLowerCase();
  if (value === "approved") return "Approved";
  if (value === "rejected") return "Rejected";
  return "Pending";
}

function safeFilePart_(value) {
  return String(value || "Form").replace(/[\\/:*?"<>|#%{}[\]]/g, "_").replace(/\s+/g, " ").trim().slice(0, 90);
}

function childFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function recordFolder_(status, formId) {
  const root = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const statusFolderName = String(status || "").toLowerCase() === "rejected" ? "Reject" : statusLabel_(status);
  const statusFolder = childFolder_(root, statusFolderName);
  return childFolder_(statusFolder, safeFilePart_(formId || "General Forms"));
}

function displayValue_(value) {
  if (Array.isArray(value)) return value.filter(item => item !== "" && item != null).join(", ");
  if (value && typeof value === "object") return "";
  return String(value == null || value === "" ? "—" : value);
}

function createRecordPdf_(record, status) {
  const statusText = statusLabel_(status);
  const formName = safeFilePart_(record.formId || "Ambaji Foods Form");
  const docNumber = String(record.docId || "—");
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "dd-MMM-yyyy HH:mm:ss");
  const fileName = statusText.toUpperCase() + "_" + formName + "_" + safeFilePart_(record.id) + ".pdf";
  const document = DocumentApp.create(fileName.replace(/\.pdf$/i, ""));
  const body = document.getBody();
  body.clear();
  const company = body.appendParagraph("Ambaji Foods (India) Private Limited, Kanpur");
  company.setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const formTitle = body.appendParagraph(formName);
  formTitle.setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const statusParagraph = body.appendParagraph("STATUS: " + statusText.toUpperCase());
  statusParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  statusParagraph.editAsText().setBold(true).setFontSize(18)
    .setForegroundColor(statusText === "Approved" ? "#168A45" : statusText === "Rejected" ? "#A1261D" : "#9A6500");
  body.appendParagraph("Document ID: " + docNumber);
  body.appendParagraph("Record ID: " + String(record.id || "—"));
  body.appendParagraph("Generated: " + timestamp);
  body.appendHorizontalRule();
  const rows = [["Field", "Finding / Value"]];
  const internalKeys = {approvals:true, approvalStatus:true, submittedAt:true, rejectedAt:true};
  Object.keys(record.data || {}).forEach(key => {
    if (internalKeys[key]) return;
    const value = displayValue_(record.data[key]);
    if (value) rows.push([String(key).replace(/([A-Z])/g, " $1").replace(/^./, char => char.toUpperCase()), value]);
  });
  const table = body.appendTable(rows);
  table.getRow(0).getCell(0).editAsText().setBold(true);
  table.getRow(0).getCell(1).editAsText().setBold(true);
  body.appendHorizontalRule();
  const approvals = (record.data || {}).approvals || {};
  body.appendParagraph("Production approval: " + statusLabel_((approvals.production || {}).status));
  body.appendParagraph("Parle approval: " + statusLabel_((approvals.parle || {}).status));
  const footer = body.appendParagraph("Apps By Prateek Agarwal");
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setBold(true);
  document.saveAndClose();
  const source = DriveApp.getFileById(document.getId());
  const pdf = recordFolder_(status, formName).createFile(source.getBlob().getAs(MimeType.PDF).setName(fileName));
  source.setTrashed(true);
  return pdf.getUrl();
}

function trashFileByUrl_(url) {
  const match = String(url || "").match(/[-\w]{25,}/);
  if (!match) return;
  try { DriveApp.getFileById(match[0]).setTrashed(true); } catch (error) {}
}

function notifyApprovalTeam_(record) {
  const data = record.data || {};
  const approvals = data.approvals || {};
  const serviceUrl = ScriptApp.getService().getUrl();
  const subject = "Approval required: Ambaji Foods Line Weight Report";
  const summary = buildLineWeightSummary_(data);
  ["production", "parle"].forEach(role => {
    const person = approvals[role] || {};
    if (!person.email || !person.token) return;
    const roleLabel = role === "production" ? "Production Manager" : "Parle Executive";
    const base = serviceUrl + "?action=approval&id=" + encodeURIComponent(record.id) +
      "&role=" + encodeURIComponent(role) + "&token=" + encodeURIComponent(person.token);
    const approveUrl = base + "&decision=approved";
    const rejectUrl = base + "&decision=rejected";
    MailApp.sendEmail({
      to: person.email,
      subject: subject,
      htmlBody:
        "<div style='font-family:Arial,sans-serif;max-width:720px;margin:auto'>" +
        "<h2>Ambaji Foods Line Weight Report</h2>" +
        "<p>Dear " + escapeHtml_(person.name || roleLabel) + ",</p>" +
        "<p>A line-weight report requires your review as <b>" + roleLabel + "</b>.</p>" +
        summary +
        "<p style='margin-top:24px'>" +
        "<a href='" + approveUrl + "' style='background:#168a45;color:#fff;padding:12px 18px;text-decoration:none;border-radius:8px;margin-right:10px'>Approve</a>" +
        "<a href='" + rejectUrl + "' style='background:#a1261d;color:#fff;padding:12px 18px;text-decoration:none;border-radius:8px'>Reject</a>" +
        "</p><p style='color:#666;font-size:12px'>Document ID: AFIPL/PRD/FRM/20</p></div>"
    });
  });
  const filler = approvals.filler || {};
  if (filler.email) {
    MailApp.sendEmail({
      to: filler.email,
      subject: "Submitted: Ambaji Foods Line Weight Report",
      htmlBody:
        "<div style='font-family:Arial,sans-serif;max-width:720px;margin:auto'>" +
        "<h2>Line Weight Report Submitted</h2><p>Your report has been saved to Google Drive and sent for approval.</p>" +
        summary + "<p style='color:#666;font-size:12px'>Document ID: AFIPL/PRD/FRM/20</p></div>"
    });
  }
}

function buildLineWeightSummary_(data) {
  function filled(values) {
    return (values || []).filter(value => value !== "" && value !== null && value !== undefined);
  }
  const oven = filled(data.ovenWeights);
  const stacker = filled(data.stackerWeights);
  return "<table style='border-collapse:collapse;width:100%;margin-top:14px'>" +
    summaryRow_("Date", data.ovenDate || data.stackerDate || "—") +
    summaryRow_("Shift", data.ovenShift || data.stackerShift || "—") +
    summaryRow_("Oven variety", data.ovenVariety || "—") +
    summaryRow_("Oven readings", oven.length) +
    summaryRow_("Stacker variety", data.stackerVariety || "—") +
    summaryRow_("Stacker readings", stacker.length) +
    summaryRow_("Document filler", data.fillerName || "—") +
    "</table>";
}

function summaryRow_(label, value) {
  return "<tr><td style='border:1px solid #bbb;padding:8px;font-weight:bold;background:#f2f2f2'>" +
    escapeHtml_(label) + "</td><td style='border:1px solid #bbb;padding:8px'>" +
    escapeHtml_(value) + "</td></tr>";
}

function handleApproval_(id, role, decision, token) {
  let title = "Approval link invalid";
  let message = "This approval request could not be verified.";
  const success = recordApproval_(id, role, decision, token);
  if (success) {
    title = decision === "approved" ? "Report approved" : "Report rejected";
    message = "Your " + decision + " decision has been recorded with the date and time.";
  }
  return HtmlService.createHtmlOutput(
    "<!doctype html><html><head><meta name='viewport' content='width=device-width,initial-scale=1'>" +
    "<title>" + escapeHtml_(title) + "</title></head><body style='margin:0;background:#f5f5f5;font-family:Arial,sans-serif'>" +
    "<div style='max-width:560px;margin:70px auto;background:#fff;border:2px solid #000;border-radius:20px;padding:30px;text-align:center'>" +
    "<div style='font-size:48px'>" + (success ? "✓" : "!") + "</div><h1>" + escapeHtml_(title) + "</h1>" +
    "<p>" + escapeHtml_(message) + "</p><p><b>Ambaji Foods · AFIPL/PRD/FRM/20</b></p>" +
    "<p style='font-size:12px;color:#666'>Apps By Prateek Agarwal</p></div></body></html>"
  );
}

function recordApproval_(id, role, decision, token) {
  const allowedRoles = ["production", "parle"];
  const allowedDecisions = ["approved", "rejected"];
  if (id && allowedRoles.indexOf(role) >= 0 && allowedDecisions.indexOf(decision) >= 0 && token) {
    const sheet = sheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
      const index = ids.indexOf(id);
      if (index >= 0) {
        const row = index + 2;
        const data = JSON.parse(String(sheet.getRange(row, 4).getValue() || "{}"));
        const approvals = data.approvals || {};
        const approval = approvals[role] || {};
        if (String(approval.token || "") === token) {
          approval.status = decision;
          approval.signedAt = new Date().toISOString();
          approvals[role] = approval;
          data.approvals = approvals;
          const productionStatus = String((approvals.production || {}).status || "pending");
          const parleStatus = String((approvals.parle || {}).status || "pending");
          if (productionStatus === "rejected" || parleStatus === "rejected") {
            data.approvalStatus = "rejected";
            data.rejectedAt = new Date().toISOString();
          } else if (productionStatus === "approved" && parleStatus === "approved") {
            data.approvalStatus = "approved";
            data.submittedAt = new Date().toISOString();
          } else {
            data.approvalStatus = "pending";
          }
          sheet.getRange(row, 4).setValue(JSON.stringify(data));
          sheet.getRange(row, 6).setValue(new Date().toISOString());
          sheet.getRange(row, 9).setValue(data.approvalStatus);
          if (data.approvalStatus === "approved" || data.approvalStatus === "rejected") {
            const oldPdf = String(sheet.getRange(row, 8).getValue() || "");
            const finalRecord = {
              id: String(sheet.getRange(row, 1).getValue() || id),
              formId: String(sheet.getRange(row, 2).getValue() || "form"),
              docId: String(sheet.getRange(row, 3).getValue() || ""),
              data: data,
              createdAt: String(sheet.getRange(row, 5).getValue() || ""),
              updatedAt: new Date().toISOString()
            };
            const finalPdf = createRecordPdf_(finalRecord, data.approvalStatus);
            sheet.getRange(row, 8).setValue(finalPdf);
            if (oldPdf && oldPdf !== finalPdf) trashFileByUrl_(oldPdf);
          }
          return true;
        }
      }
    }
  }
  return false;
}

function escapeHtml_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
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
  let status = "pending";
  let formId = "General Forms";
  let row = 0;
  const id = String(params.recordId || "");
  if (id) {
    const sheet = sheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
      const index = ids.indexOf(id);
      if (index >= 0) {
        row = index + 2;
        formId = String(sheet.getRange(row, 2).getValue() || formId);
        const data = JSON.parse(String(sheet.getRange(row, 4).getValue() || "{}"));
        status = String(data.approvalStatus || "pending");
      }
    }
  }
  const folder = recordFolder_(status, formId);
  const bytes = Utilities.base64Decode(params.base64);
  const stampedName = statusLabel_(status).toUpperCase() + "_" + safeFilePart_(params.fileName);
  const blob = Utilities.newBlob(bytes, params.mimeType || "image/png", stampedName);
  const file = folder.createFile(blob);
  if (row) sheet_().getRange(row, 7).setValue(file.getUrl());
  return file.getUrl();
}
