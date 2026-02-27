const SHEET_NAME = "Form Data";
const SPREADSHEET_ID = "1uq24US_wIAVSwyMy3SFi6tqaHFVls5C9qySgD4fogtE"; // Bound or standalone safe.
const DEBUG_SHEET_NAME = "Debug Logs";

function doPost(e) {
  const receivedAt = new Date().toISOString();
  try {
    const raw = (e && e.postData && e.postData.contents) || "{}";
    const data = JSON.parse(raw);
    const sheet = getSheet_();

    const headers = getHeaders_(data);
    ensureHeaders_(sheet, headers);

    const row = headers.map((h) => data[h] ?? "");
    sheet.appendRow(row);
    writeDebugLog_(receivedAt, "OK", "Saved to Google Sheet", raw);

    return jsonResponse_({ ok: true, message: "Saved to Google Sheet" });
  } catch (err) {
    const msg = "Error: " + (err && err.message ? err.message : String(err));
    const raw = (e && e.postData && e.postData.contents) || "";
    writeDebugLog_(receivedAt, "ERROR", msg, raw);
    return jsonResponse_({
      ok: false,
      message: msg
    });
  }
}

function doGet() {
  return jsonResponse_({ ok: true, message: "Web app is running" });
}

function getSheet_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  // For standalone Apps Script projects, active spreadsheet can be null.
  if (!ss && SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  if (!ss) {
    throw new Error("Spreadsheet not found. Use Extensions > Apps Script from the sheet OR set SPREADSHEET_ID.");
  }

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

function getHeaders_(data) {
  const preferred = [
    "Date",
    "Day",
    "Opening Balance",
    "Closing Balance",
    "Sleep",
    "Exercise",
    "Medication",
    "Session",
    "Mood (1-5)",
    "Energy (1-5)",
    "Trade No",
    "Script",
    "Reason for Entry",
    "Reason to Exit",
    "Averaging",
    "P/L",
    "Notes",
    "Overtrading",
    "Fixed Position Size",
    "ClientTimestamp"
  ];

  const extra = Object.keys(data).filter((k) => !preferred.includes(k));
  return preferred.concat(extra);
}

function ensureHeaders_(sheet, headers) {
  const current = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    : [];

  const isSame =
    current.length === headers.length &&
    current.every((v, i) => String(v) === String(headers[i]));

  if (!isSame) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function writeDebugLog_(ts, status, message, rawPayload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(DEBUG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(DEBUG_SHEET_NAME);
    sh.getRange(1, 1, 1, 4).setValues([["Timestamp", "Status", "Message", "Payload"]]);
    sh.setFrozenRows(1);
  }

  sh.appendRow([ts, status, message, String(rawPayload || "").slice(0, 50000)]);
}
