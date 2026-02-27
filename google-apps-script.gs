const SHEET_NAME = "Form Data";

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const sheet = getSheet_();

    const headers = getHeaders_(data);
    ensureHeaders_(sheet, headers);

    const row = headers.map((h) => data[h] ?? "");
    sheet.appendRow(row);

    return jsonResponse_({ ok: true, message: "Saved to Google Sheet" });
  } catch (err) {
    return jsonResponse_({
      ok: false,
      message: "Error: " + (err && err.message ? err.message : String(err))
    });
  }
}

function doGet() {
  return jsonResponse_({ ok: true, message: "Web app is running" });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
