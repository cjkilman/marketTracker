function getOrCreateSheet(ss, name, headers) {
  if (!ss || typeof ss.getSheetByName !== "function") {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && Array.isArray(headers)) {
      sheet.appendRow(headers);
    }
  } else if (headers && Array.isArray(headers)) {
    const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const same = currentHeaders.every((h, i) => h === headers[i]);
    if (!same) {
      sheet.clearContents();
      sheet.appendRow(headers);
    }
  }
  return sheet;
}