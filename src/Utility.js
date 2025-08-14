function getOrCreateSheet(ss, name, headers) {
  // Default to active spreadsheet if none provided
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  // Convert ID strings into Spreadsheet objects
  if (typeof ss === 'string') {
    ss = SpreadsheetApp.openById(ss);
  }

  // Validate object type
  if (typeof ss.getSheetByName !== 'function') {
    throw new Error('getOrCreateSheet: Provided ss is not a Spreadsheet.');
  }

  // Proceed with previous logic
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && Array.isArray(headers)) sheet.appendRow(headers);
  } else if (headers && Array.isArray(headers)) {
    const currentHeaders = sheet
      .getRange(1, 1, 1, headers.length)
      .getValues()[0];
    const same = currentHeaders.every((h, i) => h === headers[i]);
    if (!same) {
      sheet.clearContents();
      sheet.appendRow(headers);
    }
  }

  return sheet;
}