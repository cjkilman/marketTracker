// =========================================================
// TODO [Bridge Tag]:
// Add debugLog() + warnIfMismatch() helpers here if/when
// array length mismatches or noisy debugging become a pain.
// =========================================================

/**
 * Get or create a sheet, preserving headers.
 * For new sheets, limits the column count to the header length.
 * @param {SpreadsheetApp.Spreadsheet} ss - Spreadsheet object
 * @param {string} name - Sheet name
 * @param {string[]} headers - Array of header strings
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(ss, name, headers) {
  if (!ss || typeof ss.getSheetByName !== 'function') {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!Array.isArray(headers)) {
    throw new Error("getOrCreateSheet: headers must be an array of strings");
  }

  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    // Create new sheet
    sheet = ss.insertSheet(name);

    // Adjust columns to match headers exactly
    const headerCount = headers.length;
    const maxCols = sheet.getMaxColumns();
    if (maxCols > headerCount) {
      sheet.deleteColumns(headerCount + 1, maxCols - headerCount);
    } else if (maxCols < headerCount) {
      sheet.insertColumnsAfter(maxCols, headerCount - maxCols);
    }

    sheet.appendRow(headers);
  } else {
    // Existing sheet: check headers
    const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const same = currentHeaders.every((h, i) => h === headers[i]);
    if (!same) {
      sheet.clearContents();
      sheet.appendRow(headers);
    }
  }

  return sheet;
}

/**
 * Basic logging utility for informational messages.
 * @param {string} message - The message to log.
 */
function logInfo(message) {
  Logger.log("[INFO] " + message);
}