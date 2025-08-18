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
 * Parse a Config time cell into [hour, minute] in project timezone.
 * Handles strings ("11:00", "6:30 PM"), Dates, or numeric fractions (Google Sheets TIME).
 */
function _toHM(val) {
  
  const tz = _projectTZ();

  // Case: Date object (common when cell is time-formatted)
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val)) {
    const h = Number(Utilities.formatDate(val, tz, "H"));
    const m = Number(Utilities.formatDate(val, tz, "m"));
    return [h, m];
  }

  // Case: Number (fraction of a day)
  if (typeof val === "number") {
    let total = Math.round(val * 1440);             // minutes in a day
    total = ((total % 1440) + 1440) % 1440;         // wrap safely
    return [Math.floor(total / 60), total % 60];
  }

  // Case: String "HH:mm" or "h:mm AM/PM"
  const s = String(val || "").trim();
  const m = s.match(/^(\d{1,2}):(\d{2})(?:\s*(AM|PM))?$/i);
  if (m) {
    let h = Number(m[1]), mm = Number(m[2]);
    const ap = (m[3] || "").toUpperCase();
    if (ap === "AM" && h === 12) h = 0;
    if (ap === "PM" && h < 12) h += 12;
    return [h, mm];
  }

  throw new Error(`Unrecognized time value in Config: ${val}`);
}

function _projectTZ() {
  return (typeof Session !== 'undefined' && Session.getScriptTimeZone)
    ? Session.getScriptTimeZone()
    : 'Etc/UTC';
}