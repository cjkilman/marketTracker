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

/************************************************************
 * Time helpers — consistent, defensive, and log-friendly
 * Contract:
 *   _toHM(val) -> { h:number, m:number }
 *   _inWindow_(now, h, m, durationMin) -> boolean  (LOCAL tz)
 ************************************************************/

function _projectTZ() {
  return (typeof Session !== 'undefined' && Session.getScriptTimeZone)
    ? Session.getScriptTimeZone()
    : 'Etc/UTC';
}

/** Parse time-of-day from Date | number (sheet fraction) | string.
 * Accepts:
 *   Date                → local H:m
 *   number (0..1 day)  → H:m
 *   "11"               → 11:00
 *   "11:0" / "11:00"   → 11:00
 *   "6:00 PM" / "6 PM" → 18:00
 */
function _toHM(val) {
  const tz = (typeof _projectTZ === "function" ? _projectTZ() : Session.getScriptTimeZone());

  // Date object (e.g., time-formatted cell)
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val)) {
    const h = Number(Utilities.formatDate(val, tz, "H"));
    const m = Number(Utilities.formatDate(val, tz, "m"));
    _assertFinite(h, m, `Bad Date in Config: ${val}`);
    return { h, m };
  }

  // Number (fraction of day)
  if (typeof val === "number") {
    let total = Math.round(val * 1440);                       // minutes
    total = ((total % 1440) + 1440) % 1440;                   // wrap
    return { h: Math.floor(total / 60), m: total % 60 };
  }

  // String family
  const s = String(val ?? "").trim().toUpperCase();
  if (!s) throw new Error(`Time missing in Config`);

  // Accept "H", "H:M", "H AM/PM", "H:M AM/PM"
  const m = s.match(/^(\d{1,2})(?::(\d{1,2}))?\s*(AM|PM)?$/);
  if (!m) throw new Error(`Unrecognized time value in Config: "${val}"`);

  let hNum = parseInt(m[1], 10);
  let mNum = (m[2] != null ? parseInt(m[2], 10) : 0);
  const ap  = m[3]; // AM/PM or undefined

  if (ap) {
    if (ap === "PM" && hNum < 12) hNum += 12;
    if (ap === "AM" && hNum === 12) hNum = 0;
  }
  _assertHM(hNum, mNum, `Invalid time from "${val}" → (${hNum},${mNum})`);
  return { h: hNum, m: mNum };
}

function _assertFinite(h, m, msg) {
  if (!Number.isFinite(h) || !Number.isFinite(m)) throw new Error(msg);
}
function _assertHM(h, m, msg) {
  if (!(h >= 0 && h < 24) || !(m >= 0 && m < 60)) throw new Error(msg);
}

/** Local-tz window check with strict argument validation. */
function _inWindow_(now, startH, startM, durationMin) {
  if (!(now instanceof Date) || isNaN(now)) {
    throw new Error(`_inWindow_: "now" must be a valid Date, got ${now}`);
  }
  if (!Number.isInteger(startH) || !Number.isInteger(startM)) {
    throw new Error(`_inWindow_: startH/startM must be ints, got h=${startH} m=${startM}`);
  }
  if (!Number.isInteger(durationMin) || durationMin <= 0) {
    throw new Error(`_inWindow_: durationMin must be a positive int, got ${durationMin}`);
  }

  const start = new Date(now);
  start.setHours(startH, startM, 0, 0); // LOCAL tz
  const end = new Date(start.getTime() + durationMin * 60 * 1000);
  return now >= start && now < end;     // inclusive start, exclusive end
}

/** Decide phase based on config+mode at a given moment. */
function _determinePhase(config, mode, now) {
  if (mode === "open")  return { isOpenRun: true,  isCloseRun: false, allowed: true };
  if (mode === "close") return { isOpenRun: false, isCloseRun: true,  allowed: true };

  now = now || new Date();
  const DUR = 60; // minutes

  const { h: oH, m: oM } = _toHM(config?.OpenTime  ?? "11:00"); // LOCAL times
  const { h: cH, m: cM } = _toHM(config?.CloseTime ?? "18:00");

  const inOpen  = _inWindow_(now, oH, oM, DUR);
  const inClose = _inWindow_(now, cH, cM, DUR);

  // Debug trace: one glance tells you parse + window state.
  console.log(
    `[PHASE DEBUG] tz=${Session.getScriptTimeZone()} now=${now.toLocaleString()} `
    + `OpenRaw="${config?.OpenTime}"→${oH}:${String(oM).padStart(2,"0")} inOpen=${inOpen} `
    + `CloseRaw="${config?.CloseTime}"→${cH}:${String(cM).padStart(2,"0")} inClose=${inClose}`
  );

  if (inOpen)  return { isOpenRun: true,  isCloseRun: false, allowed: true };
  if (inClose) return { isOpenRun: false, isCloseRun: true,  allowed: true };
  return { isOpenRun: false, isCloseRun: false, allowed: false };
}