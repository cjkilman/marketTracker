function getCurrentMarketPrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const maxLogIDs = config["MaxLogIDs"] ? parseInt(config["MaxLogIDs"], 10) : 750;

  // Read the typeIDs list once
  const typeIDs = getTypeIDsFromItemList(maxLogIDs);

  // Read market settings once
  const marketCombos = getMarketSettings();

  // Ensure sheet exists and preserve headers
  const sheetName = "Market Prices";
  const headers = ["date", "market_id", "market_type", "type_id", "min_sell", "max_buy", "median_sell", "median_buy"];
  const sheet = getOrCreateSheet(ss, sheetName, headers);

  // Collect rows
  const rows = [];
  marketCombos.forEach(({ market_id, market_type }) => {
    const prices = getMarketPrices(typeIDs, market_id, market_type);

    typeIDs.forEach(type_id => {
      const entry = prices[type_id];
      rows.push([
        new Date(),
        market_id,
        market_type,
        type_id,
        entry.minSell,
        entry.maxBuy,
        entry.medianSell,
        entry.medianBuy
      ]);
    });
  });

  // Append under headers
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Prune old rows (configurable retention days)
  pruneOldRows(sheet, config["PriceRetentionDays"]);

  // NOTE: HistoryManager/updateHistory will also use pruneOldRows()
  // with config["HistoryRetentionDays"] — hook that in when we build it.
}

function testPrune()
{
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Ensure sheet exists and preserve headers
  const sheetName = "Market Prices";
  const headers = ["date", "market_id", "market_type", "type_id", "min_sell", "max_buy", "median_sell", "median_buy"];
  const sheet = getOrCreateSheet(ss, sheetName, headers);
  pruneOldRows(sheet,1)
}
/**
 * Prunes rows older than `retentionDays`, based on the given date column.
 * Defaults to column 1 if not specified.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} retentionDays
 * @param {number} [dateCol=1] - 1-based index of the date column
 */
function pruneOldRows(sheet, retentionDays, dateCol = 1) {
  if (!retentionDays) return;

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - retentionDays);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // nothing but headers

  const timestamps = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues().flat();

  for (let i = timestamps.length - 1; i >= 0; i--) {
    const ts = timestamps[i];
    var meep = ts instanceof Date && ts < cutoff;
    if (ts instanceof Date && ts < cutoff) {
      sheet.deleteRow(i + 2); // offset for header row
    }
  }
}


// Sanitizer helper
function sanitizeIDs(ids) {
  return [...new Set(
    ids
      .map(v => Number(v))
      .filter(v => !isNaN(v) && v > 0)
  )];
}


function getMarketSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Market Settings");
  if (!sheet) throw new Error("Market Settings sheet not found.");

  const types = ["station", "system", "region"];
  const combos = [];

  // Loop each column separately
  types.forEach((type, index) => {
    const col = 4 + index; // D, E, F
    const raw = sheet.getRange(3, col, sheet.getLastRow() - 2, 1)
      .getValues()
      .flat();

    const clean = sanitizeIDs(raw);
    clean.forEach(id => combos.push({ market_id: id, market_type: type }));
  });

  return combos;
}


function getTypeIDsFromItemList(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Item List Back End");
  if (!sheet) throw new Error("Item List Back End sheet not found.");

  let ids = sheet.getRange("A2:A" + sheet.getLastRow())
    .getValues()
    .flat();

  ids = sanitizeIDs(ids);

  if (limit && ids.length > limit) ids = ids.slice(0, limit);
  return ids;
}


function getAllMarketAssignments(limit) {
  const typeIDs = getTypeIDsFromItemList(limit);
  const marketCombos = getMarketSettings();

  return marketCombos.map(({ market_id, market_type }) => ({
    market_id,
    market_type,
    type_ids: typeIDs
  }));
}


/**
 * Fetches market prices and ensures numeric values are math-friendly.
 * Non-numeric values become null instead of "".
 */
function getMarketPrices(type_ids, market_id, market_type) {
  const data = postFetch(type_ids, market_id, market_type); // Using existing Fuz API function
  const result = {};

  type_ids.forEach(id => {
    const entry = data[id] || {};

    const minSell    = parseFloat(entry.sell?.min)    > 0 ? parseFloat(entry.sell.min)    : null;
    const maxBuy     = parseFloat(entry.buy?.max)     > 0 ? parseFloat(entry.buy.max)     : null;
    const medianSell = parseFloat(entry.sell?.median) > 0 ? parseFloat(entry.sell.median) : null;
    const medianBuy  = parseFloat(entry.buy?.median)  > 0 ? parseFloat(entry.buy.median)  : null;

    result[id] = { minSell, maxBuy, medianSell, medianBuy };
  });

  return result;
}


/**
 * Test consumer for getAllMarketAssignments().
 * Creates/clears a sheet named "Market Assignments Test"
 * and writes out all assignments.
 */
function testConsumerAssignments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Market Assignments Test";

  // Get the market assignments (array of objects)
  const assignments = getAllMarketAssignments();

  // Flatten each assignment into rows
  const rows = [];
  assignments.forEach(a => {
    a.type_ids.forEach(type_id => {
      rows.push([type_id, a.market_id, a.market_type]);
    });
  });

  // Headers
  const header = ["type_id", "market_id", "market_type"];

  // Create/clear sheet
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }

  // Write data
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  Logger.log(`Wrote ${rows.length} assignment rows to '${sheetName}'`);
}

function emergencyCompactMarketPrices() {
  const NAME = 'Market Prices';       // <-- change if needed
  const KEEP_DAYS = 14;               // keep last 14 days only
  const MAX_ROWS = 120000;            // hard cap after prune

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(NAME);
  if (!sh) throw new Error('Sheet not found: ' + NAME);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;    // avoid overlap

  try {
    const values = sh.getDataRange().getValues();
    if (values.length <= 1) {
      // nothing to prune, but still shrink excess empty rows
      trimTrailingEmptyRows(sh);
      return;
    }

    const header = values[0];
    const rows = values.slice(1);

    // Find date column by header name (fallback to col 1 if not found)
    let dateCol = header.findIndex(h => String(h).toLowerCase() === 'date');
    if (dateCol < 0) dateCol = 0;

    const cutoff = new Date(Date.now() - KEEP_DAYS*24*60*60*1000);

    // Keep only rows with valid date >= cutoff
    const recent = rows.filter(r => {
      const d = r[dateCol];
      return d && d instanceof Date && d >= cutoff;
    });

    // Enforce hard cap (keep most recent slice at the end)
    const keep = recent.slice(-Math.min(MAX_ROWS-1, recent.length));
    const out = [header].concat(keep);

    sh.clearContents(); // keep formats/filters; clears values fast
    sh.getRange(1,1,out.length, out[0].length).setValues(out);

    // Finally, remove the millions of empty rows below data
    trimTrailingEmptyRows(sh);
  } finally {
    lock.releaseLock();
  }
}



function trimTrailingEmptyRows(sh) {
  const NAME = 'Market Prices'; // change if needed
   sh = SpreadsheetApp.getActive().getSheetByName(NAME);
  if (!sh) throw new Error('Sheet not found: ' + NAME);

  // If a filter is applied, row deletes can fail—remove it first.
  const filter = sh.getFilter && sh.getFilter();
  if (filter) filter.remove();

  const last = sh.getLastRow();   // e.g., 87500
  const max  = sh.getMaxRows();   // e.g., 962306
  let extra  = max - last;
  if (extra <= 0) return;

  const BLOCK = 20000;            // delete in chunks to avoid timeouts
  while (extra > 0) {
    const n = Math.min(BLOCK, extra);
    sh.deleteRows(last + 1, n);
    extra -= n;
    // Utilities.sleep(100); // uncomment if you still hit rate limits
  }
}