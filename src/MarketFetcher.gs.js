function getCurrentMarketPrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const maxLogIDs = config["MaxLogIDs"] ? parseInt(config["MaxLogIDs"], 10) : 750;

  const typeIDs = getTypeIDsFromItemList(maxLogIDs);
  const marketCombos = getMarketSettings();

const sheet = getOrCreateSheet(
  ss,  // Add this so the function gets the correct context
  "Market Prices",
  ["type_id", "market_id", "market_type", "max_buy", "min_sell", "date"]
);

  const now = new Date();
  const cutoff = new Date(now.getTime() - 24 * 60 * 60 * 1000);

  // Keep existing data within 24 hours
  let existing = sheet.getDataRange().getValues();
  const header = existing.shift();
  existing = existing.filter(row => new Date(row[5]) >= cutoff);

  // Process each market combo
  marketCombos.forEach(({ market_id, market_type }) => {
    const data = getMarketPrices(typeIDs, market_id, market_type);
    const rows = Object.entries(data).map(([type_id, prices]) => [
      parseInt(type_id, 10),
      market_id,
      market_type,
      prices.maxBuy || "",
      prices.minSell || "",
      now
    ]);
    existing.push(...rows);
  });

  // Write back
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, existing.length, header.length).setValues(existing);
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

function getAllMarketAssignments() {
  const typeIDs = getTypeIDsFromItemList(); // Already returns sanitized list
  const marketCombos = getMarketSettings(); // [{ market_id, market_type }, ...]

  const assignments = [];

  marketCombos.forEach(({ market_id, market_type }) => {
    typeIDs.forEach(type_id => {
      assignments.push({
        type_id,
        market_id,
        market_type
      });
    });
  });

  return assignments;
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

    const minSell = parseFloat(entry.sell?.min) > 0 ? parseFloat(entry.sell.min) : null;
    const maxBuy  = parseFloat(entry.buy?.max)  > 0 ? parseFloat(entry.buy.max)  : null;

    result[id] = { minSell, maxBuy };
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

  // Map objects -> rows
  const rows = assignments.map(a => [
    a.type_id,
    a.market_id,
    a.market_type
  ]);

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

  Logger.log(`Wrote ${rows.length} assignments to '${sheetName}'`);
}