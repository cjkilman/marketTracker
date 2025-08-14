function getCurrentMarketPrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const maxLogIDs = config["MaxLogIDs"] ? parseInt(config["MaxLogIDs"], 10) : 750;

  const typeIDs = getTypeIDsFromItemList(maxLogIDs);
  const marketCombos = getMarketSettings();

  const sheet = getOrCreateSheet(
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

function getTypeIDsFromItemList(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Item List Back End");
  if (!sheet) throw new Error("Item List Back End sheet not found.");
  let ids = sheet.getRange("A2:A" + sheet.getLastRow()).getValues().flat().filter(Number);
  ids = [...new Set(ids)]; // unique
  if (limit && ids.length > limit) ids = ids.slice(0, limit);
  return ids;
}

function getMarketSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Market Settings");
  if (!sheet) throw new Error("Market Settings sheet not found.");
  const data = sheet.getRange(3, 4, sheet.getLastRow() - 2, 3).getValues();
  const combos = [];
  const types = ["station", "system", "region"];
  data.forEach(row => {
    row.forEach((id, i) => {
      if (id) combos.push({ market_id: Number(id), market_type: types[i] });
    });
  });
  return combos.filter((v, i, a) => a.findIndex(t =>
    t.market_id === v.market_id && t.market_type === v.market_type
  ) === i);
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