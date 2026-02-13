
// is this Sheet Formula 
// Something about All Items Candle Stick chaart?

/**
 * Note this projects Paused pending Time and Updates
 */


function getDailyCandlestick(type_id, market_id, market_type, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Market History");
  if (!sheet) throw new Error("Market History sheet not found.");

  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Remove header row

  // Map columns by header
  const col = {};
  headers.forEach((h, i) => col[h] = i);

  // Filter rows matching type_id, market_id, market_type, and date
  const rows = data.filter(r =>
    r[col["type_id"]] === type_id &&
    r[col["market_id"]] === market_id &&
    r[col["market_type"]] === market_type &&
    r[col["date"]].toDateString() === date.toDateString()
  );

  if (!rows.length) return null;

  const minSells = rows.map(r => r[col["min_sell"]]);
  const maxBuys  = rows.map(r => r[col["max_buy"]]);

  const high = Math.max(...minSells);
  const low  = Math.min(...maxBuys);

  // Open: first min_sell of the day
  const open = rows[0][col["min_sell"]];
  // Close: last min_sell of the day
  const close = rows[rows.length - 1][col["min_sell"]];

  return { open, close, high, low };
}