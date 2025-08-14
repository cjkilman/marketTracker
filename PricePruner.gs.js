/**
 * Remove price rows older than 24 hours from 'market prices'
 * Keeps header row intact with plain strings.
 */
function pruneExpiredPrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("market prices");
  if (!sheet) return;

  const now = new Date();
  const cutoff = now.getTime() - (24 * 60 * 60 * 1000);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const header = data[0].map(v => v != null ? String(v).trim() : "");
  const newData = [header];

  for (let i = 1; i < data.length; i++) {
    const ts = new Date(data[i][0]);
    if (!isNaN(ts) && ts.getTime() >= cutoff) {
      newData.push(data[i]);
    }
  }

  sheet.clearContents();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

/**
 * Update market history from 'market prices'
 * Header matching is case-insensitive and trim-safe.
 */
function updateMarketHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName("market_history");
  const priceSheet = ss.getSheetByName("market prices");
  if (!historySheet || !priceSheet) return;

  const findCol = (headerArr, name) =>
    headerArr.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const histHeader = historySheet.getRange(1, 1, 1, historySheet.getLastColumn())
    .getValues()[0].map(h => h != null ? String(h).trim() : "");

  const hDateIdx   = findCol(histHeader, "date");
  const hTypeIdx   = findCol(histHeader, "type_id");
  const hMarketIdx = findCol(histHeader, "market_id");
  const hMTypeIdx  = findCol(histHeader, "market_type");
  const hBuyIdx    = findCol(histHeader, "max_buy");
  const hSellIdx   = findCol(histHeader, "max_sell");

  const priceHeader = priceSheet.getRange(1, 1, 1, priceSheet.getLastColumn())
    .getValues()[0].map(h => h != null ? String(h).trim() : "");

  const pTypeIdx    = findCol(priceHeader, "type_id");
  const pMarketIdx  = findCol(priceHeader, "market_id");
  const pMTypeIdx   = findCol(priceHeader, "market_type");
  const pMaxBuyIdx  = findCol(priceHeader, "max_buy");
  const pMaxSellIdx = findCol(priceHeader, "max_sell");
  const pDateIdx    = findCol(priceHeader, "date");

  if ([pTypeIdx, pMarketIdx, pMTypeIdx, pMaxBuyIdx, pMaxSellIdx, pDateIdx].some(i => i < 0)) {
    throw new Error("market prices headers missing expected columns (type_id, market_id, market_type, max_buy, max_sell, date).");
  }

  const prices = priceSheet.getRange(2, 1, priceSheet.getLastRow() - 1, priceSheet.getLastColumn()).getValues();
  if (!prices.length) return;

  const historyData = {};
  for (let row of prices) {
    const key = [row[pTypeIdx], row[pMarketIdx], row[pMTypeIdx], row[pDateIdx]].join("|");
    if (!historyData[key]) {
      historyData[key] = {
        type_id: row[pTypeIdx],
        market_id: row[pMarketIdx],
        market_type: row[pMTypeIdx],
        max_buy: row[pMaxBuyIdx],
        max_sell: row[pMaxSellIdx],
        date: row[pDateIdx]
      };
    }
  }

  const newRows = Object.values(historyData).map(d => [
    d.date, d.type_id, d.market_id, d.market_type, d.max_buy, d.max_sell
  ]);

  if (newRows.length) {
    historySheet.insertRowsAfter(historySheet.getLastRow() || 1, newRows.length);
    historySheet.getRange(historySheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
}

/**
 * Main daily update chain.
 */
function runDailyUpdate() {
  pruneExpiredPrices();
  getCurrentMarketPrices();
  updateMarketHistory();
  calculateMedianPrices();
}