function openCloseSnapshot() {
  getCurrentMarketPrices();
  computeGroupMediansWithCurrentBuySell(true);
}

function computeGroupMediansWithCurrentBuySell(appendHistory) {
  appendHistory = appendHistory || false; // default false

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("market prices");
  var data = sheet.getDataRange().getValues();

  var headers = data[0];
  var dateCol = headers.indexOf("date_time");
  var buyCol = headers.indexOf("max_buy");
  var sellCol = headers.indexOf("min_sell");

  var props = PropertiesService.getScriptProperties();
  var lastMode = props.getProperty('lastOpenCloseMode');

  if (!lastMode && appendHistory) {
    lastMode = 'close';
  }

  var showOpen;
  if (appendHistory) {
    showOpen = lastMode !== 'open';
    props.setProperty('lastOpenCloseMode', showOpen ? 'open' : 'close');
  } else {
    showOpen = lastMode === 'open';
  }

  var groups = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var key = [row[0], row[1], row[2]].join("||");
    if (!groups[key]) {
      groups[key] = { buy: [], sell: [], rows: [] };
    }
    if (row[buyCol] != null && row[buyCol] !== "") groups[key].buy.push(row[buyCol]);
    if (row[sellCol] != null && row[sellCol] !== "") groups[key].sell.push(row[sellCol]);
    groups[key].rows.push(row);
  }

  var result = [];
  result.push([
    "type_id", "market_id", "market_type",
    "Median Max Buy", "Median Min Sell",
    "Current Max Buy", "Current Min Sell",
    "% Change Median Max Buy vs Current Max Buy",
    "% Change Median Min Sell vs Current Min Sell"
  ]);

  var now = new Date();

  for (var key in groups) {
    var g = groups[key];
    var validRows = g.rows.filter(function(row) {
      return row[dateCol] && new Date(row[dateCol]) <= now;
    });
    var latestRow = validRows.reduce(function(latest, row) {
      return (!latest || new Date(row[dateCol]) > new Date(latest[dateCol])) ? row : latest;
    }, null);

    var medBuy = median(g.buy);
    var medSell = median(g.sell);

    var currMaxBuy = latestRow ? latestRow[buyCol] : null;
    var currMinSell = latestRow ? latestRow[sellCol] : null;

    var changeBuy = (currMaxBuy && medBuy) ? ((currMaxBuy - medBuy) / medBuy) : null;
    var changeSell = (currMinSell && medSell) ? ((currMinSell - medSell) / medSell) : null;

    var parts = key.split("||");
    result.push([
      parts[0], parts[1], parts[2],
      medBuy, medSell,
      currMaxBuy, currMinSell,
      changeBuy, changeSell
    ]);
  }

  var outputSheet = ss.getSheetByName("Medians Output") || ss.insertSheet("Medians Output");
  outputSheet.clearContents();
  outputSheet.getRange(1, 1, result.length, result[0].length).setValues(result);

  if (appendHistory) {
    appendOpenCloseHistory(result, showOpen);
  }
}

function appendOpenCloseHistory(result, showOpen) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheetName = "Open/Close History";
  var histSheet = ss.getSheetByName(histSheetName) || ss.insertSheet(histSheetName);

  if (histSheet.getLastRow() === 0) {
    var header = [
      "Timestamp", "Open", "Close",
      "type_id", "market_id", "market_type",
      "Low Median", "High Median",
      "High", "Low",
      "% Change Median Max Buy vs Current Max Buy",
      "% Change Median Min Sell vs Current Min Sell"
    ];
    histSheet.appendRow(header);
  }

  var nowStr = new Date();
  var rowsToAppend = [];

  for (var i = 1; i < result.length; i++) {
    var openVal = showOpen ? result[i][5] : "";
    var closeVal = showOpen ? "" : result[i][5];
    rowsToAppend.push([
      nowStr, openVal, closeVal,
      result[i][0], result[i][1], result[i][2],
      result[i][4], result[i][3],
      result[i][5], result[i][6],
      result[i][7], result[i][8]
    ]);
  }

  histSheet.getRange(histSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
    .setValues(rowsToAppend);
}

// NEW — Build Candlestick Data
function buildCandlestickData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var histSheet = ss.getSheetByName("Open/Close History");
  if (!histSheet) throw new Error("Open/Close History sheet not found.");

  var config = getConfig();

  var data = histSheet.getDataRange().getValues();
  var headers = data[0];
  var tsCol = headers.indexOf("Timestamp");
  var openCol = headers.indexOf("Open");
  var closeCol = headers.indexOf("Close");
  var highCol = headers.indexOf("High");
  var lowCol = headers.indexOf("Low");
  var typeCol = headers.indexOf("type_id");
  var marketCol = headers.indexOf("market_id");
  var marketTypeCol = headers.indexOf("market_type");

  var filtered = data.filter(function(row, idx) {
    if (idx === 0) return false;
    return row[typeCol] == config.type_id &&
           row[marketCol] == config.market_id &&
           row[marketTypeCol] == config.market_type;
  });

  filtered.sort(function(a, b) {
    return new Date(a[tsCol]) - new Date(b[tsCol]);
  });

  var N = 1000; // last N rows
  if (filtered.length > N) {
    filtered = filtered.slice(filtered.length - N);
  }

  var candleData = [["DateTime", "Low", "Open", "Close", "High"]];
  filtered.forEach(function(row) {
    candleData.push([
      new Date(row[tsCol]),
      row[lowCol],
      row[openCol],
      row[closeCol],
      row[highCol]
    ]);
  });

  var candleSheet = ss.getSheetByName("Candlestick Data") || ss.insertSheet("Candlestick Data");
  candleSheet.clearContents();
  candleSheet.getRange(1, 1, candleData.length, candleData[0].length).setValues(candleData);
}

// Config management
function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet = ss.getSheetByName("Config");
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet("Config");
    cfgSheet.getRange("A1:B3").setValues([
      ["type_id", 34],
      ["market_id", 30002187],
      ["market_type", "system"]
    ]);
  }
  var vals = cfgSheet.getRange("A1:B3").getValues();
  var cfg = {};
  vals.forEach(function(row) {
    cfg[row[0]] = row[1];
  });
  return cfg;
}

function median(values) {
  if (!values.length) return null;
  values.sort(function(a, b) { return a - b; });
  var half = Math.floor(values.length / 2);
  if (values.length % 2)
    return values[half];
  else
    return (values[half - 1] + values[half]) / 2.0;
}