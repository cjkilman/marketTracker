function buildCandlestickData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const daysBack = parseInt(config["DaysForCandlestick"], 10) || 30;

  const historySheet = ss.getSheetByName("History");
  if (!historySheet) throw new Error("History sheet not found.");

  const lastDate = new Date();
  const firstDate = new Date();
  firstDate.setDate(lastDate.getDate() - daysBack);

  const data = historySheet.getDataRange().getValues();
  const headers = data.shift();
  const dateIndex = headers.indexOf("Date");
  const openIndex = headers.indexOf("Open");
  const highIndex = headers.indexOf("High");
  const lowIndex = headers.indexOf("Low");
  const closeIndex = headers.indexOf("Close");

  const filtered = data.filter(row => {
    const rowDate = new Date(row[dateIndex]);
    return rowDate >= firstDate;
  });

  const candleSheet = ss.getSheetByName("Candlestick Data") || ss.insertSheet("Candlestick Data");
  candleSheet.clear();
  candleSheet.appendRow(["Date", "Low", "Open", "Close", "High"]);

  filtered.forEach(row => {
    candleSheet.appendRow([
      row[dateIndex], row[lowIndex], row[openIndex], row[closeIndex], row[highIndex]
    ]);
  });
}