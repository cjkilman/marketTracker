function runMarketUpdate() {
  const startTime = new Date();
  let status = "SUCCESS";

  try {
    // 1. Remove expired intraday prices
    pruneExpiredPrices();

    // 2. Fetch fresh market prices from ESI API
    getCurrentMarketPrices();

    // 3. Append today's prices to 1-year rolling history
    updateHistory();

    // 4. Build candlestick chart data for configured days
    buildCandlestickData();

  } catch (err) {
    status = "FAILED: " + err.message;
    Logger.log("Error during run: " + err);
  } finally {
    const endTime = new Date();
    logRunStatus(startTime, endTime, status);
  }
}

function logRunStatus(startTime, endTime, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Run Log");

  // Create sheet if missing
  if (!logSheet) {
    logSheet = ss.insertSheet("Run Log");
    logSheet.appendRow(["Date", "Start Time", "End Time", "Duration (mins)", "Status"]);
  }

  const duration = ((endTime.getTime() - startTime.getTime()) / 60000).toFixed(2);
  logSheet.appendRow([
    Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy-MM-dd"),
    Utilities.formatDate(startTime, Session.getScriptTimeZone(), "HH:mm:ss"),
    Utilities.formatDate(endTime, Session.getScriptTimeZone(), "HH:mm:ss"),
    duration,
    status
  ]);
}
