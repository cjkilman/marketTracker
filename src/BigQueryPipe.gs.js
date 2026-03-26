// cjkilman/markettracker/marketTracker-dev/src/BigQueryPipe.gs.js

function streamToBigQuery(rows) {
  if (!rows || rows.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("BQ_LIVE_DATA");
  
  if (!sheet) {
    sheet = ss.insertSheet("BQ_LIVE_DATA");
    sheet.appendRow(['date', 'market_id', 'market_type', 'type_id', 'min_sell', 'max_buy', 'median_sell', 'median_buy']);
  }

  // Keep only the last 5000 rows to stay fast
  const lastRow = sheet.getLastRow();
  if (lastRow > 5000) {
    sheet.deleteRows(2, rows.length); 
  }

  const dataToAppend = rows.map(row => {
    const isArr = Array.isArray(row);
    return [
      isArr ? row[0] : row.date,
      isArr ? row[1] : row.market_id,
      isArr ? row[2] : row.market_type,
      isArr ? row[3] : row.type_id,
      isArr ? row[4] : row.min_sell,
      isArr ? row[5] : row.max_buy,
      isArr ? row[6] : row.median_sell,
      isArr ? row[7] : row.median_buy
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, dataToAppend.length, 8).setValues(dataToAppend);
  console.log(`[LOCAL] Wrote ${rows.length} rows to BQ_LIVE_DATA sheet.`);
}
/**
 * THE RESET CROWBAR
 * Logic: Deletes the existing table so the next stream job 
 * can recreate it with a fresh schema.
 */
function resetBigQueryTable() {
  const projectId = 'tenacious-tiger-345318'; 
  const datasetId = 'market_data';
  const tableId = 'market_prices';
  
  try {
    BigQuery.Tables.remove(projectId, datasetId, tableId);
    console.log(`[CROWBAR] Table ${tableId} deleted. The sandbox clock has been reset.`);
    
    console.log("[CROWBAR] Rebuilding table with fresh schema...");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Market Prices");
    
    if(sheet) {
      // Get the last 10 rows to prime the table
      const lastRow = sheet.getLastRow();
      if(lastRow > 1) {
        const startRow = Math.max(2, lastRow - 9);
        const rawData = sheet.getRange(startRow, 1, 10, 8).getValues();
        streamToBigQuery(rawData);
      }
    }
    SpreadsheetApp.getActiveSpreadsheet().toast("BigQuery Reset Successful", "Tycoon Operations");
  } catch (err) {
    console.error("[CROWBAR] Reset Failed: " + err.message);
    if (err.message.indexOf("Not found") > -1) {
      console.log("Table didn't exist anyway. Ready for fresh creation.");
    }
  }
}