// cjkilman/markettracker/marketTracker-dev/src/BigQueryPipe.gs.js

/**
 * INDUSTRIAL PIPE: Direct-to-Vault Streaming.
 * Bypasses Google Sheets entirely to prevent the "Spinning Wheel" crash.
 */
function streamToBigQuery(dataRows) {
  if (!dataRows || dataRows.length === 0) return;

  const cfg = getConfig(); // Pull IDs from your central config
  const projectId = cfg.BQ_PROJECT_ID || 'tenacious-tiger-345318';
  const datasetId = 'market_data';
  const tableId = 'market_prices_staged';

  // 1. Convert data to Newline Delimited JSON (BigQuery's required format)
  let ndjson = "";
  dataRows.forEach(row => {
    // Ensure data types match your Native BigQuery table schema perfectly
    const bqRow = {
      date: row.date.toISOString(), // Must be ISO string for BQ TIMESTAMP
      market_id: String(row.market_id),
      market_type: String(row.market_type),
      type_id: Number(row.type_id),
      min_sell: Number(row.min_sell),
      max_buy: Number(row.max_buy),
      median_sell: Number(row.median_sell),
      median_buy: Number(row.median_buy)
    };
    ndjson += JSON.stringify(bqRow) + "\n";
  });

  // 2. Build the BigQuery Upload Job
  const blob = Utilities.newBlob(ndjson, 'application/octet-stream');
  const job = {
    configuration: {
      load: {
        destinationTable: {
          projectId: projectId,
          datasetId: datasetId,
          tableId: tableId
        },
        sourceFormat: 'NEWLINE_DELIMITED_JSON',
        writeDisposition: 'WRITE_APPEND' // ADD to the vault, don't overwrite
      }
    }
  };

  // 3. FIRE DIRECTLY TO THE CLOUD
  try {
    BigQuery.Jobs.insert(job, projectId, blob);
    // Notice: We DO NOT write to BQ_LIVE_DATA or Market Prices anymore.
  } catch (e) {
    console.error(`[BQ PIPE FATAL] Failed to stream to BigQuery: ${e.message}`);
    throw e; // Pass the error back up so the worker knows it failed
  }
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