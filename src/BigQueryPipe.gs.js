// cjkilman/markettracker/marketTracker-dev/src/BigQueryPipe.gs.js

const jobConfig = {
  configuration: {
    load: {
      destinationTable: {
        projectId: 'tenacious-tiger-345318', 
        datasetId: 'market_data',
        tableId: 'market_prices'
      },
      // --- THE FIX: This MUST be inside 'load' ---
      autodetect: true, 
      // ------------------------------------------
      sourceFormat: 'NEWLINE_DELIMITED_JSON',
      writeDisposition: 'WRITE_APPEND' 
    }
  }
};

/**
 * Streams market data to BigQuery using free "Load Jobs".
 * REVISED: Handles Object inputs and includes Circuit Breaker + Schema Autodetect.
 */
function streamToBigQuery(rows) {
  // 1. CIRCUIT BREAKER CHECK
  // Fetches configuration to check if BigQuery is enabled
  const config = getConfig(); 
  if (config.BQ_ENABLED === false || config.BQ_ENABLED === "false") {
    console.warn("[CIRCUIT BREAKER] BigQuery streaming is disabled via Config.");
    return;
  }

  if (!rows || rows.length === 0) return;

  const projectId = 'tenacious-tiger-345318'; 
  const datasetId = 'market_data';
  const tableId = 'market_prices';

  // 2. Transform rows into Newline-Delimited JSON
  const jsonRows = rows.map(row => {
    const isObject = !Array.isArray(row);

    const rawDate    = isObject ? row.date        : row[0];
    const marketId   = isObject ? row.market_id   : row[1];
    const marketType = isObject ? row.market_type : row[2];
    const typeId     = isObject ? row.type_id     : row[3];
    const minSell    = isObject ? row.min_sell    : row[4];
    const maxBuy     = isObject ? row.max_buy     : row[5];
    const medSell    = isObject ? row.median_sell : row[6];
    const medBuy     = isObject ? row.median_buy  : row[7];

    let validDate;
    if (rawDate instanceof Date) {
      validDate = !isNaN(rawDate) ? rawDate.toISOString() : new Date().toISOString();
    } else if (typeof rawDate === 'string') {
      validDate = rawDate || new Date().toISOString();
    } else {
      validDate = new Date().toISOString();
    }

    return JSON.stringify({
      date: validDate,
      market_id: parseInt(marketId),
      market_type: String(marketType),
      type_id: parseInt(typeId),
      min_sell: parseFloat(minSell) || null,
      max_buy: parseFloat(maxBuy) || null,
      median_sell: parseFloat(medSell) || null,
      median_buy: parseFloat(medBuy) || null
    });
  }).join('\n');

  const blob = Utilities.newBlob(jsonRows, 'application/octet-stream');

  // 3. Configure the Load Job with THE FIX
  const jobConfig = {
    configuration: {
      load: {
        destinationTable: {
          projectId: projectId,
          datasetId: datasetId,
          tableId: tableId
        },
        // --- THE CRITICAL FIX ---
        // Tells BigQuery to infer the schema from the JSON keys
        autodetect: true, 
        // ------------------------
        sourceFormat: 'NEWLINE_DELIMITED_JSON',
        writeDisposition: 'WRITE_APPEND' 
      }
    }
  };

  try {
    const runJob = BigQuery.Jobs.insert(jobConfig, projectId, blob);
    console.log(`[BIGQUERY] Load Job started: ${runJob.jobReference.jobId}. Rows: ${rows.length}`);
  } catch (err) {
    console.error("[BIGQUERY] Pipe Error: " + err.message);
  }
}

/**
 * THE RESET CROWBAR
 * Logic: Deletes the existing table so the next stream job 
 * can recreate it with a fresh 60-day clock and schema.
 */
function resetBigQueryTable() {
  const projectId = 'tenacious-tiger-345318'; 
  const datasetId = 'market_data';
  const tableId = 'market_prices';
  
  try {
    // 1. Kill the old table
    BigQuery.Tables.remove(projectId, datasetId, tableId);
    console.log(`[CROWBAR] Table ${tableId} deleted. The sandbox clock has been reset.`);
    
    // 2. OPTIONAL: Trigger an immediate fresh stream to rebuild it
    // This ensures your formulas don't stay broken for long.
    console.log("[CROWBAR] Rebuilding table with fresh schema...");
    
    // We grab a small slice of data to "prime" the new table
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawData = ss.getSheetByName("Market_Data_Raw").getDataRange().getValues().slice(0, 10);
    
    // This call will now use the 'autodetect: true' logic we added
    streamToBigQuery(rawData);
    
    SpreadsheetApp.getActiveSpreadsheet().toast("BigQuery Reset Successful", "Tycoon Operations");
  } catch (err) {
    console.error("[CROWBAR] Reset Failed: " + err.message);
    if (err.message.indexOf("Not found") > -1) {
      console.log("Table didn't exist anyway. Ready for fresh creation.");
    }
  }
}