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

// cjkilman/markettracker/marketTracker-dev/src/BigQueryPipe.gs.js

/**
 * REVISED: Uses a hardcoded schema to prevent "No schema specified" errors
 * during fresh table creation.
 */
function streamToBigQuery(rows) {
  const config = getConfig(); 
  if (config.BQ_ENABLED === false || String(config.BQ_ENABLED).toLowerCase() === "false") {
    console.warn("[CIRCUIT BREAKER] BigQuery streaming is disabled.");
    return;
  }

  if (!rows || rows.length === 0) return;

  const projectId = 'tenacious-tiger-345318'; 
  const datasetId = 'market_data';
  const tableId = 'market_prices';

// Define the explicit schema to match your manual table exactly
  const schema = {
    fields: [
      { name: 'date', type: 'TIMESTAMP', mode: 'REQUIRED' }, // Match the REQUIRED mode
      { name: 'market_id', type: 'INTEGER', mode: 'NULLABLE' },
      { name: 'market_type', type: 'STRING', mode: 'NULLABLE' },
      { name: 'type_id', type: 'INTEGER', mode: 'NULLABLE' },
      { name: 'min_sell', type: 'FLOAT', mode: 'NULLABLE' },
      { name: 'max_buy', type: 'FLOAT', mode: 'NULLABLE' },
      { name: 'median_sell', type: 'FLOAT', mode: 'NULLABLE' },
      { name: 'median_buy', type: 'FLOAT', mode: 'NULLABLE' }
    ]
  };
  

  // Transform rows to NDJSON
  const jsonRows = rows.map(row => {
    const isObject = !Array.isArray(row);
    const rawDate = isObject ? row.date : row[0];
    
    return JSON.stringify({
      date: (rawDate instanceof Date) ? rawDate.toISOString() : new Date().toISOString(),
      market_id: parseInt(isObject ? row.market_id : row[1]),
      market_type: String(isObject ? row.market_type : row[2]),
      type_id: parseInt(isObject ? row.type_id : row[3]),
      min_sell: parseFloat(isObject ? row.min_sell : row[4]) || null,
      max_buy: parseFloat(isObject ? row.max_buy : row[5]) || null,
      median_sell: parseFloat(isObject ? row.median_sell : row[6]) || null,
      median_buy: parseFloat(isObject ? row.median_buy : row[7]) || null
    });
  }).join('\n');

  const blob = Utilities.newBlob(jsonRows, 'application/octet-stream');

  const job = {
    configuration: {
      load: {
        destinationTable: { projectId, datasetId, tableId },
        schema: schema, // Explicit schema is safer than autodetect
        sourceFormat: 'NEWLINE_DELIMITED_JSON',
        writeDisposition: 'WRITE_APPEND' 
      }
    }
  };

  try {
    const runJob = BigQuery.Jobs.insert(job, projectId, blob);
    console.log(`[BIGQUERY] Job started: ${runJob.jobReference.jobId}`);
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