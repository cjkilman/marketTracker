// cjkilman/markettracker/marketTracker-dev/src/BigQueryPipe.gs.js

/**
 * INDUSTRIAL PIPE: Direct-to-Vault Streaming.
 * Bypasses Google Sheets entirely to prevent the "Spinning Wheel" crash.
 */
function streamToBigQuery(dataRows) {
  if (!dataRows || dataRows.length === 0) return;

const cfg = getConfig(); 
  const projectId = cfg.BQ_PROJECT_ID;
  const datasetId = cfg.BQ_DATASET_ID;
  const tableId = 'market_prices_staged'; // Leave this hardcoded if it is a dedicated staging table

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

/**
 * MAINTENANCE: Automatically resets the BigQuery Circuit Breaker.
 * Intended to run once every 24 hours to recover from daily quota limits.
 */
function autoResetCircuitBreaker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Market Config");

  if (!configSheet) {
    console.error("[MAINTENANCE] Failed to run Circuit Breaker reset: Config sheet missing.");
    return;
  }

  const data = configSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "BQ_ENABLED") {
      const isCurrentlyEnabled = (data[i][1] === true || String(data[i][1]).toLowerCase() === "true");

      if (!isCurrentlyEnabled) {
        console.warn("[MAINTENANCE] Circuit Breaker was tripped. Auto-resetting to ENABLED for the new day.");
        configSheet.getRange(i + 1, 2).setValue(true);

        // Optional: Clear the safety tape so the next orchestrator pulse does a full fresh pull
        PropertiesService.getScriptProperties().deleteProperty('LAST_FUZZ_FETCH');
      } else {
        console.log("[MAINTENANCE] Circuit Breaker is already ENABLED. No action needed.");
      }
      return;
    }
  }
}

/**
 * Helper to flip the BQ_ENABLED flag from the menu OR background triggers
 */
function toggleBigQueryCircuitBreaker(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Market Config");

  if (!configSheet) {
    console.error("[CIRCUIT BREAKER] Failed: 'Market Config' sheet not found.");
    return;
  }

  const data = configSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "BQ_ENABLED") {
      const currentValue = data[i][1];

      // Determine the new state
      const newValue = (currentValue === true || String(currentValue).toLowerCase() === "true") ? false : true;

      // Write the new state back to the sheet
      configSheet.getRange(i + 1, 2).setValue(newValue);

      // Log it for the background execution logs
      const status = newValue ? "[ENABLED]" : "[DISABLED]";
      console.warn(`[CIRCUIT BREAKER] BigQuery Pipe is now ${status}`);

      // Safe Toast: Only attempt to show a UI toast if the context allows it
      try {
        ss.toast(`BigQuery Pipe is now ${status}`, "Circuit Breaker");
      } catch (e) {
        // Silently ignore UI errors when called by a background time-trigger
      }
      return;
    }
  }

  console.error("[CIRCUIT BREAKER] Failed: BQ_ENABLED setting not found.");
  try {
    ss.toast("BQ_ENABLED setting not found in Market Config.", "Error");
  } catch (e) { }
}

function runVaultMergeAndReset(ctx, targetSheetName = "Market Prices") {
  const { ss, cfg } = ctx;
  const sh = ss.getSheetByName(targetSheetName);

  if (!sh || sh.getLastRow() < 2) return;

  const isVaultOk = (String(cfg.BQ_ENABLED).toUpperCase() === "TRUE");
  if (!isVaultOk) return;

  try {
    const data = sh.getDataRange().getValues();
    const success = bigQueryBatchLoad_(cfg.BQ_PROJECT_ID, cfg.BQ_DATASET_ID, cfg.BQ_Market_Prices, data);

    if (success) {
      console.log(`[VaultGate] Success! Clearing landing pad...`);

      // RE-FETCH the row count right here to avoid "Out of Bounds" errors
      const currentRowCount = sh.getLastRow();
      const rowsToDelete = currentRowCount - 1;

      if (rowsToDelete > 0) {
        try {
          sh.deleteRows(2, rowsToDelete);
          console.log(`[VaultGate] Successfully cleared ${rowsToDelete} rows.`);
        } catch (e) {
          // If maintenance already cleared it, we just ignore the error and move on
          console.warn("[VaultGate] Sheet was modified during upload. Skipping manual clear.");
        }
      }
    }
  } catch (e) {
    console.error("Vault Error: " + e.message);
  }
}

function forceVaultMerge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = getAppContext(ss); // Handshake
  console.log("[LAUNCH] Manually triggering Vault Merge...");

  ctx.props.setProperty('exec_start_time', String(Date.now()));
  runVaultMergeAndReset(ctx); // Pass the ctx

  console.log("[SUCCESS] Vault Merge Complete.");
}

/**
 * INTERNAL HELPER: Executes the BigQuery Load Job
 * OPTIMIZED: Uses "Streaming Inserts" (Tabledata.insertAll) instead of Load Jobs.
 * Bypasses the file-upload proxy that causes 503s and chunks data to save memory.
 */
function bigQueryBatchLoad_(projectId, datasetId, tableId, dataArray) {
  const headers = dataArray[0];
  const rows = dataArray.slice(1);
  const CHUNK_SIZE = 2500; // Extremely safe memory limit for Apps Script

  console.log(`[Vault] Initiating streaming insert of ${rows.length} rows to ${tableId}...`);

  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);

    // Build the specific JSON format required for Streaming Inserts
    const insertAllRequest = {
      skipInvalidRows: false,
      ignoreUnknownValues: true, // Prevents crashes if your sheet has an extra blank column
      rows: chunk.map(row => {
        let obj = {};
        headers.forEach((h, index) => {
          let val = row[index];
          // Format dates strictly for BigQuery
          if (val instanceof Date) {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
          }
          // Only add the field if the header actually has a name
          if (h && String(h).trim() !== "") {
            obj[String(h).trim()] = val;
          }
        });
        return { json: obj };
      })
    };

    let chunkSuccess = false;
    const maxRetries = 3;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const response = BigQuery.Tabledata.insertAll(insertAllRequest, projectId, datasetId, tableId);

        // Streaming inserts do not crash the script on bad data; they return the exact error.
        if (response.insertErrors && response.insertErrors.length > 0) {
          const exactError = response.insertErrors[0].errors[0].message;
          console.error(`[Vault] Schema Error in chunk: ${exactError}`);
          return false; // Abort so we don't lose the sheet
        }

        chunkSuccess = true;
        break; // Chunk succeeded, break out of the retry loop

      } catch (e) {
        if (e.message.includes("503") || e.message.includes("Service Unavailable") || e.message.includes("timeout")) {
          console.warn(`[Vault] API hiccup on chunk ${i}. Retrying attempt ${attempt}...`);
          if (attempt === maxRetries) throw e;
          Utilities.sleep(Math.pow(2, attempt) * 1000);
        } else {
          console.error(`[Vault] Hard API Error: ${e.message}`);
          throw e; // Hard error (like Permission Denied)
        }
      }
    }

    if (!chunkSuccess) return false;

    // Force garbage collection / memory flush between chunks
    Utilities.sleep(200);
  }

  return true; // All chunks succeeded
}

/**
 * Hard-sets the BQ_ENABLED flag to prevent toggle-loops during quota hits.
 */
function setBigQueryCircuitBreaker(ss, targetState) {
  const configSheet = ss.getSheetByName("Market Config");
  if (!configSheet) return;

  const data = configSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "BQ_ENABLED") {
      configSheet.getRange(i + 1, 2).setValue(targetState);
      console.warn(`[GATE] BigQuery Pipe set to ${targetState ? '[ENABLED]' : '[DISABLED]'}`);
      return;
    }
  }
}