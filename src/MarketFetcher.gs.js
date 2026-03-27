/** MarketFetcher.gs — STATEFUL Prices runner with Predictive Scheduling & Locks
 * Manages fetching Fuzzwork prices based on state properties.
 * Called by the masterOrchestrator.
 *
 * NOTE: This file has been modified for an APPEND-ONLY workflow and implements 
 * a "Cold Start Sentinel" to preserve data integrity on all subsequent runs.
 * 
 * NOTE: This Impliments Big Query as the Primary Sources of Market Data by using 
 * Fuz Works Aggergation stored to a Sheet Table and then merged into the Google Vault
 *  */

/* global LockService, PropertiesService, SpreadsheetApp, LoggerEx, fuzAPI, getMarketPrices, getMasterMarketRequests, getOrCreateSheet, scheduleOneTimeTrigger, JOB_LEASE_DURATION_MS, pruneOldRows, _trimTrailing_, getConfig */

// --- Constants ---
const FUZZ_JOB_PREFIX = 'fuzzJob'; // Prefix for state properties
const FUZZ_PROP_STEP = FUZZ_JOB_PREFIX + 'Step';
const FUZZ_PROP_INDEX = FUZZ_JOB_PREFIX + 'RequestIndex';
const FUZZ_PROP_ROW = FUZZ_JOB_PREFIX + 'WriteRow';
const FUZZ_PROP_LEASE = FUZZ_JOB_PREFIX + 'LeaseUntil';
const FUZZ_PROP_INIT = FUZZ_JOB_PREFIX + 'Initialized'; // NEW: Cold start sentinel
const FUZZ_SHEET_FINAL = 'Market Prices';     // Final destination sheet (used for direct append)
const FUZZ_SHEET_HEADERS = ["date", "market_id", "market_type", "type_id", "min_sell", "max_buy", "median_sell", "median_buy"];

const FUZZ_BATCH_SIZE = 750; // How many requests to process per execution run
const FUZZ_TIME_LIMIT_MS = 280000;      // Soft limit (4m 40s) before rescheduling
const FUZZ_RESCHEDULE_MS = 5000;        // Delay for rescheduling (5s)
const FUZZ_DOC_LOCK_TIMEOUT = 10000;    // Wait 10s for DocumentLock on write

// --- [PRUNE CONSTANTS] ---
const PRUNE_JOB_PREFIX = 'heavyPruneJob'; // Prefix for state properties
const PRUNE_PROP_STEP = PRUNE_JOB_PREFIX + 'Step';
const PRUNE_PROP_READ_ROW = PRUNE_JOB_PREFIX + 'ReadRow';
const PRUNE_SHEET_TEMP = 'Market_Prices_Prune_Temp'; // Temp sheet for pruning
const PRUNE_BATCH_SIZE = 5000; // How many rows to read/process at a time
// --- [END OF NEW CONSTANTS] ---

const LOG_FUZZ = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('FuzzWorker') : console);

/**
 * WRAPPER: Handles Lease & Lock.
 * - Checks if 30 minutes have passed.
 * - Calls executeWithTryLock(_updateFuzzMarketDataWorker).
 * - Updates Lease Timestamp ONLY if successful.
 */
function updateFuzzMarketDataSheet(itemSource) {
  const LOG_HEADER = '[Orchestrator]';
  const LEASE_MINUTES = 30;
  const PROPS = PropertiesService.getScriptProperties();

  // --- 1. LEASE CHECK ---
  const lastRun = parseFloat(PROPS.getProperty('LAST_FUZZ_FETCH') || '0');
  const now = Date.now();
  const minutesSince = (now - lastRun) / (1000 * 60);

  if (minutesSince < LEASE_MINUTES) {
    console.log(`${LOG_HEADER} Lease Active. Skipping Fetch. (Next run in ${(LEASE_MINUTES - minutesSince).toFixed(1)} min)`);
    return;
  }

  // --- 2. EXECUTE WITH LOCK ---
  // This calls the worker with NO arguments.
  const result = executeWithTryLock(() => {
    _updateFuzzMarketDataWorker(itemSource); // Pass it to the worker
  }, 'updateFuzzMarketDataSheet');

  // --- 3. UPDATE LEASE (On Success) ---
  if (result !== null) {
    console.log(`${LOG_HEADER} Cycle Complete. Updating Lease Timestamp.`);
    PROPS.setProperty('LAST_FUZZ_FETCH', now.toString());
  } else {
    console.warn(`${LOG_HEADER} Execution skipped by ScriptLock.`);
  }
}

/**
 * Resets the state of the Fuzz market data job.
 * MODIFIED: No longer deletes FUZZ_PROP_LEASE to enforce the 30-min window.
 */
function _resetFuzzMarketDataJobState(error) {
  LOG_FUZZ.warn(`RESETTING Fuzz Market Data Job State. Reason: ${error ? error.message : 'Completion/Manual'}`);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  try {
    SCRIPT_PROP.deleteProperty(FUZZ_PROP_STEP);
    SCRIPT_PROP.deleteProperty(FUZZ_PROP_INDEX);
    SCRIPT_PROP.deleteProperty(FUZZ_PROP_ROW);

    // REMOVE OR COMMENT OUT THIS LINE:
    // SCRIPT_PROP.deleteProperty(FUZZ_PROP_LEASE); 

    // Delete potential triggers
    deleteTriggersByName('updateFuzzMarketDataSheet');
  } catch (propError) {
    LOG_FUZZ.error(`Error deleting script properties: ${propError.message}`);
  }
  LOG_FUZZ.info("Fuzz job state reset. Lease remains active until expiration.");
}




/**
 * WORKER: Smart Execution.
 * - Reads 'Item List Back End' & 'Market Settings' internally.
 * - Loops through all markets.
 */
function _updateFuzzMarketDataWorker(itemSource) {
  // Polyfill logger
  const LOG = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('FuzzWorker') : console;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. LOAD ITEMS ---
const sourceName = itemSource || "Item List Back End"; // Use the passed source OR default
const itemSheet = ss.getSheetByName(sourceName);
  if (!itemSheet) {
    console.error("[FuzzWorker] Critical: 'Item List Back End' sheet missing.");
    return;
  }

  const itemData = itemSheet.getDataRange().getValues();
  const typeIdBatch = [];

  // Skip Header, Read Col A (Index 0)
  for (let i = 1; i < itemData.length; i++) {
    const id = parseInt(itemData[i][0]);
    if (!isNaN(id)) typeIdBatch.push(id);
  }

  if (typeIdBatch.length === 0) {
    console.warn("[FuzzWorker] No Type IDs found in backend list.");
    return;
  }

  // --- 2. LOAD MARKETS ---
  const hubSheet = ss.getSheetByName("Market Settings");
  const marketsToProcess = [];

  if (hubSheet) {
    const hubData = hubSheet.getDataRange().getValues();
    // CSV Offset: Col D (Index 3)=Station, Col E (Index 4)=System
    for (let i = 1; i < hubData.length; i++) {
      const sysId = parseInt(hubData[i][4]);
      const statId = parseInt(hubData[i][3]);

      if (!isNaN(sysId)) {
        marketsToProcess.push({ market_id: sysId, market_type: 'system' });
      } else if (!isNaN(statId)) {
        marketsToProcess.push({ market_id: statId, market_type: 'station' });
      }
    }
  }

  // Fallback to Jita
  if (marketsToProcess.length === 0) {
    marketsToProcess.push({ market_id: 60003760, market_type: 'station' });
  }

  // --- 3. EXECUTE LOOP ---
  console.log(`[FuzzWorker] Starting Cycle: ${typeIdBatch.length} items across ${marketsToProcess.length} markets.`);

  marketsToProcess.forEach(activeConfig => {
    try {
      const marketData = getMarketPrices(typeIdBatch, activeConfig.market_id, activeConfig.market_type);

      if (!marketData || Object.keys(marketData).length === 0) return;

      const rowsToStream = [];
      const timestamp = new Date().toISOString();

      typeIdBatch.forEach(typeId => {
        const data = marketData[typeId];
        if (data) {
          const buy = data.buy || {};
          const sell = data.sell || {};
          if ((buy.orderCount || 0) + (sell.orderCount || 0) > 0) {
            rowsToStream.push({
              date: timestamp,
              market_id: String(activeConfig.market_id),
              market_type: activeConfig.market_type,
              type_id: parseInt(typeId, 10),
              min_sell: parseFloat(sell.min) || 0,
              max_buy: parseFloat(buy.max) || 0,
              median_sell: parseFloat(sell.median) || 0,
              median_buy: parseFloat(buy.median) || 0
            });
          }
        }
      });

      // --- CRITICAL: STREAM TO BIGQUERY ---
      if (rowsToStream.length > 0) {
        // 1. Lock the Orchestrator out while we write
        PropertiesService.getScriptProperties().setProperty('fuzz_job_active', 'true');

        console.log(`[BQ] Streaming ${rowsToStream.length} rows for Market: ${activeConfig.market_id}`);
        streamToBigQuery(rowsToStream); // This pushes data to your BQ table

        // 2. Unlock immediately after this batch finishes
        PropertiesService.getScriptProperties().setProperty('fuzz_job_active', 'false');
      }

    } catch (e) {
      console.error(`[FuzzWorker] Error on Hub ${activeConfig.market_id}: ${e.message}`);
      // Safety: Ensure flag is cleared even on error
      PropertiesService.getScriptProperties().setProperty('fuzz_job_active', 'false');
    }
  });

  console.log(`[FuzzWorker] Cycle Complete.`);

  // --- STATIC DATA REFRESH PULSE ---
  var utilitySheet = ss.getSheetByName("Utility"); // 'ss' is already defined at the top of this function!
  if (utilitySheet) {
    // 1. Throw the kill switch
    utilitySheet.getRange("B3").setValue(0);
    SpreadsheetApp.flush();
    // 2. Flip the switch back on 
    utilitySheet.getRange("B3").setValue(1);
    console.log(`[FuzzWorker] Static Data Pulse Fired.`);
  }
}


/* ---------------------- Light Mode hygiene ---------------------- */

function lightPrePrune_(sh, maxRows) {
  const f = sh.getFilter && sh.getFilter();
  if (f) f.remove();
  _trimTrailing_(sh);
  const last = sh.getLastRow();
  if (maxRows && last > maxRows) {
    const over = last - maxRows;
    _deleteInBlocks_(sh, 2, over);
    LoggerEx && LoggerEx.log('Prices: light prune removed oldest rows:', over);
  }
}

function postTighten_(sh) {
  _trimTrailing_(sh);
}

function _trimTrailing_(sh) {
  if (!sh) return;

  const lastRow = sh.getLastRow();
  const maxRows = sh.getMaxRows();
  const lastCol = sh.getLastColumn();
  const maxCols = sh.getMaxColumns();

  // 1. Trim Rows (Existing logic)
  const rowsToDelete = maxRows - (lastRow + 5);
  if (rowsToDelete > 0) {
    try {
      _deleteInBlocks_(sh, lastRow + 6, rowsToDelete);
      console.log(`[CLEANUP] Trimmed ${rowsToDelete} rows from ${sh.getName()}`);
    } catch (e) {
      console.error(`Error trimming rows: ${e.message}`);
    }
  }

  // 2. Trim Columns (NEW - Crucial for cell limit)
  // Sheets default to 26 columns (A-Z). EVE data only needs ~8-10.
  // Trimming 15 columns from a 500k row sheet frees 7.5 million cells.
  const colsToDelete = maxCols - (lastCol + 1);
  if (colsToDelete > 0 && maxCols > 5) {
    try {
      sh.deleteColumns(lastCol + 2, colsToDelete);
      console.log(`[CLEANUP] Trimmed ${colsToDelete} columns from ${sh.getName()}`);
    } catch (e) {
      console.error(`Error trimming columns: ${e.message}`);
    }
  }
}

function _deleteInBlocks_(sh, startRow, count) {
  const BLOCK = 20000;
  if (count <= 0 || !sh) return; // Add guard clause

  const frozen = sh.getFrozenRows();
  const bodyStart = frozen + 1;
  const maxRows = sh.getMaxRows();
  const bodyRows = Math.max(0, maxRows - frozen);

  if (startRow < bodyStart) startRow = bodyStart;

  const keptTop = Math.max(0, startRow - bodyStart);
  let safeCount = Math.min(count, Math.max(0, bodyRows - 1 - keptTop));

  if (safeCount <= 0) {
    if (startRow === bodyStart && count >= bodyRows) {
      sh.getRange(bodyStart, 1, 1, sh.getMaxColumns()).clearContent();
    }
    return;
  }
  let row = startRow;
  while (safeCount > 0) {
    const n = Math.min(BLOCK, safeCount);
    sh.deleteRows(row, n);
    safeCount -= n;
  }
  if (startRow === bodyStart && count >= bodyRows) {
    sh.getRange(bodyStart, 1, 1, sh.getMaxColumns()).clearContent();
  }
}


/** Batch/contiguous prune for "older than N days" assuming chronological appends. */
function pruneOldRows(sheet, retentionDays, dateCol /* 1-based */) {
  if (!retentionDays || !sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - retentionDays);

  const dates = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues().flat();

  let boundary = -1;
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i];
    if (!(d instanceof Date)) continue;
    if (d < cutoff) boundary = i;
    else break;
  }

  if (boundary >= 0) {
    const rowsToDelete = boundary + 1;
    _deleteInBlocks_(sheet, /*startRow=*/ 2, /*count=*/ rowsToDelete);
    LoggerEx && LoggerEx.log('Prices: pruned old rows (<= cutoff):', rowsToDelete);
  }
}

/* ---------------------- [NEW] 24-Hour State Check ---------------------- */

/**
 * NEW: Performs the 24-hour state check.
 * 1. Runs a "Light Prune" (pruneOldRows) on the live data.
 * 2. Resets the main Fuzz worker job state.
 *
 * This function should be put on a daily trigger.
 */
function dailyJobReset() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('DailyReset') : console);
  LOG.info("Starting 24-hour state check and job reset...");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig(); // Corrected config call
  const sheetName = FUZZ_SHEET_FINAL; // Use final sheet name for prune
  // We use getOrCreateSheet here to ensure the sheet exists and has headers,
  // but it will NOT clear contents due to the new logic in _updateFuzzMarketDataWorker.
  const sheet = getOrCreateSheet(ss, sheetName, FUZZ_SHEET_HEADERS);

  if (!sheet) {
    LOG.error(`Sheet not found: ${sheetName}. Skipping light prune.`);
  } else {
    try {
      // 1. Run the "Light Prune"
      const retentionDays = cfg.PriceRetentionDays || 1; // Use PriceRetentionDays from config
      LOG.info(`Running light prune (pruneOldRows) for ${retentionDays} day(s) on ${sheetName}...`);
      pruneOldRows(sheet, retentionDays, 1); // 1 = date column
      LOG.info("Light prune complete.");
    } catch (e) {
      LOG.error(`Light prune failed: ${e.message}`);
    }
  }

  try {
    // 2. Reset the Fuzz worker's "Start Index"
    LOG.info("Resetting Fuzz Market Data Job State...");
    _resetFuzzMarketDataJobState(new Error("Daily 24-hour reset"));
    LOG.info("Fuzz job state reset complete.");
  } catch (e) {
    LOG.error(`Fuzz job reset failed: ${e.message}`);
  }
}


/* ---------------------- [REWRITTEN] Heavy Prune (Stateful) ---------------------- */

/**
 * REWRITTEN: This is now the "Starter" function for the stateful heavy prune.
 */
function dailyHeavyPrune_Prices() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('HeavyPrune') : console);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();

  // Check if it's already running
  const currentState = SCRIPT_PROP.getProperty(PRUNE_PROP_STEP);
  // REFACTORED: Use strings directly, remove STATE_FLAGS dependency
  if (currentState && currentState !== 'COMPLETE' && currentState !== "NEW_RUN") {
    LOG.warn(`Heavy Prune is already running (State: ${currentState}). Skipping new start.`);
    // Re-schedule the worker just in case the trigger was lost
    scheduleOneTimeTrigger('_heavyPruneWorker', 5000);
    return;
  }

  LOG.info("Starting new Heavy Prune cycle.");
  // REFACTORED: Use strings directly
  SCRIPT_PROP.setProperty(PRUNE_PROP_STEP, "NEW_RUN");

  // Use executeWithTryLock
  const result = executeWithTryLock(_heavyPruneWorker, '_heavyPruneWorker');
  if (result === null) {
    LOG.warn("Heavy Prune start skipped by ScriptLock. Another process is running.");
  }
}

/**
 * NEW: The "Processing While Loop" worker for the heavy prune.
 * Note: This function remains swap-based for safety/atomicity during the deduplication process.
 */
function _heavyPruneWorker() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('PruneWorker') : console); // <-- Uses LOG
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const START_TIME = Date.now();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig(); // Corrected config call
  const sourceSheetName = FUZZ_SHEET_FINAL; // Market Prices is now the source
  const tempSheetName = PRUNE_SHEET_TEMP;

  // REFACTORED: Use strings directly
  let currentState = SCRIPT_PROP.getProperty(PRUNE_PROP_STEP) || "NEW_RUN";
  LOG.info(`Starting worker. Current State: ${currentState}`); // <-- Uses LOG

  try {
    // --- State: NEW_RUN (Start) ---
    // REFACTORED: Use strings directly
    if (currentState === "NEW_RUN") {
      LOG.info(`State: NEW_RUN. Preparing prune temp sheet.`);
      const docLock = LockService.getDocumentLock();
      if (docLock.tryLock(FUZZ_DOC_LOCK_TIMEOUT)) {
        try {
          // [FIX] Aggressively delete old sheets to free space BEFORE creating new ones
          const oldSheetName = FUZZ_SHEET_FINAL + "_Prune_Old";
          const oldSheet = ss.getSheetByName(oldSheetName);
          if (oldSheet) {
            LOG.info(`Deleting old backup sheet to free space: ${oldSheetName}`);
            ss.deleteSheet(oldSheet);
          }

          const existingTemp = ss.getSheetByName(tempSheetName);
          if (existingTemp) {
            LOG.info(`Deleting stale temp sheet to free space: ${tempSheetName}`);
            ss.deleteSheet(existingTemp);
          }
          SpreadsheetApp.flush(); // Ensure deletion is committed before insertion
          // [END FIX]

          const tempSheet = getOrCreateSheet(ss, tempSheetName, FUZZ_SHEET_HEADERS);
          // (No need to clear content since we just deleted and recreated it)

          tempSheet.hideSheet();
          SpreadsheetApp.flush();

          SCRIPT_PROP.setProperty(PRUNE_PROP_READ_ROW, '2'); // Data starts row 2

          currentState = "PROCESSING";
          SCRIPT_PROP.setProperty(PRUNE_PROP_STEP, currentState);
          LOG.info(`Temp sheet '${tempSheetName}' prepared. Transitioning to ${currentState}.`);
        } finally {
          docLock.releaseLock();
        }
      } else {
        LOG.warn(`Document Lock busy during prune setup. Rescheduling.`);
        scheduleOneTimeTrigger('_heavyPruneWorker', FUZZ_RESCHEDULE_MS);
        return;
      }
    } // --- End NEW_RUN ---

    // --- State: PROCESSING (Processing While Loop) ---
    // REFACTORED: Use strings directly
    if (currentState === "PROCESSING") {
      LOG.info(`State: PROCESSING. Reading/deduping batches.`); // <-- Uses LOG

      const sourceSheet = ss.getSheetByName(sourceSheetName);
      const tempSheet = ss.getSheetByName(tempSheetName);
      if (!sourceSheet || !tempSheet) {
        throw new Error("Missing source or temp sheet during prune processing.");
      }

      let readRow = parseInt(SCRIPT_PROP.getProperty(PRUNE_PROP_READ_ROW) || '2');
      const lastRow = sourceSheet.getLastRow();

      // --- Find Header Indices ---
      const header = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
      const lower = header.map(h => String(h).trim().toLowerCase());
      const find = (name) => lower.findIndex(h => h === name);
      const DATE = find('date'), TYPE = find('type_id'), MID = find('market_id'), MTP = find('market_type');

      // --- REFACTORED: Find price columns for filtering ---
      const MIN_SELL = find('min_sell');
      const MAX_BUY = find('max_buy');

      if (DATE < 0 || TYPE < 0 || MID < 0 || MTP < 0 || MIN_SELL < 0 || MAX_BUY < 0) {
        throw new Error('Missing required columns (date/type_id/market_id/market_type/min_sell/max_buy) in source sheet.');
      }

      // --- Processing Loop ---
      while (readRow <= lastRow) {
        // --- 1. Time Limit Check ---
        if (Date.now() - START_TIME > FUZZ_TIME_LIMIT_MS) {
          SCRIPT_PROP.setProperty(PRUNE_PROP_READ_ROW, readRow.toString());
          scheduleOneTimeTrigger('_heavyPruneWorker', FUZZ_RESCHEDULE_MS);
          LOG.warn(`Time limit hit. Saved state. Rescheduled. Next read row: ${readRow}`); // <-- Uses LOG
          return;
        }

        // --- 2. Read Batch ---
        const rowsToRead = Math.min(PRUNE_BATCH_SIZE, lastRow - readRow + 1);
        if (rowsToRead <= 0) break;

        LOG.info(`Reading ${rowsToRead} rows from ${sourceSheetName} (starting row ${readRow})...`); // <-- Uses LOG
        const data = sourceSheet.getRange(readRow, 1, rowsToRead, header.length).getValues();

        // --- 3. Process Batch (Retention, Bucket, Dedupe) ---
        const retentionDays = cfg.PriceRetentionDays || 1; // Use PriceRetentionDays from config
        const bucketMinutes = cfg.BucketMinutes || 20; // Use BucketMinutes from config

        const cutoff = new Date(Date.now() - retentionDays * 86400000);
        const msPerBucket = bucketMinutes * 60 * 1000;
        const keep = new Map(); // Keep latest record *within this batch*

        for (let i = 0; i < data.length; i++) {
          const r = data[i];
          const d = r[DATE];
          if (!(d instanceof Date) || d < cutoff) continue; // Retention filter

          // --- REFACTORED: VALID PRICE CHECK ---
          // A row is invalid *only if* BOTH min_sell AND max_buy are non-positive.
          const validMinSell = (r[MIN_SELL] != null && r[MIN_SELL] !== "" && Number(r[MIN_SELL]) > 0);
          const validMaxBuy = (r[MAX_BUY] != null && r[MAX_BUY] !== "" && Number(r[MAX_BUY]) > 0);

          // If NEITHER price is valid (min_sell is invalid AND max_buy is invalid), skip the row.
          if (!validMinSell && !validMaxBuy) {
            continue; // Skip row as it has no valid data
          }
          // --- END REFACTOR ---

          const bucket = Math.floor(d.getTime() / msPerBucket);
          const key = bucket + '|' + r[TYPE] + '|' + r[MID] + '|' + r[MTP];

          const prev = keep.get(key);
          if (!prev || (r[DATE] > prev[DATE])) {
            keep.set(key, r); // Dedupe filter
          }
        }

        const rowsToWrite = Array.from(keep.values());

        // --- 4. Write Batch (Throttle Sheet Writes / Retrigger) ---
        if (rowsToWrite.length > 0) {
          const docLock = LockService.getDocumentLock();
          if (docLock.tryLock(FUZZ_DOC_LOCK_TIMEOUT)) {
            try {
              tempSheet.getRange(tempSheet.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
              LOG.info(`Appended ${rowsToWrite.length} deduped rows to ${tempSheetName}.`); // <-- Uses LOG
            } finally {
              docLock.releaseLock();
            }
          } else {
            // RETRIGGER ON WRITE FAILURE
            LOG.warn(`Document Lock busy for prune write. Rescheduling (will re-process batch).`); // <-- Uses LOG
            scheduleOneTimeTrigger('_heavyPruneWorker', FUZZ_RESCHEDULE_MS);
            return;
          }
        }

        // --- 5. Advance Index ---
        readRow += rowsToRead;
        SCRIPT_PROP.setProperty(PRUNE_PROP_READ_ROW, readRow.toString());

      } // --- End while loop ---

      // --- Post-Loop Check ---
      if (readRow > lastRow) {
        LOG.info("All source rows processed. Transitioning to FINALIZING."); // <-- Uses LOG
        // REFACTORED: Use strings directly
        currentState = "FINALIZING";
        SCRIPT_PROP.setProperty(PRUNE_PROP_STEP, currentState);
        scheduleOneTimeTrigger('_finalizePrune', 1000); // 1 sec delay
      }
    } // --- End PROCESSING ---

  } catch (e) {
    LOG.error(`Unhandled error in prune worker: ${e.message}\nStack: ${e.stack}`); // <-- Uses LOG
    // Reset prune state on error
    SCRIPT_PROP.deleteProperty(PRUNE_PROP_STEP);
    SCRIPT_PROP.deleteProperty(PRUNE_PROP_READ_ROW);
  }
}

/**
 * NEW: The "Finalize" function for the heavy prune.
 * This runs the *second* deduplication (across all batches) and performs the atomic swap.
 */
function _finalizePrune() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('PruneFinalizer') : console);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();

  // REFACTORED: Use strings directly
  if (SCRIPT_PROP.getProperty(PRUNE_PROP_STEP) !== "FINALIZING") {
    LOG.warn(`Finalizer called in incorrect state (${SCRIPT_PROP.getProperty(PRUNE_PROP_STEP)}). Aborting.`);
    return;
  }

  LOG.info("Starting finalization: Secondary deduplication and atomic swap.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig(); // Corrected config call
  const tempSheetName = PRUNE_SHEET_TEMP;
  const finalSheetName = FUZZ_SHEET_FINAL; // Market Prices is the final sheet
  const oldSheetName = finalSheetName + "_Prune_Old"; // Unique name for clarity

  const docLock = LockService.getDocumentLock();
  try {
    if (docLock.tryLock(30000)) { // Wait up to 30s
      try {
        const tempSheet = ss.getSheetByName(tempSheetName);
        if (!tempSheet || tempSheet.getLastRow() <= 1) {
          throw new Error(`Prune temp sheet '${tempSheetName}' is missing or empty! Cannot finalize.`);
        }

        // --- 1. Secondary Deduplication (in memory) ---
        LOG.info("Reading temp sheet for final deduplication...");
        const data = tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, tempSheet.getLastColumn()).getValues();

        const bucketMinutes = cfg.BucketMinutes || 20; // Use BucketMinutes from config
        const maxRows = cfg.PricesMaxRows || 100000; // Use PricesMaxRows from config

        const msPerBucket = bucketMinutes * 60 * 1000;
        const keep = new Map();

        const header = tempSheet.getRange(1, 1, 1, tempSheet.getLastColumn()).getValues()[0];
        const lower = header.map(h => String(h).trim().toLowerCase());
        const find = (name) => lower.findIndex(h => h === name);
        const DATE = find('date'), TYPE = find('type_id'), MID = find('market_id'), MTP = find('market_type');

        // --- REFACTORED: Find price columns for filtering ---
        const MIN_SELL = find('min_sell');
        const MAX_BUY = find('max_buy');

        for (let i = 0; i < data.length; i++) {
          const r = data[i];

          // --- REFACTORED: VALID PRICE CHECK ---
          // (This check is redundant if _heavyPruneWorker worked, but good for safety)
          const validMinSell = (r[MIN_SELL] != null && r[MIN_SELL] !== "" && Number(r[MIN_SELL]) > 0);
          const validMaxBuy = (r[MAX_BUY] != null && r[MAX_BUY] !== "" && Number(r[MAX_BUY]) > 0);

          // If NEITHER price is valid, skip the row.
          if (!validMinSell && !validMaxBuy) {
            continue; // Skip this row
          }
          // --- END REFACTOR ---

          const d = r[DATE];
          if (!(d instanceof Date)) continue;

          const bucket = Math.floor(d.getTime() / msPerBucket);
          const key = bucket + '|' + r[TYPE] + '|' + r[MID] + '|' + r[MTP];

          const prev = keep.get(key);
          if (!prev || (r[DATE] > prev[DATE])) {
            keep.set(key, r);
          }
        }

        let deduped = Array.from(keep.values());
        LOG.info(`Final deduplication complete. Kept ${deduped.length} rows.`);

        // --- 2. Cap Rows ---
        if (deduped.length > maxRows) {
          deduped.sort((a, b) => a[DATE] - b[DATE]); // Sort by date ascending
          deduped = deduped.slice(deduped.length - maxRows); // Keep the newest rows
          LOG.info(`Capped rows to ${deduped.length} (max: ${maxRows}).`);
        }

        // --- 3. Rewrite Temp Sheet ---
        tempSheet.clearContents(); // Clear everything
        tempSheet.getRange(1, 1, 1, header.length).setValues([header]); // Set header
        if (deduped.length > 0) {
          tempSheet.getRange(2, 1, deduped.length, header.length).setValues(deduped);
        }
        _trimTrailing_(tempSheet);
        SpreadsheetApp.flush();
        LOG.info("Final data written to temp sheet."); // <-- CORRECTED LOG VARIABLE

        // --- 4. Atomic Swap (Prune Edition) ---
        const finalSheet = ss.getSheetByName(finalSheetName);
        const oldSheet = ss.getSheetByName(oldSheetName);

        if (oldSheet) ss.deleteSheet(oldSheet);
        if (finalSheet) finalSheet.setName(oldSheetName);
        tempSheet.setName(finalSheetName);
        tempSheet.showSheet(); // Use tempSheet handle which is now the final sheet

        // --- THIS WAS THE FIX YOU MADE ---
        SpreadsheetApp.flush(); // <-- You fixed this! (Was Spreadfuzz.flush())
        // --- END OF FIX ---

        LOG.info("Atomic sheet swap successful."); // <-- CORRECTED LOG VARIABLE

        // --- 5. Reset Prune Job State ---
        SCRIPT_PROP.deleteProperty(PRUNE_PROP_STEP);
        SCRIPT_PROP.deleteProperty(PRUNE_PROP_READ_ROW);
        LOG.info("Heavy Prune job state reset complete."); // <-- CORRECTED LOG VARIABLE

      } catch (swapError) {
        LOG.error(`CRITICAL error during prune swap: ${swapError.message}. State NOT reset.`);
        scheduleOneTimeTrigger('_finalizePrune', 60000); // Retry in 1 min
        throw swapError;
      } finally {
        docLock.releaseLock();
      }
    } else {
      LOG.warn("Document Lock busy during finalization. Rescheduling finalizer.");
      scheduleOneTimeTrigger('_finalizePrune', FUZZ_RESCHEDULE_MS);
    }
  } catch (e) {
    LOG.error(`Error in finalizer lock acquisition: ${e.message}`);
    scheduleOneTimeTrigger('_finalizePrune', 60000);
  }
}

/**
 * NEW: Resets the state of the Heavy Prune job.
 */
function _resetHeavyPruneJobState(error) {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('HeavyPrune') : console);
  LOG.warn(`RESETTING Heavy Prune Job State. Reason: ${error ? error.message : 'Manual'}`);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  try {
    SCRIPT_PROP.deleteProperty(PRUNE_PROP_STEP);
    SCRIPT_PROP.deleteProperty(PRUNE_PROP_READ_ROW);
    deleteTriggersByName('_heavyPruneWorker');
    deleteTriggersByName('_finalizePrune');
  } catch (propError) {
    LOG.error(`Error deleting script properties: ${propError.message}`);
  }
  LOG.info("Heavy Prune job state reset complete.");
}

/**
 * ONE-TIME UTILITY: Run this manually to free up cell space.
 * Deletes old backup/temp sheets that may be clogging the workbook.
 */
function emergencyCleanup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = [
    'Market Prices_Prune_Old',  // Old backup from previous swaps
    'Market_Prices_Prune_Temp'  // Stale temp sheet from failed runs
  ];

  sheetsToDelete.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      console.log(`Deleting sheet to free space: ${name}`);
      ss.deleteSheet(sheet);
    } else {
      console.log(`Sheet not found (clean): ${name}`);
    }
  });

  // Also trim empty rows from the main sheet
  const mainSheet = ss.getSheetByName('Market Prices');
  if (mainSheet) {
    console.log('Trimming trailing empty rows from Market Prices...');
    _trimTrailing_(mainSheet);
  }
}

/**
 * PATH B: The 24-Hour Expiry "Crowbar"
 * Keeps the sheet from exploding when BigQuery is offline.
 * Move to MarketFetchetr
 */
function pruneSheetToRolling24(ss) {
  if(!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FUZZ_SHEET_FINAL);
  if (!sheet) return;

  const now = new Date().getTime();
  const oneDayAgo = now - (24 * 60 * 60 * 1000);
  const data = sheet.getDataRange().getValues();
  
  // Assuming Timestamp is in Column A (Index 0)
  const rowsToKeep = data.filter((row, index) => {
    if (index === 0) return true; // Keep Header
    const rowTime = new Date(row[0]).getTime();
    return rowTime > oneDayAgo;
  });

  sheet.clearContents();
  if (rowsToKeep.length > 0) {
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
  }
  console.log(`[PRUNE] Kept ${rowsToKeep.length} rows (Last 24h).`);
}

/**
 * THE LIVE WIRE RECALL: Surgical BigQuery Puller.
 * Refined to match the E7:K standard and ConfigHandler logic.
 */
function FUZ_refreshPriceInterface(ss, targetSheetName) {
  const cfg = getConfig(); // Pull IDs from your central config
  const projectId = cfg.BQ_PROJECT_ID || 'tenacious-tiger-345318';
  
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) return console.warn(`[SKIP] Sheet "${targetSheetName}" not found.`);

  // --- 1. THE GATE (C4:D4) ---
  const settings = sheet.getRange("C4:D4").getValues()[0];
  const marketId = parseInt(settings[0]);
  const marketType = String(settings[1]).trim();

  if (isNaN(marketId) || marketType.toLowerCase().includes("loading")) {
    sheet.getRange("E4").setValue(`⚠️ Waiting for IMPORTRANGE...`);
    return; 
  }

  // --- 2. THE RECALL SQL (Optimized for E7:K) ---
  // Calculates 24h Medians + Current Prices + % Changes in the Vault
  const sql = `
    SELECT 
      type_id, 
      ROUND(AVG(median_buy), 2) as avg_buy, 
      ROUND(AVG(median_sell), 2) as avg_sell,
      ARRAY_AGG(median_buy ORDER BY date DESC LIMIT 1)[OFFSET(0)] as cur_buy,
      ARRAY_AGG(median_sell ORDER BY date DESC LIMIT 1)[OFFSET(0)] as cur_sell
    FROM \`${projectId}.market_data.market_prices_staged\`
    WHERE market_id = ${marketId}
      AND LOWER(market_type) = LOWER('${marketType}')
      AND date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 24 HOUR)
    GROUP BY type_id ORDER BY type_id ASC
  `;

  try {
    console.log(`[${targetSheetName}] Vault Recall Started...`);
    const queryResults = BigQuery.Jobs.query({query: sql, useLegacySql: false}, projectId);
    
    const data = queryResults.rows ? queryResults.rows.map(row => {
      const v = row.f.map(field => field.v);
      const avgB = parseFloat(v[1]), avgS = parseFloat(v[2]), curB = parseFloat(v[3]), curS = parseFloat(v[4]);
      // Calculate % Change here to keep the sheet light
      const chgS = avgS > 0 ? (curS - avgS) / avgS : 0;
      const chgB = avgB > 0 ? (curB - avgB) / avgB : 0;
      return [v[0], avgB, avgS, curB, curS, chgS, chgB];
    }) : [];

    if (data.length === 0) throw new Error("Vault is Empty (Waiting for next fetch)");
    
    // --- 3. THE RESPONSE (Write to E7:K) ---
    FUZ_writeToInterfaceSheet(sheet, data, "✅ Vault Synced");

  } catch (err) {
    console.warn(`[${targetSheetName}] Sync Failed: ${err.message}`);
    sheet.getRange("E4").setValue(`⚠️ Sync Error: ${err.message}`);
  }
}

/**
 * HELPER: Surgical write to the E7:K range.
 */
function FUZ_writeToInterfaceSheet(sheet, data, statusPrefix) {
  const HEADERS = [["type_id_filtered", "Median Buy", "Median Sell", "Current Buy", "Current Sell", "Sell Change", "Buy Change"]];
  
  // Clear Column E through K starting at row 8
  const lastRow = sheet.getLastRow();
  if (lastRow >= 8) sheet.getRange(8, 5, lastRow - 7, 7).clearContent();
  
  // Set Headers at Row 7
  sheet.getRange(7, 5, 1, 7).setValues(HEADERS);
  
  // Write Data at Row 8
  if (data.length > 0) {
    sheet.getRange(8, 5, data.length, 7).setValues(data);
    // Formatting: J-K as Percentages, F-I as ISK
    sheet.getRange(8, 10, data.length, 2).setNumberFormat('0.00%');
    sheet.getRange(8, 6, data.length, 4).setNumberFormat('#,##0.00 "ISK"');
  }
  
  sheet.getRange("E4").setValue(`${statusPrefix}: ${new Date().toLocaleTimeString()}`);
}

/**
 * THE GENERAL: Now with Error-Gating for IMPORTRANGE cells.
 * Targets: 'filtered prices', 'Mineral Supply Prices', 'T1 Supply Prices'.
 */
function FUZ_publishPriceInterfaces(ss) {
  if(!ss) ss = SpreadsheetApp.getActive();
  const cfg = getConfig();
  const isVaultOk = (cfg.BQ_ENABLED === true);
  
  const priceSheets = ['filtered prices', 'Mineral Supply Prices', 'T1 Supply Prices'];
  
  priceSheets.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return console.warn(`[SKIP] Interface "${sheetName}" not found.`);

    // --- 1. THE GATING CHECK (C4, D4) ---
    const mktIdRaw = sh.getRange("C4").getValue();
    const mktTypeRaw = sh.getRange("D4").getValue();

    // Helper: Detect Errors or Loading states from IMPORTRANGE
    const isError = (val) => {
      const s = String(val);
      return s.indexOf('#') === 0 || s.toLowerCase().includes('loading') || s.trim() === '';
    };

    if (isError(mktIdRaw) || isError(mktTypeRaw)) {
      console.warn(`[${sheetName}] Gated: Market Settings contain Error/Loading.`);
      sh.getRange("E4").setValue("!! ERROR: Market Settings Loading/Broken !!");
      return; // Skip this sheet and move to the next
    }

    const marketId = parseInt(mktIdRaw);
    const marketType = String(mktTypeRaw).trim();

    if (isNaN(marketId)) {
      sh.getRange("E4").setValue("!! ERROR: Market ID is not a number !!");
      return;
    }

    // --- 2. READ ITEM REQUESTS (Column B) ---
    const lastRow = sh.getLastRow();
    if (lastRow < 8) return console.warn(`[${sheetName}] No items found in Col B.`);
    
    const itemIds = sh.getRange(8, 2, lastRow - 7, 1).getValues()
                      .flat()
                      .filter(id => !isNaN(parseInt(id)) && id > 0);

    if (itemIds.length === 0) return;

    console.log(`[${sheetName}] Requesting ${itemIds.length} items for ${marketType} ${marketId}`);

    // --- 3. THE RECALL ---
    let data = [];
    try {
      if (isVaultOk) {
        data = getPricesFromVault_(cfg.BQ_PROJECT_ID, marketId, marketType, itemIds);
      } else {
        data = getPricesFromLocalBuffer_(ss.getSheetByName('Market Prices'), marketId, itemIds);
      }

      // --- 4. THE RESPONSE (Write to E7:K) ---
      writeToPriceInterface_(sh, data, isVaultOk);

    } catch (e) {
      console.error(`[${sheetName}] Sync Failed: ${e.message}`);
      sh.getRange("E4").setValue("!! Sync Error: Check Logs !!");
    }
  });
}

/**
 * BUFFER: Pulls from the local 'Market Prices' sheet if the Vault is locked.
 */
function getPricesFromLocalBuffer_(sh) {
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  // Filter for the last 24h and map to your [id, buy_med, sell_med, buy_curr, sell_curr] format
  // (Assuming your local sheet schema matches FUZZ_SHEET_HEADERS)
  return data.slice(1).map(r => [r[3], r[7], r[6], r[5], r[4]]);
}

/**
 * RECALL: Targeted SQL query for specific items.
 */
function getPricesFromVault_(projectId, mktId, mktType, ids) {
  const idString = ids.join(',');
  const sql = `
    SELECT 
      type_id, 
      ROUND(AVG(median_buy), 2) as avg_buy, 
      ROUND(AVG(median_sell), 2) as avg_sell,
      ARRAY_AGG(median_buy ORDER BY date DESC LIMIT 1)[OFFSET(0)] as curr_buy,
      ARRAY_AGG(median_sell ORDER BY date DESC LIMIT 1)[OFFSET(0)] as curr_sell
    FROM \`${projectId}.market_data.market_prices_staged\`
    WHERE market_id = ${mktId}
      AND LOWER(market_type) = LOWER('${mktType}')
      AND type_id IN (${idString})
      AND date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 24 HOUR)
    GROUP BY type_id ORDER BY type_id ASC
  `;
  
  try {
    const queryResults = BigQuery.Jobs.query({query: sql, useLegacySql: false}, projectId);
    if (!queryResults.rows) return [];
    
    return queryResults.rows.map(row => {
      const v = row.f.map(field => field.v);
      const avgB = parseFloat(v[1]), avgS = parseFloat(v[2]), curB = parseFloat(v[3]), curS = parseFloat(v[4]);
      // Calculate % Change: (Current - 24h Median) / 24h Median
      const chgS = avgS > 0 ? (curS - avgS) / avgS : 0;
      const chgB = avgB > 0 ? (curB - avgB) / avgB : 0;
      return [v[0], avgB, avgS, curB, curS, chgS, chgB];
    });
  } catch (e) {
    console.error("Vault Query Failed: " + e.message);
    return [];
  }
}

/**
 * RESPONSE: Surgical write to the E7:K range.
 */
function writeToPriceInterface_(sheet, data, isVault) {
  const HEADERS = [["type_id_filtered", "Median Buy", "Median Sell", "Current Buy", "Current Sell", "Sell Change", "Buy Change"]];
  
  // Clear only the output range
  const lastRow = sheet.getLastRow();
  if (lastRow >= 8) sheet.getRange(8, 5, lastRow - 7, 7).clearContent();
  
  // Write Headers (E7:K7)
  sheet.getRange(7, 5, 1, 7).setValues(HEADERS);
  
  if (data.length > 0) {
    sheet.getRange(8, 5, data.length, 7).setValues(data);
    
    // Formatting for Changes (J and K)
    sheet.getRange(8, 10, data.length, 2).setNumberFormat('0.00%');
    // Formatting for ISK
    sheet.getRange(8, 6, data.length, 4).setNumberFormat('#,##0.00 "ISK"');
  }
  
  const status = isVault ? "✅ Vault" : "⚠️ Buffer";
  sheet.getRange("E4").setValue(`${status} Synced: ${new Date().toLocaleTimeString()}`);
}