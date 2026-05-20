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
function updateFuzzMarketDataSheet(itemSource, ctx) {
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
  // You MUST use 'return' here so the True/False/Null passes back to the 'success' variable
  const success = executeWithTryLock(() => {
    return _updateFuzzMarketDataWorker(itemSource, ctx);
  }, 'updateFuzzMarketDataSheet');

  // --- 3. UPDATE LEASE (The State Machine Logic) ---
  if (success === true) {
    // STATE 1: 100% Pass Complete. 
    console.log(`${LOG_HEADER} 100% Pass Complete. Setting 30-minute Lease.`);
    PROPS.setProperty('LAST_FUZZ_FETCH', now.toString());
  }
  else if (success === false) {
    // STATE 2: Partial Pass / Timeout. 
    console.log(`${LOG_HEADER} Partial Pass. Clearing lease so next pulse resumes immediately.`);
    // VIP BYPASS: Deleting the lease ensures the next 3-minute trigger isn't blocked
    PROPS.deleteProperty('LAST_FUZZ_FETCH');
  }
  else {
    // STATE 3: Lock Busy (success is null or undefined)
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
 * Updated to use Type-Safe Context and tactical burner-sheet recovery.
 */
function _updateFuzzMarketDataWorker(itemSource, ctx) {
  const LOG = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('FuzzWorker') : console;

  // Handshake destructuring
  const { ss, props, cfg } = ctx;

  // --- 1. LOAD ITEMS (The Dynamic Map) ---
  const sourceName = itemSource || "Item List Back End";
  const itemSheet = ss.getSheetByName(sourceName);

  if (!itemSheet) {
    LOG.error(`Critical: '${sourceName}' sheet missing.`);
    return false;
  }

  const itemData = itemSheet.getDataRange().getValues();
  const typeIdBatch = [];

  for (let i = 1; i < itemData.length; i++) {
    const id = parseInt(itemData[i][0]);
    if (!isNaN(id)) typeIdBatch.push(id);
  }

  if (typeIdBatch.length === 0) {
    LOG.warn(`No Type IDs found in ${sourceName}.`);
    return false;
  }

  // --- 2. LOAD MARKETS ---
  const hubSheet = ss.getSheetByName("Market Settings");
  let marketsToProcess = [];

  if (hubSheet) {
    const hubData = hubSheet.getDataRange().getValues();
    for (let i = 1; i < hubData.length; i++) {
      const sysId = parseInt(hubData[i][4]);
      const statId = parseInt(hubData[i][3]);
      if (!isNaN(sysId)) marketsToProcess.push({ market_id: sysId, market_type: 'system' });
      else if (!isNaN(statId)) marketsToProcess.push({ market_id: statId, market_type: 'station' });
    }
  }

  if (marketsToProcess.length === 0) {
    marketsToProcess.push({ market_id: 60003760, market_type: 'station' }); // Jita Fallback
  }

  // --- 3. EXECUTE LOOP ---
  LOG.info(`Starting Cycle: ${typeIdBatch.length} items across ${marketsToProcess.length} markets.`);
  const startTime = Date.now();

  let startIndex = parseInt(props.getProperty('fuzz_hub_bookmark')) || 0;
  LOG.info(`Resuming from Hub Index: ${startIndex}`);

  for (let i = startIndex; i < marketsToProcess.length; i++) {
    const activeConfig = marketsToProcess[i];

    // 3-minute safety limit
    if (Date.now() - startTime > 180000) {
      LOG.warn(`[TIMEOUT] Bookmarking at Hub Index: ${i}`);
      props.setProperty('fuzz_hub_bookmark', String(i));
      props.setProperty('fuzz_pass_complete', 'false');
      return false;
    }

    try {
      const marketData = getMarketPrices(typeIdBatch, activeConfig.market_id, activeConfig.market_type);

      if (!marketData || Object.keys(marketData).length === 0) continue;

      const rowsToStream = [];
      const timestamp = new Date();

      typeIdBatch.forEach(typeId => {
        const data = marketData[typeId];
        if (data && ((data.buy?.orderCount || 0) + (data.sell?.orderCount || 0) > 0)) {
          rowsToStream.push([
            timestamp,
            String(activeConfig.market_id),
            activeConfig.market_type,
            parseInt(typeId, 10),
            parseFloat(data.sell?.min) || 0,
            parseFloat(data.buy?.max) || 0,
            parseFloat(data.sell?.median) || 0,
            parseFloat(data.buy?.median) || 0
          ]);
        }
      });

      // --- STAGE TO LOCAL TANK (Tactical & Sparring-Ready) ---
      if (rowsToStream.length > 0) {
        props.setProperty('fuzz_job_active', 'true');

        // TACTICAL: Direct reference or instant rebuild using Context ss handle
        const stagingSheet = ss.getSheetByName("Market Prices") || rebuildMarketPricesSheet(ctx);

        const lastRow = stagingSheet.getLastRow();
        const chunkSize = 4000;

        for (let j = 0; j < rowsToStream.length; j += chunkSize) {
          const chunk = rowsToStream.slice(j, j + chunkSize);
          stagingSheet.getRange(lastRow + 1 + j, 1, chunk.length, 8).setValues(chunk);
          // NO FLUSH. Let the V8 engine queue the writes natively.
        }
        props.setProperty('fuzz_job_active', 'false');
      }

    } catch (e) {
      LOG.error(`Error on Hub ${activeConfig.market_id}: ${e.message}`);
      props.setProperty('fuzz_job_active', 'false');
    }
  }

  // --- 4. COMPLETION SIGNAL ---
  LOG.info(`100% Market Pass Complete.`);
  props.setProperty('fuzz_pass_complete', 'true');
  props.setProperty('fuzz_hub_bookmark', '0');

  // --- 5. TACTICAL STATIC DATA PULSE ---
  const utilitySheet = ss.getSheetByName("Utility");
  if (utilitySheet) {
    utilitySheet.getRange("B3").setValue(0);
    // TACTICAL FLUSH: Mandatory for triggering ImportRange formula resets
    SpreadsheetApp.flush();
    utilitySheet.getRange("B3").setValue(1);
    LOG.info(`Static Data Pulse Fired.`);
  }

  return true;
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
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('PruneWorker') : console);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const START_TIME = Date.now();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const sourceSheetName = FUZZ_SHEET_FINAL;
  const tempSheetName = PRUNE_SHEET_TEMP;

  let currentState = SCRIPT_PROP.getProperty(PRUNE_PROP_STEP) || "NEW_RUN";
  LOG.info(`Starting worker. Current State: ${currentState}`);

  try {
    // --- State: NEW_RUN (Start) ---
    if (currentState === "NEW_RUN") {
      LOG.info(`State: NEW_RUN. Preparing prune temp sheet.`);
      const docLock = LockService.getDocumentLock();
      if (docLock.tryLock(FUZZ_DOC_LOCK_TIMEOUT)) {
        try {
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
          SpreadsheetApp.flush();

          const tempSheet = getOrCreateSheet(ss, tempSheetName, FUZZ_SHEET_HEADERS);
          tempSheet.hideSheet();
          SpreadsheetApp.flush();

          SCRIPT_PROP.setProperty(PRUNE_PROP_READ_ROW, '2');

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
    }

    // --- State: PROCESSING (Processing While Loop) ---
    if (currentState === "PROCESSING") {
      LOG.info(`State: PROCESSING. Reading/deduping batches.`);

      const sourceSheet = ss.getSheetByName(sourceSheetName);
      const tempSheet = ss.getSheetByName(tempSheetName);
      if (!sourceSheet || !tempSheet) {
        throw new Error("Missing source or temp sheet during prune processing.");
      }

      let readRow = parseInt(SCRIPT_PROP.getProperty(PRUNE_PROP_READ_ROW) || '2');
      const lastRow = sourceSheet.getLastRow();

      const header = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
      const lower = header.map(h => String(h).trim().toLowerCase());
      const find = (name) => lower.findIndex(h => h === name);
      const DATE = find('date'), TYPE = find('type_id'), MID = find('market_id'), MTP = find('market_type');
      const MIN_SELL = find('min_sell');
      const MAX_BUY = find('max_buy');

      if (DATE < 0 || TYPE < 0 || MID < 0 || MTP < 0 || MIN_SELL < 0 || MAX_BUY < 0) {
        throw new Error('Missing required columns (date/type_id/market_id/market_type/min_sell/max_buy) in source sheet.');
      }

      while (readRow <= lastRow) {
        if (Date.now() - START_TIME > FUZZ_TIME_LIMIT_MS) {
          SCRIPT_PROP.setProperty(PRUNE_PROP_READ_ROW, readRow.toString());
          scheduleOneTimeTrigger('_heavyPruneWorker', FUZZ_RESCHEDULE_MS);
          LOG.warn(`Time limit hit. Saved state. Rescheduled. Next read row: ${readRow}`);
          return;
        }

        const rowsToRead = Math.min(PRUNE_BATCH_SIZE, lastRow - readRow + 1);
        if (rowsToRead <= 0) break;

        LOG.info(`Reading ${rowsToRead} rows from ${sourceSheetName} (starting row ${readRow})...`);
        const data = sourceSheet.getRange(readRow, 1, rowsToRead, header.length).getValues();

        const retentionDays = cfg.PriceRetentionDays || 1;
        const bucketMinutes = cfg.BucketMinutes || 20;
        const cutoff = new Date(Date.now() - retentionDays * 86400000);
        const msPerBucket = bucketMinutes * 60 * 1000;
        const keep = new Map();

        for (let i = 0; i < data.length; i++) {
          const r = data[i];

          // --- BULLETPROOF DATE CHECK ---
          let d = r[DATE];
          if (!(d instanceof Date)) {
            d = new Date(d); // Force text strings into Date objects
          }
          if (isNaN(d.getTime()) || d < cutoff) continue;
          // ------------------------------

          const validMinSell = (r[MIN_SELL] != null && r[MIN_SELL] !== "" && Number(r[MIN_SELL]) > 0);
          const validMaxBuy = (r[MAX_BUY] != null && r[MAX_BUY] !== "" && Number(r[MAX_BUY]) > 0);

          if (!validMinSell && !validMaxBuy) {
            continue;
          }

          const bucket = Math.floor(d.getTime() / msPerBucket);
          const key = bucket + '|' + r[TYPE] + '|' + r[MID] + '|' + r[MTP];

          const prev = keep.get(key);
          if (!prev || (d > prev[DATE])) {
            keep.set(key, r);
          }
        }

        const rowsToWrite = Array.from(keep.values());

        if (rowsToWrite.length > 0) {
          const docLock = LockService.getDocumentLock();
          if (docLock.tryLock(FUZZ_DOC_LOCK_TIMEOUT)) {
            try {
              tempSheet.getRange(tempSheet.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length).setValues(rowsToWrite);
              LOG.info(`Appended ${rowsToWrite.length} deduped rows to ${tempSheetName}.`);
            } finally {
              docLock.releaseLock();
            }
          } else {
            LOG.warn(`Document Lock busy for prune write. Rescheduling (will re-process batch).`);
            scheduleOneTimeTrigger('_heavyPruneWorker', FUZZ_RESCHEDULE_MS);
            return;
          }
        }

        readRow += rowsToRead;
        SCRIPT_PROP.setProperty(PRUNE_PROP_READ_ROW, readRow.toString());

      }

      if (readRow > lastRow) {
        LOG.info("All source rows processed. Transitioning to FINALIZING.");
        currentState = "FINALIZING";
        SCRIPT_PROP.setProperty(PRUNE_PROP_STEP, currentState);
        scheduleOneTimeTrigger('_finalizePrune', 1000);
      }
    }

  } catch (e) {
    LOG.error(`Unhandled error in prune worker: ${e.message}\nStack: ${e.stack}`);
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

  if (SCRIPT_PROP.getProperty(PRUNE_PROP_STEP) !== "FINALIZING") {
    LOG.warn(`Finalizer called in incorrect state (${SCRIPT_PROP.getProperty(PRUNE_PROP_STEP)}). Aborting.`);
    return;
  }

  LOG.info("Starting finalization: Secondary deduplication and atomic swap.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const tempSheetName = PRUNE_SHEET_TEMP;
  const finalSheetName = FUZZ_SHEET_FINAL;
  const oldSheetName = finalSheetName + "_Prune_Old";

  const docLock = LockService.getDocumentLock();
  try {
    if (docLock.tryLock(30000)) {
      try {
        const tempSheet = ss.getSheetByName(tempSheetName);
        if (!tempSheet || tempSheet.getLastRow() <= 1) {
          throw new Error(`Prune temp sheet '${tempSheetName}' is missing or empty! Cannot finalize.`);
        }

        LOG.info("Reading temp sheet for final deduplication...");
        const data = tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, tempSheet.getLastColumn()).getValues();

        const bucketMinutes = cfg.BucketMinutes || 20;
        const maxRows = cfg.PricesMaxRows || 100000;

        const msPerBucket = bucketMinutes * 60 * 1000;
        const keep = new Map();

        const header = tempSheet.getRange(1, 1, 1, tempSheet.getLastColumn()).getValues()[0];
        const lower = header.map(h => String(h).trim().toLowerCase());
        const find = (name) => lower.findIndex(h => h === name);
        const DATE = find('date'), TYPE = find('type_id'), MID = find('market_id'), MTP = find('market_type');
        const MIN_SELL = find('min_sell');
        const MAX_BUY = find('max_buy');

        for (let i = 0; i < data.length; i++) {
          const r = data[i];

          const validMinSell = (r[MIN_SELL] != null && r[MIN_SELL] !== "" && Number(r[MIN_SELL]) > 0);
          const validMaxBuy = (r[MAX_BUY] != null && r[MAX_BUY] !== "" && Number(r[MAX_BUY]) > 0);

          if (!validMinSell && !validMaxBuy) {
            continue;
          }

          // --- BULLETPROOF DATE CHECK ---
          let d = r[DATE];
          if (!(d instanceof Date)) {
            d = new Date(d); // Force text strings into Date objects
          }
          if (isNaN(d.getTime())) continue;
          // ------------------------------

          const bucket = Math.floor(d.getTime() / msPerBucket);
          const key = bucket + '|' + r[TYPE] + '|' + r[MID] + '|' + r[MTP];

          const prev = keep.get(key);
          if (!prev || (d > prev[DATE])) {
            keep.set(key, r);
          }
        }

        let deduped = Array.from(keep.values());
        LOG.info(`Final deduplication complete. Kept ${deduped.length} rows.`);

        if (deduped.length > maxRows) {
          deduped.sort((a, b) => a[DATE] - b[DATE]);
          deduped = deduped.slice(deduped.length - maxRows);
          LOG.info(`Capped rows to ${deduped.length} (max: ${maxRows}).`);
        }

        tempSheet.clearContents();
        tempSheet.getRange(1, 1, 1, header.length).setValues([header]);
        if (deduped.length > 0) {
          tempSheet.getRange(2, 1, deduped.length, header.length).setValues(deduped);
        }
        _trimTrailing_(tempSheet);
        SpreadsheetApp.flush();
        LOG.info("Final data written to temp sheet.");

        const finalSheet = ss.getSheetByName(finalSheetName);
        const oldSheet = ss.getSheetByName(oldSheetName);

        if (oldSheet) ss.deleteSheet(oldSheet);
        if (finalSheet) finalSheet.setName(oldSheetName);
        tempSheet.setName(finalSheetName);
        tempSheet.showSheet();

        SpreadsheetApp.flush();
        LOG.info("Atomic sheet swap successful.");

        SCRIPT_PROP.deleteProperty(PRUNE_PROP_STEP);
        SCRIPT_PROP.deleteProperty(PRUNE_PROP_READ_ROW);
        LOG.info("Heavy Prune job state reset complete.");

      } catch (swapError) {
        LOG.error(`CRITICAL error during prune swap: ${swapError.message}. State NOT reset.`);
        scheduleOneTimeTrigger('_finalizePrune', 60000);
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
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
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
 * THE LIVE WIRE RECALL: Surgical BigQuery Puller with Local Fallback.
 * Refined to match the E7:K standard and ConfigHandler logic.
 */
/**
 * THE LIVE WIRE RECALL: Surgical BigQuery Puller with Local Fallback.
 */
function FUZ_refreshPriceInterface(ss, targetSheetName) {
  const cfg = getConfig();
  const projectId = cfg.BQ_PROJECT_ID || 'tenacious-tiger-345318';
  const isVaultOk = (cfg.BQ_ENABLED === true);

  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) return console.warn(`[SKIP] Sheet "${targetSheetName}" not found.`);

  // --- 1. THE GATE (C4:D4) ---
  const settings = sheet.getRange("C4:D4").getValues()[0];
  const marketId = parseInt(settings[0]);
  const marketType = String(settings[1]).trim();

  if (isNaN(marketId)) return;

  // --- 2. GET ITEM IDS (Column B) ---
  const lastRow = sheet.getLastRow();
  if (lastRow < 8) return;
  const itemIds = sheet.getRange(8, 2, lastRow - 7, 1).getValues().flat().filter(id => id > 0);

  // --- 3. THE RECALL (Unified Logic) ---
  let data = [];
  try {
    if (isVaultOk) {
      // Use the shared function so indexing is always perfect
      data = getPricesFromVault_(projectId, marketId, marketType, itemIds);
    } else {
      data = getPricesFromLocalBuffer_(ss.getSheetByName("Market Prices"), marketId, itemIds);
    }

    // --- 4. THE RESPONSE (Write to E:K) ---
    writeToPriceInterface_(sheet, data, isVaultOk);

  } catch (err) {
    console.warn(`[${targetSheetName}] Sync Failed: ${err.message}`);
    sheet.getRange("E4").setValue(`[!] Sync Error: ${err.message}`);
  }
}

function FUZ_writeToInterfaceSheet(sheet, data, statusPrefix) {
  // Matches Column E, F, G, H, I, J, K
  const HEADERS = [["type_id_filtered", "Current Sell", "Current Buy", "Median Sell", "Median Buy", "Sell Change", "Buy Change"]];
  
  // Clear E7:K
  const lastRow = sheet.getLastRow();
  if (lastRow >= 7) sheet.getRange(7, 5, lastRow - 6, 7).clearContent();

  sheet.getRange(7, 5, 1, 7).setValues(HEADERS);
  if (data.length > 0) {
    sheet.getRange(8, 5, data.length, 7).setValues(data);
    sheet.getRange(8, 10, data.length, 2).setNumberFormat('0.00%'); // Changes
    sheet.getRange(8, 6, data.length, 4).setNumberFormat('#,##0.00 "ISK"'); // Prices
  }
}

/**
 * THE GENERAL: Direct Cache-First Recall via fuzAPI.
 * Targets: 'filtered prices', 'Mineral Supply Prices', 'T1 Supply Prices'.
 */
function FUZ_publishPriceInterfaces(ss) {
  if (!ss) ss = SpreadsheetApp.getActive();

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

    const marketId = parseInt(mktIdRaw, 10);
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
      .filter(id => !isNaN(parseInt(id, 10)) && id > 0);

    if (itemIds.length === 0) return;

    console.log(`[${sheetName}] Requesting ${itemIds.length} items from cache client for ${marketType} ${marketId}`);

    // --- 3. THE RECALL ---
    let data = [];
    try {
      // Pull directly from modular cache client
      const priceMap = getMarketPrices(itemIds, marketId, marketType);

      // Reconstruct the 2D array structure expected by the printer [ID, MinS, MaxB, MedS, MedB]
      data = itemIds.map(id => {
        const obj = priceMap[id];
        if (obj) {
          return [
            Number(id), // Keep native numeric type to align with SDE schemas
            obj.sell?.min !== "" ? Number(obj.sell.min) : "",
            obj.buy?.max  !== "" ? Number(obj.buy.max)  : "",
            obj.sell?.median !== "" ? Number(obj.sell.median) : "",
            obj.buy?.median  !== "" ? Number(obj.buy.median)  : ""
          ];
        }
        return [Number(id), "", "", "", ""];
      });

      // 1. Enforce numeric standard format directly on the ranges
      sh.getRange("E7:E").setNumberFormat("0");
      sh.getRange("F7:K").setNumberFormat("#,##0.00");

      // --- 4. THE RESPONSE (Write to E7:K) ---
      writeToPriceInterface_(sh, data, false);

    } catch (e) {
      console.error(`[${sheetName}] Sync Failed: ${e.message}`);
      sh.getRange("E4").setValue("!! Sync Error: Check Logs !!");
    }
  });
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

function emergencyGhostCleanup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  sheets.forEach(sh => {
    if (sh.getName().includes("Merge_Temp_")) ss.deleteSheet(sh);
  });
  console.log("Workbook Cleaned.");
}

function getPricesFromLocalBuffer_(sh, marketId, itemIds) {
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const idSet = new Set(itemIds.map(id => Number(id)));
  
  return data.slice(1)
    .filter(r => idSet.has(Number(r[3])) && Number(r[1]) === Number(marketId))
    .map(r => [String(r[3]), r[4], r[5], r[6], r[7]]); // E, F, G, H, I
}

function getPricesFromVault_(projectId, marketId, marketType, itemIds) {
  const cfg = getConfig();
  const targetProject = cfg.BQ_PROJECT_ID || 'tenacious-tiger-345318';
  const tableId = cfg.BQ_Market_Prices || 'market_prices_history';
  const idList = itemIds.join(',');
  const cleanType = String(marketType).toLowerCase().trim();

  const sql = `
    WITH History AS (
      SELECT type_id, min_sell, max_buy, date FROM \`${targetProject}.market_data.${tableId}\`
      WHERE market_id = ${marketId} AND market_type = '${cleanType}' AND type_id IN (${idList})
      AND date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 24 HOUR)
    ),
    CalculatedBaselines AS (
      SELECT type_id,
        PERCENTILE_CONT(NULLIF(min_sell, 0), 0.5) OVER(PARTITION BY type_id) as med_sell_base,
        PERCENTILE_CONT(NULLIF(max_buy, 0), 0.5) OVER(PARTITION BY type_id) as med_buy_base
      FROM History
    ),
    LatestPrice AS (
      SELECT type_id, min_sell, max_buy, ROW_NUMBER() OVER(PARTITION BY type_id ORDER BY date DESC) as rn FROM History
    )
    SELECT CAST(L.type_id AS STRING), L.min_sell, L.max_buy, CAST(B.med_sell_base AS FLOAT64), CAST(B.med_buy_base AS FLOAT64)
    FROM LatestPrice L
    LEFT JOIN (SELECT DISTINCT * FROM CalculatedBaselines) B ON L.type_id = B.type_id
    WHERE L.rn = 1`;

  const queryResults = BigQuery.Jobs.query({ query: sql, useLegacySql: false }, targetProject);

  return queryResults.rows ? queryResults.rows.map(row => {
    // Just return the 5 values BigQuery found
    return row.f.map(field => field.v !== null ? field.v : "");
  }) : [];
}

function forceSystemRecovery() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = getAppContext(ss); // Get the handshake
  console.log("🛠 Recovery: Clearing lease and forcing Vault Merge...");
  ctx.props.deleteProperty('LAST_FUZZ_FETCH');
  ctx.props.setProperty('fuzz_pass_complete', 'true'); 
  
  runVaultMergeAndReset(ctx); // Pass the ctx
  console.log("✅ Recovery Complete.");
}

function writeToPriceInterface_(sh, data, isVaultOk) {
  if (!data || data.length === 0) return;

  const startRow = 8; // Data starts here
  const lastRow = sh.getLastRow();
  
  // 1. Clear ONLY E8:I (5 columns)
  if (lastRow >= startRow) {
    sh.getRange(startRow, 5, lastRow - startRow + 1, 5).clearContent();
  }

  const rawIds = sh.getRange(startRow, 2, Math.max(1, lastRow - startRow + 1), 1).getValues().flat();
  const validIds = rawIds.filter(id => id !== "" && !isNaN(Number(id)));

  const vaultMap = new Map();
  data.forEach(row => vaultMap.set(Number(row[0]), row));

  const output = validIds.map(id => {
    const cleanId = Number(id);
    return vaultMap.has(cleanId) ? vaultMap.get(cleanId) : [String(cleanId), "", "", "", ""];
  });

  // 2. THE WRITE: Exactly 5 columns (E to I)
  sh.getRange(startRow, 5, output.length, 5).setValues(output);
  
  // Format Prices (F:I)
  sh.getRange(startRow, 6, output.length, 4).setNumberFormat('#,##0.00 "ISK"');

  const ts = Utilities.formatDate(new Date(), "America/New_York", "h:mm:ss a");
  sh.getRange("E4").setValue(`✅ Vault Synced: ${ts}`);
}