/** MarketFetcher.gs — STATEFUL Prices runner with Predictive Scheduling & Locks
 * Manages fetching Fuzzwork prices based on state properties.
 * Called by the masterOrchestrator.
 *
 * NOW INCLUDES:
 * - Stateful Heavy Prune worker (`_heavyPruneWorker`)
 * - Daily Job Reset (`dailyJobReset`)
 */

/* global LockService, PropertiesService, SpreadsheetApp, LoggerEx, fuzAPI, getMasterMarketRequests, getOrCreateSheet, scheduleOneTimeTrigger, STATE_FLAGS, JOB_LEASE_DURATION_MS, mtConfig, executeWithTryLock, pruneOldRows, _trimTrailing_, getConfig */ // <-- Added getConfig

// --- Constants ---
const FUZZ_JOB_PREFIX = 'fuzzJob'; // Prefix for state properties
const FUZZ_PROP_STEP = FUZZ_JOB_PREFIX + 'Step';
const FUZZ_PROP_INDEX = FUZZ_JOB_PREFIX + 'RequestIndex';
const FUZZ_PROP_ROW = FUZZ_JOB_PREFIX + 'WriteRow';
const FUZZ_PROP_LEASE = FUZZ_JOB_PREFIX + 'LeaseUntil';
const FUZZ_SHEET_TEMP = 'Market_Prices_Temp'; // Temporary sheet for writes
const FUZZ_SHEET_FINAL = 'Market Prices';     // Final destination sheet
const FUZZ_SHEET_OLD = 'Market_Prices_Old';   // Intermediate for swap delete
const FUZZ_SHEET_HEADERS = ["date", "market_id", "market_type", "type_id", "min_sell", "max_buy", "median_sell", "median_buy"];

const FUZZ_BATCH_SIZE = 750; // How many requests to process per execution run
const FUZZ_TIME_LIMIT_MS = 280000;      // Soft limit (4m 40s) before rescheduling
const FUZZ_RESCHEDULE_MS = 5000;        // Delay for rescheduling (5s)
const FUZZ_DOC_LOCK_TIMEOUT = 10000;    // Wait 10s for DocumentLock on write

// --- [NEW PRUNE CONSTANTS] ---
const PRUNE_JOB_PREFIX = 'heavyPruneJob'; // Prefix for state properties
const PRUNE_PROP_STEP = PRUNE_JOB_PREFIX + 'Step';
const PRUNE_PROP_READ_ROW = PRUNE_JOB_PREFIX + 'ReadRow';
const PRUNE_SHEET_TEMP = 'Market_Prices_Prune_Temp'; // Temp sheet for pruning
const PRUNE_BATCH_SIZE = 5000; // How many rows to read/process at a time
// --- [END OF NEW CONSTANTS] ---

const LOG_FUZZ = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('FuzzWorker') : console);

/**
 * Public wrapper function called by the orchestrator. Uses ScriptLock.
 */
function updateFuzzMarketDataSheet() {
  // Uses executeWithTryLock from Orchestrator.gs.js
  const result = executeWithTryLock(_updateFuzzMarketDataWorker, 'updateFuzzMarketDataSheet');
  if (result === null) {
    LOG_FUZZ.warn("Execution skipped by ScriptLock. Will retry on next trigger.");
  }
}

/**
 * Resets the state of the Fuzz market data job.
 */
function _resetFuzzMarketDataJobState(error) {
  LOG_FUZZ.warn(`RESETTING Fuzz Market Data Job State. Reason: ${error ? error.message : 'Completion/Manual'}`);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  try {
    SCRIPT_PROP.deleteProperty(FUZZ_PROP_STEP);
    SCRIPT_PROP.deleteProperty(FUZZ_PROP_INDEX);
    SCRIPT_PROP.deleteProperty(FUZZ_PROP_ROW);
    SCRIPT_PROP.deleteProperty(FUZZ_PROP_LEASE);
    // Delete potential triggers
    deleteTriggersByName('updateFuzzMarketDataSheet');
    deleteTriggersByName('_finalizeFuzzDataUpdate'); // Ensure finalizer trigger is cleared
  } catch (propError) {
    LOG_FUZZ.error(`Error deleting script properties: ${propError.message}`);
  }
  LOG_FUZZ.info("Fuzz market data job state reset complete.");
}


/**
 * The core stateful worker function for fetching Fuzz data.
 */
function _updateFuzzMarketDataWorker() {
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const START_TIME = Date.now();

  // --- State Initialization & Validation ---
  let currentState = SCRIPT_PROP.getProperty(FUZZ_PROP_STEP) || STATE_FLAGS.NEW_RUN;
  LOG_FUZZ.info(`Starting worker. Current State: ${currentState}`);

  // --- Lease Management ---
  const leaseUntil = parseInt(SCRIPT_PROP.getProperty(FUZZ_PROP_LEASE) || '0', 10);
  if (START_TIME > leaseUntil) {
    LOG_FUZZ.error(`Job lease expired or not set! Lease ended at ${new Date(leaseUntil)}. Resetting job state.`);
    _resetFuzzMarketDataJobState(new Error("Job lease expired"));
    return; // Halt execution
  }
  // Extend lease if nearing expiry within this run
  if (leaseUntil - START_TIME < 60000) { // Less than 1 min left
    const newLease = START_TIME + JOB_LEASE_DURATION_MS;
    SCRIPT_PROP.setProperty(FUZZ_PROP_LEASE, newLease.toString());
    LOG_FUZZ.info(`Extended job lease until ${new Date(newLease)}`);
  }
  // --- End Lease Management ---


  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tempSheet = null; // Initialize

  try {
    // --- State: NEW_RUN (Setup Phase) ---
    if (currentState === STATE_FLAGS.NEW_RUN) {
      LOG_FUZZ.info(`State: ${STATE_FLAGS.NEW_RUN}. Preparing temporary sheet.`);

      // Use Document Lock for sheet creation/clearing
      const docLock = LockService.getDocumentLock();
      if (docLock.tryLock(FUZZ_DOC_LOCK_TIMEOUT)) {
        try {
          // Delete old sheet if it exists
          const oldSheet = ss.getSheetByName(FUZZ_SHEET_OLD);
          if (oldSheet) ss.deleteSheet(oldSheet);

          tempSheet = getOrCreateSheet(ss, FUZZ_SHEET_TEMP, FUZZ_SHEET_HEADERS);
          if (tempSheet.getLastRow() > 1) {
            tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, tempSheet.getMaxColumns()).clearContent();
          }
          tempSheet.hideSheet();
          SpreadsheetApp.flush(); // Ensure sheet operations complete

          // Initialize state for processing
          SCRIPT_PROP.setProperty(FUZZ_PROP_INDEX, '0');
          SCRIPT_PROP.setProperty(FUZZ_PROP_ROW, '2'); // Data starts row 2
          currentState = STATE_FLAGS.PROCESSING;
          SCRIPT_PROP.setProperty(FUZZ_PROP_STEP, currentState);
          LOG_FUZZ.info(`Temp sheet '${FUZZ_SHEET_TEMP}' prepared. Transitioning to ${currentState}.`);

        } finally {
          docLock.releaseLock();
        }
      } else {
        LOG_FUZZ.warn(`Document Lock busy during setup. Rescheduling.`);
        scheduleOneTimeTrigger('updateFuzzMarketDataSheet', FUZZ_RESCHEDULE_MS);
        return; // Reschedule and exit
      }
    } // --- End NEW_RUN ---


    // --- State: PROCESSING ---
    if (currentState === STATE_FLAGS.PROCESSING) {
      LOG_FUZZ.info(`State: ${STATE_FLAGS.PROCESSING}. Fetching and writing batches.`);

      let requestStartIndex = parseInt(SCRIPT_PROP.getProperty(FUZZ_PROP_INDEX) || '0');
      let nextWriteRow = parseInt(SCRIPT_PROP.getProperty(FUZZ_PROP_ROW) || '2');
      const allMarketRequests = getMasterMarketRequests(); // Get full list each time

      if (!allMarketRequests || allMarketRequests.length === 0) {
        LOG_FUZZ.warn("Master request list is empty. Resetting job.");
        _resetFuzzMarketDataJobState(new Error("Master request list empty"));
        return;
      }

      tempSheet = ss.getSheetByName(FUZZ_SHEET_TEMP); // Ensure we have the sheet object
      if (!tempSheet) {
        throw new Error(`Sheet '${FUZZ_SHEET_TEMP}' missing during PROCESSING phase.`);
      }

      let batchesProcessedThisRun = 0;

      // --- Processing Loop ---
      while (requestStartIndex < allMarketRequests.length) {
        // --- Time Limit Check ---
        if (Date.now() - START_TIME > FUZZ_TIME_LIMIT_MS) {
          SCRIPT_PROP.setProperty(FUZZ_PROP_INDEX, requestStartIndex.toString());
          SCRIPT_PROP.setProperty(FUZZ_PROP_ROW, nextWriteRow.toString());
          scheduleOneTimeTrigger('updateFuzzMarketDataSheet', FUZZ_RESCHEDULE_MS);
          LOG_FUZZ.warn(`Time limit hit after ${batchesProcessedThisRun} batches. Saved state. Rescheduled.`);
          return; // Exit current execution
        }

        // --- Prepare Batch & Group ---
        const requestEndIndex = Math.min(requestStartIndex + FUZZ_BATCH_SIZE, allMarketRequests.length);
        const requestsForThisRun = allMarketRequests.slice(requestStartIndex, requestEndIndex);
        const groupedRequests = {};
        requestsForThisRun.forEach(req => {
          const key = `${req.market_id}_${req.market_type}`;
          if (!groupedRequests[key]) groupedRequests[key] = {
            market_id: req.market_id,
            market_type: req.market_type,
            typeIDs: []
          };
          groupedRequests[key].typeIDs.push(req.type_id);
        });

        LOG_FUZZ.info(`Processing batch indices ${requestStartIndex}-${requestEndIndex - 1} (${requestsForThisRun.length} reqs, ${Object.keys(groupedRequests).length} markets).`);

        // --- Fetch Data ---
        const now = new Date();
        const rowsToWrite = [];
        let fetchErrorOccurred = false;

        Object.values(groupedRequests).forEach(({
          market_id,
          market_type,
          typeIDs
        }) => {
          try {
            const prices = getMarketPrices(typeIDs, market_id, market_type); // Calls fuzAPI internally
            typeIDs.forEach(type_id => {
              const e = prices[type_id] || {};
              rowsToWrite.push([now, market_id, market_type, type_id, e.minSell ?? null, e.maxBuy ?? null, e.medianSell ?? null, e.medianBuy ?? null]);
            });
          } catch (apiError) {
            LOG_FUZZ.error(`API error for ${market_type}:${market_id} (items: ${typeIDs.length}): ${apiError.message}`);
            fetchErrorOccurred = true;
          }
        });

        // --- Write Batch (Document Lock) ---
        if (rowsToWrite.length > 0) {
          const docLock = LockService.getDocumentLock();
          if (docLock.tryLock(FUZZ_DOC_LOCK_TIMEOUT)) {
            try {
              const range = tempSheet.getRange(nextWriteRow, 1, rowsToWrite.length, FUZZ_SHEET_HEADERS.length);
              range.setValues(rowsToWrite);
              nextWriteRow += rowsToWrite.length;
              batchesProcessedThisRun++;
              LOG_FUZZ.info(`Batch write success. ${rowsToWrite.length} rows written. Next row: ${nextWriteRow}`);
            } catch (writeError) {
              LOG_FUZZ.error(`Error during batch write: ${writeError.message}. Rescheduling.`);
              SCRIPT_PROP.setProperty(FUZZ_PROP_INDEX, requestStartIndex.toString()); // Save current index
              SCRIPT_PROP.setProperty(FUZZ_PROP_ROW, nextWriteRow.toString()); // Save potentially advanced row
              scheduleOneTimeTrigger('updateFuzzMarketDataSheet', FUZZ_RESCHEDULE_MS);
              throw writeError; // Re-throw to ensure finally block runs and exits
            } finally {
              docLock.releaseLock();
            }
          } else {
            LOG_FUZZ.warn(`Document Lock busy for write. Saving state and rescheduling.`);
            SCRIPT_PROP.setProperty(FUZZ_PROP_INDEX, requestStartIndex.toString());
            SCRIPT_PROP.setProperty(FUZZ_PROP_ROW, nextWriteRow.toString());
            scheduleOneTimeTrigger('updateFuzzMarketDataSheet', FUZZ_RESCHEDULE_MS);
            return; // Exit
          }
        } else if (!fetchErrorOccurred) {
          LOG_FUZZ.info(`No data returned/to write for batch indices ${requestStartIndex}-${requestEndIndex - 1}. Advancing.`);
        }

        // --- Advance Index Only After Successful Handling ---
        requestStartIndex = requestEndIndex;
        SCRIPT_PROP.setProperty(FUZZ_PROP_INDEX, requestStartIndex.toString());
        SCRIPT_PROP.setProperty(FUZZ_PROP_ROW, nextWriteRow.toString()); // Save row progress too

      } // --- End while loop ---

      // --- Post-Loop Check ---
      if (requestStartIndex >= allMarketRequests.length) {
        LOG_FUZZ.info("All batches processed. Transitioning to FINALIZING.");
        currentState = STATE_FLAGS.FINALIZING;
        SCRIPT_PROP.setProperty(FUZZ_PROP_STEP, currentState);
        // Immediately schedule the finalizer
        scheduleOneTimeTrigger('_finalizeFuzzDataUpdate', 1000); // 1 sec delay
      }

    } // --- End PROCESSING ---

  } catch (e) {
    LOG_FUZZ.error(`Unhandled error in worker: ${e.message}\nStack: ${e.stack}`);
    _resetFuzzMarketDataJobState(e);
  } finally {
    const duration = (Date.now() - START_TIME) / 1000;
    LOG_FUZZ.info(`Worker execution finished in ${duration.toFixed(2)}s. Final State: ${currentState}`);
  }
}

/**
 * Performs the atomic sheet swap using DocumentLock. Triggered after PROCESSING.
 */
function _finalizeFuzzDataUpdate() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('FuzzFinalizer') : console);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();

  // Ensure state is correct
  if (SCRIPT_PROP.getProperty(FUZZ_PROP_STEP) !== STATE_FLAGS.FINALIZING) {
    LOG.warn(`Finalizer called in incorrect state (${SCRIPT_PROP.getProperty(FUZZ_PROP_STEP)}). Resetting job.`);
    _resetFuzzMarketDataJobState(new Error("Finalizer called in incorrect state"));
    return;
  }

  LOG.info("Starting finalization: Atomic sheet swap.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docLock = LockService.getDocumentLock();

  try {
    if (docLock.tryLock(30000)) { // Wait up to 30s for the lock
      try {
        const tempSheet = ss.getSheetByName(FUZZ_SHEET_TEMP);
        const finalSheet = ss.getSheetByName(FUZZ_SHEET_FINAL);
        const oldSheet = ss.getSheetByName(FUZZ_SHEET_OLD);

        if (!tempSheet || tempSheet.getLastRow() <= 1) {
          throw new Error(`Temp sheet '${FUZZ_SHEET_TEMP}' is missing or empty! Cannot finalize.`);
        }

        // 1. Delete previous "Old" sheet (if exists)
        if (oldSheet) ss.deleteSheet(oldSheet);

        // 2. Rename current "Final" sheet to "Old" (if exists)
        if (finalSheet) finalSheet.setName(FUZZ_SHEET_OLD);

        // 3. Rename "Temp" sheet to "Final"
        tempSheet.setName(FUZZ_SHEET_FINAL);
        tempSheet.showSheet(); // Make it visible
        SpreadsheetApp.flush(); // Ensure changes apply

        LOG.info("Atomic sheet swap successful.");

        // 4. Reset job state ONLY on successful swap
        _resetFuzzMarketDataJobState(null); // Pass null for successful completion

      } catch (swapError) {
        LOG.error(`CRITICAL error during sheet swap: ${swapError.message}. State NOT reset. Manual intervention likely needed.`);
        scheduleOneTimeTrigger('_finalizeFuzzDataUpdate', 60000); // Retry in 1 min
        throw swapError; // Re-throw
      } finally {
        docLock.releaseLock();
      }
    } else {
      LOG.warn("Document Lock busy during finalization. Rescheduling finalizer.");
      scheduleOneTimeTrigger('_finalizeFuzzDataUpdate', FUZZ_RESCHEDULE_MS);
    }
  } catch (e) {
    LOG.error(`Error in finalizer lock acquisition: ${e.message}`);
    scheduleOneTimeTrigger('_finalizeFuzzDataUpdate', 60000); // Retry in 1 min
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
  if (!sh) return; // Add guard clause
  const used = sh.getLastRow();
  const alloc = sh.getMaxRows();
  const extra = alloc - used;
  if (extra > 0) _deleteInBlocks_(sh, used + 1, extra);
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
function pruneOldRows(sheet, retentionDays, dateCol /* 1-based */ ) {
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
  const cfg = getConfig(); // <-- CORRECTED
  const sheetName = cfg.sheets.prices;
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    LOG.error(`Sheet not found: ${sheetName}. Skipping light prune.`);
  } else {
    try {
      // 1. Run the "Light Prune"
      const retentionDays = cfg.retentionDays.prices || 1;
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
 * It just acquires a lock, sets the state, and calls the worker.
 *
 * This is what you should schedule on your daily trigger (a few hours after dailyJobReset).
 */
function dailyHeavyPrune_Prices() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('HeavyPrune') : console);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();

  // Check if it's already running
  const currentState = SCRIPT_PROP.getProperty(PRUNE_PROP_STEP);
  if (currentState && currentState !== 'COMPLETE' && currentState !== STATE_FLAGS.NEW_RUN) {
    LOG.warn(`Heavy Prune is already running (State: ${currentState}). Skipping new start.`);
    // Re-schedule the worker just in case the trigger was lost
    scheduleOneTimeTrigger('_heavyPruneWorker', 5000);
    return;
  }

  LOG.info("Starting new Heavy Prune cycle.");
  SCRIPT_PROP.setProperty(PRUNE_PROP_STEP, STATE_FLAGS.NEW_RUN);

  // Use executeWithTryLock
  const result = executeWithTryLock(_heavyPruneWorker, '_heavyPruneWorker');
  if (result === null) {
    LOG.warn("Heavy Prune start skipped by ScriptLock. Another process is running.");
  }
}

/**
 * NEW: The "Processing While Loop" worker for the heavy prune.
 * This function reads, processes, and writes in batches to avoid timeouts.
 */
function _heavyPruneWorker() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('PruneWorker') : console);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const START_TIME = Date.now();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig(); // <-- CORRECTED
  const sourceSheetName = cfg.sheets.prices;
  const tempSheetName = PRUNE_SHEET_TEMP;

  let currentState = SCRIPT_PROP.getProperty(PRUNE_PROP_STEP) || STATE_FLAGS.NEW_RUN;
  LOG.info(`Starting worker. Current State: ${currentState}`);

  try {
    // --- State: NEW_RUN (Start) ---
    if (currentState === STATE_FLAGS.NEW_RUN) {
      LOG.info(`State: ${STATE_FLAGS.NEW_RUN}. Preparing prune temp sheet.`);
      const docLock = LockService.getDocumentLock();
      if (docLock.tryLock(FUZZ_DOC_LOCK_TIMEOUT)) {
        try {
          const tempSheet = getOrCreateSheet(ss, tempSheetName, FUZZ_SHEET_HEADERS);
          if (tempSheet.getLastRow() > 1) {
            tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, tempSheet.getMaxColumns()).clearContent();
          }
          tempSheet.hideSheet();
          SpreadsheetApp.flush();

          SCRIPT_PROP.setProperty(PRUNE_PROP_READ_ROW, '2'); // Data starts row 2
          currentState = STATE_FLAGS.PROCESSING;
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
    if (currentState === STATE_FLAGS.PROCESSING) {
      LOG.info(`State: ${STATE_FLAGS.PROCESSING}. Reading/deduping batches.`);
      
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
      if (DATE < 0 || TYPE < 0 || MID < 0 || MTP < 0) {
        throw new Error('Missing required columns (date/type_id/market_id/market_type) in source sheet.');
      }
      
      // --- Processing Loop ---
      while (readRow <= lastRow) {
        // --- 1. Time Limit Check ---
        if (Date.now() - START_TIME > FUZZ_TIME_LIMIT_MS) {
          SCRIPT_PROP.setProperty(PRUNE_PROP_READ_ROW, readRow.toString());
          scheduleOneTimeTrigger('_heavyPruneWorker', FUZZ_RESCHEDULE_MS);
          LOG.warn(`Time limit hit. Saved state. Rescheduled. Next read row: ${readRow}`);
          return;
        }

        // --- 2. Read Batch ---
        const rowsToRead = Math.min(PRUNE_BATCH_SIZE, lastRow - readRow + 1);
        if (rowsToRead <= 0) break; 
        
        LOG.info(`Reading ${rowsToRead} rows from ${sourceSheetName} (starting row ${readRow})...`);
        const data = sourceSheet.getRange(readRow, 1, rowsToRead, header.length).getValues();
        
        // --- 3. Process Batch (Retention, Bucket, Dedupe) ---
        const { bucketMinutes, retentionDays } = cfg;
        const cutoff = new Date(Date.now() - retentionDays.prices * 86400000);
        const msPerBucket = (bucketMinutes || 20) * 60 * 1000;
        const keep = new Map(); // Keep latest record *within this batch*

        for (let i = 0; i < data.length; i++) {
          const r = data[i];
          const d = r[DATE];
          if (!(d instanceof Date) || d < cutoff) continue; // Retention filter

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
              LOG.info(`Appended ${rowsToWrite.length} deduped rows to ${tempSheetName}.`);
            } finally {
              docLock.releaseLock();
            }
          } else {
            // RETRIGGER ON WRITE FAILURE
            LOG.warn(`Document Lock busy for prune write. Rescheduling (will re-process batch).`);
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
        LOG.info("All source rows processed. Transitioning to FINALIZING.");
        currentState = STATE_FLAGS.FINALIZING;
        SCRIPT_PROP.setProperty(PRUNE_PROP_STEP, currentState);
        scheduleOneTimeTrigger('_finalizePrune', 1000); // 1 sec delay
      }
    } // --- End PROCESSING ---

  } catch (e) {
    LOG.error(`Unhandled error in prune worker: ${e.message}\nStack: ${e.stack}`);
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
  
  if (SCRIPT_PROP.getProperty(PRUNE_PROP_STEP) !== STATE_FLAGS.FINALIZING) {
    LOG.warn(`Finalizer called in incorrect state (${SCRIPT_PROP.getProperty(PRUNE_PROP_STEP)}). Aborting.`);
    return;
  }

  LOG.info("Starting finalization: Secondary deduplication and atomic swap.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig(); // <-- CORRECTED
  const tempSheetName = PRUNE_SHEET_TEMP;
  const finalSheetName = cfg.sheets.prices;
  const oldSheetName = FUZZ_SHEET_OLD; // Use the same "Old" sheet as the Fuzz worker

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
        
        const { bucketMinutes, maxRows } = cfg;
        const msPerBucket = (bucketMinutes || 20) * 60 * 1000;
        const keep = new Map();
        
        const header = tempSheet.getRange(1, 1, 1, tempSheet.getLastColumn()).getValues()[0];
        const lower = header.map(h => String(h).trim().toLowerCase());
        const find = (name) => lower.findIndex(h => h === name);
        const DATE = find('date'), TYPE = find('type_id'), MID = find('market_id'), MTP = find('market_type');

        for (let i = 0; i < data.length; i++) {
          const r = data[i];
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
        if (deduped.length > maxRows.prices) {
          deduped.sort((a,b) => a[DATE] - b[DATE]); // Sort by date ascending
          deduped = deduped.slice(deduped.length - maxRows.prices); // Keep the newest rows
          LOG.info(`Capped rows to ${deduped.length} (max: ${maxRows.prices}).`);
        }
        
        // --- 3. Rewrite Temp Sheet ---
        tempSheet.clearContents(); // Clear everything
        tempSheet.getRange(1, 1, 1, header.length).setValues([header]); // Set header
        if (deduped.length > 0) {
          tempSheet.getRange(2, 1, deduped.length, header.length).setValues(deduped);
        }
        _trimTrailing_(tempSheet);
        SpreadsheetApp.flush();
        LOG.info("Final data written to temp sheet.");

        // --- 4. Atomic Swap ---
        const finalSheet = ss.getSheetByName(finalSheetName);
        const oldSheet = ss.getSheetByName(oldSheetName);

        if (oldSheet) ss.deleteSheet(oldSheet);
        if (finalSheet) finalSheet.setName(oldSheetName);
        tempSheet.setName(finalSheetName);
        tempSheet.showSheet();
        SpreadsheetApp.flush();
        LOG.info("Atomic sheet swap successful.");

        // --- 5. Reset Prune Job State ---
        SCRIPT_PROP.deleteProperty(PRUNE_PROP_STEP);
        SCRIPT_PROP.deleteProperty(PRUNE_PROP_READ_ROW);
        LOG.info("Heavy Prune job state reset complete.");

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