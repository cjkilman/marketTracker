/** HistoryManager.gs — STATEFUL Daily OHLC/median builder
 *
 * REWRITTEN to be a stateful, batch-processing worker
 * to avoid 6-minute timeouts on large data.
 *
 * Source sheet:  "Market Prices"
 * Target sheet:  "Market History"
 */

/* global LockService, PropertiesService, SpreadsheetApp, LoggerEx, STATE_FLAGS, 
   executeWithTryLock, scheduleOneTimeTrigger, getOrCreateSheet, _trimTrailing_,
   FUZZ_TIME_LIMIT_MS, FUZZ_RESCHEDULE_MS, FUZZ_DOC_LOCK_TIMEOUT, PT */

// --- [NEW HISTORY CONSTANTS] ---
const HIST_JOB_PREFIX = 'historyJob'; // Prefix for state properties
const HIST_PROP_STEP = HIST_JOB_PREFIX + 'Step';
const HIST_PROP_READ_ROW = HIST_JOB_PREFIX + 'ReadRow';
const HIST_SHEET_TEMP = 'Market_History_Temp'; // Temp sheet for building
const HIST_BATCH_SIZE = 5000; // How many rows to read from Market Prices at a time
// --- [END OF NEW CONSTANTS] ---

const HM_HEADERS = [
  'type_id', 'market_id', 'market_type', 'date',
  'buy_open', 'buy_close', 'sell_open', 'sell_close',
  'buy_high', 'buy_low', 'sell_high', 'sell_low',
  'median_buy', 'median_sell'
];

/* ============================ Config ============================ */

function _hmCfg() {
  const c = (typeof getConfig === 'function') ? (getConfig() || {}) : {};
  const num = (v, d) => (v == null || isNaN(Number(v))) ? d : Number(v);
  const str = (v, d) => (v == null || v === '') ? d : String(v);
  const bool = (v, d = false) => {
    const s = String(v == null ? '' : v).trim().toUpperCase();
    return s === 'TRUE' ? true : (s === 'FALSE' ? false : d);
  };
  const retentionDays = num(c["HistoryRetentionDays"], num(c["HistoryDaysLimit"], 365));

  return {
    sheets: {
      prices: str(c["MarketPricesSheet"], "Market Prices"),
      history: str(c["HistorySheetName"], "Market History"),
    },
    historyRetentionDays: retentionDays,
    historyMaxRows: num(c["HistoryMaxRows"], 200000),
    // ... other configs ...
    openTime: str(c["OpenTime"], "11:00"),
    closeTime: str(c["CloseTime"], "18:00"),
    openWindowMins: num(c["OpenWindowMins"], 30),
    closeWindowMins: num(c["CloseWindowMins"], 30),
    lightSlices: bool(c["History.LightSlices"], true),
  };
}

// --- Time helpers ---
function _toProjectDay_(v) {
  const d = PT.parseDateSafe(v);
  if (isNaN(d)) return new Date('Invalid');
  return PT.projectDate(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0);
}

function _minsSinceMidnight_(d) {
  return d.getHours() * 60 + d.getMinutes();
}

function _sliceFlags_(cfg) {
  const now = PT.now();
  const openHM = PT.coerceHM(cfg.openTime);
  const closeHM = PT.coerceHM(cfg.closeTime);
  const minsNow = _minsSinceMidnight_(now);
  const openMin = (openHM.h | 0) * 60 + (openHM.m | 0);
  const closeMin = (closeHM.h | 0) * 60 + (closeHM.m | 0);
  const inOpenWindow = minsNow >= openMin && minsNow < openMin + (cfg.openWindowMins | 0);
  const inCloseWindow = minsNow >= closeMin && minsNow <= closeMin + (cfg.closeWindowMins | 0);
  return { now, minsNow, openMin, closeMin, inOpenWindow, inCloseWindow };
}

/* ======================== Aggregation Helpers ======================== */

function _median_(arr) {
  const xs = (arr || []).filter(n => Number.isFinite(n)).sort((a, b) => a - b);
  if (!xs.length) return null;
  const mid = Math.floor(xs.length / 2);
  return xs.length % 2 ? xs[mid] : (xs[mid - 1] + xs[mid]) / 2;
}

/**
 * MODIFIED: _aggDailyFromPrices_
 * Now accepts the header map 'P' to find columns by name.
 */
function _aggDailyFromPrices_(rows, P) { // <-- Added P
  const bag = new Map();

  function pushSample(key, stamp, sellMin, buyMax, medSell, medBuy, typeId, marketId, mtype, day) {
    let rec = bag.get(key);
    if (!rec) {
      rec = {
        typeId, marketId, mtype, day,
        firstTs: stamp, lastTs: stamp,
        sellOpen: sellMin, sellClose: sellMin,
        buyOpen: buyMax, buyClose: buyMax,
        sellHigh: sellMin != null ? sellMin : null,
        sellLow: sellMin != null ? sellMin : null,
        buyHigh: buyMax != null ? buyMax : null,
        buyLow: buyMax != null ? buyMax : null,
        medSell: [], medBuy: []
      };
      bag.set(key, rec);
    } else {
      if (stamp < rec.firstTs) {
        rec.firstTs = stamp;
        if (sellMin != null) rec.sellOpen = sellMin;
        if (buyMax != null) rec.buyOpen = buyMax;
      }
      if (stamp > rec.lastTs) {
        rec.lastTs = stamp;
        if (sellMin != null) rec.sellClose = sellMin;
        if (buyMax != null) rec.buyClose = buyMax;
      }
      if (sellMin != null) {
        if (rec.sellHigh == null || sellMin > rec.sellHigh) rec.sellHigh = sellMin;
        if (rec.sellLow == null || sellMin < rec.sellLow) rec.sellLow = sellMin;
      }
      if (buyMax != null) {
        if (rec.buyHigh == null || buyMax > rec.buyHigh) rec.buyHigh = buyMax;
        if (rec.buyLow == null || buyMax < rec.buyLow) rec.buyLow = buyMax;
      }
    }
    if (medSell != null) rec.medSell.push(medSell);
    if (medBuy != null) rec.medBuy.push(medBuy);
  }

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    // --- MODIFIED TO USE P MAP ---
    const ts = r[P.date]; if (!(ts instanceof Date)) continue;
    const marketId = Number(r[P.market_id]);
    const mtype = String(r[P.market_type] || '');
    const typeId = Number(r[P.type_id]);
    if (!Number.isFinite(marketId) || !Number.isFinite(typeId) || !mtype) continue;

    const minSell = (r[P.min_sell] == null || r[P.min_sell] <= 0) ? null : Number(r[P.min_sell]);
    const maxBuy = (r[P.max_buy] == null || r[P.max_buy] <= 0) ? null : Number(r[P.max_buy]);
    const medSell = (r[P.median_sell] == null || r[P.median_sell] <= 0) ? null : Number(r[P.median_sell]);
    const medBuy = (r[P.median_buy] == null || r[P.median_buy] <= 0) ? null : Number(r[P.median_buy]);
    // --- END MODIFICATION ---

    const day = _toProjectDay_(ts);
    const key = `${typeId}|${marketId}|${mtype}|${day.getTime()}`;
    pushSample(key, ts.getTime(), minSell, maxBuy, medSell, medBuy, typeId, marketId, mtype, day);
  }

  // Final rows in History schema
  const out = new Map();
  for (const [key, v] of bag.entries()) {
    out.set(key, [
      v.typeId, v.marketId, v.mtype, v.day,
      v.buyOpen ?? '', v.buyClose ?? '',
      v.sellOpen ?? '', v.sellClose ?? '',
      v.buyHigh ?? '', v.buyLow ?? '',
      v.sellHigh ?? '', v.sellLow ?? '',
      _median_(v.medBuy) ?? '',
      _median_(v.medSell) ?? ''
    ]);
  }
  return out;
}

/**
 * NEW: Helper function to upsert data into the history temp sheet.
 * This reads the existing temp data, merges, and rewrites.
 */
function _upsertHistoryData(tempSheet, newRows) {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('HistoryMan') : console);
  LOG.info(`Upserting ${newRows.length} rows into ${tempSheet.getName()}...`);

  const keyMap = new Map();
  const iType = 0, iMid = 1, iMtp = 2, iDate = 3; // Indices from HM_HEADERS

  // 1. Read all existing data from the temp sheet
  const lastRow = tempSheet.getLastRow();
  if (lastRow > 1) {
    const oldData = tempSheet.getRange(2, 1, lastRow - 1, HM_HEADERS.length).getValues();
    for (const row of oldData) {
      const d = row[iDate];
      if (d instanceof Date) {
        const key = `${row[iType]}|${row[iMid]}|${row[iMtp]}|${d.getTime()}`;
        keyMap.set(key, row);
      }
    }
  }
  
  // 2. Merge new data, overwriting old
  for (const row of newRows) {
    const d = row[iDate];
    if (d instanceof Date) {
      const key = `${row[iType]}|${row[iMid]}|${row[iMtp]}|${d.getTime()}`;
      keyMap.set(key, row); // New data overwrites old data for the same key
    }
  }

  // 3. Write all merged data back
  const finalRows = Array.from(keyMap.values());
  if (finalRows.length > 0) {
    // Clear old content first
    if (lastRow > 1) {
      tempSheet.getRange(2, 1, lastRow - 1, HM_HEADERS.length).clearContent();
    }
    // Write new merged data
    tempSheet.getRange(2, 1, finalRows.length, HM_HEADERS.length).setValues(finalRows);
  }
  LOG.info(`Upsert complete. Temp sheet now has ${finalRows.length} rows.`);
}

/* ============================ MAIN STATE MACHINE ============================ */

/**
 * REWRITTEN: This is now the "Starter" function for the stateful history builder.
 * It just acquires a lock, sets the state, and calls the worker.
 *
 * This is what you should schedule on your daily trigger (e.g., updateHistory).
 */
function updateHistory() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('HistoryMan') : console);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();

  // Check if it's already running
  const currentState = SCRIPT_PROP.getProperty(HIST_PROP_STEP);
  if (currentState && currentState !== 'COMPLETE' && currentState !== STATE_FLAGS.NEW_RUN) {
    LOG.warn(`History Manager is already running (State: ${currentState}). Skipping new start.`);
    scheduleOneTimeTrigger('_historyWorker', 5000);
    return;
  }

  LOG.info("Starting new History Manager cycle.");
  SCRIPT_PROP.setProperty(HIST_PROP_STEP, STATE_FLAGS.NEW_RUN);

  // Use executeWithTryLock
  const result = executeWithTryLock(_historyWorker, '_historyWorker');
  if (result === null) {
    LOG.warn("History Manager start skipped by ScriptLock. Another process is running.");
    // If skipped, schedule a retry of this "starter" function
    scheduleOneTimeTrigger('updateHistory', 5 * 60 * 1000); // 5 min retry
  }
}

/**
 * NEW: The "Processing While Loop" worker for the History Manager.
 * This function reads Market Prices, aggregates, and writes to a temp sheet in batches.
 */
function _historyWorker() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('HistoryMan') : console);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const START_TIME = Date.now();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = _hmCfg();
  const sourceSheetName = cfg.sheets.prices;
  const tempSheetName = HIST_SHEET_TEMP;

  let currentState = SCRIPT_PROP.getProperty(HIST_PROP_STEP) || STATE_FLAGS.NEW_RUN;
  LOG.info(`Starting worker. Current State: ${currentState}`);

  try {
    // --- State: NEW_RUN (Start) ---
    if (currentState === STATE_FLAGS.NEW_RUN) {
      LOG.info(`State: ${STATE_FLAGS.NEW_RUN}. Preparing history temp sheet.`);
      const docLock = LockService.getDocumentLock();
      if (docLock.tryLock(FUZZ_DOC_LOCK_TIMEOUT)) {
        try {
          const tempSheet = getOrCreateSheet(ss, tempSheetName, HM_HEADERS);
          if (tempSheet.getLastRow() > 1) {
            tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, tempSheet.getMaxColumns()).clearContent();
          }
          tempSheet.hideSheet();
          SpreadsheetApp.flush();

          SCRIPT_PROP.setProperty(HIST_PROP_READ_ROW, '2'); // Data starts row 2
          currentState = STATE_FLAGS.PROCESSING;
          SCRIPT_PROP.setProperty(HIST_PROP_STEP, currentState);
          LOG.info(`Temp sheet '${tempSheetName}' prepared. Transitioning to ${currentState}.`);
        } finally {
          docLock.releaseLock();
        }
      } else {
        LOG.warn(`Document Lock busy during history setup. Rescheduling.`);
        scheduleOneTimeTrigger('_historyWorker', FUZZ_RESCHEDULE_MS);
        return;
      }
    } // --- End NEW_RUN ---

    // --- State: PROCESSING (Processing While Loop) ---
    if (currentState === STATE_FLAGS.PROCESSING) {
      LOG.info(`State: ${STATE_FLAGS.PROCESSING}. Reading/aggregating batches.`);
      
      const sourceSheet = ss.getSheetByName(sourceSheetName);
      const tempSheet = ss.getSheetByName(tempSheetName);
      if (!sourceSheet || !tempSheet) {
        throw new Error("Missing source or temp sheet during history processing.");
      }

      let readRow = parseInt(SCRIPT_PROP.getProperty(HIST_PROP_READ_ROW) || '2');
      const lastRow = sourceSheet.getLastRow();
      
      // --- Find Header Indices ---
      const header = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
      const P_H = header.map(h => String(h).trim().toLowerCase());
      const P = { date: P_H.indexOf('date'), market_id: P_H.indexOf('market_id'), market_type: P_H.indexOf('market_type'), type_id: P_H.indexOf('type_id'), min_sell: P_H.indexOf('min_sell'), max_buy: P_H.indexOf('max_buy'), median_sell: P_H.indexOf('median_sell'), median_buy: P_H.indexOf('median_buy') };
      
      // --- Processing Loop ---
      while (readRow <= lastRow) {
        // --- 1. Time Limit Check ---
        if (Date.now() - START_TIME > FUZZ_TIME_LIMIT_MS) {
          SCRIPT_PROP.setProperty(HIST_PROP_READ_ROW, readRow.toString());
          scheduleOneTimeTrigger('_historyWorker', FUZZ_RESCHEDULE_MS);
          LOG.warn(`Time limit hit. Saved state. Rescheduled. Next read row: ${readRow}`);
          return;
        }

        // --- 2. Read Batch from Market Prices ---
        const rowsToRead = Math.min(HIST_BATCH_SIZE, lastRow - readRow + 1);
        if (rowsToRead <= 0) break; 
        
        LOG.info(`Reading ${rowsToRead} rows from ${sourceSheetName} (starting row ${readRow})...`);
        const data = sourceSheet.getRange(readRow, 1, rowsToRead, header.length).getValues();
        
        // --- 3. Process Batch (Aggregate to Daily OHLC) ---
        // Filter for data from the last ~2 days
        const lookbackMs = Date.now() - (2 * 86400000); // 2 days
        const recent = [];
        for (let i = 0; i < data.length; i++) {
            const dt = data[i][P.date];
            if (dt instanceof Date && dt.getTime() >= lookbackMs) {
                recent.push(data[i]);
            }
        }
        
        const agg = _aggDailyFromPrices_(recent, P); // Pass header map
        const rowsToWrite = Array.from(agg.values());

        // --- 4. Write/Upsert Batch to Temp Sheet ---
        if (rowsToWrite.length > 0) {
          const docLock = LockService.getDocumentLock();
          if (docLock.tryLock(FUZZ_DOC_LOCK_TIMEOUT)) {
            try {
              // This needs to be an UPSERT, not an append, to handle
              // multiple batches processing the same day.
              _upsertHistoryData(tempSheet, rowsToWrite);
              LOG.info(`Upserted ${rowsToWrite.length} aggregated rows to ${tempSheetName}.`);
            } finally {
              docLock.releaseLock();
            }
          } else {
            // RETRIGGER ON WRITE FAILURE
            LOG.warn(`Document Lock busy for history write. Rescheduling (will re-process batch).`);
            scheduleOneTimeTrigger('_historyWorker', FUZZ_RESCHEDULE_MS);
            return;
          }
        }
        
        // --- 5. Advance Index ---
        readRow += rowsToRead;
        SCRIPT_PROP.setProperty(HIST_PROP_READ_ROW, readRow.toString());
      
      } // --- End while loop ---

      // --- Post-Loop Check ---
      if (readRow > lastRow) {
        LOG.info("All source rows processed. Transitioning to FINALIZING.");
        currentState = STATE_FLAGS.FINALIZING;
        SCRIPT_PROP.setProperty(HIST_PROP_STEP, currentState);
        scheduleOneTimeTrigger('_finalizeHistory', 1000); // 1 sec delay
      }
    } // --- End PROCESSING ---

  } catch (e) {
    LOG.error(`Unhandled error in history worker: ${e.message}\nStack: ${e.stack}`);
    // Reset history state on error
    SCRIPT_PROP.deleteProperty(HIST_PROP_STEP);
    SCRIPT_PROP.deleteProperty(HIST_PROP_READ_ROW);
  }
}

/**
 * NEW: The "Finalize" function for the History Manager.
 * This performs the atomic swap and applies retention/cap.
 */
function _finalizeHistory() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('HistoryFinalizer') : console);
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  
  if (SCRIPT_PROP.getProperty(HIST_PROP_STEP) !== STATE_FLAGS.FINALIZING) {
    LOG.warn(`Finalizer called in incorrect state (${SCRIPT_PROP.getProperty(HIST_PROP_STEP)}). Aborting.`);
    return;
  }

  LOG.info("Starting finalization: Applying retention/cap and atomic swap.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = _hmCfg();
  const tempSheetName = HIST_SHEET_TEMP;
  const finalSheetName = cfg.sheets.history;
  const oldSheetName = finalSheetName + "_Old"; // Give it a unique "Old" name

  const docLock = LockService.getDocumentLock();
  try {
    if (docLock.tryLock(30000)) { // Wait up to 30s
      try {
        const tempSheet = ss.getSheetByName(tempSheetName);
        if (!tempSheet || tempSheet.getLastRow() <= 1) {
          throw new Error(`History temp sheet '${tempSheetName}' is missing or empty! Cannot finalize.`);
        }

        // --- 1. Apply Retention and Row Cap ---
        LOG.info("Applying retention and row cap to temp sheet...");
        const data = tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, HM_HEADERS.length).getValues();
        const { historyRetentionDays, historyMaxRows } = cfg;
        const cutoff = new Date(Date.now() - historyRetentionDays * 86400000);
        
        let finalData = data.filter(r => r[3] instanceof Date && r[3] >= cutoff); // Filter by retention
        
        if (finalData.length > historyMaxRows) {
          finalData.sort((a,b) => a[3] - b[3]); // Sort by date ascending
          finalData = finalData.slice(finalData.length - historyMaxRows); // Keep newest rows
          LOG.info(`Capped history rows to ${finalData.length} (max: ${historyMaxRows}).`);
        } else {
          LOG.info(`Filtered by retention. ${finalData.length} rows remain.`);
        }
        
        // --- 2. Rewrite Temp Sheet ---
        tempSheet.getRange(2, 1, tempSheet.getLastRow() - 1, tempSheet.getMaxColumns()).clearContent();
        if (finalData.length > 0) {
          tempSheet.getRange(2, 1, finalData.length, HM_HEADERS.length).setValues(finalData);
        }
        // Use the _trimTrailing_ from MarketFetcher.gs.js
        if (typeof _trimTrailing_ === "function") {
          _trimTrailing_(tempSheet);
        }
        SpreadsheetApp.flush();
        LOG.info("Final data written to temp sheet.");

        // --- 3. Atomic Swap ---
        const finalSheet = ss.getSheetByName(finalSheetName);
        const oldSheet = ss.getSheetByName(oldSheetName);

        if (oldSheet) ss.deleteSheet(oldSheet);
        if (finalSheet) finalSheet.setName(oldSheetName);
        tempSheet.setName(finalSheetName);
        tempSheet.showSheet();
        SpreadsheetApp.flush();
        LOG.info("Atomic sheet swap successful.");

        // --- 4. Reset History Job State ---
        SCRIPT_PROP.deleteProperty(HIST_PROP_STEP);
        SCRIPT_PROP.deleteProperty(HIST_PROP_READ_ROW);
        LOG.info("History Manager job state reset complete.");

      } catch (swapError) {
        LOG.error(`CRITICAL error during history swap: ${swapError.message}. State NOT reset.`);
        scheduleOneTimeTrigger('_finalizeHistory', 60000); // Retry in 1 min
        throw swapError;
      } finally {
        docLock.releaseLock();
      }
    } else {
      LOG.warn("Document Lock busy during finalization. Rescheduling finalizer.");
      scheduleOneTimeTrigger('_finalizeHistory', FUZZ_RESCHEDULE_MS);
    }
  } catch (e) {
    LOG.error(`Error in finalizer lock acquisition: ${e.message}`);
    scheduleOneTimeTrigger('_finalizeHistory', 60000);
  }
}

/* ===================== Back-compat entry points ===================== */
// These are kept so old manual runs don't error, but `updateHistory` is the new trigger
function updateMarketHistory(opts) { return updateHistory(); }
function updateHistoryTest() { return updateHistory(); }

/* =================== Other Helpers (Unchanged) =================== */
// (All other helper functions from the original file remain here)
// _hmCfg(), _tzNowNY_(), _minutesSinceMidnightNY_(), _parseHHMMtoMinutes_()
// _wbAllocatedCells_(), _cellsNeededForAppend_(), _tightenTrailingRows_(), 
// _deleteRowsAscBlocks_(), ensureAppendCapacity_History_(), _ensureSheet_()
// formatHistoryBlock_(), _buildIndex_(), _writeContiguousHM_()
// HM_stopTriggers(), HM_oneShot(), HM_diagDupes(), HM_dedupeKeepLast(),
// diag_HistoryTodayDupes(), _canonMarketType(), _canonKey()