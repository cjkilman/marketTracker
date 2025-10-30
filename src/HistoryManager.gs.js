/** HistoryManager.gs — STATEFUL Daily OHLC/median builder
 *
 * REWRITTEN to be a stateful, batch-processing worker
 * to avoid 6-minute timeouts on large data.
 *
 * Source sheet:  "Market Prices"
 * Target sheet:  "Market History"
 */

/* global LockService, PropertiesService, SpreadsheetApp, LoggerEx, 
   executeWithTryLock, scheduleOneTimeTrigger, getOrCreateSheet, _trimTrailing_,
   FUZZ_TIME_LIMIT_MS, FUZZ_RESCHEDULE_MS, FUZZ_DOC_LOCK_TIMEOUT, PT, mtConfig */

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
  // REFACTORED: Use strings directly, remove STATE_FLAGS dependency
  if (currentState && currentState !== 'COMPLETE' && currentState !== "NEW_RUN") {
    LOG.warn(`History Manager is already running (State: ${currentState}). Skipping new start.`);
    scheduleOneTimeTrigger('_historyWorker', 5000);
    return;
  }

  LOG.info("Starting new History Manager cycle.");
  // REFACTORED: Use strings directly
  SCRIPT_PROP.setProperty(HIST_PROP_STEP, "NEW_RUN");

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

  // REFACTORED: Use strings directly
  let currentState = SCRIPT_PROP.getProperty(HIST_PROP_STEP) || "NEW_RUN";
  LOG.info(`Starting worker. Current State: ${currentState}`);

  try {
    // --- State: NEW_RUN (Start) ---
    // REFACTORED: Use strings directly
    if (currentState === "NEW_RUN") {
      LOG.info(`State: NEW_RUN. Preparing history temp sheet.`);
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
          // REFACTORED: Use strings directly
          currentState = "PROCESSING";
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
    // REFACTORED: Use strings directly
    if (currentState === "PROCESSING") {
      LOG.info(`State: PROCESSING. Reading/aggregating batches.`);
      
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
        // Filter for data from the last ~2 days to catch day boundaries
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
        // REFACTORED: Use strings directly
        currentState = "FINALIZING";
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
  
  // REFACTORED: Use strings directly
  if (SCRIPT_PROP.getProperty(HIST_PROP_STEP) !== "FINALIZING") {
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


/* =================== Optional maintenance helpers (UNCHANGED) =================== */
// (These functions are unchanged from your original file)

function _wbAllocatedCells_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().reduce((sum, sh) => sum + sh.getMaxRows() * sh.getMaxColumns(), 0);
}

function _cellsNeededForAppend_(sh, addRows, needCols) {
  const cols = Math.max(needCols, sh.getMaxColumns());
  return addRows * cols;
}

function _tightenTrailingRows_(sh, rowBuffer) {
  if (!sh) return; // Add guard
  rowBuffer = rowBuffer || 2000;
  const used = Math.max(1, sh.getLastRow());
  const alloc = sh.getMaxRows();
  const keep = Math.max(used + rowBuffer, Math.min(alloc, used + rowBuffer));
  const extra = alloc - keep;
  if (extra > 0) {
     try { sh.deleteRows(keep + 1, extra); } catch(e) {}
  }
}

function _deleteRowsAscBlocks_(sh, rowsAsc) {
  if (!rowsAsc || !rowsAsc.length) return;
  for (let i = rowsAsc.length - 1; i >= 0; i--) {
    sh.deleteRow(rowsAsc[i]);
  }
}

function ensureAppendCapacity_History_(sh, addRows, needCols, retentionDays, dateColIdx1) {
  if (addRows <= 0) return;
  if (retentionDays > 0) {
    const used = sh.getLastRow();
    if (used > 1) {
      const iDate0 = (dateColIdx1 || 4) - 1;
      const lastCol = sh.getLastColumn();
      const data = sh.getRange(2, 1, used - 1, lastCol).getValues();
      const cutoff = new Date(Date.now() - retentionDays * 24 * 60 * 60 * 1000);
      const del = [];
      for (let i = 0; i < data.length; i++) {
        const d = data[i][iDate0];
        if (d instanceof Date && d < cutoff) del.push(i + 2);
      }
      if (del.length) _deleteRowsAscBlocks_(sh, del);
    }
  }
  _tightenTrailingRows_(sh, 2000);
  const CAP = (typeof WORKBOOK_CAP === 'number') ? WORKBOOK_CAP : 10000000;
  const SAFETY = (typeof WORKBOOK_SAFETY === 'number') ? WORKBOOK_SAFETY : 200000;
  const want = _cellsNeededForAppend_(sh, addRows, needCols);
  const have = Math.max(0, CAP - SAFETY - _wbAllocatedCells_());
  if (want <= have) return;
  const cols = Math.max(needCols, sh.getMaxColumns());
  const maxExtraRows = Math.floor(have / Math.max(1, cols));
  const msg = [
    '[History] Workbook cap guard: insufficient capacity.',
    'Need ~' + want.toLocaleString() + ' cells, free ~' + have.toLocaleString() + '.',
    'At current width (' + cols + ' cols), max additional rows: ~' + maxExtraRows.toLocaleString() + '.',
    'Shorten history retention, trim other sheets, or run heavy prunes separately.'
  ].join(' ');
  (LoggerEx?.error || console.error)(msg);
  throw new Error(msg);
}

function _ensureSheet_(ss, name, headers) {
  if (typeof getOrCreateSheet === 'function') {
    return getOrCreateSheet(ss, name, headers);
  }
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && headers.length) {
    const hdr = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    const needs = !hdr[0] || headers.some((h, i) => hdr[i] !== h);
    if (needs) {
      sh.clear();
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function formatHistoryBlock_(sh, startRow, nRows, idx1) {
  if (!nRows) return;
  const nfInt = '#,##0';
  const nfDate = 'yyyy-mm-dd';
  sh.getRange(startRow, idx1.date, nRows, 1).setNumberFormat(nfDate);
  const numericCols = [
    idx1.buy_open, idx1.buy_close, idx1.sell_open, idx1.sell_close,
    idx1.buy_high, idx1.buy_low, idx1.sell_high, idx1.sell_low,
    idx1.median_buy, idx1.median_sell
  ];
  numericCols.forEach(c => sh.getRange(startRow, c, nRows, 1).setNumberFormat(nfInt));
}

function _buildIndex_(vals) {
  if (!vals || vals.length === 0) return { idx: new Map(), cols: 0 };
  const H = vals[0]; const cols = H.length;
  const iType = H.indexOf('type_id');
  const iMid = H.indexOf('market_id');
  const iMtp = H.indexOf('market_type');
  const iDate = H.indexOf('date');
  const toNYMidnight = (v) => {
    if (v instanceof Date) return _toProjectDay_(v);
    const s = String(v || '').trim();
    if (!s) return null;
    const n = Number(s);
    if (Number.isFinite(n)) {
      const d = new Date(Math.round((n - 25569) * 86400000));
      return _toProjectDay_((d instanceof Date ? d : new Date(d)));
    }
    const t = Date.parse(s);
    if (isNaN(t)) return null;
    return _toProjectDay_(new Date(t));
  };
  const idx = new Map();
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const t = Number(row[iType]);
    const m = Number(row[iMid]);
    const mt = String(row[iMtp] || '').trim();
    const d = toNYMidnight(row[iDate]);
    if (!Number.isFinite(t) || !Number.isFinite(m) || !mt || !d) continue;
    idx.set(`${t}|${m}|${mt}|${+d}`, r + 1);
  }
  return { idx, cols };
}

function _writeContiguousHM_(sh, updates, totalCols) {
  if (!updates.length) return;
  updates.sort((a, b) => a.rn - b.rn);
  for (let i = 0; i < updates.length;) {
    const start = updates[i].rn;
    const block = [updates[i].row];
    let j = i + 1;
    while (j < updates.length && updates[j].rn === updates[j - 1].rn + 1) {
      block.push(updates[j].row);
      j++;
    }
    sh.getRange(start, 1, block.length, totalCols).setValues(block);
    i = j;
  }
}

function HM_stopTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['HM_update', 'updateHistory', 'updateHistoryTest', '_historyWorker', '_finalizeHistory'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function HM_oneShot() { updateHistory(); } // Just call the new starter

function HM_diagDupes(limit) {
  limit = Number(limit || 10);
  const sh = SpreadsheetApp.getActive().getSheetByName('Market History');
  if (!sh) { console.log('no Market History'); return; }
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) { console.log('empty'); return; }
  const H = vals[0]; const iT=H.indexOf('type_id'), iM=H.indexOf('market_id'), iMt=H.indexOf('market_type'), iD=H.indexOf('date');
  const toNY = (v)=> {
    if (v instanceof Date) return _toProjectDay_(v);
    const s=String(v||'').trim(); if(!s) return null;
    const n=Number(s); if (Number.isFinite(n)) return _toProjectDay_(new Date(Math.round((n-25569)*86400000)));
    const t=Date.parse(s); if (isNaN(t)) return null;
    return _toProjectDay_(new Date(t));
  };
  const ct = {};
  for (let r=1;r<vals.length;r++){
    const d = toNY(vals[r][iD]); if(!d) continue;
    const key = `${Math.floor(vals[r][iT])}|${Math.floor(vals[r][iM])}|${String(vals[r][iMt]).trim()}|${+d}`;
    ct[key]=(ct[key]||0)+1;
  }
  const top = Object.entries(ct).filter(([,c])=>c>1).sort((a,b)=>b[1]-a[1]).slice(0,limit);
  console.log('Top dup keys (key,count):', top);
}

function HM_dedupeKeepLast() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Market History');
  if (!sh) return;
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return;
  const H = vals[0]; const iT=H.indexOf('type_id'), iM=H.indexOf('market_id'), iMt=H.indexOf('market_type'), iD=H.indexOf('date');
  const toNY = (v)=> {
    if (v instanceof Date) return _toProjectDay_(v);
    const s=String(v||'').trim(); if(!s) return null;
    const n=Number(s); if (Number.isFinite(n)) return _toProjectDay_(new Date(Math.round((n-25569)*86400000)));
    const t=Date.parse(s); if (isNaN(t)) return null;
    return _toProjectDay_(new Date(t));
  };
  const keepRow = new Map();
  for (let r=1;r<vals.length;r++){
    const d = toNY(vals[r][iD]); if (!d) continue;
    const key = `${Math.floor(vals[r][iT])}|${Math.floor(vals[r][iM])}|${String(vals[r][iMt]).trim()}|${+d}`;
    keepRow.set(key, r+1); // last occurrence wins
  }
  const deleteRows = [];
  const kept = new Set(keepRow.values());
  for (let r=2;r<=vals.length;r++){
    if (!kept.has(r)) deleteRows.push(r);
  }
  deleteRows.sort((a,b)=>a-b);
  for (let i=deleteRows.length-1;i>=0;i--) sh.deleteRow(deleteRows[i]);
  console.log('HM_dedupeKeepLast removed:', deleteRows.length);
}

function diag_HistoryTodayDupes() {
  const cfg = _hmCfg();
  const sh = SpreadsheetApp.getActive().getSheetByName(cfg.sheets.history);
  if (!sh) return Logger.log('History sheet not found');
  const vals = sh.getDataRange().getValues();
  const h = vals[0];
  const I = { type: h.indexOf('type_id'), mid: h.indexOf('market_id'),
              mtp: h.indexOf('market_type'), date: h.indexOf('date') };
  const today = _toProjectDay_(PT.now());
  const seen = new Map(), dupes = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const d0 = row[I.date]; if (!(d0 instanceof Date)) continue;
    const d = _toProjectDay_(d0); if (+d !== +today) continue;
    const key = `${row[I.type]}|${row[I.mid]}|${String(row[I.mtp]).trim().toLowerCase()}|${+d}`;
    if (seen.has(key)) dupes.push({ rowNumber: r+1, key });
    else seen.set(key, r+1);
  }
  Logger.log({ today, dupes: dupes.length });
  return dupes;
}

function _canonMarketType(v) { return String(v || '').trim().toLowerCase(); }
function _canonKey(typeId, marketId, marketType, date) {
  return `${Math.floor(typeId)}|${Math.floor(marketId)}|${_canonMarketType(marketType)}|${+_toProjectDay_(date)}`;
}