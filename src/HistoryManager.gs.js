/***********************
 * HistoryManager.gs
 * - Builds daily history from "Market Prices"
 * - Day High  = max(min_sell) over last 24h
 * - Day Low   = min(max_buy)  over last 24h
 * - Open/Close snapshots from getMarketPrices(...)
 * - Retention via HistoryRetentionDays (default 365)
 * - Production sheet: "Market History"
 * - Test sheet:       "History (TEST)"
 ***********************/

const HISTORY_PROD_SHEET = "Market History";
const HISTORY_TEST_SHEET = "History (TEST)";

/**
 * Production entrypoint – safe for triggers.
 * Uses "auto" mode (respects open/close time windows).
 */
function updateHistory() {
  if (AbortCloseIfNoHistoryYet()) return; 
  return _updateHistoryCore({
    testMode: false,
    mode: "auto"
  });
}

/**
 * Test entrypoint – writes to "History (TEST)".
 * Defaults to "auto", or pass "open"/"close" for debugging.
 */
function updateHistoryTestMode(modeOverride) {
  return _updateHistoryCore({
    testMode: true,
    mode: modeOverride || "auto"
  });
}

/**
 * Core: fetch → build rows → append → prune.
 * - testMode: writes to "History (TEST)" instead of prod history sheet.
 * - Daily extremes per key (type_id|market_id|market_type):
 *    - buy_open  = first max_buy seen today
 *    - buy_close = current snapshot max_buy
 *    - sell_open = first min_sell seen today
 *    - sell_close= current snapshot min_sell
 *    - daily_high= max(min_sell) today (sell ceiling)
 *    - daily_low = min(max_buy)  today (buy floor)
 */
function _updateHistoryCore(opts) {
  const testMode = !!(opts && opts.testMode);
  const mode = (opts && opts.mode) || "auto";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();

  const maxLogIDs     = parseInt(config["MaxLogIDs"] || 750, 10);
  const retentionDays = parseInt(config["HistoryRetentionDays"] || 365, 10);

  const now = new Date();

  const phase = _determinePhase(config, mode, now); // 1-hour trigger windows
  if (!phase.allowed) {
    Logger.log("[updateHistory] Skipped: outside Open/Close trigger windows (auto mode).");
    return;
  }

  const typeIDs      = getTypeIDsFromItemList(maxLogIDs);
  const marketCombos = getMarketSettings();

  // ---- 1) Fetch fresh market data into a flat list
  const marketData = [];
  marketCombos.forEach(({ market_id, market_type }) => {
    const prices = getMarketPrices(typeIDs, market_id, market_type);
    typeIDs.forEach(type_id => {
      const p = prices[type_id] || {};
      marketData.push({
        date: now,
        type_id,
        market_id,
        market_type,
        min_sell:    p.minSell ?? null,     // sell side price
        max_buy:     p.maxBuy  ?? null,     // buy side price
        median_sell: p.medianSell ?? null,
        median_buy:  p.medianBuy  ?? null
      });
    });
  });

// ---- 2) Upsert today's rows (respect open/close phase)
  const targetSheetName = testMode ? HISTORY_TEST_SHEET : HISTORY_PROD_SHEET;
  _ensureHistoryHeader(targetSheetName, [
    "type_id","market_id","market_type","date",
    "buy_open","buy_close","sell_open","sell_close",
    "daily_high","daily_low","median_buy","median_sell"
  ]);

  const targetSheet = ss.getSheetByName(targetSheetName);
  _upsertHistoryRows(targetSheet, marketData, phase);

  // ---- 3) Prune (date is column 4 in History)
  pruneOldRows(targetSheet, retentionDays, 4);
}

/**
 * Build daily rows with dual-side open/close + daily_high/low.
 * phase = { isOpenRun, isCloseRun }
 * Seeds today's state from the PROD sheet so opens persist even when testing.
 */
function _buildDailyRowsForHistory(marketData, phase) {
  if (!marketData || marketData.length === 0) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName(HISTORY_PROD_SHEET); // seed from prod

  const today0 = new Date(marketData[0].date); today0.setHours(0,0,0,0);

  // state[key] = { buy_open, sell_open, high, low }
  const state = new Map();

  if (prodSheet && prodSheet.getLastRow() > 1) {
    const headers = prodSheet.getRange(1,1,1,prodSheet.getLastColumn()).getValues()[0];
    const idx = {}; headers.forEach((h,i)=> idx[h]=i);

    const required = ["type_id","market_id","market_type","date","buy_open","sell_open","daily_high","daily_low"];
    const haveAll = required.every(h => h in idx);

    if (haveAll) {
      const data = prodSheet.getRange(2,1,prodSheet.getLastRow()-1,prodSheet.getLastColumn()).getValues();
      for (let i=0;i<data.length;i++) {
        const r = data[i];
        const d = new Date(r[idx["date"]]); d.setHours(0,0,0,0);
        if (d.getTime() !== today0.getTime()) continue;

        const key = `${r[idx["type_id"]]}|${r[idx["market_id"]]}|${r[idx["market_type"]]}`;
        const cur = state.get(key) || { buy_open: null, sell_open: null, high: null, low: null };

        const buy_open  = r[idx["buy_open"]];
        const sell_open = r[idx["sell_open"]];
        const hi        = r[idx["daily_high"]];
        const lo        = r[idx["daily_low"]];

        if (typeof buy_open  === "number" && cur.buy_open  == null) cur.buy_open  = buy_open;
        if (typeof sell_open === "number" && cur.sell_open == null) cur.sell_open = sell_open;
        if (typeof hi        === "number") cur.high = cur.high == null ? hi : Math.max(cur.high, hi);
        if (typeof lo        === "number") cur.low  = cur.low  == null ? lo : Math.min(cur.low,  lo);

        state.set(key, cur);
      }
    }
  }

  // Build new snapshot rows + update today's state by phase
  const out = [];
  for (const rec of marketData) {
    const key = `${rec.type_id}|${rec.market_id}|${rec.market_type}`;
    const cur = state.get(key) || { buy_open: null, sell_open: null, high: null, low: null };

    // Always update extremes
    if (typeof rec.min_sell === "number") cur.high = cur.high == null ? rec.min_sell : Math.max(cur.high, rec.min_sell);
    if (typeof rec.max_buy  === "number") cur.low  = cur.low  == null ? rec.max_buy  : Math.min(cur.low,  rec.max_buy);

    // OPEN phase: set opens if not already set; do NOT set closes
    if (phase.isOpenRun) {
      if (cur.buy_open  == null && typeof rec.max_buy  === "number") cur.buy_open  = rec.max_buy;
      if (cur.sell_open == null && typeof rec.min_sell === "number") cur.sell_open = rec.min_sell;
    }

    // CLOSE phase: set closes from this snapshot; do NOT change opens
    const buy_close  = phase.isCloseRun && (typeof rec.max_buy  === "number") ? rec.max_buy  : null;
    const sell_close = phase.isCloseRun && (typeof rec.min_sell === "number") ? rec.min_sell : null;

    state.set(key, cur);

    out.push([
      rec.type_id,
      rec.market_id,
      rec.market_type,
      rec.date,
      cur.buy_open,
      buy_close,
      cur.sell_open,
      sell_close,
      cur.high,
      cur.low,
      rec.median_buy,
      rec.median_sell
    ]);
  }

  return out;
}

/** Ensure target history sheet exists with exact headers, limited to header count on create. */
function _ensureHistoryHeader(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet(ss, sheetName, headers);
}

/** Batch-append rows to a sheet (no header changes). */
function appendRowsToSheet(sheetName, rows) {
  if (!rows || rows.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/** --------- Time-window helpers (1-hour trigger windows) --------- */

function _inWindow_(now, startH, startM, durationMin) {
  const start = new Date(now);
   start.setHours(startH, startM, 0, 0);
  const end   = new Date(start.getTime() + durationMin * 60 * 1000);
  return now >= start && now < end;
}

/**
 * Returns { isOpenRun, isCloseRun, allowed }
 * - mode "auto": allowed only if within the 60m window after OpenTime or CloseTime
 * - mode "open"/"close": forced, always allowed
 */
function _determinePhase(config, mode, now) {
  if (mode === "open")  return { isOpenRun: true,  isCloseRun: false, allowed: true };
  if (mode === "close") return { isOpenRun: false, isCloseRun: true,  allowed: true };

  const [oH, oM] = _toHM(config.OpenTime  || "11:00");
  const [cH, cM] = _toHM(config.CloseTime || "18:00");
  const DUR = 60; // minutes

  const inOpen  = _inWindow_(now, oH, oM, DUR);
  const inClose = _inWindow_(now, cH, cM, DUR);

  if (inOpen)  return { isOpenRun: true,  isCloseRun: false, allowed: true  };
  if (inClose) return { isOpenRun: false, isCloseRun: true,  allowed: true  };
  return { isOpenRun: false, isCloseRun: false, allowed: false };
}

/**
 * Upserts today's history rows into target sheet.
 * - OPEN pass: set buy_open/sell_open (first seen), closes null; update daily_high/low.
 * - CLOSE pass: set buy_close/sell_close (current snapshot); update daily_high/low.
 * - If no open row exists at close time, create a single row with open=close.
 */
function _upsertHistoryRows(sheet, marketData, phase) {
  if (!marketData || marketData.length === 0) return;

  // Header index map
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const idx = {}; headers.forEach((h,i)=> idx[h]=i+1); // 1-based for setValue convenience

  // Build index of today's existing rows by key
  const today0 = new Date(marketData[0].date); today0.setHours(0,0,0,0);
  const lastRow = sheet.getLastRow();
  const rowMap = new Map(); // key -> rowIndex
  if (lastRow > 1) {
    const data = sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getValues();
    for (let r=0; r<data.length; r++) {
      const row = data[r];
      const d = new Date(row[idx["date"]-1]); d.setHours(0,0,0,0);
      if (d.getTime() !== today0.getTime()) continue;
      const key = _keyForRow(row, idx);
      rowMap.set(key, r+2); // store 1-based row index in sheet
    }
  }

  // For performance, accumulate appends and do one batch append
  const appends = [];

  for (const rec of marketData) {
    const key = `${rec.type_id}|${rec.market_id}|${rec.market_type}`;
    let rowIndex = rowMap.get(key);

    // Compute high/low incrementally
    const highVal = (typeof rec.min_sell === "number") ? rec.min_sell : null; // sell-side high
    const lowVal  = (typeof rec.max_buy  === "number") ? rec.max_buy  : null; // buy-side low

    if (!rowIndex) {
      // No row yet for today → create one based on phase
      if (phase.isOpenRun) {
        appends.push([
          rec.type_id, rec.market_id, rec.market_type, rec.date,
          (typeof rec.max_buy  === "number") ? rec.max_buy  : null, // buy_open
          null,                                                     // buy_close
          (typeof rec.min_sell === "number") ? rec.min_sell : null, // sell_open
          null,                                                     // sell_close
          highVal,                                                  // daily_high
          lowVal,                                                   // daily_low
          rec.median_buy, rec.median_sell
        ]);
      } else {
        // Close run but open row missing (edge case): set open=close to current snapshot
        const buySnap  = (typeof rec.max_buy  === "number") ? rec.max_buy  : null;
        const sellSnap = (typeof rec.min_sell === "number") ? rec.min_sell : null;
        appends.push([
          rec.type_id, rec.market_id, rec.market_type, rec.date,
          buySnap,                         // buy_open (seed)
          buySnap,                         // buy_close
          sellSnap,                        // sell_open (seed)
          sellSnap,                        // sell_close
          highVal,                         // daily_high
          lowVal,                          // daily_low
          rec.median_buy, rec.median_sell
        ]);
      }
      continue; // handle updates after append batch
    }

    // Row exists for today → update columns in-place
    // Read existing values needed for high/low carry
    const existing = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Derive new high/low
    let newHigh = existing[idx["daily_high"]-1];
    let newLow  = existing[idx["daily_low"]-1];
    if (typeof highVal === "number") newHigh = (typeof newHigh === "number") ? Math.max(newHigh, highVal) : highVal;
    if (typeof lowVal  === "number") newLow  = (typeof newLow  === "number") ? Math.min(newLow,  lowVal)  : lowVal;

    // OPEN: set opens only if not set; CLOSE: set closes to snapshot
    if (phase.isOpenRun) {
      if (existing[idx["buy_open"]-1]  == null && typeof rec.max_buy  === "number")
        sheet.getRange(rowIndex, idx["buy_open"]).setValue(rec.max_buy);
      if (existing[idx["sell_open"]-1] == null && typeof rec.min_sell === "number")
        sheet.getRange(rowIndex, idx["sell_open"]).setValue(rec.min_sell);
    } else if (phase.isCloseRun) {
      if (typeof rec.max_buy  === "number")
        sheet.getRange(rowIndex, idx["buy_close"]).setValue(rec.max_buy);
      if (typeof rec.min_sell === "number")
        sheet.getRange(rowIndex, idx["sell_close"]).setValue(rec.min_sell);
    }

    // Always refresh highs/lows and medians to latest known
    sheet.getRange(rowIndex, idx["daily_high"]).setValue(newHigh);
    sheet.getRange(rowIndex, idx["daily_low"]).setValue(newLow);
    sheet.getRange(rowIndex, idx["median_buy"]).setValue(rec.median_buy);
    sheet.getRange(rowIndex, idx["median_sell"]).setValue(rec.median_sell);
  }

  // Perform one batch append if needed, then update rowMap for those new rows (not strictly necessary today)
  if (appends.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, appends.length, headers.length).setValues(appends);
    // (Optionally populate rowMap for appended rows if you upsert multiple passes in one invocation)
  }
}

/** Build the row key from an existing row using header indices. */
function _keyForRow(row, idx) {
  return `${row[idx["type_id"]-1]}|${row[idx["market_id"]-1]}|${row[idx["market_type"]-1]}`;
}

// Call near the top of updateHistory()
function AbortCloseIfNoHistoryYet() {
  const cfg = getConfig();
  const phase = _determinePhase(cfg, "auto", new Date());
  if (!phase.allowed || !phase.isCloseRun) return false; // only care during close window

  const histName = cfg.HistorySheetName || "Market History";
  const sh = SpreadsheetApp.getActive().getSheetByName(histName);
  if (!sh) {
    Logger.log(`[ABORT] Close run aborted: "${histName}" does not exist (no open-run creation yet).`);
    return true; // signal caller to abort
  }
  return false;
}

// Uses your Config + Utility._toHM and stays in Project TZ
function projectTZ_() {
  const c = getConfig();
  return c["TZ.Project"] || SpreadsheetApp.getActive().getSpreadsheetTimeZone();
}

// Build today's window [OpenTime .. now] in Project TZ
function projectDayWindowNow_() {
  const c = getConfig();
  const tz = projectTZ_();

  const now = new Date();
  const y  = Number(Utilities.formatDate(now, tz, "yyyy"));
  const M  = Number(Utilities.formatDate(now, tz, "MM")) - 1;
  const d  = Number(Utilities.formatDate(now, tz, "dd"));

  const { h: oH, m: oM } = _toHM(c["OpenTime"] || "11:00");  // <-- your Utility.js
  const start = new Date(y, M, d, oH, oM, 0, 0);              // Date in Project TZ

  return { start, endNow: now };
}

// Build today's full window [OpenTime .. CloseTime] in Project TZ
function projectDayWindowFull_() {
  const c = getConfig();
  const tz = projectTZ_();

  const now = new Date();
  const y  = Number(Utilities.formatDate(now, tz, "yyyy"));
  const M  = Number(Utilities.formatDate(now, tz, "MM")) - 1;
  const d  = Number(Utilities.formatDate(now, tz, "dd"));

  const { h: oH, m: oM } = _toHM(c["OpenTime"]  || "11:00");
  const { h: cH, m: cM } = _toHM(c["CloseTime"] || "18:00");

  const start  = new Date(y, M, d, oH, oM, 0, 0);
  const endDay = new Date(y, M, d, cH, cM, 59, 999);

  return { start, endDay };
}

// Date-only (00:00) in Project TZ for one-row-per-day keys
function dateOnlyProjectTZ_(dt) {
  const tz = projectTZ_();
  const y  = Number(Utilities.formatDate(dt, tz, "yyyy"));
  const M  = Number(Utilities.formatDate(dt, tz, "MM")) - 1;
  const d  = Number(Utilities.formatDate(dt, tz, "dd"));
  return new Date(y, M, d); // midnight in Project TZ
}

function isSameProjectDay_(a, bDateOnly) {
  if (!(a instanceof Date) || !(bDateOnly instanceof Date)) return false;
  const tz = projectTZ_();
  const A = Utilities.formatDate(a,       tz, "yyyy-MM-dd");
  const B = Utilities.formatDate(bDateOnly, tz, "yyyy-MM-dd");
  return A === B;
}