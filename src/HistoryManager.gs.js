/** HistoryManager.gs — Daily OHLC/median builder with cap guard (no auto-prune)
 *
 * Depends on:
 *   - WORKBOOK_CAP, WORKBOOK_SAFETY (defined globally, e.g., in Utility.js)
 *   - LoggerEx (optional; falls back to console)
 *   - (optional) getConfig()/mtConfig() for settings
 *
 * Source sheet:  "Market Prices"
 *   [date, market_id, market_type, type_id, min_sell, max_buy, median_sell, median_buy]
 *
 * Target sheet:  "Market History"
 *   [type_id, market_id, market_type, date,
 *    buy_open, buy_close, sell_open, sell_close,
 *    buy_high, buy_low, sell_high, sell_low,
 *    median_buy, median_sell]
 *
 * Heavy prunes should be run separately.
 */

// TODO: Address updateHistory() Midday Append creep. Option: replace with ESI market History for Candle Sticks

/* ============================ Config ============================ */

function _hmCfg() {
  const c = (typeof getConfig === 'function') ? (getConfig() || {}) : {};
  const num = (v, d) => (v == null || isNaN(Number(v))) ? d : Number(v);
  const str = (v, d) => (v == null || v === '') ? d : String(v);
  const bool = (v, d=false) => {
    const s = String(v==null? '' : v).trim().toUpperCase();
    return s === 'TRUE' ? true : (s === 'FALSE' ? false : d);
  };

  // Retention: prefer HistoryRetentionDays; fall back to HistoryDaysLimit
  const retentionDays = num(c["HistoryRetentionDays"], num(c["HistoryDaysLimit"], 365));

  return {
    sheets: {
      prices: str(c["MarketPricesSheet"], "Market Prices"),
      history: str(c["HistorySheetName"], "Market History"),
    },
    // Slices & schedule
    openTime: str(c["OpenTime"], "11:00"),
    closeTime: str(c["CloseTime"], "18:00"),
    openWindowMins: num(c["OpenWindowMins"], 30),
    closeWindowMins: num(c["CloseWindowMins"], 30),
    sliceHours: num(c["SliceHours"], 3),                 // “comfort” slices
    lightSlices: bool(c["History.LightSlices"], true),   // medians only at close when TRUE

    // History retention & max rows (heavy prunes run separately)
    historyRetentionDays: retentionDays,
    historyMaxRows: num(c["HistoryMaxRows"], 200000),

    // Writer batching: cells budget → rows per batch computed inside HM_update
    historyChunkCells: num(c["History.ChunkSize"], 5000),

    // Misc the engine may consult
    daysForCandlestick: num(c["DaysForCandlestick"], 30),
    rebuildAlways: bool(c["RebuildAlways"], false),

    // Extras you already have (not directly used here but left for symmetry)
    pricesMaxRows: num(c["PricesMaxRows"], 100000),
    priceRetentionDays: num(c["PriceRetentionDays"], 1),
    bucketMinutes: num(c["BucketMinutes"], 20),
    chunkSize: num(c["ChunkSize"], 75), // (Fetcher’s chunk; not used by History)
    marketDefaults: {
      type_id: num(c["type_id"], 34),
      market_id: num(c["market_id"], 30002187),
      market_type: str(c["market_type"], "system"),
      maxLogIDs: num(c["MaxLogIDs"], 750),
    },
  };
}
function _tzNowNY_() { return new Date(Utilities.formatDate(new Date(), 'America/New_York', "yyyy/MM/dd HH:mm:ss")); }
function _minutesSinceMidnightNY_(d) { return +Utilities.formatDate(d, 'America/New_York', 'H')*60 + +Utilities.formatDate(d, 'America/New_York', 'mm'); }
function _parseHHMMtoMinutes_(s) { const m=/^(\d{1,2}):(\d{2})$/.exec(String(s||'')); if(!m) return null; return Math.min(23,Math.max(0,+m[1]))*60 + Math.min(59,Math.max(0,+m[2])); }

// ---- Time helpers (use ProjectTime/PT) ----
// Normalize any date-like value to midnight in the *project* time zone.
function _toProjectDay_(v) {
  const d = PT.parseDateSafe(v);                 // robust: Date | number | ISO-ish
  if (isNaN(d)) return new Date('Invalid');
  return PT.projectDate(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0);
}

// Minutes since local midnight in project TZ (PT.now() already matches project tz)
function _minsSinceMidnight_(d) {
  return d.getHours() * 60 + d.getMinutes();
}




/**
 * Slice flags computed in project TZ:
 * - open/close windows use your config strings and PT.coerceHM (handles "11", "11:00", "11am", 1130, [11,30], Date, etc.)
 * - no hard-coded "America/New_York"
 */
function _sliceFlags_(cfg) {
  const now = PT.now(); // project-local "now"
  const openHM  = PT.coerceHM(cfg.openTime);   // {h,m}
  const closeHM = PT.coerceHM(cfg.closeTime);  // {h,m}

  const minsNow  = _minsSinceMidnight_(now);
  const openMin  = (openHM.h|0)  * 60 + (openHM.m|0);
  const closeMin = (closeHM.h|0) * 60 + (closeHM.m|0);

  const inOpenWindow  = minsNow >= openMin  && minsNow <  openMin  + (cfg.openWindowMins|0);
  const inCloseWindow = minsNow >= closeMin && minsNow <= closeMin + (cfg.closeWindowMins|0);

  return { now, minsNow, openMin, closeMin, inOpenWindow, inCloseWindow };
}



const HM_HEADERS = [
  'type_id','market_id','market_type','date',
  'buy_open','buy_close','sell_open','sell_close',
  'buy_high','buy_low','sell_high','sell_low',
  'median_buy','median_sell'
];

/* ====================== Sheet/Cap Utilities ====================== */

function _wbAllocatedCells_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().reduce((sum, sh) => sum + sh.getMaxRows() * sh.getMaxColumns(), 0);
}

function _cellsNeededForAppend_(sh, addRows, needCols) {
  const cols = Math.max(needCols, sh.getMaxColumns());
  return addRows * cols;
}

function _tightenTrailingRows_(sh, rowBuffer) {
  rowBuffer = rowBuffer || 2000;
  const used = Math.max(1, sh.getLastRow());
  const alloc = sh.getMaxRows();
  const keep = Math.max(used + rowBuffer, Math.min(alloc, used + rowBuffer));
  const extra = alloc - keep;
  if (extra > 0) sh.deleteRows(keep + 1, extra);
}

function _deleteRowsAscBlocks_(sh, rowsAsc) {
  if (!rowsAsc || !rowsAsc.length) return;
  // Delete from the end; Sheets reindexes rows as you delete
  for (let i = rowsAsc.length - 1; i >= 0; i--) {
    sh.deleteRow(rowsAsc[i]);
  }
}

/**
 * Ensure there’s enough workbook capacity to append `addRows` at width `needCols`.
 * Strategy: optional retention prune (older than N days) → tighten → capacity check → throw if insufficient.
 */
function ensureAppendCapacity_History_(sh, addRows, needCols, retentionDays, dateColIdx1) {
  if (addRows <= 0) return;

  // 1) Optional simple retention prune on this sheet (older than N days)
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

  // 2) Tighten trailing allocation on this sheet
  _tightenTrailingRows_(sh, 2000);

  // 3) Capacity check against global cap
  const CAP    = (typeof WORKBOOK_CAP === 'number') ? WORKBOOK_CAP : 10000000;
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

/** Use global getOrCreateSheet if present; otherwise safe local fallback. */
function _ensureSheet_(ss, name, headers) {
  if (typeof getOrCreateSheet === 'function') {
    return getOrCreateSheet(ss, name, headers);
  }
  // Fallback (limited): create/verify header row
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

/* ======================== Aggregation Helpers ======================== */



function _median_(arr) {
  const xs = (arr || []).filter(n => Number.isFinite(n)).sort((a, b) => a - b);
  if (!xs.length) return null;
  const mid = Math.floor(xs.length / 2);
  return xs.length % 2 ? xs[mid] : (xs[mid - 1] + xs[mid]) / 2;
}

/**
 * Build per-day OHLC + medians from raw price samples.
 * rows = [date, market_id, market_type, type_id, min_sell, max_buy, median_sell, median_buy]
 */
function _aggDailyFromPrices_(rows) {
  const bag = new Map();

  function pushSample(key, stamp, sellMin, buyMax, medSell, medBuy, typeId, marketId, mtype, day) {
    let rec = bag.get(key);
    if (!rec) {
      rec = {
        typeId, marketId, mtype, day,
        firstTs: stamp, lastTs: stamp,
        sellOpen: sellMin, sellClose: sellMin,
        buyOpen:  buyMax,  buyClose:  buyMax,
        sellHigh: sellMin != null ? sellMin : null,
        sellLow:  sellMin != null ? sellMin : null,
        buyHigh:  buyMax  != null ? buyMax  : null,
        buyLow:   buyMax  != null ? buyMax  : null,
        medSell: [], medBuy: []
      };
      bag.set(key, rec);
    } else {
      if (stamp < rec.firstTs) {
        rec.firstTs = stamp;
        if (sellMin != null) rec.sellOpen = sellMin;
        if (buyMax  != null) rec.buyOpen  = buyMax;
      }
      if (stamp > rec.lastTs) {
        rec.lastTs = stamp;
        if (sellMin != null) rec.sellClose = sellMin;
        if (buyMax  != null) rec.buyClose  = buyMax;
      }
      if (sellMin != null) {
        if (rec.sellHigh == null || sellMin > rec.sellHigh) rec.sellHigh = sellMin;
        if (rec.sellLow  == null || sellMin < rec.sellLow)  rec.sellLow  = sellMin;
      }
      if (buyMax != null) {
        if (rec.buyHigh == null || buyMax > rec.buyHigh) rec.buyHigh = buyMax;
        if (rec.buyLow  == null || buyMax < rec.buyLow)  rec.buyLow  = buyMax;
      }
    }
    if (medSell != null) rec.medSell.push(medSell);
    if (medBuy  != null) rec.medBuy.push(medBuy);
  }

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const ts = r[0]; if (!(ts instanceof Date)) continue;
    const marketId = Number(r[1]);
    const mtype = String(r[2] || '');
    const typeId = Number(r[3]);
    if (!Number.isFinite(marketId) || !Number.isFinite(typeId) || !mtype) continue;

    const minSell = (r[4] == null || r[4] <= 0) ? null : Number(r[4]);
    const maxBuy  = (r[5] == null || r[5] <= 0) ? null : Number(r[5]);
    const medSell = (r[6] == null || r[6] <= 0) ? null : Number(r[6]);
    const medBuy  = (r[7] == null || r[7] <= 0) ? null : Number(r[7]);

    const day = _toProjectDay_(ts);
    const key = `${typeId}|${marketId}|${mtype}|${day.getTime()}`;
    pushSample(key, ts.getTime(), minSell, maxBuy, medSell, medBuy, typeId, marketId, mtype, day);
  }

  // Final rows in History schema
  const out = new Map();
  for (const [key, v] of bag.entries()) {
    out.set(key, [
      v.typeId, v.marketId, v.mtype, v.day,
      v.buyOpen ?? '',  v.buyClose ?? '',
      v.sellOpen ?? '', v.sellClose ?? '',
      v.buyHigh ?? '',  v.buyLow ?? '',
      v.sellHigh ?? '', v.sellLow ?? '',
      _median_(v.medBuy)  ?? '',
      _median_(v.medSell) ?? ''
    ]);
  }
  return out;
}

/* ======================== Formatting ======================== */

function formatHistoryBlock_(sh, startRow, nRows, idx1) {
  if (!nRows) return;
  const nfInt  = '#,##0';
  const nfDate = 'yyyy-mm-dd';
  sh.getRange(startRow, idx1.date, nRows, 1).setNumberFormat(nfDate);
  const numericCols = [
    idx1.buy_open, idx1.buy_close, idx1.sell_open, idx1.sell_close,
    idx1.buy_high, idx1.buy_low, idx1.sell_high, idx1.sell_low,
    idx1.median_buy, idx1.median_sell
  ];
  numericCols.forEach(c => sh.getRange(startRow, c, nRows, 1).setNumberFormat(nfInt));
}

/* ======================== Upsert Helpers ======================== */

function _buildIndex_(vals) {
  // vals includes header row at [0]
  if (!vals || vals.length === 0) return { idx: new Map(), cols: 0 };
  const H = vals[0]; const cols = H.length;
  const iType = H.indexOf('type_id');
  const iMid  = H.indexOf('market_id');
  const iMtp  = H.indexOf('market_type');
  const iDate = H.indexOf('date');

  // Coerce to local (NY) midnight — accept Date, serial, or ISO string
  const toNYMidnight = (v) => {
    if (v instanceof Date) return _toProjectDay_(v);
    const s = String(v || '').trim();
    if (!s) return null;
    const n = Number(s);
    if (Number.isFinite(n)) {
      // Excel/Sheets serial → JS Date (UTC midnight), then to NY midnight
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
    const t  = Number(row[iType]);
    const m  = Number(row[iMid]);
    const mt = String(row[iMtp] || '').trim();
    const d  = toNYMidnight(row[iDate]);
    if (!Number.isFinite(t) || !Number.isFinite(m) || !mt || !d) continue;
    idx.set(`${t}|${m}|${mt}|${+d}`, r + 1); // 1-based sheet row
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

/* ============================ MAIN ============================ */
/**
 * Build daily candles from the last N hours of Market Prices and upsert into Market History.
 * Options:
 *   - lookbackHours (number)
 *   - retentionDays (number)   // only used for pre-prune before capacity check
 *   - writeChunk (number)
 *   - test / dryRun (boolean)  // compute/log only; no writes
 *   - mode: "close"            // ensure we include late-day samples (boosts lookback)
 *   - auto: boolean            // telemetry only
 */
function HM_update(opts) {
  // ---- options & config ----
  opts = opts || {};
  const cfg   = _hmCfg();
  const slice = _sliceFlags_(cfg);

  const MODE    = String(opts.mode || '').toLowerCase();    // e.g., "close"
  const IS_TEST = !!(opts.test || opts.dryRun);             // dry-run (no writes)
  const IS_AUTO = !!opts.auto;                              // telemetry only

  let LOOKBACK_H = Number(opts.lookbackHours || cfg.lookbackHours || 26);
  const RETAIN_D = Number(opts.retentionDays || cfg.historyRetentionDays || 365);
  if (MODE === 'close') LOOKBACK_H = Math.max(LOOKBACK_H, 30);

  // ---- sheets ----
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const shPrices  = ss.getSheetByName(cfg.sheets.prices);
  const shHistory = _ensureSheet_(ss, cfg.sheets.history, HM_HEADERS);
  if (!shPrices) { (LoggerEx?.warn || console.warn)('[HM] Missing "Market Prices"'); return; }

  // ---- price header + bounds ----
  const P_H = shPrices.getRange(1, 1, 1, shPrices.getLastColumn()).getValues()[0];
  const P = { date:1, market_id:2, market_type:3, type_id:4, min_sell:5, max_buy:6, median_sell:7, median_buy:8 };

  const lastRow = shPrices.getLastRow();
  if (lastRow < 2) { (LoggerEx?.log || console.log)('[HM] No prices data'); return; }

  // ---- read recent "Market Prices" window (scan upward until cutoff) ----
  const cutoffMs = Date.now() - LOOKBACK_H * 60 * 60 * 1000;
  const data = shPrices.getRange(2, 1, lastRow - 1, P_H.length).getValues();
  const recent = [];
  for (let i = data.length - 1; i >= 0; i--) {
    const dt = data[i][P.date - 1];
    if (!(dt instanceof Date)) continue;
    if (dt.getTime() < cutoffMs) break;
    recent.push([
      dt,
      data[i][P.market_id - 1], data[i][P.market_type - 1], data[i][P.type_id - 1],
      data[i][P.min_sell - 1],  data[i][P.max_buy - 1],
      data[i][P.median_sell - 1], data[i][P.median_buy - 1]
    ]);
  }
  recent.reverse();
  if (!recent.length) { (LoggerEx?.log || console.log)('[HM] No recent prices within lookback=', LOOKBACK_H); return; }

  // ---- aggregate prices → daily OHLC/medians ----
  const agg = _aggDailyFromPrices_(recent); // Map<key, row[]>

  // ---- target day set (NY midnight) to restrict History indexing ----
  const daySet = new Set();
  let minDayNY = null;
  for (const [, row] of agg) {
    const d = row[3]; // date field in history schema
    if (d instanceof Date) {
      const ny = _toProjectDay_(d);
      daySet.add(+ny);
      if (!minDayNY || ny < minDayNY) minDayNY = ny;
    }
  }

  // ---- index only rows in History for the target day range (FAST PATH) ----
  // Columns we need to build keys: type_id(1), market_id(2), market_type(3), date(4)
  const IDX1 = {
    date: 4,
    buy_open: 5, buy_close: 6, sell_open: 7, sell_close: 8,
    buy_high: 9, buy_low: 10, sell_high: 11, sell_low: 12,
    median_buy: 13, median_sell: 14
  };
  const COLS = HM_HEADERS.length;

  const H_lastRow = shHistory.getLastRow();
  let existingIdx = new Map();

  if (H_lastRow >= 2 && minDayNY) {
    // Read just the date column, find where the earliest target day starts (scan backward; append-only assumption)
    const dateCol = shHistory.getRange(2, IDX1.date, H_lastRow - 1, 1).getValues().flat();
    let startIdx = 0; // 0-based into dateCol
    for (let i = dateCol.length - 1; i >= 0; i--) {
      const v = dateCol[i];
      if (!(v instanceof Date)) continue;
      const ny = _toProjectDay_(v);
      if (ny < minDayNY) { startIdx = i + 1; break; }
    }
    // small buffer just in case
    startIdx = Math.max(0, startIdx - 10);
    const rowsToRead = dateCol.length - startIdx;
    if (rowsToRead > 0) {
      const block = shHistory.getRange(2 + startIdx, 1, rowsToRead, 4).getValues(); // [type_id, market_id, market_type, date]
      const baseRow1 = 2 + startIdx; // 1-based sheet row of block[0]
      for (let i = 0; i < block.length; i++) {
        const r = block[i];
        const t = Number(r[0]), m = Number(r[1]), mt = String(r[2]||'').trim(), d = r[3];
        if (!Number.isFinite(t) || !Number.isFinite(m) || !mt || !(d instanceof Date)) continue;
        const dny = _toProjectDay_(d);
        if (!daySet.has(+dny)) continue; // keep only rows in our target days
        const key = `${Math.floor(t)}|${Math.floor(m)}|${mt}|${+dny}`;
        existingIdx.set(key, baseRow1 + i);
      }
    }
  }

  // ---- partition into updates vs appends ----
  const updates = [];
  const appends = [];
  agg.forEach((row, key) => {
    const rn = existingIdx.get(key);
    if (rn) updates.push({ rn, row }); else appends.push(row);
  });

  // ---- LightSlices: blank medians (today only) when not in close window ----
  if (cfg.lightSlices && !slice.inCloseWindow) {
    const todayNY = _toProjectDay_(slice.now);
    const isTodayNY = (d) => (d instanceof Date) && (+_toProjectDay_(d) === +todayNY);

    // Appends
    for (let i = 0; i < appends.length; i++) {
      const d = appends[i][IDX1.date - 1];
      if (isTodayNY(d)) {
        appends[i][IDX1.median_buy  - 1] = '';
        appends[i][IDX1.median_sell - 1] = '';
      }
    }
    // Updates
    for (let i = 0; i < updates.length; i++) {
      const row = updates[i].row;
      const d = row[IDX1.date - 1];
      if (isTodayNY(d)) {
        row[IDX1.median_buy  - 1] = '';
        row[IDX1.median_sell - 1] = '';
      }
    }
  }

  // ---- serialize writes ----
  const dlock = LockService.getDocumentLock();
  if (!dlock.tryLock(250)) {
    try { dlock.waitLock(5000); }
    catch (e) { (LoggerEx?.warn || console.warn)('[HM] Document busy; skipping'); return; }
  }

  try {
    if (IS_TEST) {
      const CAP    = (typeof WORKBOOK_CAP === 'number') ? WORKBOOK_CAP : 10000000;
      const SAFETY = (typeof WORKBOOK_SAFETY === 'number') ? WORKBOOK_SAFETY : 200000;
      const want   = _cellsNeededForAppend_(shHistory, appends.length, COLS);
      const have   = Math.max(0, CAP - SAFETY - _wbAllocatedCells_());
      (LoggerEx?.log || console.log)('[HM][DRY-RUN]', {
        mode: MODE || null, auto: IS_AUTO,
        updates: updates.length, appends: appends.length,
        wantCells: want, freeCells: have
      });
      return { dryRun:true, mode:MODE, auto:IS_AUTO, updates:updates.length, appends:appends.length, wantCells:want, freeCells:have };
    }

    // Guard capacity on this sheet (may pre-prune by retention)
    ensureAppendCapacity_History_(shHistory, appends.length, COLS, RETAIN_D, IDX1.date);

    // Updates as contiguous runs
    _writeContiguousHM_(shHistory, updates, COLS);

    // Appends in chunks (config is cells-budget → convert to rows) + format
    if (appends.length) {
      const start = shHistory.getLastRow() + 1;
      const cellsBudget = Math.max(1000, Number(cfg.historyChunkCells || cfg.writeChunk || 5000)); // cells per batch
      const rowsPerBatch = Math.max(50, Math.floor(cellsBudget / COLS));
      const CH = rowsPerBatch;

      for (let i = 0; i < appends.length; i += CH) {
        const seg = appends.slice(i, i + CH);
        shHistory.getRange(start + i, 1, seg.length, COLS).setValues(seg);
      }
      formatHistoryBlock_(shHistory, start, appends.length, IDX1);
    }

    SpreadsheetApp.flush();
    (LoggerEx?.log || console.log)('[HM] upserts=', {
      updates: updates.length, appends: appends.length,
      mode: MODE || null, auto: IS_AUTO || false
    });

  } finally {
    try { dlock.releaseLock(); } catch (_) {}
  }
}



/* ===================== Back-compat entry points ===================== */

function updateMarketHistory(opts) { return HM_update(opts); }

// Your originals — keep triggers intact
function updateHistory()       { return HM_update({ auto: true,  test: false }); }
function updateHistoryTest()   { return HM_update({ auto: false, test: true, mode: 'close' }); }

/* =================== Optional maintenance helpers =================== */
/** Stop any triggers that mistakenly point at HM_update (or your wrappers). */
function HM_stopTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['HM_update','updateHistory','updateHistoryTest'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/** Quick one-shot manual run (adjust lookback if you need to backfill today’s close). */
function HM_oneShot() { HM_update({ lookbackHours: 30 }); }

/** Diagnose duplicate keys in Market History (top N). */
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

/** Dedupe Market History keeping the **last** row per key (type|market|mtype|day). */
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
  const sh  = SpreadsheetApp.getActive().getSheetByName(cfg.sheets.history);
  if (!sh) return Logger.log('History sheet not found');

  const vals = sh.getDataRange().getValues();
  const h = vals[0];
  const I = { type: h.indexOf('type_id'), mid: h.indexOf('market_id'),
              mtp: h.indexOf('market_type'), date: h.indexOf('date') };

  const today = _toProjectDay_(PT.now()); // uses your ProjectTime
  const seen = new Map(), dupes = [];

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const d0  = row[I.date]; if (!(d0 instanceof Date)) continue;
    const d   = _toProjectDay_(d0); if (+d !== +today) continue;

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
