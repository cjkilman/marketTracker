/** HistoryManager.gs — Daily OHLC/median builder with cap guard (no auto-prune)
 *
 * Depends on:
 *   - WORKBOOK_CAP, WORKBOOK_SAFETY (from utility.gs)
 *   - LoggerEx (optional; falls back to console)
 *   - (optional) mtConfig() / getConfig() for settings
 *
 * Sheets:
 *   - Source:  "Market Prices"
 *     Columns: [date, market_id, market_type, type_id, min_sell, max_buy, median_sell, median_buy]
 *   - Target:  "Market History"
 *     Columns: [type_id, market_id, market_type, date,
 *               buy_open, buy_close, sell_open, sell_close,
 *               buy_high, buy_low, sell_high, sell_low,
 *               median_buy, median_sell]
 */

/* ----------------------- Config / helpers ----------------------- */

function _hmCfg() {
  // Allow overrides via getConfig()/mtConfig(), else defaults.
  const c = (typeof getConfig === 'function') ? (getConfig() || {}) : {};
  const pickNum = (k, d) => {
    const v = c[k]; const n = Number(v);
    return Number.isFinite(n) ? n : d;
  };
  const pickStr = (k, d) => {
    const v = c[k]; return (v == null || v === '') ? d : String(v);
  };
  return {
    sheets: {
      prices: pickStr('MarketPricesSheet', 'Market Prices'),
      history: pickStr('HistorySheetName', 'Market History'),
    },
    // Keep 365d of history unless overridden
    historyRetentionDays: pickNum('HistoryRetentionDays', 365),
    // How far back to look in "Prices" when computing today's candle
    lookbackHours: pickNum('HistoryLookbackHours', 26),
    // Append chunk size
    writeChunk: pickNum('HistoryWriteChunk', 500),
  };
}

const HM_HEADERS = [
  'type_id','market_id','market_type','date',
  'buy_open','buy_close','sell_open','sell_close',
  'buy_high','buy_low','sell_high','sell_low',
  'median_buy','median_sell'
];

/* ----------------------- Cap + sheet utilities ----------------------- */

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
  let start = rowsAsc[0], prev = start, count = 1;
  for (let i = 1; i <= rowsAsc.length; i++) {
    const r = rowsAsc[i];
    if (r === prev + 1) { prev = r; count++; continue; }
    sh.deleteRows(start, count);
    for (let j = i; j < rowsAsc.length; j++) rowsAsc[j] -= count; // reindex
    if (i >= rowsAsc.length) break;
    start = prev = rowsAsc[i]; count = 1;
  }
}

/**
 * Guard history appends against the 10M workbook cap.
 * Strategy: optional retention prune → tighten → check cap → error if insufficient.
 */
function ensureAppendCapacity_History_(sh, addRows, needCols, retentionDays, dateColIdx1) {
  if (addRows <= 0) return;

  // 1) Pre-retention prune (simple "older than N days")
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

  // 2) Tighten this sheet's trailing allocation to avoid waste
  _tightenTrailingRows_(sh, 2000);

  // 3) Capacity check
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

function getOrCreateSheet(ss, name, headers) {
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

/* ----------------------- Aggregation helpers ----------------------- */

function _toNYDay_(dt) {
  // Normalize a JS Date to midnight in America/New_York (as a Date object)
  const tz = 'America/New_York';
  const y = +Utilities.formatDate(dt, tz, 'yyyy');
  const m = +Utilities.formatDate(dt, tz, 'MM') - 1;
  const d = +Utilities.formatDate(dt, tz, 'dd');
  return new Date(y, m, d, 0, 0, 0, 0);
}

function _median_(arr) {
  const xs = arr.filter(n => Number.isFinite(n)).sort((a, b) => a - b);
  if (!xs.length) return null;
  const mid = Math.floor(xs.length / 2);
  return xs.length % 2 ? xs[mid] : (xs[mid - 1] + xs[mid]) / 2;
}

function _aggDailyFromPrices_(rows) {
  // rows: array of [date, market_id, market_type, type_id, min_sell, max_buy, median_sell, median_buy]
  // Returns a Map keyed by "type|market|mtype|dateSerial" -> aggregated record array for History.
  const bag = new Map();

  // Helper to push sample
  function touch(key, stamp, sellMin, buyMax, medSell, medBuy, typeId, marketId, mtype, day) {
    let rec = bag.get(key);
    if (!rec) {
      rec = {
        typeId, marketId, mtype, day,
        // opens/closes use earliest/latest timestamp observed
        firstTs: stamp, lastTs: stamp,
        sellOpen: sellMin, sellClose: sellMin,
        buyOpen: buyMax,  buyClose:  buyMax,
        sellHigh: (sellMin != null ? sellMin : null),
        sellLow:  (sellMin != null ? sellMin : null),
        buyHigh:  (buyMax  != null ? buyMax  : null),
        buyLow:   (buyMax  != null ? buyMax  : null),
        medSell: [], medBuy: []
      };
      bag.set(key, rec);
    } else {
      // open/close
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
      // highs/lows
      if (sellMin != null) {
        if (rec.sellHigh == null || sellMin > rec.sellHigh) rec.sellHigh = sellMin;
        if (rec.sellLow  == null || sellMin < rec.sellLow)  rec.sellLow  = sellMin;
      }
      if (buyMax != null) {
        if (rec.buyHigh == null || buyMax > rec.buyHigh) rec.buyHigh = buyMax;
        if (rec.buyLow  == null || buyMax < rec.buyLow)  rec.buyLow  = buyMax;
      }
    }
    // medians
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

    const day = _toNYDay_(ts);
    const key = `${typeId}|${marketId}|${mtype}|${day.getTime()}`;
    touch(key, ts.getTime(), minSell, maxBuy, medSell, medBuy, typeId, marketId, mtype, day);
  }

  // build final row arrays for write
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

/* ----------------------- Formatting ----------------------- */

function formatHistoryBlock_(sh, startRow, nRows, idx1) {
  if (!nRows) return;
  const nfInt  = '#,##0';
  const nfDate = 'yyyy-mm-dd';
  // date (col 4 by default)
  sh.getRange(startRow, idx1.date, nRows, 1).setNumberFormat(nfDate);
  // numeric fields
  const numericCols = [
    idx1.buy_open, idx1.buy_close, idx1.sell_open, idx1.sell_close,
    idx1.buy_high, idx1.buy_low, idx1.sell_high, idx1.sell_low,
    idx1.median_buy, idx1.median_sell
  ];
  numericCols.forEach(c => sh.getRange(startRow, c, nRows, 1).setNumberFormat(nfInt));
}

/* ----------------------- Upsert writer ----------------------- */

function _buildIndex_(vals) {
  // vals includes header row at [0]
  if (!vals || vals.length === 0) return { idx: new Map(), cols: 0 };
  const H = vals[0]; const cols = H.length;
  const iType = H.indexOf('type_id');
  const iMid  = H.indexOf('market_id');
  const iMtp  = H.indexOf('market_type');
  const iDate = H.indexOf('date');
  const idx = new Map();
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const t = Number(row[iType]), m = Number(row[iMid]);
    const mt = String(row[iMtp] || '');
    const d  = row[iDate] instanceof Date ? row[iDate] : null;
    if (!Number.isFinite(t) || !Number.isFinite(m) || !mt || !d) continue;
    const key = `${t}|${m}|${mt}|${_toNYDay_(d).getTime()}`;
    idx.set(key, r + 1); // sheet row number (1-based)
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

/* ----------------------- MAIN: build + upsert ----------------------- */

/**
 * Build daily candles from the last N hours of Market Prices and upsert into Market History.
 * Safe to call from a time trigger; uses a DocumentLock to serialize writes.
 */
function HM_update(opts) {
  opts = opts || {};
  const cfg = _hmCfg();
  const LOOKBACK_H = Number(opts.lookbackHours || cfg.lookbackHours || 26);
  const RETAIN_D   = Number(opts.retentionDays || cfg.historyRetentionDays || 365);
  const CHUNK      = Number(opts.writeChunk || cfg.writeChunk || 500);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPrices  = ss.getSheetByName(cfg.sheets.prices);
  const shHistory = getOrCreateSheet(ss, cfg.sheets.history, HM_HEADERS);
  if (!shPrices) { (LoggerEx?.warn || console.warn)('[HM] Missing "Market Prices"'); return; }

  // Resolve column indexes in Prices (1-based for readability)
  // Prices headers are fixed per MarketFetcher
  const P_H = shPrices.getRange(1, 1, 1, shPrices.getLastColumn()).getValues()[0];
  const P = {
    date: 1,         // "date"
    market_id: 2,    // "market_id"
    market_type: 3,  // "market_type"
    type_id: 4,      // "type_id"
    min_sell: 5,     // "min_sell"
    max_buy: 6,      // "max_buy"
    median_sell: 7,  // "median_sell"
    median_buy: 8    // "median_buy"
  };

  // Pull recent window from Prices
  const lastRow = shPrices.getLastRow();
  if (lastRow < 2) { (LoggerEx?.log || console.log)('[HM] No prices data'); return; }

  const data = shPrices.getRange(2, 1, lastRow - 1, P_H.length).getValues();
  const cutoffMs = Date.now() - LOOKBACK_H * 60 * 60 * 1000;
  const recent = [];
  for (let i = data.length - 1; i >= 0; i--) {
    const dt = data[i][P.date - 1];
    if (!(dt instanceof Date)) continue;
    if (dt.getTime() < cutoffMs) break; // data is chronological; stop scanning
    recent.push([
      dt,
      data[i][P.market_id - 1], data[i][P.market_type - 1], data[i][P.type_id - 1],
      data[i][P.min_sell - 1],  data[i][P.max_buy - 1],
      data[i][P.median_sell - 1], data[i][P.median_buy - 1]
    ]);
  }
  recent.reverse(); // restore chronological order

  if (!recent.length) { (LoggerEx?.log || console.log)('[HM] No recent prices within lookback'); return; }

  // Aggregate into daily rows
  const agg = _aggDailyFromPrices_(recent); // Map<key, row[]>

  // Build existing index on History (for upserts)
  const H_lastRow = shHistory.getLastRow();
  const H_vals = shHistory.getRange(1, 1, Math.max(1, H_lastRow), HM_HEADERS.length).getValues();
  if (!H_vals[0][0]) {
    // Ensure header if sheet was empty
    shHistory.getRange(1, 1, 1, HM_HEADERS.length).setValues([HM_HEADERS]);
    shHistory.setFrozenRows(1);
    H_vals[0] = HM_HEADERS.slice();
  }
  const { idx: existingIdx, cols: totalCols } = _buildIndex_(H_vals);
  const COLS = HM_HEADERS.length;

  // Partition into updates vs appends
  const updates = [];
  const appends = [];
  agg.forEach((row, key) => {
    const rn = existingIdx.get(key);
    if (rn) updates.push({ rn, row }); else appends.push(row);
  });

  // 1-based index map for formatting
  const IDX1 = {
    date: 4,
    buy_open: 5, buy_close: 6, sell_open: 7, sell_close: 8,
    buy_high: 9, buy_low: 10, sell_high: 11, sell_low: 12,
    median_buy: 13, median_sell: 14
  };

  // WRITE (serialized with DocumentLock)
  const dlock = LockService.getDocumentLock();
  if (!dlock.tryLock(250)) {
    try { dlock.waitLock(5000); }
    catch (e) { (LoggerEx?.warn || console.warn)('[HM] Document busy; skipping'); return; }
  }
  try {
    // Guard append capacity before any writes
    ensureAppendCapacity_History_(shHistory, /*addRows=*/appends.length, /*needCols=*/COLS, RETAIN_D, /*dateColIdx1=*/IDX1.date);

    // updates in contiguous blocks
    _writeContiguousHM_(shHistory, updates, COLS);

    // appends in chunks + format
    if (appends.length) {
      const start = shHistory.getLastRow() + 1;
      const CH = Math.max(50, CHUNK);
      for (let i = 0; i < appends.length; i += CH) {
        const seg = appends.slice(i, i + CH);
        shHistory.getRange(start + i, 1, seg.length, COLS).setValues(seg);
      }
      formatHistoryBlock_(shHistory, start, appends.length, IDX1);
    }
    SpreadsheetApp.flush();
    (LoggerEx?.log || console.log)('[HM] upserts=', { updates: updates.length, appends: appends.length });

  } finally {
    try { dlock.releaseLock(); } catch (_) {}
  }
}

/* Convenience alias if you prefer a different entry name */
function updateMarketHistory(opts) { return HM_update(opts); }
