/** MarketFetcher.gs — Prices runner with Light/Heavy Prune + Lock
 * Assumes sheet "Market Prices" exists with header in row 1:
 * ["date","market_id","market_type","type_id","min_sell","max_buy","median_sell","median_buy"]
 * Depends on:
 * - getConfig(), getMarketSettings(), getTypeIDsFromItemList(), fuzAPI.requestItems()
 * - LoggerEx
 */

/* ---------------------- Config helpers ---------------------- */

function mtConfig() {
  const c = (typeof getConfig === 'function') ? (getConfig() || {}) : {};
  const num = (v, d) => (v == null || isNaN(Number(v))) ? d : Number(v);
  const str = (v, d) => (v == null || v === '') ? d : String(v);
  return {
    sheets: {
      prices:  str(c["MarketPricesSheet"], "Market Prices"),
      history: str(c["HistorySheetName"],  "Market History"),
    },
    retentionDays: {
      prices:  num(c["PriceRetentionDays"],   1),
      history: num(c["HistoryRetentionDays"], 365),
    },
    maxRows: {
      prices:  num(c["PricesMaxRows"],  100000),
      history: num(c["HistoryMaxRows"], 200000),
    },
    chunkSize:     num(c["ChunkSize"],     75),
    maxLogIDs:     num(c["MaxLogIDs"],    750),
    bucketMinutes: num(c["BucketMinutes"], 20),
    // optional:
    daysForCandlestick: num(c["DaysForCandlestick"], 30),
    openTime:           str(c["OpenTime"],  "11:00"),
    closeTime:          str(c["CloseTime"], "18:00"),
    rebuildAlways:      String(c["RebuildAlways"] || "FALSE").toUpperCase() === "TRUE",
  };
}
function getBucketMinutes() { return mtConfig().bucketMinutes; }

/* ---------------------- Utilities ---------------------- */

function sanitizeIDs(ids) {
  return [...new Set(ids.map(v => Number(v)).filter(v => !isNaN(v) && v > 0))];
}
function getMarketSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Market Settings");
  if (!sheet) throw new Error("Market Settings sheet not found.");
  const types = ["station", "system", "region"];
  const combos = [];
  types.forEach((type, i) => {
    const col = 4 + i; // D,E,F
    const raw = sheet.getRange(3, col, Math.max(0, sheet.getLastRow() - 2), 1).getValues().flat();
    sanitizeIDs(raw).forEach(id => combos.push({ market_id: id, market_type: type }));
  });
  return combos;
}
function getTypeIDsFromItemList(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Item List Back End");
  if (!sheet) throw new Error("Item List Back End sheet not found.");
  let ids = sheet.getRange("A2:A" + sheet.getLastRow()).getValues().flat();
  ids = sanitizeIDs(ids);
  if (limit && ids.length > limit) ids = ids.slice(0, limit);
  return ids;
}


/**
 * Coerces v into a positive finite number, else null.
 * - Accepts numbers or strings (e.g. "12,345.67").
 * - Treats 0/negatives/NaN/undefined as null.
 */
const toPosNumberOrNull = (v) => {
  if (v == null) return null;
  const n = typeof v === 'number' ? v : Number(String(v).replace(/[,\s]/g, ''));
  return Number.isFinite(n) && n > 0 ? n : null;
};

/** Map Fuz result into math-friendly numbers (null for non-numeric/≤0). */
function getMarketPrices(typeIds, marketId, marketType) {
  let dataMap = {};
  try {
    // [PATCH START] Replace old postFetch with new fuzAPI.requestItems
    // fuzAPI.requestItems returns an Array of { type_id, buy, sell, ... } objects
    const resultsArray = fuzAPI.requestItems(marketId, marketType, typeIds) || [];
    
    // Convert array to a map { type_id: {buy:..., sell:...} } for easier access by ID
    resultsArray.forEach(item => {
        dataMap[item.type_id] = item;
    });
    // [PATCH END]
  } catch (e) {
    // If the call blows up, return all-null rows for requested ids.
    return Object.fromEntries(
      [...new Set(typeIds)].map(id => [id, {
        minSell: null, maxBuy: null, medianSell: null, medianBuy: null
      }])
    );
  }

  const out = {};
  for (const id of typeIds) {
    // Accessing the new FuzDataObject structure: dataMap?.[id] = { buy: {max, median, ...}, sell: {...} }
    const e = dataMap?.[id] ?? {};
    const sell = e?.sell ?? {};
    const buy  = e?.buy  ?? {};
    out[id] = {
      // The keys min, max, median, etc., are now accessed from the new FuzDataObject structure.
      minSell:    toPosNumberOrNull(sell.min),
      maxBuy:     toPosNumberOrNull(buy.max),
      medianSell: toPosNumberOrNull(sell.median),
      medianBuy:  toPosNumberOrNull(buy.median),
    };
  }
  return out;
}


/* ---------------------- Entry: record prices (baseline) ---------------------- */

function getCurrentMarketPrices() {
  const cfg = mtConfig();
  const SHEET_NAME = cfg.sheets.prices;
  const RETENTION  = cfg.retentionDays.prices;
  const MAX_ROWS   = cfg.maxRows.prices;
  const MAX_LOG    = cfg.maxLogIDs;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) { LoggerEx && LoggerEx.warn("Prices: skip — lock busy"); return; }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const headers = ["date","market_id","market_type","type_id","min_sell","max_buy","median_sell","median_buy"];
    const sheet = getOrCreateSheet(ss, SHEET_NAME, headers);

    // Light pre-prune (fast)
    lightPrePrune_(sheet, MAX_ROWS);

    // Collect prices
    const typeIDs = getTypeIDsFromItemList(MAX_LOG);
    const marketCombos = getMarketSettings();
    if (!typeIDs.length || !marketCombos.length) {
      LoggerEx && LoggerEx.warn("Prices: no typeIDs or no markets; aborting run.");
      return;
    }
    const now = new Date(); // timestamp stored in local tz (NY), not UTC)
    const rows = [];
    marketCombos.forEach(({ market_id, market_type }) => {
      const prices = getMarketPrices(typeIDs, market_id, market_type);
      typeIDs.forEach(type_id => {
        const e = prices[type_id] || {};
        rows.push([now, market_id, market_type, type_id, e.minSell ?? null, e.maxBuy ?? null, e.medianSell ?? null, e.medianBuy ?? null]);
      });
    });

    if (rows.length) {
      const start = sheet.getLastRow() + 1;
      sheet.getRange(start, 1, rows.length, headers.length).setValues(rows);
      LoggerEx && LoggerEx.log("Prices: wrote", rows.length, "rows @", start);
    } else {
      LoggerEx && LoggerEx.warn("Prices: no rows produced this run.");
    }

    // Tighten & simple retention (fast)
    postTighten_(sheet);
    if (RETENTION) pruneOldRows(sheet, RETENTION, /*dateCol=*/1);
  } catch (e) {
    LoggerEx && LoggerEx.error("getCurrentMarketPrices failed:", e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------- Light Mode hygiene ---------------------- */

function lightPrePrune_(sh, maxRows) {
  const f = sh.getFilter && sh.getFilter(); if (f) f.remove();
  _trimTrailing_(sh);
  const last = sh.getLastRow();
  if (maxRows && last > maxRows) {
    const over = last - maxRows;
    _deleteInBlocks_(sh, 2, over);
    LoggerEx && LoggerEx.log('Prices: light prune removed oldest rows:', over);
  }
}
function postTighten_(sh) { _trimTrailing_(sh); }
function _trimTrailing_(sh) {
  const used = sh.getLastRow();
  const alloc = sh.getMaxRows();
  const extra = alloc - used;
  if (extra > 0) _deleteInBlocks_(sh, used + 1, extra);
}
// Deletes rows in blocks but never removes the last non-frozen row.
// If the caller tries to delete the entire body, we leave one row and clear it.
function _deleteInBlocks_(sh, startRow, count) {
  const BLOCK = 20000;

  if (count <= 0) return;

  const frozen = sh.getFrozenRows();
  const bodyStart = frozen + 1;                  // first non-frozen row
  const maxRows   = sh.getMaxRows();
  const bodyRows  = Math.max(0, maxRows - frozen);

  // normalize startRow to body
  if (startRow < bodyStart) startRow = bodyStart;

  // how many rows exist before our start within the body?
  const keptTop = Math.max(0, startRow - bodyStart);

  // We must leave at least 1 non-frozen row:
  // deletable = bodyRows - 1 - keptTop
  let safeCount = Math.min(count, Math.max(0, bodyRows - 1 - keptTop));
  if (safeCount <= 0) {
    // If the intent was "wipe everything", just clear the single kept row
    if (startRow === bodyStart && count >= bodyRows) {
      sh.getRange(bodyStart, 1, 1, sh.getMaxColumns()).clearContent();
    }
    return;
  }

  // Delete in chunks (row index stays the same because rows collapse upward)
  let row = startRow;
  while (safeCount > 0) {
    const n = Math.min(BLOCK, safeCount);
    sh.deleteRows(row, n);
    safeCount -= n;
  }

  // If the caller intended to delete the whole body, clear the one kept row
  if (startRow === bodyStart && count >= bodyRows) {
    sh.getRange(bodyStart, 1, 1, sh.getMaxColumns()).clearContent();
  }
}


/** Batch/contiguous prune for "older than N days" assuming chronological appends. */
function pruneOldRows(sheet, retentionDays, dateCol /* 1-based */) {
  if (!retentionDays) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - retentionDays);

  // Read the date column once
  const dates = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues().flat();

  // Find the last index we should delete (contiguous block from the top)
  let boundary = -1; // last index in 'dates' to delete (0-based over data rows)
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i];
    if (!(d instanceof Date)) continue;            // skip blanks
    if (d < cutoff) boundary = i; else break;      // as soon as we hit fresh data, stop (chronological)
  }

  if (boundary >= 0) {
    const rowsToDelete = boundary + 1;            // convert 0-based to count
    _deleteInBlocks_(sheet, /*startRow=*/2, /*count=*/rowsToDelete);
    LoggerEx && LoggerEx.log('Prices: pruned old rows (<= cutoff):', rowsToDelete);
  }
}

/* ---------------------- Heavy Prune (daily) ---------------------- */

function dailyHeavyPrune_Prices() {
  const CFG = mtConfig();
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheets.prices);
  if (!sh) { LoggerEx && LoggerEx.warn('HeavyPrune: sheet not found'); return; }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(2000)) { LoggerEx && LoggerEx.warn('HeavyPrune skip: lock busy'); return; }
  try {
    const f = sh.getFilter && sh.getFilter(); if (f) f.remove();
    heavyPruneSheet_(sh, CFG.retentionDays.prices, CFG.maxRows.prices, CFG.bucketMinutes);
  } finally {
    lock.releaseLock();
  }
}

function dailyHeavyPrune_History() {
  const CFG = mtConfig();
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.sheets.history);
  if (!sh) return;
  heavyPruneSheet_(sh, CFG.retentionDays.history, CFG.maxRows.history, CFG.bucketMinutes);
}

/** Heavy prune: retention → dedupe (20m buckets) → cap → rewrite once → tighten */
function heavyPruneSheet_(sh, retentionDays, maxRows, bucketMinutes) {
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  const header = sh.getRange(1,1,1,lastCol).getValues()[0];
  const lower  = header.map(h => String(h).trim().toLowerCase());
  const find = (name) => lower.findIndex(h => h === name);

  // Resolve key column indices (0-based for arrays)
  const DATE = find('date');
  const TYPE = find('type_id');
  const MID  = find('market_id');
  const MTP  = find('market_type');
  if (DATE < 0 || TYPE < 0 || MID < 0 || MTP < 0) {
    LoggerEx && LoggerEx.warn('HeavyPrune: missing required columns (date/type_id/market_id/market_type)');
    return;
  }

  // Read all data rows once
  const data = sh.getRange(2,1,lastRow-1,lastCol).getValues();

  // 1) Retention window
  const cutoff = new Date(Date.now() - retentionDays * 86400000);
  const windowed = [];
  for (let i=0;i<data.length;i++) {
    const r = data[i];
    const d = r[DATE];
    if (d instanceof Date && d >= cutoff) windowed.push(r);
  }

  // 2) Dedupe by 20-min buckets — key = bucket|type_id|market_id|market_type
  const msPerBucket = (bucketMinutes || 20) * 60 * 1000;
  const keep = new Map();
  for (let i=0;i<windowed.length;i++) {
    const r = windowed[i];
    const d = r[DATE];
    if (!(d instanceof Date)) continue; // keep dedupe pure on dated rows
    const bucket = Math.floor(d.getTime() / msPerBucket);
    const key = bucket + '|' + r[TYPE] + '|' + r[MID] + '|' + r[MTP];

    const prev = keep.get(key);
    if (!prev || (r[DATE] > prev[DATE])) keep.set(key, r); // newest wins
  }
  let deduped = Array.from(keep.values());

  // 3) Enforce cap (keep newest by date)
  if (deduped.length > maxRows) {
    deduped.sort((a,b) => a[DATE] - b[DATE]);            // oldest first
    deduped = deduped.slice(deduped.length - maxRows);   // keep newest maxRows
  }

  // 4) Rewrite once (header + data) then tighten
  sh.clearContents();
  sh.getRange(1,1,1,lastCol).setValues([header]);
  if (deduped.length) sh.getRange(2,1,deduped.length,lastCol).setValues(deduped);

  // Tighten trailing blanks
  const used = sh.getLastRow(), alloc = sh.getMaxRows();
  if (alloc > used) sh.deleteRows(used + 1, alloc - used);

  LoggerEx && LoggerEx.log('HeavyPrune kept:', deduped.length, 'rows | window(days):', retentionDays, '| bucket(min):', bucketMinutes);
}