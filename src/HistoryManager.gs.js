/***********************
 * History Manager (clean, commented)
 * Scope: uses global getOrCreateSheet, PT, and Utility only.
 ***********************/

/*********************** Entrypoints ***********************/
function updateHistory() { return HM_update({ auto: true, test: false }); }
function updateHistoryTest() { return HM_update({ auto: false, test: true, mode: "close" }); }

/*********************** Core ***********************/
function HM_update(opts) {
  opts = opts || {};
  let mode = "open";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = (typeof getConfig === "function") ? (getConfig() || {}) : {};

  // 👇 choose sheet by mode
  const baseSheetName = cfg["History.SheetName"] || "Market History";
  const testSheetName = cfg["History.TestSheetName"] || "History (TEST)";
  const sheetName = (opts.test ? testSheetName : baseSheetName);

  const retentionDays = parseInt(cfg["History.RetentionDays"] || cfg["RetentionDays"] || 365, 10);
  const CHUNK = parseInt(cfg["History.ChunkSize"] || cfg["ChunkSize"] || 250, 10);

  const nowLocal = new Date();
  const appends = [];
  const updates = [];
  // Group rows by (type|market|mtype)
  const groups = new Map();

  // (optional: make it obvious in logs which sheet we’re targeting)
  LoggerEx.log(`[History] Target sheet: ${sheetName} (test=${!!opts.test})`);
  // ... rest of HM_update unchanged ...

  // 🔧 HOIST THESE LINES (so the loop can use dateKeyToday)
  const dateCell = (typeof PT !== 'undefined' && PT && typeof PT.todayAt === 'function')
    ? PT.todayAt(0, 0, 0)
    : new Date(nowLocal.getFullYear(), nowLocal.getMonth(), nowLocal.getDate());
  const dateKeyToday = (typeof PT !== 'undefined' && PT && typeof PT.yyyymmdd === 'function')
    ? PT.yyyymmdd(dateCell)
    : (nowLocal.getFullYear() + "-" + String(nowLocal.getMonth() + 1).padStart(2, "0") + "-" + String(nowLocal.getDate()).padStart(2, "0"));


  // Phase selection → "open" | "close"
  const rawPhase = opts.auto ? HM_selectPhase(cfg, nowLocal) : (opts.mode || "open");
  mode = (typeof rawPhase === "string") ? rawPhase.toLowerCase() : (rawPhase && typeof rawPhase.name === "string" ? rawPhase.name.toLowerCase() : String(rawPhase).toLowerCase());

  if (mode === "close" && !ss.getSheetByName(sheetName)) {
    LoggerEx.warn(`[History] Close-phase but no history sheet yet → aborting.`);
    return;
  }

  // Source sheet

  const pricesSheet = getOrCreateSheet(ss, "Market Prices", ["date", "market_id", "market_type", "type_id", "min_sell", "max_buy", "median_sell", "median_buy"]);
  const prices = pricesSheet.getDataRange().getValues();
  if (!prices || prices.length <= 1) { LoggerEx.log("[History] Market Prices is empty."); return; }

  const header = prices.shift();
  const IDX = _buildIDX(header);

  // Window filter (~26h)
  const cutoff = new Date(nowLocal.getTime() - 26 * 60 * 60 * 1000);
  _logWindowStats(prices, IDX);
  const windowRows = prices.filter(r => { const d = PT.parseDateSafe(r[IDX.date]); return !isNaN(d) && d >= cutoff; });
  if (!windowRows.length) { LoggerEx.log("[History] No recent rows in window; nothing to do."); return; }


  // dateKeyToday already exists earlier in HM_update (from dateCell)
  for (const r of windowRows) {
    const key = [r[IDX.type_id], r[IDX.market_id], r[IDX.market_type]].join("|");

    // create group if needed (add earliest trackers)
    let g = groups.get(key);
    if (!g) {
      g = {
        type_id: r[IDX.type_id],
        market_id: r[IDX.market_id],
        market_type: r[IDX.market_type],
        minSells: [],
        maxBuys: [],
        latest: { buy: null, sell: null, ts: new Date(0) },
        earliestBuy: { ts: null, val: null },   // NEW
        earliestSell: { ts: null, val: null }    // NEW
      };
      groups.set(key, g);
    }

    // parse values
    const sellVal = Number(r[IDX.min_sell]);
    const buyVal = Number(r[IDX.max_buy]);
    const ts = PT.parseDateSafe(r[IDX.date]);
    if (isNaN(ts)) continue;

    // accumulate daily ranges/medians pools
    if (Number.isFinite(sellVal) && sellVal > 0) g.minSells.push(sellVal);
    if (Number.isFinite(buyVal) && buyVal > 0) g.maxBuys.push(buyVal);

    // latest tick in window (for close/open snapshots)
    if (ts > g.latest.ts) {
      g.latest.ts = ts;
      if (Number.isFinite(buyVal) && buyVal > 0) g.latest.buy = buyVal;
      if (Number.isFinite(sellVal) && sellVal > 0) g.latest.sell = sellVal;
    }

    // EARLIEST OF *TODAY* (for backfilling *_open on close)
    if (PT.yyyymmdd(ts) === dateKeyToday) {
      if (Number.isFinite(buyVal) && buyVal > 0 && (!g.earliestBuy.ts || ts < g.earliestBuy.ts)) {
        g.earliestBuy = { ts, val: buyVal };
      }
      if (Number.isFinite(sellVal) && sellVal > 0 && (!g.earliestSell.ts || ts < g.earliestSell.ts)) {
        g.earliestSell = { ts, val: sellVal };
      }
    }
  }


  // Target sheet
  const HISTORY_HEADER = [
    "type_id", "market_id", "market_type", "date",
    "buy_open", "buy_close", "sell_open", "sell_close",
    "buy_high", "buy_low", "sell_high", "sell_low",
    "median_buy", "median_sell"
  ];
  const historySheet = getOrCreateSheet(ss, sheetName, HISTORY_HEADER);
  // Build target index (dynamic, based on sheet header)
  const histHeader = historySheet.getRange(1, 1, 1, historySheet.getLastColumn())
    .getValues()[0].map(h => String(h || "").trim().toLowerCase());
  const TIDX = {};
  histHeader.forEach((h, i) => TIDX[h] = i);

  // Convenience accessors (required cols)
  const iType = TIDX["type_id"];
  const iMarket = TIDX["market_id"];
  const iMtype = TIDX["market_type"];
  const iDate = TIDX["date"];
  const iBuyOpen = TIDX["buy_open"];
  const iBuyClose = TIDX["buy_close"];
  const iSellOpen = TIDX["sell_open"];
  const iSellClose = TIDX["sell_close"];
  const iBuyHigh = TIDX["buy_high"];
  const iBuyLow = TIDX["buy_low"];
  const iSellHigh = TIDX["sell_high"];
  const iSellLow = TIDX["sell_low"];
  const iMedBuy = TIDX["median_buy"];
  const iMedSell = TIDX["median_sell"];

  const COLS = histHeader.length; // should be 14


  const existing = historySheet.getDataRange().getValues();
  if (existing.length > 0) existing.shift(); // remove header
  const index = new Map();
  for (let i = 0; i < existing.length; i++) {
    const row = existing[i];
    const dateKey = PT.yyyymmdd(row[iDate]);
    const key = [row[iType], row[iMarket], row[iMtype], dateKey].join("|");
    index.set(key, { rowNumber: i + 2, cur: row });
  }





  for (const g of groups.values()) {
    // per-side ranges + medians
    const buyHigh = g.maxBuys.length ? Math.max.apply(null, g.maxBuys) : "";
    const buyLow = g.maxBuys.length ? Math.min.apply(null, g.maxBuys) : "";
    const sellHigh = g.minSells.length ? Math.max.apply(null, g.minSells) : "";
    const sellLow = g.minSells.length ? Math.min.apply(null, g.minSells) : "";
    const medBuy = g.maxBuys.length ? Utility.median(g.maxBuys) : "";
    const medSell = g.minSells.length ? Utility.median(g.minSells) : "";

    const hasSellSide = g.minSells.length > 0 || (Number.isFinite(g.latest.sell) && g.latest.sell > 0);
    const hasBuySide = g.maxBuys.length > 0 || (Number.isFinite(g.latest.buy) && g.latest.buy > 0);
    if (!hasSellSide && !hasBuySide) continue;

    // build base row
    const base = new Array(COLS).fill("");
    base[iType] = g.type_id;
    base[iMarket] = g.market_id;
    base[iMtype] = g.market_type;
    base[iDate] = dateCell;

    if (mode === "open") {
      if (Number.isFinite(g.latest.buy) && g.latest.buy > 0) base[iBuyOpen] = g.latest.buy;
      if (Number.isFinite(g.latest.sell) && g.latest.sell > 0) base[iSellOpen] = g.latest.sell;
    } else { // close
      if (Number.isFinite(g.latest.buy) && g.latest.buy > 0) base[iBuyClose] = g.latest.buy;
      if (Number.isFinite(g.latest.sell) && g.latest.sell > 0) base[iSellClose] = g.latest.sell;

      // backfill *_open if still empty using earliest-of-today
      if ((base[iBuyOpen] === "" || base[iBuyOpen] == null) &&
        g.earliestBuy && Number.isFinite(g.earliestBuy.val) && g.earliestBuy.val > 0) {
        base[iBuyOpen] = g.earliestBuy.val;
      }
      if ((base[iSellOpen] === "" || base[iSellOpen] == null) &&
        g.earliestSell && Number.isFinite(g.earliestSell.val) && g.earliestSell.val > 0) {
        base[iSellOpen] = g.earliestSell.val;
      }
    }

    // per-side ranges + medians
    base[iBuyHigh] = buyHigh;
    base[iBuyLow] = buyLow;
    base[iSellHigh] = sellHigh;
    base[iSellLow] = sellLow;
    base[iMedBuy] = medBuy;
    base[iMedSell] = medSell;

    if (base[iBuyHigh] !== "" && base[iBuyLow] !== "" && base[iBuyLow] > base[iBuyHigh]) {
      LoggerEx.warn(`[HM WARN] buy_low > buy_high for ${g.type_id}|${g.market_id}|${g.market_type}`);
    }
    if (base[iSellHigh] !== "" && base[iSellLow] !== "" && base[iSellLow] > base[iSellHigh]) {
      LoggerEx.warn(`[HM WARN] sell_low > sell_high for ${g.type_id}|${g.market_id}|${g.market_type}`);
    }


    // merge / append
    const composite = [g.type_id, g.market_id, g.market_type, dateKeyToday].join("|");
    const hit = index.get(composite);
    if (hit) {
      const merged = hit.cur.slice();

      const maybeCopy = (curIdx, baseIdx) => {
        const curVal = merged[curIdx];
        const newVal = base[baseIdx];
        const curBlank = (curVal === "" || curVal == null);
        const newGood = Number.isFinite(newVal) && newVal > 0;
        if (curBlank && newGood) merged[curIdx] = newVal;
      };
      maybeCopy(iBuyOpen, iBuyOpen);
      maybeCopy(iSellOpen, iSellOpen);
      maybeCopy(iBuyClose, iBuyClose);
      maybeCopy(iSellClose, iSellClose);

      // always refresh ranges & medians from window
      merged[iBuyHigh] = base[iBuyHigh];
      merged[iBuyLow] = base[iBuyLow];
      merged[iSellHigh] = base[iSellHigh];
      merged[iSellLow] = base[iSellLow];
      merged[iMedBuy] = base[iMedBuy];
      merged[iMedSell] = base[iMedSell];

      if (mode === "close") {
        if ((merged[iBuyOpen] === "" || merged[iBuyOpen] == null) &&
          g.earliestBuy && Number.isFinite(g.earliestBuy.val) && g.earliestBuy.val > 0) {
          merged[iBuyOpen] = g.earliestBuy.val;
          LoggerEx.log(`[HM] Backfilled buy_open  → ${merged[iBuyOpen]} for ${g.type_id}|${g.market_id}|${g.market_type}`);
        }
        if ((merged[iSellOpen] === "" || merged[iSellOpen] == null) &&
          g.earliestSell && Number.isFinite(g.earliestSell.val) && g.earliestSell.val > 0) {
          merged[iSellOpen] = g.earliestSell.val;
          LoggerEx.log(`[HM] Backfilled sell_open → ${merged[iSellOpen]} for ${g.type_id}|${g.market_id}|${g.market_type}`);
        }
      }

      updates.push({ rowNumber: hit.rowNumber, values: merged });
    } else {
      appends.push(base);
    }
  }
  // collect quick stats
  let statBoth = 0, statBuyOnly = 0, statSellOnly = 0;
  for (const g of groups.values()) {
    const hasBuy = g.maxBuys.length || (Number.isFinite(g.latest.buy) && g.latest.buy > 0);
    const hasSell = g.minSells.length || (Number.isFinite(g.latest.sell) && g.latest.sell > 0);
    if (hasBuy && hasSell) statBoth++;
    else if (hasBuy) statBuyOnly++;
    else if (hasSell) statSellOnly++;
  }
  LoggerEx.log(`[HM] window mix — both=${statBoth} buyOnly=${statBuyOnly} sellOnly=${statSellOnly}`);

  // === WRITE SECTION (serialize with DocumentLock) ===
  var dlock = LockService.getDocumentLock();
  if (!dlock.tryLock(250)) {
    try { dlock.waitLock(5000); } catch (e) {
      LoggerEx.warn("[HM] Document busy; skipping writes this run.");
      return;
    }
  }
  try {
    // write updates (contiguous batches in as few calls as possible)
    writeContiguous_(historySheet, updates, COLS);

    // write appends
    if (appends.length) {
      const startRow = historySheet.getLastRow() + 1;
      for (let i = 0; i < appends.length; i += CHUNK) {
        const seg = appends.slice(i, i + CHUNK);
        historySheet.getRange(startRow + i, 1, seg.length, COLS).setValues(seg);
      }
    }

    // retention (optional to include under the same lock — recommended)
    if (retentionDays > 0) {
      const all = historySheet.getDataRange().getValues();
      if (all.length > 1) {
        all.shift();
        const cutoffRet = new Date(nowLocal.getTime() - retentionDays * 24 * 60 * 60 * 1000);
        const toDelete = [];
        for (let i = 0; i < all.length; i++) {
          const d = PT.parseDateSafe(all[i][iDate]);
          if (!isNaN(d) && d < cutoffRet) toDelete.push(i + 2);
        }
        for (let i = toDelete.length - 1; i >= 0; i--) historySheet.deleteRow(toDelete[i]);
      }
    }

    SpreadsheetApp.flush(); // push changes ASAP while we still hold the lock
  } finally {
    try { dlock.releaseLock(); } catch (_) { }
  }


  LoggerEx.log(`[History] Done. mode=${mode} groups=${groups.size} updates=${updates.length} appends=${appends.length}`);
}

/*********************** Phase selection ***********************/
function HM_selectPhase(cfg, nowLocal) {
  try {
    if (typeof projectDayWindowNow_ === "function") {
      const win = projectDayWindowNow_();
      const phaseStr = (typeof win?.phase === "string") ? win.phase : (win?.phase?.name || String(win?.phase || "open"));
      _logPhase(nowLocal, cfg?.OpenTime, cfg?.CloseTime, phaseStr);
      return phaseStr;
    }
  } catch (e) { LoggerEx.warn(`[PHASE DEBUG] projectDayWindowNow_ threw: ${e && e.stack || e}`); }

  const openHM = (typeof Utility !== "undefined" && Utility && typeof Utility.toHM === "function") ? Utility.toHM(cfg?.OpenTime ?? "11:00") : PT.coerceHM(cfg?.OpenTime ?? "11:00");
  const closeHM = (typeof Utility !== "undefined" && Utility && typeof Utility.toHM === "function") ? Utility.toHM(cfg?.CloseTime ?? "18:00") : PT.coerceHM(cfg?.CloseTime ?? "18:00");
  const DUR = parseInt(cfg["History.WindowMinutes"] || 60, 10);
  const inOpen = Utility.inWindow(nowLocal, openHM.h, openHM.m, DUR);
  const inClose = Utility.inWindow(nowLocal, closeHM.h, closeHM.m, DUR);
  _logPhase(nowLocal, cfg?.OpenTime, cfg?.CloseTime, (inOpen ? "open" : (inClose ? "close" : "none")));
  if (inOpen) return "open";
  if (inClose) return "close";
  return "open";
}

function _logPhase(nowLocal, openRaw, closeRaw, phase) {
  const tz = Session.getScriptTimeZone() || "UTC";
  LoggerEx.log(`[PHASE DEBUG] tz=${tz} now=${nowLocal.toLocaleString()} OpenRaw="${openRaw}" CloseRaw="${closeRaw}" phase=${phase}`);
}

/** getOrCreateSheet lives in Utility; use Utility.getOrCreateSheet(ss, name, headers). */

/**
 * Resolve column indexes from header row (case/space-insensitive) with aliases.
 * Throws if a required column is missing (the error lists tried names + header).
 */
function _buildIDX(header) {
  const norm = h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_");
  const map = new Map(); header.forEach((h, i) => map.set(norm(h), i));
  const pick = (...cands) => { for (const c of cands) if (map.has(c)) return map.get(c); throw new Error(`Missing column; tried ${cands.join(",")} | header=[${header.join(" | ")}].`); };
  const IDX = {
    type_id: pick("type_id", "typeid", "item_id", "itemid"),
    market_id: pick("market_id", "marketid", "location_id", "locationid"),
    market_type: pick("market_type", "markettype", "location_type", "locationtype", "scope"),
    date: pick("date", "timestamp", "time", "datetime", "date_time"),
    max_buy: pick("max_buy", "buy", "best_buy", "buy_price", "buy_max"),
    min_sell: pick("min_sell", "sell", "best_sell", "sell_price", "sell_min"),
  };
  LoggerEx.log(`[IDX] type_id=${IDX.type_id} market_id=${IDX.market_id} market_type=${IDX.market_type} date=${IDX.date} max_buy=${IDX.max_buy} min_sell=${IDX.min_sell}`);
  return IDX;
}

function writeContiguous_(sheet, updates, cols) {
  if (!updates || !updates.length) return;

  // Tweak these if your item list grows a lot
  var MAX_ROWS_PER_FLUSH = 5000;    // ~5k rows/flush; with 14 cols ≈ 70k cells
  var SLEEP_BETWEEN_MS = 30;      // tiny breather between flushes

  // Sort by row so we can coalesce adjacent writes
  updates.sort(function (a, b) { return a.rowNumber - b.rowNumber; });

  var start = updates[0].rowNumber;
  var buf = [updates[0].values];

  function flush() {
    sheet.getRange(start, 1, buf.length, cols).setValues(buf);
    SpreadsheetApp.flush();                     // push to server
    if (SLEEP_BETWEEN_MS) Utilities.sleep(SLEEP_BETWEEN_MS);
  }

  for (var i = 1; i < updates.length; i++) {
    var u = updates[i];
    var expectedNext = start + buf.length;

    if (u.rowNumber === expectedNext && buf.length < MAX_ROWS_PER_FLUSH) {
      // still contiguous and under cap
      buf.push(u.values);
    } else {
      // either a gap or we hit the cap — flush and start new block
      flush();
      start = u.rowNumber;
      buf = [u.values];
    }
  }
  // last block
  flush();
}

/**
 * Log oldest/newest timestamps and number of bad dates for a quick sanity check.
 */
function _logWindowStats(rows, IDX) {
  let minTs = null, maxTs = null, bad = 0;
  for (const r of rows) {
    const d = PT.parseDateSafe(r[IDX.date]);
    if (isNaN(d)) { bad++; continue; }
    if (!minTs || d < minTs) minTs = d;
    if (!maxTs || d > maxTs) maxTs = d;
  }
  LoggerEx.log(`[WINDOW] total=${rows.length} valid=${rows.length - bad} badDates=${bad} oldest=${minTs ? minTs.toLocaleString() : "-"} newest=${maxTs ? maxTs.toLocaleString() : "-"}`);
}
