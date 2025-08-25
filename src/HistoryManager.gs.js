/***********************
 * History Manager (clean, commented)
 * Scope: uses global getOrCreateSheet, PT, and Utility only.
 ***********************/

/*********************** Entrypoints ***********************/
function updateHistory()      { return HM_update({ auto: true,  test: false }); }
function updateHistoryTest()  { return HM_update({ auto: false, test: true, mode: "open" }); }

/*********************** Core ***********************/
function HM_update(opts) {
  opts = opts || {};
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = (typeof getConfig === "function") ? (getConfig() || {}) : {};

  const sheetName     = cfg["History.SheetName"] || "Market History";
  const retentionDays = parseInt(cfg["History.RetentionDays"] || cfg["RetentionDays"] || 365, 10);
  const CHUNK         = parseInt(cfg["History.ChunkSize"]    || cfg["ChunkSize"]    || 250, 10);

  const nowLocal = new Date();

  // Phase selection → "open" | "close"
  const rawPhase = opts.auto ? HM_selectPhase(cfg, nowLocal) : (opts.mode || "open");
  const mode = (typeof rawPhase === "string") ? rawPhase.toLowerCase() : (rawPhase && typeof rawPhase.name === "string" ? rawPhase.name.toLowerCase() : String(rawPhase).toLowerCase());

  if (mode === "close" && !ss.getSheetByName(sheetName)) {
    _log(`[History] Close-phase but no history sheet yet → aborting.`);
    return;
  }

  // Source sheet
  const pricesSheet = getOrCreateSheet(ss, "Market Prices", ["date","market_id","market_type","type_id","min_sell","max_buy","median_sell","median_buy"]);
  const prices = pricesSheet.getDataRange().getValues();
  if (!prices || prices.length <= 1) { _log("[History] Market Prices is empty."); return; }

  const header = prices.shift();
  const IDX = _buildIDX(header);

  // Window filter (~26h)
  const cutoff = new Date(nowLocal.getTime() - 26 * 60 * 60 * 1000);
  _logWindowStats(prices, IDX);
  const windowRows = prices.filter(r => { const d = PT.parseDateSafe(r[IDX.date]); return !isNaN(d) && d >= cutoff; });
  if (!windowRows.length) { _log("[History] No recent rows in window; nothing to do."); return; }

  // Group rows by (type|market|mtype)
  const groups = new Map();
  for (const r of windowRows) {
    const key = [r[IDX.type_id], r[IDX.market_id], r[IDX.market_type]].join("|");
    let g = groups.get(key);
    if (!g) { g = { type_id:r[IDX.type_id], market_id:r[IDX.market_id], market_type:r[IDX.market_type], minSells:[], maxBuys:[], latest:{buy:null, sell:null, ts:new Date(0)} }; groups.set(key, g); }
    const sellVal = Number(r[IDX.min_sell]); if (Number.isFinite(sellVal) && sellVal > 0) g.minSells.push(sellVal);
    const buyVal  = Number(r[IDX.max_buy]);  if (Number.isFinite(buyVal)  && buyVal  > 0) g.maxBuys.push(buyVal);
    const ts = PT.parseDateSafe(r[IDX.date]); if (ts > g.latest.ts) { g.latest.ts = ts; if (Number.isFinite(buyVal) && buyVal > 0) g.latest.buy = buyVal; if (Number.isFinite(sellVal) && sellVal > 0) g.latest.sell = sellVal; }
  }

  // Target sheet
  const historySheet = getOrCreateSheet(ss, sheetName, ["type_id","market_id","market_type","date","buy_open","buy_close","sell_open","sell_close","daily_high","daily_low","median_buy","median_sell"]);

  const existing = historySheet.getDataRange().getValues();
  if (existing.length && existing[0] && existing[0][0] !== "type_id") { /* header already ensured by getOrCreateSheet */ }
  if (existing.length > 0) existing.shift();
  const HIDX = { type_id:0, market_id:1, market_type:2, date:3, buy_open:4, buy_close:5, sell_open:6, sell_close:7, daily_high:8, daily_low:9, median_buy:10, median_sell:11 };

  const index = new Map();
  for (let i=0;i<existing.length;i++) {
    const row = existing[i]; const dateKey = PT.yyyymmdd(row[HIDX.date]);
    index.set([row[HIDX.type_id], row[HIDX.market_id], row[HIDX.market_type], dateKey].join("|"), { rowNumber: i+2, cur: row });
  }

  const dateCell = (typeof PT !== 'undefined' && PT && typeof PT.todayAt === 'function') ? PT.todayAt(0,0,0) : new Date(nowLocal.getFullYear(), nowLocal.getMonth(), nowLocal.getDate());
  const dateKeyToday = PT.yyyymmdd(dateCell);
  const appends = []; const updates = [];

  for (const g of groups.values()) {
    const dayHigh = g.minSells.length ? Math.max.apply(null, g.minSells) : "";
    const dayLow  = g.maxBuys.length ? Math.min.apply(null, g.maxBuys) : "";
    const medBuy  = g.maxBuys.length ? Utility.median(g.maxBuys)      : "";
    const medSell = g.minSells.length ? Utility.median(g.minSells)    : "";

    const hasSellSide = g.minSells.length > 0 || (Number.isFinite(g.latest.sell) && g.latest.sell > 0);
    const hasBuySide  = g.maxBuys.length > 0 || (Number.isFinite(g.latest.buy)  && g.latest.buy  > 0);
    if (!hasSellSide && !hasBuySide) continue;

    const base = [ g.type_id, g.market_id, g.market_type, dateCell, "", "", "", "", dayHigh, dayLow, medBuy, medSell ];
    if (mode === "open") {
      if (Number.isFinite(g.latest.buy)  && g.latest.buy  > 0) base[HIDX.buy_open]  = g.latest.buy;
      if (Number.isFinite(g.latest.sell) && g.latest.sell > 0) base[HIDX.sell_open] = g.latest.sell;
    } else {
      if (Number.isFinite(g.latest.buy)  && g.latest.buy  > 0) base[HIDX.buy_close]  = g.latest.buy;
      if (Number.isFinite(g.latest.sell) && g.latest.sell > 0) base[HIDX.sell_close] = g.latest.sell;
    }

    const composite = [g.type_id, g.market_id, g.market_type, dateKeyToday].join("|");
    const hit = index.get(composite);
    if (hit) {
      const merged = hit.cur.slice();
      const maybeCopy = (curIdx, baseIdx) => { const curVal = merged[curIdx]; const newVal = base[baseIdx]; const curBlank = (curVal === "" || curVal == null); const newGood = Number.isFinite(newVal) && newVal > 0; if (curBlank && newGood) merged[curIdx] = newVal; };
      maybeCopy(HIDX.buy_open,  HIDX.buy_open);
      maybeCopy(HIDX.sell_open, HIDX.sell_open);
      maybeCopy(HIDX.buy_close, HIDX.buy_close);
      maybeCopy(HIDX.sell_close,HIDX.sell_close);
      merged[HIDX.daily_high]  = base[HIDX.daily_high];
      merged[HIDX.daily_low]   = base[HIDX.daily_low];
      merged[HIDX.median_buy]  = base[HIDX.median_buy];
      merged[HIDX.median_sell] = base[HIDX.median_sell];
      updates.push({ rowNumber: hit.rowNumber, values: merged });
    } else {
      appends.push(base);
    }
  }

  if (updates.length) {
    for (let i = 0; i < updates.length; i += CHUNK) {
      const batch = updates.slice(i, i + CHUNK);
      batch.forEach(b => historySheet.getRange(b.rowNumber, 1, 1, 12).setValues([b.values]));
    }
  }

  if (appends.length) {
    const startRow = historySheet.getLastRow() + 1;
    for (let i = 0; i < appends.length; i += CHUNK) {
      const seg = appends.slice(i, i + CHUNK);
      historySheet.getRange(startRow + i, 1, seg.length, 12).setValues(seg);
    }
  }

  if (retentionDays > 0) {
    const all = historySheet.getDataRange().getValues();
    if (all.length > 1) {
      all.shift();
      const cutoffRet = new Date(nowLocal.getTime() - retentionDays * 24 * 60 * 60 * 1000);
      const toDelete = [];
      for (let i = 0; i < all.length; i++) { const d = PT.parseDateSafe(all[i][HIDX.date]); if (!isNaN(d) && d < cutoffRet) toDelete.push(i + 2); }
      for (let i = toDelete.length - 1; i >= 0; i--) historySheet.deleteRow(toDelete[i]);
    }
  }

  _log(`[History] Done. mode=${mode} groups=${groups.size} updates=${updates.length} appends=${appends.length}`);
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
  } catch (e) { _log(`[PHASE DEBUG] projectDayWindowNow_ threw: ${e && e.stack || e}`); }

  const openHM  = (typeof Utility !== "undefined" && Utility && typeof Utility.toHM === "function") ? Utility.toHM(cfg?.OpenTime  ?? "11:00") : PT.coerceHM(cfg?.OpenTime  ?? "11:00");
  const closeHM = (typeof Utility !== "undefined" && Utility && typeof Utility.toHM === "function") ? Utility.toHM(cfg?.CloseTime ?? "18:00") : PT.coerceHM(cfg?.CloseTime ?? "18:00");
  const DUR = parseInt(cfg["History.WindowMinutes"] || 60, 10);
  const inOpen  = Utility.inWindow(nowLocal, openHM.h,  openHM.m,  DUR);
  const inClose = Utility.inWindow(nowLocal, closeHM.h, closeHM.m, DUR);
  _logPhase(nowLocal, cfg?.OpenTime, cfg?.CloseTime, (inOpen ? "open" : (inClose ? "close" : "none")));
  if (inOpen)  return "open";
  if (inClose) return "close";
  return "open";
}

function _logPhase(nowLocal, openRaw, closeRaw, phase) {
  const tz = Session.getScriptTimeZone() || "UTC";
  _log(`[PHASE DEBUG] tz=${tz} now=${nowLocal.toLocaleString()} OpenRaw="${openRaw}" CloseRaw="${closeRaw}" phase=${phase}`);
}

/** getOrCreateSheet lives in Utility; use Utility.getOrCreateSheet(ss, name, headers). */

// Logger shims — prefer LoggerEx; fallback to console
function _log()      { try { LoggerEx.log.apply(LoggerEx, arguments); }   catch (e) { try{ console.log([].slice.call(arguments).join(' ')); }catch(_){} } }
function _logWarn()  { try { LoggerEx.warn.apply(LoggerEx, arguments); }  catch (e) { try{ console.warn([].slice.call(arguments).join(' ')); }catch(_){} } }
function _logError() { try { LoggerEx.error.apply(LoggerEx, arguments); } catch (e) { try{ console.error([].slice.call(arguments).join(' ')); }catch(_){} } }
function _logDebug() { try { LoggerEx.debug.apply(LoggerEx, arguments); } catch (e) { try{ console.log([].slice.call(arguments).join(' ')); }catch(_){} } }


/**
 * Resolve column indexes from header row (case/space-insensitive) with aliases.
 * Throws if a required column is missing (the error lists tried names + header).
 */
function _buildIDX(header) {
  const norm = h => String(h || "").trim().toLowerCase().replace(/\s+/g, "_");
  const map = new Map(); header.forEach((h, i) => map.set(norm(h), i));
  const pick = (...cands) => { for (const c of cands) if (map.has(c)) return map.get(c); throw new Error(`Missing column; tried ${cands.join(",")} | header=[${header.join(" | ")}].`); };
  const IDX = {
    type_id:     pick("type_id","typeid","item_id","itemid"),
    market_id:   pick("market_id","marketid","location_id","locationid"),
    market_type: pick("market_type","markettype","location_type","locationtype","scope"),
    date:        pick("date","timestamp","time","datetime","date_time"),
    max_buy:     pick("max_buy","buy","best_buy","buy_price","buy_max"),
    min_sell:    pick("min_sell","sell","best_sell","sell_price","sell_min"),
  };
  _log(`[IDX] type_id=${IDX.type_id} market_id=${IDX.market_id} market_type=${IDX.market_type} date=${IDX.date} max_buy=${IDX.max_buy} min_sell=${IDX.min_sell}`);
  return IDX;
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
  _log(`[WINDOW] total=${rows.length} valid=${rows.length-bad} badDates=${bad} oldest=${minTs?minTs.toLocaleString():"-"} newest=${maxTs?maxTs.toLocaleString():"-"}`);
}
