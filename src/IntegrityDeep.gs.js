/** ===== Deep Integrity (heavy – use when needed) ===== **/

/** ---------------- Schema ---------------- */
const REQUIRED_MP_KEYS  = ["date","market_id","market_type","type_id","max_buy","min_sell"];
const OPTIONAL_MP_KEYS  = ["median_sell","median_buy"];

const REQUIRED_MH_KEYS  = ["type_id","market_id","market_type","date","buy_open","buy_close","sell_open","sell_close"];
const OPTIONAL_MH_KEYS  = ["daily_high","daily_low","median_buy","median_sell"];

// Limit scan + log volume
const ID_CHECK_WINDOW_HOURS = 3;  // set 0 to disable (scan all)
const ID_WARN_SAMPLE_LIMIT  = 20;  // per warning-kind per sheet

// If false, we only WARN/INFO when Market History is missing and continue.
// If true, missing sheet is an error and Deep fails.
const ID_REQUIRE_MARKET_HISTORY = false;

/** ---------------- Header mapping (order-agnostic) ---------------- */
function _id_asDate(v) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) return v;
  const t = (typeof v === "string") ? Date.parse(v) : NaN;
  return isNaN(t) ? null : new Date(t);
}

function _normHeader_(h) {
  return String(h || "")
    .replace(/\u00A0/g, " ")     // NBSP → space
    .trim()
    .toLowerCase()
    .replace(/[^\w]+/g, "_")     // spaces/punct → _
    .replace(/^_+|_+$/g, "");
}
function _aliasHeader_(k) {
  const alias = {
    // common variants
    typeid: "type_id",
    type_id: "type_id",
    marketid: "market_id",
    market_id: "market_id",
    market_type: "market_type",
    "market-type": "market_type",
    minsell: "min_sell",
    "min_sell": "min_sell",
    maxbuy: "max_buy",
    "max_buy": "max_buy",
    mediansell: "median_sell",
    "median_sell": "median_sell",
    medianbuy: "median_buy",
    "median_buy": "median_buy",
    date: "date",
    buyopen: "buy_open",
    buy_close: "buy_close",
    buyclose: "buy_close",
    sellopen: "sell_open",
    sell_close: "sell_close",
    sellclose: "sell_close",
    dailyhigh: "daily_high",
    dailylow: "daily_low"
  };
  return alias[k] || k;
}
function _headerIndex_(headers) {
  const idx = {};
  headers.forEach((h, i) => {
    const key = _aliasHeader_(_normHeader_(h));
    if (key) idx[key] = i;
  });
  return idx;
}
function _noticeHeaderChanges_(headers, requiredKeys, optionalKeys, sheetName, log) {
  const idx = _headerIndex_(headers);
  const seen = Object.keys(idx);

  const requiredMissing = requiredKeys.filter(k => !(k in idx));
  const known = new Set([ ...requiredKeys, ...optionalKeys ]);
  const newCols = seen.filter(k => !known.has(k));

  if (newCols.length) {
    _id_info(log, `[NOTICE] ${sheetName}: new columns detected → ${JSON.stringify(newCols)}`, sheetName);
    _id_info(log, `[NOTICE] Consider adding to OPTIONAL_* and updating tests.`, sheetName);
  }
  if (requiredMissing.length) {
    _id_err(log, `Missing required columns → ${JSON.stringify(requiredMissing)}`, sheetName);
  }
  return { idx, requiredMissing, newCols };
}

/** Reassemble a row into this canonical order (missing keys become null). */
function _canonicalRow_(row, idx, order) {
  return order.map(k => (k in idx ? row[idx[k]] : null));
}

/** ---------------- Public: main deep integrity ---------------- */
function Integrity_Deep() {
  const log = (typeof Log !== "undefined") ? Log.for("Integrity") : null;
  try {
    const ok1 = _id_checkMarketPrices(log);
    const ok2 = _id_checkMarketHistory(log);
    const ok3 = _id_crossKeys(log);
    const allOk = ok1 && ok2 && ok3;
    const msg = allOk ? "Integrity_Deep: PASS" : "Integrity_Deep: FAIL (see logs)";
    log ? (allOk ? log.info(msg) : log.error(msg)) : Logger.log(msg);
    return allOk;
  } catch (e) {
    log ? log.error("Integrity_Deep: exception", { err: String(e) }) : Logger.log("Integrity_Deep exception: " + e);
    return false;
  }
}

/** ---------------- Market Prices check (order-agnostic) ---------------- */
function _id_checkMarketPrices(log) {
  const sh = SpreadsheetApp.getActive().getSheetByName("Market Prices");
  if (!sh) { _id_err(log, "Missing sheet", "Market Prices"); return false; }

  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const { idx, requiredMissing } = _noticeHeaderChanges_(header, REQUIRED_MP_KEYS, OPTIONAL_MP_KEYS, "Market Prices", log);

  if (requiredMissing.length) return false;

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) { _id_info(log, "No data rows", "Market Prices"); return true; }

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Compute cutoff
  const cutoff = ID_CHECK_WINDOW_HOURS > 0
    ? new Date(Date.now() - ID_CHECK_WINDOW_HOURS * 60 * 60 * 1000)
    : null;

  let errors = 0, warns = 0;
  let warnMissingBuy = 0, warnMissingSell = 0, warnInverted = 0;

  const sampleCount = { missingBuy: 0, missingSell: 0, inverted: 0 };

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const [date, market_id, market_type, type_id, max_buy, min_sell, median_sell, median_buy]
      = _canonicalRow_(row, idx, ["date","market_id","market_type","type_id","max_buy","min_sell","median_sell","median_buy"]);

    // Skip by date window if configured
    if (cutoff) {
      const d = _id_asDate(date);
      if (d && d < cutoff) continue;
    }

    // basic validations
    if (!_id_posInt(type_id)) errors += _id_row(log, "type_id must be a positive integer", "Market Prices", i + 2);
    if (!_id_posInt(market_id)) errors += _id_row(log, "market_id must be a positive integer", "Market Prices", i + 2);
    if (_id_blank(market_type)) errors += _id_row(log, "market_type required", "Market Prices", i + 2);

    // allow blanks (no active orders) → warn; non-blank non-numeric → error
    if (_id_blank(max_buy)) {
      warnMissingBuy++;
      if (sampleCount.missingBuy < ID_WARN_SAMPLE_LIMIT) {
        warns += _id_warn(log, "max_buy missing (no active buy orders?)", "Market Prices", i + 2);
        sampleCount.missingBuy++;
      }
    } else if (!_id_num(max_buy)) {
      errors += _id_row(log, "max_buy must be numeric", "Market Prices", i + 2);
    }

    if (_id_blank(min_sell)) {
      warnMissingSell++;
      if (sampleCount.missingSell < ID_WARN_SAMPLE_LIMIT) {
        warns += _id_warn(log, "min_sell missing (no active sell orders?)", "Market Prices", i + 2);
        sampleCount.missingSell++;
      }
    } else if (!_id_num(min_sell)) {
      errors += _id_row(log, "min_sell must be numeric", "Market Prices", i + 2);
    }

    if (_id_num(max_buy) && _id_num(min_sell) && Number(max_buy) > Number(min_sell)) {
      warnInverted++;
      if (sampleCount.inverted < ID_WARN_SAMPLE_LIMIT) {
        warns += _id_warn(log, "max_buy > min_sell (spread inverted?)", "Market Prices", i + 2);
        sampleCount.inverted++;
      }
    }

    if (!_id_date(date)) errors += _id_row(log, "date must be a valid date", "Market Prices", i + 2);
  }

  // Summaries
  if (warnMissingBuy   > sampleCount.missingBuy)
    _id_info(log, `${warnMissingBuy} rows with missing max_buy (sampled ${sampleCount.missingBuy})`, "Market Prices");
  if (warnMissingSell  > sampleCount.missingSell)
    _id_info(log, `${warnMissingSell} rows with missing min_sell (sampled ${sampleCount.missingSell})`, "Market Prices");
  if (warnInverted     > sampleCount.inverted)
    _id_info(log, `${warnInverted} rows with inverted spread (sampled ${sampleCount.inverted})`, "Market Prices");

  _id_report(log, "Market Prices", data.length, errors, warns);
  return errors === 0;
}

/** ---------------- Market History check (order-agnostic) ---------------- */
function _id_checkMarketHistory(log) {
  const sh = SpreadsheetApp.getActive().getSheetByName("Market History");
  if (!sh) {
    if (ID_REQUIRE_MARKET_HISTORY) {
      _id_err(log, "Missing sheet", "Market History");
      return false;
    } else {
      _id_warn(log, "Skipped (sheet missing)", "Market History");
      return true; // ← treat as skipped, not failure
    }
  }

  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const { idx, requiredMissing } = _noticeHeaderChanges_(header, REQUIRED_MH_KEYS, OPTIONAL_MH_KEYS, "Market History", log);

  if (requiredMissing.length) return false;

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) { _id_info(log, "No data rows", "Market History"); return true; }

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let errors = 0, warns = 0;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const [
      type_id, market_id, market_type, date,
      buy_open, buy_close, sell_open, sell_close,
      daily_high, daily_low, median_buy, median_sell
    ] = _canonicalRow_(row, idx, [
      "type_id","market_id","market_type","date",
      "buy_open","buy_close","sell_open","sell_close",
      "daily_high","daily_low","median_buy","median_sell"
    ]);

    if (!_id_posInt(type_id)) errors += _id_row(log, "type_id must be a positive integer", "Market History", i + 2);
    if (!_id_posInt(market_id)) errors += _id_row(log, "market_id must be a positive integer", "Market History", i + 2);
    if (_id_blank(market_type)) errors += _id_row(log, "market_type required", "Market History", i + 2);
    if (!_id_date(date)) errors += _id_row(log, "date must be a valid date", "Market History", i + 2);

    // numeric checks
    const numOK = [buy_open, buy_close, sell_open, sell_close].every(_id_num);
    if (!numOK) errors += _id_row(log, "open/close fields must be numeric", "Market History", i + 2);

    // range sanity if highs/lows present
    if (_id_num(daily_high) && _id_num(daily_low) && Number(daily_high) >= Number(daily_low)) {
      const values = [buy_open, buy_close, sell_open, sell_close].filter(_id_num).map(Number);
      const out = values.some(v => v > Number(daily_high) || v < Number(daily_low));
      if (out) warns += _id_warn(log, "one or more price fields outside [daily_low, daily_high]", "Market History", i + 2);
    }
  }

  _id_report(log, "Market History", data.length, errors, warns);
  return errors === 0;
}

/** ---------------- Cross-key check (uses normalized headers) ---------------- */
function _id_crossKeys(log) {
  const ss = SpreadsheetApp.getActive();
  const mp = ss.getSheetByName("Market Prices");
  const mh = ss.getSheetByName("Market History");

  // Market Prices is core: if it's missing, that's always an error.
  if (!mp) { _id_err(log, "Missing sheet", "Market Prices"); return false; }

  // Market History is optional unless required by flag.
  if (!mh) {
    if (ID_REQUIRE_MARKET_HISTORY) {
      _id_err(log, "Missing sheet", "Market History");
      return false;                 // fail when required
    } else {
      _id_info(log, "Skipped (sheet missing)", "Cross");
      return true;                  // pass when optional
    }
  }

  // Both present → do the key comparison (warn-only)
  const keysMP = _id_keys(mp, ["type_id","market_id","market_type"]);
  const keysMH = _id_keys(mh, ["type_id","market_id","market_type"]);
  let missing = 0;
  for (const k of keysMH) if (!keysMP.has(k)) missing++;
  if (missing > 0) _id_warn(log, "Some history keys not present in Market Prices", "Cross", null, { missing });
  else _id_info(log, "Cross-keys OK", "Cross");

  return true; // comparison is warn-only
}

/** ---------------- Utilities ---------------- */
function _id_report(log, name, rows, errors, warns) {
  const msg = errors === 0 && warns === 0
    ? `${name}: OK`
    : errors === 0 ? `${name}: OK with warnings (${warns})`
    : `${name}: FAILED (errors=${errors}, warns=${warns})`;
  if (typeof Log !== "undefined") {
    if (errors === 0) Log.for("Integrity").info(msg, { rows });
    else Log.for("Integrity").error(msg, { rows });
  } else {
    Logger.log(msg);
  }
}
function _id_row(log, msg, sheet, row) { _id_err(log, `${msg}`, sheet, row); return 1; }
function _id_err(log, msg, sheet, row) {
  if (log) log.error(msg, { sheet, row });
  else Logger.log(`[ERROR] ${sheet} row ${row || "-"}: ${msg}`);
}
function _id_warn(log, msg, sheet, row, meta) {
  if (log) log.warn(msg, Object.assign({ sheet, row }, meta || {}));
  else Logger.log(`[WARN] ${sheet} row ${row || "-"}: ${msg}`);
  return 1;
}
function _id_info(log, msg, sheet) {
  if (log) log.info(msg, { sheet });
  else Logger.log(`[INFO] ${sheet}: ${msg}`);
}
function _id_num(v) { return v !== "" && v !== null && !isNaN(Number(v)); }
function _id_blank(v) { return v === "" || v === null; }
function _id_posInt(v) { return _id_num(v) && Number.isInteger(Number(v)) && Number(v) > 0; }
function _id_date(v) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) return true;
  const t = (typeof v === "string") ? Date.parse(v) : NaN;
  return !isNaN(t);
}
/** Build a Set of keys "type_id|market_id|market_type" using normalized headers. */
function _id_keys(sheet, cols) {
  const lastCol = Math.max(3, sheet.getLastColumn());
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idxMap = _headerIndex_(header);
  // ensure all requested cols exist
  for (const c of cols) if (!(c in idxMap)) return new Set();

  const last = sheet.getLastRow();
  const set = new Set();
  if (last <= 1) return set;

  const data = sheet.getRange(2, 1, last - 1, lastCol).getValues();
  for (const r of data) {
    const type_id = r[idxMap["type_id"]];
    const market_id = r[idxMap["market_id"]];
    const market_type = String(r[idxMap["market_type"]] || "").toLowerCase();
    set.add(`${type_id}|${market_id}|${market_type}`);
  }
  return set;
}