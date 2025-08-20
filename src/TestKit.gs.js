/** ================== Quick Wins Helpers ================== **/

function sanityCheck() {
  const needed = [
    ["Market Prices", ["type_id","market_id","market_type","max_buy","min_sell","date"]],
    ["Market History", ["type_id","market_id","market_type","date","buy_open","buy_close","sell_open","sell_close","daily_high","daily_low","median_buy","median_sell"]],
  ];
  const ss = SpreadsheetApp.getActive();
  for (const [name, headers] of needed) {
    const sh = ss.getSheetByName(name);
    if (!sh) throw new Error(`Missing sheet: ${name}`);
    const got = sh.getRange(1,1,1,headers.length).getValues()[0];
    if (got.join("|") !== headers.join("|")) {
      throw new Error(`Header mismatch on ${name}. Expected: ${headers.join(", ")}`);
    }
  }
  return true;
}

function readUsedRange(sheet) {
  const lastRow = Math.max(1, sheet.getLastRow());
  const lastCol = Math.max(1, sheet.getLastColumn());
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

// Standalone timing helper (works with or without Logger)
function timeIt(label, fn) {
  const t0 = Date.now();
  const out = fn();
  const ms = Date.now() - t0;
  try {
    // If Logger is available, mirror to it
    if (typeof Log !== "undefined") Log.for("timeIt").info(label, { ms });
  } catch (_){}
  Logger.log(`${label}: ${ms} ms`);
  return out;
}

function safeAppendRows(sheet, rows) {
  if (!rows || !rows.length) return;
  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

const CACHE_NS = "EVE_TRACKER";
function cacheGetAll(keys){
  const c = CacheService.getScriptCache();
  const kv = c.getAll(keys.map(k=>`${CACHE_NS}:${k}`));
  const out = {};
  keys.forEach(k => { out[k] = kv[`${CACHE_NS}:${k}`] || null; });
  return out;
}
function cacheSetAll(obj, seconds){
  const c = CacheService.getScriptCache();
  const kv = {};
  Object.keys(obj).forEach(k => kv[`${CACHE_NS}:${k}`] = String(obj[k]));
  c.putAll(kv, seconds || 300);
}

function isTestMode(){ 
  return PropertiesService.getScriptProperties().getProperty("TEST_MODE")==="true";
}
function enableTestMode(){ PropertiesService.getScriptProperties().setProperty("TEST_MODE","true"); }
function disableTestMode(){ PropertiesService.getScriptProperties().setProperty("TEST_MODE","false"); }

function protectHeaders(sheetName){
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return;
  const p = sh.protect().setDescription(`${sheetName} headers`);
  p.setUnprotectedRanges([sh.getRange(1,1,1,sh.getLastColumn())]);
}

function logTriggers(){
  ScriptApp.getProjectTriggers().forEach(t=>{
    Logger.log(`${t.getHandlerFunction()} — ${t.getEventType()}`);
  });
}

/** ================== Test Harness Switches ================== **/

function runHistoryTest() {
  const useLogger = (typeof Log !== "undefined");
  if (useLogger) return Log.for("Harness").time("updateHistoryTest", () => updateHistoryTest());
  return timeIt("updateHistoryTest", () => updateHistoryTest());
}
function runHistoryProd() {
  const useLogger = (typeof Log !== "undefined");
  if (useLogger) return Log.for("Harness").time("updateHistory", () => updateHistory());
  return timeIt("updateHistory", () => updateHistory());
}
function runTestModeOn(){ enableTestMode(); }
function runTestModeOff(){ disableTestMode(); }

/** ================== Table Integrity Check ================== **/

function TableIntegrity_Check() {
  const log = (typeof Log !== "undefined") ? Log.for("Integrity") : null;

  const EXPECT_MP = ["type_id","market_id","market_type","max_buy","min_sell","date"];
  const EXPECT_MH = ["type_id","market_id","market_type","date","buy_open","buy_close","sell_open","sell_close","daily_high","daily_low","median_buy","median_sell"];

  const ss = SpreadsheetApp.getActive();
  const shMP = ss.getSheetByName("Market Prices");
  const shMH = ss.getSheetByName("Market History");

  let ok = true;

  ok = checkSheet(shMP, "Market Prices", EXPECT_MP, validateRow_MP, log) && ok;
  ok = checkSheet(shMH, "Market History", EXPECT_MH, validateRow_MH, log) && ok;

  // Cross-table sanity
  try {
    const keysMP = collectKeys(shMP, EXPECT_MP);
    const keysMH = collectKeys(shMH, EXPECT_MH);
    let missingInMP = 0;
    for (const k of keysMH) if (!keysMP.has(k)) missingInMP++;
    if (missingInMP > 0) {
      log ? log.warn("Some history keys aren’t present in Market Prices", { missing_count: missingInMP })
          : Logger.log(`[WARN] Some history keys aren’t present in Market Prices: ${missingInMP}`);
    } else {
      log ? log.info("Cross-table key check OK (history keys present in prices).")
          : Logger.log("Cross-table key check OK (history keys present in prices).");
    }
  } catch (e) {
    log ? log.warn("Cross-table key check skipped due to error", { err: String(e) })
        : Logger.log(`[WARN] Cross-table key check skipped: ${e}`);
  }

  if (ok) { log ? log.info("TableIntegrity_Check PASS") : Logger.log("TableIntegrity_Check PASS"); }
  else    { log ? log.error("TableIntegrity_Check FAIL (see above logs)") : Logger.log("TableIntegrity_Check FAIL"); }

  return ok;
}

function checkSheet(sheet, name, expectedHeaders, rowValidator, log) {
  if (!sheet) {
    log ? log.error("Missing sheet", { sheet: name }) : Logger.log(`[ERROR] Missing sheet: ${name}`);
    return false;
  }
  const header = sheet.getRange(1,1,1,expectedHeaders.length).getValues()[0];
  const headerMatch = header.join("|") === expectedHeaders.join("|");
  if (!headerMatch) {
    log ? log.error("Header mismatch", { sheet: name, expected: expectedHeaders, got: header })
        : Logger.log(`[ERROR] Header mismatch on ${name}`);
    return false;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    log ? log.info("No data rows to validate", { sheet: name })
        : Logger.log(`[INFO] No data rows to validate: ${name}`);
    return true;
  }
  const data = sheet.getRange(2,1,lastRow-1,expectedHeaders.length).getValues();

  let errors = 0, warns = 0;
  for (let i=0;i<data.length;i++){
    const row = data[i];
    const res = rowValidator(row);
    if (!res) continue;
    if (res.error && res.error.length){
      errors += res.error.length;
      res.error.forEach(msg => log ? log.error(msg, { sheet: name, row: i+2, row_preview: safePreview(row) })
                                   : Logger.log(`[ERROR] ${name} row ${i+2}: ${msg}`));
    }
    if (res.warn && res.warn.length){
      warns += res.warn.length;
      res.warn.forEach(msg => log ? log.warn(msg, { sheet: name, row: i+2, row_preview: safePreview(row) })
                                  : Logger.log(`[WARN] ${name} row ${i+2}: ${msg}`));
    }
  }

  if (errors === 0 && warns === 0) {
    log ? log.info("Sheet OK", { sheet: name, rows_checked: data.length })
        : Logger.log(`[INFO] Sheet OK: ${name}, rows: ${data.length}`);
    return true;
  } else if (errors === 0) {
    log ? log.info("Sheet OK with warnings", { sheet: name, rows_checked: data.length, warns })
        : Logger.log(`[INFO] Sheet OK with warnings: ${name}, warns: ${warns}`);
    return true;
  } else {
    log ? log.error("Sheet FAILED", { sheet: name, rows_checked: data.length, errors, warns })
        : Logger.log(`[ERROR] Sheet FAILED: ${name}, errors: ${errors}, warns: ${warns}`);
    return false;
  }
}

/** ---------- Validators ---------- **/
function validateRow_MP(row){
  const errs = [], warns = [];
  const [type_id, market_id, market_type, max_buy, min_sell, dt] = row;

  if (!isPositiveInt(type_id)) errs.push("type_id must be positive integer");
  if (!isPositiveInt(market_id)) errs.push("market_id must be positive integer");

  const mt = String(market_type || "").toLowerCase();
  if (!["region","system","station"].includes(mt)) errs.push("market_type must be region|system|station");

  if (!isNumberish(max_buy)) errs.push("max_buy must be numeric");
  if (!isNumberish(min_sell)) errs.push("min_sell must be numeric");

  if (isNumberish(max_buy) && isNumberish(min_sell) && Number(max_buy) > Number(min_sell)) {
    warns.push("max_buy > min_sell (spread inverted?)");
  }

  if (!isValidDate(dt)) errs.push("date must be a valid date");
  return { error: errs, warn: warns };
}

function validateRow_MH(row){
  const errs = [], warns = [];
  const [type_id, market_id, market_type, dt, bo, bc, so, sc, hi, lo, mb, ms] = row;

  if (!isPositiveInt(type_id)) errs.push("type_id must be positive integer");
  if (!isPositiveInt(market_id)) errs.push("market_id must be positive integer");
  const mt = String(market_type || "").toLowerCase();
  if (!["region","system","station"].includes(mt)) errs.push("market_type must be region|system|station");

  if (!isValidDate(dt)) errs.push("date must be a valid date");

  ["buy_open","buy_close","sell_open","sell_close","daily_high","daily_low","median_buy","median_sell"]
    .forEach((field, idx) => {
      const v = row[4 + idx];
      if (!isBlank(v) && !isNumberish(v)) errs.push(`${field} must be numeric or blank`);
    });

  if (isNumberish(hi) && isNumberish(lo) && Number(hi) < Number(lo)) errs.push("daily_high < daily_low");

  const values = [bo, bc, so, sc, mb, ms].filter(isNumberish).map(Number);
  if (isNumberish(hi) && isNumberish(lo) && values.length) {
    const outOfRange = values.some(v => v > Number(hi) || v < Number(lo));
    if (outOfRange) warns.push("one or more price fields outside [daily_low, daily_high]");
  }

  return { error: errs, warn: warns };
}

/** ---------- Cross-table key collection ---------- **/
function collectKeys(sheet, expectedHeaders){
  const lastRow = sheet.getLastRow();
  const set = new Set();
  if (lastRow <= 1) return set;
  const data = sheet.getRange(2,1,lastRow-1,expectedHeaders.length).getValues();
  const idxType = expectedHeaders.indexOf("type_id");
  const idxMarket = expectedHeaders.indexOf("market_id");
  const idxTypeStr = expectedHeaders.indexOf("market_type");
  for (const r of data) {
    const k = `${r[idxType]}|${r[idxMarket]}|${String(r[idxTypeStr] || "").toLowerCase()}`;
    set.add(k);
  }
  return set;
}

/** ---------- Small utilities ---------- **/
function isNumberish(v){ return v !== "" && v !== null && !isNaN(Number(v)); }
function isPositiveInt(v){ return isNumberish(v) && Number.isInteger(Number(v)) && Number(v) > 0; }
function isBlank(v){ return v === "" || v === null; }
function isValidDate(v){
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) return true;
  const t = (typeof v === "string") ? Date.parse(v) : NaN;
  return !isNaN(t);
}
function safePreview(row){ return row.slice(0, 6); }

/** ================== Quick Start ================== **/
function QuickWins_SmokeTest() {
  Logger.log("Running QuickWins_SmokeTest…");
  sanityCheck();

  const ss = SpreadsheetApp.getActive();
  const mp = ss.getSheetByName("Market Prices");
  const rows = readUsedRange(mp);
  Logger.log(`Market Prices rows read: ${rows.length}`);

  timeIt("NoOp-100k", () => { for (let i=0;i<100000;i++){} });

  cacheSetAll({ foo: "123", bar: "xyz" }, 60);
  const got = cacheGetAll(["foo","bar","baz"]);
  Logger.log(`Cache foo=${got.foo} bar=${got.bar} baz=${got.baz}`);

  protectHeaders("Market Prices");
  Logger.log("QuickWins_SmokeTest completed.");
}

function TableIntegrity_Run() {
  // Optional: bump verbosity temporarily
  if (typeof setLogLevel === "function") setLogLevel("INFO");
  return TableIntegrity_Check();
}

function Debug_Gatekeeper() {
  const cfg = getConfig();             // your config fetch
  const tz  = _projectTZ();
  const now = new Date();

  const [oH,oM] = _toHM(cfg.OpenTime);
  const [cH,cM] = _toHM(cfg.CloseTime);
  const DUR = 60; // minutes window for open/close

  // window membership
  const inOpen  = _inWindow_(now, oH, oM, DUR);
  const inClose = _inWindow_(now, cH, cM, DUR);

  // pretty prints
  const openStart = _fmt(tz, _atHM(now, oH, oM));
  const openEnd   = _fmt(tz, _addMin(_atHM(now, oH, oM), DUR));
  const closeStart= _fmt(tz, _atHM(now, cH, cM));
  const closeEnd  = _fmt(tz, _addMin(_atHM(now, cH, cM), DUR));

  Logger.log("=== Gatekeeper Sanity ===");
  Logger.log(`Now:        ${_fmt(tz, now)} (${tz})`);
  Logger.log(`Open win :  ${openStart} → ${openEnd}  (inside=${inOpen})`);
  Logger.log(`Close win:  ${closeStart} → ${closeEnd} (inside=${inClose})`);
  Logger.log(`Status  :   ${inOpen ? "INSIDE OPEN" : inClose ? "INSIDE CLOSE" : "OUTSIDE BOTH"}`);

  const cases = [
    { mode:"auto",  expect: (inOpen || inClose) },
    { mode:"open",  expect: true  },
    { mode:"close", expect: true  }
  ];

  cases.forEach(({mode, expect}) => {
    const r = _determinePhase(cfg, mode, now);
    const pass = (r.allowed === expect);
    Logger.log(
      `Mode=${mode.padEnd(5)} | allowed=${r.allowed} | isOpen=${r.isOpenRun} | isClose=${r.isCloseRun} | EXPECT=${expect} | ${pass?"PASS":"FAIL"}`
    );
  });
}

/* ---------- tiny helpers ---------- */
function _fmt(tz, d){ return Utilities.formatDate(d, tz, "yyyy-MM-dd HH:mm"); }
function _atHM(base, h, m){ const d=new Date(base); d.setHours(h,m,0,0); return d; }
function _addMin(d, min){ return new Date(d.getTime() + min*60000); }

function Debug_Gatekeeper_At(isoLike) {
  const tz  = _projectTZ();
  const mock = new Date(isoLike); // e.g., "2025-08-18T11:15:00"
  Logger.log(`--- Simulating now=${_fmt(tz, mock)} ---`);
  const cfg = getConfig();
  const [oH,oM] = _toHM(cfg.OpenTime);
  const [cH,cM] = _toHM(cfg.CloseTime);
  const DUR=60;
  const inOpen  = _inWindow_(mock, oH, oM, DUR);
  const inClose = _inWindow_(mock, cH, cM, DUR);
  Logger.log(`Inside OPEN=${inOpen}, Inside CLOSE=${inClose}`);
  ["auto","open","close"].forEach(mode=>{
    const r = _determinePhase(cfg, mode, mock);
    Logger.log(`Mode=${mode} → allowed=${r.allowed}, isOpen=${r.isOpenRun}, isClose=${r.isCloseRun}`);
  });
}

function historySanityPeek(sampleCount = 3) {
  const cfg = getConfig();
  const PRICES = cfg["MarketPricesSheet"] || "Market Prices";
  const HISTORY = cfg["HistorySheetName"] || "Market History";

  const ss = SpreadsheetApp.getActive();
  const psh = ss.getSheetByName(PRICES);
  const hsh = ss.getSheetByName(HISTORY);
  if (!psh || !hsh) { Logger.log("Missing sheets"); return; }

  // --- resolve headers (Prices) ---
  const ph = psh.getRange(1,1,1,psh.getLastColumn()).getValues()[0].map(x=>String(x).toLowerCase().trim());
  const P_DATE = ph.indexOf("date"), P_TID = ph.indexOf("type_id"), P_MID = ph.indexOf("market_id"), P_MTP = ph.indexOf("market_type");
  const P_MINSELL = ph.indexOf("min_sell"), P_MAXBUY = ph.indexOf("max_buy");
  if ([P_DATE,P_TID,P_MID,P_MTP,P_MINSELL,P_MAXBUY].some(i=>i<0)) { Logger.log("Prices header mismatch"); return; }

  // today's window in Project TZ
  const { start, endNow } = projectDayWindowNow_();

  // --- read Prices today ---
  const pr = psh.getLastRow();
  if (pr < 2) { Logger.log("No prices data"); return; }
  const pvals = psh.getRange(2,1,pr-1,psh.getLastColumn()).getValues();

  const groups = new Map(); // key -> {first,last,hiSell,loBuy}
  for (const r of pvals) {
    const d = r[P_DATE]; if (!(d instanceof Date)) continue;
    if (d < start || d > endNow) continue;
    const key = `${r[P_TID]}|${r[P_MID]}|${r[P_MTP]}`;
    let g = groups.get(key);
    if (!g) { g = { first:null, last:null, hiSell:null, loBuy:null }; groups.set(key, g); }
    const minSell = Number(r[P_MINSELL]); const maxBuy = Number(r[P_MAXBUY]);
    if (!g.first || d < g.first.d) g.first = { d, minSell, maxBuy };
    if (!g.last  || d > g.last.d)  g.last  = { d, minSell, maxBuy };
    if (isFinite(minSell)) g.hiSell = (g.hiSell==null)? minSell : Math.max(g.hiSell, minSell);
    if (isFinite(maxBuy))  g.loBuy  = (g.loBuy ==null)? maxBuy  : Math.min(g.loBuy,  maxBuy);
  }

  // --- summary ---
  Logger.log(`Prices today: keys=${groups.size} window=[${start}]..[${endNow}]`);
  if (!groups.size) return;

  // sample a few keys
  const keys = Array.from(groups.keys()).slice(0, sampleCount);
  for (const k of keys) {
    const [tid, mid, mtp] = k.split("|");
    const g = groups.get(k);
    Logger.log(
      `· ${tid}@${mid}/${mtp} ` +
      `open(buy/sell)=(${g.first?.maxBuy ?? '-'}/${g.first?.minSell ?? '-'}) ` +
      `close(buy/sell)=(${g.last?.maxBuy ?? '-'}/${g.last?.minSell ?? '-'}) ` +
      `dayHi=${g.hiSell ?? '-'} dayLo=${g.loBuy ?? '-'}`
    );
  }

  // --- history today count (quick sanity) ---
  const hh = hsh.getRange(1,1,1,hsh.getLastColumn()).getValues()[0].map(x=>String(x).toLowerCase().trim());
  const H_DATE = hh.indexOf("date");
  let todayCount = 0;
  if (hsh.getLastRow() > 1 && H_DATE >= 0) {
    const hvals = hsh.getRange(2,1,hsh.getLastRow()-1,hsh.getLastColumn()).getValues();
    const dayOnly = dateOnlyProjectTZ_(start);
    for (const r of hvals) if (isSameProjectDay_(r[H_DATE], dayOnly)) todayCount++;
  }
  Logger.log(`History today: rows=${todayCount}`);
}


/******************************
 * _determinePhase() Test Harness
 * - Assumes _determinePhase, _inWindow_, _toHM already exist.
 * - Config times are LOCAL (project timezone).
 ******************************/

/** Helper: build a Date for “today” at local h:m:s */
function _localToday(h, m, s=0, ms=0) {
  const d = new Date();
  d.setHours(h, m, s, ms);
  return d;
}

/** Tiny assert */
function _assert(name, cond, details="") {
  if (!cond) {
    console.error(`❌ FAIL: ${name} ${details ? " → " + details : ""}`);
    return false;
  }
  console.log(`✅ PASS: ${name}`);
  return true;
}

/** Compare the phase result to expected flags */
function _expectPhase(name, res, exp) {
  const ok =
    res.allowed   === exp.allowed &&
    res.isOpenRun === exp.isOpenRun &&
    res.isCloseRun=== exp.isCloseRun;
  return _assert(
    name,
    ok,
    `got {allowed:${res.allowed}, open:${res.isOpenRun}, close:${res.isCloseRun}} ` +
    `exp {allowed:${exp.allowed}, open:${exp.isOpenRun}, close:${exp.isCloseRun}}`
  );
}

/** Main: run a suite of time-window tests */
function test_determinePhase() {
  const config = {
    OpenTime:  "11:00",  // LOCAL
    CloseTime: "18:00",  // LOCAL
  };

  let fails = 0;
  const now_10_59 = _localToday(10,59,59);
  const now_11_00 = _localToday(11,0,0);
  const now_11_30 = _localToday(11,30,0);
  const now_11_59 = _localToday(11,59,59);
  const now_12_00 = _localToday(12,0,0);
  const now_13_00 = _localToday(13,0,0);
  const now_18_30 = _localToday(18,30,0);

  // --- OPEN window tests (duration 60, end is exclusive) ---
  fails += !_expectPhase("Open: just before window (10:59:59)",
    _determinePhase(config, "auto", now_10_59),
    {allowed:false, isOpenRun:false, isCloseRun:false});

  fails += !_expectPhase("Open: at window start (11:00:00)",
    _determinePhase(config, "auto", now_11_00),
    {allowed:true, isOpenRun:true, isCloseRun:false});

  fails += !_expectPhase("Open: mid window (11:30)",
    _determinePhase(config, "auto", now_11_30),
    {allowed:true, isOpenRun:true, isCloseRun:false});

  fails += !_expectPhase("Open: last second inside (11:59:59)",
    _determinePhase(config, "auto", now_11_59),
    {allowed:true, isOpenRun:true, isCloseRun:false});

  fails += !_expectPhase("Open: exclusive end (12:00:00) → outside",
    _determinePhase(config, "auto", now_12_00),
    {allowed:false, isOpenRun:false, isCloseRun:false});

  // --- CLOSE window test ---
  fails += !_expectPhase("Close: mid window (18:30)",
    _determinePhase(config, "auto", now_18_30),
    {allowed:true, isOpenRun:false, isCloseRun:true});

  // --- Outside both windows ---
  fails += !_expectPhase("Outside both (13:00)",
    _determinePhase(config, "auto", now_13_00),
    {allowed:false, isOpenRun:false, isCloseRun:false});

  // --- Mode overrides ---
  fails += !_expectPhase("Mode override: open",
    _determinePhase(config, "open", now_13_00),
    {allowed:true, isOpenRun:true, isCloseRun:false});

  fails += !_expectPhase("Mode override: close",
    _determinePhase(config, "close", now_11_30),
    {allowed:true, isOpenRun:false, isCloseRun:true});

  // --- Optional: AM/PM parsing robustness (if your _toHM supports it) ---
  const configAmPm = { OpenTime: "11:00 AM", CloseTime: "6:00 PM" };
  fails += !_expectPhase("AM/PM parsing: open mid (11:30)",
    _determinePhase(configAmPm, "auto", now_11_30),
    {allowed:true, isOpenRun:true, isCloseRun:false});
  fails += !_expectPhase("AM/PM parsing: close mid (18:30)",
    _determinePhase(configAmPm, "auto", now_18_30),
    {allowed:true, isOpenRun:false, isCloseRun:true});

  // Summary
  if (fails) {
    console.error(`\n❌ ${fails} test(s) failed in test_determinePhase`);
    throw new Error(`${fails} _determinePhase test(s) failed`);
  } else {
    console.log("\n🎉 All _determinePhase tests passed");
  }
}

/** Quick single-shot to reproduce your 11:53 case */
function test_1153_case() {
  const config = { OpenTime: "11:00", CloseTime: "18:00" };
  const now = _localToday(11,53,0);
  const res = _determinePhase(config, "auto", now);
  console.log("TZ:", Session.getScriptTimeZone(), "now:", now.toLocaleString());
  console.log("Result:", JSON.stringify(res));
  return res;
}