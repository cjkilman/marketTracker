/** ===== Deep Integrity (heavy – use when needed) ===== **/
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

function _id_checkMarketPrices(log){
  const EXPECT = ["type_id","market_id","market_type","max_buy","min_sell","date"];
  const sh = SpreadsheetApp.getActive().getSheetByName("Market Prices");
  if (!sh) { _id_err(log,"Missing sheet","Market Prices"); return false; }

  const header = sh.getRange(1,1,1,EXPECT.length).getValues()[0];
  if (header.join("|") !== EXPECT.join("|")) { _id_err(log,"Header mismatch","Market Prices"); return false; }

  const last = sh.getLastRow();
  if (last <= 1) { _id_info(log,"No data rows","Market Prices"); return true; }

  const data = sh.getRange(2,1,last-1,EXPECT.length).getValues();
  let errors = 0, warns = 0;
  for (let i=0;i<data.length;i++){
    const [type_id, market_id, market_type, max_buy, min_sell, dt] = data[i];
    if (!_id_posInt(type_id))  errors += _id_row(log,"type_id must be positive integer","Market Prices",i+2);
    if (!_id_posInt(market_id))errors += _id_row(log,"market_id must be positive integer","Market Prices",i+2);
    const mt = String(market_type||"").toLowerCase();
    if (!["region","system","station"].includes(mt)) errors += _id_row(log,"market_type must be region|system|station","Market Prices",i+2);
    if (!_id_num(max_buy))     errors += _id_row(log,"max_buy must be numeric","Market Prices",i+2);
    if (!_id_num(min_sell))    errors += _id_row(log,"min_sell must be numeric","Market Prices",i+2);
    if (_id_num(max_buy) && _id_num(min_sell) && Number(max_buy) > Number(min_sell))
      warns += _id_warn(log,"max_buy > min_sell (spread inverted?)","Market Prices",i+2);
    if (!_id_date(dt))         errors += _id_row(log,"date must be a valid date","Market Prices",i+2);
  }
  _id_report(log,"Market Prices",data.length,errors,warns);
  return errors === 0;
}

function _id_checkMarketHistory(log){
  const EXPECT = ["type_id","market_id","market_type","date","buy_open","buy_close","sell_open","sell_close","daily_high","daily_low","median_buy","median_sell"];
  const sh = SpreadsheetApp.getActive().getSheetByName("Market History");
  if (!sh) { _id_err(log,"Missing sheet","Market History"); return false; }

  const header = sh.getRange(1,1,1,EXPECT.length).getValues()[0];
  if (header.join("|") !== EXPECT.join("|")) { _id_err(log,"Header mismatch","Market History"); return false; }

  const last = sh.getLastRow();
  if (last <= 1) { _id_info(log,"No data rows","Market History"); return true; }

  const data = sh.getRange(2,1,last-1,EXPECT.length).getValues();
  let errors = 0, warns = 0;
  for (let i=0;i<data.length;i++){
    const [type_id, market_id, market_type, dt, bo, bc, so, sc, hi, lo, mb, ms] = data[i];
    if (!_id_posInt(type_id))  errors += _id_row(log,"type_id must be positive integer","Market History",i+2);
    if (!_id_posInt(market_id))errors += _id_row(log,"market_id must be positive integer","Market History",i+2);
    const mt = String(market_type||"").toLowerCase();
    if (!["region","system","station"].includes(mt)) errors += _id_row(log,"market_type must be region|system|station","Market History",i+2);
    if (!_id_date(dt))         errors += _id_row(log,"date must be a valid date","Market History",i+2);

    const fields = [["buy_open",bo],["buy_close",bc],["sell_open",so],["sell_close",sc],
                    ["daily_high",hi],["daily_low",lo],["median_buy",mb],["median_sell",ms]];
    for (const [name,val] of fields){
      if (!_id_blank(val) && !_id_num(val)) errors += _id_row(log,`${name} must be numeric or blank`,"Market History",i+2);
    }
    if (_id_num(hi) && _id_num(lo) && Number(hi) < Number(lo))
      errors += _id_row(log,"daily_high < daily_low","Market History",i+2);

    const values = [bo,bc,so,sc,mb,ms].filter(_id_num).map(Number);
    if (_id_num(hi) && _id_num(lo) && values.length){
      const out = values.some(v => v > Number(hi) || v < Number(lo));
      if (out) warns += _id_warn(log,"one or more price fields outside [daily_low, daily_high]","Market History",i+2);
    }
  }
  _id_report(log,"Market History",data.length,errors,warns);
  return errors === 0;
}

function _id_crossKeys(log){
  const ss = SpreadsheetApp.getActive();
  const mp = ss.getSheetByName("Market Prices");
  const mh = ss.getSheetByName("Market History");
  if (!mp || !mh) return false;

  const keysMP = _id_keys(mp, ["type_id","market_id","market_type"]);
  const keysMH = _id_keys(mh, ["type_id","market_id","market_type"]);

  let missing = 0;
  for (const k of keysMH) if (!keysMP.has(k)) missing++;

  if (missing > 0) _id_warn(log,"Some history keys not present in Market Prices","Cross",null,{missing});
  else _id_info(log,"Cross-keys OK","Cross");
  return true; // warn-only; don’t fail deep check on this
}

/** ——— Utilities ——— */
function _id_report(log,name,rows,errors, warns){
  const msg = errors===0 && warns===0 ? `${name}: OK`
            : errors===0 ? `${name}: OK with warnings (${warns})`
            : `${name}: FAILED (errors=${errors}, warns=${warns})`;
  if (typeof Log !== "undefined") {
    if (errors===0) Log.for("Integrity").info(msg, { rows }); else Log.for("Integrity").error(msg, { rows });
  } else {
    Logger.log(msg);
  }
}
function _id_row(log,msg,sheet,row){ _id_err(log,`${msg}`,sheet,row); return 1; }
function _id_err(log,msg,sheet,row){ 
  if (log) log.error(msg, { sheet, row }); else Logger.log(`[ERROR] ${sheet} row ${row||"-"}: ${msg}`); 
}
function _id_warn(log,msg,sheet,row,meta){
  if (log) log.warn(msg, Object.assign({ sheet, row }, meta||{})); else Logger.log(`[WARN] ${sheet} row ${row||"-"}: ${msg}`);
  return 1;
}
function _id_info(log,msg,sheet){ if (log) log.info(msg, { sheet }); else Logger.log(`[INFO] ${sheet}: ${msg}`); }

function _id_num(v){ return v !== "" && v !== null && !isNaN(Number(v)); }
function _id_blank(v){ return v === "" || v === null; }
function _id_posInt(v){ return _id_num(v) && Number.isInteger(Number(v)) && Number(v) > 0; }
function _id_date(v){
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) return true;
  const t = (typeof v === "string") ? Date.parse(v) : NaN;
  return !isNaN(t);
}
function _id_keys(sheet, cols){
  // assumes headers in row 1
  const header = sheet.getRange(1,1,1,Math.max(3, sheet.getLastColumn())).getValues()[0];
  const idx = cols.map(c => header.indexOf(c));
  const last = sheet.getLastRow();
  const set = new Set();
  if (last <= 1) return set;
  const data = sheet.getRange(2,1,last-1,Math.max(...idx)+1).getValues();
  for (const r of data){
    const k = `${r[idx[0]]}|${r[idx[1]]}|${String(r[idx[2]]||"").toLowerCase()}`;
    set.add(k);
  }
  return set;
}