/***********************
 * HistoryManager.gs (CLEAN)
 * - Builds/updates daily history rows from "Market Prices"
 *   headers: [type_id, market_id, market_type, max_buy, min_sell, date]
 * - Day High  = max(min_sell) over window
 * - Day Low   = min(max_buy)  over window
 * - Open/Close snapshots from collector (fill-if-blank at open, set at close)
 * - Retention via History.RetentionDays (default 365)
 * - Prod sheet: "Market History" | Test sheet: "History (TEST)"
 * - Anchored to PROJECT-LOCAL open/close times; batch read/write; LoggerEx
 ***********************/

const HISTORY_PROD_SHEET = "Market History";
const HISTORY_TEST_SHEET = "History (TEST)";

/** Prod entrypoint — safe for triggers */
function updateHistory() {
  return _updateHistoryCore({ auto: true, mode: "auto", testMode: false });
}

/** Debug entrypoint — manual run into test sheet */
function updateHistoryTest() {
  return _updateHistoryCore({ auto: false, mode: "open", testMode: true });
}


function seedHistoryOpenOnce(){
  return _updateHistoryCore({ auto:false, mode:"open", testMode:false, skipPrune:true });
}

/**
 * Core update history logic.
 * @param {{auto:boolean,mode:("open"|"close"|"auto"),testMode:boolean,skipPrune?:boolean}} opts
 */
function _updateHistoryCore(opts) {
  opts = opts || { auto: true, mode: "auto", testMode: false };
  const L = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('History') : {log:Logger.log, warn:Logger.log, error:Logger.log, debug:Logger.log, startTimer:()=>({stamp:Logger.log})});
  const T = L.startTimer('updateHistory');

  const cfg = getConfig();
  const now = new Date();

  // Phase (LOCAL clock)
  const phase = (opts.auto ? _determinePhaseLocal(now, cfg) : (opts.mode || "open"));
  const isClose = phase === "close";
  const isOpen  = phase === "open";
  const sheetName = opts.testMode ? HISTORY_TEST_SHEET : HISTORY_PROD_SHEET;

  // Safety: on close, abort if no sheet/data yet (sheet created at open)
  if (isClose && AbortCloseIfNoHistoryYet(sheetName)) {
    return L.warn("Abort: close-run with no history sheet yet.");
  }

  const ss = SpreadsheetApp.getActive();
  const pricesSheet = ss.getSheetByName("Market Prices");
  if (!pricesSheet) return L.error("No 'Market Prices' sheet found – nothing to do.");

  const hist = getOrCreateSheet(ss, sheetName, [
    "type_id","market_id","market_type","date",
    "buy_open","buy_close","sell_open","sell_close",
    "daily_high","daily_low","median_buy","median_sell"
  ]);

  // Config (LOCAL times); accepts HH:MM strings or Date (Sheets time-only)
  const openVal   = (cfg && (cfg["OpenTime"]))  || "11:00";
  const closeVal  = (cfg && (cfg["CloseTime"])) || "18:00";
  const sliceHours= parseInt((cfg && (cfg["SliceHours"])) || 3, 10); // comfort default

  const openLocal  = todayAtHM_Local(openVal);
  const closeLocal = todayAtHM_Local(closeVal);

  // Window selection (LOCAL day anchored)
  let rangeStart, rangeEnd;
  if (isClose) {
    rangeStart = openLocal;  rangeEnd = closeLocal;   // full-day sweep
  } else {
    const sliceStart = new Date(now.getTime() - sliceHours*3600*1000);
    rangeStart = new Date(Math.max(openLocal.getTime(), sliceStart.getTime()));
    rangeEnd   = now;
  }
  L.log("phase=%s open=%s close=%s", phase, openLocal, closeLocal);
  L.debug("window %s → %s", rangeStart, rangeEnd);

  // Read only needed columns within window
  const needCols = ["type_id","market_id","market_type","max_buy","min_sell","date"];
  const recent   = _readPricesInWindow(pricesSheet, rangeStart, rangeEnd, needCols);
  const pHead    = recent.headerIdx; const pData = recent.rows;
  T.stamp("rows in window=" + pData.length);
  if (!pData.length) return L.warn("No market price rows in selected window.");

  // Aggregate by key (type|market|kind|date)
  const byKey = new Map();
  for (const r of pData) {
    const type_id = r[pHead.type_id], market_id = r[pHead.market_id], market_type = r[pHead.market_type];
    const max_buy = Number(r[pHead.max_buy])||0, min_sell = Number(r[pHead.min_sell])||0;
    const d = new Date(r[pHead.date]);
    const dateOnly = _dateYMD(d);
    const k = `${type_id}|${market_id}|${market_type}|${dateOnly}`;
    let o = byKey.get(k);
    if (!o) {
      o = { type_id, market_id, market_type, dateOnly,
            // for close run:
            buys: [], sells: [],
            // for light slices:
            daily_low: null, daily_high: null,
            earliestTS: null, open_buy_candidate: null, open_sell_candidate: null
          };
      byKey.set(k, o);
    }

    if (isClose || !lightSlices) {
      // heavy path (close pass or you’ve disabled LightSlices)
      if (max_buy>0)  o.buys.push(max_buy);
      if (min_sell>0) o.sells.push(min_sell);
    } else {
      // light path (open/midday): track only min/max and earliest sample
      if (max_buy>0)  o.daily_low  = (o.daily_low==null ? max_buy  : Math.min(o.daily_low,  max_buy));
      if (min_sell>0) o.daily_high = (o.daily_high==null? min_sell : Math.max(o.daily_high, min_sell));
      if (!o.earliestTS || d < o.earliestTS) {
        o.earliestTS = d;
        if (max_buy>0)  o.open_buy_candidate  = max_buy;
        if (min_sell>0) o.open_sell_candidate = min_sell;
      }
    }
  }

    for (const o of byKey.values()) {
      if (isClose || !lightSlices) {
        // heavy path: final medians + final H/L
        o.buys.sort((a,b)=>a-b);
        o.sells.sort((a,b)=>a-b);
        o.median_buy  = _median(o.buys);
        o.median_sell = _median(o.sells);
        o.daily_low   = o.buys.length  ? Math.min(...o.buys)  : null;
        o.daily_high  = o.sells.length ? Math.max(...o.sells) : null;
      }
    }

  // Index existing history
  const hRows = _readTable(hist); const hHead = hRows.headerIdx; const hData = hRows.rows;
  const index = new Map();
  for (let i=0;i<hData.length;i++){
    const r=hData[i];
    const k=`${r[hHead.type_id]}|${r[hHead.market_id]}|${r[hHead.market_type]}|${_dateYMD(new Date(r[hHead.date]))}`;
    if(!index.has(k)) index.set(k, i+2);
  }

  // Prepare updates
  const updates=[];
  const setCell=(row,col,val)=>{ if(val!==null&&val!==undefined&&val!=='') row[hHead[col]] = val; return row; };

  for (const o of byKey.values()) {
    const k = `${o.type_id}|${o.market_id}|${o.market_type}|${o.dateOnly}`;
    let rowIndex=index.get(k), rowVals;
    if (rowIndex) {
      rowVals = hist.getRange(rowIndex,1,1,hRows.header.length).getValues()[0];
    } else {
      rowVals = new Array(hRows.header.length).fill("");
      rowVals[hHead.type_id]     = o.type_id;
      rowVals[hHead.market_id]   = o.market_id;
      rowVals[hHead.market_type] = o.market_type;
      rowVals[hHead.date]        = new Date(o.dateOnly);
      rowIndex = hist.getLastRow() + 1;
    }

    // Always refresh daily aggregates; write medians at close only (keep daytime light)
    // Always refresh daily aggregates
    setCell(rowVals, "daily_high", o.daily_high);
    setCell(rowVals, "daily_low",  o.daily_low);

    // Snapshots
    if (isOpen) {
      if (!rowVals[hHead.buy_open]  && (o.open_buy_candidate  != null))
        setCell(rowVals, "buy_open",  o.open_buy_candidate);
      if (!rowVals[hHead.sell_open] && (o.open_sell_candidate != null))
        setCell(rowVals, "sell_open", o.open_sell_candidate);
    }

    if (isClose) {
      setCell(rowVals, "median_buy",  o.median_buy);
      setCell(rowVals, "median_sell", o.median_sell);
      setCell(rowVals, "buy_close",   o.median_buy);
      setCell(rowVals, "sell_close",  o.median_sell);
    }

    updates.push({rowIndex, rowVals});
  }

  // Batch apply (contiguous coalescing + cell-chunk guard)
  if (updates.length){
    updates.sort((a,b)=>a.rowIndex-b.rowIndex);
    const headerLen=hRows.header.length;
    const chunkCells=parseInt((cfg && (cfg["History.ChunkSize"]))||5000,10);

    let blockStart=null, blockVals=[];
    const flush=()=>{
      if(!blockVals.length) return;
      const maxRowsPer=Math.max(1,Math.floor(chunkCells/headerLen));
      for(let i=0;i<blockVals.length;i+=maxRowsPer){
        const slice=blockVals.slice(i,i+maxRowsPer);
        hist.getRange(blockStart+i,1,slice.length,headerLen).setValues(slice);
      }
      blockStart=null; blockVals=[];
    };

    for(const u of updates){
      if(blockStart===null){ blockStart=u.rowIndex; blockVals=[u.rowVals]; }
      else if(u.rowIndex===blockStart+blockVals.length){ blockVals.push(u.rowVals); }
      else { flush(); blockStart=u.rowIndex; blockVals=[u.rowVals]; }
    }
    flush();
  }
  T.stamp("writes done rows="+updates.length);

  // Prune (skip in daytime if requested)
  const retentionDays = parseInt((cfg && (cfg["History.RetentionDays"] || cfg["RetentionDays"])) || 365, 10);
  if (!opts.skipPrune) _pruneHistoryByAge(hist, retentionDays);

  L.log("done phase=%s, rows=%s", phase, updates.length);
}

/** Abort close-run if no sheet/data yet */
function AbortCloseIfNoHistoryYet(sheetName){
  const sh=SpreadsheetApp.getActive().getSheetByName(sheetName);
  if(!sh) return true; return !(sh.getLastRow()>1);
}

/***********************
 * Helpers (time, math, tables)
 ***********************/


function _determinePhaseLocal(nowLocal, cfg) {
  try {
    const o = normalizeHM((cfg && cfg["OpenTime"])  || "11:00");
    const c = normalizeHM((cfg && cfg["CloseTime"]) || "18:00");
    const hm = nowLocal.getHours()*60 + nowLocal.getMinutes();
    const oMin = o.h*60 + o.m;
    const cMin = c.h*60 + c.m;

    if (oMin <= cMin) {
      // Same-day session (e.g., 11:00 → 18:00)
      return (hm >= oMin && hm < cMin) ? "open" : "close";
    } else {
      // Overnight session (e.g., 18:00 → 11:00 next day)
      // Open if hm >= open OR hm < close
      return (hm >= oMin || hm < cMin) ? "open" : "close";
    }
  } catch (e) {
    return "open"; // safe default
  }
}

function _parseHM(s){ var m=String(s||"").match(/^(\d{1,2}):(\d{2})$/); if(!m) throw new Error('Bad HM:'+s); return {h:parseInt(m[1],10), m:parseInt(m[2],10)}; }
function _dateYMD(d){ var y=d.getFullYear(); var m=(d.getMonth()+1).toString().padStart(2,'0'); var day=d.getDate().toString().padStart(2,'0'); return y+"-"+m+"-"+day; }
function _median(arr){ if(!arr||!arr.length) return null; var a=arr.slice().sort(function(x,y){return x-y}); var n=a.length; return (n%2)?a[(n-1)/2]:(a[n/2-1]+a[n/2])/2; }

function _readTable(sh){ var lastR=sh.getLastRow(), lastC=sh.getLastColumn(); if(!lastR||!lastC) return {header:[],headerIdx:{},rows:[]}; var rawHeader=sh.getRange(1,1,1,lastC).getValues()[0].map(String); var hLen=rawHeader.length; while(hLen>1 && (!rawHeader[hLen-1]||rawHeader[hLen-1]=='')) hLen--; var header=rawHeader.slice(0,hLen); var headerIdx=Object.fromEntries(header.map(function(h,i){return [h,i]})); var rows=(lastR>1)?sh.getRange(2,1,lastR-1,hLen).getValues():[]; return {header:header, headerIdx:headerIdx, rows:rows}; }

/** Read only needed columns within a LOCAL time window */
function _readPricesInWindow(pricesSheet, start, end, needCols){
  var lastR=pricesSheet.getLastRow(), lastC=pricesSheet.getLastColumn();
  if(lastR<2||!lastC) return { header:needCols, headerIdx:Object.fromEntries(needCols.map(function(n,i){return [n,i]})), rows:[] };
  var hdr=pricesSheet.getRange(1,1,1,lastC).getValues()[0].map(String);
  var idx1=Object.fromEntries(hdr.map(function(h,i){return [h,i+1]}));
  for (var i=0;i<needCols.length;i++){ var n=needCols[i]; if(!idx1[n]) throw new Error('Missing column in Market Prices: '+n); }

  var dcol=idx1['date'];
  var dates=pricesSheet.getRange(2,dcol,lastR-1,1).getValues();
  var first=-1,last=-1;
  for (var r=0;r<dates.length;r++){
    var cell=dates[r][0]; var d=(cell instanceof Date)?cell:new Date(cell);
    if (d>=start && d<=end){ if(first===-1) first=r; last=r; }
  }
  if(first===-1) return { header:needCols, headerIdx:Object.fromEntries(needCols.map(function(n,i){return [n,i]})), rows:[] };

  var cols=needCols.map(function(n){return idx1[n]}).sort(function(a,b){return a-b});
  var c1=cols[0], c2=cols[cols.length-1]; var height=last-first+1;
  var block=pricesSheet.getRange(first+2, c1, height, c2-c1+1).getValues();
  var rel={}; for (var j=0;j<needCols.length;j++){ rel[needCols[j]] = idx1[needCols[j]] - c1; }
  var rows=block.map(function(rw){ return [ rw[rel.type_id], rw[rel.market_id], rw[rel.market_type], rw[rel.max_buy], rw[rel.min_sell], rw[rel.date] ]; });
  var headerIdx=Object.fromEntries(needCols.map(function(n,i){return [n,i]}));
  return { header:needCols, headerIdx:headerIdx, rows:rows };
}


/** Prune old rows */
function _pruneHistoryByAge(sh,days){ if(!days||days<=0) return; var t=_readTable(sh); var idx=t.headerIdx; var now=Date.now(); var keep=[t.header]; for (var i=0;i<t.rows.length;i++){ var r=t.rows[i]; var d=new Date(r[idx.date]); if(!isFinite(d.getTime())) continue; var age=(now-d.getTime())/86400000; if(age<=days) keep.push(r); } sh.clearContents(); sh.getRange(1,1,keep.length, keep[0].length).setValues(keep); }
