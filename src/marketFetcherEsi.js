/***** marketFetcherESI.gs  —  MAIN PROJECT FILE ********************************
 * Contains BOTH sides:
 *   1) ENGINE: pulls ESI region history via GESI → caches & publishes slim table
 *   2) CLIENT: marketStatDataCache(...)  (cache-only; reads CacheService via your
 *      _getCachedFuz and _normalizeOrder from marketStatData)
 *
 * SLIM IS LAW:
 *   - Engine publishes named range: MarketResultESI_Region (type_id, region_id, vol30, ts, status)
 *   - Clients import that single range; all other math stays local (Fuz books + DoB).
 *******************************************************************************/

/* ============================ CONFIG: ENGINE ============================ */
// Sources can be a Named Range OR "Sheet!A1" range
const ITEMS_SOURCE        = "'Item List Back End'!A2:A";   // type_id list
const REGIONS_SOURCE      = "'Market Settings'!F3:F";      // region_id list

// Engine sheets & export name
const SHEET_REGION_CACHE  = 'Cache_Market_ESI_Region';
const SHEET_PUBLISH       = 'Publish_ESI_Region';
const NR_MARKET_RESULT    = 'MarketResultESI_Region';

// Worker cadence & batching
const WORKER_EVERY_MIN    = 10;     // chunk worker interval (minutes)
const BATCH_SIZE          = 25;    // items per chunk per region
const CALL_PACE_MS        = 500;   // sleep between GESI calls (ms)

// Daily kickoff time (EVE DT ~11:00 UTC)
const DAILY_UTC_HOUR      = 11;
const DAILY_UTC_MIN       = 20;

// Engine headers
const HEADERS_REGION      = ['type_id','region_id','volume30_region','velocity_region','last_updated','status'];
const HEADERS_PUBLISH     = ['type_id','region_id','volume30_region','last_updated','status'];


/** ---------- Rate limit + backoff (lightweight) ---------- **/
const PACER_PROP_GAP  = 'esi_gap_ms';
const PACER_PROP_LAST = 'esi_last_ms';
const PACER_MIN_MS    = 400;   // start ~2.5 rps
const PACER_MAX_MS    = 5000;  // cap 5s between calls

function isRateLimitedError_(e){
  const s = String(e && e.message || e || '');
  return /bandwidth quota exceeded|error limit|too many|429|420/i.test(s);
}

function rateLimitPacer_(){
  const p = PropertiesService.getScriptProperties();
  const gap  = Math.max(PACER_MIN_MS, Math.min(PACER_MAX_MS, Number(p.getProperty(PACER_PROP_GAP)) || PACER_MIN_MS));
  const last = Number(p.getProperty(PACER_PROP_LAST)) || 0;
  const now  = Date.now();
  const wait = Math.max(0, last + gap - now);
  if (wait > 0) Utilities.sleep(wait);
  p.setProperty(PACER_PROP_LAST, String(Date.now()));
  return gap;
}

function bumpPacer_(){
  const p = PropertiesService.getScriptProperties();
  const cur  = Math.max(PACER_MIN_MS, Number(p.getProperty(PACER_PROP_GAP)) || PACER_MIN_MS);
  const next = Math.min(PACER_MAX_MS, Math.ceil(cur * 1.5));
  p.setProperty(PACER_PROP_GAP, String(next));
  return next;
}

function relaxPacer_(){
  const p = PropertiesService.getScriptProperties();
  const cur  = Math.max(PACER_MIN_MS, Number(p.getProperty(PACER_PROP_GAP)) || PACER_MIN_MS);
  const next = Math.max(PACER_MIN_MS, Math.floor(cur * 0.8));
  p.setProperty(PACER_PROP_GAP, String(next));
  return next;
}

function withRetries_(fn, tries=3, baseDelay=800){
  let attempt = 0;
  while (true){
    try {
      rateLimitPacer_();              // global pacing
      return fn();                    // do the call
    } catch(e){
      if (isRateLimitedError_(e) && attempt < tries-1){
        const gap = bumpPacer_();
        const jitter = Math.floor(Math.random()*300);
        const backoff = baseDelay * Math.pow(2, attempt) + jitter;
        (LoggerEx?.warn || console.warn)(`Rate-limited: gap=${gap}ms, backoff=${backoff}ms, try=${attempt+1}/${tries}`);
        Utilities.sleep(backoff);
        attempt++;
        continue;
      }
      throw e; // non-rate error or out of tries
    }
  }
}

function fetchVol30One_(typeId, regionId){
  try {
    const hist = withRetries_(
      () => GESI.getClient().setFunction('markets_region_history')
             .executeRaw({ type_id: typeId, region_id: regionId }),
      3, 800
    ); // → [{date:"YYYY-MM-DD", volume:...}, ...]

    if (!Array.isArray(hist) || !hist.length) return { status:'NOT_FOUND', vol30:0, vel:0 };
    const vol30 = sumLastNDaysObj_(hist, 30);
    return { status: (vol30>0 ? 'OK' : 'NO_DATA_30D'), vol30, vel: vol30/30 };
  } catch(e){
    if (isRateLimitedError_(e)) return { status:'ERR_RATE' };     // don’t write a row; retry later
    (LoggerEx?.error || console.error)('ESI error:', e?.message, e?.stack);
    return { status:'ERR_ESI', vol30:0, vel:0 };
  }
}



/* ============================ CONFIG: CLIENT ============================ */
/* marketStatDataCache uses your CacheService via _getCachedFuz and _normalizeOrder.
 * It needs NO sheet/range config. In your Sheets formulas, add your own FUZ_TICK heartbeat.
 */

function testESIHistory() {
  const rid = 10000043; // Domain
  const tid = 34;       // Tritanium
  const res = GESI.markets_region_history(tid,  rid );
  console.log('history rows:', Array.isArray(res) ? res.length : res);
}

/* =============================== MENUS ================================= */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Engine')
    .addItem('Setup Engine (first run)', 'setupEngine')
    .addSeparator()
    .addItem('Install daily kickoff', 'installDailyKickoff')
    .addItem('Kickoff now (mark STALE & start worker)', 'kickoffMarketHistoryRefresh')
    .addItem('Stop worker', 'stopWorker_')
    .addSeparator()
    .addItem('Publish now', 'publishMarketResultESIRegion')
    .addItem('Debug sources (log counts)', 'debugListSources')
    .addToUi();
}


/* =============================== SETUP ================================= */
function setupEngine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCacheMarketESIRegion_(ss);
  ensurePublishSheet_(ss);
  publishMarketResultESIRegion();
  SpreadsheetApp.getUi().alert('Engine setup complete ✅');
}

function ensureCacheMarketESIRegion_(ss) {
  let sh = ss.getSheetByName(SHEET_REGION_CACHE);
  if (!sh) sh = ss.insertSheet(SHEET_REGION_CACHE);
  const hdr = sh.getRange(1,1,1,HEADERS_REGION.length).getValues()[0];
  const needs = !hdr[0] || HEADERS_REGION.some((h,i)=>hdr[i]!==h);
  if (needs) {
    sh.clear();
    sh.getRange(1,1,1,HEADERS_REGION.length).setValues([HEADERS_REGION]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensurePublishSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_PUBLISH);
  if (!sh) sh = ss.insertSheet(SHEET_PUBLISH);
  const hdr = sh.getRange(1,1,1,HEADERS_PUBLISH.length).getValues()[0];
  const needs = !hdr[0] || HEADERS_PUBLISH.some((h,i)=>hdr[i]!==h);
  if (needs) {
    sh.clear();
    sh.getRange(1,1,1,HEADERS_PUBLISH.length).setValues([HEADERS_PUBLISH]);
    sh.setFrozenRows(1);
  }
  // heartbeat (outside named range)
  sh.getRange('G1').setValue('published_at');
  sh.getRange('H1').setNumberFormat('yyyy-mm-dd"T"hh:mm:ss"Z"');
  SpreadsheetApp.getActive().setNamedRange(NR_MARKET_RESULT, sh.getRange('A:E'));
}


/* ============================ SCHEDULING =============================== */
function installDailyKickoff() {
  // clear old triggers for a clean slate
  ScriptApp.getProjectTriggers().forEach(t => {
    const h = t.getHandlerFunction();
    if (h === 'kickoffMarketHistoryRefresh' || h === 'marketFetchChunk') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('kickoffMarketHistoryRefresh')
    .timeBased().everyDays(1).atHour(DAILY_UTC_HOUR).nearMinute(DAILY_UTC_MIN)
    .inTimezone('Etc/UTC').create();
  SpreadsheetApp.getUi().alert(`Daily kickoff installed @ ${DAILY_UTC_HOUR}:${DAILY_UTC_MIN} UTC ✅`);
}


/* ========================= KICKOFF → WORKER ============================ */
function kickoffMarketHistoryRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = ss.getSheetByName(SHEET_REGION_CACHE);
  if (!cache) throw new Error('Missing '+SHEET_REGION_CACHE);

  markAllCacheStale_(cache);

  const props = PropertiesService.getScriptProperties();
  props.setProperty('mf_cursor', JSON.stringify({ ri:0, ii:0 }));
  props.setProperty('mf_job_active', '1');
  props.setProperty('mf_job_started', new Date().toISOString());

  publishMarketResultESIRegion(); // clients see stale→fresh tick

  ScriptApp.newTrigger('marketFetchChunk').timeBased().everyMinutes(WORKER_EVERY_MIN).create();
}


/* ===================== WORKER: ESI → CACHE (GESI) ====================== */
function marketFetchChunk() {
  // prevent overlap
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cache = ss.getSheetByName(SHEET_REGION_CACHE);
    if (!cache) return stopWorker_('no cache');

    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('mf_job_active') !== '1') return stopWorker_('no active job');

    const regions = readListFlex_(REGIONS_SOURCE, {
          numeric: true, integer: true, dropZeros: true, validator: isValidRegionId_
        });

    const items = readListFlex_(ITEMS_SOURCE, { numeric: true, integer: true });

    if (!items.length || !regions.length) return stopWorker_('no items/regions');

    const cur = JSON.parse(props.getProperty('mf_cursor') || '{"ri":0,"ii":0}');
    let ri = Math.min(cur.ri||0, regions.length-1);
    let ii = Math.min(cur.ii||0, items.length-1);

    const regionId = regions[ri];
    const end  = Math.min(ii + BATCH_SIZE, items.length);
    const chunk = items.slice(ii, end);

    // --- replacement for the per-item loop ---
const rowsOut = [];
let processed = 0;
let hitRate   = false;

for (const typeId of chunk) {
  const res = fetchVol30One_(typeId, regionId);   // uses retries + pacer

  if (res.status === 'ERR_RATE') {                // hit bandwidth limit → pause chunk
    hitRate = true;
    break;                                        // leave remaining items for next run
  }

  // Write concrete results (OK, NO_DATA_30D, NOT_FOUND, ERR_ESI)
  rowsOut.push([
    typeId,
    regionId,
    res.vol30 || 0,
    res.vel   || 0,
    new Date(),
    res.status
  ]);

  // Optional debug: highlight surprising zeros
  if ((res.vol30|0) === 0 && res.status === 'OK') {
    (LoggerEx?.info || console.info)('zero vol30', {
      typeId, regionId
      // tip: you can log a sample hist row inside fetchVol30One_ if needed
    });
  }

  processed++;
}

// Flush this chunk’s results
if (rowsOut.length) upsertRegionCache_(cache, rowsOut);
publishMarketResultESIRegion();

// Advance cursor only by processed records
ii += processed;
if (ii >= items.length) { ii = 0; ri++; }

// Relax pacer if we completed the chunk without rate limits
if (!hitRate) relaxPacer_();

// (keep your existing "done?" check / cursor persistence below)


    upsertRegionCache_(cache, rowsOut);
    publishMarketResultESIRegion(); // progressive publish

    // advance cursor
    ii = end;
    if (ii >= items.length) { ii = 0; ri++; }

    // done?
    if (ri >= regions.length) {
      props.deleteProperty('mf_job_active');
      props.deleteProperty('mf_cursor');
      return stopWorker_('done');
    }
    props.setProperty('mf_cursor', JSON.stringify({ ri, ii }));

  } finally {
    try { lock.releaseLock(); } catch(e){}
  }
}

function isValidRegionId_(n) {
  n = Math.floor(Number(n));
  return Number.isFinite(n) && n >= 10000000 && n < 20000000;
}


/* ============================ CLIENT SIDE ============================== */
/**
 * marketStatDataCache — cache-only (no network, no tables).
 * Reads Fuz book stats from CacheService via your _getCachedFuz helper.
 * If a key isn’t cached yet, returns "".
 *
 * @param {number|range} type_ids
 * @param {string}       location_type  "region" | "system" | "station"
 * @param {number}       location_id
 * @param {string}       order_type     "sell" | "buy" (default "sell")
 * @param {string}       order_level    "volume" (default); supports min|max|avg|median if your cache holds them
 * @return {any[][]|number|string}      same shape as type_ids
 */
function marketStatDataCache(type_ids, location_type, location_id, order_type, order_level) {
  if (type_ids == null) return "";

  const lt = String(location_type || "").toLowerCase();
  if (!["region","system","station"].includes(lt)) {
    return Array.isArray(type_ids) ? type_ids.map(() => [""]) : "";
  }

  const norm = (typeof _normalizeOrder === "function")
    ? _normalizeOrder(order_type, order_level)
    : { type: (String(order_type || "sell").toLowerCase()==="buy"?"buy":"sell"),
        level: String(order_level || "volume").toLowerCase()
      };

  // preserve input shape
  const in2D = Array.isArray(type_ids) ? type_ids : [[type_ids]];
  const rows = in2D.length, cols = in2D[0].length;

  // flatten ids, keep nulls for placeholders
  const flatIds = [];
  for (let r=0; r<rows; r++) for (let c=0; c<cols; c++) {
    const n = Number(in2D[r][c]); flatIds.push(Number.isFinite(n) ? n : null);
  }

  if (typeof _getCachedFuz !== "function") {
    throw new Error("_getCachedFuz not found — ensure marketStatData defines it (CacheService-backed).");
  }

  const uniq = Array.from(new Set(flatIds.filter(n => n != null)));
  const { have } = _getCachedFuz(uniq, Number(location_id), lt); // {have:{[typeId]:row}}

  const pick = (row) => {
    if (!row || !row[norm.type]) return null;
    const v = row[norm.type][norm.level];  // min|max|avg|median|volume
    const num = Number(v);
    return Number.isFinite(num) ? num : null;
  };

  let k = 0;
  const out = Array.from({ length: rows }, () => Array(cols).fill(""));
  for (let r=0; r<rows; r++) for (let c=0; c<cols; c++) {
    const id = flatIds[k++];
    out[r][c] = (id == null) ? "" : (pick(have[id]) ?? "");
  }
  return Array.isArray(type_ids) ? out : out[0][0];
}


/* =============================== HELPERS ================================ */
function publishMarketResultESIRegion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = ss.getSheetByName(SHEET_REGION_CACHE);
  const pub   = ss.getSheetByName(SHEET_PUBLISH);
  if (!cache || !pub) throw new Error('Missing sheets');

  const vals = cache.getDataRange().getValues();
  const out = [HEADERS_PUBLISH];
  if (vals.length > 1) {
    const h  = vals[0];
    const ix = {
      t:h.indexOf('type_id'), r:h.indexOf('region_id'),
      v:h.indexOf('volume30_region'), u:h.indexOf('last_updated'), s:h.indexOf('status')
    };
    for (let i=1; i<vals.length; i++) {
      const row = vals[i];
      if (!row[ix.t]) continue;
      out.push([row[ix.t], row[ix.r], row[ix.v], row[ix.u], row[ix.s]]);
    }
  }
  pub.clearContents();
  pub.getRange(1,1,out.length,HEADERS_PUBLISH.length).setValues(out);
  // heartbeat
  pub.getRange('H1').setNumberFormat('yyyy-mm-dd"T"hh:mm:ss"Z"').setValue(new Date());
  SpreadsheetApp.getActive().setNamedRange(NR_MARKET_RESULT, pub.getRange('A:E'));
}

function markAllCacheStale_(sheet) {

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const h = data[0];
  const iStatus = h.indexOf('status');
  const iUpd    = h.indexOf('last_updated');
  if (iStatus === -1) throw new Error('status col missing');
  sheet.getRange(2, iStatus+1, data.length-1, 1)
       .setValues(Array(data.length-1).fill(['STALE']));
  if (iUpd !== -1) {
    sheet.getRange(2, iUpd+1, data.length-1, 1)
         .setValues(Array(data.length-1).fill([new Date()]));
  }
}

function stopWorker_(reason) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'marketFetchChunk')
    .forEach(t => ScriptApp.deleteTrigger(t));
  try { publishMarketResultESIRegion(); } catch(e){}
  if (reason) console.log('marketFetchChunk: stopped ('+reason+')');
}

// Sum last 30 days (objects or arrays)
function sumLastNDaysObj_(hist, days) {
  const startMs = Date.now() - days * 864e5;
  let sum = 0;
  for (const r of Array.isArray(hist) ? hist : []) {
    const t = Date.parse((r.date || '') + 'T00:00:00Z');
    if (!isNaN(t) && t >= startMs) sum += Number(r.volume || 0);
  }
  return sum;
}

// Read list from Named Range OR "Sheet!A1" (numeric IDs; de-dupe; preserve order)
// Read list from Named Range OR "Sheet!A1" (numeric/string IDs; de-dupe; preserve order)
function readListFlex_(spec, opts) {
  opts = Object.assign({
    numeric: true,       // coerce to number
    integer: true,       // floor numbers (IDs)
    dropZeros: true,     // ignore 0
    dedupe: true,
    validator: null      // fn(val)->boolean
  }, opts || {});

  const ss = SpreadsheetApp.getActive();
  let range = (spec && typeof spec.getValues === 'function') ? spec : ss.getRangeByName(spec);
  if (!range && typeof spec === 'string' && spec.includes('!')) {
    const m = spec.match(/^'?([^'!]+)'?\!(.+)$/);
    if (!m) throw new Error('Invalid list spec: ' + spec);
    const sh = ss.getSheetByName(m[1]);
    if (!sh) throw new Error('Sheet not found: ' + m[1]);
    range = sh.getRange(m[2]);
  }
  if (!range) throw new Error('List source not found: ' + spec);

  const seen = new Set(), out = [];
  const vals = range.getValues().flat();

  for (let v of vals) {
    // normalize to trimmed string first to catch blanks & whitespace
    let s = String(v).trim();
    if (!s.length) continue;                     // ← ignore blanks

    let val;
    if (opts.numeric) {
      let n = Number(s);                         // handles "1.0000043E7"
      if (!Number.isFinite(n)) continue;
      if (opts.integer) n = Math.floor(n);
      if (opts.dropZeros && n === 0) continue;   // ← avoid the "0" poison
      val = n;
    } else {
      val = s;
    }

    if (opts.validator && !opts.validator(val)) continue;
    if (!opts.dedupe || !seen.has(val)) {
      if (opts.dedupe) seen.add(val);
      out.push(val);
    }
  }
  return out;
}


// Upsert (type_id, region_id) rows into Cache_Market_ESI_Region
function upsertRegionCache_(sheet, rows) {
  if (!rows.length) return;
  const data = sheet.getDataRange().getValues();
  const h = data[0];
  const iType = h.indexOf('type_id');
  const iReg  = h.indexOf('region_id');
  const map = new Map();
  for (let r=1; r<data.length; r++) {
    const t = data[r][iType], g = data[r][iReg];
    if (t==='' || g==='') continue;
    map.set(`${t}:${g}`, r+1);
  }
  const updates = [], appends = [];
  for (const row of rows) {
    const [typeId, regionId] = row;
    const rn = map.get(`${typeId}:${regionId}`);
    if (rn) updates.push({ rn, row }); else appends.push(row);
  }
  updates.forEach(u => sheet.getRange(u.rn, 1, 1, HEADERS_REGION.length).setValues([u.row]));
  if (appends.length) {
    const start = sheet.getLastRow()+1;
    sheet.getRange(start, 1, appends.length, HEADERS_REGION.length).setValues(appends);
  }
}


/* =============================== DEBUG ================================= */
function debugListSources() {
  const items   = readListFlex_(ITEMS_SOURCE,   {numeric:true});
  const regions = readListFlex_(REGIONS_SOURCE, {numeric:true});
  console.log({ items_count: items.length, regions_count: regions.length,
                sample_items: items.slice(0,5), sample_regions: regions.slice(0,5) });
}
