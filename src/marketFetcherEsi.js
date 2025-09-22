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
const ITEMS_SOURCE = "'Item List Back End'!A2:A";   // type_id list
const REGIONS_SOURCE = "'Market Settings'!F3:F";      // region_id list

// Engine sheets & export name
const SHEET_REGION_CACHE = 'Cache_Market_ESI_Region';
const SHEET_PUBLISH = 'Publish_ESI_Region';
const NR_MARKET_RESULT = 'MarketResultESI_Region';

// Worker cadence & batching
const WORKER_EVERY_MIN = 5;     // chunk worker interval (minutes)
const BATCH_SIZE = 25;    // items per chunk per region
const CALL_PACE_MS = 500;   // sleep between GESI calls (ms)

// Daily kickoff time (EVE DT ~11:00 UTC)
const DAILY_UTC_HOUR = 11;
const DAILY_UTC_MIN = 20;

// Engine headers
const HEADERS_REGION = ['type_id', 'region_id', 'volume30_region', 'velocity_region', 'last_updated', 'status'];
const HEADERS_PUBLISH = ['type_id', 'region_id', 'volume30_region', 'last_updated', 'status'];


/** ---------- Rate limit + backoff (lightweight) ---------- **/
const PACER_PROP_GAP = 'esi_gap_ms';
const PACER_PROP_LAST = 'esi_last_ms';
const PACER_MIN_MS = 400;   // start ~2.5 rps
const PACER_MAX_MS = 5000;  // cap 5s between calls

function isRateLimitedError_(e) {
  const s = String(e && e.message || e || '');
  return /bandwidth quota exceeded|error limit|too many|429|420/i.test(s);
}

function rateLimitPacer_() {
  const p = PropertiesService.getScriptProperties();
  const gap = Math.max(PACER_MIN_MS, Math.min(PACER_MAX_MS, Number(p.getProperty(PACER_PROP_GAP)) || PACER_MIN_MS));
  const last = Number(p.getProperty(PACER_PROP_LAST)) || 0;
  const now = Date.now();
  const wait = Math.max(0, last + gap - now);
  if (wait > 0) Utilities.sleep(wait);
  p.setProperty(PACER_PROP_LAST, String(Date.now()));
  return gap;
}

function bumpPacer_() {
  const p = PropertiesService.getScriptProperties();
  const cur = Math.max(PACER_MIN_MS, Number(p.getProperty(PACER_PROP_GAP)) || PACER_MIN_MS);
  const next = Math.min(PACER_MAX_MS, Math.ceil(cur * 1.5));
  p.setProperty(PACER_PROP_GAP, String(next));
  return next;
}

function relaxPacer_() {
  const p = PropertiesService.getScriptProperties();
  const cur = Math.max(PACER_MIN_MS, Number(p.getProperty(PACER_PROP_GAP)) || PACER_MIN_MS);
  const next = Math.max(PACER_MIN_MS, Math.floor(cur * 0.8));
  p.setProperty(PACER_PROP_GAP, String(next));
  return next;
}

function withRetries_(fn, tries = 3, baseDelay = 800) {
  let attempt = 0;
  while (true) {
    try {
      rateLimitPacer_();              // global pacing
      return fn();                    // do the call
    } catch (e) {
      if (isRateLimitedError_(e) && attempt < tries - 1) {
        const gap = bumpPacer_();
        const jitter = Math.floor(Math.random() * 300);
        const backoff = baseDelay * Math.pow(2, attempt) + jitter;
        (LoggerEx?.warn || console.warn)(`Rate-limited: gap=${gap}ms, backoff=${backoff}ms, try=${attempt + 1}/${tries}`);
        Utilities.sleep(backoff);
        attempt++;
        continue;
      }
      throw e; // non-rate error or out of tries
    }
  }
}

function fetchVol30One_(typeId, regionId) {
  try {
    const hist = withRetries_(
      () => GESI.getClient()
                .setFunction('markets_region_history')
                .executeRaw({ type_id: typeId, region_id: regionId }),
      3, 800
    ); // -> [{date:"YYYY-MM-DD", volume:..., ...}]

    if (!Array.isArray(hist) || hist.length === 0) {
      return { status: 'NOT_FOUND', vol30: "", vel: "" };
    }

    const vol30 = sumLastNDaysObj_(hist, 30); // sums 'volume' last 30d (UTC)

    if (vol30 > 0) {
      return { status: 'OK', vol30: vol30, vel: vol30 / 30 };
    }
    // zero over the last 30 days → treat as no data for math purposes
    return { status: 'NO_DATA_30D', vol30: "", vel: "" };

  } catch (e) {
    if (isRateLimitedError_(e)) {
      return { status: 'ERR_RATE' }; // do not write a row; retry next run
    }
    (LoggerEx?.error || console.error)('ESI error:', e?.message || e, e?.stack);
    return { status: 'ERR_ESI', vol30: "", vel: "" };
  }
}




/* ============================ CONFIG: CLIENT ============================ */
/* marketStatDataCache uses your CacheService via _getCachedFuz and _normalizeOrder.
 * It needs NO sheet/range config. In your Sheets formulas, add your own FUZ_TICK heartbeat.
 */

function testESIHistory() {
  const rid = 10000043; // Domain
  const tid = 34;       // Tritanium
  const res = GESI.markets_region_history(tid, rid);
  console.log('history rows:', Array.isArray(res) ? res.length : res);
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
  const hdr = sh.getRange(1, 1, 1, HEADERS_REGION.length).getValues()[0];
  const needs = !hdr[0] || HEADERS_REGION.some((h, i) => hdr[i] !== h);
  if (needs) {
    sh.clear();
    sh.getRange(1, 1, 1, HEADERS_REGION.length).setValues([HEADERS_REGION]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensurePublishSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_PUBLISH);
  if (!sh) sh = ss.insertSheet(SHEET_PUBLISH);
  const hdr = sh.getRange(1, 1, 1, HEADERS_PUBLISH.length).getValues()[0];
  const needs = !hdr[0] || HEADERS_PUBLISH.some((h, i) => hdr[i] !== h);
  if (needs) {
    sh.clear();
    sh.getRange(1, 1, 1, HEADERS_PUBLISH.length).setValues([HEADERS_PUBLISH]);
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
  if (!cache) throw new Error('Missing ' + SHEET_REGION_CACHE);

  markAllCacheStale_(cache);

  const props = PropertiesService.getScriptProperties();
  props.setProperty('mf_cursor', JSON.stringify({ ri: 0, ii: 0 }));
  props.setProperty('mf_job_active', '1');
  props.setProperty('mf_job_started', new Date().toISOString());

  publishMarketResultESIRegion(); // clients see stale→fresh tick

  ScriptApp.newTrigger('marketFetchChunk').timeBased().everyMinutes(WORKER_EVERY_MIN).create();
}


/* ===================== WORKER: ESI → CACHE (GESI) ====================== */
function marketFetchChunk() {
  var t0 = Date.now();
  var rid = Utilities.getUuid().slice(0, 8); // run id for grouping

  // prevent overlap
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    LoggerEx.warn('mf.lock.busy rid=' + rid);
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cache = ss.getSheetByName(SHEET_REGION_CACHE);
    if (!cache) {
      LoggerEx.warn('mf.early.no_cache rid=' + rid);
      return stopWorker_('no cache');
    }

    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('mf_job_active') !== '1') {
      LoggerEx.log('mf.early.no_active_job rid=' + rid);
      return stopWorker_('no active job');
    }

    var regions = readListFlex_(REGIONS_SOURCE, {
      numeric: true, integer: true, dropZeros: true, validator: isValidRegionId_
    });
    var items = readListFlex_(ITEMS_SOURCE, { numeric: true, integer: true });

    if (!items.length || !regions.length) {
      LoggerEx.warn('mf.early.no_items_or_regions rid=' + rid + ' items=' + items.length + ' regions=' + regions.length);
      return stopWorker_('no items/regions');
    }

    var cur = JSON.parse(props.getProperty('mf_cursor') || '{"ri":0,"ii":0}');
    var ri = Math.min(cur.ri || 0, regions.length - 1);
    var ii = Math.min(cur.ii || 0, items.length - 1);

    var regionId = regions[ri];
    var end = Math.min(ii + BATCH_SIZE, items.length);
    var chunk = items.slice(ii, end);

    LoggerEx.log('mf.run.start rid=' + rid + ' regions=' + regions.length + ' items=' + items.length +
                 ' cursorIn=' + JSON.stringify({ri:ri, ii:ii}) +
                 ' region=' + regionId + ' window=' + JSON.stringify({from:ii, to:end, size:chunk.length}));

    var rowsOut = [];
    var processed = 0;
    var hitRate = false;
    var ok = 0, err = 0;

    for (var c = 0; c < chunk.length; c++) {
      var typeId = chunk[c];
      var res = fetchVol30One_(typeId, regionId); // retries + pacer inside

      if (res.status === 'ERR_RATE') {
        hitRate = true;
        LoggerEx.warn('mf.rate.hit rid=' + rid + ' region=' + regionId + ' type=' + typeId + ' processed=' + processed);
        break; // leave remaining items for next run
      }

      if (res.status === 'OK') {
        ok++;
        rowsOut.push([typeId, regionId, res.vol30, res.vel, new Date(), res.status]);
      } else {
        err++;
        rowsOut.push([typeId, regionId, "", "", new Date(), res.status]);
        LoggerEx.warn('mf.item.err rid=' + rid + ' region=' + regionId + ' type=' + typeId + ' status=' + res.status);
      }

      processed++;
    }

    // Flush once
    if (rowsOut.length) {
      upsertRegionCache_(cache, rowsOut);
      publishMarketResultESIRegion();
      LoggerEx.log('mf.flush.chunk rid=' + rid + ' rows=' + rowsOut.length);
    } else {
      LoggerEx.log('mf.flush.empty rid=' + rid);
    }

    // Advance cursor only by actual work
    var cursorBefore = { ri: ri, ii: ii };
    if (hitRate && processed === 0) {
      if (typeof hardenPacer_ === 'function') hardenPacer_();
      LoggerEx.warn('mf.cursor.stay_due_to_rate rid=' + rid + ' region=' + regionId +
                    ' cursor=' + JSON.stringify(cursorBefore));
      // keep ri/ii unchanged
    } else {
      ii += processed;
      if (ii >= items.length) { ii = 0; ri++; }
      if (!hitRate && typeof relaxPacer_ === 'function') relaxPacer_();
    }

    var done = (ri >= regions.length);
    var cursorAfter = done ? null : { ri: ri, ii: ii };

    LoggerEx.log('mf.run.summary rid=' + rid + ' region=' + regionId +
                 ' ok=' + ok + ' err=' + err + ' hitRate=' + hitRate +
                 ' processed=' + processed + ' elapsedMs=' + (Date.now()-t0) +
                 ' cursorBefore=' + JSON.stringify(cursorBefore) +
                 ' cursorAfter=' + JSON.stringify(cursorAfter) + ' done=' + done);

    if (done) {
      props.deleteProperty('mf_job_active');
      props.deleteProperty('mf_cursor');
      LoggerEx.log('mf.done rid=' + rid);
      return stopWorker_('done');
    }

    props.setProperty('mf_cursor', JSON.stringify({ ri: ri, ii: ii }));

  } catch (e) {
    LoggerEx.error('mf.run.exception rid=' + rid + ' msg=' + (e && e.message));
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
    LoggerEx.log('mf.run.end rid=' + rid + ' ms=' + (Date.now() - t0));
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
  if (!["region", "system", "station"].includes(lt)) {
    return Array.isArray(type_ids) ? type_ids.map(() => [""]) : "";
  }

  const norm = (typeof _normalizeOrder === "function")
    ? _normalizeOrder(order_type, order_level)
    : {
      type: (String(order_type || "sell").toLowerCase() === "buy" ? "buy" : "sell"),
      level: String(order_level || "volume").toLowerCase()
    };

  // preserve input shape
  const in2D = Array.isArray(type_ids) ? type_ids : [[type_ids]];
  const rows = in2D.length, cols = in2D[0].length;

  // flatten ids, keep nulls for placeholders
  const flatIds = [];
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
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
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
    const id = flatIds[k++];
    out[r][c] = (id == null) ? "" : (pick(have[id]) ?? "");
  }
  return Array.isArray(type_ids) ? out : out[0][0];
}


/* =============================== HELPERS ================================ */
function publishMarketResultESIRegion() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const cache = ss.getSheetByName(SHEET_REGION_CACHE);
  const pub   = ss.getSheetByName(SHEET_PUBLISH);
  if (!cache || !pub) throw new Error('Missing sheets');

  const keepPairs = getConfigPairSet_();               // ← NEW

  const vals = cache.getDataRange().getValues();
  const out  = [HEADERS_PUBLISH];

  if (vals.length > 1) {
    const h  = vals[0];
    const ix = { t:h.indexOf('type_id'), r:h.indexOf('region_id'),
                 v:h.indexOf('volume30_region'), u:h.indexOf('last_updated'),
                 s:h.indexOf('status') };

    for (let i = 1; i < vals.length; i++) {
      const row = vals[i];
      const t = row[ix.t], r = row[ix.r];
      if (t === "" || r === "") continue;
      if (!keepPairs.has(`${Math.floor(t)}:${Math.floor(r)}`)) continue;  // ← NEW

      let status = String(row[ix.s] || "").toUpperCase();
      if (status === "STALE") status = "STALE-ESI";

      const keepNumeric = (status === "OK" || status === "STALE-ESI");
      const vol = keepNumeric ? row[ix.v] : "";

      out.push([t, r, vol, row[ix.u], status]);
    }
  }

  pub.clearContents();
  pub.getRange(1, 1, out.length, HEADERS_PUBLISH.length).setValues(out);
  pub.getRange('H1').setNumberFormat('yyyy-mm-dd\"T\"hh:mm:ss\"Z\"').setValue(new Date());
  SpreadsheetApp.getActive().setNamedRange(NR_MARKET_RESULT, pub.getRange('A:E'));
}



function markAllCacheStale_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const h = data[0];
  const iStatus = h.indexOf('status');
  if (iStatus === -1) throw new Error('status col missing');

  // Build status-only updates, conditional per row
  const newStatuses = [];
  for (let r = 1; r < data.length; r++) {
    const cur = String(data[r][iStatus] || "").toUpperCase();
    // Flip only good rows to STALE-ESI; leave others as-is
    const next = (cur === "OK" || cur === "STALE-ESI") ? "STALE-ESI" : cur;
    newStatuses.push([next]);
  }

  // Write just the status column; do NOT overwrite dates or numbers
  sheet.getRange(2, iStatus + 1, newStatuses.length, 1).setValues(newStatuses);
}


function stopWorker_(reason) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'marketFetchChunk')
    .forEach(t => ScriptApp.deleteTrigger(t));
  try { publishMarketResultESIRegion(); } catch (e) { }
  if (reason) console.log('marketFetchChunk: stopped (' + reason + ')');
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
  const iReg = h.indexOf('region_id');
  const map = new Map();
  for (let r = 1; r < data.length; r++) {
    const t = data[r][iType], g = data[r][iReg];
    if (t === '' || g === '') continue;
    map.set(`${t}:${g}`, r + 1);
  }
  const updates = [], appends = [];
  for (const row of rows) {
    const [typeId, regionId] = row;
    const rn = map.get(`${typeId}:${regionId}`);
    if (rn) updates.push({ rn, row }); else appends.push(row);
  }
  updates.forEach(u => sheet.getRange(u.rn, 1, 1, HEADERS_REGION.length).setValues([u.row]));
  if (appends.length) {
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, appends.length, HEADERS_REGION.length).setValues(appends);
  }
}


/* =============================== DEBUG ================================= */
function debugListSources() {
  const items = readListFlex_(ITEMS_SOURCE, { numeric: true });
  const regions = readListFlex_(REGIONS_SOURCE, { numeric: true });
  console.log({
    items_count: items.length, regions_count: regions.length,
    sample_items: items.slice(0, 5), sample_regions: regions.slice(0, 5)
  });
}


function _pairKey_(t, r){ return `${Math.floor(t)}:${Math.floor(r)}`; }

function getConfigPairSet_(){
  const items   = readListFlex_(ITEMS_SOURCE,   { numeric:true, integer:true });
  const regions = readListFlex_(REGIONS_SOURCE, { numeric:true, integer:true, dropZeros:true, validator: isValidRegionId_ });
  const set = new Set();
  for (const t of items) for (const r of regions) set.add(_pairKey_(t,r));
  return set;
}


function pruneCacheToConfig_(mode='tombstone'){
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_REGION_CACHE);
  if (!sh) throw new Error('Missing '+SHEET_REGION_CACHE);

  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return;

  const h   = vals[0];
  const iT  = h.indexOf('type_id');
  const iR  = h.indexOf('region_id');
  const iV  = h.indexOf('volume30_region');
  const iVel= h.indexOf('velocity_region');
  const iU  = h.indexOf('last_updated');
  const iS  = h.indexOf('status');

  const keepPairs = getConfigPairSet_();

  if (mode === 'hard'){
    // Rebuild the sheet with only configured rows
    const out = [h];
    for (let i=1;i<vals.length;i++){
      const row = vals[i];
      if (keepPairs.has(_pairKey_(row[iT], row[iR]))) out.push(row);
    }
    sh.clear();
    sh.getRange(1,1,out.length,out[0].length).setValues(out);
    sh.setFrozenRows(1);
    return;
  }

  // tombstone: mark non-config rows as REMOVED and blank their numbers
  const updates = [];
  for (let i=1;i<vals.length;i++){
    const row = vals[i];
    const keep = keepPairs.has(_pairKey_(row[iT], row[iR]));
    if (!keep) {
      row[iV] = "";               // blank numerics
      if (iVel > -1) row[iVel] = "";
      if (iU > -1)   row[iU] = new Date();
      row[iS] = "REMOVED";
      updates.push({ rn:i+1, row });
    }
  }
  updates.forEach(u => sh.getRange(u.rn, 1, 1, h.length).setValues([u.row]));
}


/**
 * Build a values-only table per client from Publish_ESI_Region.
 * Input:  Sheet "Interface Clients" → A:Client, B:region_id_heartbeat
 * Source: Sheet "Publish_ESI_Region" (A:E = type_id,region_id,volume30_region,last_updated,status)
 * Output: Sheet "Publish_ESI_Region_<client>" with headers:
 *         [type_id, volume30_region, last_updated, status]
 * Named range for client: MarketResultESI_Region_<client>
 */
function ESI_publishClientInterfaces() {
  const ss    = SpreadsheetApp.getActive();
  const srcSh = ss.getSheetByName('Publish_ESI_Region');
  const cliSh = ss.getSheetByName('Interface Clients');
  if (!srcSh || !cliSh) throw new Error('Missing required sheets');

  const src = srcSh.getDataRange().getValues();
  if (src.length < 2) return;

  // Build region → rows from source (A:E fixed schema)
  // A=0 type_id, B=1 region_id, C=2 volume30_region, D=3 last_updated, E=4 status
  const byRegion = new Map();
  for (let r = 1; r < src.length; r++) {
    const row = src[r];
    const rid = Number(row[1]);
    if (!Number.isFinite(rid)) continue;
    const packed = [row[0], row[2], row[3], row[4]]; // [type_id, vol30, last_updated, status]
    if (!byRegion.has(rid)) byRegion.set(rid, []);
    byRegion.get(rid).push(packed);
  }

  // Read clients: A:Client, B:region_id_heartbeat, (optional) C:status_filter like "OK|STALE-ESI"
  const lastCliRow = cliSh.getLastRow();
  if (lastCliRow < 2) return;
  const cliRows = cliSh.getRange(2, 1, lastCliRow - 1, 3).getValues();

  const OUT_HEADERS = ['type_id','volume30_region','last_updated','status'];

  // helper: normalize status to uppercase with hyphens
  const normStatus = s => String(s || '').trim().toUpperCase().replace(/_/g, '-');

  for (const [clientRaw, ridRaw, filtRaw] of cliRows) {
    const client = String(clientRaw || '').trim();
    const rid    = Number(ridRaw);
    if (!client || !Number.isFinite(rid)) continue;

    // Build allowed set (default: OK|STALE-ESI)
    const allowed = new Set(
      String(filtRaw || 'OK|STALE-ESI')
        .split(/[|,]/)
        .map(x => normStatus(x))
        .filter(Boolean)
    );

    // Filter rows by region + status
    const rows = (byRegion.get(rid) || []).filter(r => allowed.has(normStatus(r[3])));

    const outName = 'Publish_ESI_Region_' + client;
    const outSh   = getOrCreateSheet(ss, outName, OUT_HEADERS); // your Utility helper

    // Clear data rows, keep header
    const last = outSh.getLastRow();
    if (last > 1) outSh.getRange(2, 1, last - 1, OUT_HEADERS.length).clearContent();

    // Write values-only
    if (rows.length) {
      outSh.getRange(2, 1, rows.length, OUT_HEADERS.length).setValues(rows);
      outSh.getRange(2, 3, rows.length, 1).setNumberFormat('yyyy-mm-dd'); // last_updated
    }

    // Named range for clients to IMPORTRANGE
    const height = Math.max(1, rows.length + 1);
    ss.setNamedRange('MarketResultESI_Region_' + client,
      outSh.getRange(1, 1, height, OUT_HEADERS.length));
  }
}

/**
 * Install/replace a time trigger to republish client interfaces.
 * @param {number} everyMinutes  One of: 1, 5, 10, 15, 30 (default 10)
 */
function ESI_installClientInterfaceRefresh(everyMinutes) {
  const allowed = [1,5,10,15,30];
  const n = Number(everyMinutes || 5);
  if (!allowed.includes(n)) throw new Error('everyMinutes must be 1, 5, 10, 15, or 30.');
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'ESI_publishClientInterfaces')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('ESI_publishClientInterfaces').timeBased().everyMinutes(n).create();
}

/** Remove the refresh trigger */
function ESI_stopClientInterfaceRefresh() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'ESI_publishClientInterfaces')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

