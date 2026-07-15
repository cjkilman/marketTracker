/***** marketFetcherESI.gs  —  MAIN PROJECT FILE ********************************
 * Contains BOTH sides:
 * 1) ENGINE: pulls ESI region history via GESI → caches & publishes slim table
 * 2) CLIENT: marketStatDataCache(...)  (cache-only; reads CacheService via your
 * _getCachedFuz and _normalizeOrder from marketStatData)
 *
 * SLIM IS LAW:
 * - Engine publishes named range: MarketResultESI_Region (type_id, region_id, vol30, ts, status)
 * - Clients import that single range; all other math stays local (Fuz books + DoB).
 *******************************************************************************/

/* ============================ CONFIG: ENGINE ============================ */
// Sources can be a Named Range OR "Sheet!A1" range
const ITEMS_SOURCE = "'Item List Back End'!A2:B";   // type_id list
const REGIONS_SOURCE = "'Market Settings'!F3:F";      // region_id list

// Engine sheets & export name
const SHEET_REGION_CACHE = 'Cache_Market_ESI_Region';
const SHEET_PUBLISH = 'Publish_ESI_Region';
const NR_MARKET_RESULT = 'MarketResultESI_Region';

// Worker cadence & batching
const WORKER_EVERY_MIN = 1;     // chunk worker interval (minutes)
const BATCH_SIZE = 200;    // items per chunk per region
const CALL_PACE_MS = 100;   // sleep between GESI calls (ms)

// Daily kickoff time (EVE DT ~11:00 UTC)
const DAILY_UTC_HOUR = 11;
const DAILY_UTC_MIN = 20;

// Engine headers
const HEADERS_REGION = ['type_id', 'region_id', 'volume30_region', 'velocity30_region', 'volume7_region', 'velocity7_region', 'volume5_region', 'velocity5_region', 'last_updated', 'status'];
const HEADERS_PUBLISH = ['type_id', 'region_id', 'volume30_region', 'volume7_region', 'volume5_region', 'last_updated', 'status'];

/* ============================ ESI SYSTEM HOOKS ============================ */
const getMarketEsiHooks = (authChar) => ({
  onLock: (err) => {
    LoggerEx.error('FATAL: Quota Exceeded mid-fetch. Gate locked.');
    stopWorker_('Google Quota Exceeded'); 
  },
  onAuthError: (err) => {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('mf_blocked_auth', '1');
    props.deleteProperty('mf_job_active');
    
    const charName = authChar || 'Default Character';
    LoggerEx.error(`Auth Failure for [${charName}] on ${err.endpoint}. Hint: Re-authorize GESI.`);
    
    stopWorker_('auth required');
  }
});

/***** GESI batching + governor (≤300 req/min) ********************************/
var ESI_TB_RATE_PER_MIN = 300;   // tokens added per minute
var ESI_TB_BURST = 300;   // max token bucket
var ESI_GROUP_SIZE = 50;    // requests per fetchAll group
var ESI_MICRO_BREATH_MS = 100;   // tiny pause between groups

function tbConsume_(need) {
  var props = PropertiesService.getScriptProperties();
  var now = Date.now();
  var last = +props.getProperty('esi_tb_last_ms') || 0;
  var tok = +props.getProperty('esi_tb_tokens') || 0;

  if (!last) { last = now; tok = ESI_TB_BURST; } // start full to avoid first-stall

  var ratePerMs = ESI_TB_RATE_PER_MIN / 60000;                  // tokens/ms
  tok = Math.min(ESI_TB_BURST, tok + ratePerMs * (now - last)); // refill

  if (tok >= need) {
    tok -= need;
    props.setProperty('esi_tb_tokens', String(tok));
    props.setProperty('esi_tb_last_ms', String(now));
    return 0;
  }

  var shortfall = need - tok;
  var waitMs = Math.ceil(shortfall / ratePerMs);
  props.setProperty('esi_tb_tokens', '0');
  props.setProperty('esi_tb_last_ms', String(now));
  (LoggerEx?.info || console.info)('esi.tb.wait', { ms: waitMs, need: need });
  return waitMs;
}

function isBeforeNoonET_() {
  return +Utilities.formatDate(new Date(), 'America/New_York', 'H') < 12;
}
function getBatchSize_() {      // AM: 200, PM: 60
  return isBeforeNoonET_() ? 200 : 60;
}

function summarizeHistoryMulti_(rows) {
  if (!rows || !rows.length) return { vol30: 0, vel30: 0, vol7: 0, vel7: 0, vol5: 0, vel5: 0 };

  var v30 = 0, v7 = 0, v5 = 0;
  var daysToCount = Math.min(rows.length, 30);

  for (var k = 1; k <= daysToCount; k++) {
    var dailyVol = +rows[rows.length - k].volume || 0;
    v30 += dailyVol;
    if (k <= 7) v7 += dailyVol;
    if (k <= 5) v5 += dailyVol;
  }

  var n30 = daysToCount;
  var n7 = Math.min(rows.length, 7);
  var n5 = Math.min(rows.length, 5);

  return {
    vol30: v30, vel30: (n30 ? v30 / n30 : 0),
    vol7: v7, vel7: (n7 ? v7 / n7 : 0),
    vol5: v5, vel5: (n5 ? v5 / n5 : 0)
  };
}

function fetchHistoryBatchGESI_(regionId, typeIds, authChar) {
  // THE FIX: Only pass authChar if it actually contains a string
  const authClient = authChar ? GESI.getClient(authChar) : GESI.getClient();
  
  const hooks = getMarketEsiHooks(authChar);
  const service = ESI.forEndpoint(authClient, 'markets_region_history', hooks, null);

  const wait = tbConsume_(typeIds.length);
  if (wait > 0) Utilities.sleep(Math.min(wait, 30000));

  return typeIds.map(typeId => {
    if (ESI.isLocked()) {
      return { typeId: typeId, status: 'ERR_RATE' };
    }

    const response = service.get({ 
      region_id: regionId, 
      type_id: typeId 
    });

    if (response.error) {
      if (response.error === 'LOCKED' || response.error.includes('429') || response.error.includes('420')) {
        return { typeId: typeId, status: 'ERR_RATE' };
      }
      if (response.error === 'AUTH_FAILURE') {
        return { typeId: typeId, status: 'ERR_AUTH' };
      }
      return { typeId: typeId, status: 'ERR_ESI' };
    }

    const hist = response.data;
    if (!Array.isArray(hist) || hist.length === 0) {
      return { typeId: typeId, status: 'NOT_FOUND' };
    }

    const m = summarizeHistoryMulti_(hist);
    return {
      typeId: typeId,
      status: 'OK',
      vol30: m.vol30, vel30: m.vel30,
      vol7:  m.vol7,  vel7:  m.vel7,
      vol5:  m.vol5,  vel5:  m.vel5
    };
  });
}

function fetchVol30One_(typeId, regionId, authChar) {
  // THE FIX: Only pass authChar if it actually contains a string
  const authClient = authChar ? GESI.getClient(authChar) : GESI.getClient(); 
  
  const service = ESI.forEndpoint(authClient, 'markets_region_history', getMarketEsiHooks(authChar), null);
  
  const response = service.get({ region_id: regionId, type_id: typeId });

  if (response.error) {
    if (response.error === 'LOCKED' || response.error.includes('429') || response.error.includes('420')) {
      return { status: 'ERR_RATE', vol30: "", vel: "" };
    }
    const log = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('FETCH_ONE') : console;
    log.error(`ESI error for Item ${typeId} in Region ${regionId}: ${response.error}`);
    return { status: 'ERR_ESI', vol30: "", vel: "" };
  }

  const hist = response.data;
  if (!Array.isArray(hist) || hist.length === 0) {
    return { status: 'NOT_FOUND', vol30: "", vel: "" };
  }

  const m = summarizeHistoryMulti_(hist);
  if (m.vol30 > 0) return { status: 'OK', vol30: m.vol30, vel: m.vel30 };
  
  return { status: 'NO_DATA_30D', vol30: "", vel: "" };
}

/* ============================ CONFIG: CLIENT ============================ */
function testESIHistory() {
  const rid = 10000043; // Domain
  const tid = 34;       // Tritanium
  const client = GESI.getClient().setFunction('markets_region_history');
  const resp = UrlFetchApp.fetch(client.buildRequest({
    region_id: rid, type_id: tid, show_column_headings: false, version: 'latest'
  }));
  const rows = JSON.parse(resp.getContentText());
  console.log('history rows:', Array.isArray(rows) ? rows.length : rows);
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
  sh.getRange('G1').setValue('published_at');
  sh.getRange('H1').setNumberFormat('yyyy-mm-dd"T"hh:mm:ss"Z"');
  SpreadsheetApp.getActive().setNamedRange(NR_MARKET_RESULT, sh.getRange('A:E'));
}

/* ============================ SCHEDULING =============================== */
function installDailyKickoff() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const h = t.getHandlerFunction();
    if (h === 'kickoffMarketHistoryRefresh' || h === 'marketFetchChunk') {
      ScriptApp.deleteTrigger(t);
    }
  });
  
  ScriptApp.newTrigger('kickoffMarketHistoryRefresh')
    .timeBased().everyDays(1).atHour(DAILY_UTC_HOUR).nearMinute(DAILY_UTC_MIN)
    .inTimezone('Etc/UTC').create();
  
  // NON-BLOCKING FIX: Log to console and use toast instead of alert
  const msg = "Daily kickoff installed @ " + DAILY_UTC_HOUR + ":" + DAILY_UTC_MIN + " UTC";
  console.log("[SUCCESS] " + msg);
  
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Trigger Setup");
  } catch (e) {
    // Safety fallback if executed outside spreadsheet runtime context
  }
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

  ScriptApp.newTrigger('marketFetchChunk').timeBased().everyMinutes(WORKER_EVERY_MIN).create();
}

/* ===================== WORKER: ESI → CACHE (GESI) ====================== */
function marketFetchChunk() {
  const props = PropertiesService.getScriptProperties();
  
  if (ESI.isLocked()) {
    LoggerEx.error('ABORT: ESI global quota is locked. Skipping execution to protect account.');
    stopWorker_('Quota Lock Active');
    return;
  }
  
  var t0 = Date.now();
  var rid = Utilities.getUuid().slice(0, 8);

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
    var BS = (typeof getBatchSize_ === 'function') ? getBatchSize_() : BATCH_SIZE;
    var end = Math.min(ii + BS, items.length);
    LoggerEx.info('mf.batch.params', { BS: BS, cadence_min: WORKER_EVERY_MIN });

    var chunk = items.slice(ii, end);

    LoggerEx.log('mf.run.start rid=' + rid + ' regions=' + regions.length + ' items=' + items.length +
      ' cursorIn=' + JSON.stringify({ ri: ri, ii: ii }) +
      ' region=' + regionId + ' window=' + JSON.stringify({ from: ii, to: end, size: chunk.length }));

    var rowsOut = [];
    var processed = 0;
    var hitRate = false;
    var ok = 0, err = 0, errAuth = 0;

    var results = fetchHistoryBatchGESI_(regionId, chunk);

    for (var i = 0; i < results.length; i++) {
      var r = results[i];

      if (r.status === 'ERR_RATE') {
        hitRate = true;
        LoggerEx.warn('mf.rate.hit rid=' + rid + ' region=' + regionId + ' processed=' + processed);
        break;
      }

      if (r.status === 'ERR_AUTH') {
        errAuth++;
        rowsOut.push([r.typeId, regionId, "", "", "", "", "", "", new Date(), 'ERR_AUTH']);
        continue;
      }

      if (r.status === 'OK') {
        ok++;
        rowsOut.push([r.typeId, regionId, r.vol30, r.vel30, r.vol7, r.vel7, r.vol5, r.vel5, new Date(), 'OK']);
        if (r.vol30 === 0) (LoggerEx?.info || console.info)('mf.vol30.zero', { regionId: regionId, typeId: r.typeId });
        processed++;
      } else {
        err++;
        rowsOut.push([r.typeId, regionId, "", "", "", "", "", "", new Date(), r.status]);
        LoggerEx.warn('mf.item.err rid=' + rid + ' region=' + regionId + ' type=' + r.typeId + ' status=' + r.status);
        processed++;
      }
    }

    if (rowsOut.length) {
      upsertRegionCache_(cache, rowsOut);
      LoggerEx.log('mf.flush.chunk rid=' + rid + ' rows=' + rowsOut.length);
    } else {
      LoggerEx.log('mf.flush.empty rid=' + rid);
    }

    if (errAuth > 0) {
      props.setProperty('mf_blocked_auth', '1');
      props.deleteProperty('mf_job_active');
      LoggerEx.error('mf.auth.block rid=' + rid + ' region=' + regionId + ' errAuth=' + errAuth + ' hint="Sheets → Add-ons → GESI → Authorize Character"');
      return stopWorker_('auth required');
    }

    var cursorBefore = { ri: ri, ii: ii };
    if (hitRate && processed === 0) {
      LoggerEx.warn('mf.cursor.stay_due_to_rate rid=' + rid + ' region=' + regionId +
        ' cursor=' + JSON.stringify(cursorBefore));
    } else {
      ii += processed;
      if (ii >= items.length) { ii = 0; ri++; }
    }

    var done = (ri >= regions.length);
    var cursorAfter = done ? null : { ri: ri, ii: ii };

    LoggerEx.log('mf.run.summary rid=' + rid + ' region=' + regionId +
      ' ok=' + ok + ' err=' + err + ' hitRate=' + hitRate +
      ' processed=' + processed + ' elapsedMs=' + (Date.now() - t0) +
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
    const errorMsg = String(e && e.message);
    LoggerEx.error('mf.run.exception rid=' + rid + ' msg=' + errorMsg);
    throw e; 

  } finally {
    try { lock.releaseLock(); } catch (e2) { }
    LoggerEx.log('mf.run.end rid=' + rid + ' ms=' + (Date.now() - t0));
  }
}

function isValidRegionId_(n) {
  n = Math.floor(Number(n));
  return Number.isFinite(n) && n >= 10000000 && n < 20000000;
}

/**
 * GAP FILLER: Scans 'Item List Back End' vs 'Cache_Market_ESI_Region'.
 */
function fillMissingEsiHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cacheSh = ss.getSheetByName('Cache_Market_ESI_Region');

  const expectedIds = readListFlex_(ITEMS_SOURCE, { numeric: true });

  const existingIds = new Set();
  if (cacheSh && cacheSh.getLastRow() > 1) {
    const data = cacheSh.getRange(2, 1, cacheSh.getLastRow() - 1, 1).getValues();
    data.forEach(r => existingIds.add(Number(r[0])));
  }

  const missing = expectedIds.filter(id => !existingIds.has(id));

  if (missing.length === 0) {
    console.log("[GapFiller] Cache is complete. No missing items.");
    return;
  }

  console.log(`[GapFiller] Found ${missing.length} missing items. Fetching...`);

  const regions = readListFlex_(REGIONS_SOURCE, {
    numeric: true, integer: true, dropZeros: true, validator: isValidRegionId_
  });

  regions.forEach(rid => {
    const results = fetchHistoryBatchGESI_(rid, missing);
    const rowsOut = [];
    results.forEach(r => {
      if (r.status === 'OK') {
        rowsOut.push([
          r.typeId, rid, r.vol30, r.vel30, r.vol7, r.vel7, r.vol5, r.vel5, new Date(), 'OK'
        ]);
      } else {
        rowsOut.push([
          r.typeId, rid, "", "", "", "", "", "", new Date(), r.status
        ]);
      }
    });

    if (rowsOut.length) {
      upsertRegionCache_(ensureCacheMarketESIRegion_(ss), rowsOut);
      console.log(`[GapFiller] Saved ${rowsOut.length} items for Region ${rid}`);
    }
  });

  publishMarketResultESIRegion();
  console.log("[GapFiller] Complete. Cache & Publish updated.");
}

/* =============================== HELPERS ================================ */
function publishMarketResultESIRegion(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = ss.getSheetByName(SHEET_REGION_CACHE);
  const pub = ss.getSheetByName(SHEET_PUBLISH);
  if (!cache || !pub) throw new Error('Missing sheets');

  const keepPairs = getConfigPairSet_();

  const vals = cache.getDataRange().getValues();
  const out = [HEADERS_PUBLISH];

  if (vals.length > 1) {
    const h = vals[0];
    const ix = {
      t: h.indexOf('type_id'), r: h.indexOf('region_id'),
      v30: h.indexOf('volume30_region'),
      v7: h.indexOf('volume7_region'),
      v5: h.indexOf('volume5_region'),
      u: h.indexOf('last_updated'), s: h.indexOf('status')
    };

    for (let i = 1; i < vals.length; i++) {
      const row = vals[i];
      const t = row[ix.t], r = row[ix.r];
      if (t === "" || r === "") continue;
      if (!keepPairs.has(`${Math.floor(t)}:${Math.floor(r)}`)) continue;

      let status = String(row[ix.s] || "").toUpperCase();
      if (status === "STALE") status = "STALE-ESI";

      const keepNumeric = (status === "OK" || status === "STALE-ESI");
      const vol30 = keepNumeric ? row[ix.v30] : "";
      const vol7 = keepNumeric ? row[ix.v7] : "";
      const vol5 = keepNumeric ? row[ix.v5] : "";

      out.push([t, r, vol30, vol7, vol5, row[ix.u], status]);
    }
  }

  pub.clearContents();
  pub.getRange(1, 1, out.length, HEADERS_PUBLISH.length).setValues(out);
  pub.getRange('H1').setNumberFormat('yyyy-mm-dd\"T\"hh:mm:ss\"Z\"').setValue(new Date());

  if (out.length > 1) {
    pub.getRange(2, 3, out.length - 1, 3).setNumberFormat('0');
  }

  SpreadsheetApp.getActive().setNamedRange(NR_MARKET_RESULT, pub.getRange('A:E'));
}

/**
 * Build a values-only table per client from Publish_ESI_Region.
 * Output: Sheet "Publish_ESI_Region_<client>" with headers:
 * [type_id, volume30_region, last_updated, status]
 */
function ESI_publishVolumeInterfaces(ss) {
  if (!ss) ss = SpreadsheetApp.getActive();
  const srcSh = ss.getSheetByName('Publish_ESI_Region');
  const cliSh = ss.getSheetByName('Interface Clients');
  
  console.log("[PUBLISH] Starting Sync... Checking source sheets.");

  if (!srcSh || !cliSh) {
    console.error("[PUBLISH] Failed: One or more sheets missing.");
    return;
  }

  const src = srcSh.getDataRange().getValues();
  console.log("[PUBLISH] Source rows found: " + src.length);
  if (src.length < 2) {
    console.warn("[PUBLISH] Aborted: 'Publish_ESI_Region' is empty.");
    return;
  }

  const lastCliRow = cliSh.getLastRow();
  console.log("[PUBLISH] Interface Clients found: " + (lastCliRow - 1));
  if (lastCliRow < 2) {
    console.warn("[PUBLISH] Aborted: 'Interface Clients' has no entries.");
    return;
  }

  const normStatus = (s) => String(s || '').trim().toUpperCase().replace(/_/g, '-');
  const toInt = (v) => {
    const n = Number(String(v).replace(/[^\d\-]/g, ''));
    return Number.isFinite(n) ? n : NaN;
  };
  const toDate = (v) => {
    if (v instanceof Date) return v;
    const s = String(v || '').trim();
    if (!s) return null;
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (m) return new Date(Date.UTC(+m[1], +m[2]-1, +m[3]));
    const n = Number(s);
    if (Number.isFinite(n)) return new Date(Math.round((n - 25569) * 86400000)); 
    const t = Date.parse(s);
    return isNaN(t) ? null : new Date(t);
  };

  const byRegion = new Map();

  for (let r = 1; r < src.length; r++) {
    const row = src[r];
    const typeId = toInt(row[0]);
    const regionId = toInt(row[1]);
    if (!Number.isFinite(typeId) || !Number.isFinite(regionId)) continue;

    const vol30 = Number(row[2]) || 0;
    const last = toDate(row[3]);      
    const status = normStatus(row[4]);

    if (!byRegion.has(regionId)) byRegion.set(regionId, new Map());
    const bag = byRegion.get(regionId);

    const prev = bag.get(typeId);
    if (!prev) {
      bag.set(typeId, [typeId, vol30, last || '', status]);
    } else {
      const prevDate = prev[2] instanceof Date ? prev[2] : toDate(prev[2]);
      if ((last && !prevDate) || (last && prevDate && last > prevDate)) {
        bag.set(typeId, [typeId, vol30, last, status]);
      }
    }
  }

  if (lastCliRow < 2) return;
  const cliRows = cliSh.getRange(2, 1, lastCliRow - 1, 3).getValues();

  const OUT_HEADERS = ['type_id', 'volume30_region', 'last_updated', 'status'];

  for (const [clientRaw, ridRaw, filtRaw] of cliRows) {
    const client = String(clientRaw || '').trim();
    const regionId = toInt(ridRaw);
    if (!client || !Number.isFinite(regionId)) continue;

    const filt = String(filtRaw || 'OK|STALE-ESI').trim();
    const allowAll = normStatus(filt) === 'ALL';
    const allowed = new Set(
      allowAll ? [] : filt.split(/[|,]/).map((x) => normStatus(x)).filter(Boolean)
    );

    const bag = byRegion.get(regionId);
    const rows = [];
    if (bag && bag.size) {
      for (const [, rec] of bag) {
        if (!rec) continue;
        const st = normStatus(rec[3]);
        if (allowAll || allowed.has(st)) rows.push(rec);
      }
    }

    rows.sort((a, b) => (b[1] - a[1]) || (a[0] - b[0]));

    for (let i = 0; i < rows.length; i++) {
      const d = rows[i][2];
      rows[i][2] = toDate(d) || '';
    }

    const outName = 'Publish_ESI_Region_' + client;
    const outSh = getOrCreateSheet(ss, outName, OUT_HEADERS);

    const last = outSh.getLastRow();
    if (last > 1) outSh.getRange(2, 1, last - 1, OUT_HEADERS.length).clearContent();

    if (rows.length) {
      outSh.getRange(2, 1, rows.length, OUT_HEADERS.length).setValues(rows);
      outSh.getRange(2, 1, rows.length, 1).setNumberFormat('0');              
      outSh.getRange(2, 2, rows.length, 1).setNumberFormat('#,##0');          
      outSh.getRange(2, 3, rows.length, 1).setNumberFormat('yyyy-mm-dd');     
    }

    const height = Math.max(1, rows.length + 1);
    ss.setNamedRange(
      'MarketResultESI_Region_' + client,
      outSh.getRange(1, 1, height, OUT_HEADERS.length)
    );
  }
}

function markAllCacheStale_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const h = data[0];
  const iStatus = h.indexOf('status');
  if (iStatus === -1) throw new Error('status col missing');

  const newStatuses = [];
  for (let r = 1; r < data.length; r++) {
    const cur = String(data[r][iStatus] || "").toUpperCase();
    const next = (cur === "OK" || cur === "STALE-ESI") ? "STALE-ESI" : cur;
    newStatuses.push([next]);
  }

  sheet.getRange(2, iStatus + 1, newStatuses.length, 1).setValues(newStatuses);
}

function stopWorker_(reason) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'marketFetchChunk')
    .forEach(t => ScriptApp.deleteTrigger(t));
  try { publishMarketResultESIRegion(); } catch (e) { }
  if (reason) console.log('marketFetchChunk: stopped (' + reason + ')');
}

function sumLastNDaysObj_(hist, days) {
  const startMs = Date.now() - days * 864e5;
  let sum = 0;
  for (const r of Array.isArray(hist) ? hist : []) {
    const t = Date.parse((r.date || '') + 'T00:00:00Z');
    if (!isNaN(t) && t >= startMs) sum += Number(r.volume || 0);
  }
  return sum;
}

function readListFlex_(spec, opts) {
  opts = Object.assign({
    numeric: true, integer: true, dropZeros: true, dedupe: true, validator: null
  }, opts || {});

  const ss = SpreadsheetApp.getActive();
  
  let rangeObj = (spec && typeof spec.getValues === 'function') ? spec : ss.getRangeByName(spec);
  
  if (!rangeObj && typeof spec === 'string' && spec.includes('!')) {
    const m = spec.match(/^'?([^'!]+)'?\!(.+)$/);
    if (m) {
      const sh = ss.getSheetByName(m[1]);
      if (sh) rangeObj = sh.getRange(m[2]);
    }
  }

  if (!rangeObj) {
    console.error(`[LIST_ERROR] Could not find source: ${spec}`);
    return [];
  }

  const seen = new Set(), out = [];
  const vals = rangeObj.getValues(); 

  for (let i = 0; i < vals.length; i++) {
    let typeIdRaw = String(vals[i][0]).trim(); 
    let itemName = vals[i][1] ? String(vals[i][1]).trim() : "Unknown Item";

    if (String(typeIdRaw).includes("#N/A") || 
        String(typeIdRaw).includes("Check Name") || 
        String(typeIdRaw).includes("#VALUE!") || 
        String(typeIdRaw).includes("#REF!")) {
      continue; 
    }

    if (!typeIdRaw.length) continue; 

    let val;
    if (opts.numeric) {
      let n = Number(typeIdRaw);
      if (!Number.isFinite(n)) {
        console.warn(`[LIST_WARN] Non-numeric ID for "${itemName}": Found "${typeIdRaw}".`);
        continue;
      }
      if (opts.integer) n = Math.floor(n);
      if (opts.dropZeros && n === 0) continue;
      val = n;
    } else {
      val = typeIdRaw;
    }

    if (opts.validator && !opts.validator(val)) continue;
    if (!opts.dedupe || !seen.has(val)) {
      if (opts.dedupe) seen.add(val);
      out.push(val);
    }
  }
  return out;
}

// FAST BLOCK UPDATER
function upsertRegionCache_(sheet, rows) {
  if (!rows.length) return;

  const lastRow = sheet.getLastRow();
  const numCols = rows[0].length;

  if (lastRow <= 1) {
    sheet.getRange(2, 1, rows.length, numCols).setValues(rows);
    return;
  }

  const lookupKeys = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = new Map();

  for (let r = 0; r < lookupKeys.length; r++) {
    const t = lookupKeys[r][0];
    const g = lookupKeys[r][1];
    if (t === '' || g === '') continue;
    map.set(`${t}:${g}`, r + 2);
  }

  const appends = [];
  const updates = [];

  for (const row of rows) {
    const targetSheetRow = map.get(`${row[0]}:${row[1]}`);
    if (targetSheetRow !== undefined) {
      updates.push({ rn: targetSheetRow, data: row });
    } else {
      appends.push(row);
    }
  }

  if (updates.length > 0) {
    updates.sort((a, b) => a.rn - b.rn);

    let blockStart = updates[0].rn;
    let blockData = [updates[0].data];

    for (let i = 1; i < updates.length; i++) {
      const prevRn = updates[i - 1].rn;
      const currRn = updates[i].rn;

      if (currRn === prevRn + 1) {
        blockData.push(updates[i].data);
      } else {
        sheet.getRange(blockStart, 1, blockData.length, numCols).setValues(blockData);
        blockStart = currRn;
        blockData = [updates[i].data];
      }
    }
    sheet.getRange(blockStart, 1, blockData.length, numCols).setValues(blockData);
  }

  if (appends.length > 0) {
    sheet.getRange(lastRow + 1, 1, appends.length, numCols).setValues(appends);
  }
}

function _pairKey_(t, r) { return `${Math.floor(t)}:${Math.floor(r)}`; }

function getConfigPairSet_() {
  const items = readListFlex_(ITEMS_SOURCE, { numeric: true, integer: true });
  const regions = readListFlex_(REGIONS_SOURCE, { numeric: true, integer: true, dropZeros: true, validator: isValidRegionId_ });
  const set = new Set();
  for (const t of items) for (const r of regions) set.add(_pairKey_(t, r));
  return set;
}

function pruneCacheToConfig_(mode = 'tombstone') {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_REGION_CACHE);
  if (!sh) throw new Error('Missing ' + SHEET_REGION_CACHE);

  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return;

  const h = vals[0];
  const iT = h.indexOf('type_id');
  const iR = h.indexOf('region_id');
  const iV = h.indexOf('volume30_region');
  const iVel = h.indexOf('velocity_region');
  const iU = h.indexOf('last_updated');
  const iS = h.indexOf('status');

  const keepPairs = getConfigPairSet_();

  if (mode === 'hard') {
    const out = [h];
    for (let i = 1; i < vals.length; i++) {
      const row = vals[i];
      if (keepPairs.has(_pairKey_(row[iT], row[iR]))) out.push(row);
    }
    sh.clear();
    sh.getRange(1, 1, out.length, out[0].length).setValues(out);
    sh.setFrozenRows(1);
    return;
  }

  const updates = [];
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const keep = keepPairs.has(_pairKey_(row[iT], row[iR]));
    if (!keep) {
      row[iV] = "";
      if (iVel > -1) row[iVel] = "";
      if (iU > -1) row[iU] = new Date();
      row[iS] = "REMOVED";
      updates.push({ rn: i + 1, row });
    }
  }
  updates.forEach(u => sh.getRange(u.rn, 1, 1, h.length).setValues([u.row]));
}

function ESI_installClientInterfaceRefresh(everyMinutes) {
  const allowed = [1, 5, 10, 15, 30];
  const n = Number(everyMinutes || 5);
  if (!allowed.includes(n)) throw new Error('everyMinutes must be 1, 5, 10, 15, or 30.');
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'ESI_publishVolumeInterfaces')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('ESI_publishVolumeInterfaces').timeBased().everyMinutes(n).create();
}

function ESI_stopClientInterfaceRefresh() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'ESI_publishVolumeInterfaces')
    .forEach(t => ScriptApp.deleteTrigger(t));
}