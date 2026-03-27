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


/** ---------- Rate limit + backoff (lightweight) ---------- **/
const PACER_PROP_GAP = 'esi_gap_ms';
const PACER_PROP_LAST = 'esi_last_ms';
const PACER_MIN_MS = 400;   // start ~2.5 rps
const PACER_MAX_MS = 5000;  // cap 5s between calls

/***** GESI batching + governor (≤300 req/min) ********************************/

// Hard cap
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

// ----- Blacksmoke/GESI client -----
function getGESIHistoryClient_() {
  // Requires GESI library + authorized character (Sheets Add-ons → GESI → Authorize)
  return GESI.getClient().setFunction('markets_region_history');
}

function isBeforeNoonET_() {
  return +Utilities.formatDate(new Date(), 'America/New_York', 'H') < 12;
}
function getBatchSize_() {      // AM: 200, PM: 60
  return isBeforeNoonET_() ? 200 : 60;
}



function buildHistoryRequests_(client, regionId, typeIds) {
  return typeIds.map(function (tid) {
    return client.buildRequest({
      region_id: regionId,
      type_id: tid,
      show_column_headings: false,
      version: 'latest'
    });
  });
}

function summarizeHistoryMulti_(rows) {
  if (!rows || !rows.length) return { vol30: 0, vel30: 0, vol7: 0, vel7: 0, vol5: 0, vel5: 0 };
  
  var v30 = 0, v7 = 0, v5 = 0;
  // ESI returns chronological. Count backwards from the newest day.
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

function mapGESIRespToResult_(resp, typeId) {
  var code = resp.getResponseCode();
  if (code === 200) {
    try {
      var arr = JSON.parse(resp.getContentText()) || [];
      var m = summarizeHistoryMulti_(arr);
      return { 
        typeId: typeId, status: 'OK', 
        vol30: m.vol30, vel30: m.vel30,
        vol7: m.vol7, vel7: m.vel7,
        vol5: m.vol5, vel5: m.vel5
      };
    } catch (e) {
      return { typeId: typeId, status: 'ERR_PARSE' };
    }
  }
  if (code === 401 || code === 403) return { typeId: typeId, status: 'ERR_AUTH', code: code };
  if (code === 420 || code === 429) return { typeId: typeId, status: 'ERR_RATE', code: code };
  if (code >= 500) return { typeId: typeId, status: 'ERR_5XX', code: code };
  if (code >= 400) return { typeId: typeId, status: 'ERR_4XX', code: code };
  return { typeId: typeId, status: 'ERR_' + code, code: code };
}

function fetchHistoryBatchGESI_(regionId, typeIds) {
  var client = getGESIHistoryClient_();
  var out = [];
  
  // Start with the configured max size (e.g., 50)
  // If we hit errors, we will lower this variable dynamically for this run.
  var currentBatchSize = ESI_GROUP_SIZE; 
  var i = 0;

  while (i < typeIds.length) {
    // 1. Slice the next batch based on current dynamic size
    var end = Math.min(i + currentBatchSize, typeIds.length);
    var ids = typeIds.slice(i, end);

    // 2. Pay the rate-limit cost (Token Bucket)
    var wait = tbConsume_(ids.length);
    if (wait > 0) Utilities.sleep(Math.min(wait, 30000));

    var reqs = buildHistoryRequests_(client, regionId, ids);
    
    try {
      // 3. Attempt the fetch
      var resps = UrlFetchApp.fetchAll(reqs);

      // --- Success Path ---
      
      // Log error budget (optional, using first response)
      try {
        if (resps.length > 0) {
          var hdr = resps[0].getAllHeaders && resps[0].getAllHeaders();
          var remain = +(hdr && (hdr['x-esi-error-limit-remain'] || hdr['X-Esi-Error-Limit-Remain']) || -1);
          if (remain >= 0) (LoggerEx?.info || console.info)('esi.errbudget', { remain: remain });
        }
      } catch (_) { }

      // Process results
      for (var k = 0; k < resps.length; k++) {
        out.push(mapGESIRespToResult_(resps[k], ids[k]));
      }

      // Advance the cursor by the ACTUAL number of items processed
      i += ids.length; 
      
      // Optional: If we are running successfully at a low batch size, 
      // you could slowly increment currentBatchSize here, but usually 
      // it's safer to stay low once a run has proven unstable.
      
      Utilities.sleep(ESI_MICRO_BREATH_MS);

    } catch (e) {
      // --- Error Path ---
      var msg = String(e.message || "");
      var isBandwidth = msg.includes("Bandwidth") || msg.includes("quota") || msg.includes("limit");

      // If it's a bandwidth error and we can still reduce the size:
      if (isBandwidth && currentBatchSize > 1) {
        var oldSize = currentBatchSize;
        // Dynamic Adjustment: Cut batch size in half
        currentBatchSize = Math.max(1, Math.floor(currentBatchSize / 2));
        
        (LoggerEx?.warn || console.warn)('esi.bandwidth.adapt', {
           msg: "Bandwidth limit hit. Reducing batch size and retrying.",
           oldSize: oldSize,
           newSize: currentBatchSize,
           index: i
        });
        
        // Wait a moment for the quota bucket to drain
        Utilities.sleep(3000);
        
        // LOOP REPEATS: We do NOT increment 'i', so we retry the same items with the smaller batch.
      } else {
        // Fatal error: Not bandwidth related, OR we are already at batch size 1
        throw e;
      }
    }
  }
  
  return out;
}



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
  const client = GESI.getClient().setFunction('markets_region_history');
  const resp = UrlFetchApp.fetch(client.buildRequest({
    region_id: rid,
    type_id: tid,
    show_column_headings: false,
    version: 'latest'
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

  // publishMarketResultESIRegion(); // <-- REFACTORED: Removed.
  // This is too "heavy" to call here. The `markAllCacheStale_` call
  // is sufficient. Publishing will be handled by the ESI_publishClientInterfaces
  // timer and the stopWorker_() function.

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
    var BS = (typeof getBatchSize_ === 'function') ? getBatchSize_() : BATCH_SIZE;
    var end = Math.min(ii + BS, items.length);
    LoggerEx.info('mf.batch.params', { BS: BS, cadence_min: WORKER_EVERY_MIN });

    var chunk = items.slice(ii, end);

    LoggerEx.log('mf.run.start rid=' + rid + ' regions=' + regions.length + ' items=' + items.length +
      ' cursorIn=' + JSON.stringify({ ri: ri, ii: ii }) +
      ' region=' + regionId + ' window=' + JSON.stringify({ from: ii, to: end, size: chunk.length }));

    // --- REFACTORED: Removed duplicate variable declarations ---
    var rowsOut = [];
    var processed = 0;
    var hitRate = false;
    var ok = 0, err = 0, errAuth = 0;
    // --- END REFACTOR ---

    // --- GESI batch (auth handled by GESI, paced by token bucket) ---
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


    // Flush once
    if (rowsOut.length) {
      upsertRegionCache_(cache, rowsOut);
      
      // publishMarketResultESIRegion(); // <-- REFACTORED: Removed.
      // This function is too "heavy" to run inside the 1-minute worker loop.
      // It will be called by stopWorker_() when the job is done, or by
      // the ESI_publishClientInterfaces timer.
      
      LoggerEx.log('mf.flush.chunk rid=' + rid + ' rows=' + rowsOut.length);
    } else {
      LoggerEx.log('mf.flush.empty rid=' + rid);
    }

    // Auth fail-fast: stop job, preserve cursor. No recycling.
    if (errAuth > 0) {
      var props = PropertiesService.getScriptProperties();
      props.setProperty('mf_blocked_auth', '1'); // breadcrumb if you want to gate startup
      props.deleteProperty('mf_job_active');
      LoggerEx.error('mf.auth.block rid=' + rid + ' region=' + regionId + ' errAuth=' + errAuth + ' hint="Sheets → Add-ons → GESI → Authorize Character"');
      return stopWorker_('auth required');
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
    LoggerEx.error('mf.run.exception rid=' + rid + ' msg=' + (e && e.message));
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
 * Fetches any item that is missing from the Cache.
 */
function fillMissingEsiHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cacheSh = ss.getSheetByName('Cache_Market_ESI_Region');
  
  // 1. Get Expected IDs (From Source)
  const expectedIds = readListFlex_(ITEMS_SOURCE, { numeric: true });
  
  // 2. Get Existing IDs (From Cache)
  const existingIds = new Set();
  if (cacheSh && cacheSh.getLastRow() > 1) {
    // Cache has data: [type_id, region_id, ...]
    // We only care about type_id (Col A) for now, or you can check pair key.
    const data = cacheSh.getRange(2, 1, cacheSh.getLastRow()-1, 1).getValues();
    data.forEach(r => existingIds.add(Number(r[0])));
  }

  // 3. Find Gaps
  const missing = expectedIds.filter(id => !existingIds.has(id));

  if (missing.length === 0) {
    console.log("[GapFiller] Cache is complete. No missing items.");
    return;
  }

  console.log(`[GapFiller] Found ${missing.length} missing items. Fetching...`);

  // 4. Manual Fetch Loop (Reusing your internal logic)
  // We target your default regions from 'Market Settings'
  const regions = readListFlex_(REGIONS_SOURCE, { 
    numeric: true, integer: true, dropZeros: true, validator: isValidRegionId_ 
  });
  
  regions.forEach(rid => {
    // Reuse your existing batch fetcher
    const results = fetchHistoryBatchGESI_(rid, missing);
    
    // Format for Cache Upsert
    const rowsOut = [];
    results.forEach(r => {
      if (r.status === 'OK') {
        rowsOut.push([r.typeId, rid, r.vol30, r.vel, new Date(), 'OK']);
      } else {
        rowsOut.push([r.typeId, rid, "", "", new Date(), r.status]);
      }
    });

    // Save to Cache
    if (rowsOut.length) {
      upsertRegionCache_(ensureCacheMarketESIRegion_(ss), rowsOut);
      console.log(`[GapFiller] Saved ${rowsOut.length} items for Region ${rid}`);
    }
  });

  // 5. Update Publish Sheet immediately
  publishMarketResultESIRegion();
  console.log("[GapFiller] Complete. Cache & Publish updated.");
}


/* =============================== HELPERS ================================ */
function publishMarketResultESIRegion(ss) {
  if(!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = ss.getSheetByName(SHEET_REGION_CACHE);
  const pub = ss.getSheetByName(SHEET_PUBLISH);
  if (!cache || !pub) throw new Error('Missing sheets');

  const keepPairs = getConfigPairSet_();               // ← NEW

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
  SpreadsheetApp.getActive().setNamedRange(NR_MARKET_RESULT, pub.getRange('A:E'));
}

/**
 * Build a values-only table per client from Publish_ESI_Region.
 * Input:  Sheet "Interface Clients" → A:Client, B:region_id_heartbeat
 * Source: Sheet "Publish_ESI_Region" (A:E = type_id,region_id,volume30_region,last_updated,status)
 * Output: Sheet "Publish_ESI_Region_<client>" with headers:
 *         [type_id, volume30_region, last_updated, status]
 * Named range for client: MarketResultESI_Region_<client>
 */
function ESI_publishVolumeInterfaces() {
  const ss = SpreadsheetApp.getActive();
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

  // ---- helpers ----
  const normStatus = (s) => String(s || '').trim().toUpperCase().replace(/_/g, '-');
  const toInt = (v) => {
    const n = Number(String(v).replace(/[^\d\-]/g, ''));
    return Number.isFinite(n) ? n : NaN;
  };
  const toDate = (v) => {
    if (v instanceof Date) return v;
    const s = String(v || '').trim();
    if (!s) return null;
    // Try YYYY-MM-DD first
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (m) return new Date(Date.UTC(+m[1], +m[2]-1, +m[3]));
    // Try serial / general number
    const n = Number(s);
    if (Number.isFinite(n)) return new Date(Math.round((n - 25569) * 86400000)); // Excel serial fallback
    // Last resort: Date.parse
    const t = Date.parse(s);
    return isNaN(t) ? null : new Date(t);
  };

  // ---- group rows by region ----
  // Build region → deduped rows keyed by type_id (prefer latest last_updated)
  /** @type {Map<number, Map<number, any[]>>} */
  const byRegion = new Map();

  // src schema A..E: [type_id, region_id, vol30, last_updated, status]
  for (let r = 1; r < src.length; r++) {
    const row = src[r];
    const typeId = toInt(row[0]);
    const regionId = toInt(row[1]);
    if (!Number.isFinite(typeId) || !Number.isFinite(regionId)) continue;

    const vol30 = Number(row[2]) || 0;
    const last = toDate(row[3]);       // Date or null
    const status = normStatus(row[4]);

    if (!byRegion.has(regionId)) byRegion.set(regionId, new Map());
    const bag = byRegion.get(regionId);

    // dedupe by type_id: keep the row with the most recent last_updated
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

  // ---- read clients ----
  if (lastCliRow < 2) return;
  const cliRows = cliSh.getRange(2, 1, lastCliRow - 1, 3).getValues();

  const OUT_HEADERS = ['type_id', 'volume30_region', 'last_updated', 'status'];

  for (const [clientRaw, ridRaw, filtRaw] of cliRows) {
    const client = String(clientRaw || '').trim();
    const regionId = toInt(ridRaw);
    if (!client || !Number.isFinite(regionId)) continue;

    // Build allowed set (default OK|STALE-ESI; "ALL" skips filtering)
    const filt = String(filtRaw || 'OK|STALE-ESI').trim();
    const allowAll = normStatus(filt) === 'ALL';
    const allowed = new Set(
      allowAll
        ? []
        : filt.split(/[|,]/).map((x) => normStatus(x)).filter(Boolean)
    );

    // Pull region rows and apply status filter
    const bag = byRegion.get(regionId);
    const rows = [];
    if (bag && bag.size) {
      for (const [, rec] of bag) {
        if (!rec) continue;
        const st = normStatus(rec[3]);
        if (allowAll || allowed.has(st)) rows.push(rec);
      }
    }

    // Optional: sort by volume desc, then type_id
    rows.sort((a, b) => (b[1] - a[1]) || (a[0] - b[0]));

    // Coerce last_updated to Date objects for formatting
    for (let i = 0; i < rows.length; i++) {
      const d = rows[i][2];
      rows[i][2] = toDate(d) || '';
    }

    const outName = 'Publish_ESI_Region_' + client;
    const outSh = getOrCreateSheet(ss, outName, OUT_HEADERS);

    // Clear old data rows, keep header
    const last = outSh.getLastRow();
    if (last > 1) outSh.getRange(2, 1, last - 1, OUT_HEADERS.length).clearContent();

    if (rows.length) {
      outSh.getRange(2, 1, rows.length, OUT_HEADERS.length).setValues(rows);
      // number/date formats
      outSh.getRange(2, 1, rows.length, 1).setNumberFormat('0');              // type_id
      outSh.getRange(2, 2, rows.length, 1).setNumberFormat('#,##0');          // volume30_region
      outSh.getRange(2, 3, rows.length, 1).setNumberFormat('yyyy-mm-dd');     // last_updated
    }

    // Update named range (header + data)
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

/**
 * REVISED: List reader that catches and reports #N/A errors
 */
function readListFlex_(spec, opts) {
  opts = Object.assign({
    numeric: true,
    integer: true,
    dropZeros: true,
    dedupe: true,
    validator: null
  }, opts || {});

  const ss = SpreadsheetApp.getActive();
  
  // 1. Establish the source
  let rangeObj = (spec && typeof spec.getValues === 'function') ? spec : ss.getRangeByName(spec);
  
  if (!rangeObj && typeof spec === 'string' && spec.includes('!')) {
    const m = spec.match(/^'?([^'!]+)'?\!(.+)$/);
    if (m) {
      const sh = ss.getSheetByName(m[1]);
      if (sh) rangeObj = sh.getRange(m[2]);
    }
  }

  // 2. Safety check
  if (!rangeObj) {
    console.error(`[LIST_ERROR] Could not find source: ${spec}`);
    return [];
  }

  const seen = new Set(), out = [];
  const vals = rangeObj.getValues(); // This is the 2D array [ID, Name]

  for (let i = 0; i < vals.length; i++) {
    let typeIdRaw = String(vals[i][0]).trim(); 
    // Column B (Index 1) holds the Name for our diagnostic log
    let itemName = vals[i][1] ? String(vals[i][1]).trim() : "Unknown Item";

// --- THE SILENT GATE ---
    // If the ID is a known spreadsheet error, skip this row but keep going.
    if (String(typeIdRaw).includes("#N/A") || 
        String(typeIdRaw).includes("Check Name") || 
        String(typeIdRaw).includes("#VALUE!") || 
        String(typeIdRaw).includes("#REF!")) {
      
      continue; // <--- CORRECT: Skips this item, proceeds to the next.
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
    // Rebuild the sheet with only configured rows
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

  // tombstone: mark non-config rows as REMOVED and blank their numbers
  const updates = [];
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const keep = keepPairs.has(_pairKey_(row[iT], row[iR]));
    if (!keep) {
      row[iV] = "";               // blank numerics
      if (iVel > -1) row[iVel] = "";
      if (iU > -1) row[iU] = new Date();
      row[iS] = "REMOVED";
      updates.push({ rn: i + 1, row });
    }
  }
  updates.forEach(u => sh.getRange(u.rn, 1, 1, h.length).setValues([u.row]));
}




/**
 * Install/replace a time trigger to republish client interfaces.
 * @param {number} everyMinutes  One of: 1, 5, 10, 15, 30 (default 10)
 */
function ESI_installClientInterfaceRefresh(everyMinutes) {
  const allowed = [1, 5, 10, 15, 30];
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

