// FuzzworkSubset_Filter.gs.js
// Build a *small* subset from https://market.fuzzwork.co.uk/aggregatecsv.csv.gz
// so Sheets doesn't choke. Filters by scope (station/system/region) + type_ids + side.
// It *does not* write the full CSV to a sheet. It fetches, filters, writes only matches.
// Also caches the raw CSV to Drive for offline fallback.

/* global UrlFetchApp, Utilities, CacheService, SpreadsheetApp, DriveApp, Session */

// ----------------------------- Config ----------------------------------------
const FUZZ_AGG_URL        = "https://market.fuzzwork.co.uk/aggregatecsv.csv.gz";
const FUZZ_CACHE_DIR      = "FuzzworkCache";         // Drive folder to keep raw CSV copies
const FUZZ_SUBSET_SHEET   = "Fuzz_Agg_Subset";       // filtered, tiny table for lookups
const FUZZ_CFG_SHEET      = "Fuzz_Config";           // optional config helper sheet
const WRITE_CHUNK_ROWS    = 5000;                     // chunk size for writing to sheet

// ----------------------------- Public entrypoints -----------------------------
/**
 * Read config from Fuzz_Config and build the subset.
 * Fuzz_Config headers (row 1): scope_field | scope_id | typeid_range_a1 | include_buy | include_sell
 * Examples: scope_field = station|system|region, scope_id = 60003760, typeid_range_a1 = 'Market Orders'!D8:D
 */
function refreshFuzzSubsetFromConfig() {
  withDocLock_(function(){
    const cfg = readFuzzConfig_();
    const csv = getAggregatesCsvTextPreferOnline_();
    const table = filterAggregatesCsv_(csv, cfg);
    writeTableChunked_(FUZZ_SUBSET_SHEET, table, WRITE_CHUNK_ROWS);
    note_(FUZZ_SUBSET_SHEET, `subset rows=${table.length-1} scope=${cfg.scope_field}:${cfg.scope_id} types=${cfg.typeIds.size} buy=${cfg.include_buy} sell=${cfg.include_sell}`);
  });
}

/** Same as above, but force offline (use latest cached CSV in Drive). */
function refreshFuzzSubsetFromCache() {
  withDocLock_(function(){
    const cfg = readFuzzConfig_();
    const csv = loadLatestCachedCsvText_();
    const table = filterAggregatesCsv_(csv, cfg);
    writeTableChunked_(FUZZ_SUBSET_SHEET, table, WRITE_CHUNK_ROWS);
    note_(FUZZ_SUBSET_SHEET, `subset (cached) rows=${table.length-1} scope=${cfg.scope_field}:${cfg.scope_id} types=${cfg.typeIds.size} buy=${cfg.include_buy} sell=${cfg.include_sell}`);
  });
}

/** Convenience: direct call with params (no config sheet needed). */
function refreshFuzzSubset(scope_field, scope_id, typeid_range_a1, include_buy, include_sell) {
  withDocLock_(function(){
    const cfg = normalizeConfig_({scope_field, scope_id, typeid_range_a1, include_buy, include_sell});
    const csv = getAggregatesCsvTextPreferOnline_();
    const table = filterAggregatesCsv_(csv, cfg);
    writeTableChunked_(FUZZ_SUBSET_SHEET, table, WRITE_CHUNK_ROWS);
    note_(FUZZ_SUBSET_SHEET, `subset rows=${table.length-1} scope=${cfg.scope_field}:${cfg.scope_id} types=${cfg.typeIds.size} buy=${cfg.include_buy} sell=${cfg.include_sell}`);
  });
}

// ----------------------------- Core logic -------------------------------------
function filterAggregatesCsv_(csvText, cfg){
  if (!csvText) throw new Error("Empty CSV text");
  const table = Utilities.parseCsv(csvText); // full parse is OK since we *don't* write full table
  if (!table || table.length === 0) throw new Error("Parsed CSV is empty");

  // Build header index map (case-insensitive, normalize underscores/etc.)
  const headers = table[0].map(h => String(h||"").trim());
  const idx = indexMap_(headers);

  // Validate presence of columns we need
  const scopeCol = idx[cfg.scope_field];
  if (scopeCol == null) throw new Error(`CSV missing scope column: ${cfg.scope_field}`);

  const typeCol = idx.type_id ?? idx.type ?? idx.typeid;
  if (typeCol == null) throw new Error("CSV missing type id column");

  const isBuyCol = idx.is_buy ?? idx.isbuy ?? idx.buy;
  if (isBuyCol == null) throw new Error("CSV missing is_buy column");

  // Keep only needed columns to shrink result (project a practical set)
  const keepCols = pickExisting_(idx, [
    cfg.scope_field,            // region/system/station
    'type_id','type','typeid',  // whichever exists
    'is_buy','isbuy','buy',
    'median','fivepercent','weightedaverage','min','max','volume','numorders','orderset'
  ]);

  const out = [ keepCols.map(c => headers[c]) ]; // header row

  // Build typeId set for fast membership (if provided)
  const hasFilterTypes = cfg.typeIds && cfg.typeIds.size > 0;

  for (let r = 1; r < table.length; r++) {
    const row = table[r];
    if (!row || row.length === 0) continue;

    // Scope match
    const scopeVal = asInt_(row[scopeCol]);
    if (scopeVal !== cfg.scope_id) continue;

    // Type match
    const typeVal = asInt_(row[typeCol]);
    if (hasFilterTypes && !cfg.typeIds.has(typeVal)) continue;

    // Side match
    const isBuy = asBool_(row[isBuyCol]);
    if (isBuy && !cfg.include_buy) continue;
    if (!isBuy && !cfg.include_sell) continue;

    // Project selected columns
    const proj = new Array(keepCols.length);
    for (let i = 0; i < keepCols.length; i++) proj[i] = row[keepCols[i]];
    out.push(proj);
  }

  // If no matches, keep just header with a note row to avoid formula errors
  if (out.length === 1) out.push(["No matches for current filter"]);
  return out;
}

// ----------------------------- Fetch/cache ------------------------------------
function getAggregatesCsvTextPreferOnline_(){
  try {
    const resp = UrlFetchApp.fetch(FUZZ_AGG_URL, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true });
    const code = resp.getResponseCode();
    if (code !== 200) throw new Error("HTTP "+code);
    const blob = Utilities.ungzip(resp.getBlob());
    const text = blob.getDataAsString("UTF-8");
    // Save to Drive as timestamped CSV for offline
    const fname = "aggregatecsv_" + nowStr_() + ".csv";
    getOrCreateFolder_(FUZZ_CACHE_DIR).createFile(blob.setName(fname));
    return text;
  } catch(e) {
    // Fallback to cached file
    return loadLatestCachedCsvText_();
  }
}

function loadLatestCachedCsvText_(){
  const folder = getOrCreateFolder_(FUZZ_CACHE_DIR);
  const files = folder.getFiles();
  let latest = null;
  while (files.hasNext()) {
    const f = files.next();
    if (!latest || f.getDateCreated() > latest.getDateCreated()) latest = f;
  }
  if (!latest) throw new Error("No cached Fuzz aggregates in Drive folder: "+FUZZ_CACHE_DIR);
  return latest.getBlob().getDataAsString("UTF-8");
}

// ----------------------------- Config helpers ---------------------------------
function readFuzzConfig_(){
  // Create if missing
  const sh = getOrCreateSheet_(FUZZ_CFG_SHEET);
  if (sh.getLastRow() < 1) sh.appendRow(["scope_field","scope_id","typeid_range_a1","include_buy","include_sell"]);
  if (sh.getLastRow() < 2) sh.appendRow(["station", 60003760, "'Market Orders'!D8:D", true, true]);

  const vals = sh.getRange(2,1,1,5).getValues()[0];
  const raw = { scope_field: vals[0], scope_id: vals[1], typeid_range_a1: vals[2], include_buy: vals[3], include_sell: vals[4] };
  return normalizeConfig_(raw);
}

function normalizeConfig_(raw){
  const scope_field = String(raw.scope_field||"").toLowerCase().trim();
  const scope_id    = asInt_(raw.scope_id);
  if (!scope_field || !/^(station|system|region)$/.test(scope_field)) throw new Error("scope_field must be station|system|region");
  if (!Number.isFinite(scope_id)) throw new Error("scope_id must be a number");

  // Resolve type ids from a range (string like "'Market Orders'!D8:D") or array
  let typeIds = new Set();
  const a1 = raw.typeid_range_a1;
  if (a1) {
    try {
      const sh = SpreadsheetApp.getActive().getRange(String(a1)).getValues();
      for (let i=0;i<sh.length;i++) {
        const v = asInt_(sh[i][0]);
        if (Number.isFinite(v)) typeIds.add(v);
      }
    } catch(e) {
      // if not a valid range, ignore
    }
  }

  const include_buy  = !!raw.include_buy;
  const include_sell = raw.include_sell === undefined ? true : !!raw.include_sell;

  return { scope_field, scope_id, typeIds, include_buy, include_sell };
}

// ----------------------------- Sheet helpers ----------------------------------


function writeTableChunked_(sheetName, table, chunkRows) {
  const sh = getOrCreateSheet_(sheetName);
  sh.clearContents();
  if (!table || !table.length) return;
  const cols = Math.max(1, ...table.map(r => r.length));
  const norm = table.map(r => { const x = r.slice(); if (x.length < cols) x.length = cols; return x; });

  let start = 0;
  while (start < norm.length) {
    const end = Math.min(start + chunkRows, norm.length);
    const slice = norm.slice(start, end);
    sh.getRange(start + 1, 1, slice.length, cols).setValues(slice);
    start = end;
  }
  if (norm.length > 1) sh.setFrozenRows(1);
}

function note_(sheetName, text){
  const sh = getOrCreateSheet_(sheetName);
  const cols = sh.getLastColumn() || 1;
  sh.getRange(1, cols).setNote(text + "\n" + new Date().toISOString());
}

// ----------------------------- Small utils ------------------------------------
function indexMap_(headers){
  // normalize: lowercase, remove non-alnum/underscore
  const norm = h => String(h||"").toLowerCase().replace(/[^a-z0-9_]+/g, "");
  const map = Object.create(null);
  for (let i=0;i<headers.length;i++) {
    const key = norm(headers[i]);
    if (!key) continue;
    map[key] = i;
  }
  // Provide friendly aliases
  const alias = (a,b) => { if (map[a] != null && map[b] == null) map[b] = map[a]; };
  alias('type','type_id'); alias('typeid','type_id');
  alias('isbuy','is_buy'); alias('buy','is_buy');
  alias('regionid','region'); alias('systemid','system'); alias('stationid','station');
  return map;
}

function asInt_(v){
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : NaN;
}
function asBool_(v){
  const s = String(v).toLowerCase();
  return s === '1' || s === 'true' || s === 't' || s === 'yes';
}
function getOrCreateFolder_(name){
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function withDocLock_(fn){
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function nowStr_(){
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
}
