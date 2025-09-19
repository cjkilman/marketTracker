
//JITA SELL

function testfuzAPI()
{
  let ids = [
    16239,
16243,
24030,
32881,
17366,
16273,
34206,
34202,
34203,
34205,
34204,
34201,
19761,
42695,
42830
  ];
  return fuzzApiPriceDataJitaSell(ids);
}

/**
* Generic API function to get a list of minimal prices for an array of type_id's
* @param {range} A vertical range of type_ids.
* @param {market_hub_id} market region ID, Defaults to Jita:60003760
* @param {order_type} sell or buy
* @param {order_level} min,max,average,mean
* @return minSell for each type_id. This can be configured differently.
* @customfunction
* Author: unknown
* Modified by CJ Kilman 11/19/2021 Added Configuration options and a Safe request buffer to avoide overloading the service get requesgt method
*/
function fuzzApiPriceDataJitaSell(type_ids, market_hub = 60003760, order_type = null ,order_level=null)
{
    if (!type_ids) throw 'type_ids is required';
    if(!Array.isArray(type_ids)) type_ids = [type_ids];
    type_ids = type_ids.filter(Number) ;

    if(order_type == null &&  order_level == null){
      order_type = "sell";
      order_level = "min";
    }

    if(order_type == null &&  order_level.toLowerCase() == "max") order_type = "buy";
    if(order_type == null &&  order_level.toLowerCase() == "min") order_type= "sell";
    if(order_type.toLowerCase() == "buy" &&  order_level == null) order_level = "max";
    if(order_type.toLowerCase() == "sell" &&  order_level == null) order_level = "min";
    order_type = order_type.toLowerCase(); order_level = order_level.toLowerCase();

      let price_data;
      var result=[];
  

  // Capture overflow buffer

      price_data = postFetch(type_ids,market_hub,"station")
      for(var i=0 ; i < type_ids.length ; i++)
      {
        try
        {
          let value = parseFloat(price_data[type_ids[i]][order_type][order_level]);
          if(!isNaN(value))
            result = result.concat(value);
          else
            result = result.concat("");
        }
        catch(error) // Value not on market, Leave Blank cell
        {
          result = result.concat("");
        }
      }
    
    
    return result;
  }

/**
* Fuzz market API for the given types
*
* @param {range} range A vertical range of type_ids.
* @param {string} string Jita, Amarr, Dodixie, Rens, Hek, Defaults to Jita.
* @param {string} string sell or buy. Defaults to sell.
* @param {string} string min, max, or avg. Defaults to min.
* @return result for each type_id. This can be configured differently.
* @customfunction
* Author: unknown
* Modified by CJ Kilman 11/19/2021 Added configuration options and a Safe request buffer to avoid overloading the service get request method
* Modified by Snowdevil / Highfly Chastot 12/16/2021 Added functionality for choosing hub, type, and level. Little refactoring, could use more.
*/
function fuzzPriceDataByHub(type_ids, market_hub = "Jita", order_type = "sell", order_level = null) 
{
    // Safety net
    if (!type_ids) throw 'type_ids is required';
    // Select hub ID, can ONLY use major trade hubs with this API
try{
  market_hub  = market_hub.toLowerCase();
}catch{}
    switch (market_hub) {
    case 'amarr':
        market_hub = 60008494;
        break;
    case 'dodixie':
        market_hub = 60011866;
        break;
    case 'rens':
        market_hub = 60004588;
        break;
    case 'hek':
        market_hub = 60005686;
        break; 
     
    case 'jita':
        market_hub = 60003760;
        break;
    default:
     
     
    }
    
    //deal with defaults on most used order types

      if(order_level==null)
      {
        switch(order_type.toLowerCase()){
    
        case 'buy':
              order_level = "max";
              break;
        case 'sell':
        default:
          order_level= "min";
            }
          }
    // result
    return fuzzApiPriceDataJitaSell(type_ids,market_hub,order_type,order_level);
  

  }

  /**
 * Get stats for type_ids from Fuzzworks-like payload.
 * - Relies on postFetch(...) to handle caching.
 * - Supports: buy/sell × min|max|avg|median|volume
 * - Returns array aligned to input type_ids; "" where unavailable.
 */
/**
 * Get stats for type_ids from Fuzzworks payload.
 * - Cache-first: read cache; only fetch missing ids (via postFetch).
 * - Supports: buy/sell × min|max|avg|median|volume
 * - Returns array aligned with input order; "" where unavailable.
 */
function marketStatData(type_ids, location_type, location_id, order_type, order_level) {
  if (!type_ids) throw new Error("type_ids is required");
  const input = Array.isArray(type_ids) ? type_ids : [type_ids];

  // normalize ids but keep placeholders so we can re-align later
  const ids = input.map(v => {
    const n = Number(v);
    return Number.isFinite(n) ? n : null;
  });

  // location guard
  const lt = String(location_type || "").toLowerCase();
  if (!["region","system","station"].includes(lt)) {
    throw new Error("Location Undefined");
  }

  // normalize fields
    const { type: side, level: lvl } = _normalizeOrder(order_type, order_level);

  // unique valid ids
  const uniq = Array.from(new Set(ids.filter(n => n != null)));

  // 1) cache-first
  const { have, missing } = _getCachedFuz(uniq, location_id, lt);

  // 2) fetch only truly missing ids, then merge
  if (missing.length) {
    const fetched = postFetch(missing, location_id, lt) || {};
    // merge into have
    for (const id of missing) {
      if (fetched[id] != null) have[id] = fetched[id];
    }
  }

  // 3) picker strictly for Fuzzworks fields
  function pick(row) {
    if (!row || !row[side]) return null;
    const node = row[side];
    const v = node[lvl];            // min|max|avg|median|volume
    const num = Number(v);
    return Number.isFinite(num) ? num : null;
  }

  // 4) map back to original order
  const out = ids.map(id => (id == null ? "" : (pick(have[id]) ?? "")));

  return Array.isArray(type_ids) ? out : out[0];
}


/**
 * Normalize order_type/order_level for Fuzzworks API.
 * Defaults: sell/min. Accepts common synonyms.
 */
function _normalizeOrder(order_type, order_level) {
  let type  = order_type ? String(order_type).toLowerCase() : null;
  let level = order_level ? String(order_level).toLowerCase() : null;

  // synonyms
  if (type === "bid") type = "buy";
  if (type === "ask") type = "sell";
  const levelAliases = { mean: "avg", average: "avg", med: "median", vol: "volume", qty: "volume", quantity: "volume" };
  if (level && levelAliases[level]) level = levelAliases[level];

  // defaults
  if (!type && !level)        { type = "sell"; level = "min"; }
  else if (!type && level)    { type = (level === "max") ? "buy" : "sell"; }
  else if (type && !level)    { level = (type === "buy") ? "max" : "min"; }

  const validTypes  = ["buy","sell"];
  const validLevels = ["min","max","avg","median","volume"];
  if (!validTypes.includes(type))  throw new Error("order_type must be 'buy' or 'sell'");
  if (!validLevels.includes(level)) throw new Error("order_level must be one of 'min','max','avg','median','volume'");
  return { type, level };
}


/** Small helper: run fn while holding a document lock */
function withDocLock(fn, ms = 30000) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(ms);
  try { return fn(); }
  finally { lock.releaseLock(); }
}

/**
 * Pull Price data from Fuzzworks using POST request, with safe, batched cache writes.
 * - Caches per (location_type, location_id, type_id)
 * - Returns an object keyed by type_id containing the full buy/sell aggregate
 */
function postFetch(type_ids, location_id, location_type = "station") {
  if (!type_ids) throw new Error('type_ids is required');
  if (!Array.isArray(type_ids)) type_ids = [type_ids];

  // normalize + dedupe, but keep an index for final merge
  const ids = type_ids.map(Number).filter(Number.isFinite);
  if (!ids.length) return {};
  const uniq = Array.from(new Set(ids));

  const lt = String(location_type).toLowerCase();
  if (!["region","system","station"].includes(lt)) {
    throw new Error("Invalid location_type; use 'region', 'system', or 'station'");
  }

  const cache = CacheService.getDocumentCache();
  const ttlSec = 30 * 60; // 30 min

  // 1) First cache read (no lock)
const keys = uniq.map(id => _fuzKey(lt, location_id, id));
  const got  = cache.getAll(keys);

  const result = {};
  const missing = [];

  uniq.forEach((id, i) => {
    const k = keys[i];
    const raw = got[k];
    if (raw) {
      try { result[id] = JSON.parse(raw); }
      catch (e) { missing.push(id); }
    } else {
      missing.push(id);
    }
  });

  // 2) If anything missing, fetch them in one POST (still no lock)
  let fetched = {};
  if (missing.length) {
    const url = "https://market.fuzzwork.co.uk/aggregates/";
    const payload = {};
    payload[lt] = location_id;
    payload.types = missing.join(",");
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
   let resp;
  try {
    resp = UrlFetchApp.fetch(url, options);
  } catch (e) {
    throw new Error(`Fuzzworks fetch failed: ${e && e.message ? e.message : e}`);
  }
    const code = resp.getResponseCode();
    if (code !== 200) {
      throw new Error(`Fuzzworks error ${code}: ${resp.getContentText().slice(0,200)}`);
    }
    fetched = JSON.parse(resp.getContentText() || "{}");
  }

  if (missing.length) {
    // 3) Short critical section: recheck + write cache + finalize result
    withDocLock(() => {
      // Re-check: maybe another runner already cached some of these
      const missKeys = missing.map(id => `fuz:${lt}:${location_id}:${id}`);
      const nowGot = cache.getAll(missKeys);

      // Build a putAll payload only for still-missing
      const toPut = {};
      missing.forEach((id, i) => {
        const k = missKeys[i];
        if (!nowGot[k]) {
          const row = fetched[id];
          if (row) {
            toPut[k] = JSON.stringify(row);
            result[id] = row; // also into our return map
          }
        } else {
          // another runner filled it while we fetched
          try { result[id] = JSON.parse(nowGot[k]); }
          catch (e) { /* ignore; if parse fails we simply skip caching */ }
        }
      });

      // Chunked cache writes
      const entries = Object.entries(toPut);
      const CHUNK = 80; // conservative
      for (let i = 0; i < entries.length; i += CHUNK) {
        const slice = Object.fromEntries(entries.slice(i, i + CHUNK));
        if (Object.keys(slice).length) cache.putAll(slice, ttlSec);
      }
    });
  }

  // 4) Return only the requested ids, in object form keyed by id
  return result;
}

  /** Build the per-id cache key used by postFetch */
function _fuzKey(location_type, location_id, type_id) {
  return `fuz:${location_type}:${location_id}:${type_id}`;
}

/** Try to read a batch from the document cache */
function _getCachedFuz(type_ids, location_id, location_type) {
  const cache = CacheService.getDocumentCache();
  const keys = type_ids.map(id => _fuzKey(location_type, location_id, id));
  const raw = cache.getAll(keys);

  const have = {};
  const missing = [];
  for (let i = 0; i < type_ids.length; i++) {
    const id = type_ids[i];
    const k = keys[i];
    const s = raw[k];
    if (s) {
      try {
        have[id] = JSON.parse(s);
      } catch (e) {
        // bad JSON in cache? treat as missing
        missing.push(id);
      }
    } else {
      missing.push(id);
    }
  }
  return { have, missing };
}

