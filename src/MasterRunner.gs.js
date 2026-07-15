// --- CONFIGURATION CONSTANTS ---
// Item IDs: Assumed to be in 'Item List Back End' column A, starting at row 2.
const ITEM_ID_SOURCE = "'Item List Back End'!A2:A"; 
const MARKET_SETTINGS_SHEET = "Market Settings"; 


/**
 * --- UTILITY: READ ITEM ID LIST ---
 * Reads a single-column range, coerces values to numbers, drops non-positive/non-finite values, and dedupes.
 */
function _readItemIDList(spec) {
  const ss = SpreadsheetApp.getActive();
  let range = ss.getRangeByName(spec);

  // Fallback to getting range by A1 notation
  if (!range && typeof spec === 'string' && spec.includes('!')) {
    const parts = spec.match(/^'?([^'!]+)'?\!(.+)$/);
    if (parts) range = ss.getSheetByName(parts[1])?.getRange(parts[2]);
  }
  if (!range) {
    (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('MasterRunner') : console).error(`Source range not found: ${spec}`);
    return [];
  }

  return range.getValues()
    .flat()
    .map(v => Math.floor(Number(v)))
    .filter(n => Number.isFinite(n) && n > 0)
    .filter((n, i, a) => a.indexOf(n) === i); // Dedupe
}


/**
 * --- UTILITY: DYNAMIC MARKET ID READER ---
 * Reads Market Settings D2:F to get the market types and corresponding IDs.
 */
function _readMarketIDsFromSettings() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MARKET_SETTINGS_SHEET);
  if (!sh) {
    (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('MasterRunner') : console).error(`Market Settings sheet not found.`);
    return [];
  }

  const maxRows = sh.getLastRow();
  if (maxRows < 3) return [];
  
  // Read headers (Row 2, Columns D-F) and all data (Row 3 to end, Columns D-F)
  const headers = sh.getRange(2, 4, 1, 3).getValues()[0].map(h => String(h || '').trim().toLowerCase()); // D2:F2
  const data = sh.getRange(3, 4, maxRows - 2, 3).getValues(); // D3:F[last row]
  
  const marketPairs = [];
  
  // Iterate through data rows
  for (const row of data) {
    // Iterate through the three columns (Station, System, Region)
    for (let i = 0; i < 3; i++) {
      const market_id = Math.floor(Number(row[i]));
      const market_type = headers[i];

      if (Number.isFinite(market_id) && market_id > 0) {
        marketPairs.push({
          market_id: market_id,
          market_type: market_type
        });
      }
    }
  }

  return marketPairs;
}

/**
 * --- MASTER FUNCTION DEFINITION (Optimized for Velocity Momentum) ---
 * Builds the full matrix of {type_id, market_id, market_type} requests,
 * prioritizing high-velocity and surging items first.
 */
function getMasterMarketRequestsRegion() {
  const itemIDs = _readItemIDList(ITEM_ID_SOURCE);
  const marketPairs = _readMarketIDsFromSettings(); 

  if (itemIDs.length === 0 || marketPairs.length === 0) {
    (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('MasterRunner') : console).warn('Market request list is empty.');
    return [];
  }

  // --- MOMENTUM PRIORITIZATION LAYER ---
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cacheSh = ss.getSheetByName('Cache_Market_ESI_Region');
    
    if (cacheSh && cacheSh.getLastRow() > 1) {
      const cacheData = cacheSh.getDataRange().getValues();
      const h = cacheData[0];
      const iType = h.indexOf('type_id');
      const iVel30 = h.indexOf('velocity30_region');
      const iVel5 = h.indexOf('velocity5_region');
      const iStatus = h.indexOf('status');

      // Map to store maximum momentum score per item type
      const momentumMap = new Map();

      for (let r = 1; r < cacheData.length; r++) {
        const tId = Number(cacheData[r][iType]);
        if (!tId) continue;

        const status = String(cacheData[r][iStatus] || "");
        let score = 1.0; // Baseline entry score

        // Push dead items or persistent budget-draining errors to the absolute bottom
        if (status.includes('ERR') || status === 'NOT_FOUND') {
          score = 0.0;
        } else {
          const v30 = parseFloat(cacheData[r][iVel30]) || 0;
          const v5 = parseFloat(cacheData[r][iVel5]) || 0;

          if (v30 > 0) {
            score = v5 / v30; // High ratio = surging short-term momentum
          } else if (v5 > 0) {
            score = 2.0; // Captures sudden new market movement
          }
        }

        const maxScore = momentumMap.get(tId) || 0;
        if (score > maxScore) momentumMap.set(tId, score);
      }

      // Sort the baseline itemIDs array before matrix construction
      itemIDs.sort((a, b) => {
        const scoreA = momentumMap.has(a) ? momentumMap.get(a) : 1.0;
        const scoreB = momentumMap.has(b) ? momentumMap.get(b) : 1.0;
        return scoreB - scoreA; // Descending order (highest priority first)
      });
    }
  } catch (e) {
    // Fail-safe: If caching fails, continue with default sorting order without breaking
    (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('MasterRunner') : console).warn('Velocity priority skipped: ' + e.message);
  }
  // --- END PRIORITIZATION LAYER ---

  const requests = [];
  const processedKeys = new Set();
  
  for (const pair of marketPairs) {
    for (const type_id of itemIDs) {
      const key = `${type_id}|${pair.market_id}|${pair.market_type}`;
      
      if (!processedKeys.has(key)) {
        requests.push({
          type_id: type_id,
          market_id: pair.market_id,
          market_type: pair.market_type
        });
        processedKeys.add(key);
      }
    }
  }
  
  // Note: Sorting strictly by market_id preserves your chunk sequence rules, 
  // but because itemIDs entered the loop pre-sorted by momentum, items within 
  // each market group maintain their priority order perfectly.
  requests.sort((a, b) => a.market_id - b.market_id);

  return requests;
}

/**
 * --- MASTER FUNCTION DEFINITION (Required by MarketFetcher.gs.js) ---
 * Builds the full matrix of {type_id, market_id, market_type} requests.
 */
function getMasterMarketRequests() {
  const itemIDs = _readItemIDList(ITEM_ID_SOURCE);
  const marketPairs = _readMarketIDsFromSettings(); // Get all available market ID/type combinations

  if (itemIDs.length === 0 || marketPairs.length === 0) {
    (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('MasterRunner') : console).warn('Market request list is empty.');
    return [];
  }

  const requests = [];
  const processedKeys = new Set();
  
  // Loop Type ID's for each market ID and market type to build the list
  for (const pair of marketPairs) {
    for (const type_id of itemIDs) {
      const key = `${type_id}|${pair.market_id}|${pair.market_type}`;
      
      // Ensure we don't accidentally double-add (e.g. if the same ID appears twice in the sheet)
      if (!processedKeys.has(key)) {
        requests.push({
          type_id: type_id,
          market_id: pair.market_id,
          market_type: pair.market_type
        });
        processedKeys.add(key);
      }
    }
  }
  
  // Sort requests numerically by market_id
  requests.sort((a, b) => a.market_id - b.market_id);

  return requests;
}