/***********************
 * HistoryManager.gs
 * - Builds daily history from "Market Prices"
 * - Day High  = max(min_sell) over last 24h
 * - Day Low   = min(max_buy) over last 24h
 * - Open/Close snapshots from getMarketPrices(...)
 * - Retention via HistoryRetentionDays (default 365)
 * - Production sheet: "Market History"
 * - Test sheet:       "History (TEST)"
 ***********************/

/**
 * Production entrypoint – respects open/close windows and writes to "Market History".
 * Set two time-based triggers to call this around your OpenTime and CloseTime windows.
 */
function updateHistory() {
  return _updateHistoryCore({ testMode: false });
}

/**
 * Test entrypoint – bypasses window checks and writes to "History (TEST)".
 * Use this while developing.
 */
function updateHistoryTestMode() {
  return _updateHistoryCore({ testMode: true });
}

/** ---------------- Core ---------------- **/
function _updateHistoryCore(opts) {
  var testMode = !!(opts && opts.testMode);

  // --- Config
  var config = getConfig();
  var openStr  = (config["OpenTime"]  !== undefined ? String(config["OpenTime"])  : "11:00");
  var closeStr = (config["CloseTime"] !== undefined ? String(config["CloseTime"]) : "18:00");
  var retentionDays = parseInt(config["HistoryRetentionDays"], 10);
  if (isNaN(retentionDays) || retentionDays <= 0) retentionDays = 365;

  var now = new Date();
  var tz = Session.getScriptTimeZone();
  var todayKey = Utilities.formatDate(now, tz, "yyyy-MM-dd"); // key for today’s rows

  // --- Window check (skip only in production)
  if (!testMode) {
    var withinOpen = _isWithinWindow(now, openStr);
    var withinClose = _isWithinWindow(now, closeStr);
    if (!withinOpen && !withinClose) {
      console.log("⏸ Not within open/close window. Skipping update.");
      return;
    }
  }

  var mode = _decideMode(now, openStr, closeStr); // "OPEN" or "CLOSE"
  if (testMode) {
    console.log("⏩ Test mode active. Mode evaluated as:", mode);
  }

  // --- Sheet targets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historyName = testMode ? "History (TEST)" : "Market History";
  var header = [
    "Date","Type_ID","Market_ID","Market_Type",
    "Day_High","Day_Low",
    "Open_Buy","Open_Sell",
    "Close_Buy","Close_Sell"
  ];
  var historySheet = getOrCreateSheet(ss, historyName, header);

  // --- Source sheets
  var pricesSheet = ss.getSheetByName("Market Prices");
  if (!pricesSheet) throw new Error("Market Prices sheet not found.");

  // --- Pull Type IDs & Markets (reuse your helpers)
  var maxLog = config["MaxLogIDs"] ? parseInt(config["MaxLogIDs"], 10) : 750;
  var typeIDs = getTypeIDsFromItemList(maxLog);
  var marketCombos = getMarketSettings();

  if (!typeIDs.length || !marketCombos.length) {
    console.log("Nothing to do. typeIDs length =", typeIDs.length, "markets length =", marketCombos.length);
    return;
  }

  // --- Build 24h high/low map from Market Prices
  var cutoff = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  var priceData = pricesSheet.getDataRange().getValues();
  priceData.shift(); // drop header: ["type_id","market_id","market_type","max_buy","min_sell","date"]

  var hiLoMap = {}; // key -> { dayHigh, dayLow }
  for (var i = 0; i < priceData.length; i++) {
    var r = priceData[i];
    var type_id = r[0];
    var market_id = r[1];
    var market_type = r[2];
    var max_buy = r[3];
    var min_sell = r[4];
    var stamp = new Date(r[5]);

    if (isNaN(stamp) || stamp < cutoff) continue;

    var key = _key(todayKey, type_id, market_id, market_type);

    // Day High = max(min_sell), Day Low = min(max_buy)
    var entry = hiLoMap[key] || { dayHigh: null, dayLow: null };
    if (typeof min_sell === "number" && !isNaN(min_sell)) {
      entry.dayHigh = (entry.dayHigh == null) ? min_sell : Math.max(entry.dayHigh, min_sell);
    }
    if (typeof max_buy === "number" && !isNaN(max_buy)) {
      entry.dayLow = (entry.dayLow == null) ? max_buy : Math.min(entry.dayLow, max_buy);
    }
    hiLoMap[key] = entry;
  }

  // --- Snapshot prices for this moment (Open/Close)
  // For each market combo, query current prices once per combo for all typeIDs
  var snapMap = {}; // key -> { openBuy?, openSell?, closeBuy?, closeSell? } (only fills based on mode)
  for (var m = 0; m < marketCombos.length; m++) {
    var mc = marketCombos[m];
    var market_id = mc.market_id;
    var market_type = mc.market_type;

    var snap = getMarketPrices(typeIDs, market_id,market_type ); // { [type_id]: { minSell, maxBuy } }

      for (var t = 0; t < typeIDs.length; t++) {
        var tid = typeIDs[t];
        var p = snap[tid] || {};
        var key2 = _key(todayKey, tid, market_id, market_type);

        var holder = snapMap[key2] || {};
        if (mode === "OPEN") {
          holder.openBuy  = (Number(p.maxBuy)  > 0) ? Number(p.maxBuy)  : null;
          holder.openSell = (Number(p.minSell) > 0) ? Number(p.minSell) : null;
        } else {
          holder.closeBuy  = (Number(p.maxBuy)  > 0) ? Number(p.maxBuy)  : null;
          holder.closeSell = (Number(p.minSell) > 0) ? Number(p.minSell) : null;
        }
        snapMap[key2] = holder;
      }
  }

  // --- Merge into existing history
  var existing = historySheet.getDataRange().getValues();
  var histHeader = existing.shift(); // drop header

  // Build index for today's rows
  var idx = {}; // key -> row array (without header)
  var others = []; // rows not for today (keep as-is)
  for (var e = 0; e < existing.length; e++) {
    var row = existing[e];
    if (!row || !row.length) continue;
    var d = row[0];
    if (d === todayKey) {
      var k = _key(d, row[1], row[2], row[3]); // Date + type + market id + type
      idx[k] = row;
    } else {
      others.push(row);
    }
  }

  // Construct the set of keys we care about today (union of hiLoMap and snapMap)
  var keysToday = new Set([].concat(Object.keys(hiLoMap), Object.keys(snapMap)));

  var mergedToday = [];
  keysToday.forEach(function(k) {
    var parts = k.split("|");
    var date = parts[0];
    var tid = Number(parts[1]);
    var mid = Number(parts[2]);
    var mtype = parts[3];

    var base = idx[k] || [
      date, tid, mid, mtype, // 0..3
      "", "",                // Day_High, Day_Low
      "", "",                // Open_Buy, Open_Sell
      "", ""                 // Close_Buy, Close_Sell
    ];

    var hl = hiLoMap[k];
    if (hl) {
      if (hl.dayHigh != null) base[4] = hl.dayHigh;
      if (hl.dayLow  != null) base[5] = hl.dayLow;
    }

    var snapv = snapMap[k];
    if (snapv) {
      if (snapv.openBuy   !== undefined && snapv.openBuy   !== null) base[6] = snapv.openBuy;
      if (snapv.openSell  !== undefined && snapv.openSell  !== null) base[7] = snapv.openSell;
      if (snapv.closeBuy  !== undefined && snapv.closeBuy  !== null) base[8] = snapv.closeBuy;
      if (snapv.closeSell !== undefined && snapv.closeSell !== null) base[9] = snapv.closeSell;
    }

    mergedToday.push(base);
  });

  // --- Apply retention (drop older than retentionDays)
  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - retentionDays);

  var kept = others.filter(function(row) {
    var d = _parseDateYMD(row[0], tz);
    return d >= cutoffDate;
  });

  // Rebuild sheet
  var out = [histHeader].concat(kept, mergedToday);
  historySheet.clearContents();
  historySheet.getRange(1, 1, out.length, histHeader.length).setValues(out);

  console.log("✅ History updated. Mode:", mode,
              "Today rows:", mergedToday.length,
              "Kept older rows:", kept.length);
}

/** ---------------- Helpers ---------------- **/
function _key(dateStr, type_id, market_id, market_type) {
  return [dateStr, type_id, market_id, market_type].join("|");
}

function _parseDateYMD(ymd, tz) {
  // ymd: "yyyy-MM-dd"
  var parts = String(ymd).split("-");
  var y = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1, d = parseInt(parts[2], 10);
  var dt = new Date(y, m, d, 12, 0, 0); // midday to avoid DST edge
  return dt;
}

function _isWithinWindow(now, hhmm) {
  // "HH:mm" => 1-hour window starting at that time (local tz)
  var hourMin = _parseHHmm(hhmm);
  var start = new Date(now);
  start.setHours(hourMin.h, hourMin.m, 0, 0);
  var end = new Date(start.getTime() + 60 * 60 * 1000);
  return now >= start && now < end;
}

function _decideMode(now, openStr, closeStr) {
  // If within open window -> OPEN, else if within close -> CLOSE, else choose closest upcoming
  if (_isWithinWindow(now, openStr)) return "OPEN";
  if (_isWithinWindow(now, closeStr)) return "CLOSE";
  // Outside both: pick whichever window is closer in time (used only for test visibility/debug)
  var o = _parseHHmm(openStr), c = _parseHHmm(closeStr);
  var nowMin = now.getHours() * 60 + now.getMinutes();
  var oMin = o.h * 60 + o.m;
  var cMin = c.h * 60 + c.m;
  return (Math.abs(nowMin - oMin) <= Math.abs(nowMin - cMin)) ? "OPEN" : "CLOSE";
}

function _parseHHmm(timeLike) {
  // Accepts Date | number | string with lots of variants
  if (timeLike instanceof Date) {
    return { h: timeLike.getHours(), m: timeLike.getMinutes() };
  }

  let s = String(timeLike || "").trim().toLowerCase();
  if (!s) return { h: 0, m: 0 };

  // If a range is provided, take the first segment (e.g., "16:00–17:00")
  // splits on hyphen, en dash, em dash, or the word 'to'
  s = s.split(/\s*(?:-|–|—|\bto\b)\s*/)[0].trim();

  // Try "H[:MM][am|pm]" patterns
  // Examples: "4pm", "4:30 pm", "16:00", "04:05", "16"
  const m = s.match(/^(\d{1,2})(?::\s*(\d{1,2}))?\s*(am|pm)?$/i);
  if (m) {
    let h = parseInt(m[1], 10);
    let min = m[2] != null ? parseInt(m[2], 10) : 0;
    const ap = m[3] ? m[3].toLowerCase() : null;

    if (isNaN(h)) h = 0;
    if (isNaN(min)) min = 0;

    // 12-hour clock handling
    if (ap) {
      if (ap === 'am') {
        if (h === 12) h = 0;
      } else if (ap === 'pm') {
        if (h !== 12) h = (h + 12) % 24;
      }
    }

    // Clamp to sane ranges
    h = Math.max(0, Math.min(23, h));
    min = Math.max(0, Math.min(59, min));
    return { h, m: min };
  }

  // Last resort: find first HH:MM anywhere in the string
  const m2 = s.match(/(\d{1,2})\s*:\s*(\d{1,2})/);
  if (m2) {
    let h = parseInt(m2[1], 10);
    let min = parseInt(m2[2], 10);
    if (isNaN(h)) h = 0;
    if (isNaN(min)) min = 0;
    h = Math.max(0, Math.min(23, h));
    min = Math.max(0, Math.min(59, min));
    return { h, m: min };
  }

  // If nothing matches but there's a leading integer, treat it as hour
  const m3 = s.match(/^(\d{1,2})/);
  if (m3) {
    let h = parseInt(m3[1], 10);
    if (isNaN(h)) h = 0;
    h = Math.max(0, Math.min(23, h));
    return { h, m: 0 };
  }

  // Fallback
  return { h: 0, m: 0 };
}