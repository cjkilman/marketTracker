/** ===== Quick Sanity: presence-only, order/case-insensitive ===== **/
function Integrity_Quick() {
  const ss = SpreadsheetApp.getActive();

  _iq_requireHeaders(
    ss.getSheetByName("Market Prices"),
    ["type_id","market_id","market_type","max_buy","min_sell","date"],
    "Market Prices"
  );

  _iq_requireHeaders(
    ss.getSheetByName("Market History"),
    ["type_id","market_id","market_type","date","buy_open","buy_close","sell_open","sell_close","daily_high","daily_low","median_buy","median_sell"],
    "Market History"
  );

  (typeof Log !== "undefined")
    ? Log.for("Integrity").info("Integrity_Quick: PASS")
    : Logger.log("Integrity_Quick: PASS");
  return true;
}

function _iq_requireHeaders(sheet, required, name){
  if (!sheet) throw new Error(`Missing sheet: ${name}`);
  const got = sheet
    .getRange(1,1,1, Math.max(sheet.getLastColumn(), required.length))
    .getValues()[0]
    .map(v => String(v||"").trim().toLowerCase());
  const set = new Set(got);
  for (const need of required.map(h => h.toLowerCase())) {
    if (!set.has(need)) {
      throw new Error(`Missing required header on ${name}: "${need}"`);
    }
  }
}