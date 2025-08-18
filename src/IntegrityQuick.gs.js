/** ===== Quick Integrity (fast daily check) ===== **/
function Integrity_Quick() {
  const ss = SpreadsheetApp.getActive();
  _iq_checkHeaders(ss.getSheetByName("Market Prices"), [
    "type_id","market_id","market_type","max_buy","min_sell","date"
  ], "Market Prices");

  _iq_checkHeaders(ss.getSheetByName("Market History"), [
    "type_id","market_id","market_type","date","buy_open","buy_close","sell_open","sell_close","daily_high","daily_low","median_buy","median_sell"
  ], "Market History");

  const msg = "Integrity_Quick: PASS";
  if (typeof Log !== "undefined") Log.for("Integrity").info(msg); else Logger.log(msg);
  return true;
}

function _iq_checkHeaders(sheet, expected, name) {
  if (!sheet) throw new Error(`Missing sheet: ${name}`);
  const header = sheet.getRange(1,1,1,expected.length).getValues()[0];
  if (header.join("|") !== expected.join("|")) {
    throw new Error(`Header mismatch on ${name}`);
  }
}