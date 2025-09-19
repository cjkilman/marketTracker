// --- Config ---
const timeDelay = 300;                 // ms between 0→1
const UTILITY_SHEET = "Utility";
const TICK = { BOOK: "B3", STATIC: "C3", ESI: "D3" };

const ss = SpreadsheetApp.getActiveSpreadsheet();
const util = ss.getSheetByName(UTILITY_SHEET);

// Flip a single tick cell: 0 then 1
function bumpTick_(a1) {
  util.getRange(a1).setValue(0);
  Utilities.sleep(timeDelay);
  util.getRange(a1).setValue(1);
}

// Reset all ticks to 0
function resetAllTicks_() {
  util.getRange(TICK.BOOK).setValue(0);
  util.getRange(TICK.STATIC).setValue(0);
  util.getRange(TICK.ESI).setValue(0);
}

// Public: master refresh (order can matter if you want)
function refreshData() {
  resetAllTicks_();
  bumpTick_(TICK.ESI);    // force ESI-dependent calcs (if any)
  bumpTick_(TICK.STATIC); // static lookups, etc.
  bumpTick_(TICK.BOOK);   // book/cache calcs last
}

// Optional: granular buttons
function refreshAllData(){ resetAllTicks_(); }
function refreshDynamicData(){ bumpTick_(TICK.BOOK); }
function refreshStaticData(){  bumpTick_(TICK.STATIC); }
function refreshESI(){         bumpTick_(TICK.ESI); }


