/**
 * Main.gs
 * * Entry point for custom menus and trigger setup.
 * All core logic is in other modules.
 */

/* global SpreadsheetApp, _deleteExistingTriggers, masterOrchestrator, dailyJobReset, dailyHeavyPrune_Prices, updateHistory, buildAllCandlesticks, _resetFuzzMarketDataJobState, _resetEsiHistoryJobState, Logger */

/**
 * Creates the "Admin Tools" menu when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Admin Tools')
    .addItem('Create Triggers', 'createTriggers')
    .addItem('Delete All Triggers', '_deleteExistingTriggers')
    .addSeparator()
    .addItem('Run 15-Min Orchestrator', 'masterOrchestrator')
    .addItem('Run Daily Reset (Light Prune)', 'dailyJobReset')
    .addItem('Run Daily Heavy Prune', 'dailyHeavyPrune_Prices')
    .addItem('Run Daily History (OHLC) Build', 'updateHistory')
    .addSeparator()
    .addItem('Manual: Reset Fuzz Job State', '_resetFuzzMarketDataJobState_MENU')
    .addItem('Manual: Reset ESI Job State', '_resetEsiHistoryJobState_MENU')
    .addToUi();
}

/**
 * Wrapper for menu item to reset Fuzz job.
 */
function _resetFuzzMarketDataJobState_MENU() {
  // _resetFuzzMarketDataJobState is in MarketFetcher.gs.js
  if (typeof _resetFuzzMarketDataJobState === 'function') {
    _resetFuzzMarketDataJobState(new Error("Manual Reset from Menu"));
    SpreadsheetApp.getUi().alert('Fuzz Job State has been reset.');
  } else {
    SpreadsheetApp.getUi().alert('Error: _resetFuzzMarketDataJobState function not found.');
  }
}

/**
 * Wrapper for menu item to reset ESI job.
 */
function _resetEsiHistoryJobState_MENU() {
  // _resetEsiHistoryJobState is in marketFetcherEsi.js
  if (typeof _resetEsiHistoryJobState === 'function') {
    _resetEsiHistoryJobState(new Error("Manual Reset from Menu"));
    SpreadsheetApp.getUi().alert('ESI Job State has been reset.');
  } else {
    SpreadsheetApp.getUi().alert('Error: _resetEsiHistoryJobState function not found.');
  }
}

/**
 * Deletes all triggers in the project.
 */
function _deleteExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
  Logger.log(`Deleted ${triggers.length} triggers.`);
}

/**
 * REVISED: Creates all necessary triggers for the project.
 * This now points to the new "Starter" functions for the daily jobs.
 * The History/Candlestick layer is PARKED (commented out).
 */
function createTriggers() {
  _deleteExistingTriggers(); // Deletes all triggers
  const log = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('Triggers') : console;

  // 1. The 15-Minute "Pipper"
  ScriptApp.newTrigger('masterOrchestrator')
    .timeBased()
    .everyMinutes(15)
    .create();
  log.info("Created 15-minute trigger for masterOrchestrator.");

  // 2. The 24-Hour State Check (Light Prune + Fuzz Reset)
  // Runs daily at 11 PM
  ScriptApp.newTrigger('dailyJobReset')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();
  log.info("Created daily 11 PM trigger for dailyJobReset.");

  // 3. The Heavy Prune Starter
  // Runs daily at 2 AM (giving the reset plenty of time)
  ScriptApp.newTrigger('dailyHeavyPrune_Prices')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  log.info("Created daily 2 AM trigger for dailyHeavyPrune_Prices.");

  // --- PARKED JOBS (Candlestick Layer) ---
  
  // 4. The History/Candlestick Starter
  // ScriptApp.newTrigger('updateHistory')
  //   .timeBased()
  //   .everyDays(1)
  //   .atHour(3)
  //   .create();
  // log.info("Created daily 3 AM trigger for updateHistory.");

  // 5. Candlestick Builder
  // ScriptApp.newTrigger('buildAllCandlesticks')
  //   .timeBased()
  //   .everyDays(1)
  //   .atHour(4)
  //   .create();
  // log.info("Created daily 4 AM trigger for buildAllCandlesticks.");
  
  log.info("--- History/Candlestick jobs are PARKED and triggers were NOT created. ---");

  SpreadsheetApp.getUi().alert('All active triggers have been created successfully. (History/Candlestick jobs are parked).');
}