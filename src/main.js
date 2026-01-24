// Global Property Service
const SCRIPT_PROPS = PropertiesService.getScriptProperties();

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
    .addItem('Run SDE Update (Full)', 'sde_job_START') // New SDE Start
    .addSeparator()
    .addItem('Manual: Reset Fuzz Job State', '_resetFuzzMarketDataJobState_MENU')
    .addItem('Manual: Reset ESI Job State', '_resetEsiHistoryJobState_MENU')
    .addItem('Manual: Reset SDE Job State', '_resetSdeJobState_MENU') // New SDE Reset
    .addToUi();
}

function GET_SDE_CONFIG() {
  return [
    { 
      name: "SDE_invTypes", 
      file: "invTypes.csv", 
      cols: ["typeID", "groupID", "typeName", "volume", "marketGroupID", "published"] ,
      published: true
    }
  ];
}

function GET_UTILITY_CONFIG() {
  return { sheetName: "Utility", range: "B3:C3" }; // Adjust to your actual "Off" switch location
}

/**
 * NEW: Manual Reset for SDE Job
 */
function _resetSdeJobState_MENU() {
  if (typeof sde_job_FINALIZE === 'function') {
    sde_job_FINALIZE(); 
    SpreadsheetApp.getUi().alert('SDE Job State has been force-reset.');
  } else {
    SpreadsheetApp.getUi().alert('Error: SDE controller not found.');
  }
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
  // --- REFACTORED: This function does not exist in your ESI file. ---
  // if (typeof _resetEsiHistoryJobState === 'function') {
  //   _resetEsiHistoryJobState(new Error("Manual Reset from Menu"));
  //   SpreadsheetApp.getUi().alert('ESI Job State has been reset.');
  // } else {
  //   SpreadsheetApp.getUi().alert('Error: _resetEsiHistoryJobState function not found.');
  // }
  
  // --- We will reset it manually instead ---
  try {
    PropertiesService.getScriptProperties().deleteProperty('mf_job_active');
    PropertiesService.getScriptProperties().deleteProperty('mf_cursor');
    deleteTriggersByName('marketFetchChunk');
    SpreadsheetApp.getUi().alert('ESI Job State has been force-reset.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error resetting ESI state: ' + e.message);
  }
}

/**
 * NEW: Wrapper for menu item to reset Heavy Prune job.
 */
function _resetHeavyPruneJobState_MENU() {
  // _resetHeavyPruneJobState is in MarketFetcher.gs.js
  if (typeof _resetHeavyPruneJobState === 'function') {
    _resetHeavyPruneJobState(new Error("Manual Reset from Menu"));
    SpreadsheetApp.getUi().alert('Heavy Prune Job State has been reset.');
  } else {
    SpreadsheetApp.getUi().alert('Error: _resetHeavyPruneJobState function not found.');
  }
}

// ... (rest of your main.js file) ...