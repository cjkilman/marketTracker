function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Sheet Tools')
    .addItem('Refresh All Data', 'refreshData')
    .addItem('Update SDE Data', 'importSDE')
    .addItem("pull Prices", "getCurrentMarketPrices")
    .createMenu('Engine')
    .addItem('Setup Engine (first run)', 'setupEngine')
    .addSeparator()
    .addItem('Install daily kickoff', 'installDailyKickoff')
    .addItem('Kickoff now (mark STALE & start worker)', 'kickoffMarketHistoryRefresh')
    .addItem('Stop worker', 'stopWorker_')
    .addSeparator()
    .addItem('Publish now', 'publishMarketResultESIRegion')
    .addItem('Debug sources (log counts)', 'debugListSources')
    .addItem('Update Regions & Items', 'applyConfigChange_')
    .addToUi();
}

/**
 * REVISED: Creates all necessary triggers for the project.
 * This now points to the new "Starter" functions for the daily jobs.
 */
function createTriggers() {
  _deleteExistingTriggers(); // Deletes all triggers

  // 1. The 15-Minute "Pipper"
  ScriptApp.newTrigger('masterOrchestrator')
    .timeBased()
    .everyMinutes(15)
    .create();

  // 2. The 24-Hour State Check (Light Prune + Fuzz Reset)
  // Runs daily at 11 PM
  ScriptApp.newTrigger('dailyJobReset')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .create();

  // 3. The Heavy Prune Starter
  // Runs daily at 2 AM (giving the reset plenty of time)
  ScriptApp.newTrigger('dailyHeavyPrune_Prices')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  // 4. The History/Candlestick Starter
  // Runs daily at 3 AM
  ScriptApp.newTrigger('updateHistory')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  // 5. Candlestick Builder
  ScriptApp.newTrigger('buildAllCandlesticks')
    .timeBased()
    .everyDays(1)
    .atHour(4)
    .create();

  SpreadsheetApp.getUi().alert('All triggers have been created successfully.');
}


function getStructureNames(structureIDs) {
  if (!(Array.isArray(structureIDs))) { structureIDs = [[structureIDs]] };
  var output = [];
  for (var i = 0; i < structureIDs.length; i++) {
    var data = GESI.universe_structures_structure(structureIDs[i][0], GESI.getMainCharacter(), false);
    output.push(data[0][0]);
  }
  return output;
}

function pull_SDE() {
  // Lock Formulas from running
  const haltFormulas = [[0, 0]];

  var thisSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var loadingHelper = thisSpreadSheet.getRangeByName("'Utility'!B3:C3");
  const backupSettings = loadingHelper.getValues();
  loadingHelper.setValues(haltFormulas);

  try {

    const sdePages = [
      new SdePage(
        "SDE_invTypes",
        "invTypes.csv",
        // Optional headers,  
        // invTypes is 100+ megabytes. Select columns needed to help it load faster. 
        ["typeID", "groupID", "typeName", "volume"]
      ),
      new SdePage(
        "SDE_staStations",
        "staStations.csv",
        // Optional headers,  
        // invTypes is 100+ megabytes. Select columns needed to help it load faster. 
        ["stationID", "security", "stationTypeID", "corporationID", "solarSystemID", "regionID", "stationName"]
      ),
      new SdePage(
        "SDE_industryActivityProducts",
        "industryActivityProducts.csv",
        []
      )
    ];
    sdePages.forEach(buildSDEs);
  }
  finally {
    // release lock
    loadingHelper.setValues(backupSettings);
  }
}

function importSDE() {

  // Display an alert box with a title, message, input field, and "Yes" and "No" buttons. The
  // user can also close the dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Updating the SDE',
    'Updating the SDE may take a few minutes. In the meantime do not close the window otherwise you will have to restart. Continue?',
    ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {

    pull_SDE();
  } else if (response == ui.Button.NO) {
    ui.alert('SDE unchanged.');
  } else {
    ui.alert('SDE unchanged.');
  }
}

/**
 *Get Character Names from thier ID
 *
 * @param {*} charIds
 * @return {*} 
 */
function getChacterNameFromID(charIds, show_column_headings = true) {
  if (!charIds) throw "undefined charIds";
  if (!Array.isArray(charIds)) charIds = [charIds];
  charIds = charIds.filter(Number);

  let chars = [];
  if (show_column_headings) chars = chars.concat("chacacter_name");
  const rowIdx = show_column_headings ? 1 : 0;

  for (I = 0; I < charIds.length; I++) {
    try {
      const char = GESI.characters_character(Number(charIds[I]), show_column_headings);
      chars = chars.concat(char[rowIdx][7]);
    }
    catch (e) {
      throw e;
    }
  }
  Logger.log(chars);
  return chars;
}

/**
 * Enhances Google Sheets' native "query" method.  Allows you to specify column-names instead of using the column letters in the SQL statement (no spaces allowed in identifiers)
 * 
 * Sample : =query(data!A1:I,SQL("data!A1:I1","SELECT Owner-Name,Owner-Email,Type,Account-Name",false),true)
 *  
 * Params : useColNums (boolean) : false/default = generate "SELECT A, B, C" syntax 
 *                                 true = generate "SELECT Col1, Col2, Col3" syntax
 * reference: https://productforums.google.com/forum/#!topic/docs/vTgy3hgj4M4
 * by: Matthew Quinlan
 */
function sqlFromHeaderNames(rangeName, queryString, useColNums) {

  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let range;
  try {
    range = ss.getRange(rangeName);
  }
  catch (e) {
    range = ss.getRangeByName(rangeName);
  }

  let headers = range.getValues()[0];

  for (var i = 0; i < headers.length; i++) {
    if (headers[i].length < 1) continue;
    var re = new RegExp("\\b" + headers[i] + "\\b", "gm");
    if (useColNums) {
      var columnName = "Col" + Math.floor(i + 1);
      queryString = queryString.replace(re, columnName);
    }
    else {
      var columnLetter = range.getCell(1, i + 1).getA1Notation().split(/[0-9]/)[0];
      queryString = queryString.replace(re, columnLetter);
    }
  }
  //Logger.log(queryString);
  return queryString;
}
