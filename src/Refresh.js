//Globalvariables

const timeDelay = 300;
const utilitySheetName = "Utility"

//Use with Care to avoid conflicts
const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
const utilitySheet = activeSheet.getSheetByName(utilitySheetName);

// Function to refresh data 
// sets cells to 0 to retrigger sheet formulas. 
// 1 prevents updating on cell contents change
// This compensates for a refresh bug.

function refreshData() {
    refreshAllData();
    refreshESI();
    refreshStaticData();
    refreshDynamicData();

}
  
function refreshAllData () {

  // Load sheet with variable  
  const resetRow = [
      [0,0,0]
  ];  
    
  utilitySheet.getRange(utilitySheetName +'!B3:D3')
      .setValues(resetRow);

}

function refreshDynamicData(){

  //Time delay before changing back
  Utilities.sleep(timeDelay)


  // Define reference link to specified cell
  utilitySheet.getRange(utilitySheetName +'!B3')
    .setValue(1);
 
}
function refreshESI(){

  //Time delay before changing back
  Utilities.sleep(timeDelay)


  // Define reference link to specified cell
  utilitySheet.getRange(utilitySheetName +'!D3')
    .setValue(1);
 
}

function refreshStaticData(){

    //Time delay before changing back
    Utilities.sleep(timeDelay)

    //Load sheet with variable
    utilitySheet.getRange(utilitySheetName +'!C3')
      .setValue(1);

}