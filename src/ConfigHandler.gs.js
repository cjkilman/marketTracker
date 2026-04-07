function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = getOrCreateSheet(ss, "Market Config", [
    "Setting", "Value", "Notes"
  ]);

  if (configSheet.getLastRow() < 2) {
    // Default settings
const defaults = [
  ["type_id", 34, "Default item type ID (Tritanium)"],
  ["market_id", 30002187, "Amarr system ID"],
  ["market_type", "system", "System or region market type"],
  ["BQ_ENABLED", true, "MASTER SWITCH: TRUE = BigQuery Vault | FALSE = Sheet Backup"],
  ["BQ_QUOTA_GIB", 30.72, "Daily BigQuery Free Tier Limit (Hard Stop)"],
  ["BQ_PROJECT_ID", "tenacious-tiger-345318", "Your BigQuery Project ID"],
  ["BQ_DATASET_ID", "market_data", "Your BigQuery Dataset ID"],
  
  // FIX: This must be the actual BigQuery Table Name (No Spaces)
  ["BQ_Market_Prices", "market_prices_history", "The BigQuery Database Table"],
  
  // FIX: This is the local Google Sheet Tab Name (Spaces OK)
  ["HistorySheetName", "Market History", "The local Spreadsheet Tab"],
  
  ["MaxLogIDs", 700, "Max item IDs to keep in logs"],
  ["DaysForCandlestick", 30, "Days to build candlestick chart"],
  ["RebuildAlways", false, "Always rebuild history (Slow)"],
  ["OpenTime", "11:00", "Daily open window start (HH:mm)"],
  ["CloseTime", "18:00", "Daily close window start (HH:mm)"]
];
    configSheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }

  return configSheet;
}

function FORCE_SYNC_CONFIG() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Market Config");
  
  const correctedData = [
    ["type_id", 34, "Default item type ID (Tritanium)"],
    ["market_id", 30002187, "Amarr system ID"],
    ["market_type", "system", "System or region market type"],
    ["BQ_ENABLED", true, "MASTER SWITCH: TRUE = BigQuery Vault | FALSE = Sheet Backup"],
    ["BQ_QUOTA_GIB", 30.72, "Daily BigQuery Free Tier Limit (Hard Stop)"],
    ["BQ_PROJECT_ID", "tenacious-tiger-345318", "Your BigQuery Project ID"],
    ["BQ_DATASET_ID", "market_data", "Your BigQuery Dataset ID"],
    ["BQ_Market_Prices", "market_prices_history", "The BigQuery Database Table"],
    ["HistorySheetName", "Market History", "The local Spreadsheet Tab"],
    ["MaxLogIDs", 700, "Max item IDs to keep in logs"],
    ["DaysForCandlestick", 30, "Days to build candlestick chart"],
    ["RebuildAlways", false, "Always rebuild history (Slow)"],
    ["OpenTime", "11:00", "Daily open window start (HH:mm)"],
    ["CloseTime", "18:00", "Daily close window start (HH:mm)"]
  ];

  // This forces the write regardless of whether the sheet is empty
  sh.getRange(2, 1, correctedData.length, 3).setValues(correctedData);
  console.log("✅ CONFIG FORCIBLY UPDATED. Keys are now synced.");
}

/**
 * Updates a specific config value in the 'Market Config' sheet.
 * Used by the Orchestrator to kill BQ_ENABLED if the wallet-locker is hit.
 */
function setMarketConfig(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Market Config");
  if (!sheet) return;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      console.log(`[CONFIG] ${key} updated to: ${value}`);
      return;
    }
  }
  console.warn(`[CONFIG] Key '${key}' not found. No update made.`);
}

function getConfig() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Market Config");
  if (!configSheet) throw new Error("Market Config sheet not found.");

  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) throw new Error("Market Config is empty or missing data.");

  const data = configSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const config = {};
  data.forEach(([key, value]) => {
    config[key] = value;
  });
  return config;
}