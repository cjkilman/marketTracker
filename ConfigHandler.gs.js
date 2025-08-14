function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = getOrCreateSheet(ss, "Market Config", [
    "Setting", "Value", "Notes"
  ]);

  if (configSheet.getLastRow() < 2) {
    // Default settings
    const defaults = [
      ["type_id", 34, "Default item type ID"],
      ["market_id", 30002187, "Amarr system ID"],
      ["market_type", "system", "System or region market type"],
      ["MaxLogIDs", 700, "Max item IDs to keep in logs (blank = keep all)"],
      ["DaysForCandlestick", 30, "Days to build candlestick chart"],
      ["RebuildAlways", false, "Always rebuild history"],
      ["OpenTime", "11:00", "Daily open window start (HH:mm)"],
      ["CloseTime", "18:00", "Daily close window start (HH:mm)"]
    ];
    configSheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
  }

  return configSheet;
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