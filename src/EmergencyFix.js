/**
 * EMERGENCY IN-PLACE PRUNE
 * Deletes old rows directly from 'Market Prices' without creating a temp sheet.
 * Use this when the workbook is too full to run the standard Heavy Prune.
 */
function emergencyPruneInPlace() {
  const RETENTION_DAYS = 2; // Keep only 2 days of data (matches your config)
  const SHEET_NAME = 'Market Prices';
  const DATE_COL_IDX = 1; // Column A is date
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    console.error(`Sheet '${SHEET_NAME}' not found!`);
    return;
  }
  
  // 1. Clean up the failed temp sheet from the crash
  const badTemp = ss.getSheetByName('Market_Prices_Prune_Temp');
  if (badTemp) {
    console.warn("Deleting leftover temp sheet from failed run...");
    ss.deleteSheet(badTemp);
  }

  // 2. Calculate Cutoff
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - RETENTION_DAYS);
  console.log(`Pruning '${SHEET_NAME}' entries older than: ${cutoff.toISOString()}`);

  // 3. Find the boundary (Binary search would be faster, but linear is safer for mixed data)
  // We'll read just the Date column to save memory
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const dates = sheet.getRange(2, DATE_COL_IDX, lastRow - 1, 1).getValues().flat();
  let deleteCount = 0;
  
  // Assuming chronological order (oldest at top), find the first row that IS NEW enough
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i];
    if (d instanceof Date && d < cutoff) {
      deleteCount++;
    } else {
      // Found the first "safe" date, stop counting
      break;
    }
  }
  
  if (deleteCount > 0) {
    console.log(`Deleting ${deleteCount} old rows...`);
    
    // Delete in chunks to avoid timeouts
    const CHUNK_SIZE = 5000;
    while (deleteCount > 0) {
      const toDelete = Math.min(deleteCount, CHUNK_SIZE);
      // Always delete from row 2 (shifting subsequent rows up)
      sheet.deleteRows(2, toDelete); 
      deleteCount -= toDelete;
      console.log(`Deleted ${toDelete} rows. Remaining: ${deleteCount}`);
      SpreadsheetApp.flush(); // Commit changes
    }
    console.log("Emergency prune complete.");
  } else {
    console.log("No rows older than retention limit found.");
  }
}