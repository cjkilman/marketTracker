/**
 * MASTER DISPATCHER: Rotates through all listed market-pull sheets.
 * Hook this to your 30-minute trigger.
 */
function masterMarketRefresh() {
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const NOW_MS = Date.now();

  // 1. ORCHESTRATOR LEASE CHECK (Check once for the whole fleet)
  const leaseUntil = parseInt(SCRIPT_PROP.getProperty('fuzzJobLeaseUntil') || '0', 10);
  if (NOW_MS < leaseUntil) {
    console.warn(`[ORCHESTRATOR] Engine Busy until ${new Date(leaseUntil).toLocaleTimeString()}. Skipping refresh cycle.`);
    return;
  }

  // 2. THE FLEET: List every sheet that uses the C4/D4/A7 layout
  const marketSheets = [
    'filtered prices', 
    'Mineral Supply Prices', 
    'T1 Supply Prices',
  ];

// 3. ROTATE AND REFRESH
  marketSheets.forEach(sheetName => {
    try {
      refreshFilteredPrices(sheetName);
      SpreadsheetApp.flush(); // Force the write to finish
      Utilities.sleep(500);   // Tiny gap to keep the UI snappy
    } catch (e) {
      console.error(`[CRITICAL] Failed to refresh sheet "${sheetName}": ${e.message}`);
    }
  });
}

/**
 * WORKER: Hardened BigQuery Puller for a specific sheet.
 * Now acts as a modular unit for the Dispatcher.
 */
function refreshFilteredPrices(targetSheetName) {
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const projectId = 'tenacious-tiger-345318';
  
  // Default to main sheet if called manually without an argument
  if (!targetSheetName) targetSheetName = 'filtered prices';
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);

  if (!sheet) {
    console.warn(`[SKIP] Sheet "${targetSheetName}" not found in workbook.`);
    return;
  }

  // --- 1. INPUT GATING (Handle ImportRange Failures) ---
  const marketId = sheet.getRange("C4").getValue();
  const marketType = sheet.getRange("D4").getValue();

  // Validate inputs: Abort if #N/A or Loading
  const isIdValid = marketId && !isNaN(marketId) && marketId !== "#N/A" && marketId !== "";
  const isTypeValid = marketType && marketType !== "#N/A" && marketType !== "" && marketType.toLowerCase() !== "loading...";

  if (!isIdValid || !isTypeValid) {
    console.warn(`[${targetSheetName}] GATE BLOCKED: Settings missing or invalid. Retrying next cycle.`);
    sheet.getRange("E4").setValue(`⚠️ Delayed: Settings Loading... (${new Date().toLocaleTimeString()})`);
    return; 
  }

  // --- 2. THE BIGQUERY PULL ---
  const sql = `
    SELECT date, type_id, max_buy, min_sell 
    FROM (
      SELECT *, ROW_NUMBER() OVER(PARTITION BY type_id ORDER BY date DESC) as rn
      FROM \`${projectId}.market_data.market_prices\`
      WHERE market_id = ${marketId}
      AND LOWER(market_type) = LOWER('${marketType}')
      AND date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 30 DAY)
    )
    WHERE rn = 1
    ORDER BY type_id ASC
  `;

  try {
    const queryResults = BigQuery.Jobs.query({query: sql, useLegacySql: false}, projectId);
    const rows = queryResults.rows;

    if (!rows || rows.length === 0) {
      console.warn(`[${targetSheetName}] No data found in BQ for Market ${marketId}.`);
      sheet.getRange("E4").setValue("⚠️ No Data Found in BQ");
      return;
    }

    const data = rows.map(row => [
      new Date(row.f[0].v), 
      row.f[1].v, 
      row.f[2].v, 
      row.f[3].v
    ]);

    // --- 3. THE WRITE (Starting Row 7) ---
    sheet.getRange("A7:D").clearContent();
    sheet.getRange(7, 1, 1, 4).setValues([["Date", "Type ID", "Max Buy", "Min Sell"]]);
    sheet.getRange(8, 1, data.length, 4).setValues(data);
    
    // Individual heartbeat for this sheet
    sheet.getRange("E4").setValue(`✅ Synced: ${new Date().toLocaleTimeString()}`);
    console.log(`[${targetSheetName}] Success: ${data.length} items.`);
    
  } catch (err) {
    console.error(`[${targetSheetName}] BQ Error: ${err.message}`);
    sheet.getRange("E4").setValue("❌ BQ Query Error");
  }
}