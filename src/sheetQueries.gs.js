/**
 * MASTER DISPATCHER: Rotates through all listed market-pull sheets.
 * Hook this to your 30-minute trigger.
 */
function masterMarketRefresh() {
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const NOW_MS = Date.now();

// --- THE NEW GATE (Multi-task friendly) ---
  const isWorkerActive = SCRIPT_PROPS.getProperty('fuzz_job_active') === 'true';
  
  if (isWorkerActive) {
    console.warn("[REFRESH] Market Fetcher is currently STREAMING. Skipping refresh to avoid partial data.");
    return; 
  }

  // If we get here, the engine is either IDLE or on LEASE (Cooling down).
  // It is 100% safe to refresh the sheets now.
  console.log("[REFRESH] Engine is parked. Proceeding with Display Sheet update...");

  // 2. THE FLEET: List every sheet that uses the C4/D4/A7 layout
  const marketSheets = [
    'filtered prices', 
    'Mineral Supply Prices', 
    'T1 Supply Prices',
  ];

// 3. ROTATE AND REFRESH
  marketSheets.forEach(sheetName => {
    try {
      refreshPriceInterfaceSheets(sheetName);
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
function refreshPriceInterfaceSheets(targetSheetName) {
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

// --- OPTIMIZED SINGLE QUERY ---
  const sql = `
    SELECT 
      type_id, 
      ROUND(AVG(median_buy), 2) as median_buy_24h, 
      ROUND(AVG(median_sell), 2) as median_sell_24h,
      ARRAY_AGG(median_buy ORDER BY date DESC LIMIT 1)[OFFSET(0)] as current_buy,
      ARRAY_AGG(median_sell ORDER BY date DESC LIMIT 1)[OFFSET(0)] as current_sell
    FROM \`${projectId}.market_data.market_prices\`
    WHERE market_id = ${marketId}
      AND LOWER(market_type) = LOWER('${marketType}')
      /* Narrowing to 24 hours reduces data scan by ~96% */
      AND date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 24 HOUR)
    GROUP BY type_id
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

    // --- 2. DATA MAPPING (Corrected to match the 5 SQL columns) ---
    const data = rows.map(row => [
      row.f[0].v, // type_id (Replaces Date in Column E)
      row.f[1].v, // Median Buy (24h Average)
      row.f[2].v, // Median Sell (24h Average)
      row.f[3].v, // Current Buy (Latest record)
      row.f[4].v  // Current Sell (Latest record)
    ]);

    // --- 3. THE WRITE (Aligning with Row 7, Column E) ---
    // Clear old data from Column E to I
    sheet.getRange("E7:I").clearContent();

    // Set Headers at E7 (Row 7, Column 5)
    sheet.getRange(7, 5, 1, 5).setValues([[
      "type_id_filtered", 
      "Median Buy", 
      "Median Sell", 
      "Current Buy", 
      "Current Sell"
    ]]);
    
    // Set Data starting at E8
    sheet.getRange(8, 5, data.length, 5).setValues(data);

    // Heartbeat update
    sheet.getRange("E4").setValue(`✅ Synced: ${new Date().toLocaleTimeString()}`);
    console.log(`[${targetSheetName}] Success: ${data.length} items.`);
    
  } catch (err) {
    console.error(`[${targetSheetName}] BQ Error: ${err.message}`);
    sheet.getRange("E4").setValue("❌ BQ Query Error");
  }
}