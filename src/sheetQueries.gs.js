/**
 * Wrapper for the scheduler (Hardened against null/undefined event objects)
 */
function refreshPriceInterfaceSheetsManual(e) {
  let targetSheetName;

  // 1. Check if 'e' exists at all
  if (e) {
    // If called directly with a string: ManualRefresh("Mineral Supply Prices")
    if (typeof e === 'string') {
      targetSheetName = e;
    } 
    // If called by a Time Trigger, 'e' is an object containing 'triggerUid'
    else if (e.triggerUid) {
      targetSheetName = PropertiesService.getScriptProperties().getProperty('last_scheduled_sheet');
    }
  }

  // 2. Fallback for manual IDE runs (Default to main sheet if e is null)
  if (!targetSheetName) {
    console.warn("[IDE RUN] No event object detected. Defaulting to 'filtered prices'.");
    targetSheetName = 'filtered prices';
  }

  // 3. Final validation to prevent the [object Object] crash
  if (targetSheetName === "[object Object]") {
    console.error("[ERROR] Resolved target was an object string. Aborting.");
    return;
  }

  console.log(`[EXECUTE] Background pulse starting for: ${targetSheetName}`);
  refreshPriceInterfaceSheets(null, targetSheetName, null);
}


/**
 * WORKER: Surgical BigQuery Puller with ID Filtering & API Fallback.
 * Targets Column E:I on the target sheet using Item List Back End (A2:A).
 */
function refreshPriceInterfaceSheets(ss, targetSheetName, uniqueIds) {
  const projectId = 'tenacious-tiger-345318';
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);

  if (!sheet) {
    console.warn(`[SKIP] Sheet "${targetSheetName}" not found.`);
    return;
  }

  // --- 1. ID FILTERING (The Cost Shield) ---
  if (!uniqueIds) {
    const itemSheet = ss.getSheetByName("Item List Back End");
    const rawIds = itemSheet.getRange(2, 1, Math.max(1, itemSheet.getLastRow() - 1), 1).getValues();
    uniqueIds = [...new Set(rawIds.flat().map(id => parseInt(id)).filter(id => !isNaN(id) && id > 0))];
  }
  
  if (uniqueIds.length === 0) return;
  const idString = uniqueIds.join(',');

  // --- 2. INPUT GATING ---
  const marketId = sheet.getRange("C4").getValue();
  const marketType = sheet.getRange("D4").getValue();

  if (!marketId || marketId === "#N/A" || !marketType || marketType.toLowerCase().includes("loading")) {
    sheet.getRange("E4").setValue(`⚠️ Waiting for Settings...`);
    return; 
  }

  // --- 3. THE RECALL SQL ---
  const sql = `
    SELECT 
      type_id, ROUND(AVG(median_buy), 2), ROUND(AVG(median_sell), 2),
      ARRAY_AGG(median_buy ORDER BY date DESC LIMIT 1)[OFFSET(0)],
      ARRAY_AGG(median_sell ORDER BY date DESC LIMIT 1)[OFFSET(0)]
    FROM \`${projectId}.market_data.market_prices\`
    WHERE market_id = ${marketId}
      AND LOWER(market_type) = LOWER('${marketType}')
      AND date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 24 HOUR)
      AND type_id IN (${idString})
    GROUP BY type_id ORDER BY type_id ASC
  `;

  try {
    const queryResults = BigQuery.Jobs.query({query: sql, useLegacySql: false}, projectId);
    const data = queryResults.rows ? queryResults.rows.map(row => row.f.map(field => field.v)) : [];

    if (data.length === 0) throw new Error("Empty BQ result");
    writeToInterfaceSheet(sheet, data, "✅ Synced");

  } catch (err) {
    console.warn(`[${targetSheetName}] BQ Fail (Quota/Error). Starting API Fallback...`);
    
    // --- 4. THE API FALLBACK (The Fail-Safe) ---
    try {
      const apiData = getMarketPrices(uniqueIds, marketId, marketType);
      const fallbackData = uniqueIds.map(id => {
        const item = apiData[id] || {};
        return [id, item.buy?.median || 0, item.sell?.median || 0, item.buy?.max || 0, item.sell?.min || 0];
      });
      writeToInterfaceSheet(sheet, fallbackData, "⚠️ API Fallback");
    } catch (apiErr) {
      sheet.getRange("E4").setValue("❌ Sync Failed");
    }
  }
}

/**
 * HELPER: Aligns data with your Row 7 Header / Row 8 Start format.
 */
function writeToInterfaceSheet(sheet, data, statusPrefix) {
  sheet.getRange("E8:I").clearContent();
  sheet.getRange(7, 5, 1, 5).setValues([["type_id_filtered", "Median Buy", "Median Sell", "Current Buy", "Current Sell"]]);
  if (data.length > 0) sheet.getRange(8, 5, data.length, 5).setValues(data);
  sheet.getRange("E4").setValue(`${statusPrefix}: ${new Date().toLocaleTimeString()}`);
}


/**
 * MASTER DISPATCHER: The heartbeat of your 132-slot farm.
 * Schedules background pulses for each market sheet to avoid timeouts.
 */
function masterMarketRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  
  // 1. Data Integrity Gate: Don't pulse if BigQuery is currently being streamed to
  if (props.getProperty('fuzz_job_active') === 'true') {
    console.warn("[REFRESH] BQ Stream active. Skipping pulse to avoid data collision.");
    return; 
  }

  // 2. The Sheet Queue
  const marketSheets = ['filtered prices', 'Mineral Supply Prices', 'T1 Supply Prices'];

  // 3. ID Filtering: Get unique IDs once to share across the whole engine
  const itemSheet = ss.getSheetByName("Item List Back End");
  const lastRow = itemSheet.getLastRow();
  if (lastRow < 2) {
    console.error("[ERROR] Item List Back End is empty.");
    return;
  }
  
  const rawIds = itemSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const uniqueIds = [...new Set(rawIds.flat().map(id => parseInt(id)).filter(id => !isNaN(id) && id > 0))];

  // 4. Action: Run the first sheet immediately (Main Money Printer)
  console.log(`[PULSE] Starting immediate refresh for: ${marketSheets[0]}`);
  refreshPriceInterfaceSheets(ss, marketSheets[0], uniqueIds);

// 5. Stagger: Use a more robust hand-off
  marketSheets.slice(1).forEach((sheetName, index) => {
    const delayMs = (index + 1) * 35000; // Increased to 35s to allow for API latency
    
    // We use a unique property key for each scheduled slot
    const slotKey = `scheduled_sheet_slot_${index}`;
    props.setProperty(slotKey, sheetName);
    
    // Pass the slotKey as the function name via a small wrapper if needed, 
    // or keep the current property logic but be aware of the race condition.
    props.setProperty('last_scheduled_sheet', sheetName); 
    
    scheduleOneTimeTrigger('refreshPriceInterfaceSheetsManual', delayMs);
    console.log(`[STAGGER] Scheduled ${sheetName} for pulse in ${delayMs/1000}s`);
  });
}