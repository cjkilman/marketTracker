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

  // --- 1. INPUT GATING ---
  const marketId = sheet.getRange("C4").getValue();
  const marketType = sheet.getRange("D4").getValue();

  if (!marketId || marketId === "#N/A" || !marketType || marketType.toLowerCase().includes("loading")) {
    sheet.getRange("E4").setValue(`⚠️ Waiting for Settings...`);
    return; 
  }

  // --- 2. THE RECALL SQL (Optimized) ---
  // We removed the massive 4,000+ item IN() clause. 
  // BigQuery will just return everything for this market from the last 24h.
const sql = `
    SELECT 
      type_id, ROUND(AVG(median_buy), 2), ROUND(AVG(median_sell), 2),
      ARRAY_AGG(median_buy ORDER BY date DESC LIMIT 1)[OFFSET(0)],
      ARRAY_AGG(median_sell ORDER BY date DESC LIMIT 1)[OFFSET(0)]
    FROM \`${projectId}.market_data.market_prices_staged\`
    WHERE market_id = ${marketId}
      AND LOWER(market_type) = LOWER('${marketType}')
      AND date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 24 HOUR)
    GROUP BY type_id ORDER BY type_id ASC
  `;

  try {
    console.log(`[${targetSheetName}] Requesting data from BigQuery...`);
    const queryResults = BigQuery.Jobs.query({query: sql, useLegacySql: false}, projectId);
    
    console.log(`[${targetSheetName}] Query returned. Parsing rows...`);
    const data = queryResults.rows ? queryResults.rows.map(row => row.f.map(field => field.v)) : [];

    if (data.length === 0) throw new Error("Empty BQ result (Waiting for next stream)");
    
    console.log(`[${targetSheetName}] Writing ${data.length} rows to the sheet...`);
    writeToInterfaceSheet(sheet, data, "✅ Synced");
    console.log(`[${targetSheetName}] Write complete!`);

  } catch (err) {
console.warn(`[${targetSheetName}] BQ Query Failed: ${err.message}`);
    sheet.getRange("E4").setValue(`BQ Error: ${err.message}`);
    
    // Auto-trip the circuit breaker on quota errors
    if (err.message.toLowerCase().includes("quota exceeded")) {
      console.error("Quota limit hit. Engaging circuit breaker automatically.");
      toggleBigQueryCircuitBreaker();
    }
    return;
  }
}

function archiveMarketDataToBigQuery() {
  const projectId = 'tenacious-tiger-345318';
  
  // This SQL moves data from the "Sheet" to the "Internal Vault"
  // It only grabs rows that aren't already there (based on timestamp/market/type)
  const sql = `
    INSERT INTO \`tenacious-tiger-345318.market_data.market_prices_history\`
    SELECT s.* FROM \`tenacious-tiger-345318.market_data.market_prices_staged\` s
    WHERE NOT EXISTS (
      SELECT 1 FROM \`tenacious-tiger-345318.market_data.market_prices_history\` h
      WHERE h.date = s.date 
        AND h.market_id = s.market_id 
        AND h.type_id = s.type_id
    )
  `;

  try {
    BigQuery.Jobs.query({query: sql, useLegacySql: false}, projectId);
    console.log("[VAULT] Successfully archived current sheet data to history.");
  } catch (err) {
    console.error("[VAULT ERROR] Archive failed: " + err);
  }
}

function checkLastBqError() {
  const projectId = 'tenacious-tiger-345318';
  
  try {
    console.log("Checking BigQuery's internal logs for the last 5 jobs...");
    const jobList = BigQuery.Jobs.list(projectId, { maxResults: 5, stateFilter: 'done' });
    
    if (!jobList.jobs || jobList.jobs.length === 0) {
      console.log("No recent jobs found.");
      return;
    }
    
    jobList.jobs.forEach(job => {
      // We have to get the full job details to see the exact row errors
      const fullJob = BigQuery.Jobs.get(projectId, job.jobReference.jobId);
      
      if (fullJob.status && fullJob.status.errorResult) {
        console.error(`🔥 FAILED JOB: ${fullJob.jobReference.jobId}`);
        console.error(`Error: ${fullJob.status.errorResult.message}`);
        
        // If there are specific row errors, print them
        if (fullJob.status.errors) {
          fullJob.status.errors.forEach(err => {
            console.error(` -> Detail: ${err.message}`);
          });
        }
      } else {
        console.log(`✅ Passed Job: ${fullJob.jobReference.jobId}`);
      }
    });
    
  } catch (err) {
    console.error(`Failed to fetch job history: ${err.message}`);
  }
}

function peekAtBigQuery() {
  const projectId = 'tenacious-tiger-345318';
  const sql = `
    SELECT date, market_id, market_type, type_id, min_sell 
    FROM \`${projectId}.market_data.market_prices\` 
    ORDER BY date DESC LIMIT 3
  `;
  
  try {
    console.log("Looking inside BigQuery...");
    const result = BigQuery.Jobs.query({query: sql, useLegacySql: false}, projectId);
    
    if (!result.rows || result.rows.length === 0) {
      console.error("🛑 The table is COMPLETELY EMPTY. (The upload skipped or failed).");
    } else {
      console.log("✅ DATA FOUND! Here is exactly what is saved:");
      result.rows.forEach((row, index) => {
        const vals = row.f.map(field => field.v);
        console.log(`Row ${index + 1}: Date: ${vals[0]} | Market: ${vals[1]} | Type: ${vals[2]} | Item: ${vals[3]} | Sell: ${vals[4]}`);
      });
    }
  } catch(e) {
    console.error("Query failed: " + e.message);
  }
}

function forceStartEngine() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('LAST_FUZZ_FETCH');
  props.deleteProperty('fuzz_lease_timestamp');
  props.deleteProperty('fuzz_job_active');
  props.deleteProperty('exec_start_time');
  
  console.log("🛑 All safety timers cleared.");
  console.log("🚀 Forcing full Fuzzwork fetch and BigQuery upload...");
  masterOrchestrator();
}

function getLatestPricesQuery(marketId) {
  // Ensure the project ID 'tenacious-tiger-345318' is correct here
  return `SELECT * FROM \`tenacious-tiger-345318.market_data.market_prices\`
          WHERE market_id = ${marketId}
          AND date > TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 24 HOUR)
          ORDER BY date DESC`;
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