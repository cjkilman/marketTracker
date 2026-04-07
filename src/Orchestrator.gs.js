/** Orchestrator.gs
 * Manages execution flow, concurrency, and state persistence
 * using PropertiesService and Locks.
 */

/* global LoggerEx, PropertiesService, ScriptApp, LockService, SpreadsheetApp, 
   updateFuzzMarketDataSheet, updateEsiHistorySheet */

// --- Constants ---
const STATE_FLAGS = {
  NEW_RUN: 'NEW_RUN',
  PROCESSING: 'PROCESSING',
  FINALIZING: 'FINALIZING',
  COMPLETE: 'COMPLETE'
};
const JOB_LEASE_DURATION_MS = 1800000; // 5 minutes

// Global lock depth counters
var EXECUTION_LOCK_DEPTH_TRY = 0;
var EXECUTION_LOCK_DEPTH_WAIT = 0;


/**
 * GHOST RECOVERY
 * Locates orphaned Merge_Temp_ sheets and attempts to push their 
 * stranded data into the BigQuery Vault before deleting them.
 */
function emergencyGhostCleanup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = getAppContext(ss);

  const isVaultOk = (String(ctx.cfg.BQ_ENABLED).toUpperCase() === "TRUE");
  if (!isVaultOk) {
    console.warn("[RECOVERY] Vault is currently LOCKED. Cannot merge ghosts at this time.");
    return;
  }

  const projectId = ctx.cfg.BQ_PROJECT_ID;
  const datasetId = ctx.cfg.BQ_DATASET_ID;
  const tableId = ctx.cfg.BQ_Market_Prices;

  const sheets = ss.getSheets();
  let ghostsRecovered = 0;

  sheets.forEach(sh => {
    const name = sh.getName();

    if (name.includes("Merge_Temp_")) {
      console.log(`[RECOVERY] Found ghost sheet: ${name}. Extracting payload...`);
      const data = sh.getDataRange().getValues();

      // Only upload if there is actual data (more than just the header row)
      if (data.length > 1) {
        try {
          console.log(`[RECOVERY] Transmitting ${data.length - 1} rows to Vault...`);

          // Push directly to BigQuery using the helper
          const success = bigQueryBatchLoad_(projectId, datasetId, tableId, data);

          if (success) {
            console.log(`[RECOVERY] Upload successful. Burning sidecar: ${name}`);
            ss.deleteSheet(sh);
            ghostsRecovered++;
          } else {
            console.error(`[RECOVERY] Vault rejected payload for ${name}. Sheet retained.`);
          }
        } catch (e) {
          console.error(`[RECOVERY] Failed to merge ${name}: ${e.message}`);
        }
      } else {
        // Sheet is empty or only has headers, burn it safely without uploading
        console.log(`[RECOVERY] ${name} is empty. Burning sidecar.`);
        ss.deleteSheet(sh);
      }
    }
  });

  console.log(`[RECOVERY] Workbook Cleaned. Salvaged ${ghostsRecovered} ghost payloads.`);
}

function rebuildMarketPricesSheet(ctx) {
  const targetSS = ctx ? ctx.ss : SpreadsheetApp.getActiveSpreadsheet();
  const NAME = "Market Prices";
  const HEADERS = [["date", "market_id", "market_type", "type_id", "min_sell", "max_buy", "median_sell", "median_buy"]];

  let sh = targetSS.getSheetByName(NAME);
  if (!sh) {
    sh = targetSS.insertSheet(NAME);
    sh.getRange(1, 1, 1, 8).setValues(HEADERS);
  }
  return sh;
}




/**

 * Ensures all functions share the same Spreadsheet and Config handles.
 */
function getAppContext(ss) {
  return {
    ss: ss || SpreadsheetApp.getActiveSpreadsheet(),
    props: PropertiesService.getScriptProperties(),
    cfg: getConfig(),
    timestamp: Date.now()
  };
}

/**
 * Creates a one-time trigger for a function after a delay. Deletes existing triggers first.
 */
function scheduleOneTimeTrigger(functionName, delayMs) {
  if (typeof functionName !== 'string' || functionName.trim() === '') {
    throw new Error(`CRITICAL SCHEDULER ERROR: Invalid function name: ${functionName}`);
  }
  try {
    deleteTriggersByName(functionName); // Prevent duplicates
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .after(delayMs)
      .create();
    // Use LoggerEx if available, otherwise console.log
    const log = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('Scheduler') : console;
    log.info(`Scheduled one-time trigger for ${functionName} in ~${Math.round(delayMs / 1000)}s.`);
  } catch (e) {
    console.error(`Failed to schedule trigger for ${functionName}: ${e.message}. Stack: ${e.stack}`);
  }
}

function isTimeForInterfaceSync() {
  const props = PropertiesService.getScriptProperties();
  const lastSync = parseInt(props.getProperty('last_interface_sync') || '0', 10);
  const now = new Date().getTime();

  // Only sync once every 5 minutes
  if (now - lastSync > 300000) {
    props.setProperty('last_interface_sync', String(now));
    return true;
  }
  return false;
}

/**
 * Deletes all time-based triggers for a specific function name.
 */
function deleteTriggersByName(functionName) {
  if (typeof functionName !== 'string' || functionName.trim() === '') return 0;
  let deletedCount = 0;
  try {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === functionName && trigger.getEventType() === ScriptApp.EventType.CLOCK) {
        try {
          ScriptApp.deleteTrigger(trigger);
          deletedCount++;
        } catch (e) {
          console.warn(`Could not delete trigger (ID: ${trigger.getUniqueId()}) for ${functionName}: ${e.message}`);
        }
      }
    });
    if (deletedCount > 0) {
      // Use LoggerEx if available
      const log = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('Scheduler') : console;
      log.info(`Deleted ${deletedCount} existing clock trigger(s) for ${functionName}.`);
    }
  } catch (e) {
    console.error(`Error accessing/deleting triggers for ${functionName}: ${e.message}.`);
  }
  return deletedCount;
}

/**
 * Attempts to lock (ScriptLock, 30s try) and execute a function. Skips if locked.
 * Returns function result on success, null on skip. Manages nested calls.
 */
function executeWithTryLock(func, funcName) {
  const log = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('LockManager') : console;
 // --- GLOBAL CONFIGURATION ---
// STRICT: Use ScriptLock only. DocumentLock is incompatible with this architecture.
  const lock = LockService.getScriptLock();
  let functionResult = null;

  if (lock.tryLock(30000)) { // 30-second tryLock
    const isOuterLock = (EXECUTION_LOCK_DEPTH_TRY === 0);
    EXECUTION_LOCK_DEPTH_TRY++;
    try {
      if (isOuterLock) deleteTriggersByName(funcName); // Clean retry trigger if we got the lock
      log.info(`--- ${isOuterLock ? 'Starting' : 'Entering nested'} Execution (TryLock): ${funcName} ---`);

      functionResult = func(); // Execute and store result

      log.info(`--- ${isOuterLock ? 'Finished' : 'Exiting nested'} Execution (TryLock): ${funcName} ---`);
    } catch (e) {
      log.error(`${funcName} failed: ${e.message}\nStack: ${e.stack}`);
      throw e; // Re-throw error to halt execution if needed
    } finally {
      EXECUTION_LOCK_DEPTH_TRY--;
      try {
        lock.releaseLock();
      } catch (lockError) {
        log.error(`CRITICAL: Failed to release Script Lock for ${funcName}: ${lockError.message}`);
      }
    }
    return functionResult; // Return actual result
  } else {
    log.warn(`${funcName} skipped: Script Lock busy.`);
    return null; // Indicate skip
  }
}

/**
 * Executes a function within a strict ScriptLock.
 * Prevents "DocumentLock" drift and parallel job collisions.
 * * @param {Function} func - The function to execute.
 * @param {string} funcName - Name for logging/trigger management.
 * @return {*} - Returns the result of the executed function.
 */
function executeWithWaitLock(func, funcName) {
  // Fallback for LoggerEx to keep it ASCII/Console based
  const log = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('LockManager') : console;
  
  // STRICT: ScriptLock only. Do not use DocumentLock.
  const lock = LockService.getScriptLock();
  const LOCK_TIMEOUT_MS = 30000;
  let functionResult = null;

  // Initialize depth counter if undefined
  if (typeof EXECUTION_LOCK_DEPTH_WAIT === 'undefined') {
    var EXECUTION_LOCK_DEPTH_WAIT = 0; 
  }

  try {
    lock.waitLock(LOCK_TIMEOUT_MS);
  } catch (e) {
    log.error(`[LOCK FAIL] ${funcName} could not acquire ScriptLock after ${LOCK_TIMEOUT_MS/1000}s.`);
    throw new Error(`Lock Timeout: ${funcName}`);
  }

  const isOuterLock = (EXECUTION_LOCK_DEPTH_WAIT === 0);
  EXECUTION_LOCK_DEPTH_WAIT++;

  try {
    // If this is the primary entry point, clear any redundant triggers
    if (isOuterLock && typeof deleteTriggersByName === 'function') {
      deleteTriggersByName(funcName);
    }

    log.info(`--- ${isOuterLock ? 'START' : 'NESTED'} : ${funcName} ---`);

    // Execute the actual logic
    functionResult = func();

    log.info(`--- ${isOuterLock ? 'FINISH' : 'EXIT'} : ${funcName} ---`);
    
  } catch (e) {
    log.error(`[EXECUTION ERROR] ${funcName}: ${e.message}`);
    // No .Flush() here - let the error bubble up naturally
    throw e;
  } finally {
    EXECUTION_LOCK_DEPTH_WAIT--;
    try {
      lock.releaseLock();
    } catch (lockError) {
      log.error(`CRITICAL: Lock release failed for ${funcName}`);
    }
  }

  return functionResult;
}
/**
 * Helper to check if it's time to refresh the Google Sheet displays
 */
function isTimeForDisplayRefresh() {
  // Logic: Always return true to refresh whenever the Orchestrator pulses, 
  // or add a timer check (e.g., every 5 minutes).
  return true;
}

/** Helper to check if we have enough execution time left (in seconds) */
function hasFuel(secondsNeeded) {
  const props = PropertiesService.getScriptProperties();
  const startTime = parseInt(props.getProperty('exec_start_time') || Date.now());
  const elapsedSeconds = (Date.now() - startTime) / 1000;
  return (300 - elapsedSeconds) > secondsNeeded; // 300s = 5 minutes (Leaves 1 min safety buffer)
}



/**
 * THE RESUME BUTTON
 * Clears ONLY the 30-minute cooldown lease, leaving the Bookmark intact, 
 * then fires the Orchestrator to finish the pass.
 */
function resumeEngine() {
  const props = PropertiesService.getScriptProperties();

  // Clear ONLY the cooldown timers
  props.deleteProperty('fuzz_lease_timestamp');
  props.deleteProperty('LAST_FUZZ_FETCH');

  console.log("🛑 Cooldown bypassed. Resuming from Bookmark...");

  // Fire the engine
  masterOrchestrator();
}


/**
 * THE ENDGAME ORCHESTRATOR
 * Implements the strict ETL (Extract, Stage, Load) Pipeline.
 */
function masterOrchestrator() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = getAppContext(ss);

  const props = ctx.props;
  props.setProperty('exec_start_time', String(startTime));

  // ==========================================
  // 1. MANDATORY RECOVERY: Vault Merge & Clean
  // ==========================================
  const isVaultOk = (String(ctx.cfg.BQ_ENABLED).toUpperCase() === "TRUE");

  if (isVaultOk) {
    const targetSheetName = ctx.cfg.MarketPricesSheet || "Market Prices";
    const stagingSheet = ctx.ss.getSheetByName(targetSheetName);
    const lastRow = stagingSheet ? stagingSheet.getLastRow() : 0;

    // If Vault is enabled and we have data in the staging area, 
    // we MUST merge and clean before doing anything else.
    if (lastRow > 1) {
      console.log("[RECOVERY] BigQuery Quota Reset detected. Merging staged data to Vault...");
      try {
        runVaultMergeAndReset(ctx);
        console.log("[RECOVERY] Vault Merge and Clean successful.");
      } catch (e) {
        console.error("[RECOVERY ERROR] Failed to merge backlog: " + e.message);
        // Exit to prevent over-stuffing the local sheet if merge fails
        return;
      }
    }
  }

  const itemSource = isVaultOk ? "SDE_invTypes" : "Item List Back End";
  console.log(`[GATE] Vault: ${isVaultOk ? "OPEN" : "LOCKED"}. Source: ${itemSource}`);

  // ==========================================
  // 2. RUN MARKET FETCHER (Extract & Stage)
  // ==========================================
  try {
    console.log("--- PHASE 1: FETCHING ---");
    // Worker runs, APPENDS chunked data to 'Market Prices', and sets 'fuzz_pass_complete'
    updateFuzzMarketDataSheet(itemSource, ctx);

  } catch (e) {
    if (e.message.includes("quota")) {
      console.error("[!] CRITICAL QUOTA HIT [!] Tripping Safety Switch.");
      toggleBigQueryCircuitBreaker();
      return;
    }
    console.warn("Fetch Failed: " + e.message);
  }

  // ==========================================
  // 3. VAULT GATE (Merge vs. Prune)
  // ==========================================
  if (hasFuel(120)) {
    console.log("--- PHASE 2: VAULT GATE ---");

    // Read the signal from the worker to ensure it finished all markets
    const isPassComplete = (props.getProperty('fuzz_pass_complete') === 'true');

    if (isVaultOk && isPassComplete) {
      // * OK: Big Q Merge Reset Sheet
      console.log("[PATH A] Full Pass Confirmed. Merging to Vault...");
      runVaultMergeAndReset(ctx);
      props.deleteProperty('fuzz_pass_complete'); // Reset signal for next cycle
    }
    else if (!isVaultOk) {
      // * Locked: Trim Sheet Items with 24 Hour Expirettes
      console.log("[PATH B] BigQuery Locked. Pruning local sheet to Rolling 24h...");
      pruneSheetToRolling24(ss);
    }
    else {
      // Safety: Vault is open, but the pass didn't finish (timed out/errored).
      console.warn("[SKIP] Merge Aborted: Market pass was partial. Data remains in staging.");
    }
  }

  // ==========================================
  // 4. SET TRIGGER UI REFRESH
  // ==========================================
  console.log("[STAGGER] Task 2 Complete. Scheduling UI Pulse for +1 Minute...");
  scheduleOneTimeTrigger('orchestratorTaskUI', 60000);

  // ==========================================
  // 5. MAINTENANCE & GARBAGE COLLECTION
  // ==========================================
  if (hasFuel(60)) {
    console.log("--- PHASE 4: MAINTENANCE ---");
    runMaintenanceJobs();
  } else {
    console.warn("[Maintenance] Skipped: Insufficient execution fuel remaining.");
  }
}

/**
 * MAINTENANCE: Automatically resets the BigQuery Circuit Breaker.
 * Intended to run once every 24 hours to recover from daily quota limits.
 */
function autoResetCircuitBreaker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Market Config");

  if (!configSheet) {
    console.error("[MAINTENANCE] Failed to run Circuit Breaker reset: Config sheet missing.");
    return;
  }

  const data = configSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "BQ_ENABLED") {
      const isCurrentlyEnabled = (data[i][1] === true || String(data[i][1]).toLowerCase() === "true");

      if (!isCurrentlyEnabled) {
        console.warn("[MAINTENANCE] Circuit Breaker was tripped. Auto-resetting to ENABLED for the new day.");
        configSheet.getRange(i + 1, 2).setValue(true);

        // Optional: Clear the safety tape so the next orchestrator pulse does a full fresh pull
        PropertiesService.getScriptProperties().deleteProperty('LAST_FUZZ_FETCH');
      } else {
        console.log("[MAINTENANCE] Circuit Breaker is already ENABLED. No action needed.");
      }
      return;
    }
  }
}

/**
 * Helper to flip the BQ_ENABLED flag from the menu OR background triggers
 */
function toggleBigQueryCircuitBreaker(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Market Config");

  if (!configSheet) {
    console.error("[CIRCUIT BREAKER] Failed: 'Market Config' sheet not found.");
    return;
  }

  const data = configSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === "BQ_ENABLED") {
      const currentValue = data[i][1];

      // Determine the new state
      const newValue = (currentValue === true || String(currentValue).toLowerCase() === "true") ? false : true;

      // Write the new state back to the sheet
      configSheet.getRange(i + 1, 2).setValue(newValue);

      // Log it for the background execution logs
      const status = newValue ? "[ENABLED]" : "[DISABLED]";
      console.warn(`[CIRCUIT BREAKER] BigQuery Pipe is now ${status}`);

      // Safe Toast: Only attempt to show a UI toast if the context allows it
      try {
        ss.toast(`BigQuery Pipe is now ${status}`, "Circuit Breaker");
      } catch (e) {
        // Silently ignore UI errors when called by a background time-trigger
      }
      return;
    }
  }

  console.error("[CIRCUIT BREAKER] Failed: BQ_ENABLED setting not found.");
  try {
    ss.toast("BQ_ENABLED setting not found in Market Config.", "Error");
  } catch (e) { }
}

function runVaultMergeAndReset(ctx, targetSheetName = "Market Prices") {
  const { ss, cfg } = ctx;
  const sh = ss.getSheetByName(targetSheetName);

  if (!sh || sh.getLastRow() < 2) return;

  const isVaultOk = (String(cfg.BQ_ENABLED).toUpperCase() === "TRUE");
  if (!isVaultOk) return;

  try {
    const data = sh.getDataRange().getValues();
    const success = bigQueryBatchLoad_(cfg.BQ_PROJECT_ID, cfg.BQ_DATASET_ID, cfg.BQ_Market_Prices, data);

    if (success) {
      console.log(`[VaultGate] Success! Clearing landing pad...`);

      // RE-FETCH the row count right here to avoid "Out of Bounds" errors
      const currentRowCount = sh.getLastRow();
      const rowsToDelete = currentRowCount - 1;

      if (rowsToDelete > 0) {
        try {
          sh.deleteRows(2, rowsToDelete);
          console.log(`[VaultGate] Successfully cleared ${rowsToDelete} rows.`);
        } catch (e) {
          // If maintenance already cleared it, we just ignore the error and move on
          console.warn("[VaultGate] Sheet was modified during upload. Skipping manual clear.");
        }
      }
    }
  } catch (e) {
    console.error("Vault Error: " + e.message);
  }
}



function KILL_STUCK_PRUNE() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('heavyPruneJobStep');
  props.deleteProperty('heavyPruneJobReadRow');
  console.log("Heavy Prune state cleared.");
}

/**
 * ROUND-ROBIN MAINTENANCE SCHEDULER
 * Executes one garbage collection or maintenance chore per cycle to save RAM and prevent timeouts.
 */
function runMaintenanceJobs() {
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const NOW_MS = new Date().getTime();

  // 1. Job Registry with targeted intervals (in milliseconds)
  const JOB_QUEUE = [
    { name: 'dailyHeavyPrune_Prices', interval: 86400000 },
    { name: 'dailyJobReset', interval: 86400000 },
    { name: 'autoResetCircuitBreaker', interval: 86400000 }, // NEW: 24 hour reset
    { name: 'emergencyCleanup', interval: 604800000 }
  ];

  const QUEUE_INDEX_KEY = 'MAINTENANCE_QUEUE_INDEX';
  let currentIndex = parseInt(SCRIPT_PROP.getProperty(QUEUE_INDEX_KEY) || '0', 10);
  if (currentIndex >= JOB_QUEUE.length) currentIndex = 0;

  let iterations = 0;

  // 2. Loop through the queue until we find one job that is due
  while (iterations < JOB_QUEUE.length) {
    const job = JOB_QUEUE[currentIndex];
    const lastRunKey = 'LAST_RUN_' + job.name;
    const lastRunTs = parseInt(SCRIPT_PROP.getProperty(lastRunKey) || '0', 10);
    const isDue = (NOW_MS - lastRunTs) >= job.interval;

    // 3. Execution Logic
    if (isDue) {
      console.log(`[Maintenance] Dispatching: ${job.name}`);

      try {
        // Safer Dispatcher: Maps string names directly to the global functions
        const jobFunctions = {
          'dailyHeavyPrune_Prices': () => { if (typeof dailyHeavyPrune_Prices === 'function') dailyHeavyPrune_Prices(); },
          'dailyJobReset': () => { if (typeof dailyJobReset === 'function') dailyJobReset(); },
          'emergencyCleanup': () => { if (typeof emergencyCleanup === 'function') emergencyCleanup(); },
          'autoResetCircuitBreaker': () => { if (typeof autoResetCircuitBreaker === 'function') autoResetCircuitBreaker(); } // NEW: Map the function
        };

        if (jobFunctions[job.name]) {
          jobFunctions[job.name](); // Execute the chore

          // Mark it complete and advance the queue index for the next Orchestrator run
          SCRIPT_PROP.setProperty(lastRunKey, NOW_MS.toString());
          SCRIPT_PROP.setProperty(QUEUE_INDEX_KEY, ((currentIndex + 1) % JOB_QUEUE.length).toString());
          console.log(`[Maintenance] ${job.name} completed successfully.`);

          return; // EXIT EARLY: Only do ONE chore per Orchestrator pulse to prevent Google timeouts.
        } else {
          console.warn(`[Maintenance] Function ${job.name} is not defined in the workspace.`);
        }
      } catch (e) {
        console.error(`[Maintenance] Critical Failure in ${job.name}: ${e.message}`);
      }
    }

    // Move to the next item if the current one isn't due
    currentIndex = (currentIndex + 1) % JOB_QUEUE.length;
    iterations++;
  }

  console.log("[Maintenance] Cycle Complete: All jobs are currently within their wait windows.");
}

function manualUIReset() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('LAST_PRICE_PULSE');
  props.deleteProperty('LAST_VOL_PULSE');
  console.log("🛠 UI Leases cleared. Next pulse will force a full sync.");
}

/**
 * TASK 3: UI DISTRIBUTION (RE-PRIORITIZED)
 * Price Interfaces are now the First Strike.
 */
function orchestratorTaskUI() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = Date.now();
  const LEASE_MS = 30 * 60 * 1000;

  props.setProperty('exec_start_time', String(now));
  console.log("--- PHASE 3: UI PULSE (Priority: Prices) ---");

  // --- TASK 1: PRICE INTERFACES (Moved to Front) ---
  // Targets: 'filtered prices', 'Mineral Supply Prices', etc.
  const lastPriceSync = parseInt(props.getProperty('LAST_PRICE_PULSE') || '0', 10);
  if (now - lastPriceSync > LEASE_MS) {
    if (hasFuel(120)) { // Higher fuel priority for the main interface
      console.log("[UI] Step 1: Slicing data for Client Interfaces...");
      FUZ_publishPriceInterfaces(ss);
      props.setProperty('LAST_PRICE_PULSE', String(Date.now()));
    } else {
      console.warn("[UI] Low fuel for Price update. Rescheduling pulse.");
      scheduleOneTimeTrigger('orchestratorTaskUI', 45000);
      return; // Exit early to save remaining fuel for the next attempt
    }
  } else {
    console.log("[UI] Price lease active. Skipping.");
  }

  // --- TASK 2: REGIONAL VOLUMES ---
  const lastVolSync = parseInt(props.getProperty('LAST_VOL_PULSE') || '0', 10);
  if (now - lastVolSync > LEASE_MS) {
    if (hasFuel(60)) {
      console.log("[UI] Step 2: Updating Volume Heartbeat...");
      ESI_publishVolumeInterfaces(ss);
      props.setProperty('LAST_VOL_PULSE', String(Date.now()));
    } else {
      console.warn("[UI] Low fuel for Volume update. Skipping.");
    }
  } else {
    console.log("[UI] Volume lease active. Skipping.");
  }
}

function forceStartEngine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = getAppContext(ss); // Handshake
  const props = ctx.props;

  props.deleteProperty('LAST_FUZZ_FETCH');
  props.deleteProperty('fuzz_lease_timestamp');
  props.deleteProperty('fuzz_job_active');

  console.log("🛑 Safety timers cleared. 🚀 Forcing full fetch...");
  masterOrchestrator();
}

function forceVaultMerge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = getAppContext(ss); // Handshake
  console.log("[LAUNCH] Manually triggering Vault Merge...");

  ctx.props.setProperty('exec_start_time', String(Date.now()));
  runVaultMergeAndReset(ctx); // Pass the ctx

  console.log("[SUCCESS] Vault Merge Complete.");
}

/**
 * INTERNAL HELPER: Executes the BigQuery Load Job
 * OPTIMIZED: Uses "Streaming Inserts" (Tabledata.insertAll) instead of Load Jobs.
 * Bypasses the file-upload proxy that causes 503s and chunks data to save memory.
 */
function bigQueryBatchLoad_(projectId, datasetId, tableId, dataArray) {
  const headers = dataArray[0];
  const rows = dataArray.slice(1);
  const CHUNK_SIZE = 2500; // Extremely safe memory limit for Apps Script

  console.log(`[Vault] Initiating streaming insert of ${rows.length} rows to ${tableId}...`);

  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);

    // Build the specific JSON format required for Streaming Inserts
    const insertAllRequest = {
      skipInvalidRows: false,
      ignoreUnknownValues: true, // Prevents crashes if your sheet has an extra blank column
      rows: chunk.map(row => {
        let obj = {};
        headers.forEach((h, index) => {
          let val = row[index];
          // Format dates strictly for BigQuery
          if (val instanceof Date) {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
          }
          // Only add the field if the header actually has a name
          if (h && String(h).trim() !== "") {
            obj[String(h).trim()] = val;
          }
        });
        return { json: obj };
      })
    };

    let chunkSuccess = false;
    const maxRetries = 3;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        const response = BigQuery.Tabledata.insertAll(insertAllRequest, projectId, datasetId, tableId);

        // Streaming inserts do not crash the script on bad data; they return the exact error.
        if (response.insertErrors && response.insertErrors.length > 0) {
          const exactError = response.insertErrors[0].errors[0].message;
          console.error(`[Vault] Schema Error in chunk: ${exactError}`);
          return false; // Abort so we don't lose the sheet
        }

        chunkSuccess = true;
        break; // Chunk succeeded, break out of the retry loop

      } catch (e) {
        if (e.message.includes("503") || e.message.includes("Service Unavailable") || e.message.includes("timeout")) {
          console.warn(`[Vault] API hiccup on chunk ${i}. Retrying attempt ${attempt}...`);
          if (attempt === maxRetries) throw e;
          Utilities.sleep(Math.pow(2, attempt) * 1000);
        } else {
          console.error(`[Vault] Hard API Error: ${e.message}`);
          throw e; // Hard error (like Permission Denied)
        }
      }
    }

    if (!chunkSuccess) return false;

    // Force garbage collection / memory flush between chunks
    Utilities.sleep(200);
  }

  return true; // All chunks succeeded
}



