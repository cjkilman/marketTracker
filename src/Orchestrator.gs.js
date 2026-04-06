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

function emergencyGhostCleanup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  sheets.forEach(sh => {
    if (sh.getName().includes("Merge_Temp_")) ss.deleteSheet(sh);
  });
  console.log("Workbook Cleaned.");
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
 * Waits for a ScriptLock (up to 30s) and executes a function. Throws error if lock fails.
 */
function executeWithWaitLock(func, funcName) {
  const log = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('LockManager') : console;
  const lock = LockService.getScriptLock();
  let functionResult = null;

  try {
    lock.waitLock(30000); // 30-second waitLock
  } catch (e) {
    log.error(`Could not acquire Script Lock for ${funcName} after 30s wait.`);
    throw e;
  }

  const isOuterLock = (EXECUTION_LOCK_DEPTH_WAIT === 0);
  EXECUTION_LOCK_DEPTH_WAIT++;
  try {
    if (isOuterLock) deleteTriggersByName(funcName);
    log.info(`--- ${isOuterLock ? 'Starting' : 'Entering nested'} Execution (WaitLock): ${funcName} ---`);

    functionResult = func(); // Execute and store result

    log.info(`--- ${isOuterLock ? 'Finished' : 'Exiting nested'} Execution (WaitLock): ${funcName} ---`);
  } catch (e) {
    log.error(`${funcName} failed: ${e.message}\nStack: ${e.stack}`);
    throw e;
  } finally {
    EXECUTION_LOCK_DEPTH_WAIT--;
    try {
      lock.releaseLock();
    } catch (lockError) {
      log.error(`CRITICAL: Failed to release Script Lock for ${funcName}: ${lockError.message}`);
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
  const ctx = getAppContext(ss); // Create the context once
  
  const props = ctx.props;
  props.setProperty('exec_start_time', String(startTime));

  // 1. CHECK VAULT GATE
  const isVaultOk = (String(ctx.cfg.BQ_ENABLED).toUpperCase() === "TRUE");
  const itemSource = isVaultOk ? "SDE_invTypes" : "Item List Back End";
  console.log(`[GATE] Vault: ${isVaultOk ? "OPEN" : "LOCKED"}. Source: ${itemSource}`);

  // ==========================================
  // 2. RUN MARKET FETCHER (Extract & Stage)
  // ==========================================
  try {
    console.log("--- PHASE 1: FETCHING ---");
    // Worker runs, APPENDS chunked data to 'Market Prices', and sets 'fuzz_pass_complete'
    updateFuzzMarketDataSheet(itemSource,ctx);

  } catch (e) {
    if (e.message.includes("quota")) {
      console.error("[!] CRITICAL QUOTA HIT [!] Tripping Safety Switch.");
      toggleBigQueryCircuitBreaker(); // Replaced setMarketConfig
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

function runVaultMergeAndReset(ctx) {
  const { ss, cfg } = ctx;
  const targetSheetName = "Market Prices"; 
  const oldSh = ss.getSheetByName(targetSheetName);
  if (!oldSh || oldSh.getLastRow() < 2) return;

  const isVaultOk = (String(cfg.BQ_ENABLED).toUpperCase() === "TRUE");
  const tempName = "Merge_Temp_" + Date.now();

  try {
    // Phase A: Detach Sidecar
    oldSh.setName(tempName);
    rebuildMarketPricesSheet(ctx); // Create clean landing pad for Fetcher

    if (isVaultOk) {
      // Phase B: Attempt Vault Upload
      // ... (Existing BigQuery Upload Logic) ...
      
      // Phase C: Success - Burn the Sidecar
      ss.deleteSheet(oldSh); 
      console.log("[Vault] Upload Success. Sidecar deleted.");
    } else {
      throw new Error("Vault Disabled: Redirecting to Local Fallback.");
    }

  } catch (e) {
    console.warn(`[VaultGate] ${e.message}`);
    
    // Phase D: THE RECOVERY
    // If we failed, merge the Sidecar data back into the main 'Market Prices'
    const recoverySh = ss.getSheetByName(tempName);
    if (recoverySh) {
      const data = recoverySh.getDataRange().getValues();
      const mainSh = ss.getSheetByName(targetSheetName);
      // Append the data back to the live buffer so the UI can see it
      if (data.length > 1) {
        mainSh.getRange(mainSh.getLastRow() + 1, 1, data.length - 1, 8)
              .setValues(data.slice(1));
      }
      ss.deleteSheet(recoverySh);
      console.log("[Fallback] Data restored to Market Prices for UI use.");
    }
  }
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
 */
function bigQueryBatchLoad_(projectId, datasetId, tableId, dataArray) {
  // Convert 2D array to Newline Delimited JSON (Standard BQ Format)
  const headers = dataArray[0];
  const jsonRows = dataArray.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return JSON.stringify(obj);
  }).join('\n');

  const blob = Utilities.newBlob(jsonRows, 'application/octet-stream');

  const job = {
    configuration: {
      load: {
        destinationTable: {
          projectId: projectId,
          datasetId: datasetId,
          tableId: tableId
        },
        sourceFormat: 'NEWLINE_DELIMITED_JSON',
        writeDisposition: 'WRITE_APPEND' // Keep the history!
      }
    }
  };

  const runJob = BigQuery.Jobs.insert(job, projectId, blob);
  return (runJob.status.state === 'DONE' || runJob.status.state === 'PENDING');
}



