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

function rebuildMarketPricesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const HEADERS = ["date", "market_id", "market_type", "type_id", "min_sell", "max_buy", "median_sell", "median_buy"];
  
  // Creates the sheet, adds headers, and deletes columns I-Z automatically
  const sheet = getOrCreateSheet(ss, 'Market Prices', HEADERS);
  
  console.log("Market Prices sheet rebuilt and optimized!");
}



function forceStartEngine() {
  const props = PropertiesService.getScriptProperties();
  
  // 1. Clear the "Safety Tape"
  props.deleteProperty('LAST_FUZZ_FETCH'); 
  props.deleteProperty('fuzz_lease_timestamp');
  props.deleteProperty('fuzz_job_active');
  
  console.log("🛑 Safety timers cleared.");
  console.log("🚀 Forcing full fetch and BigQuery upload...");
  
  // 2. Fire the Orchestrator
  masterOrchestrator(); 
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

/**
 * FUEL GAUGE: Returns true if we have enough time to start a new task.
 * @param {number} requiredSeconds - Minimum buffer needed (default 30s)
 */
function hasFuel(requiredSeconds = 30) {
  const startTime = PropertiesService.getScriptProperties().getProperty('exec_start_time');
  if (!startTime) return true; // Fallback if not set
  
  const elapsed = (Date.now() - parseInt(startTime)) / 1000;
  const limit = 360; // Google's 6-minute limit
  
  return (limit - elapsed) > requiredSeconds;
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
  const props = PropertiesService.getScriptProperties();
  props.setProperty('exec_start_time', String(startTime)); 

  // ==========================================
  // 1. CHECK VAULT GATE (Pre-Flight)
  // ==========================================
  const cfg = getConfig(); 
  const isVaultOk = (String(cfg.BQ_ENABLED).toUpperCase() === "TRUE");
  
  // Select Source based on Gate Status
  const itemSource = isVaultOk ? "SDE_invTypes" : "Item List Back End";
  console.log(`[GATE] Vault: ${isVaultOk ? "OPEN" : "LOCKED"}. Source: ${itemSource}`);

  // ==========================================
  // 2. RUN MARKET FETCHER (Extract & Stage)
  // ==========================================
  try {
    console.log("--- PHASE 1: FETCHING ---");
    // Worker runs, APPENDS chunked data to 'Market Prices', and sets 'fuzz_pass_complete'
    updateFuzzMarketDataSheet(itemSource); 

    console.log("[ANESTHESIA] Flushing staged results...");
    SpreadsheetApp.flush(); 
  } catch (e) {
    if (e.message.includes("quota")) {
      console.error("!!! CRITICAL QUOTA HIT !!! Tripping Safety Switch.");
      setMarketConfig("BQ_ENABLED", false); 
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
      runVaultMergeAndReset(ss); 
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
    
    console.log("[ANESTHESIA] Flushing Gate results...");
    SpreadsheetApp.flush();
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
 * ROUND-ROBIN MAINTENANCE SCHEDULER
 * Executes one garbage collection or maintenance chore per cycle to save RAM and prevent timeouts.
 */
function runMaintenanceJobs() {
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const NOW_MS = new Date().getTime();

  // 1. Job Registry with targeted intervals (in milliseconds)
  const JOB_QUEUE = [
    { name: 'dailyHeavyPrune_Prices', interval: 86400000 }, // 24 hours
    { name: 'dailyJobReset', interval: 86400000 },          // 24 hours
    { name: 'emergencyCleanup', interval: 604800000 }       // 7 days (Weekly deep clean)
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
        // This prevents the V8 engine from blocking dynamic eval() calls.
        const jobFunctions = {
          'dailyHeavyPrune_Prices': () => { if (typeof dailyHeavyPrune_Prices === 'function') dailyHeavyPrune_Prices(); },
          'dailyJobReset': () => { if (typeof dailyJobReset === 'function') dailyJobReset(); },
          'emergencyCleanup': () => { if (typeof emergencyCleanup === 'function') emergencyCleanup(); }
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
/**
 * TASK 3: UI DISTRIBUTION (THE HEARTBEAT + THE LIVE WIRE)
 * Runs in its own execution window after BigQuery has settled.
 */
function orchestratorTaskUI() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('exec_start_time', String(startTime)); 

  console.log("--- PHASE 3: UI PULSE ---");

  // 1. Update Regional Volumes (System A)
  if (hasFuel(120)) { 
    console.log("[UI] Step 1: Updating Volume Heartbeat...");
    ESI_publishVolumeInterfaces(ss); 
  }

  // 2. Distribute Everything (Consolidated Sync)
  if (hasFuel(90)) {
    console.log("[UI] Step 2: Slicing data for Client Interfaces...");
    FUZ_publishPriceInterfaces(ss); 
  } else {
    console.warn("[STAGGER] Low fuel. Rescheduling UI pulse.");
    scheduleOneTimeTrigger('orchestratorTaskUI', 30000);
  }
}

/**
 * THE VAULT BYPASS
 * Pushes the currently staged 162k rows into BigQuery and clears the sheet.
 */
function forceVaultMerge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log("[LAUNCH] Manually triggering Vault Merge for staged data...");
  
  // THE FIX: Reset the stopwatch so the fuel gauge reads 100%
  PropertiesService.getScriptProperties().setProperty('exec_start_time', String(Date.now()));
  
  // Drive the truck to the Vault
  runVaultMergeAndReset(ss); 
  
  console.log("[SUCCESS] Vault Merge Complete. Pipeline is now fully operational.");
}

function runVaultMergeAndReset(ss) {
  const LOG_HEADER = '[VaultGate]';
  const cfg = getConfig();
  const sheetName = "Market Prices"; 
  const sh = ss.getSheetByName(sheetName);

  if (!sh) {
    console.error(`${LOG_HEADER} Error: ${sheetName} not found.`);
    return;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    console.log(`${LOG_HEADER} Sheet is empty. Nothing to merge.`);
    return;
  }

  if (!hasFuel(120)) {
    console.warn(`${LOG_HEADER} Low fuel. Aborting merge to prevent partial data loss.`);
    return;
  }

  try {
    const rawData = sh.getDataRange().getValues();
    let ndjson = "";

    // Start at index 1 to skip the header row
    for (let i = 1; i < rawData.length; i++) {
      let row = rawData[i];
      let dateCell = row[0];

      // Format date to UTC strict
      let formattedDate = "";
      if (dateCell instanceof Date) {
        formattedDate = Utilities.formatDate(dateCell, "UTC", "yyyy-MM-dd HH:mm:ss");
      } else if (typeof dateCell === 'string') {
        let parsed = new Date(dateCell);
        if (!isNaN(parsed)) {
           formattedDate = Utilities.formatDate(parsed, "UTC", "yyyy-MM-dd HH:mm:ss");
        } else {
           formattedDate = dateCell; 
        }
      }

      // Explicitly map every column to the exact BigQuery schema name
      const bqRow = {
        date: formattedDate,
        market_id: String(row[1]),
        market_type: String(row[2]),
        type_id: Number(row[3]),
        min_sell: Number(row[4]) || 0,
        max_buy: Number(row[5]) || 0,
        median_sell: Number(row[6]) || 0,
        median_buy: Number(row[7]) || 0
      };

      ndjson += JSON.stringify(bqRow) + "\n";
    }

    console.log(`${LOG_HEADER} Preparing to bury ${rawData.length - 1} explicitly mapped rows in the Vault...`);

    const projectId = 'tenacious-tiger-345318'; 
    const datasetId = 'market_data';            
    const tableId   = 'market_prices_staged';   

    const blob = Utilities.newBlob(ndjson, 'application/octet-stream');

    const job = {
      configuration: {
        load: {
          destinationTable: {
            projectId: projectId,
            datasetId: datasetId,
            tableId: tableId
          },
          sourceFormat: 'NEWLINE_DELIMITED_JSON',
          writeDisposition: 'WRITE_APPEND'
        }
      }
    };

    // Fire the payload directly via the Advanced Service, ignoring the helper function
    BigQuery.Jobs.insert(job, projectId, blob);

    console.log(`${LOG_HEADER} Merge Successful. BigQuery confirmed receipt.`);
    
    const lastCol = sh.getLastColumn();
    sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    console.log(`${LOG_HEADER} Sheet reset. Buffer is clean.`);

  } catch (e) {
    console.error(`${LOG_HEADER} CRITICAL FAILURE: ${e.message}`);
  }
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



