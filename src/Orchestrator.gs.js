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
 * THE ENDGAME ORCHESTRATOR
 * The single source of truth for the 132-slot farm.
 */
function masterOrchestrator() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('exec_start_time', String(startTime)); 

  // --- PRE-FLIGHT: THE SAFETY COP ---
  const cfg = getConfig(); 
  const isVaultOk = (cfg.BQ_ENABLED === true);
  
  // SELECT SOURCE: SDE (Industrial) vs Item List (Tactical Backup)
  const itemSource = isVaultOk ? "SDE_Master_List" : "Item List Back End";
  console.log(`[GATE] Vault: ${isVaultOk ? "OPEN" : "LOCKED"}. Source: ${itemSource}`);

  // --- TASK 1: MARKET INGESTION ---
  try {
    console.log("--- PHASE 1: FETCHING ---");
    // This function now APPENDS to 'BQ_LIVE_DATA'
    updateFuzzMarketDataSheet(itemSource); 

    // ANESTHESIA: Force Google to commit the new rows before the Gate
    console.log("[ANESTHESIA] Flushing fetch results...");
    SpreadsheetApp.flush(); 

  } catch (e) {
    if (e.message.includes("quota")) {
      console.error("!!! CRITICAL QUOTA HIT !!! Tripping Safety Switch.");
      setMarketConfig("BQ_ENABLED", false); // Hard Stop on B10
      return; // Exit: Path B will take over at the next 3:00 AM reset cycle
    }
    console.warn("Fetch Failed: " + e.message);
  }

  // --- TASK 2: THE VAULT GATE (Merge vs. Prune) ---
  if (hasFuel(120)) { 
    console.log("--- PHASE 2: VAULT GATE ---");
    if (isVaultOk) {
      console.log("[PATH A] Merging Sheet to BigQuery Vault...");
      runVaultMergeAndReset(ss); 
    } else {
      console.log("[PATH B] BigQuery Locked. Pruning to Rolling 24h...");
      pruneSheetToRolling24(ss); 
    }
    
    // ANESTHESIA: Force the "Clean Slate" before we pulse the UI
    console.log("[ANESTHESIA] Flushing Gate results...");
    SpreadsheetApp.flush();
  }

// --- THE STAGGERED HANDOFF ---
  // Instead of running Task 3 now, we schedule it for 1 minute from now.
  // This allows BigQuery to "settle" and resets our execution timer.
  console.log("[STAGGER] Task 2 Complete. Scheduling UI Pulse for +1 Minute...");
  scheduleOneTimeTrigger('orchestratorTaskUI', 60000);
}

/**
 * TASK 3: UI DISTRIBUTION
 * Runs in its own execution window with dedicated fuel monitoring.
 */
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
 * TASK 2: VAULT GATE (PATH A)
 * Merges the 'Market Prices' sheet data into BigQuery and wipes the buffer.
 */
function runVaultMergeAndReset(ss) {
  const LOG_HEADER = '[VaultGate]';
  const cfg = getConfig();
  const sheetName = "Market Prices"; // FUZZ_SHEET_FINAL
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

  // 1. FUEL CHECK (Need a full tank for BigQuery operations)
  if (!hasFuel(120)) {
    console.warn(`${LOG_HEADER} Low fuel. Aborting merge to prevent partial data loss.`);
    return;
  }

  try {
    // 2. PREPARE THE DATA
    // Grab everything including headers to ensure schema alignment
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    console.log(`${LOG_HEADER} Preparing to bury ${rows.length} rows in the Vault...`);

    // 3. BIGQUERY LOAD JOB
    // We use a LOAD job because it's free and handles batching better than streaming for large sets
    const projectId = cfg.BQ_PROJECT_ID; 
    const datasetId = cfg.BQ_DATASET_ID;
    const tableId = cfg.BQ_TABLE_ID;

    if (!projectId || !datasetId || !tableId) {
      throw new Error("Missing BigQuery Configuration (Project/Dataset/Table ID).");
    }

    const blob = Utilities.newBlob(JSON.stringify(rows), 'application/octet-stream');
    
    // Using the BigQuery Pipe logic (Batch Load)
    const jobSuccess = bigQueryBatchLoad_(projectId, datasetId, tableId, data);

    if (jobSuccess) {
      console.log(`${LOG_HEADER} Merge Successful. BigQuery confirmed receipt.`);
      
      // 4. THE RESET (Only happens on success!)
      // Leave the headers, kill the data.
      const lastCol = sh.getLastColumn();
      sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      
      // OPTIONAL: Trim the sheet back down to 100 rows to keep it snappy
      _trimTrailing_(sh);
      
      console.log(`${LOG_HEADER} Sheet reset. Buffer is clean.`);
    } else {
      throw new Error("BigQuery Job failed to verify success signal.");
    }

  } catch (e) {
    console.error(`${LOG_HEADER} CRITICAL FAILURE: ${e.message}`);
    // WE DO NOT RESET THE SHEET HERE. 
    // This allows Path B (Pruning) to handle the data on the next pulse if the Vault is down.
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



