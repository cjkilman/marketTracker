/** Orchestrator.gs
 * Manages execution flow, concurrency, and state persistence
 * using PropertiesService and Locks.
 * * ScriptLock for Concurrency
 * Sheet Lock for Data Integrity
 */

/* global LoggerEx, PropertiesService, ScriptApp, LockService, SpreadsheetApp, 
   FUZ_publishPriceInterfaces, ESI_publishVolumeInterfaces, getConfig */

// --- Global Lock Depth Counters ---
var EXECUTION_LOCK_DEPTH_TRY = 0;
var EXECUTION_LOCK_DEPTH_WAIT = 0;

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
    const log = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('Scheduler') : console;
    log.info(`Scheduled one-time trigger for ${functionName} in ~${Math.round(delayMs / 1000)}s.`);
  } catch (e) {
    console.error(`Failed to schedule trigger for ${functionName}: ${e.message}. Stack: ${e.stack}`);
  }
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
  // GLOBAL CONFIGURATION: ScriptLock only. DocumentLock is incompatible with this architecture.
  const lock = LockService.getScriptLock();
  let functionResult = null;

  if (lock.tryLock(30000)) { // 30-second tryLock
    const isOuterLock = (EXECUTION_LOCK_DEPTH_TRY === 0);
    EXECUTION_LOCK_DEPTH_TRY++;
    try {
      if (isOuterLock) deleteTriggersByName(funcName); 
      log.info(`--- ${isOuterLock ? 'Starting' : 'Entering nested'} Execution (TryLock): ${funcName} ---`);

      functionResult = func(); 

      log.info(`--- ${isOuterLock ? 'Finished' : 'Exiting nested'} Execution (TryLock): ${funcName} ---`);
    } catch (e) {
      log.error(`${funcName} failed: ${e.message}\nStack: ${e.stack}`);
      throw e; 
    } finally {
      EXECUTION_LOCK_DEPTH_TRY--;
      try {
        lock.releaseLock();
      } catch (lockError) {
        log.error(`CRITICAL: Failed to release Script Lock for ${funcName}: ${lockError.message}`);
      }
    }
    return functionResult; 
  } else {
    log.warn(`${funcName} skipped: Script Lock busy.`);
    return null; 
  }
}

/**
 * Executes a function within a strict ScriptLock.
 * Prevents parallel job collisions.
 */
function executeWithWaitLock(func, funcName) {
  const log = (typeof LoggerEx !== 'undefined') ? LoggerEx.withTag('LockManager') : console;
  const lock = LockService.getScriptLock();
  const LOCK_TIMEOUT_MS = 30000;
  let functionResult = null;

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
    if (isOuterLock && typeof deleteTriggersByName === 'function') {
      deleteTriggersByName(funcName);
    }

    log.info(`--- ${isOuterLock ? 'START' : 'NESTED'} : ${funcName} ---`);

    functionResult = func();

    log.info(`--- ${isOuterLock ? 'FINISH' : 'EXIT'} : ${funcName} ---`);
    
  } catch (e) {
    log.error(`[EXECUTION ERROR] ${funcName}: ${e.message}`);
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

/** Helper to check if we have enough execution time left (in seconds) */
function hasFuel(secondsNeeded) {
  const props = PropertiesService.getScriptProperties();
  const startTime = parseInt(props.getProperty('exec_start_time') || Date.now());
  const elapsedSeconds = (Date.now() - startTime) / 1000;
  return (300 - elapsedSeconds) > secondsNeeded; // 300s = 5 minutes (Leaves 1 min safety buffer)
}

/**
 * THE ENDGAME ORCHESTRATOR
 * Lightweight loop conductor. Schedules UI distributions directly.
 */
function masterOrchestrator() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ctx = getAppContext(ss);

  ctx.props.setProperty('exec_start_time', String(startTime));
  console.log("--- ORCHESTRATOR LOOP PULSE ---");

  // Directly prompt the UI layer update task
  console.log("[STAGGER] Scheduling UI Pulse for +1 Minute...");
  scheduleOneTimeTrigger('orchestratorTaskUI', 60000);

  // Maintenance Check
  if (hasFuel(60)) {
    console.log("--- MAINTENANCE CHECK ---");
    runMaintenanceJobs();
  } else {
    console.warn("[MAINTENANCE] Skipped: Insufficient execution fuel remaining.");
  }
}

/**
 * ROUND-ROBIN MAINTENANCE SCHEDULER
 * Retained clean interface shell for future zero-cost tasks.
 */
function runMaintenanceJobs() {
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const NOW_MS = new Date().getTime();

  // All legacy database staging / heavy pruning tasks removed
  const JOB_QUEUE = [];

  const QUEUE_INDEX_KEY = 'MAINTENANCE_QUEUE_INDEX';
  let currentIndex = parseInt(SCRIPT_PROP.getProperty(QUEUE_INDEX_KEY) || '0', 10);
  if (currentIndex >= JOB_QUEUE.length) currentIndex = 0;

  let iterations = 0;

  while (iterations < JOB_QUEUE.length) {
    const job = JOB_QUEUE[currentIndex];
    const lastRunKey = 'LAST_RUN_' + job.name;
    const lastRunTs = parseInt(SCRIPT_PROP.getProperty(lastRunKey) || '0', 10);
    const isDue = (NOW_MS - lastRunTs) >= job.interval;

    if (isDue) {
      console.log(`[MAINTENANCE] Dispatching: ${job.name}`);
      try {
        const jobFunctions = {};

        if (jobFunctions[job.name]) {
          jobFunctions[job.name](); 
          SCRIPT_PROP.setProperty(lastRunKey, NOW_MS.toString());
          SCRIPT_PROP.setProperty(QUEUE_INDEX_KEY, ((currentIndex + 1) % JOB_QUEUE.length).toString());
          console.log(`[MAINTENANCE] ${job.name} completed successfully.`);
          return; 
        } else {
          console.warn(`[MAINTENANCE] Function ${job.name} is not defined in the workspace.`);
        }
      } catch (e) {
        console.error(`[MAINTENANCE] Critical Failure in ${job.name}: ${e.message}`);
      }
    }
    currentIndex = (currentIndex + 1) % JOB_QUEUE.length;
    iterations++;
  }
  console.log("[MAINTENANCE] Cycle Complete: No active maintenance tasks pending.");
}

/**
 * MANUAL RESET
 * Clears active UI update limits to force fresh processing on the next loop execution.
 */
function manualUIReset() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('LAST_PRICE_PULSE');
  props.deleteProperty('LAST_VOL_PULSE');
  console.log("[RESET] UI Leases cleared. Next pulse will force a full sync.");
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

  // --- CONFIG DRIFT PRE-CHECK ---
  let forcePriceSync = false;
  const priceSheets = ['filtered prices', 'Mineral Supply Prices', 'T1 Supply Prices'];
  
  priceSheets.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    
    const mktIdRaw = sh.getRange("C4").getValue();
    const mktTypeRaw = sh.getRange("D4").getValue();
    
    // Guard clause: Skip checking if cells are currently empty, broken, or loading
    if (mktIdRaw === "" || String(mktIdRaw).indexOf('#') === 0 || String(mktTypeRaw).toLowerCase().includes('loading')) {
      return;
    }
    
    // Construct a unique snapshot string for the current market configuration
    const currentVal = String(mktIdRaw).trim() + '|' + String(mktTypeRaw).trim().toLowerCase();
    const propKey = 'UI_CONF_' + sheetName.replace(/\s+/g, '_');
    const lastVal = props.getProperty(propKey);
    
    // If we have a historical value and it does not match what is currently on the sheet, trip the gate
    if (lastVal && lastVal !== currentVal) {
      console.log(`[DRIFT] Configuration change on sheet "${sheetName}": [${lastVal}] -> [${currentVal}]`);
      forcePriceSync = true;
    }
    
    // Cache the updated value back to script properties
    props.setProperty(propKey, currentVal);
  });

  // --- TASK 1: PRICE INTERFACES ---
  const lastPriceSync = parseInt(props.getProperty('LAST_PRICE_PULSE') || '0', 10);
  
  // Slices immediately if a configuration drift was flagged OR the 30-minute lease expired
  if (forcePriceSync || (now - lastPriceSync > LEASE_MS)) {
    if (hasFuel(120)) { 
      console.log("[UI] Step 1: Slicing data for Client Interfaces...");
      FUZ_publishPriceInterfaces(ss);
      props.setProperty('LAST_PRICE_PULSE', String(Date.now()));
    } else {
      console.warn("[UI] Low fuel for Price update. Rescheduling pulse.");
      scheduleOneTimeTrigger('orchestratorTaskUI', 45000);
      return; 
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

/**
 * FORCE TRIGGER
 * Instantly fires the orchestrator loop chain.
 */
function forceStartEngine() {
  console.log("[RESET] Manual engine activation override. Forcing full loop execution...");
  masterOrchestrator();
}