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

/**
 * Helper to ensure we don't spam the ESI publish
 */
function isTimeForInterfaceSync() {
  const lastSync = parseInt(SCRIPT_PROPS.getProperty('last_esi_sync') || '0', 10);
  const now = new Date().getTime();
  
  // Only sync once every 5 minutes
  if (now - lastSync > 300000) {
    SCRIPT_PROPS.setProperty('last_esi_sync', String(now));
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

function isTimeForInterfaceSync() {
  const lastSync = parseInt(SCRIPT_PROPS.getProperty('last_interface_sync') || '0', 10);
  const now = new Date().getTime();
  // If this property was never set, it stays 0, and (now - 0) is always > 300000.
  // BUT, if it was set to a time in the future by mistake, it will never run.
  return (now - lastSync) > 300000; 
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

function masterOrchestrator() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('exec_start_time', String(Date.now())); // Mark the start

  // --- TASK 1: FUZZ ---
  updateFuzzMarketDataSheet();

// --- TASK 2: SYNC ---
  // We bump the fuel requirement slightly (60s) because we are doing two steps now.
  if (hasFuel(60)) { 
    console.log("Fuel Good. Syncing Interfaces...");
    
    // STEP A: Pull data from Cache -> Publish Sheet
    // This ensures 'Publish_ESI_Region' is never empty, even if the worker is running.
    if (typeof publishMarketResultESIRegion === 'function') {
      try {
        publishMarketResultESIRegion(); 
        console.log("Intermediate Publish Table Updated.");
      } catch (e) {
        console.warn("Intermediate Publish Failed (Skipping Sync): " + e.message);
      }
    }

    // STEP B: Push data from Publish Sheet -> Client Sheets
    ESI_publishClientInterfaces();
    
  } else {
    console.warn("Low Fuel! Skipping Sync to avoid hard timeout.");
  }

  // --- TASK 3: REFRESH ---
  if (hasFuel(30)) { // Need 30s buffer for Refreshes
    console.log("Fuel Good. Running Refreshes...");
    masterMarketRefresh();
  }
}


