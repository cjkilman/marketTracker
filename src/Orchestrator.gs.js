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
const JOB_LEASE_DURATION_MS = 300000; // 5 minutes

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
 * REVISED: Master orchestrator (the "pipper") triggered every 15 minutes.
 * This function now implements the "bump" logic by checking job leases.
 * It will "bump" (start) any job whose lease has expired.
 */
function masterOrchestrator() {
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('Orchestrator') : console);
  LOG.info("Master orchestrator running. Checking job leases...");

  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const NOW_MS = Date.now();

  const ESI_STAGNATION_TIMEOUT_MS = 30 * 60 * 1000; // Define a 30-minute threshold for no progress

  // --- 1. Check Fuzz Snapshot Job (MarketFetcher.gs.js) ---
  const fuzzLease = parseInt(SCRIPT_PROP.getProperty('fuzzJobLeaseUntil') || '0', 10);
  if (fuzzLease > NOW_MS) {
    LOG.warn(`Fuzz job is already active (Lease expires in ${((fuzzLease - NOW_MS) / 60000).toFixed(1)} min). Skipping dispatch.`);
  } else {
    LOG.info("Fuzz job is not active. Dispatching 'bump' (starting updateFuzzMarketDataSheet).");
    const newFuzzLease = NOW_MS + JOB_LEASE_DURATION_MS; // Give it a new 5 min lease
    SCRIPT_PROP.setProperty('fuzzJobLeaseUntil', newFuzzLease.toString());
    updateFuzzMarketDataSheet(); // Call the job
  }

  // --- 2. Check "ESI" Job (marketFetcherEsi.js) - HANG DETECTION LOGIC ---
  const esiActive = SCRIPT_PROP.getProperty('mf_job_active');
  const currentCursor = SCRIPT_PROP.getProperty('mf_cursor');
  const lastCheckedCursor = SCRIPT_PROP.getProperty('mf_hang_check_cursor');
  const lastCheckTime = parseInt(SCRIPT_PROP.getProperty('mf_hang_check_time_ms') || '0', 10);

  if (esiActive === '1') {
    const isFirstCheck = lastCheckedCursor === null;
    
    // Stagnation is detected if the cursor hasn't moved AND the check time exceeds 30 minutes
    const isStagnated = (currentCursor === lastCheckedCursor && NOW_MS - lastCheckTime > ESI_STAGNATION_TIMEOUT_MS);
    
    if (isFirstCheck) {
        LOG.info("ESI Job is active. Starting progress monitoring.");
    } else if (isStagnated) {
        LOG.error(`ESI Job HANG DETECTED! Cursor (${currentCursor}) has not advanced for over 30 minutes. Forcing state reset.`);
        
        // Execute the necessary cleanup and reset
        if (typeof _resetEsiHistoryJobState === 'function') {
            _resetEsiHistoryJobState(new Error("Progress Stagnation Detected by Orchestrator"));
        } else {
            LOG.error("Cannot reset ESI state: _resetEsiHistoryJobState function is missing.");
        }
    } 
    
    // Always update the check properties for the NEXT 15-minute cycle
    if (currentCursor !== null) {
        SCRIPT_PROP.setProperty('mf_hang_check_cursor', currentCursor);
        SCRIPT_PROP.setProperty('mf_hang_check_time_ms', String(NOW_MS));
    }

    LOG.warn(`ESI job is active and progressing. Skipping dispatch.`);
}

  LOG.info("Master orchestrator finished lease checks.");
}