/**
 * Orchestrator for MarketTracker jobs. Manages execution flow, concurrency,
 * and state persistence using PropertiesService and Locks.
 */

// --- Global Lock Depth Counters ---
var EXECUTION_LOCK_DEPTH_TRY = 0;

// --- State Machine Constants ---
const STATE_FLAGS = {
  NEW_RUN: 'NEW_RUN',
  PROCESSING: 'PROCESSING',
  COMPLETE: 'COMPLETE'
};

// --- Lease Duration ---
const JOB_LEASE_DURATION_MS = 280000; // 4 minutes 40 seconds

/** Creates a one-time trigger. */
function scheduleOneTimeTrigger(functionName, delayMs) { /* ... implementation ... */
  if (typeof functionName !== 'string' || functionName.trim() === '') {
    throw new Error(`CRITICAL SCHEDULER ERROR: Invalid function name: ${functionName}`);
  }
  try {
    deleteTriggersByName(functionName);
    ScriptApp.newTrigger(functionName).timeBased().after(delayMs).create();
    console.log(`Scheduled one-time trigger for ${functionName} in ~${Math.round(delayMs / 60000)} min.`);
  } catch (e) {
    console.error(`Failed to schedule trigger for ${functionName}: ${e.message}.`);
  }
}

/** Deletes triggers by name. */
function deleteTriggersByName(functionName) { /* ... implementation ... */
  if (typeof functionName !== 'string' || functionName.trim() === '') return 0;
  let deletedCount = 0;
  try {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === functionName && trigger.getEventType() === ScriptApp.EventType.CLOCK) {
        try { ScriptApp.deleteTrigger(trigger); deletedCount++; }
        catch (e) { console.warn(`Could not delete trigger (ID: ${trigger.getUniqueId()}) for ${functionName}: ${e.message}`); }
      }
    });
  } catch (e) { console.error(`Error accessing/deleting triggers for ${functionName}: ${e.message}.`); }
  return deletedCount;
}

/** Executes a function with ScriptLock (tryLock 30s). */
function executeWithTryLock(func, funcName) { /* ... implementation ... */
  const lock = LockService.getScriptLock();
  let functionResult = null;
  if (lock.tryLock(30000)) {
    const isOuterLock = (EXECUTION_LOCK_DEPTH_TRY === 0);
    EXECUTION_LOCK_DEPTH_TRY++;
    try {
      if (isOuterLock) deleteTriggersByName(funcName);
      console.log(`--- ${isOuterLock ? 'Starting' : 'Entering nested'} Execution (TryLock): ${funcName} ---`);
      functionResult = func();
      console.log(`--- ${isOuterLock ? 'Finished' : 'Exiting nested'} Execution (TryLock): ${funcName} ---`);
    } catch (e) {
      console.error(`${funcName} failed: ${e.message}\nStack: ${e.stack}`);
      throw e;
    } finally {
      EXECUTION_LOCK_DEPTH_TRY--;
      try { lock.releaseLock(); }
      catch (lockError) { console.error(`CRITICAL: Failed to release Script Lock for ${funcName}: ${lockError.message}`); }
    }
    return functionResult;
  } else {
    console.warn(`${funcName} skipped: Script Lock busy.`);
    return null; // Indicate skip
  }
}

/**
 * Master orchestrator - runs every 15 minutes.
 * Alternates between Cache Refresh and Price Snapshot Update.
 */
function masterOrchestrator() {
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const currentMinute = new Date().getMinutes();
  const NOW_MS = Date.now();
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('Orchestrator') : console);

  LOG.info(`Master orchestrator running (Minute: ${currentMinute}).`);

  // --- Determine which job to run based on the 15-minute interval ---
  // Simple alternation: 0-14 & 30-44 for Cache, 15-29 & 45-59 for Snapshot
  const runCacheWarmer = (currentMinute >= 0 && currentMinute < 15) || (currentMinute >= 30 && currentMinute < 45);
  const runPriceSnapshot = (currentMinute >= 15 && currentMinute < 30) || (currentMinute >= 45 && currentMinute < 60);

  // --- Job Leases ---
  const fuzzLeaseUntil = parseInt(SCRIPT_PROP.getProperty('fuzzJobLeaseUntil') || '0', 10);
  const cacheLeaseUntil = parseInt(SCRIPT_PROP.getProperty('cacheWarmerLeaseUntil') || '0', 10); // Use a separate lease for the warmer
  const isFuzzJobActive = fuzzLeaseUntil > NOW_MS;
  const isCacheWarmerActive = cacheLeaseUntil > NOW_MS;

  // --- Check Expired Leases ---
  if (!isFuzzJobActive && fuzzLeaseUntil > 0) {
      LOG.warn(`Fuzz job lease expired. Clearing lease.`);
      SCRIPT_PROP.deleteProperty('fuzzJobLeaseUntil');
      // _resetFuzzMarketDataJobState(new Error("Lease expired")); // Optional reset
  }
  if (!isCacheWarmerActive && cacheLeaseUntil > 0) {
      LOG.warn(`Cache Warmer lease expired. Clearing lease.`);
      SCRIPT_PROP.deleteProperty('cacheWarmerLeaseUntil');
      // Cache warmer manages its own resume state, usually no need to reset fully.
  }

  // --- Dispatch Logic ---
  if (runCacheWarmer) {
      LOG.info("In Cache Warmer window.");
      if (isCacheWarmerActive) {
          LOG.info(`Cache Warmer is already active. Skipping dispatch.`);
      } else if (isFuzzJobActive) {
          LOG.warn(`Cache Warmer window, but Fuzz Price Snapshot job is active. Skipping Cache Warmer dispatch.`);
      } else {
          LOG.info("Dispatching Cache Warmer.");
          const newLease = NOW_MS + JOB_LEASE_DURATION_MS; // Use standard lease
          SCRIPT_PROP.setProperty('cacheWarmerLeaseUntil', newLease.toString());
          LOG.info(`Set Cache Warmer lease until ${new Date(newLease)}`);
          // Call the wrapper which uses TryLock and handles retries/rescheduling
          const result = executeWithTryLock(triggerCacheWarmerWithRetry, 'triggerCacheWarmerWithRetry');
           if (result === null) {
              LOG.warn("Cache Warmer dispatch skipped by ScriptLock. Lease remains.");
              // Lease will prevent immediate re-dispatch by next orchestrator run.
              // triggerCacheWarmerWithRetry's internal logic might schedule a retry if needed.
           }
      }
  } else if (runPriceSnapshot) {
      LOG.info("In Price Snapshot Update window.");
      if (isFuzzJobActive) {
          LOG.info(`Fuzz Price Snapshot job is already active. Skipping dispatch.`);
      } else if (isCacheWarmerActive) {
           LOG.warn(`Price Snapshot window, but Cache Warmer job is active. Skipping Price Snapshot dispatch.`);
      } else {
          LOG.info("Dispatching Fuzz Price Snapshot Update.");
          const newLease = NOW_MS + JOB_LEASE_DURATION_MS;
          SCRIPT_PROP.setProperty('fuzzJobLeaseUntil', newLease.toString());
          LOG.info(`Set Fuzz job lease until ${new Date(newLease)}`);
          const result = updateFuzzMarketDataSheet(); // Calls wrapper
          if (result === null) {
              LOG.warn("Fuzz Price Snapshot dispatch skipped by ScriptLock. Lease remains.");
          }
      }
  } else {
      // Should not happen with the current minute logic, but good failsafe.
      LOG.warn(`No job scheduled for current minute (${currentMinute}).`);
  }

  LOG.info("Master orchestrator finished.");
}


// --- REMOVED triggerDailyHeavyPrunes ---
// --- REMOVED triggerEsiHistoryUpdate ---


/**
 * Run ONCE to set up the main orchestrator trigger. Clears old ones first.
 */
function setupOrchestratorTrigger() { // Renamed for clarity
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('Setup') : console);
  LOG.info("Setting up/Resetting the main Orchestrator trigger...");

  const ORCHESTRATOR_FUNCTION = 'masterOrchestrator';
  const TRIGGER_INTERVAL_MINUTES = 15; // Keep the 15-min check cycle

  // --- Clear ALL existing time-based triggers ---
  let deletedCount = 0;
  try {
    ScriptApp.getProjectTriggers().forEach(trigger => {
       if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
            try { ScriptApp.deleteTrigger(trigger); deletedCount++; }
            catch(e) { LOG.warn(`Could not delete trigger ID ${trigger.getUniqueId()}: ${e.message}`);}
       }
    });
     LOG.info(`Deleted ${deletedCount} existing clock triggers.`);
  } catch (e) {
     LOG.error(`Error accessing triggers for cleanup: ${e.message}`);
     SpreadsheetApp.getUi().alert(`Error clearing old triggers: ${e.message}`);
     return;
  }

  // --- Setup the single Master Orchestrator Trigger ---
  try {
    ScriptApp.newTrigger(ORCHESTRATOR_FUNCTION)
      .timeBased()
      .everyMinutes(TRIGGER_INTERVAL_MINUTES)
      .create();
    LOG.info(`SUCCESS: Created ${TRIGGER_INTERVAL_MINUTES}-minute trigger for ${ORCHESTRATOR_FUNCTION}.`);
    SpreadsheetApp.getUi().alert(`Orchestrator trigger set to run every ${TRIGGER_INTERVAL_MINUTES} minutes.`);

  } catch (e) {
    LOG.error(`Failed to create orchestrator trigger: ${e.message}.`);
    SpreadsheetApp.getUi().alert(`Failed to create orchestrator trigger: ${e.message}`);
  }
}