// Add this function (or ensure it exists) in MarketFetcher.gs.js or Orchestrator.gs.js

/**
 * Wrapper function for the cache warmer.
 * Attempts to run the cache warmer using executeWithTryLock.
 * If skipped due to lock, it schedules a one-time retry trigger for itself.
 * If completed fully, logs completion.
 */
function triggerCacheWarmerWithRetry() {
  const SCRIPT_PROP = PropertiesService.getScriptProperties();
  const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('CacheWarmer') : console);

  // --- Check if Price Snapshot Job is Finalizing (if applicable) ---
  // If your Fuzz job had a finalizing step, check it here. Since we removed it, this might not be needed.
  // const fuzzStep = SCRIPT_PROP.getProperty('fuzzJobStep');
  // if (fuzzStep === STATE_FLAGS.FINALIZING) { // Assuming STATE_FLAGS is global or accessible
  //   LOG.warn("Cache Warmer: Skipping execution, Fuzz Price job is finalizing.");
  //   return;
  // }

  const funcToRun = fuzzworkCacheRefresh_TimeGated; // Assumes this function exists
  const funcName = 'fuzzworkCacheRefresh_TimeGated';
  const wrapperFuncName = 'triggerCacheWarmerWithRetry';

  const retryDelayMs = 2 * 60 * 1000; // 2 minutes retry delay

  LOG.info(`Wrapper ${wrapperFuncName} called. Attempting to run ${funcName}...`);

  // Assumes executeWithTryLock is globally available (from Orchestrator.gs.js)
  const result = executeWithTryLock(funcToRun, funcName); // result is true (full run), false (incomplete), or null (skipped)

  if (result === null) {
    // --- Case 1: Skipped due to Script Lock ---
    LOG.warn(`${funcName} was skipped due to Script Lock. Scheduling retry for ${wrapperFuncName}.`);
    scheduleOneTimeTrigger(wrapperFuncName, retryDelayMs); // Assumes scheduleOneTimeTrigger is global

  } else if (result === true) {
    // --- Case 2: Ran AND Completed Fully ---
    LOG.info(`${funcName} completed a full run successfully.`);
    // Cache warmer is done, no need to trigger price snapshot here, orchestrator handles timing.
    // Clear the lease for the cache warmer
    SCRIPT_PROP.deleteProperty('cacheWarmerLeaseUntil');

  } else if (result === false) {
    // --- Case 3: Ran but did NOT complete fully (hit time limit) ---
    LOG.info(`${funcName} ran but hit its time limit and rescheduled itself.`);
    // Lease remains, inner function handles rescheduling.

  } else {
    // --- Case 4: Unexpected return value ---
    LOG.warn(`${funcName} returned unexpected value: ${result}`);
    // Clear lease on unexpected outcome
     SCRIPT_PROP.deleteProperty('cacheWarmerLeaseUntil');
  }
}

/**
 * Cache refresh function (placeholder - ensure your actual implementation exists).
 * Processes the Fuzzworks cache queue in batches within time limits.
 */
function fuzzworkCacheRefresh_TimeGated() {
    const LOG = (typeof LoggerEx !== 'undefined' ? LoggerEx.withTag('CacheWarmerCore') : console);
    const SCRIPT_PROP = PropertiesService.getScriptProperties();
    const START_TIME = Date.now();
    const TIME_LIMIT_MS = 270000; // 4m 30s
    const PROP_KEY_RESUME = 'cacheRefresh_lastIndex';
    const PROP_KEY_COOLDOWN = 'cacheRefresh_lastFullCompletion'; // Tracks last full run
    const SUB_BATCH_SIZE = 2500; // How many items to process per execution
    const COOLDOWN_MINUTES = 10; // Minimum time between full runs

    LOG.info("Starting Fuzzworks cache refresh cycle (Time Gated)...");
    let completedFullRun = false;

    try {
        // --- Cooldown Check ---
        const lastCompletionTime = parseInt(SCRIPT_PROP.getProperty(PROP_KEY_COOLDOWN) || '0', 10);
        const cooldownMs = COOLDOWN_MINUTES * 60 * 1000;
        if (START_TIME < lastCompletionTime + cooldownMs) {
            LOG.info(`Cache refresh cooldown active (last run finished ${((START_TIME - lastCompletionTime)/60000).toFixed(1)} min ago). Skipping.`);
            // IMPORTANT: Clear the lease if skipping due to cooldown, otherwise orchestrator thinks it's running
            SCRIPT_PROP.deleteProperty('cacheWarmerLeaseUntil');
            return true; // Return true as if completed, to prevent retry loops
        }

        // --- Get Requests ---
        // !! IMPORTANT: Ensure 'getMasterBatchFromControlTable' is defined globally or imported !!
        if (typeof getMasterBatchFromControlTable !== 'function') {
            throw new Error("Dependency 'getMasterBatchFromControlTable' is missing.");
        }
        const allRequests = getMasterBatchFromControlTable();
        if (!allRequests || allRequests.length === 0) {
            LOG.info("Cache Refresh: Control Table empty. Resetting state.");
            SCRIPT_PROP.deleteProperty(PROP_KEY_RESUME);
            SCRIPT_PROP.deleteProperty(PROP_KEY_COOLDOWN); // Clear cooldown too
             SCRIPT_PROP.deleteProperty('cacheWarmerLeaseUntil'); // Clear lease
            return true;
        }

        // --- Determine Starting Point ---
        const resumeIndexRaw = SCRIPT_PROP.getProperty(PROP_KEY_RESUME);
        let startIndex = resumeIndexRaw ? parseInt(resumeIndexRaw, 10) : 0;
        if (isNaN(startIndex) || startIndex < 0 || startIndex >= allRequests.length) startIndex = 0;
        if (startIndex === 0) LOG.info(`Cache refresh starting/restarting from index 0.`);
        else LOG.info(`Cache refresh resuming from index ${startIndex}.`);

        let itemsProcessedThisRun = 0;

        // --- Processing Loop ---
        while (startIndex < allRequests.length) {
            // Time Limit Check
            if (Date.now() - START_TIME > TIME_LIMIT_MS) {
                SCRIPT_PROP.setProperty(PROP_KEY_RESUME, startIndex.toString());
                LOG.warn(`⚠️ Cache refresh time limit hit after ${itemsProcessedThisRun} items. Next run starts at index ${startIndex}. RESCHEDULING SELF.`);
                scheduleOneTimeTrigger('triggerCacheWarmerWithRetry', 30 * 1000); // Reschedule the *wrapper*
                return false; // Did not complete fully
            }

            // Process sub-batch
            const endIndex = Math.min(startIndex + SUB_BATCH_SIZE, allRequests.length);
            const currentSubBatch = allRequests.slice(startIndex, endIndex);
            if (currentSubBatch.length > 0) {
                 LOG.info(`Processing cache refresh batch: Indices ${startIndex} to ${endIndex - 1}`);
                try {
                    // Call the core fuzAPI function that handles cache checks and fetches
                    fuzAPI.getDataForRequests(currentSubBatch); // Assumes fuzAPI is global
                    itemsProcessedThisRun += currentSubBatch.length;
                } catch (apiError) {
                     LOG.error(`Error refreshing cache batch indices ${startIndex}-${endIndex - 1}: ${apiError.message}. Skipping batch.`);
                     // Optionally add retry logic for the specific batch here
                }
            }
            startIndex = endIndex;
        } // End while

        // --- Full Completion ---
        SCRIPT_PROP.deleteProperty(PROP_KEY_RESUME);
        SCRIPT_PROP.setProperty(PROP_KEY_COOLDOWN, START_TIME.toString()); // Use START_TIME as completion time
        LOG.info(`Cache refresh: Successfully processed all ${allRequests.length} items. Cooldown set. Index reset.`);
        completedFullRun = true;

    } catch (e) {
        LOG.error(`Unhandled error during cache refresh: ${e.message}\nStack: ${e.stack}`);
        completedFullRun = false;
        // Consider whether to reset PROP_KEY_RESUME on error or let it retry
    } finally {
        const duration = (Date.now() - START_TIME) / 1000;
        LOG.info(`Cache refresh execution block finished in ${duration.toFixed(2)}s. Full run completed: ${completedFullRun}`);
        // Lease is cleared by the wrapper (triggerCacheWarmerWithRetry) upon completion.
    }
    return completedFullRun;
}