/* eslint-disable no-console */
/* eslint-disable no-unused-vars */

/**
 * SDE_Job_Controller.gs
 * Stateful, multi-step SDE import job resilient to execution limits.
 */

// --- Safely define global constants ---
if (typeof KEY_JOB_RUNNING === 'undefined') { var KEY_JOB_RUNNING = 'SDE_JOB_RUNNING'; }
if (typeof KEY_JOB_LIST === 'undefined') { var KEY_JOB_LIST = 'SDE_JOB_LIST'; }
if (typeof KEY_JOB_INDEX === 'undefined') { var KEY_JOB_INDEX = 'SDE_JOB_INDEX'; }
if (typeof KEY_BACKUP_SETTINGS === 'undefined') { var KEY_BACKUP_SETTINGS = 'SDE_BACKUP_SETTINGS'; }
if (typeof GLOBAL_STATE_KEY === 'undefined') { var GLOBAL_STATE_KEY = 'GLOBAL_SYSTEM_STATE'; }
if (typeof KEY_JOB_CHUNK_INDEX === 'undefined') { var KEY_JOB_CHUNK_INDEX = 'SDE_JOB_CHUNK_INDEX'; }

var SS;
function getSS() {
  if (!SS) { SS = SpreadsheetApp.getActiveSpreadsheet(); }
  return SS;
}

// -----------------------------------------------------------------------------
// --- SDE ENGINE LIBRARY (sdeLib) ---
// -----------------------------------------------------------------------------

const sdeLib = () => {
  let _sheetCache = {};

  const downloadTextData = (csvFile) => {
    console.time("downloadTextData( csvFile:" + csvFile + " )");
    const baseURL = 'https://raw.githubusercontent.com/cjkilman/eve-sde-dump/main/' + csvFile;
    const csvContent = UrlFetchApp.fetch(baseURL).getContentText();
    console.timeEnd("downloadTextData( csvFile:" + csvFile + " )");
    return csvContent.trim().replace(/\n$/, "");
  };

  /**
   * Internal helper: Ensures a sheet exists without destructive clearing.
   */
  const getOrCreateSheet = (activeSpreadsheet, sheetName) => {
    let sheet = activeSpreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = activeSpreadsheet.insertSheet(sheetName, activeSpreadsheet.getNumSheets());
      sheet.setName(sheetName);
    }
    return sheet;
  };



  /**
   * Prepares the sheet. Row 1 is cleared to be overwritten by sanitized headers.
   */
  const createOrClearSdeSheet = (activeSpreadsheet, sheetName) => {
    let sheet = getOrCreateSheet(activeSpreadsheet, sheetName);
    sheet.clearContents();
    _sheetCache[sheetName] = sheet;
    return sheet;
  };

  const CSVToArray = (strData, strDelimiter = ",", headers = null, publishedOnly = true) => {
    const allLines = Utilities.parseCsv(strData, strDelimiter.charCodeAt(0));
    if (allLines.length === 0) return [];

    const rawHeaders = allLines[0].map(h => h.trim());

    let colIndices = (headers && headers.length > 0)
      ? headers.map(h => rawHeaders.indexOf(h)).filter(idx => idx !== -1)
      : rawHeaders.map((_, i) => i);

    const publishIdx = rawHeaders.indexOf("published");
    const marketGroupIdx = rawHeaders.indexOf("marketGroupID");

    let arrData = [];

    for (let i = 1; i < allLines.length; i++) {
      const cols = allLines[i];

      // Gates
      if (publishedOnly === true && publishIdx !== -1) {
        const pubValue = String(cols[publishIdx]).trim();
        if (pubValue !== '1' && pubValue.toLowerCase() !== 'true') continue;
      }
      if (marketGroupIdx !== -1) {
        const mgValue = String(cols[marketGroupIdx]).trim().toLowerCase();
        if (mgValue === "" || mgValue === "null" || mgValue === "0") continue;
      }

      // --- ESCAPE LOGIC ---
      let sanitizedRow = colIndices.map(idx => {
        let val = String(cols[idx] || "").trim();

        // 1. If it's a Number -> Add ' to prevent scientific notation (e.g. '34)
        if (val !== "" && !isNaN(val)) {
          return "'" + val;
        }

        // 2. If it starts with ' -> Add ANOTHER ' so it displays correctly (e.g. ''Arbalest')
        if (val.startsWith("'")) {
          return "'" + val;
        }

        // 3. Else -> Leave it alone
        return val;
      });
      arrData.push(sanitizedRow);
    }

    // Headers: Always tick them to be safe
    const finalHeaders = (headers && headers.length > 0) ? headers : rawHeaders;
    arrData.unshift(finalHeaders.map(h => "'" + h));

    return arrData;
  };

  function _writeChunkInternal(dataChunk, startRow, numCols, sheetName) {
    const chunkStartTime = new Date().getTime();
    const docLock = LockService.getDocumentLock();
    if (!docLock.tryLock(5000)) return { success: false, duration: 0 };

    try {
      let workSheet = _sheetCache[sheetName];
      if (!workSheet) throw new Error(`Sheet '${sheetName}' not in cache.`);
      workSheet.getRange(startRow, 1, dataChunk.length, numCols).setValues(dataChunk);
    } finally {
      docLock.releaseLock();
    }
    return { success: true, duration: new Date().getTime() - chunkStartTime };
  }

  class SdePage {
    constructor(sheet, csvFile, headers = null, backupRanges = null, publishedOnly = true) {
      this.sheet = sheet;
      this.csvFile = csvFile;
      this.headers = (headers && !Array.isArray(headers)) ? [headers] : headers;
      this.backupRanges = (backupRanges && !Array.isArray(backupRanges)) ? [backupRanges] : backupRanges;
      this.publishedOnly = (publishedOnly == null) ? true : publishedOnly;
    }
  }

  const buildSDEs = (sdePage, scriptStartTime) => {
    if (sdePage == null) throw "sdePage is required";
    const activeSpreadsheet = getSS();

    const csvContent = downloadTextData(sdePage.csvFile);
    const csvData = CSVToArray(csvContent, ",", sdePage.headers, sdePage.publishedOnly);

    if (!csvData || csvData.length < 1) return true;

    const numCols = csvData[0].length;
    let currentRow = parseInt(SCRIPT_PROPS.getProperty(KEY_JOB_CHUNK_INDEX) || '0', 10);
    let finalSheetReference;

    if (currentRow === 0) {
      finalSheetReference = createOrClearSdeSheet(activeSpreadsheet, sdePage.sheet);
    } else {
      finalSheetReference = activeSpreadsheet.getSheetByName(sdePage.sheet);
      _sheetCache[sdePage.sheet] = finalSheetReference;
    }

    while (currentRow < csvData.length) {
      // 285000ms = 4.75 minutes (Safety margin for 6-minute limit)
      if ((new Date().getTime() - scriptStartTime) > 285000) {
        SCRIPT_PROPS.setProperty(KEY_JOB_CHUNK_INDEX, currentRow.toString());
        return false;
      }

      const chunkSize = 2000;
      const chunkEnd = Math.min(currentRow + chunkSize, csvData.length);
      const chunk = csvData.slice(currentRow, chunkEnd);

      let result = _writeChunkInternal(chunk, currentRow + 1, numCols, sdePage.sheet);
      if (result.success) {
        currentRow = chunkEnd;
      } else {
        Utilities.sleep(1000);
      }
    }

    // --- Final Trimming ---
    // This runs only when the loop completes successfully
    SCRIPT_PROPS.setProperty(KEY_JOB_CHUNK_INDEX, '0');

    const maxCols = finalSheetReference.getMaxColumns();
    const maxRows = finalSheetReference.getMaxRows();
    const dataRows = csvData.length;

    if (maxCols > numCols) {
      finalSheetReference.deleteColumns(numCols + 1, maxCols - numCols);
    }
    if (maxRows > dataRows) {
      finalSheetReference.deleteRows(dataRows + 1, maxRows - dataRows);
    }

    return true;
  };

  // This closes the sdeLib arrow function
  return { SdePage, buildSDEs };
};

// -----------------------------------------------------------------------------
// --- CONTROLLER FUNCTIONS ---
// -----------------------------------------------------------------------------

function sde_job_START() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const { SdePage } = sdeLib();
    let configRaw = (typeof GET_SDE_CONFIG === 'function') ? GET_SDE_CONFIG() : [];
    const sdePages = configRaw.map(item => new SdePage(item.name, item.file, item.cols));

    SCRIPT_PROPS.setProperty(KEY_JOB_RUNNING, 'true');
    SCRIPT_PROPS.setProperty(KEY_JOB_LIST, JSON.stringify(sdePages));
    SCRIPT_PROPS.setProperty(KEY_JOB_INDEX, '0');
    SCRIPT_PROPS.setProperty(KEY_JOB_CHUNK_INDEX, '0');

    _deleteTriggersFor('sde_job_PROCESS');
    ScriptApp.newTrigger('sde_job_PROCESS').timeBased().after(1000).create();
  } finally {
    lock.releaseLock();
  }
}

function sde_job_PROCESS() {
  const SCRIPT_START_TIME = new Date().getTime();
  const jobList = JSON.parse(SCRIPT_PROPS.getProperty(KEY_JOB_LIST));
  let jobIndex = parseInt(SCRIPT_PROPS.getProperty(KEY_JOB_INDEX), 10);

  if (jobIndex >= jobList.length) {
    sde_job_FINALIZE();
    return;
  }

  const SDE = sdeLib();
  const currentJob = jobList[jobIndex];
  const sdePage = new SDE.SdePage(currentJob.sheet, currentJob.csvFile, currentJob.headers, currentJob.backupRanges, currentJob.publishedOnly);

  if (SDE.buildSDEs(sdePage, SCRIPT_START_TIME)) {
    SCRIPT_PROPS.setProperty(KEY_JOB_INDEX, (jobIndex + 1).toString());
  }

  _deleteTriggersFor('sde_job_PROCESS');
  ScriptApp.newTrigger('sde_job_PROCESS').timeBased().after(2000).create();
}

// Ensure this is added to your controller logic to prevent "orphan" triggers
function sde_job_KILL_ALL_TRIGGERS() {
  _deleteTriggersFor('sde_job_PROCESS');
  SCRIPT_PROPS.deleteProperty(KEY_JOB_RUNNING);
  console.log("All SDE jobs terminated manually.");
}

function sde_job_FINALIZE() {
  SCRIPT_PROPS.deleteProperty(KEY_JOB_RUNNING);
  _deleteTriggersFor('sde_job_PROCESS');
  console.log('--- SDE JOB COMPLETE ---');
}

function _deleteTriggersFor(fn) {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); });
}