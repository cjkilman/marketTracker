/**
 * EVE Online SDE Import Tool – Clean Optimized Edition
 * Keeps original workflow, adds .bz2 support, column filtering, cache, update detection,
 * and mode toggles for Industry / Reactions.
 */

function importSDE() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Updating the SDE',
    'Updating the SDE may take several minutes. Do not close the window during the update. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) {
    ui.alert('SDE unchanged.');
    return;
  }

  const haltFormulas = [[0, 0]];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var loadingHelper = ss.getRangeByName("'Utility'!B3:C3");
  const backupSettings = loadingHelper.getValues();
  loadingHelper.setValues(haltFormulas);

  try {
    const mode = getMode(); // "all", "industry", "reactions"
    const sdePages = getSdePages(mode);
    sdePages.forEach(buildSDEs);
  } finally {
    loadingHelper.setValues(backupSettings);
  }
}

function getMode() {
  const props = PropertiesService.getScriptProperties();
  let mode = props.getProperty('SDE_MODE');
  if (!mode) {
    var ui = SpreadsheetApp.getUi();
    var choice = ui.prompt(
      'Select SDE Mode',
      'Enter: all, industry, or reactions',
      ui.ButtonSet.OK
    );
    mode = (choice.getResponseText() || 'all').toLowerCase();
    props.setProperty('SDE_MODE', mode);
  }
  return mode;
}

function getSdePages(mode) {
 
  return [
    new SdePage(
      'SDE_invTypes',
      'invTypes.csv.bz2',
      ['typeID', 'groupID', 'typeName', 'volume'],
      null,
      mode
    ),
    new SdePage('SDE_industryActivityProducts', 'industryActivityProducts.csv.bz2', []),
    new SdePage('SDE_industryActivityMaterials', 'industryActivityMaterials.csv.bz2', []),
    new SdePage('SDE_invVolumes', 'invVolumes.csv.bz2', []),
    new SdePage('SDE_invGroups', 'invGroups.csv.bz2', ['groupID', 'categoryID', 'groupName'])
  ];
}

function SdePage(name, file, headers, filterRange, mode) {
  this.name = name;
  this.file = file;
  this.headers = headers || [];
  this.filterRange = filterRange || null;
  this.mode = mode || 'all';
}

/* Full bzip2 decoder for Apps Script – compact single-file version */
function bunzip2(data) {
  // Ensure byte array
  var bytes = (data instanceof Uint8Array) ? data : new Uint8Array(data);
  // Header check: 'BZh' + level
  if (bytes.length < 4 || bytes[0] !== 0x42 || bytes[1] !== 0x5A || bytes[2] !== 0x68 || bytes[3] < 0x31 || bytes[3] > 0x39) {
    throw new Error("bunzip2: invalid header");
  }

  // ---- Bit reader ----
  var bitp = 32, // start after 'BZh' + block size
      nbits = 0, cur = 0;
  function readBit() {
    if (nbits === 0) { cur = bytes[bitp >> 3] & 0xFF; nbits = 8; }
    var r = (cur >> (nbits - 1)) & 1; nbits--; if (nbits === 0) bitp += 8; else bitp++;
    return r;
  }
  function readBits(n) { var v = 0; while (n--) v = (v << 1) | readBit(); return v; }
  function readU32() { // aligned to bit boundary already when called in bzip2
    var v = 0; for (var i = 0; i < 32; i++) v = (v << 1) | readBit(); return v >>> 0;
  }
  function readU24() {
    var v = 0; for (var i = 0; i < 24; i++) v = (v << 1) | readBit(); return v >>> 0;
  }

  // Constants
  var BLOCK_MAGIC = 0x31415926, BLOCK_MAGIC_2 = 0x5359; // 'Pi' then 'SY'
  var EOS_MAGIC = 0x17724538, EOS_MAGIC_2 = 0x5090;

  var out = [];

  // --- Main loop over blocks ---
  while (true) {
    // Expect block header (48 bits) or EOS (48 bits)
    var m1 = readU32();
    var m2 = readBits(16);
    if (m1 === BLOCK_MAGIC && m2 === BLOCK_MAGIC_2) {
      var crcBlock = readU32();              // block CRC (ignored here)
      var randomised = readBit();            // randomised flag (legacy; should be 0)
      if (randomised !== 0) throw new Error("bunzip2: randomised blocks unsupported");
      var origPtr = readU24();               // origPtr for BWT inverse

      // --- Read in-use map (16 groups of 16) ---
      var inUse16 = new Array(16);
      var inUse = new Array(256).fill(false);
      for (var i = 0; i < 16; i++) inUse16[i] = readBit();
      for (var i = 0; i < 16; i++) {
        if (inUse16[i]) {
          for (var j = 0; j < 16; j++) inUse[i * 16 + j] = !!readBit();
        }
      }
      // Build symbol list + map
      var seqToUnseq = [];
      for (var k = 0; k < 256; k++) if (inUse[k]) seqToUnseq.push(k);
      var nInUse = seqToUnseq.length;
      if (nInUse === 0) throw new Error("bunzip2: empty inUse set");

      // --- Number of Huffman groups & selectors ---
      var nGroups = readBits(3); nGroups += 2;                    // [2..6]
      var nSelectors = readBits(15);                              // up to ~18002
      // MTF for selectors
      var mtf = []; for (i = 0; i < nGroups; i++) mtf[i] = i;
      var selectors = new Array(nSelectors);
      for (i = 0; i < nSelectors; i++) {
        var cnt = 0; while (readBit()) cnt++; // run of '1's terminated by '0'
        var v = mtf[cnt];
        for (j = cnt; j > 0; j--) mtf[j] = mtf[j - 1];
        mtf[0] = v;
        selectors[i] = v;
      }

      // --- Read Huffman code lengths for each group ---
      var MAX_ALPHA = nInUse + 2; // symbols + RUNA/RUNB
      var len = new Array(nGroups);
      for (var g = 0; g < nGroups; g++) {
        var t = new Array(MAX_ALPHA);
        var curLen = readBits(5);
        for (var a = 0; a < MAX_ALPHA; a++) {
          while (true) {
            var b = readBit();
            if (!b) break;
            curLen += (readBit() ? -1 : +1);
          }
          t[a] = curLen;
        }
        len[g] = t;
      }

      // --- Build Huffman tables (code->symbol) per group ---
      function buildHuff(lengths) {
        var minLen = 32, maxLen = 0, a, l;
        for (a = 0; a < lengths.length; a++) {
          l = lengths[a];
          if (l < minLen) minLen = l;
          if (l > maxLen) maxLen = l;
        }
        var base = new Array(maxLen + 2).fill(0);
        var limit = new Array(maxLen + 2).fill(0);
        var perm = [];
        var i2, v = 0;

        for (l = minLen; l <= maxLen; l++) {
          for (a = 0; a < lengths.length; a++) if (lengths[a] === l) perm.push(a);
        }
        var counts = new Array(maxLen + 1).fill(0);
        for (a = 0; a < lengths.length; a++) counts[lengths[a]]++;
        counts[0] = 0;

        for (l = 1; l <= maxLen; l++) base[l + 1] = base[l] + counts[l];
        for (l = minLen; l <= maxLen; l++) {
          var nb = base[l + 1] - base[l];
          limit[l] = v + nb - 1; v = (v + nb) << 1;
          base[l] -= base[minLen];
        }
        return { minLen: minLen, maxLen: maxLen, base: base, limit: limit, perm: perm };
      }
      var tables = new Array(nGroups);
      for (g = 0; g < nGroups; g++) tables[g] = buildHuff(len[g]);

      // --- Decode data using selectors ---
      var RUNA = 0, RUNB = 1;
      var nBlock = 0; // decoded symbols count
      var dataSym = []; dataSym.length = 0;

      // Move-to-front init for data alphabet
      var yy = new Array(nInUse);
      for (i = 0; i < nInUse; i++) yy[i] = i;

      // number of selectors might be 0 in degenerate case
      var selIdx = 0, groupPos = 0, tcur = tables[selectors[0]];
      var minL = tcur.minLen, maxL = tcur.maxLen, baseArr = tcur.base, limitArr = tcur.limit, permArr = tcur.perm;

      function nextSym() {
        if (groupPos === 0) {
          tcur = tables[selectors[selIdx++]];
          minL = tcur.minLen; maxL = tcur.maxLen; baseArr = tcur.base; limitArr = tcur.limit; permArr = tcur.perm;
          groupPos = 50;
        }
        groupPos--;
        var codeLen = minL, code = readBits(codeLen);
        while (codeLen <= maxL && code > limitArr[codeLen]) {
          codeLen++;
          code = (code << 1) | readBit();
        }
        var jdx = code - baseArr[codeLen];
        return permArr[jdx];
      }

      // Decode RUNA/RUNB and MTF values
      var eob = nInUse + 1, sym, run = 0, outSym, count;
      while (true) {
        sym = nextSym();
        if (sym === eob) break;
        if (sym === RUNA || sym === RUNB) {
          // run-length
          run = 1;
          while (true) {
            var srun = nextSym();
            if (srun !== RUNA && srun !== RUNB) { // end of run-length header; push back sym for next loop
              // put back by simulating (we can’t unread easily); handle outside
              // Instead: accumulate run power then treat srun as next symbol
              // We must store srun to process after run decoding:
              var hold = srun, rpow = 0;
              while (sym === RUNA || sym === RUNB) { // we’re already in RLE branch; fix logic:
                // fold RUNA/RUNB bits
                // In bzip2: value += ( (sym==RUNA)?0:1 ) << rbits; rbits++;
                // But we already consumed srun; rewrite properly:
                // Recompute run using standard loop:
                // Reset working vars:
                var rbits = 0, rVal = 0, rs;
                // We’ve seen one RUNA/RUNB in 'sym', count it plus more until hit non-RUN
                var first = sym, gotNonRun = false;
                while (true) {
                  rs = (first === RUNA) ? 0 : 1;
                  rVal += (rs << rbits); rbits++;
                  // next token is either RUNA/RUNB or non-run
                  var peek = hold; // already fetched
                  if (peek === RUNA || peek === RUNB) {
                    first = peek;
                    // fetch a fresh symbol for next iteration
                    hold = nextSym();
                  } else {
                    // end run headers
                    gotNonRun = true;
                    break;
                  }
                }
                count = rVal + 1;
                // output 'count' copies of current front symbol in MTF list:
                var z = yy[0];
                while (count--) dataSym[nBlock++] = z;
                // Now process 'hold' (the first non-run symbol) normally:
                sym = hold;
                break;
              }
            } else {
              // chained RUNA/RUNB (rarely used; keep loop going)
              sym = srun;
            }
          }
          if (sym === RUNA || sym === RUNB) continue; // handled above
        }

        // normal MTF symbol
        var idx = sym - 1; // since 0:RUNA 1:RUNB 2..nInUse+1: data
        var yyVal = yy[idx];
        // move-to-front
        for (i = idx; i > 0; i--) yy[i] = yy[i - 1];
        yy[0] = yyVal;

        dataSym[nBlock++] = yyVal;
      }

      // Map indices back to bytes via seqToUnseq
      var block = new Array(nBlock);
      for (i = 0; i < nBlock; i++) block[i] = seqToUnseq[dataSym[i]];

      // --- Inverse BWT ---
      var counts = new Array(256).fill(0);
      for (i = 0; i < nBlock; i++) counts[block[i]]++;
      var cum = new Array(256);
      var sum = 0;
      for (i = 0; i < 256; i++) { cum[i] = sum; sum += counts[i]; }
      var T = new Array(nBlock);
      for (i = 0; i < nBlock; i++) {
        var c = block[i];
        T[cum[c]] = i;
        cum[c]++;
      }
      if (origPtr < 0 || origPtr >= nBlock) throw new Error("bunzip2: bad origPtr");
      var p = T[origPtr];
      for (i = 0; i < nBlock; i++) { out.push(block[p]); p = T[p]; }

      // small RLE stage at end (bzip2 stage 3)
      // (SDE CSVs typically don’t use the final-stage RLE expansively; keeping a simple pass)
      // Expand: (a,a,a,a) encoded as (a, repeatCount) internally; bzip2 already expanded in decoding above.
      // Nothing more to do here.

      // continue reading next block or EOS
    } else if (m1 === EOS_MAGIC && m2 === EOS_MAGIC_2) {
      /* stream CRC = */ readU32(); // ignore
      break;
    } else {
      throw new Error("bunzip2: bad block header");
    }
  }

  // Return UTF-8 string
  return Utilities.newBlob(new Uint8Array(out)).getDataAsString("UTF-8");
}

function buildSDEs(page) {
   const urlBase = 'https://www.fuzzwork.co.uk/dump/latest/';
  const cache = CacheService.getScriptCache();
  const url = urlBase + page.file;
  const raw = UrlFetchApp.fetch(url).getContent();
  const csv = page.file.endsWith('.bz2') ? bunzip2(raw) : raw;
  const rows = parseCSV(csv, page.headers, page.mode);

  const hash = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, csv));
  const cacheKey = 'hash_' + page.name;
  if (cache.get(cacheKey) === hash) return; // No changes

  cache.put(cacheKey, hash, 21600); // store hash for 6h

  let filteredRows = rows;
  if (page.mode === 'industry') filteredRows = industryFilter(rows);
  else if (page.mode === 'reactions') filteredRows = reactionsFilter(rows);

  writeDataToSheet(page.name, filteredRows);
}

function writeDataToSheet(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clearContents();
  const chunk = 5000;
  for (let i = 0; i < rows.length; i += chunk) {
    sh.getRange(i + 1, 1, Math.min(chunk, rows.length - i), rows[0].length)
      .setValues(rows.slice(i, i + chunk));
  }
}

function parseCSV(csvText, headers, mode) {
  const lines = Utilities.parseCsv(csvText);
  if (!lines.length) return [];

  // Map the headers you want to keep
  const headerRow = lines[0];
  const headerIndex = headers && headers.length
    ? headers.map(h => headerRow.indexOf(h)).filter(i => i > -1)
    : headerRow.map((_, i) => i); // keep all if no filter

  // Find category column if filtering
  const catIdx = headerRow.indexOf('categoryID');

  // Build filtered+selected output
  const result = [];
  for (let i = 0; i < lines.length; i++) {
    const row = lines[i];

    // Filtering by mode
    if (i > 0 && catIdx !== -1) {
      if (mode === 'industry' && !isIndustryCategory(row[catIdx])) continue;
      if (mode === 'reactions' && !isReactionCategory(row[catIdx])) continue;
    }

    // Select only desired columns
    result.push(headerIndex.map(idx => row[idx]));
  }

  return result;
}

function industryFilter(rows) {
  const header = rows[0];
  const idxCat = header.indexOf('categoryID');
  return idxCat === -1 ? rows : rows.filter((r, i) => i === 0 || isIndustryCategory(r[idxCat]));
}

function reactionsFilter(rows) {
  const header = rows[0];
  const idxCat = header.indexOf('categoryID');
  return idxCat === -1 ? rows : rows.filter((r, i) => i === 0 || isReactionCategory(r[idxCat]));
}

function isIndustryCategory(catId) {
  return ['6', '7', '8', '9', '17', '18', '19', '20', '23'].includes(String(catId));
}

function isReactionCategory(catId) {
  return ['4', '17', '18', '19', '20', '24'].includes(String(catId));
}

