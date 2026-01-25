/**
 * Streams market data to BigQuery using free "Load Jobs".
 * REVISED: Handles Object inputs from FuzzWorker correctly.
 */
function streamToBigQuery(rows) {
  if (!rows || rows.length === 0) return;

  const projectId = 'tenacious-tiger-345318'; 
  const datasetId = 'market_data';
  const tableId = 'market_prices';

  // 1. Transform rows into Newline-Delimited JSON
  const jsonRows = rows.map(row => {
    // Determine if input is Object (Worker) or Array (Legacy)
    const isObject = !Array.isArray(row);

    // Extract values safely based on input type
    const rawDate    = isObject ? row.date        : row[0];
    const marketId   = isObject ? row.market_id   : row[1];
    const marketType = isObject ? row.market_type : row[2];
    const typeId     = isObject ? row.type_id     : row[3];
    const minSell    = isObject ? row.min_sell    : row[4];
    const maxBuy     = isObject ? row.max_buy     : row[5];
    const medSell    = isObject ? row.median_sell : row[6];
    const medBuy     = isObject ? row.median_buy  : row[7];

    // Validate Date
    let validDate;
    if (rawDate instanceof Date) {
      validDate = !isNaN(rawDate) ? rawDate.toISOString() : new Date().toISOString();
    } else if (typeof rawDate === 'string') {
      // If it's already an ISO string (which our worker sends), use it.
      // If empty/invalid, fallback to NOW.
      validDate = rawDate || new Date().toISOString();
    } else {
      validDate = new Date().toISOString();
    }

    return JSON.stringify({
      date: validDate,
      market_id: parseInt(marketId),
      market_type: String(marketType),
      type_id: parseInt(typeId),
      min_sell: parseFloat(minSell) || null,
      max_buy: parseFloat(maxBuy) || null,
      median_sell: parseFloat(medSell) || null,
      median_buy: parseFloat(medBuy) || null
    });
  }).join('\n');

  // Convert to a blob for the upload job
  const blob = Utilities.newBlob(jsonRows, 'application/octet-stream');

  // 2. Configure the Load Job
  const jobConfig = {
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

  try {
    const runJob = BigQuery.Jobs.insert(jobConfig, projectId, blob);
    console.log(`[BIGQUERY] Load Job started: ${runJob.jobReference.jobId}. Rows: ${rows.length}`);
  } catch (err) {
    console.error("[BIGQUERY] Pipe Error: " + err.message);
  }
}