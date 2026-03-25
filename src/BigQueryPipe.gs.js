function streamToBigQuery(rows) {
  if (!rows || rows.length === 0) return;

  const projectId = 'tenacious-tiger-345318';
  const datasetId = 'market_data';
  const tableId = 'market_prices';

  // 1. UNIVERSAL MAPPER: Handles both Objects (from fetcher) and Arrays (from sheet)
  const jsonRows = rows.map(row => {
    const isArr = Array.isArray(row);
    
    // Safely extract date
    let rawDate = isArr ? row[0] : row.date;
    let isoDate = (rawDate instanceof Date) ? rawDate.toISOString() : new Date(rawDate || Date.now()).toISOString();

    // Map everything securely
    const obj = {
      date: isoDate,
      market_id: parseInt(isArr ? row[1] : row.market_id) || 0,
      market_type: String(isArr ? row[2] : row.market_type || 'system'),
      type_id: parseInt(isArr ? row[3] : row.type_id) || 0,
      min_sell: parseFloat(isArr ? row[4] : row.min_sell) || 0,
      max_buy: parseFloat(isArr ? row[5] : row.max_buy) || 0,
      median_sell: parseFloat(isArr ? row[6] : row.median_sell) || 0,
      median_buy: parseFloat(isArr ? row[7] : row.median_buy) || 0
    };
    return JSON.stringify(obj);
  }).join('\n');

  // 2. Create the Blob
  const blob = Utilities.newBlob(jsonRows, 'application/octet-stream');

  // 3. Configure the Job
  const job = {
    configuration: {
      load: {
        destinationTable: { projectId, datasetId, tableId },
        sourceFormat: 'NEWLINE_DELIMITED_JSON', // Back to the safest standard
        writeDisposition: 'WRITE_APPEND',
        schema: {
          fields: [
            { name: 'date', type: 'TIMESTAMP', mode: 'REQUIRED' },
            { name: 'market_id', type: 'INTEGER' },
            { name: 'market_type', type: 'STRING' },
            { name: 'type_id', type: 'INTEGER' },
            { name: 'min_sell', type: 'FLOAT' },
            { name: 'max_buy', type: 'FLOAT' },
            { name: 'median_sell', type: 'FLOAT' },
            { name: 'median_buy', type: 'FLOAT' }
          ]
        }
      }
    }
  };

  try {
    const result = BigQuery.Jobs.insert(job, projectId, blob);
    console.log(`[BQ] Upload Success! Job ID: ${result.jobReference.jobId}`);
  } catch (err) {
    console.error(`[BQ] Upload Failed: ${err.message}`);
  }
}