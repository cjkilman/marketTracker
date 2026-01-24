/**
 * Streams market data to BigQuery using free "Load Jobs".
 * This method is completely free and bypasses the Sandbox streaming restriction.
 */
function streamToBigQuery(rows) {
  const projectId = 'tenacious-tiger-345318'; 
  const datasetId = 'market_data';
  const tableId = 'market_prices';

  // 1. Transform rows into Newline-Delimited JSON (Free Tier requirement)
  const jsonRows = rows.map(row => {
    return JSON.stringify({
      date: row[0] instanceof Date ? row[0].toISOString() : new Date(row[0]).toISOString(),
      market_id: parseInt(row[1]),
      market_type: String(row[2]),
      type_id: parseInt(row[3]),
      min_sell: parseFloat(row[4]) || null,
      max_buy: parseFloat(row[5]) || null,
      median_sell: parseFloat(row[6]) || null,
      median_buy: parseFloat(row[7]) || null
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
        writeDisposition: 'WRITE_APPEND' // Appends to existing history
      }
    }
  };

  try {
    // 3. Execute as a Job (Free) instead of a Stream (Paid)
    const runJob = BigQuery.Jobs.insert(jobConfig, projectId, blob);
    console.log(`[BIGQUERY] Load Job started: ${runJob.jobReference.jobId}. Rows: ${rows.length}`);
  } catch (err) {
    console.error("[BIGQUERY] Free-Tier Pipe Error: " + err.message);
  }
}