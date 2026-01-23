/**
 * Streams market data rows directly to BigQuery instead of the Sheet.
 */
function streamToBigQuery(rows) {
  const projectId = 'tenacious-tiger-345318'; // <--- CHANGE THIS
  const datasetId = 'market_data';
  const tableId = 'market_prices';

  // Map sheet-style rows to BigQuery JSON rows
  const jsonRows = rows.map(row => ({
    json: {
      date: row[0] instanceof Date ? row[0].toISOString() : new Date(row[0]).toISOString(),
      market_id: parseInt(row[1]),
      market_type: String(row[2]),
      type_id: parseInt(row[3]),
      min_sell: parseFloat(row[4]) || null,
      max_buy: parseFloat(row[5]) || null,
      median_sell: parseFloat(row[6]) || null,
      median_buy: parseFloat(row[7]) || null
    }
  }));

  const insertRequest = {
    rows: jsonRows
  };

  try {
    const response = BigQuery.Tabledata.insertAll(insertRequest, projectId, datasetId, tableId);
    
    if (response.insertErrors) {
      console.error("[BIGQUERY] Insert Errors: ", JSON.stringify(response.insertErrors));
    } else {
      console.log(`[BIGQUERY] Successfully pushed ${rows.length} rows to BigQuery.`);
    }
  } catch (err) {
    console.error("[BIGQUERY] Pipeline Error: " + err.message);
  }
}