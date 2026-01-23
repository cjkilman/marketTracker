/**
 * Replaces the LET formula by pulling the latest prices from BigQuery.
 * Trigger this manually or on a timer.
 */
function refreshFilteredPrices() {
  const projectId = 'tenacious-tiger-345318';
  const targetSheetName = 'filtered prices'; // Match your sheet name
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);
  
  // 1. Get criteria from the sheet (matching your old C4 and D4 cells)
  const marketId = sheet.getRange("C4").getValue();
  const marketType = sheet.getRange("D4").getValue();

  // 2. SQL Query to get the LATEST entry for every type_id in that market
  const sql = `
    SELECT 
      date, type_id, max_buy, min_sell 
    FROM (
      SELECT *, 
             ROW_NUMBER() OVER(PARTITION BY type_id ORDER BY date DESC) as rn
      FROM \`${projectId}.market_data.market_prices\`
      WHERE market_id = ${marketId} 
      AND lower(market_type) = lower('${marketType}')
    )
    WHERE rn = 1
    ORDER BY type_id ASC
  `;

  const queryResults = BigQuery.Jobs.query({query: sql, useLegacySql: false}, projectId);
  const rows = queryResults.rows;

  if (!rows || rows.length === 0) {
    console.warn("No data found in BigQuery for these criteria.");
    return;
  }

  // 3. Format results for the sheet
  const headers = [["Date", "Type ID", "Max Buy", "Min Sell"]];
  const data = rows.map(row => [
    new Date(row.f[0].v), // Date
    row.f[1].v,           // Type ID
    row.f[2].v,           // Max Buy
    row.f[3].v            // Min Sell
  ]);

  // 4. Clear old data and write new data (starting at Row 7 to avoid overwriting your config)
  sheet.getRange("A7:D").clearContent();
  sheet.getRange(7, 1, data.length, 4).setValues(data);
  
  console.log(`Updated ${data.length} items from BigQuery.`);
}