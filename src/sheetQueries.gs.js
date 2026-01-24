function refreshFilteredPrices() {
  const projectId = 'tenacious-tiger-345318';
  const targetSheetName = 'filtered prices';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);
  
  // 1. Get criteria from the sheet
  const marketId = sheet.getRange("C4").getValue();
  const marketType = sheet.getRange("D4").getValue();

  // --- NEW GATE: Validate ImportRange Inputs ---
  console.log(`[GATE CHECK] Market ID: "${marketId}" | Market Type: "${marketType}"`);

  // Check for empty, loading, or error states from ImportRange
  if (!marketId || isNaN(marketId) || marketId === "" || marketId === "#N/A") {
    console.warn("🛑 ABORT: Market ID is invalid or still loading from IMPORTRANGE.");
    return;
  }
  
  if (!marketType || marketType === "" || marketType === "#N/A" || marketType.toLowerCase() === "loading...") {
    console.warn("🛑 ABORT: Market Type is invalid or still loading from IMPORTRANGE.");
    return;
  }

  console.log("✅ Inputs validated. Proceeding to BigQuery.");

  // 2. Build the SQL (Using the Safety 1=1 logic we discussed)
  const sql = `
    SELECT 
      date, type_id, max_buy, min_sell 
    FROM (
      SELECT *, 
             ROW_NUMBER() OVER(PARTITION BY type_id ORDER BY date DESC) as rn
      FROM \`${projectId}.market_data.market_prices\`
      WHERE 1=1
      AND market_id = ${marketId}
      AND LOWER(market_type) = LOWER('${marketType}')
      AND date >= TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 30 DAY)
    )
    WHERE rn = 1
    ORDER BY type_id ASC
  `;

  try {
    const queryResults = BigQuery.Jobs.query({query: sql, useLegacySql: false}, projectId);
    const rows = queryResults.rows;

    if (!rows || rows.length === 0) {
      console.warn(`[BQ] No results found for Market ${marketId} (${marketType}) in the last 30 days.`);
      return;
    }

    // 3. Format results for the sheet
    const data = rows.map(row => [
      new Date(row.f[0].v), 
      row.f[1].v,           
      row.f[2].v,           
      row.f[3].v            
    ]);

    // 4. Clear and Write (Starting Row 7)
    // We write the header once, then the data below it.
    sheet.getRange("A7:D").clearContent();
    sheet.getRange(7, 1, 1, 4).setValues([["Date", "Type ID", "Max Buy", "Min Sell"]]);
    sheet.getRange(8, 1, data.length, 4).setValues(data);
    
    console.log(`[SUCCESS] Updated ${data.length} items from BigQuery.`);
  } catch (err) {
    console.error(`[BQ ERROR] SQL failed: ${err.message}`);
  }
}