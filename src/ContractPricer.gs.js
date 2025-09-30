// ContractPricer.gs.js — evaluates a contract against hub reference prices
var ContractPricer = (function(){
const CONTRACT_ITEMS_SHEET = 'contracts_items'; // raw ESI: contract_id | type_id | quantity | is_included
const TRACKER_SHEET = 'Minerals Tracker';
const OUT_SHEET = 'Contract Pricing';


function price(contractId, hubRef, mode){
const ss=SpreadsheetApp.getActive();
const ci = ss.getSheetByName(CONTRACT_ITEMS_SHEET);
if(!ci) throw new Error('Missing '+CONTRACT_ITEMS_SHEET);
const cv = ci.getDataRange().getValues(); const ch = cv.shift().map(String);
const cID=ch.indexOf('contract_id'), cT=ch.indexOf('type_id'), cQ=ch.indexOf('quantity'), cInc=ch.indexOf('is_included');


const items = cv.filter(r=>String(r[cID])==String(contractId) && (r[cInc]===true || r[cInc]==='TRUE'))
.map(r=>({type_id:r[cT], qty:Number(r[cQ])||0}));


const tr=ss.getSheetByName(TRACKER_SHEET); if(!tr) throw new Error('Run MineralTracker.build() first');
const tv=tr.getDataRange().getValues(); const th=tv.shift().map(String);
const tType=th.indexOf('type_id'), tHub=th.indexOf('Hub');
const tSellMin=th.indexOf('sell_min (now)'), tMedSell=th.indexOf('median_sell_24h');


const rows=[]; let total=0;
for (const it of items){
// pick ref
const refRow = tv.find(r=>r[tType]==it.type_id && r[tHub]==hubRef);
if(!refRow){ rows.push(['', it.type_id, '', it.qty, null, null, 'no ref']); continue; }
let refPrice = null;
if (mode==='median_sell') refPrice = refRow[tMedSell];
else /* sell_min */ refPrice = refRow[tSellMin];


const line = refPrice!=null ? refPrice * it.qty : null;
if(line!=null) total += line;
rows.push(['', it.type_id, '', it.qty, refPrice, line, '']);
}


// write out
const out = ss.getSheetByName(OUT_SHEET) || ss.insertSheet(OUT_SHEET);
out.getRange('B1').setValue(contractId);
out.getRange('B2').setValue(hubRef);
out.getRange('B3').setValue(mode||'sell_min');
out.getRange('B4').setValue(0.08);


const hdr=['line','type_id','Item Name','Qty','Ref price','Line value','Notes'];
const data=[hdr].concat(rows);
out.getRange(6,1,data.length,data[0].length).setValues(data);


// Totals & markdown
const last = 6 + rows.length;
out.getRange(last+1,4,1,2).setValues([["Subtotal:", total]]);
out.getRange(last+2,4,1,2).setValues([["Fair offer (1 - markdown_pct):", total * (1 - Number(out.getRange('B4').getValue()||0))]]);
}


return { price: price };
})();