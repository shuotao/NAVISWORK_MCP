/**
 * Full BOQ sync: ensure all 29 groups exist + import all items from BOQ Excel
 */
const WebSocket = require('ws');
const XLSX = require('xlsx');
const crypto = require('crypto');

const EXCEL_PATH = 'C:/Users/Admin/Downloads/Appendix A - MSI-1-MH-BQ-G-00256_Rev_1_Bill of Quantity.xlsx';
const wb = XLSX.readFile(EXCEL_PATH);

const skipSheets = ['Cover Sheet', 'Cost Summary', 'Preamble', '1. Preli. Works'];

// All 29 BOQ sheets with WBS
const allBQ = [
  {name:'2.1 Civil Works',wbs:'2'},{name:'2.2 Infra Works',wbs:'2'},{name:'2.3 PHE Works',wbs:'2'},
  {name:'3. Structure',wbs:'3'},
  {name:'4.1 Non-CR Architecture',wbs:'4'},{name:'4.2 CR Architecture',wbs:'4'},
  {name:'5.1 Non-CR HVAC',wbs:'5'},{name:'5.2 CR HVAC',wbs:'5'},
  {name:'6.1 Exhaust Ducting',wbs:'6'},{name:'6.2 Exhaust Equipment',wbs:'6'},
  {name:'7.1 Utility_CDA',wbs:'7'},{name:'7.2 Utility_HPCDA',wbs:'7'},{name:'7.3 PV',wbs:'7'},
  {name:'7.4 PCW',wbs:'7'},{name:'7.5 ICA+SW',wbs:'7'},
  {name:'8 Waste Water',wbs:'8'},
  {name:'9.1 ACS',wbs:'9'},{name:'9.2 FAS',wbs:'9'},{name:'9.3 CCTV',wbs:'9'},
  {name:'9.4 PA',wbs:'9'},{name:'9.5 FMCS',wbs:'9'},{name:'9.6 IT',wbs:'9'},
  {name:'10.1 Non-CR Elect.',wbs:'10'},{name:'10.2 CR Elect.',wbs:'10'},
  {name:'11.1 INTERNAL HYDRANT SYSTEM',wbs:'11'},{name:'11.2 SPRINKLER SYSTEM',wbs:'11'},
  {name:'11.3 FOAM SYSTEM',wbs:'11'},{name:'11.4 GASIOUS SYSTEM',wbs:'11'},
  {name:'12. Gas',wbs:'12'}
];

// Parse BOQ items from Excel
function findHeaderRow(data) {
  for (var i = 0; i < Math.min(data.length, 15); i++) {
    var row = data[i];
    if (!row) continue;
    var joined = row.map(function(c) { return String(c || '').toUpperCase().trim(); }).join(' ');
    var hasSno = joined.match(/S\.?\s*NO|SI\.?\s*NO|SR\.?\s*NO/);
    var hasDesc = joined.indexOf('DESCRIPTION') >= 0 || joined.indexOf('ITEM') >= 0;
    if (hasSno && hasDesc) return i;
  }
  return -1;
}

function findColumns(headerRow) {
  var cols = { sno: -1, desc: -1, unit: -1, qty: -1 };
  headerRow.forEach(function(cell, i) {
    var v = String(cell || '').toUpperCase().trim();
    if (v.match(/^S\.?NO|^SI\.?NO|^SR/)) cols.sno = i;
    if (v.match(/DESCRIPTION/) || v === 'ITEM') cols.desc = i;
    if (v.match(/^UNITS?$/)) cols.unit = i;
    if (v.match(/^(QUANTITY|QTY\.?|TOTAL QTY)$/)) cols.qty = i;
  });
  return cols;
}

function parseSheet(sheetName) {
  var ws = wb.Sheets[sheetName];
  if (!ws) return [];
  var data = XLSX.utils.sheet_to_json(ws, {header:1});
  var headerIdx = findHeaderRow(data);
  if (headerIdx < 0) return [];
  var cols = findColumns(data[headerIdx]);
  if (cols.desc < 0) return [];

  var items = [];
  for (var i = headerIdx + 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) continue;
    var desc = cols.desc >= 0 ? row[cols.desc] : null;
    if (!desc || String(desc).trim() === '' || String(desc).length > 300) continue;
    desc = String(desc).trim();
    if (desc.match(/^(General Note|Scope of Work|The scope|Handling|Quality Control|BOQ,|Protection|Surrounding|Barricading|Note\s*[-:])/i)) continue;

    var sno = cols.sno >= 0 && row[cols.sno] != null ? String(row[cols.sno]).trim() : '';
    var unit = cols.unit >= 0 && row[cols.unit] ? String(row[cols.unit]).trim() : '';
    var qty = cols.qty >= 0 && row[cols.qty] != null && !isNaN(Number(row[cols.qty])) ? Number(row[cols.qty]) : 0;

    var isGroup = (sno.match(/^\d+$/) && !unit);
    if (!isGroup && desc.length > 3) {
      var itemName = sno ? (sno + ' - ' + desc) : desc;
      if (itemName.length > 200) itemName = itemName.substring(0, 200);
      items.push({ name: itemName, desc: desc, unit: unit, qty: qty });
    }
  }
  return items;
}

// WebSocket execution
var ws = new WebSocket('ws://localhost:2233/');
var step = 'init';
var existingGroups = {};
var missingGroups = [];
var groupInsertIdx = 0;
var sheetIdx = 0;
var groupIdMap = {}; // name → ID

function send(sql, id) {
  ws.send(JSON.stringify({ command: 'quantification_exec_sql', parameters: { sql: sql }, requestId: id }));
}

ws.on('open', function() {
  console.log('Connected. Checking existing groups...\n');
  step = 'check_groups';
  send("SELECT ID, Name FROM TK_ItemGroup ORDER BY ID", 'cg');
});

ws.on('message', function(data) {
  var r = JSON.parse(data.toString());
  if (!r.Success) { console.log('ERROR [' + r.RequestId + ']: ' + r.Error); }

  if (step === 'check_groups') {
    if (r.Data && r.Data.rows) {
      r.Data.rows.forEach(function(row) {
        existingGroups[row.Name] = row.ID;
        groupIdMap[row.Name] = row.ID;
      });
    }
    // Find missing groups
    allBQ.forEach(function(bq) {
      if (!existingGroups[bq.name] && !existingGroups[bq.name.trim()]) {
        missingGroups.push(bq);
      }
    });
    console.log('Existing groups: ' + Object.keys(existingGroups).length);
    console.log('Missing groups: ' + missingGroups.length);

    if (missingGroups.length > 0) {
      step = 'insert_groups';
      insertNextGroup();
    } else {
      step = 'sync_items';
      sheetIdx = 0;
      syncNextSheet();
    }

  } else if (step === 'insert_groups') {
    if (r.RequestId === 'get_id') {
      if (r.Data && r.Data.rows && r.Data.rows.length > 0) {
        var lastId = r.Data.rows[0].id;
        var gname = missingGroups[groupInsertIdx - 1].name;
        groupIdMap[gname] = lastId;
      }
      insertNextGroup();
    } else {
      // After INSERT, get the ID
      send("SELECT MAX(ID) as id FROM TK_ItemGroup", 'get_id');
    }

  } else if (step === 'sync_items') {
    if (r.RequestId === 'count') {
      var bq = allBQ[sheetIdx];
      var existingCount = (r.Data && r.Data.rows) ? r.Data.rows[0].cnt : 0;
      var items = parseSheet(bq.name) || parseSheet(bq.name.trim()) || [];
      var groupId = groupIdMap[bq.name] || groupIdMap[bq.name.trim()];

      if (!groupId) {
        console.log('  [SKIP] No group ID for ' + bq.name);
        sheetIdx++;
        syncNextSheet();
        return;
      }

      if (items.length === 0) {
        console.log('  [SKIP] No parseable items in Excel');
        sheetIdx++;
        syncNextSheet();
        return;
      }

      if (existingCount >= items.length) {
        console.log('  [OK] Already has ' + existingCount + ' items (Excel: ' + items.length + ')');
        sheetIdx++;
        syncNextSheet();
        return;
      }

      // Delete existing items and re-import
      if (existingCount > 0) {
        console.log('  Replacing ' + existingCount + ' → ' + items.length + ' items');
      } else {
        console.log('  Importing ' + items.length + ' items');
      }

      // Build batch SQL
      var sqls = [];
      if (existingCount > 0) {
        sqls.push("DELETE FROM TK_Item WHERE Parent = " + groupId);
      }
      items.forEach(function(item) {
        var guid = crypto.randomUUID();
        var name = item.name.replace(/'/g, "''");
        var desc = item.desc.replace(/'/g, "''");
        var unit = item.unit.replace(/'/g, "''");
        sqls.push("INSERT INTO TK_Item (Name, Description, WBS, Parent, Status, CatalogId, Color, Transparency, LineThickness, CountSymbol, CountSize, Count_Formula, PrimaryQuantity_Formula) VALUES ('" +
          name + "', '" + desc + "', '', " + groupId + ", NULL, '" + guid + "', 0, 0.0, 1, 0, 5, '=1', '=" + item.qty + "')");
      });

      step = 'batch_insert';
      batchSqls = sqls;
      batchIdx = 0;
      batchNext();

    } else {
      sheetIdx++;
      syncNextSheet();
    }

  } else if (step === 'batch_insert') {
    batchIdx++;
    batchNext();
  }
});

function insertNextGroup() {
  if (groupInsertIdx >= missingGroups.length) {
    console.log('\nAll groups created. Syncing items...\n');
    step = 'sync_items';
    sheetIdx = 0;
    syncNextSheet();
    return;
  }
  var g = missingGroups[groupInsertIdx];
  var guid = crypto.randomUUID();
  console.log('  Creating group: ' + g.name);
  groupInsertIdx++;
  send("INSERT INTO TK_ItemGroup (Name, WBS, Parent, Status, CatalogId) VALUES ('" +
    g.name.replace(/'/g, "''") + "', '" + g.wbs + "', NULL, NULL, '" + guid + "')", 'ig');
}

function syncNextSheet() {
  if (sheetIdx >= allBQ.length) {
    finalize();
    return;
  }
  var bq = allBQ[sheetIdx];
  var groupId = groupIdMap[bq.name] || groupIdMap[bq.name.trim()];
  process.stdout.write('[' + (sheetIdx+1) + '/' + allBQ.length + '] ' + bq.name.padEnd(35));

  if (!groupId) {
    console.log(' [NO GROUP]');
    sheetIdx++;
    syncNextSheet();
    return;
  }

  step = 'sync_items';
  send("SELECT COUNT(*) as cnt FROM TK_Item WHERE Parent = " + groupId, 'count');
}

var batchSqls = [];
var batchIdx = 0;
function batchNext() {
  if (batchIdx >= batchSqls.length) {
    step = 'sync_items';
    // Don't increment sheetIdx here - the message handler will do it
    ws.send(JSON.stringify({ command: 'quantification_exec_sql', parameters: { sql: "SELECT 1" }, requestId: 'batch_done' }));
    return;
  }
  // Execute in chunks of 10 for speed
  var chunk = batchSqls.slice(batchIdx, batchIdx + 10);
  var sql = chunk.join('; ');
  batchIdx += 10;
  send(sql, 'bi');
}

function finalize() {
  console.log('\n=== SYNC COMPLETE ===');
  send("SELECT g.Name, COUNT(i.ID) as cnt FROM TK_ItemGroup g LEFT JOIN TK_Item i ON i.Parent = g.ID GROUP BY g.ID ORDER BY g.ID", 'final');
  step = 'final';
}

ws.on('message', function(data) {}); // handled above via single handler

// Override - use single message handler
var origListeners = ws.listeners('message');
if (origListeners.length > 1) {
  ws.removeListener('message', origListeners[origListeners.length - 1]);
}

// Final handler is already in the main one via step === 'final' check...
// Actually need to handle 'final' step
var _origHandler = origListeners[0];
ws.removeAllListeners('message');
ws.on('message', function(data) {
  var r = JSON.parse(data.toString());
  if (!r.Success && r.Error) {
    // Only log real errors, skip constraint warnings for batch inserts
    if (step !== 'batch_insert') console.log('ERROR: ' + r.Error);
  }

  if (step === 'final') {
    if (r.Data && r.Data.rows) {
      console.log('');
      r.Data.rows.forEach(function(row) {
        console.log('  ' + String(row.Name).padEnd(35) + ' | ' + row.cnt + ' items');
      });
      console.log('\nTotal groups: ' + r.Data.rows.length);
      var total = r.Data.rows.reduce(function(s, row) { return s + row.cnt; }, 0);
      console.log('Total items: ' + total);
    }
    ws.close();
    process.exit(0);
    return;
  }

  _origHandler(data);
});

ws.on('error', function(e) { console.log('ERROR:', e.message); process.exit(1); });
setTimeout(function() { console.log('TIMEOUT at sheet ' + sheetIdx); process.exit(1); }, 600000);
