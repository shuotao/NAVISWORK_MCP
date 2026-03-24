/**
 * BQ Sync: compare BOQ Excel with current Quantification, import missing sheets
 */
const WebSocket = require('ws');
const XLSX = require('xlsx');
const fs = require('fs');

const EXCEL_PATH = 'C:/Users/Admin/Downloads/Appendix A - MSI-1-MH-BQ-G-00256_Rev_1_Bill of Quantity.xlsx';
const ws = new WebSocket('ws://localhost:2233/');

function send(cmd, params, id) {
  ws.send(JSON.stringify({ command: cmd, parameters: params, requestId: id }));
}

// BOQ sheets that should be in Quantification (skip Cover, Cost Summary, Preamble, Preli)
const bqSheets = [
  '2.1 Civil Works', '2.2 Infra Works', '2.3 PHE Works',
  '3. Structure',
  '4.1 Non-CR Architecture', '4.2 CR Architecture',
  '5.1 Non-CR HVAC', '5.2 CR HVAC',
  '6.1 Exhaust Ducting', '6.2 Exhaust Equipment',
  '7.1 Utility_CDA', '7.2 Utility_HPCDA', '7.3 PV', '7.4 PCW', '7.5 ICA+SW',
  '8 Waste Water',
  '9.1 ACS', '9.2 FAS', '9.3 CCTV', '9.4 PA', '9.5 FMCS', '9.6 IT',
  '10.1 Non-CR Elect.', '10.2 CR Elect.',
  '11.1 INTERNAL HYDRANT SYSTEM', '11.2 SPRINKLER SYSTEM',
  '11.3 FOAM SYSTEM ', '11.4 GASIOUS SYSTEM',
  '12. Gas'
];

var step = '';

ws.on('open', function () {
  console.log('Connected. Querying current Quantification groups...\n');
  step = 'query';
  send('quantification_get_item_groups', {}, 'qg');
});

ws.on('message', function (data) {
  var r = JSON.parse(data.toString());

  if (step === 'query') {
    if (!r.Success) {
      console.log('Error querying Quantification:', r.Error);
      ws.close(); process.exit(1); return;
    }

    var existing = {};
    if (r.Data && r.Data.groups) {
      r.Data.groups.forEach(function (g) {
        existing[g.Name] = g;
      });
    }

    console.log('=== CURRENT QUANTIFICATION GROUPS ===');
    Object.keys(existing).forEach(function (k) {
      console.log('  [EXISTS] ' + k + ' (WBS:' + existing[k].WBS + ')');
    });

    // Find missing
    var missing = bqSheets.filter(function (s) {
      return !existing[s] && !existing[s.trim()];
    });

    console.log('\n=== MISSING FROM QUANTIFICATION ===');
    missing.forEach(function (s) { console.log('  [MISSING] ' + s); });
    console.log('\nTotal existing: ' + Object.keys(existing).length);
    console.log('Total missing: ' + missing.length);

    if (missing.length === 0) {
      console.log('\nAll BOQ sheets already exist!');
      ws.close(); process.exit(0); return;
    }

    // Build import items for missing sheets (just the group headers)
    var importItems = missing.map(function (sheet) {
      var wbs = sheet.split(' ')[0].replace('.', '');
      return { sheet: sheet, sno: sheet.split(' ')[0], desc: sheet, unit: 'LS', qty: 0 };
    });

    console.log('\nImporting ' + importItems.length + ' missing groups...');
    step = 'import';
    send('quantification_import_bq', { items: importItems }, 'imp');

  } else if (step === 'import') {
    if (r.Success) {
      console.log('\n=== IMPORT RESULT ===');
      console.log(JSON.stringify(r.Data, null, 2));
    } else {
      console.log('Import failed:', r.Error);
    }
    ws.close();
    process.exit(0);
  }
});

ws.on('error', function (e) { console.log('ERROR:', e.message); process.exit(1); });
setTimeout(function () { console.log('TIMEOUT'); process.exit(1); }, 60000);
