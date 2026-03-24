const WebSocket = require('ws');
const crypto = require('crypto');
const ws = new WebSocket('ws://localhost:2233/');

function send(cmd, params, id) {
  ws.send(JSON.stringify({ command: cmd, parameters: params, requestId: id }));
}

var missing = [
  { name: '5.1 Non-CR HVAC', wbs: '5' },
  { name: '5.2 CR HVAC', wbs: '5' },
  { name: '6.1 Exhaust Ducting', wbs: '6' },
  { name: '7.2 Utility_HPCDA', wbs: '7' },
  { name: '9.1 ACS', wbs: '9' },
  { name: '9.2 FAS', wbs: '9' },
  { name: '9.3 CCTV', wbs: '9' },
  { name: '9.4 PA', wbs: '9' },
  { name: '11.1 INTERNAL HYDRANT SYSTEM', wbs: '11' },
  { name: '11.2 SPRINKLER SYSTEM', wbs: '11' },
  { name: '11.3 FOAM SYSTEM', wbs: '11' },
  { name: '11.4 GASIOUS SYSTEM', wbs: '11' },
  { name: '12. Gas', wbs: '12' },
];

var idx = 0;

ws.on('open', function () {
  console.log('Inserting ' + missing.length + ' missing BQ groups...\n');
  insertNext();
});

function insertNext() {
  if (idx >= missing.length) {
    console.log('\nDone! Verifying...');
    send('quantification_exec_sql', {
      sql: "SELECT Id, Name, WBS FROM TK_ItemGroup ORDER BY Id"
    }, 'verify');
    return;
  }
  var item = missing[idx];
  var guid = crypto.randomUUID();
  var sql = "INSERT INTO TK_ItemGroup (Name, WBS, Parent, Status, CatalogId) VALUES ('" +
    item.name + "', '" + item.wbs + "', NULL, NULL, '" + guid + "')";
  console.log('  [' + (idx + 1) + '/' + missing.length + '] ' + item.name);
  send('quantification_exec_sql', { sql: sql }, 'ins_' + idx);
}

ws.on('message', function (data) {
  var r = JSON.parse(data.toString());

  if (r.RequestId && r.RequestId.startsWith('ins_')) {
    if (r.Success) {
      idx++;
      insertNext();
    } else {
      console.log('  ERROR: ' + r.Error);
      idx++;
      insertNext();
    }
  } else if (r.RequestId === 'verify') {
    if (r.Success && r.Data && r.Data.rows) {
      console.log('\n=== ALL QUANTIFICATION GROUPS ===');
      r.Data.rows.forEach(function (row) {
        console.log('  [' + row.ID + '] ' + row.Name + ' (WBS:' + row.WBS + ')');
      });
      console.log('\nTotal: ' + r.Data.rows.length + ' groups');
    }
    ws.close();
    process.exit(0);
  }
});

ws.on('error', function (e) { console.log('ERROR:', e.message); process.exit(1); });
setTimeout(function () { console.log('TIMEOUT'); process.exit(1); }, 60000);
