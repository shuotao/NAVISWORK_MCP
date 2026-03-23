/**
 * Match B001 model elements → BQ Items, write to Quantification TK_Object
 * Only processes items under DE_MSI_001_ALL_B001_Federated Model.nwd
 */

const WebSocket = require('ws');
const fs = require('fs');

const bqItems = JSON.parse(fs.readFileSync(__dirname + '/bq_items.json', 'utf8'));

// NWC short name → BQ Sheet
const nwcToBQ = {
  'CUB_A_ARCH': '4.1 Non-CR Architecture',
  'FAB_A_ARCH': '4.2 CR Architecture',
  'CUB_S_STRU': '3. Structure',
  'FAB_S_STRU': '3. Structure',
  'CUB_M_CHWT': '7.4 PCW',
  'FAB_M_CHWT': '7.4 PCW',
  'CUB_M_DUCT': '5.1 Non-CR HVAC',
  'FAB_M_DUCT': '5.2 CR HVAC',
  'CUB_P_PLUM': '2.3 PHE Works',
  'FAB_P_PLUM': '2.3 PHE Works',
  'CUB_F_FPRT': '11.1 INTERNAL HYDRANT SYSTEM',
  'FAB_F_FPRT': '11.2 SPRINKLER SYSTEM',
  'CUB_E_ELEC': '10.1 Non-CR Elect.',
  'FAB_E_ELEC': '10.2 CR Elect.',
  'CUB_I_INST': '9.5 FMCS',
  'FAB_I_INST': '9.5 FMCS',
  'SIT_C_CIVIL': '2.1 Civil Works',
  'ALL_M_EQPM': '6.2 Exhaust Equipment',
  'ALL_D_PRWT': '8 Waste Water',
  'ALL_N_GASS': '12. Gas',
  'CUB_M_EGEX': '6.1 Exhaust Ducting',
  'FAB_M_EGEX': '6.2 Exhaust Equipment',
  'ALL_D_INAP': '7.1 Utility_CDA',
};

// 從 path 解析 NWC 來源
function getNwcFromPath(path) {
  if (!path || path.indexOf('B001_Federated Model.nwd') < 0) return null;
  var parts = path.split(' > ');
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i];
    if (p.indexOf('B001_') >= 0 && (p.endsWith('.nwc') || p.endsWith('.nwd')) && p !== 'DE_MSI_001_ALL_B001_Federated Model.nwd') {
      return p.replace('DE_MSI_001_', '').replace('_B001_', '_').replace('.nwc', '').replace('.nwd', '');
    }
  }
  return null;
}

// 取得 BQ Group ID (已在 DB 中)
var bqGroupIds = {}; // sheet name → TK_ItemGroup.ID
var bqItemIds = {};  // sheet name → [TK_Item.ID, ...]

var ws;
var phase = 'init';
var scanCategories = [
  'Ducts', 'Duct Fittings', 'Duct Accessories',
  'Pipes', 'Pipe Fittings', 'Pipe Accessories',
  'Walls', 'Doors', 'Floors',
  'Structural Columns', 'Structural Framing',
  'Cable Trays', 'Electrical Equipment', 'Lighting Fixtures',
  'Mechanical Equipment', 'Sprinklers', 'Fire Alarm Devices',
  'Communication Devices', 'Security Devices', 'Data Devices',
  'Plumbing Equipment', 'Plumbing Fixtures',
  'Conduits',
];
var catIdx = 0;
var matchResults = {}; // bqSheet → { count, guids[] }

function send(cmd, params, id) {
  ws.send(JSON.stringify({ command: cmd, parameters: params, requestId: id }));
}

function connect() {
  ws = new WebSocket('ws://localhost:2233/');

  ws.on('open', function () {
    console.log('Connected.\n');
    // Step 1: 讀取 Quantification 中的 Group IDs
    console.log('Step 1: Reading Quantification groups...');
    send('quantification_exec_sql', {
      sql: 'SELECT ID, Name FROM TK_ItemGroup ORDER BY ID'
    }, 'getGroups');
  });

  ws.on('message', function (data) {
    var r = JSON.parse(data.toString());

    if (r.RequestId === 'getGroups') {
      if (r.Success && r.Data.rows) {
        r.Data.rows.forEach(function (row) { bqGroupIds[row.Name] = row.ID; });
        console.log('  Groups: ' + Object.keys(bqGroupIds).length);
      }
      // 讀取 Items
      send('quantification_exec_sql', {
        sql: 'SELECT ID, Name, Parent FROM TK_Item ORDER BY ID'
      }, 'getItems');
      return;
    }

    if (r.RequestId === 'getItems') {
      if (r.Success && r.Data.rows) {
        r.Data.rows.forEach(function (row) {
          // 找出這個 item 屬於哪個 group
          var groupId = row.Parent;
          for (var sheet in bqGroupIds) {
            if (bqGroupIds[sheet] === groupId) {
              if (!bqItemIds[sheet]) bqItemIds[sheet] = [];
              bqItemIds[sheet].push(row.ID);
              break;
            }
          }
        });
        console.log('  Items loaded');
      }
      // Step 2: 掃描模型
      console.log('\nStep 2: Scanning B001 elements by category...\n');
      scanNext();
      return;
    }

    if (phase === 'scan_select') {
      var count = r.Data ? r.Data.foundCount : 0;
      if (count === 0) {
        catIdx++;
        scanNext();
        return;
      }
      phase = 'scan_props';
      send('batch_get_properties', {
        fields: ['Element.Category', 'Element.Family', 'Element.Type',
          'Element.System Name', 'Element.Size'],
        maxResults: 50
      }, 'prop_' + catIdx);
      return;
    }

    if (phase === 'scan_props') {
      var cat = scanCategories[catIdx];
      if (r.Success && r.Data && r.Data.rows) {
        var b001Count = 0;
        r.Data.rows.forEach(function (row) {
          var nwc = getNwcFromPath(row.path);
          if (!nwc) return; // 不在 B001 下
          b001Count++;

          var bqSheet = nwcToBQ[nwc];
          if (!bqSheet) return;

          if (!matchResults[bqSheet]) matchResults[bqSheet] = { sampleCount: 0, samples: [] };
          matchResults[bqSheet].sampleCount++;
          if (matchResults[bqSheet].samples.length < 5) {
            matchResults[bqSheet].samples.push({
              category: cat,
              family: row['Element.Family'],
              type: row['Element.Type'],
              systemName: row['Element.System Name'],
              size: row['Element.Size'],
              nwc: nwc
            });
          }
        });
        console.log('  ' + cat + ': ' + b001Count + ' in B001 (of ' + r.Data.totalItems + ' total)');
      }
      catIdx++;
      scanNext();
      return;
    }

    if (r.RequestId === 'summary') {
      console.log('\n========================================');
      console.log('MATCHING SUMMARY (B001 only)');
      console.log('========================================\n');

      Object.keys(matchResults).sort().forEach(function (sheet) {
        var m = matchResults[sheet];
        var bqCount = bqItemIds[sheet] ? bqItemIds[sheet].length : 0;
        console.log(sheet.padEnd(30) + ' | BQ items: ' + bqCount + ' | Model samples: ' + m.sampleCount);
        m.samples.forEach(function (s) {
          console.log('  → ' + s.category + ' | ' + (s.family || '') + ' | ' + (s.systemName || '') + ' | ' + (s.size || ''));
        });
        console.log('');
      });

      fs.writeFileSync(__dirname + '/b001_matching_results.json', JSON.stringify(matchResults, null, 2));
      console.log('Saved to b001_matching_results.json');
      ws.close();
      process.exit(0);
    }
  });

  ws.on('error', function (e) { console.log('ERROR:', e.message); process.exit(1); });
}

function scanNext() {
  if (catIdx >= scanCategories.length) {
    send('quantification_exec_sql', { sql: 'SELECT 1' }, 'summary');
    return;
  }
  phase = 'scan_select';
  var cat = scanCategories[catIdx];
  send('select_items_by_search', {
    category: 'Element', property: 'Category', value: cat
  }, 'sel_' + catIdx);
}

connect();
setTimeout(function () { console.log('TIMEOUT'); process.exit(1); }, 600000);
