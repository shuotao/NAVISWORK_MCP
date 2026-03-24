/**
 * Full B001 Matching: scan each Category, filter B001 by path, aggregate by Type
 * Output: { bqSheet → { typeKey → count } } for Quantification linking
 */
const WebSocket = require('ws');
const fs = require('fs');

const nwcToBQ = {
  // Architecture
  'CUB_A_ARCH': '4.1 Non-CR Architecture',
  'FAB_A_ARCH': '4.2 CR Architecture',
  // Structure
  'CUB_S_STRU': '3. Structure', 'FAB_S_STRU': '3. Structure',
  // HVAC
  'CUB_M_DUCT': '5.1 Non-CR HVAC', 'FAB_M_DUCT': '5.2 CR HVAC',
  'CUB_M_HVAC': '5.1 Non-CR HVAC', 'FAB_M_HVAC': '5.2 CR HVAC',
  // Exhaust
  'CUB_M_EGEX': '6.1 Exhaust Ducting', 'FAB_M_EGEX': '6.2 Exhaust Equipment',
  'ALL_M_EQPM': '6.2 Exhaust Equipment',
  // Process Utility
  'ALL_D_INAP': '7.1 Utility_CDA',
  'ALL_D_HPCD': '7.2 Utility_HPCDA',
  'ALL_D_PVAC': '7.3 PV',
  'CUB_M_CHWT': '7.4 PCW', 'FAB_M_CHWT': '7.4 PCW',
  'ALL_D_ICASW': '7.5 ICA+SW',
  // PHE / Civil / Infra
  'CUB_P_PLUM': '2.3 PHE Works', 'FAB_P_PLUM': '2.3 PHE Works',
  'SIT_C_CIVIL': '2.1 Civil Works',
  'SIT_I_INFRA': '2.2 Infra Works',
  // Waste Water
  'ALL_D_PRWT': '8 Waste Water',
  // ELV
  'CUB_I_INST': '9.5 FMCS', 'FAB_I_INST': '9.5 FMCS',
  'CUB_I_ACS': '9.1 ACS', 'FAB_I_ACS': '9.1 ACS',
  'CUB_I_FAS': '9.2 FAS', 'FAB_I_FAS': '9.2 FAS',
  'CUB_I_CCTV': '9.3 CCTV', 'FAB_I_CCTV': '9.3 CCTV',
  'CUB_I_PA': '9.4 PA', 'FAB_I_PA': '9.4 PA',
  'CUB_I_IT': '9.6 IT', 'FAB_I_IT': '9.6 IT',
  // Electrical
  'CUB_E_ELEC': '10.1 Non-CR Elect.', 'FAB_E_ELEC': '10.2 CR Elect.',
  // Fire Protection
  'CUB_F_FPRT': '11.1 INTERNAL HYDRANT SYSTEM', 'FAB_F_FPRT': '11.2 SPRINKLER SYSTEM',
  'ALL_F_FOAM': '11.3 FOAM SYSTEM',
  'ALL_F_GASS': '11.4 GASIOUS SYSTEM',
  // Gas
  'ALL_N_GASS': '12. Gas',
};

function getNwcKey(path) {
  if (!path) return null;
  var parts = path.split(' > ');
  // L1 必須是 B001 Federated Model（不是 combined）
  if (parts.length < 3 || parts[1] !== 'DE_MSI_001_ALL_B001_Federated Model.nwd') return null;
  // L2 是 NWC/NWD 來源
  var nwc = parts[2];
  return nwc.replace('DE_MSI_001_', '').replace('_B001_', '_').replace('.nwc', '').replace('.nwd', '');
}

// Categories to scan (skip Center line, Lines etc)
var categories = [
  'Doors', 'Walls', 'Floors', 'Windows', 'Roofs', 'Stairs', 'Curtain Panels',
  'Structural Columns', 'Structural Framing', 'Structural Foundations',
  'Ducts', 'Duct Fittings', 'Duct Accessories', 'Duct Insulations',
  'Pipes', 'Pipe Fittings', 'Pipe Accessories', 'Pipe Insulations',
  'Cable Trays', 'Cable Tray Fittings', 'Conduits', 'Conduit Fittings',
  'Electrical Equipment', 'Lighting Fixtures',
  'Mechanical Equipment',
  'Sprinklers', 'Fire Alarm Devices',
  'Communication Devices', 'Security Devices', 'Data Devices',
  'Plumbing Equipment', 'Plumbing Fixtures',
  'Air Terminals', 'Specialty Equipment',
];

var catIdx = 0;
var allResults = {}; // bqSheet → { category → { type → count } }
var ws = new WebSocket('ws://localhost:2233/');

function send(cmd, params, id) {
  ws.send(JSON.stringify({ command: cmd, parameters: params, requestId: id }));
}

var phase = '';
ws.on('open', function () {
  console.log('Connected. Scanning ' + categories.length + ' categories...\n');
  nextCat();
});

function nextCat() {
  if (catIdx >= categories.length) {
    outputResults();
    return;
  }
  phase = 'select';
  send('select_items_by_search', {
    category: 'Element', property: 'Category', value: categories[catIdx]
  }, 'sel');
}

ws.on('message', function (data) {
  var r = JSON.parse(data.toString());
  var cat = categories[catIdx];

  if (phase === 'select') {
    var total = r.Data ? r.Data.foundCount : 0;
    if (total === 0) { catIdx++; nextCat(); return; }
    phase = 'props';
    // 超大 Category（>10000）取 5000 樣本，其他全取
    var limit = total > 10000 ? 5000 : total;
    send('batch_get_properties', {
      fields: ['Element.Category', 'Element.Family', 'Element.Type', 'Element.System Name', 'Element.Size'],
      maxResults: limit
    }, 'prop');
  } else if (phase === 'props') {
    if (r.Success && r.Data && r.Data.rows) {
      var b001Count = 0;
      r.Data.rows.forEach(function (row) {
        var nwcKey = getNwcKey(row.path);
        if (!nwcKey) return;
        b001Count++;

        var bqSheet = nwcToBQ[nwcKey] || 'UNMAPPED:' + nwcKey;

        // Sub-classify I_INST by Revit Category
        if (nwcKey === 'CUB_I_INST' || nwcKey === 'FAB_I_INST') {
          if (cat === 'Fire Alarm Devices') bqSheet = '9.2 FAS';
          else if (cat === 'Security Devices') bqSheet = '9.3 CCTV';
          else if (cat === 'Communication Devices') bqSheet = '9.4 PA';
          else if (cat === 'Data Devices') bqSheet = '9.6 IT';
          else if (cat === 'Electrical Equipment') bqSheet = '9.1 ACS';
          // else stays as 9.5 FMCS
        }

        var typeKey = (row['Element.Family'] || cat) + ' | ' + (row['Element.Type'] || '');
        var sysName = row['Element.System Name'];
        if (sysName) typeKey += ' [' + sysName + ']';
        var size = row['Element.Size'];
        if (size) typeKey += ' {' + size + '}';

        if (!allResults[bqSheet]) allResults[bqSheet] = {};
        if (!allResults[bqSheet][cat]) allResults[bqSheet][cat] = {};
        allResults[bqSheet][cat][typeKey] = (allResults[bqSheet][cat][typeKey] || 0) + 1;
      });
      process.stdout.write(cat.padEnd(25) + ' | total: ' + r.Data.totalItems + ' | B001: ' + b001Count + '\n');
    } else {
      process.stdout.write(cat.padEnd(25) + ' | error: ' + (r.Error || '').substring(0, 50) + '\n');
    }
    catIdx++;
    nextCat();
  }
});

function outputResults() {
  console.log('\n\n========================================');
  console.log('B001 MODEL ↔ BQ MATCHING RESULTS');
  console.log('========================================\n');

  var report = [];
  Object.keys(allResults).sort().forEach(function (sheet) {
    report.push('\n=== ' + sheet + ' ===');
    var cats = allResults[sheet];
    Object.keys(cats).forEach(function (cat) {
      var types = cats[cat];
      Object.keys(types).sort().forEach(function (typeKey) {
        report.push('  ' + cat.padEnd(22) + ' | ' + typeKey + ' × ' + types[typeKey]);
      });
    });
  });

  var output = report.join('\n');
  console.log(output);

  fs.writeFileSync(__dirname + '/b001_full_match.txt', output);
  fs.writeFileSync(__dirname + '/b001_full_match.json', JSON.stringify(allResults, null, 2));
  console.log('\n\nSaved to b001_full_match.txt / .json');

  ws.close();
  process.exit(0);
}

ws.on('error', function (e) { console.log('ERROR:', e.message); process.exit(1); });
setTimeout(function () { console.log('TIMEOUT at cat ' + catIdx + ': ' + categories[catIdx]); process.exit(1); }, 600000);
