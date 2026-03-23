/**
 * BQ ↔ Model Element Matching Engine
 *
 * Strategy:
 * 1. NWC source → BQ Sheet (discipline filter)
 * 2. System Name/Type → specific BQ group (system filter)
 * 3. Category + Type + Size → BQ item (element filter)
 */

const WebSocket = require('ws');
const fs = require('fs');
const path = require('path');

const bqItems = JSON.parse(fs.readFileSync(path.join(__dirname, 'bq_items.json'), 'utf8'));
const nwcMapping = JSON.parse(fs.readFileSync(path.join(__dirname, 'nwc_bq_mapping.json'), 'utf8'));

// ─── System Name → BQ Sheet 對應 ───
const systemToBQ = {
  // HVAC
  'Exhaust Air': '6.1 Exhaust Ducting',
  'Supply Air': '5.1 Non-CR HVAC',
  'Return Air': '5.1 Non-CR HVAC',
  'Fresh Air': '5.1 Non-CR HVAC',
  // Piping Systems
  'PROCESS WATER': '7.4 PCW',
  'CHILLED WATER': '7.4 PCW',
  'PCW': '7.4 PCW',
  'PWL': '7.4 PCW',
  'CDA': '7.1 Utility_CDA',
  'HPCDA': '7.1 Utility_CDA',
  'PV': '7.3 PV',
  'PROCESS VACUUM': '7.3 PV',
  'ICA': '7.5 ICA+SW',
  'SCRUBBER WATER': '7.5 ICA+SW',
  'DIW': '8 Waste Water',
  'UPW': '8 Waste Water',
  'WASTE': '8 Waste Water',
  'DRAIN': '8 Waste Water',
  'MMW': '8 Waste Water',
  'IPAL': '8 Waste Water',
  'APM': '8 Waste Water',
  'CCW': '8 Waste Water',
  'IW': '8 Waste Water',
  'WW': '8 Waste Water',
  'WWR': '8 Waste Water',
  'WWS': '8 Waste Water',
  'SWPD': '8 Waste Water',
  'TMAH': '8 Waste Water',
  'NG': '12. Gas',
  'GN2': '12. Gas',
  'PN2': '12. Gas',
  'PO2': '12. Gas',
  'PCO2': '12. Gas',
  'H2': '12. Gas',
  'PAR': '12. Gas',
  'Fire Protection': '11.2 SPRINKLER SYSTEM',
  'Sprinkler': '11.2 SPRINKLER SYSTEM',
  'FP': '11.1 INTERNAL HYDRANT SYSTEM',
};

// ─── Category → BQ Sheet 對應 ───
const categoryToBQ = {
  'Walls': '4.1 Non-CR Architecture',
  'Doors': '4.1 Non-CR Architecture',
  'Windows': '4.1 Non-CR Architecture',
  'Floors': '4.1 Non-CR Architecture',
  'Roofs': '4.1 Non-CR Architecture',
  'Curtain Panels': '4.1 Non-CR Architecture',
  'Curtain Wall Mullions': '4.1 Non-CR Architecture',
  'Stairs': '4.1 Non-CR Architecture',
  'Railings': '4.1 Non-CR Architecture',
  'Structural Columns': '3. Structure',
  'Structural Framing': '3. Structure',
  'Structural Foundations': '3. Structure',
  'Structural Connections': '3. Structure',
  'Ducts': '5.1 Non-CR HVAC',
  'Duct Fittings': '5.1 Non-CR HVAC',
  'Duct Accessories': '5.1 Non-CR HVAC',
  'Duct Insulations': '5.1 Non-CR HVAC',
  'Pipes': '2.3 PHE Works',
  'Pipe Fittings': '2.3 PHE Works',
  'Pipe Accessories': '2.3 PHE Works',
  'Pipe Insulations': '2.3 PHE Works',
  'Sprinklers': '11.2 SPRINKLER SYSTEM',
  'Fire Alarm Devices': '9.2 FAS',
  'Security Devices': '9.1 ACS',
  'Communication Devices': '9.1 ACS',
  'Data Devices': '9.6 IT',
  'Cable Trays': '10.1 Non-CR Elect.',
  'Cable Tray Fittings': '10.1 Non-CR Elect.',
  'Conduits': '10.1 Non-CR Elect.',
  'Conduit Fittings': '10.1 Non-CR Elect.',
  'Electrical Equipment': '10.1 Non-CR Elect.',
  'Electrical Fixtures': '10.1 Non-CR Elect.',
  'Lighting Fixtures': '10.1 Non-CR Elect.',
  'Lighting Devices': '10.1 Non-CR Elect.',
  'Mechanical Equipment': '6.2 Exhaust Equipment',
  'Plumbing Equipment': '2.3 PHE Works',
  'Plumbing Fixtures': '2.3 PHE Works',
  'Air Terminals': '5.1 Non-CR HVAC',
  'Generic Models': null, // 需要根據 NWC 來源判斷
  'Specialty Equipment': null,
};

// ─── Size 解析：從 "150 mmø" 或 "300 mmx100 mm" 提取 NB ───
function parseSize(sizeStr) {
  if (!sizeStr) return null;
  var m = sizeStr.match(/(\d+)\s*mm/);
  return m ? parseInt(m[1]) : null;
}

// ─── 從 System Name/Type 判斷所屬 BQ Sheet ───
function matchSystemToBQ(systemName, systemType) {
  var combined = ((systemName || '') + ' ' + (systemType || '')).toUpperCase();
  for (var key in systemToBQ) {
    if (combined.indexOf(key.toUpperCase()) >= 0) {
      return systemToBQ[key];
    }
  }
  return null;
}

// ─── 從 NWC 路徑判斷是 CUB (Non-CR) 還是 FAB (CR) ───
function isCR(ancestorPath) {
  return ancestorPath && ancestorPath.indexOf('FAB') >= 0;
}

// ─── 主匹配邏輯 ───
function matchElement(element) {
  var cat = element['Element.Category'];
  var systemName = element['Element.System Name'];
  var systemType = element['Element.System Type'];
  var size = parseSize(element['Element.Size']);
  var family = element['Element.Family'];
  var type = element['Element.Type'];
  var ancestorPath = element.path || '';
  var cr = isCR(ancestorPath);

  // 1. 先用 System Name/Type 做精準匹配
  var bqSheet = matchSystemToBQ(systemName, systemType);

  // 2. 如果沒有 System，用 Category 匹配
  if (!bqSheet && cat) {
    bqSheet = categoryToBQ[cat];
  }

  // 3. CR/Non-CR 修正
  if (bqSheet && cr) {
    if (bqSheet === '4.1 Non-CR Architecture') bqSheet = '4.2 CR Architecture';
    if (bqSheet === '5.1 Non-CR HVAC') bqSheet = '5.2 CR HVAC';
    if (bqSheet === '10.1 Non-CR Elect.') bqSheet = '10.2 CR Elect.';
  }

  return {
    bqSheet: bqSheet,
    matchedBy: bqSheet ? (systemName ? 'system' : 'category') : 'unmatched',
    category: cat,
    systemName: systemName,
    size: size,
    family: family,
    type: type,
    cr: cr
  };
}

// ─── WebSocket 連線並執行匹配 ───
var ws = new WebSocket('ws://localhost:2233/');
var results = {};
var categorySamples = [
  'Ducts', 'Duct Fittings', 'Duct Accessories', 'Duct Insulations',
  'Pipes', 'Pipe Fittings', 'Pipe Accessories', 'Pipe Insulations',
  'Walls', 'Doors', 'Windows', 'Floors',
  'Structural Columns', 'Structural Framing', 'Structural Foundations',
  'Cable Trays', 'Cable Tray Fittings',
  'Electrical Equipment', 'Lighting Fixtures',
  'Mechanical Equipment', 'Sprinklers', 'Fire Alarm Devices',
  'Communication Devices', 'Security Devices', 'Data Devices',
  'Plumbing Equipment', 'Plumbing Fixtures',
  'Air Terminals', 'Conduits', 'Conduit Fittings',
];

var catIdx = 0;
var allMatched = [];
var phase = 'select';

function send(cmd, params, id) {
  ws.send(JSON.stringify({ command: cmd, parameters: params, requestId: id }));
}

function processNextCategory() {
  if (catIdx >= categorySamples.length) {
    // 完成 — 輸出結果
    outputResults();
    return;
  }
  phase = 'select';
  var cat = categorySamples[catIdx];
  console.log('Scanning: ' + cat + ' (' + (catIdx + 1) + '/' + categorySamples.length + ')...');
  send('select_items_by_search', {
    category: 'Element', property: 'Category', value: cat
  }, 'sel_' + catIdx);
}

ws.on('open', function () {
  console.log('Connected. Starting matching engine...\n');
  processNextCategory();
});

ws.on('message', function (data) {
  var r = JSON.parse(data.toString());
  var cat = categorySamples[catIdx];

  if (phase === 'select') {
    var count = r.Data ? r.Data.foundCount : 0;
    results[cat] = { total: count, matched: {} };
    if (count === 0) {
      catIdx++;
      processNextCategory();
      return;
    }
    phase = 'props';
    // 取樣最多 20 個，讀關鍵屬性
    send('batch_get_properties', {
      fields: ['Element.Category', 'Element.Family', 'Element.Type',
        'Element.System Name', 'Element.System Type', 'Element.Size', 'Element.Mark'],
      maxResults: 20
    }, 'prop_' + catIdx);
  } else if (phase === 'props') {
    if (r.Success && r.Data && r.Data.rows) {
      var rows = r.Data.rows;
      rows.forEach(function (row) {
        var match = matchElement(row);
        var key = match.bqSheet || 'UNMATCHED';
        if (!results[cat].matched[key]) results[cat].matched[key] = 0;
        results[cat].matched[key]++;
        allMatched.push({
          category: cat,
          family: row['Element.Family'],
          type: row['Element.Type'],
          systemName: row['Element.System Name'],
          size: row['Element.Size'],
          bqSheet: match.bqSheet,
          matchedBy: match.matchedBy,
          total: results[cat].total
        });
      });
      console.log('  Found ' + results[cat].total + ' → ' + JSON.stringify(results[cat].matched));
    } else {
      console.log('  Error reading properties:', r.Error);
    }
    catIdx++;
    processNextCategory();
  }
});

function outputResults() {
  console.log('\n\n========================================');
  console.log('MATCHING ENGINE RESULTS');
  console.log('========================================\n');

  // 彙總
  var bqTotals = {};
  for (var cat in results) {
    for (var sheet in results[cat].matched) {
      if (!bqTotals[sheet]) bqTotals[sheet] = { categories: [], estimatedTotal: 0 };
      bqTotals[sheet].categories.push(cat + '(' + results[cat].total + ')');
      bqTotals[sheet].estimatedTotal += results[cat].total;
    }
  }

  console.log('BQ Sheet ← Model Categories Mapping:\n');
  Object.keys(bqTotals).sort().forEach(function (sheet) {
    var info = bqTotals[sheet];
    console.log(sheet.padEnd(30) + ' | ~' + info.estimatedTotal + ' elements');
    console.log('  Categories: ' + info.categories.join(', '));
    console.log('');
  });

  // 存檔
  var output = {
    summary: bqTotals,
    details: allMatched,
    categoryResults: results
  };
  fs.writeFileSync(path.join(__dirname, 'matching_results.json'), JSON.stringify(output, null, 2));
  console.log('\nSaved to matching_results.json');

  ws.close();
  process.exit(0);
}

ws.on('error', function (e) { console.log('ERROR:', e.message); process.exit(1); });
setTimeout(function () { console.log('TIMEOUT'); process.exit(1); }, 600000);
