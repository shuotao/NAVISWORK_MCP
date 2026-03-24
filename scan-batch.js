const WebSocket = require('ws');
const fs = require('fs');

var allNwc = [
  'DE_MSI_001_CUB_B001_A_ARCH.nwc',
  'DE_MSI_001_FAB_B001_A_ARCH.nwc',
  'DE_MSI_001_CUB_B001_S_STRU.nwc',
  'DE_MSI_001_FAB_B001_S_STRU.nwc',
  'DE_MSI_001_CUB_B001_M_CHWT.nwc',
  'DE_MSI_001_CUB_B001_M_DUCT.nwc',
  'DE_MSI_001_FAB_B001_M_CHWT.nwc',
  'DE_MSI_001_FAB_B001_M_DUCT.nwc',
  'DE_MSI_001_CUB_B001_P_PLUM.nwc',
  'DE_MSI_001_FAB_B001_P_PLUM.nwc',
  'DE_MSI_001_CUB_B001_F_FPRT.nwc',
  'DE_MSI_001_FAB_B001_F_FPRT.nwc',
  'DE_MSI_001_CUB_B001_E_ELEC.nwc',
  'DE_MSI_001_FAB_B001_E_ELEC.nwc',
  'DE_MSI_001_CUB_B001_I_INST.nwc',
  'DE_MSI_001_FAB_B001_I_INST.nwc',
  'DE_MSI_001_SIT_B001_C_CIVIL.nwc',
  'DE_MSI_001_ALL_B001_M_EQPM.nwd',
  'DE_MSI_001_ALL_B001_D_PRWT.nwd',
  'DE_MSI_001_ALL_B001_N_GASS.nwd',
  'DE_MSI_001_CUB_B001_M_EGEX.nwc',
  'DE_MSI_001_FAB_B001_M_EGEX.nwc',
  'DE_MSI_001_ALL_D_INAP.nwd',
];

var nwcToBQ = {
  'CUB_A_ARCH': '4.1 Non-CR Architecture', 'FAB_A_ARCH': '4.2 CR Architecture',
  'CUB_S_STRU': '3. Structure', 'FAB_S_STRU': '3. Structure',
  'CUB_M_CHWT': '7.4 PCW', 'FAB_M_CHWT': '7.4 PCW',
  'CUB_M_DUCT': '5.1 Non-CR HVAC', 'FAB_M_DUCT': '5.2 CR HVAC',
  'CUB_P_PLUM': '2.3 PHE Works', 'FAB_P_PLUM': '2.3 PHE Works',
  'CUB_F_FPRT': '11.1 INTERNAL HYDRANT SYSTEM', 'FAB_F_FPRT': '11.2 SPRINKLER SYSTEM',
  'CUB_E_ELEC': '10.1 Non-CR Elect.', 'FAB_E_ELEC': '10.2 CR Elect.',
  'CUB_I_INST': '9.5 FMCS', 'FAB_I_INST': '9.5 FMCS',
  'SIT_C_CIVIL': '2.1 Civil Works', 'ALL_M_EQPM': '6.2 Exhaust Equipment',
  'ALL_D_PRWT': '8 Waste Water', 'ALL_N_GASS': '12. Gas',
  'CUB_M_EGEX': '6.1 Exhaust Ducting', 'FAB_M_EGEX': '6.2 Exhaust Equipment',
  'ALL_D_INAP': '7.1 Utility_CDA',
};

var instSubMap = {
  'Fire Alarm Devices': '9.2 FAS', 'Security Devices': '9.3 CCTV',
  'Communication Devices': '9.4 PA', 'Data Devices': '9.6 IT',
  'Electrical Equipment': '9.1 ACS',
};

// Batch control from CLI args
var batchStart = parseInt(process.argv[2]) || 0;
var batchSize = parseInt(process.argv[3]) || 5;
var nwcFiles = allNwc.slice(batchStart, batchStart + batchSize);

var idx = 0;
var allResults = {};
var ws = new WebSocket('ws://localhost:2233/');

function send(cmd, params, id) {
  ws.send(JSON.stringify({ command: cmd, parameters: params, requestId: id }));
}

ws.on('open', function () {
  console.log('Batch ' + batchStart + '-' + (batchStart + nwcFiles.length - 1) + ' (' + nwcFiles.length + ' files)\n');
  scanNext();
});

function scanNext() {
  if (idx >= nwcFiles.length) { save(); return; }
  send('scan_subtree', { name: nwcFiles[idx], maxItems: 100000 }, 'scan_' + idx);
}

ws.on('message', function (data) {
  var r = JSON.parse(data.toString());
  var nwc = nwcFiles[idx];
  var key = nwc.replace('DE_MSI_001_', '').replace('_B001_', '_').replace('.nwc', '').replace('.nwd', '');
  var baseBQ = nwcToBQ[key] || 'UNMAPPED:' + key;
  var isInst = (key === 'CUB_I_INST' || key === 'FAB_I_INST');

  if (r.Success && r.Data) {
    process.stdout.write(key.padEnd(16) + ' → ' + baseBQ.padEnd(30) + ' | geo: ' + (r.Data.totalGeometry || 0) + '\n');
    (r.Data.categories || []).forEach(function (cat) {
      var bq = isInst && instSubMap[cat.category] ? instSubMap[cat.category] : baseBQ;
      if (!allResults[bq]) allResults[bq] = {};
      if (!allResults[bq][cat.category]) allResults[bq][cat.category] = {};
      cat.types.forEach(function (t) {
        allResults[bq][cat.category][t.typeKey] = (allResults[bq][cat.category][t.typeKey] || 0) + t.count;
      });
    });
  } else {
    process.stdout.write(key.padEnd(16) + ' → ERROR: ' + (r.Error || '') + '\n');
  }
  idx++;
  scanNext();
});

function save() {
  var file = __dirname + '/b001_full_scan.json';
  var existing = {};
  try { existing = JSON.parse(fs.readFileSync(file, 'utf8')); } catch(e) {}
  Object.keys(allResults).forEach(function (bq) {
    if (!existing[bq]) existing[bq] = {};
    Object.keys(allResults[bq]).forEach(function (cat) {
      if (!existing[bq][cat]) existing[bq][cat] = {};
      Object.keys(allResults[bq][cat]).forEach(function (tk) {
        existing[bq][cat][tk] = (existing[bq][cat][tk] || 0) + allResults[bq][cat][tk];
      });
    });
  });
  fs.writeFileSync(file, JSON.stringify(existing, null, 2));
  console.log('\nBatch done. Merged into b001_full_scan.json');
  ws.close();
  process.exit(0);
}

ws.on('error', function (e) { console.log('ERROR:', e.message); process.exit(1); });
setTimeout(function () { console.log('TIMEOUT at idx ' + idx); process.exit(1); }, 300000);
