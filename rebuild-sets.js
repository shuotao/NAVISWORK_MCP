/**
 * Rebuild all Search Sets: delete all existing, then create clean set
 */
const WebSocket = require('ws');
const ws = new WebSocket('ws://localhost:2233/');

function send(cmd, params, id) {
  ws.send(JSON.stringify({ command: cmd, parameters: params, requestId: id }));
}

var sets = [
  // 1. CSA
  { folder: '1. CSA', name: 'Architecture - Non-CR (CUB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_CUB_B001_A_ARCH.nwc' },
  { folder: '1. CSA', name: 'Architecture - CR (FAB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_FAB_B001_A_ARCH.nwc' },
  { folder: '1. CSA', name: 'Structure - Non-CR (CUB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_CUB_B001_S_STRU.nwc' },
  { folder: '1. CSA', name: 'Structure - CR (FAB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_FAB_B001_S_STRU.nwc' },
  { folder: '1. CSA', name: 'Civil Works (SIT)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_SIT_B001_C_CIVIL.nwc' },
  // 2. HVAC
  { folder: '2. HVAC', name: 'SA - Supply Air', cat: 'Element', prop: 'System Abbreviation', val: 'SA' },
  { folder: '2. HVAC', name: 'RA - Return Air', cat: 'Element', prop: 'System Abbreviation', val: 'RA' },
  { folder: '2. HVAC', name: 'EA - Exhaust Air', cat: 'Element', prop: 'System Abbreviation', val: 'EA' },
  { folder: '2. HVAC', name: 'MA - Make Up Air', cat: 'Element', prop: 'System Abbreviation', val: 'MA' },
  { folder: '2. HVAC', name: 'OA - Outside Air', cat: 'Element', prop: 'System Abbreviation', val: 'OA' },
  { folder: '2. HVAC', name: 'HVAC - Non-CR (CUB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_CUB_B001_M_DUCT.nwc' },
  { folder: '2. HVAC', name: 'HVAC - CR (FAB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_FAB_B001_M_DUCT.nwc' },
  // 3. Exhaust
  { folder: '3. Exhaust', name: 'AE - Acid Exhaust', cat: 'Element', prop: 'System Abbreviation', val: 'AE' },
  { folder: '3. Exhaust', name: 'ACID EXHAUST (Full)', cat: 'Element', prop: 'System Abbreviation', val: 'ACID EXHAUST' },
  { folder: '3. Exhaust', name: 'Exhaust Ducting (CUB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_CUB_B001_M_EGEX.nwc' },
  { folder: '3. Exhaust', name: 'Exhaust Equipment (FAB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_FAB_B001_M_EGEX.nwc' },
  { folder: '3. Exhaust', name: 'Exhaust Equipment (Shared)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_ALL_B001_M_EQPM.nwd' },
  // 4. Process Utility
  { folder: '4. Process Utility', name: 'CDA - Compressed Dry Air', cat: 'Element', prop: 'System Abbreviation', val: 'CDA' },
  { folder: '4. Process Utility', name: 'PCWS - Process Cooling Water Supply', cat: 'Element', prop: 'System Abbreviation', val: 'PCWS' },
  { folder: '4. Process Utility', name: 'PCWR - Process Cooling Water Return', cat: 'Element', prop: 'System Abbreviation', val: 'PCWR' },
  { folder: '4. Process Utility', name: 'CHWS - Chilled Water Supply', cat: 'Element', prop: 'System Abbreviation', val: 'CHWS' },
  { folder: '4. Process Utility', name: 'CHWR - Chilled Water Return', cat: 'Element', prop: 'System Abbreviation', val: 'CHWR' },
  { folder: '4. Process Utility', name: 'HWS - Hot Water Supply', cat: 'Element', prop: 'System Abbreviation', val: 'HWS' },
  { folder: '4. Process Utility', name: 'HWR - Hot Water Return', cat: 'Element', prop: 'System Abbreviation', val: 'HWR' },
  { folder: '4. Process Utility', name: 'CWS - Cold Water Supply', cat: 'Element', prop: 'System Abbreviation', val: 'CWS' },
  { folder: '4. Process Utility', name: 'CWR - Cold Water Return', cat: 'Element', prop: 'System Abbreviation', val: 'CWR' },
  { folder: '4. Process Utility', name: 'Utility All (CDA/HPCDA/PV/ICA)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_ALL_D_INAP.nwd' },
  // 5. PHE
  { folder: '5. PHE', name: 'DCW - Domestic Cold Water', cat: 'Element', prop: 'System Abbreviation', val: 'DCW' },
  { folder: '5. PHE', name: 'SAN - Sanitary', cat: 'Element', prop: 'System Abbreviation', val: 'SAN' },
  { folder: '5. PHE', name: 'Plumbing - CUB', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_CUB_B001_P_PLUM.nwc' },
  { folder: '5. PHE', name: 'Plumbing - FAB', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_FAB_B001_P_PLUM.nwc' },
  // 6. Waste Water
  { folder: '6. Waste Water', name: 'Waste Water (All)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_ALL_B001_D_PRWT.nwd' },
  // 7. Gas
  { folder: '7. Gas', name: 'Gas Piping (All)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_ALL_B001_N_GASS.nwd' },
  // 8. Electrical
  { folder: '8. Electrical', name: 'Electrical - Non-CR (CUB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_CUB_B001_E_ELEC.nwc' },
  { folder: '8. Electrical', name: 'Electrical - CR (FAB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_FAB_B001_E_ELEC.nwc' },
  // 9. ELV
  { folder: '9. ELV', name: 'ELV / FMCS - CUB', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_CUB_B001_I_INST.nwc' },
  { folder: '9. ELV', name: 'ELV / FMCS - FAB', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_FAB_B001_I_INST.nwc' },
  // 10. Fire Protection
  { folder: '10. Fire Protection', name: 'Fire - Hydrant (CUB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_CUB_B001_F_FPRT.nwc' },
  { folder: '10. Fire Protection', name: 'Fire - Sprinkler (FAB)', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_FAB_B001_F_FPRT.nwc' },
  // 11. PCW
  { folder: '11. PCW', name: 'PCW - CUB', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_CUB_B001_M_CHWT.nwc' },
  { folder: '11. PCW', name: 'PCW - FAB', cat: 'Item', prop: 'Source File', val: 'DE_MSI_001_FAB_B001_M_CHWT.nwc' },
];

var idx = 0;
var step = 'delete';

ws.on('open', function () {
  console.log('Step 1: Delete all existing Search Sets...');
  send('delete_selection_sets', { all: true }, 'del');
});

ws.on('message', function (data) {
  var r = JSON.parse(data.toString());

  if (step === 'delete') {
    console.log(r.Success ? 'Deleted: ' + r.Data.deletedCount : 'ERR: ' + r.Error);
    console.log('\nStep 2: Creating ' + sets.length + ' Search Sets...\n');
    step = 'create';
    nextSet();
  } else if (step === 'create') {
    console.log(r.Success ? 'OK' : 'ERR: ' + (r.Error || '').substring(0, 50));
    idx++;
    setTimeout(nextSet, 200);
  }
});

function nextSet() {
  if (idx >= sets.length) {
    console.log('\nDone! ' + sets.length + ' Search Sets created (0 duplicates).');
    ws.close();
    process.exit(0);
    return;
  }
  var s = sets[idx];
  process.stdout.write('[' + (idx + 1) + '/' + sets.length + '] ' + s.name.padEnd(40));
  send('create_search_set', {
    name: s.name, category: s.cat, property: s.prop, value: s.val, folder: s.folder
  }, 'ss');
}

ws.on('error', function (e) { console.log('\nERROR:', e.message); process.exit(1); });
setTimeout(function () { console.log('\nTIMEOUT'); process.exit(1); }, 300000);
