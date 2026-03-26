/**
 * Create Search Sets based on project RFP system classifications
 * Systems from MSI-1-MH-DOC-G-00252_Rev_1.docx
 */
const WebSocket = require('ws');
const ws = new WebSocket('ws://localhost:2233/');

function send(cmd, params, id) {
  ws.send(JSON.stringify({ command: cmd, parameters: params, requestId: id }));
}

// ─── Systems by discipline (from RFP abbreviations + scope) ───
var sets = [
  // === CSA ===
  { folder: '1. CSA', name: 'Architecture - Non-CR (CUB)', type: 'nwc', value: 'DE_MSI_001_CUB_B001_A_ARCH.nwc' },
  { folder: '1. CSA', name: 'Architecture - CR (FAB)', type: 'nwc', value: 'DE_MSI_001_FAB_B001_A_ARCH.nwc' },
  { folder: '1. CSA', name: 'Structure - Non-CR (CUB)', type: 'nwc', value: 'DE_MSI_001_CUB_B001_S_STRU.nwc' },
  { folder: '1. CSA', name: 'Structure - CR (FAB)', type: 'nwc', value: 'DE_MSI_001_FAB_B001_S_STRU.nwc' },
  { folder: '1. CSA', name: 'Civil Works (SIT)', type: 'nwc', value: 'DE_MSI_001_SIT_B001_C_CIVIL.nwc' },

  // === HVAC ===
  { folder: '2. HVAC', name: 'SA - Supply Air', type: 'sys', value: 'SA' },
  { folder: '2. HVAC', name: 'RA - Return Air', type: 'sys', value: 'RA' },
  { folder: '2. HVAC', name: 'EA - Exhaust Air', type: 'sys', value: 'EA' },
  { folder: '2. HVAC', name: 'MA - Make Up Air', type: 'sys', value: 'MA' },
  { folder: '2. HVAC', name: 'OA - Outside Air', type: 'sys', value: 'OA' },

  // === Exhaust ===
  { folder: '3. Exhaust', name: 'AE - Acid Exhaust', type: 'sys', value: 'AE' },
  { folder: '3. Exhaust', name: 'ACID EXHAUST (Full System)', type: 'sys', value: 'ACID EXHAUST' },
  { folder: '3. Exhaust', name: 'Exhaust Ducting (CUB)', type: 'nwc', value: 'DE_MSI_001_CUB_B001_M_EGEX.nwc' },
  { folder: '3. Exhaust', name: 'Exhaust Equipment (FAB)', type: 'nwc', value: 'DE_MSI_001_FAB_B001_M_EGEX.nwc' },
  { folder: '3. Exhaust', name: 'Exhaust Equipment (Shared)', type: 'nwc', value: 'DE_MSI_001_ALL_B001_M_EQPM.nwd' },

  // === Process Utility ===
  { folder: '4. Process Utility', name: 'CDA - Compressed Dry Air', type: 'sys', value: 'CDA' },
  { folder: '4. Process Utility', name: 'PCWS - Process Cooling Water Supply', type: 'sys', value: 'PCWS' },
  { folder: '4. Process Utility', name: 'PCWR - Process Cooling Water Return', type: 'sys', value: 'PCWR' },
  { folder: '4. Process Utility', name: 'CHWS - Chilled Water Supply', type: 'sys', value: 'CHWS' },
  { folder: '4. Process Utility', name: 'CHWR - Chilled Water Return', type: 'sys', value: 'CHWR' },
  { folder: '4. Process Utility', name: 'HWS - Hot Water Supply', type: 'sys', value: 'HWS' },
  { folder: '4. Process Utility', name: 'HWR - Hot Water Return', type: 'sys', value: 'HWR' },
  { folder: '4. Process Utility', name: 'CWS - Cold Water Supply', type: 'sys', value: 'CWS' },
  { folder: '4. Process Utility', name: 'CWR - Cold Water Return', type: 'sys', value: 'CWR' },
  { folder: '4. Process Utility', name: 'Utility All (CDA/HPCDA/PV/ICA)', type: 'nwc', value: 'DE_MSI_001_ALL_D_INAP.nwd' },

  // === PHE / Plumbing ===
  { folder: '5. PHE', name: 'DCW - Domestic Cold Water', type: 'sys', value: 'DCW' },
  { folder: '5. PHE', name: 'SAN - Sanitary', type: 'sys', value: 'SAN' },
  { folder: '5. PHE', name: 'Plumbing - CUB', type: 'nwc', value: 'DE_MSI_001_CUB_B001_P_PLUM.nwc' },
  { folder: '5. PHE', name: 'Plumbing - FAB', type: 'nwc', value: 'DE_MSI_001_FAB_B001_P_PLUM.nwc' },

  // === Waste Water ===
  { folder: '6. Waste Water', name: 'Waste Water (All)', type: 'nwc', value: 'DE_MSI_001_ALL_B001_D_PRWT.nwd' },

  // === Gas ===
  { folder: '7. Gas', name: 'Gas Piping (All)', type: 'nwc', value: 'DE_MSI_001_ALL_B001_N_GASS.nwd' },

  // === Electrical ===
  { folder: '8. Electrical', name: 'Electrical - Non-CR (CUB)', type: 'nwc', value: 'DE_MSI_001_CUB_B001_E_ELEC.nwc' },
  { folder: '8. Electrical', name: 'Electrical - CR (FAB)', type: 'nwc', value: 'DE_MSI_001_FAB_B001_E_ELEC.nwc' },

  // === ELV ===
  { folder: '9. ELV', name: 'ELV / FMCS (All)', type: 'nwc', value: 'DE_MSI_001_CUB_B001_I_INST.nwc' },
  { folder: '9. ELV', name: 'ELV / FMCS - FAB', type: 'nwc', value: 'DE_MSI_001_FAB_B001_I_INST.nwc' },

  // === Fire Protection ===
  { folder: '10. Fire Protection', name: 'Fire - Hydrant (CUB)', type: 'nwc', value: 'DE_MSI_001_CUB_B001_F_FPRT.nwc' },
  { folder: '10. Fire Protection', name: 'Fire - Sprinkler (FAB)', type: 'nwc', value: 'DE_MSI_001_FAB_B001_F_FPRT.nwc' },

  // === PCW (dedicated NWC) ===
  { folder: '11. PCW', name: 'PCW - CUB', type: 'nwc', value: 'DE_MSI_001_CUB_B001_M_CHWT.nwc' },
  { folder: '11. PCW', name: 'PCW - FAB', type: 'nwc', value: 'DE_MSI_001_FAB_B001_M_CHWT.nwc' },
];

var idx = 0;

ws.on('open', function () {
  console.log('Creating ' + sets.length + ' Search Sets from project RFP...\n');
  nextSet();
});

function nextSet() {
  if (idx >= sets.length) {
    console.log('\nDone! ' + sets.length + ' Search Sets created.');
    ws.close();
    process.exit(0);
    return;
  }
  var s = sets[idx];
  process.stdout.write('[' + (idx + 1) + '/' + sets.length + '] ' + s.name.padEnd(40));

  if (s.type === 'sys') {
    // Search by System Abbreviation
    send('create_search_set', {
      name: s.name, category: 'Element', property: 'System Abbreviation', value: s.value, folder: s.folder
    }, 'ss');
  } else {
    // Search by source file (NWC)
    send('create_search_set', {
      name: s.name, category: 'Item', property: 'Source File', value: s.value, folder: s.folder
    }, 'ss');
  }
}

ws.on('message', function (data) {
  var r = JSON.parse(data.toString());
  console.log(r.Success ? 'OK' : 'ERR: ' + (r.Error || '').substring(0, 50));
  idx++;
  setTimeout(nextSet, 300);
});

ws.on('error', function (e) { console.log('\nERROR:', e.message); process.exit(1); });
setTimeout(function () { console.log('\nTIMEOUT at ' + idx); process.exit(1); }, 300000);
