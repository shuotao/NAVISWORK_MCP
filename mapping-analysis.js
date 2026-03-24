/**
 * Final analysis: which BOQ sheets have NWC files vs which don't
 */

// 23 NWC files under B001 and their mapped keys
var nwcFiles = [
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
  'CUB_A_ARCH': '4.1 Non-CR Architecture',
  'FAB_A_ARCH': '4.2 CR Architecture',
  'CUB_S_STRU': '3. Structure', 'FAB_S_STRU': '3. Structure',
  'CUB_M_CHWT': '7.4 PCW', 'FAB_M_CHWT': '7.4 PCW',
  'CUB_M_DUCT': '5.1 Non-CR HVAC', 'FAB_M_DUCT': '5.2 CR HVAC',
  'CUB_P_PLUM': '2.3 PHE Works', 'FAB_P_PLUM': '2.3 PHE Works',
  'CUB_F_FPRT': '11.1 INTERNAL HYDRANT SYSTEM', 'FAB_F_FPRT': '11.2 SPRINKLER SYSTEM',
  'CUB_E_ELEC': '10.1 Non-CR Elect.', 'FAB_E_ELEC': '10.2 CR Elect.',
  'CUB_I_INST': '9.5 FMCS', 'FAB_I_INST': '9.5 FMCS',
  'SIT_C_CIVIL': '2.1 Civil Works',
  'ALL_M_EQPM': '6.2 Exhaust Equipment',
  'ALL_D_PRWT': '8 Waste Water', 'ALL_N_GASS': '12. Gas',
  'CUB_M_EGEX': '6.1 Exhaust Ducting', 'FAB_M_EGEX': '6.2 Exhaust Equipment',
  'ALL_D_INAP': '7.1 Utility_CDA',
};

// All BOQ sheets
var allBQ = [
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
  '11.3 FOAM SYSTEM', '11.4 GASIOUS SYSTEM',
  '12. Gas'
];

// Map each NWC file to its key and BQ sheet
console.log('=== NWC → BQ MAPPING (all 23 files) ===\n');
var coveredBQ = new Set();
nwcFiles.forEach(function (nwc) {
  var key = nwc.replace('DE_MSI_001_', '').replace('_B001_', '_').replace('.nwc', '').replace('.nwd', '');
  var bq = nwcToBQ[key] || 'UNMAPPED';
  coveredBQ.add(bq);
  console.log('  ' + key.padEnd(15) + ' → ' + bq);
});

console.log('\n=== BOQ SHEETS STATUS ===\n');
var hasNwc = [];
var noNwc = [];
allBQ.forEach(function (bq) {
  if (coveredBQ.has(bq)) {
    hasNwc.push(bq);
    console.log('  ✓ ' + bq);
  } else {
    noNwc.push(bq);
    console.log('  ✗ ' + bq + '  ← NO NWC FILE');
  }
});

console.log('\nCovered: ' + hasNwc.length + '/' + allBQ.length);
console.log('Missing: ' + noNwc.length);

console.log('\n=== ANALYSIS OF MISSING SHEETS ===\n');
var analysis = {
  '2.2 Infra Works': 'No SIT_I_INFRA NWC. May be in B000 site model.',
  '7.2 Utility_HPCDA': 'No HPCDA NWC. May not be modeled yet or bundled in another utility.',
  '7.3 PV': 'No PV NWC. May not be modeled yet.',
  '7.5 ICA+SW': 'No ICA/SW NWC. May not be modeled yet.',
  '9.1 ACS': 'Likely inside I_INST (mapped to 9.5 FMCS). Need element-level sub-classification.',
  '9.2 FAS': 'Likely inside I_INST. Fire Alarm Devices (170 found) should map here instead of FMCS.',
  '9.3 CCTV': 'Likely inside I_INST. Security Devices (147 found) should map here.',
  '9.4 PA': 'Likely inside I_INST. Communication Devices (106 found) should map here.',
  '9.6 IT': 'Likely inside I_INST. Data Devices (197 found) should map here.',
  '11.3 FOAM SYSTEM': 'Likely inside F_FPRT. Need element-level sub-classification by System Type.',
  '11.4 GASIOUS SYSTEM': 'Likely inside F_FPRT. Need element-level sub-classification by System Type.',
};
noNwc.forEach(function (bq) {
  console.log('  ' + bq + ':');
  console.log('    ' + (analysis[bq] || 'Unknown'));
});
