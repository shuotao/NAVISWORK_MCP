/**
 * MEP System Isolation: for each system, hide siblings + save viewpoint
 * Uses optimized hide_all_except (sibling-level toggle, no descendant iteration)
 */
const WebSocket = require('ws');
const ws = new WebSocket('ws://localhost:2233/');

function send(cmd, params, id) {
  ws.send(JSON.stringify({ command: cmd, parameters: params, requestId: id }));
}

var systems = [
  { show: ['DE_MSI_001_CUB_B001_M_DUCT.nwc'], name: 'HVAC - Non-CR (CUB)' },
  { show: ['DE_MSI_001_FAB_B001_M_DUCT.nwc'], name: 'HVAC - CR (FAB)' },
  { show: ['DE_MSI_001_CUB_B001_M_DUCT.nwc', 'DE_MSI_001_FAB_B001_M_DUCT.nwc'], name: 'HVAC - All' },
  { show: ['DE_MSI_001_CUB_B001_M_CHWT.nwc'], name: 'PCW - Non-CR (CUB)' },
  { show: ['DE_MSI_001_FAB_B001_M_CHWT.nwc'], name: 'PCW - CR (FAB)' },
  { show: ['DE_MSI_001_CUB_B001_M_CHWT.nwc', 'DE_MSI_001_FAB_B001_M_CHWT.nwc'], name: 'PCW - All' },
  { show: ['DE_MSI_001_CUB_B001_P_PLUM.nwc', 'DE_MSI_001_FAB_B001_P_PLUM.nwc'], name: 'Plumbing (PHE)' },
  { show: ['DE_MSI_001_CUB_B001_F_FPRT.nwc'], name: 'Fire - Hydrant (CUB)' },
  { show: ['DE_MSI_001_FAB_B001_F_FPRT.nwc'], name: 'Fire - Sprinkler (FAB)' },
  { show: ['DE_MSI_001_CUB_B001_F_FPRT.nwc', 'DE_MSI_001_FAB_B001_F_FPRT.nwc'], name: 'Fire Protection - All' },
  { show: ['DE_MSI_001_CUB_B001_E_ELEC.nwc'], name: 'Electrical - Non-CR (CUB)' },
  { show: ['DE_MSI_001_FAB_B001_E_ELEC.nwc'], name: 'Electrical - CR (FAB)' },
  { show: ['DE_MSI_001_CUB_B001_E_ELEC.nwc', 'DE_MSI_001_FAB_B001_E_ELEC.nwc'], name: 'Electrical - All' },
  { show: ['DE_MSI_001_CUB_B001_I_INST.nwc', 'DE_MSI_001_FAB_B001_I_INST.nwc'], name: 'ELV / FMCS' },
  { show: ['DE_MSI_001_CUB_B001_M_EGEX.nwc'], name: 'Exhaust Ducting (CUB)' },
  { show: ['DE_MSI_001_FAB_B001_M_EGEX.nwc'], name: 'Exhaust Equipment (FAB)' },
  { show: ['DE_MSI_001_CUB_B001_M_EGEX.nwc', 'DE_MSI_001_FAB_B001_M_EGEX.nwc', 'DE_MSI_001_ALL_B001_M_EQPM.nwd'], name: 'Exhaust - All' },
  { show: ['DE_MSI_001_ALL_B001_D_PRWT.nwd'], name: 'Waste Water' },
  { show: ['DE_MSI_001_ALL_B001_N_GASS.nwd'], name: 'Gas Piping' },
  { show: ['DE_MSI_001_ALL_D_INAP.nwd'], name: 'Utility (CDA/HPCDA/PV/ICA)' },
];

var FOLDER = 'MEP Systems';
var idx = 0;
var step = '';

ws.on('open', function () {
  console.log('Creating ' + systems.length + ' MEP system viewpoints...\n');
  step = 'unhide_init';
  send('unhide_all', {}, 'init');
});

ws.on('message', function (data) {
  var r = JSON.parse(data.toString());

  if (step === 'unhide_init') {
    nextSystem();
  } else if (step === 'hide') {
    if (!r.Success) {
      console.log('HIDE ERR: ' + (r.Error || '').substring(0, 50));
      idx++; nextSystem(); return;
    }
    // Select first shown NWC for zoom
    step = 'select';
    send('select_subtree', { name: systems[idx].show[0] }, 'sel');
  } else if (step === 'select') {
    step = 'zoom';
    send('zoom_to_selection', {}, 'zoom');
  } else if (step === 'zoom') {
    step = 'save';
    send('save_viewpoint', { name: systems[idx].name, folder: FOLDER }, 'save');
  } else if (step === 'save') {
    console.log(r.Success ? 'OK' : 'ERR: ' + (r.Error || '').substring(0, 50));
    idx++; nextSystem();
  } else if (step === 'final') {
    console.log('\nDone! ' + systems.length + ' viewpoints in "' + FOLDER + '" folder.');
    ws.close(); process.exit(0);
  }
});

function nextSystem() {
  if (idx >= systems.length) {
    step = 'final';
    send('unhide_all', {}, 'final');
    return;
  }
  var sys = systems[idx];
  process.stdout.write('[' + (idx + 1) + '/' + systems.length + '] ' + sys.name.padEnd(35) + ' ');
  step = 'hide';
  if (sys.show.length === 1) {
    send('hide_all_except', { name: sys.show[0] }, 'hide');
  } else {
    send('hide_all_except', { names: sys.show }, 'hide');
  }
}

ws.on('error', function (e) { console.log('ERROR:', e.message); process.exit(1); });
setTimeout(function () { console.log('TIMEOUT at ' + idx); process.exit(1); }, 600000);
