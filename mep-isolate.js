/**
 * MEP System Isolation v4:
 * 1. Create Search Sets for each system (safe, no hide/show)
 * 2. For each system: select → set top view → save viewpoint
 * No isolate_selection — just save viewpoint with top view camera angle
 */
const WebSocket = require('ws');
const ws = new WebSocket('ws://localhost:2233/');

function send(cmd, params, id) {
  ws.send(JSON.stringify({ command: cmd, parameters: params, requestId: id }));
}

var systems = [
  { abbr: 'SAN', name: 'Sanitary (SAN)' },
  { abbr: 'AE', name: 'Acid Exhaust (AE)' },
  { abbr: 'CHWR', name: 'Chilled Water Return (CHWR)' },
  { abbr: 'OA', name: 'Outside Air (OA)' },
  { abbr: 'ACID EXHAUST', name: 'Acid Exhaust (Full)' },
  { abbr: 'DCW', name: 'Domestic Cold Water (DCW)' },
  { abbr: 'CHWS', name: 'Chilled Water Supply (CHWS)' },
  { abbr: 'CWR', name: 'Cold Water Return (CWR)' },
  { abbr: 'CWS', name: 'Cold Water Supply (CWS)' },
  { abbr: 'HWS', name: 'Hot Water Supply (HWS)' },
  { abbr: 'MA', name: 'Make Up Air (MA)' },
  { abbr: 'RA', name: 'Return Air (RA)' },
  { abbr: 'PCWR', name: 'PCW Return (PCWR)' },
  { abbr: 'PCWS', name: 'PCW Supply (PCWS)' },
  { abbr: 'HWR', name: 'Hot Water Return (HWR)' },
  { abbr: 'EA', name: 'Exhaust Air (EA)' },
  { abbr: 'SA', name: 'Supply Air (SA)' },
  { abbr: 'CDA', name: 'CDA Utility' },
];

var FOLDER = 'MEP Systems';
var idx = 0;
var step = '';
var phase = 'search_sets'; // Phase 1: search sets, Phase 2: viewpoints

ws.on('open', function () {
  console.log('=== Phase 1: Creating ' + systems.length + ' Search Sets ===\n');
  nextSearchSet();
});

// ─── Phase 1: Create Search Sets ───
function nextSearchSet() {
  if (idx >= systems.length) {
    console.log('\n=== Phase 2: Creating viewpoints (Top View) ===\n');
    idx = 0;
    phase = 'viewpoints';
    nextViewpoint();
    return;
  }
  var sys = systems[idx];
  process.stdout.write('[' + (idx + 1) + '/' + systems.length + '] ' + sys.name.padEnd(35));
  step = 'create_ss';
  send('create_search_set', {
    name: sys.name,
    category: 'Element',
    property: 'System Abbreviation',
    value: sys.abbr,
    folder: FOLDER
  }, 'ss');
}

// ─── Phase 2: Select + Top View + Save Viewpoint ───
function nextViewpoint() {
  if (idx >= systems.length) {
    console.log('\nDone! ' + systems.length + ' search sets + viewpoints created in "' + FOLDER + '".');
    ws.close();
    process.exit(0);
    return;
  }
  var sys = systems[idx];
  process.stdout.write('[' + (idx + 1) + '/' + systems.length + '] ' + sys.name.padEnd(35));
  step = 'vp_select';
  send('select_items_by_search', {
    category: 'Element', property: 'System Abbreviation', value: sys.abbr
  }, 'vpsel');
}

ws.on('message', function (data) {
  var r = JSON.parse(data.toString());

  // Phase 1: Search Sets
  if (phase === 'search_sets') {
    if (step === 'create_ss') {
      console.log(r.Success ? 'OK' : 'ERR: ' + (r.Error || '').substring(0, 50));
      idx++;
      setTimeout(nextSearchSet, 500);
    }
    return;
  }

  // Phase 2: Viewpoints
  if (step === 'vp_select') {
    var count = r.Data ? r.Data.foundCount : 0;
    process.stdout.write(count + ' → ');
    step = 'vp_top';
    send('set_view_top', {}, 'top');
  } else if (step === 'vp_top') {
    step = 'vp_zoom';
    send('zoom_to_selection', {}, 'zoom');
  } else if (step === 'vp_zoom') {
    step = 'vp_save';
    send('save_viewpoint', { name: systems[idx].name + ' (Top)' }, 'save');
  } else if (step === 'vp_save') {
    console.log(r.Success ? 'SAVED' : 'ERR: ' + (r.Error || '').substring(0, 40));
    idx++;
    setTimeout(nextViewpoint, 2000);
  }
});

ws.on('error', function (e) { console.log('\nERROR:', e.message); process.exit(1); });
setTimeout(function () { console.log('\nTIMEOUT at ' + idx); process.exit(1); }, 600000);
