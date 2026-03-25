/**
 * Cross-reference: Model quantities vs BOQ quantities
 * Compares scan_subtree results with Quantification data
 */
const WebSocket = require('ws');
const fs = require('fs');

// Model data
var modelData = JSON.parse(fs.readFileSync(__dirname + '/b001_full_scan.json', 'utf8'));

var ws = new WebSocket('ws://localhost:2233/');

function send(sql, id) {
  ws.send(JSON.stringify({ command: 'quantification_exec_sql', parameters: { sql: sql }, requestId: id }));
}

ws.on('open', function () {
  // Get all groups with their items and quantities
  send(`SELECT g.Name as groupName, g.ID as groupId,
    COUNT(i.ID) as itemCount,
    SUM(CASE WHEN i.PrimaryQuantity_Formula IS NOT NULL THEN 1 ELSE 0 END) as withQty
    FROM TK_ItemGroup g
    LEFT JOIN TK_Item i ON i.Parent = g.ID
    GROUP BY g.ID
    ORDER BY g.Name`, 'groups');
});

var step = 'groups';
var bqGroups = {};
var detailQueue = [];
var detailIdx = 0;

ws.on('message', function (data) {
  var r = JSON.parse(data.toString());

  if (step === 'groups') {
    if (r.Data && r.Data.rows) {
      r.Data.rows.forEach(function (row) {
        bqGroups[row.groupName] = { id: row.groupId, itemCount: row.itemCount, withQty: row.withQty };
        detailQueue.push(row.groupName);
      });
    }
    step = 'details';
    detailIdx = 0;
    fetchNextDetail();

  } else if (step === 'details') {
    var groupName = detailQueue[detailIdx];
    if (r.Data && r.Data.rows) {
      bqGroups[groupName].items = r.Data.rows.map(function (row) {
        // Parse quantity from formula like "=125" or "=0"
        var qtyStr = row.PrimaryQuantity_Formula || '';
        var qty = parseFloat(qtyStr.replace('=', '')) || 0;
        return { name: row.Name, qty: qty, formula: qtyStr };
      });
      // Sum total BQ quantity
      bqGroups[groupName].totalBQQty = bqGroups[groupName].items.reduce(function (s, i) { return s + i.qty; }, 0);
    }
    detailIdx++;
    fetchNextDetail();
  }
});

function fetchNextDetail() {
  if (detailIdx >= detailQueue.length) {
    generateReport();
    return;
  }
  var gid = bqGroups[detailQueue[detailIdx]].id;
  send("SELECT Name, PrimaryQuantity_Formula FROM TK_Item WHERE Parent = " + gid + " ORDER BY Name", 'det');
}

function generateReport() {
  var report = [];
  report.push('# Model vs BOQ Cross-Reference Report');
  report.push('Generated: ' + new Date().toISOString().split('T')[0]);
  report.push('');
  report.push('## Summary');
  report.push('');
  report.push('| BQ Sheet | Model Elements | BOQ Items | BOQ Total Qty | Status |');
  report.push('|----------|---------------|-----------|---------------|--------|');

  var allSheets = new Set();
  Object.keys(modelData).forEach(function (k) { allSheets.add(k); });
  Object.keys(bqGroups).forEach(function (k) { allSheets.add(k); });

  var sorted = Array.from(allSheets).sort();
  var modelOnly = [], bqOnly = [], matched = [], zeroQty = [];

  sorted.forEach(function (sheet) {
    var mCount = 0;
    if (modelData[sheet]) {
      Object.keys(modelData[sheet]).forEach(function (cat) {
        Object.keys(modelData[sheet][cat]).forEach(function (t) {
          mCount += modelData[sheet][cat][t];
        });
      });
    }

    var bq = bqGroups[sheet];
    var bqItems = bq ? bq.itemCount : 0;
    var bqQty = bq ? (bq.totalBQQty || 0) : 0;

    var status;
    if (mCount > 0 && bqItems > 0) {
      status = 'MATCHED';
      matched.push(sheet);
      if (bqQty === 0) { status = 'QTY=0'; zeroQty.push(sheet); }
    } else if (mCount > 0 && bqItems === 0) {
      status = 'MODEL ONLY';
      modelOnly.push(sheet);
    } else if (mCount === 0 && bqItems > 0) {
      status = 'BOQ ONLY';
      bqOnly.push(sheet);
    } else {
      status = 'EMPTY';
    }

    report.push('| ' + sheet.padEnd(33) + ' | ' +
      String(mCount).padStart(13) + ' | ' +
      String(bqItems).padStart(9) + ' | ' +
      String(Math.round(bqQty)).padStart(13) + ' | ' + status + ' |');
  });

  report.push('');
  report.push('## Findings');
  report.push('');
  report.push('### Matched (' + matched.length + ' sheets)');
  report.push('Both model and BOQ have data: ' + matched.join(', '));
  report.push('');

  if (zeroQty.length > 0) {
    report.push('### BOQ Qty = 0 (' + zeroQty.length + ' sheets)');
    report.push('BOQ items exist but all quantities are zero: ' + zeroQty.join(', '));
    report.push('');
  }

  if (modelOnly.length > 0) {
    report.push('### Model Only (' + modelOnly.length + ' sheets)');
    report.push('Model has elements but no BOQ items: ' + modelOnly.join(', '));
    report.push('');
  }

  if (bqOnly.length > 0) {
    report.push('### BOQ Only (' + bqOnly.length + ' sheets)');
    report.push('BOQ items exist but no model elements: ' + bqOnly.join(', '));
    report.push('');
  }

  // Detail per matched sheet: model categories vs BOQ items
  report.push('---');
  report.push('');
  report.push('## Detail per BQ Sheet');
  report.push('');

  sorted.forEach(function (sheet) {
    if (!modelData[sheet] && !bqGroups[sheet]) return;
    report.push('### ' + sheet);
    report.push('');

    // Model side
    if (modelData[sheet]) {
      report.push('**Model Elements:**');
      report.push('| Category | Count |');
      report.push('|----------|-------|');
      var cats = modelData[sheet];
      Object.keys(cats).sort().forEach(function (cat) {
        var count = 0;
        Object.keys(cats[cat]).forEach(function (t) { count += cats[cat][t]; });
        report.push('| ' + cat.padEnd(30) + ' | ' + String(count).padStart(5) + ' |');
      });
      report.push('');
    } else {
      report.push('**Model:** No elements found');
      report.push('');
    }

    // BOQ side - top items with qty > 0
    var bq = bqGroups[sheet];
    if (bq && bq.items && bq.items.length > 0) {
      var withQty = bq.items.filter(function (i) { return i.qty > 0; });
      report.push('**BOQ Items** (' + bq.items.length + ' total, ' + withQty.length + ' with qty > 0):');
      if (withQty.length > 0) {
        report.push('| Item | Qty |');
        report.push('|------|-----|');
        withQty.slice(0, 20).forEach(function (item) {
          var name = item.name.length > 60 ? item.name.substring(0, 60) + '...' : item.name;
          report.push('| ' + name.padEnd(62) + ' | ' + String(item.qty).padStart(8) + ' |');
        });
        if (withQty.length > 20) report.push('| ... and ' + (withQty.length - 20) + ' more items |  |');
      } else {
        report.push('All quantities are 0');
      }
      report.push('');
    } else {
      report.push('**BOQ:** No items');
      report.push('');
    }
  });

  var output = report.join('\n');
  fs.writeFileSync(__dirname + '/Cross_Reference_Report.md', output);
  console.log(output);
  console.log('\n\nSaved: Cross_Reference_Report.md');

  ws.close();
  process.exit(0);
}

ws.on('error', function (e) { console.log('ERROR:', e.message); process.exit(1); });
setTimeout(function () { console.log('TIMEOUT'); process.exit(1); }, 60000);
