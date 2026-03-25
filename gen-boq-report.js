/**
 * Generate BOQ comparison report - exact original format + model columns appended
 * Copies each sheet cell-by-cell, preserves merges, adds model data columns at the end
 */
const XLSX = require('xlsx');
const fs = require('fs');

const EXCEL_PATH = 'C:/Users/Admin/Downloads/Appendix A - MSI-1-MH-BQ-G-00256_Rev_1_Bill of Quantity.xlsx';
const origWb = XLSX.readFile(EXCEL_PATH, { cellStyles: true });
const modelData = JSON.parse(fs.readFileSync(__dirname + '/b001_full_scan.json', 'utf8'));

var newWb = XLSX.utils.book_new();

// Helper: find header row index and its column count
function findHeaderRow(data) {
  for (var i = 0; i < Math.min(data.length, 15); i++) {
    var row = data[i];
    if (!row) continue;
    var joined = row.map(function (c) { return String(c || '').toUpperCase().trim(); }).join(' ');
    var hasSno = joined.match(/S\.?\s*NO|SI\.?\s*NO|SR\.?\s*NO/);
    var hasDesc = joined.indexOf('DESCRIPTION') >= 0 || joined.indexOf('ITEM') >= 0;
    if (hasSno && hasDesc) return i;
  }
  return -1;
}

origWb.SheetNames.forEach(function (sheetName) {
  var wsOrig = origWb.Sheets[sheetName];
  if (!wsOrig) return;

  // Deep copy the original worksheet
  var wsNew = {};
  Object.keys(wsOrig).forEach(function (key) {
    if (key === '!ref' || key === '!merges' || key === '!cols' || key === '!rows') {
      wsNew[key] = JSON.parse(JSON.stringify(wsOrig[key]));
    } else if (key.charAt(0) !== '!') {
      // Copy cell
      wsNew[key] = Object.assign({}, wsOrig[key]);
    } else {
      wsNew[key] = wsOrig[key];
    }
  });

  // Get data and range
  var data = XLSX.utils.sheet_to_json(wsOrig, { header: 1 });
  var range = XLSX.utils.decode_range(wsOrig['!ref'] || 'A1');
  var origMaxCol = range.e.c;

  // Find model data for this sheet
  var modelSheet = modelData[sheetName] || modelData[sheetName.trim()];

  // Determine where to put model columns (2 cols gap after original)
  var modelStartCol = origMaxCol + 2;

  // Build model summary
  var modelCats = [];
  var totalModelElements = 0;
  if (modelSheet) {
    Object.keys(modelSheet).sort().forEach(function (cat) {
      var catTotal = 0;
      var types = [];
      Object.keys(modelSheet[cat]).forEach(function (t) {
        catTotal += modelSheet[cat][t];
        types.push({ name: t, count: modelSheet[cat][t] });
      });
      totalModelElements += catTotal;
      types.sort(function (a, b) { return b.count - a.count; });
      modelCats.push({ category: cat, count: catTotal, types: types });
    });
  }

  // Write model header at the header row
  var headerIdx = findHeaderRow(data);
  if (headerIdx < 0) headerIdx = 6; // default

  // Title row for model section
  var titleRow = Math.max(0, headerIdx - 2);
  setCell(wsNew, titleRow, modelStartCol, 'MODEL COMPARISON', 's');
  setCell(wsNew, titleRow, modelStartCol + 1, totalModelElements > 0 ? totalModelElements + ' elements' : 'Not in B001 model', 's');

  // Model column headers
  setCell(wsNew, headerIdx, modelStartCol, 'Model Category', 's');
  setCell(wsNew, headerIdx, modelStartCol + 1, 'Model Count', 's');
  setCell(wsNew, headerIdx, modelStartCol + 2, 'Top Family | Type', 's');
  setCell(wsNew, headerIdx, modelStartCol + 3, 'Type Count', 's');

  // Write model categories starting from headerIdx + 1
  var modelRow = headerIdx + 1;
  modelCats.forEach(function (mc) {
    setCell(wsNew, modelRow, modelStartCol, mc.category, 's');
    setCell(wsNew, modelRow, modelStartCol + 1, mc.count, 'n');
    // Top type
    if (mc.types.length > 0) {
      setCell(wsNew, modelRow, modelStartCol + 2, mc.types[0].name, 's');
      setCell(wsNew, modelRow, modelStartCol + 3, mc.types[0].count, 'n');
    }
    modelRow++;
    // Additional top types (up to 3)
    for (var ti = 1; ti < Math.min(mc.types.length, 4); ti++) {
      setCell(wsNew, modelRow, modelStartCol + 2, mc.types[ti].name, 's');
      setCell(wsNew, modelRow, modelStartCol + 3, mc.types[ti].count, 'n');
      modelRow++;
    }
  });

  // Total row
  if (modelCats.length > 0) {
    modelRow++;
    setCell(wsNew, modelRow, modelStartCol, 'TOTAL', 's');
    setCell(wsNew, modelRow, modelStartCol + 1, totalModelElements, 'n');
  }

  // Update range to include model columns
  var newMaxCol = modelStartCol + 3;
  var newMaxRow = Math.max(range.e.r, modelRow);
  wsNew['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: newMaxRow, c: newMaxCol } });

  // Set column widths for model columns
  if (!wsNew['!cols']) wsNew['!cols'] = [];
  while (wsNew['!cols'].length <= newMaxCol) wsNew['!cols'].push(null);
  wsNew['!cols'][modelStartCol] = { wch: 28 };
  wsNew['!cols'][modelStartCol + 1] = { wch: 12 };
  wsNew['!cols'][modelStartCol + 2] = { wch: 45 };
  wsNew['!cols'][modelStartCol + 3] = { wch: 12 };

  // Add sheet
  var name = sheetName.trim();
  if (name.length > 31) name = name.substring(0, 31);
  try {
    XLSX.utils.book_append_sheet(newWb, wsNew, name);
    console.log('[OK] ' + name.padEnd(32) + ' | model: ' + totalModelElements);
  } catch (e) {
    console.log('[ERR] ' + name + ': ' + e.message);
  }
});

function setCell(ws, r, c, val, type) {
  var addr = XLSX.utils.encode_cell({ r: r, c: c });
  if (type === 'n') {
    ws[addr] = { t: 'n', v: val };
  } else {
    ws[addr] = { t: 's', v: String(val) };
  }
}

var outPath = __dirname + '/BOQ_vs_Model_Report.xlsx';
XLSX.writeFile(newWb, outPath);
console.log('\nSaved: BOQ_vs_Model_Report.xlsx (' + newWb.SheetNames.length + ' sheets)');
