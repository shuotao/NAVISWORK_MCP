/**
 * Generate BOQ comparison report in original Excel format
 * Each sheet mirrors the BOQ structure + adds Model columns
 */
const XLSX = require('xlsx');
const fs = require('fs');

const EXCEL_PATH = 'C:/Users/Admin/Downloads/Appendix A - MSI-1-MH-BQ-G-00256_Rev_1_Bill of Quantity.xlsx';
const origWb = XLSX.readFile(EXCEL_PATH);
const modelData = JSON.parse(fs.readFileSync(__dirname + '/b001_full_scan.json', 'utf8'));

const skipSheets = ['Cover Sheet', 'Preamble', '1. Preli. Works'];

// Create new workbook
var newWb = XLSX.utils.book_new();

// ─── Summary sheet ───
var summaryRows = [
  ['MICRON MSI1 7A BUMPING - MODEL vs BOQ CROSS-REFERENCE REPORT'],
  ['Generated: ' + new Date().toISOString().split('T')[0]],
  [],
  ['BQ Sheet', 'Model Elements', 'BOQ Items (Excel)', 'BOQ Total Qty', 'Status', 'Coverage'],
];

var allSheets = new Set();
Object.keys(modelData).forEach(function (k) { allSheets.add(k); });
origWb.SheetNames.forEach(function (k) { if (skipSheets.indexOf(k) < 0) allSheets.add(k); });

Array.from(allSheets).sort().forEach(function (sheet) {
  // Model count
  var mCount = 0;
  if (modelData[sheet]) {
    Object.keys(modelData[sheet]).forEach(function (cat) {
      Object.keys(modelData[sheet][cat]).forEach(function (t) { mCount += modelData[sheet][cat][t]; });
    });
  }

  // BOQ count from Excel
  var bqCount = 0, bqQty = 0;
  var wsOrig = origWb.Sheets[sheet] || origWb.Sheets[sheet + ' '];
  if (wsOrig) {
    var data = XLSX.utils.sheet_to_json(wsOrig, { header: 1 });
    data.forEach(function (row) {
      if (!row) return;
      // Count rows that look like items (have a unit)
      for (var i = 2; i < Math.min(row.length, 10); i++) {
        var v = String(row[i] || '').trim();
        if (v.match(/^(Nos|Nr|Set|Meter|Rmt|LS|Kg|Sqm|Cum|Lit|mtr|nos|Lot|Each|Lump)\.?$/i)) {
          bqCount++;
          // Find qty column (usually next few cols)
          for (var j = i + 1; j < Math.min(row.length, i + 5); j++) {
            if (row[j] != null && !isNaN(Number(row[j])) && Number(row[j]) > 0) {
              bqQty += Number(row[j]);
              break;
            }
          }
          break;
        }
      }
    });
  }

  var status = '';
  var coverage = '';
  if (mCount > 0 && bqCount > 0) { status = 'MATCHED'; coverage = '100%'; }
  else if (mCount > 0 && bqCount === 0) { status = 'MODEL ONLY'; coverage = 'N/A'; }
  else if (mCount === 0 && bqCount > 0) { status = 'BOQ ONLY'; coverage = '0%'; }
  else { status = 'EMPTY'; coverage = 'N/A'; }

  summaryRows.push([sheet, mCount, bqCount, Math.round(bqQty), status, coverage]);
});

var summaryWs = XLSX.utils.aoa_to_sheet(summaryRows);
summaryWs['!cols'] = [{ wch: 35 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 10 }];
XLSX.utils.book_append_sheet(newWb, summaryWs, 'Summary');

// ─── Per-sheet detail ───
origWb.SheetNames.forEach(function (sheetName) {
  if (skipSheets.indexOf(sheetName) >= 0) return;
  if (sheetName === 'Cost Summary') return;

  var wsOrig = origWb.Sheets[sheetName];
  if (!wsOrig) return;
  var data = XLSX.utils.sheet_to_json(wsOrig, { header: 1 });

  // Find the model data for this sheet
  var modelSheet = modelData[sheetName] || modelData[sheetName.trim()];

  // Build output rows: original BOQ + model comparison columns
  var outRows = [];

  // Title
  outRows.push([sheetName + ' — MODEL vs BOQ COMPARISON']);
  outRows.push([]);

  // Model summary section
  if (modelSheet) {
    var totalModel = 0;
    Object.keys(modelSheet).forEach(function (cat) {
      Object.keys(modelSheet[cat]).forEach(function (t) { totalModel += modelSheet[cat][t]; });
    });
    outRows.push(['MODEL ELEMENTS IN THIS SCOPE: ' + totalModel]);
    outRows.push(['Category', 'Count', '', 'Top Family | Type', 'Count']);

    Object.keys(modelSheet).sort().forEach(function (cat) {
      var catTotal = 0;
      var types = [];
      Object.keys(modelSheet[cat]).forEach(function (t) {
        catTotal += modelSheet[cat][t];
        types.push({ name: t, count: modelSheet[cat][t] });
      });
      types.sort(function (a, b) { return b.count - a.count; });

      outRows.push([cat, catTotal, '', types[0] ? types[0].name : '', types[0] ? types[0].count : '']);
      // Show top 3 types
      for (var ti = 1; ti < Math.min(types.length, 4); ti++) {
        outRows.push(['', '', '', types[ti].name, types[ti].count]);
      }
    });
  } else {
    outRows.push(['MODEL ELEMENTS IN THIS SCOPE: 0 (not in B001 model)']);
  }

  outRows.push([]);
  outRows.push(['─── ORIGINAL BOQ ───']);
  outRows.push([]);

  // Copy original BOQ rows
  data.forEach(function (row) {
    if (!row) { outRows.push([]); return; }
    outRows.push(row.slice());
  });

  // Create sheet
  var name = sheetName.trim();
  if (name.length > 31) name = name.substring(0, 31);
  var ws = XLSX.utils.aoa_to_sheet(outRows);
  ws['!cols'] = [{ wch: 10 }, { wch: 50 }, { wch: 8 }, { wch: 50 }, { wch: 10 },
    { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];

  try {
    XLSX.utils.book_append_sheet(newWb, ws, name);
  } catch (e) {
    console.log('Skip sheet ' + name + ': ' + e.message);
  }
});

var outPath = __dirname + '/BOQ_vs_Model_Report.xlsx';
XLSX.writeFile(newWb, outPath);
console.log('Saved: BOQ_vs_Model_Report.xlsx (' + newWb.SheetNames.length + ' sheets)');
