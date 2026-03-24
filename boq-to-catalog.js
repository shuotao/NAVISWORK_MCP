/**
 * BOQ Excel → Navisworks Quantification Catalog XML
 * Reads all BOQ sheets, extracts items, outputs Project_BoQ_Catalog.xml
 */
const XLSX = require('xlsx');
const fs = require('fs');

const EXCEL_PATH = 'C:/Users/Admin/Downloads/Appendix A - MSI-1-MH-BQ-G-00256_Rev_1_Bill of Quantity.xlsx';
const wb = XLSX.readFile(EXCEL_PATH);

// Skip non-BOQ sheets
const skipSheets = ['Cover Sheet', 'Cost Summary', 'Preamble', '1. Preli. Works'];

function escXml(s) {
  if (!s) return '';
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function findHeaderRow(data) {
  for (var i = 0; i < Math.min(data.length, 15); i++) {
    var row = data[i];
    if (!row) continue;
    var joined = row.map(function(c) { return String(c || '').toUpperCase().trim(); }).join(' ');
    // Match various S.NO patterns + DESCRIPTION or ITEM
    var hasSno = joined.match(/S\.?\s*NO|SI\.?\s*NO|SR\.?\s*NO/);
    var hasDesc = joined.indexOf('DESCRIPTION') >= 0 || joined.indexOf('ITEM') >= 0;
    if (hasSno && hasDesc) return i;
  }
  return -1;
}

function findColumns(headerRow) {
  var cols = { sno: -1, desc: -1, unit: -1, qty: -1 };
  headerRow.forEach(function(cell, i) {
    var v = String(cell || '').toUpperCase().trim();
    if (v.match(/^S\.?NO|^SI\.?NO|^SR/)) cols.sno = i;
    if (v.match(/DESCRIPTION/) || v === 'ITEM') cols.desc = i;
    if (v.match(/^UNITS?$/)) cols.unit = i;
    if (v.match(/^(QUANTITY|QTY\.?|TOTAL QTY)$/)) cols.qty = i;
  });
  return cols;
}

var catalog = [];

wb.SheetNames.forEach(function(sheetName) {
  if (skipSheets.indexOf(sheetName) >= 0) return;

  var ws = wb.Sheets[sheetName];
  var data = XLSX.utils.sheet_to_json(ws, { header: 1 });

  var headerIdx = findHeaderRow(data);
  if (headerIdx < 0) {
    console.log('[SKIP] ' + sheetName + ' - no header row found');
    return;
  }

  var cols = findColumns(data[headerIdx]);
  if (cols.desc < 0) {
    console.log('[SKIP] ' + sheetName + ' - no description column');
    return;
  }

  // Extract WBS from sheet name
  var wbs = sheetName.split(' ')[0].replace(/\.$/, '');

  var items = [];
  var currentParent = null;

  for (var i = headerIdx + 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) continue;

    var sno = cols.sno >= 0 ? row[cols.sno] : null;
    var desc = cols.desc >= 0 ? row[cols.desc] : null;
    var unit = cols.unit >= 0 ? row[cols.unit] : null;
    var qty = cols.qty >= 0 ? row[cols.qty] : null;

    if (!desc || String(desc).trim() === '') continue;
    desc = String(desc).trim();

    // Skip general notes and long descriptions (>300 chars)
    if (desc.length > 300) continue;
    // Skip rows that are just notes
    if (desc.match(/^(General Note|Scope of Work|The scope|Handling|Quality Control|BOQ,|Protection|Surrounding|Barricading|Note\s*[-:])/i)) continue;

    var snoStr = sno != null ? String(sno).trim() : '';
    var unitStr = unit ? String(unit).trim() : '';
    var qtyVal = qty != null && !isNaN(Number(qty)) ? Number(qty) : 0;

    // Determine if this is a parent/group item (integer sno, no unit) or a leaf item
    var isGroup = (snoStr.match(/^\d+$/) && !unitStr);

    if (isGroup) {
      currentParent = desc;
    }

    items.push({
      sno: snoStr,
      desc: desc,
      unit: unitStr,
      qty: qtyVal,
      isGroup: isGroup,
      parent: isGroup ? null : currentParent
    });
  }

  if (items.length > 0) {
    catalog.push({ sheet: sheetName, wbs: wbs, items: items });
    console.log('[OK] ' + sheetName + ' : ' + items.length + ' items');
  }
});

// Generate XML
var xml = '<?xml version="1.0" encoding="utf-8"?>\n';
xml += '<Catalog>\n';

catalog.forEach(function(sheet) {
  xml += '  <ItemGroup Name="' + escXml(sheet.sheet) + '" WBS="' + escXml(sheet.wbs) + '">\n';

  var currentGroup = null;
  sheet.items.forEach(function(item) {
    if (item.isGroup) {
      if (currentGroup) xml += '    </SubGroup>\n';
      currentGroup = item.desc;
      xml += '    <SubGroup Name="' + escXml(item.sno + ' - ' + item.desc) + '">\n';
    } else {
      var itemName = item.sno ? (item.sno + ' - ' + item.desc) : item.desc;
      if (itemName.length > 200) itemName = itemName.substring(0, 200);
      xml += '      <Item';
      xml += ' Name="' + escXml(itemName) + '"';
      if (item.unit) xml += ' Unit="' + escXml(item.unit) + '"';
      if (item.qty) xml += ' Quantity="' + item.qty + '"';
      xml += ' />\n';
    }
  });
  if (currentGroup) xml += '    </SubGroup>\n';

  xml += '  </ItemGroup>\n';
});

xml += '</Catalog>\n';

fs.writeFileSync(__dirname + '/Project_BoQ_Catalog.xml', xml, 'utf8');
console.log('\n\nSaved: Project_BoQ_Catalog.xml (' + catalog.length + ' sheets)');

// Also save a CSV summary for reference
var csv = 'Sheet,WBS,S.NO,Description,Unit,Qty\n';
catalog.forEach(function(sheet) {
  sheet.items.forEach(function(item) {
    if (!item.isGroup) {
      csv += '"' + sheet.sheet + '","' + sheet.wbs + '","' + item.sno + '","' +
        item.desc.replace(/"/g, '""') + '","' + item.unit + '",' + item.qty + '\n';
    }
  });
});
fs.writeFileSync(__dirname + '/BoQ_Items.csv', csv, 'utf8');
console.log('Saved: BoQ_Items.csv');
