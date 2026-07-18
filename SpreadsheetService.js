/**
 * SpreadsheetService.gs
 * All reading/writing of sheet data lives here.
 */

function getSheetDataAsObjects_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    // skip fully blank rows
    if (values[i].every(cell => cell === '' || cell === null)) continue;
    const obj = { __row: i + 1 }; // 1-indexed sheet row, for status write-back
    headers.forEach((h, idx) => obj[h] = values[i][idx]);
    rows.push(obj);
  }
  return { headers, rows };
}

function getTemplateRegistry_() {
  const { rows } = getSheetDataAsObjects_('Template Registry');
  return rows.filter(r => String(r['Status']).toLowerCase() === 'active');
}

function getTemplateByName_(name) {
  const registry = getTemplateRegistry_();
  const match = registry.find(r => r['Template Name'] === name);
  if (!match) throw new Error('Template not found or inactive: ' + name);
  return match;
}

function addTemplateToRegistry_(templateName, docId, sourceSheet, outputFolderId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Template Registry');
  sheet.appendRow([templateName, docId, sourceSheet, outputFolderId, 'Active']);
}

function getSetting_(key) {
  const { rows } = getSheetDataAsObjects_('Settings');
  const row = rows.find(r => r['Setting'] === key);
  return row ? row['Value'] : null;
}

function writeStatus_(sourceSheet, rowNumber, columnName, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let colIndex = headers.indexOf(columnName);
  if (colIndex === -1) {
    colIndex = headers.length;
    sheet.getRange(1, colIndex + 1).setValue(columnName);
  }
  sheet.getRange(rowNumber, colIndex + 1).setValue(value);
}

function appendLog_(templateName, docName, status, notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
  sheet.appendRow([
    new Date(),
    Session.getActiveUser().getEmail(),
    templateName,
    docName,
    status,
    notes || ''
  ]);
}
