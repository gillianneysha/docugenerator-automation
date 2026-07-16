/**
 * Code.gs
 * Entry point: menu creation only. No business logic here.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Document Automation')
    .addItem('Generate Documents', 'showGenerateSidebar')
    .addItem('Register Template', 'showRegisterTemplateSidebar')
    .addItem('Manage Templates', 'showManageTemplatesSidebar')
    .addSeparator()
    .addItem('Initialize Sheets (first-time setup)', 'initializeSheets')
    .addToUi();
}

/**
 * Run this ONCE on a fresh spreadsheet to create the required tabs
 * (Template Registry, Settings, Logs) if they don't exist yet.
 * Safe to re-run; it won't overwrite existing sheets.
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheetIfMissing_(ss, 'Template Registry',
    ['Template Name', 'Google Doc ID', 'Source Sheet', 'Output Folder ID', 'Status']);

  createSheetIfMissing_(ss, 'Settings',
    ['Setting', 'Value']);

  createSheetIfMissing_(ss, 'Logs',
    ['Date', 'User', 'Template', 'Document Name', 'Status', 'Notes']);

  SpreadsheetApp.getUi().alert('Setup complete. Fill in the Settings sheet (Default Output Folder, Default PDF Folder, Company Name), then register your first template.');
}

function createSheetIfMissing_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
