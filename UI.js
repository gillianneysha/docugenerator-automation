/**
 * UI.gs
 * Functions that show sidebars and bridge HTML <-> server logic.
 */

function showRegisterTemplateSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('HTML/RegisterTemplate')
    .setTitle('Register Template');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showGenerateSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('HTML/Sidebar')
    .setTitle('Generate Documents');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showManageTemplatesSidebar() {
  // MVP: reuse the register sidebar's list view, or build out later.
  SpreadsheetApp.getUi().alert('Manage Templates: list is in the Template Registry sheet for now. Set Status to "Inactive" to disable a template.');
}

// --- called from HTML via google.script.run ---

function getSheetNames_UI() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

function getFolderPicker_UI() {
  // MVP: user pastes a folder ID/URL manually in the sidebar instead of a full Drive picker.
  return true;
}

function registerTemplate_UI(templateName, sourceSheet, outputFolderIdOrUrl) {
  const docId = DocumentApp.getActiveDocument
    ? null
    : null; // placeholder - actual doc comes from the sidebar's "current doc" input
  return null; // see RegisterTemplate.html: it collects the doc URL and parses the ID itself
}

function extractIdFromUrl_(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : url;
}

function registerTemplateFromForm_UI(templateName, docUrl, sourceSheet, folderUrl) {
  const docId = extractIdFromUrl_(docUrl);
  const folderId = extractIdFromUrl_(folderUrl);
  return registerTemplate(templateName, docId, sourceSheet, folderId);
}

function generateDocuments_UI(templateName, rowMode, outputFormat, nameTemplate) {
  let selectedRows = [];
  if (rowMode === 'selected') {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const selection = sheet.getActiveRange();
    selectedRows = [];
    for (let r = selection.getRow(); r < selection.getRow() + selection.getNumRows(); r++) {
      selectedRows.push(r);
    }
  }
  return generateDocuments(templateName, rowMode, selectedRows, outputFormat, nameTemplate);
}
