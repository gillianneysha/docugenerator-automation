function showRegisterTemplateSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile("HTML/RegisterTemplate").setTitle(
      "Register Template",
    ),
  );
}
function showGenerateSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile("HTML/Sidebar").setTitle(
      "Generate Documents",
    ),
  );
}
function showManageTemplatesSidebar() {
  SpreadsheetApp.getUi().alert(
    'Manage Templates: edit rows directly in the Template Registry sheet. Set Status to "Inactive" to disable.',
  );
}
function showLogsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Logs");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Logs sheet not found.");
    return;
  }
  ss.setActiveSheet(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.setActiveRange(sheet.getRange(lastRow, 1));
}
function showSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Settings sheet not found.");
    return;
  }
  ss.setActiveSheet(sheet);
}
function showRevalidateResults() {
  const results = revalidateAllTemplates_UI();
  if (results.length === 0) {
    SpreadsheetApp.getUi().alert("No templates registered yet.");
    return;
  }
  const lines = results.map(
    (r) =>
      (r.ok ? "✅ " : "⚠️ ") +
      r.templateName +
      " (" +
      r.status +
      "): " +
      r.message,
  );
  SpreadsheetApp.getUi().alert(
    "Template Validation Results\n\n" + lines.join("\n"),
  );
}

function getSheetNames_UI() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map((s) => s.getName());
}
function extractIdFromUrl_(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : url;
}
function registerTemplateFromForm_UI(
  templateName,
  docUrl,
  sourceSheet,
  folderUrl,
) {
  return registerTemplate(
    templateName,
    extractIdFromUrl_(docUrl),
    sourceSheet,
    extractIdFromUrl_(folderUrl),
  );
}
function getTemplateRows_UI(templateName) {
  const template = getTemplateByName_(templateName);
  const { headers, rows } = getSheetDataAsObjects_(template["Source Sheet"]);
  const labelHeaders = headers.slice(0, 3);
  return rows.map((r) => ({
    row: r.__row,
    label: labelHeaders
      .map((h) => r[h])
      .filter((v) => v !== "" && v !== null && v !== undefined)
      .join(" — "),
  }));
}
function getTemplateHeaders_UI(templateName) {
  const template = getTemplateByName_(templateName);
  return getSheetDataAsObjects_(template["Source Sheet"]).headers;
}
function getColumnValues_UI(templateName, column) {
  const template = getTemplateByName_(templateName);
  const values = getSheetDataAsObjects_(template["Source Sheet"])
    .rows.map((r) => r[column])
    .filter((v) => v !== "" && v !== null && v !== undefined);
  return [...new Set(values)].sort();
}
function getFilteredRows_UI(templateName, column, value) {
  const template = getTemplateByName_(templateName);
  return getSheetDataAsObjects_(template["Source Sheet"])
    .rows.filter((r) => String(r[column]) === String(value))
    .map((r) => r.__row);
}
function generateDocuments_UI(
  templateName,
  rowMode,
  selectedRows,
  outputFormat,
  nameTemplate,
) {
  return generateDocuments(
    templateName,
    rowMode,
    selectedRows || [],
    outputFormat,
    nameTemplate,
  );
}
