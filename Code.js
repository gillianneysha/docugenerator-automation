function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Document Automation")
    .addItem("Generate Documents", "showGenerateSidebar")
    .addItem("Register Template", "showRegisterTemplateSidebar")
    .addItem("Manage Templates", "showManageTemplatesSidebar")
    .addItem("Re-validate Templates", "showRevalidateResults")
    .addSeparator()
    .addItem("View Logs", "showLogsSheet")
    .addItem("Archive Old Logs", "archiveOldLogs")
    .addItem("Settings", "showSettingsSheet")
    .addSeparator()
    .addItem("Initialize Sheets (first-time setup)", "initializeSheets")
    .addToUi();
}

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createSheetIfMissing_(ss, "Template Registry", [
    "Template Name",
    "Google Doc ID",
    "Source Sheet",
    "Output Folder ID",
    "Status",
  ]);
  createSheetIfMissing_(ss, "Settings", ["Setting", "Value"]);
  createSheetIfMissing_(ss, "Logs", [
    "Date",
    "User",
    "Template",
    "Document Name",
    "Status",
    "Notes",
  ]);
  createSheetIfMissing_(ss, "Logs Archive", [
    "Date",
    "User",
    "Template",
    "Document Name",
    "Status",
    "Notes",
  ]);
  createSheetIfMissing_(ss, "Clause Library", ["Clause Key", "Option", "Text"]);

  const hasDefaultFolder = getSetting_("Default Output Folder");
  if (!hasDefaultFolder) {
    const folders = initializeDriveFolders_();
    setSetting_("Default Output Folder", folders.generatedDocsId);
    setSetting_("Default PDF Folder", folders.generatedPdfsId);
    if (!getSetting_("Company Name")) setSetting_("Company Name", "");
    SpreadsheetApp.getUi().alert(
      "Setup complete.\n\nCreated Drive folder structure at: " +
        folders.rootUrl +
        "\n\nDefault Output Folder and Default PDF Folder have been set automatically. " +
        'Fill in "Company Name" in Settings, then register your first template.',
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "Sheets are already set up. Drive folders were not recreated.",
    );
  }
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
