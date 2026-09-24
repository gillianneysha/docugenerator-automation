/**
 * SpreadsheetService.gs
 * All reading/writing of sheet data lives here.
 */

const TEMPLATE_REGISTRY_HEADERS_ = [
  "Template Name",
  "File Name Pattern",
  "Google Doc ID",
  "Source Sheet",
  "Output Folder ID",
  "Status",
];
const DEFAULT_FILE_NAME_PATTERN_ = "{{CLIENT NAME}} - {{EMPLOYEE NAME}}";

function ensureTemplateRegistrySchema_(ss) {
  let sheet = ss.getSheetByName("Template Registry");
  if (!sheet) {
    createSheetIfMissing_(ss, "Template Registry", TEMPLATE_REGISTRY_HEADERS_);
    return;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const values = sheet.getRange(1, 1, sheet.getLastRow(), lastColumn).getDisplayValues();
  const headers = values[0].map((header) => String(header).trim());
  const rows = values.slice(1).filter((row) => row.some((cell) => cell !== ""));
  const headerIndexes = {};
  headers.forEach((header, index) => {
    if (header) headerIndexes[header] = index;
  });

  const normalizedRows = rows.map((row) => {
    const getValue = (header) =>
      headerIndexes[header] === undefined ? "" : row[headerIndexes[header]] || "";
    let pattern = getValue("File Name Pattern");
    let docId = getValue("Google Doc ID");
    let sourceSheet = getValue("Source Sheet");
    let outputFolderId = getValue("Output Folder ID");
    let status = getValue("Status");

    // Older five-column writes landed one position too early after the new
    // pattern column was added. Repair only the recognizable shifted shape.
    const looksShifted =
      !status &&
      outputFolderId.toLowerCase() === "active" &&
      /^[\w-]{25,}$/.test(pattern);
    if (looksShifted) {
      const legacyDocId = pattern;
      const legacySourceSheet = docId;
      const legacyOutputFolderId = sourceSheet;
      const legacyStatus = outputFolderId;
      pattern = DEFAULT_FILE_NAME_PATTERN_;
      docId = legacyDocId;
      sourceSheet = legacySourceSheet;
      outputFolderId = legacyOutputFolderId;
      status = legacyStatus;
    }

    return [
      getValue("Template Name"),
      pattern || DEFAULT_FILE_NAME_PATTERN_,
      docId,
      sourceSheet,
      outputFolderId,
      status || "Active",
    ];
  });

  const schemaMatches =
    headers.length === TEMPLATE_REGISTRY_HEADERS_.length &&
    TEMPLATE_REGISTRY_HEADERS_.every((header, index) => headers[index] === header);
  if (schemaMatches && normalizedRows.every((row, index) =>
    row.every((value, column) => value === values[index + 1][column]),
  )) return;

  sheet.clearContents();
  sheet
    .getRange(1, 1, 1 + normalizedRows.length, TEMPLATE_REGISTRY_HEADERS_.length)
    .setValues([TEMPLATE_REGISTRY_HEADERS_, ...normalizedRows]);
  sheet.setFrozenRows(1);
}

function getSheetDataAsObjects_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet not found: " + sheetName);

  // getDisplayValues() = exactly what's shown in the cell (commas, currency,
  // date formats) instead of the raw value. Formatting stays controlled
  // entirely from the sheet — no code changes needed if HR changes a format.
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map((h) => String(h).trim());
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    if (values[i].every((cell) => cell === "")) continue;
    const obj = { __row: i + 1 };
    headers.forEach((h, idx) => (obj[h] = values[i][idx]));
    rows.push(obj);
  }
  return { headers, rows };
}

function getTemplateRegistry_() {
  const { rows } = getSheetDataAsObjects_("Template Registry");
  return rows.filter((r) => String(r["Status"]).toLowerCase() === "active");
}

function getTemplateByName_(name) {
  const registry = getTemplateRegistry_();
  const match = registry.find((r) => r["Template Name"] === name);
  if (!match) throw new Error("Template not found or inactive: " + name);
  return match;
}

function addTemplateToRegistry_(
  templateName,
  fileNamePattern,
  docId,
  sourceSheet,
  outputFolderId,
) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template Registry");
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0]
    .map((header) => String(header).trim());
  const row = Array(headers.length).fill("");
  const values = {
    "Template Name": templateName,
    "File Name Pattern": fileNamePattern,
    "Google Doc ID": docId,
    "Source Sheet": sourceSheet,
    "Output Folder ID": outputFolderId,
    Status: "Active",
  };
  headers.forEach((header, index) => {
    if (Object.prototype.hasOwnProperty.call(values, header)) row[index] = values[header];
  });
  sheet.appendRow(row);
}

function getSetting_(key) {
  const { rows } = getSheetDataAsObjects_("Settings");
  const row = rows.find((r) => r["Setting"] === key);
  return row ? row["Value"] : null;
}

function setSetting_(key, value) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function writeStatus_(sourceSheet, rowNumber, columnName, value, linkUrl) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let colIndex = headers.indexOf(columnName);
  if (colIndex === -1) {
    colIndex = headers.length;
    sheet.getRange(1, colIndex + 1).setValue(columnName);
  }
  const cell = sheet.getRange(rowNumber, colIndex + 1);
  if (linkUrl) {
    const richValue = SpreadsheetApp.newRichTextValue()
      .setText(value)
      .setLinkUrl(linkUrl)
      .build();
    cell.setRichTextValue(richValue);
  } else {
    cell.setValue(value);
  }
}

function appendLog_(templateName, docName, status, notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  sheet.appendRow([
    new Date(),
    Session.getActiveUser().getEmail(),
    templateName,
    docName,
    status,
    notes || "",
  ]);
}

function getTemplateRows_UI(templateName) {
  const template = getTemplateByName_(templateName);
  const { headers, rows } = getSheetDataAsObjects_(template["Source Sheet"]);

  // Use up to the first 3 columns as a readable label, e.g. "E-1002 — Juan Dela Cruz — Sales"
  const labelHeaders = headers.slice(0, 3);
  return rows.map((r) => ({
    row: r.__row,
    label: labelHeaders
      .map((h) => r[h])
      .filter((v) => v !== "" && v !== null && v !== undefined)
      .join(" — "),
  }));
}

function getAllTemplateRows_() {
  return getSheetDataAsObjects_("Template Registry").rows;
}

function getTemplateHeaders_UI(templateName) {
  const template = getTemplateByName_(templateName);
  const { headers } = getSheetDataAsObjects_(template["Source Sheet"]);
  return headers;
}

function getColumnValues_UI(templateName, column) {
  const template = getTemplateByName_(templateName);
  const { rows } = getSheetDataAsObjects_(template["Source Sheet"]);
  const values = rows
    .map((r) => r[column])
    .filter((v) => v !== "" && v !== null && v !== undefined)
    .map((v) =>
      v instanceof Date
        ? Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(v),
    );
  return [...new Set(values)].sort();
}

function getFilteredRows_UI(templateName, column, value) {
  const template = getTemplateByName_(templateName);
  const { rows } = getSheetDataAsObjects_(template["Source Sheet"]);
  return rows
    .filter((r) => {
      const cell = r[column];
      const cellStr =
        cell instanceof Date
          ? Utilities.formatDate(
              cell,
              Session.getScriptTimeZone(),
              "yyyy-MM-dd",
            )
          : String(cell);
      return cellStr === value;
    })
    .map((r) => r.__row);
}

function getAllSettings_() {
  const { rows } = getSheetDataAsObjects_("Settings");
  const map = {};
  rows.forEach((r) => {
    if (r["Setting"]) map[r["Setting"]] = r["Value"];
  });
  return map;
}

/** Moves Logs rows older than `monthsOld` months into "Logs Archive". */
function archiveOldLogs(monthsOld) {
  monthsOld = monthsOld || 6;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logs = ss.getSheetByName("Logs");
  const archive =
    ss.getSheetByName("Logs Archive") ||
    createSheetIfMissing_(ss, "Logs Archive", [
      "Date",
      "User",
      "Template",
      "Document Name",
      "Status",
      "Notes",
    ]);

  const data = logs.getDataRange().getValues();
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - monthsOld);

  const toKeep = [data[0]];
  const toArchive = [];
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0];
    if (rowDate instanceof Date && rowDate < cutoff) toArchive.push(data[i]);
    else toKeep.push(data[i]);
  }

  if (toArchive.length === 0) {
    SpreadsheetApp.getUi().alert(
      "No log entries older than " + monthsOld + " months. Nothing to archive.",
    );
    return;
  }

  archive
    .getRange(archive.getLastRow() + 1, 1, toArchive.length, toArchive[0].length)
    .setValues(toArchive);
  logs.clearContents();
  logs.getRange(1, 1, toKeep.length, toKeep[0].length).setValues(toKeep);
  logs.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(
    "Archived " + toArchive.length + " log entries to \"Logs Archive\".",
  );
}
