/**
 * SpreadsheetService.gs
 * All reading/writing of sheet data lives here.
 */

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
  docId,
  sourceSheet,
  outputFolderId,
) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template Registry");
  sheet.appendRow([templateName, docId, sourceSheet, outputFolderId, "Active"]);
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

function writeStatus_(sourceSheet, rowNumber, columnName, value) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sourceSheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let colIndex = headers.indexOf(columnName);
  if (colIndex === -1) {
    colIndex = headers.length;
    sheet.getRange(1, colIndex + 1).setValue(columnName);
  }
  sheet.getRange(rowNumber, colIndex + 1).setValue(value);
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

function getClauseText_(clauseKey, option) {
  const { rows } = getSheetDataAsObjects_("Clause Library");
  const match = rows.find(
    (r) =>
      String(r["Clause Key"]).trim() === String(clauseKey).trim() &&
      String(r["Option"]).trim() === String(option).trim(),
  );
  return match ? match["Text"] : "";
}

function resolveClauseTemplate_(text, rowData) {
  if (!text) return "";
  let resolved = text.replace(/\\n/g, "\n");
  resolved = resolved.replace(/{{\s*([^{}]+?)\s*}}/g, (match, rawKey) => {
    const key = rawKey.trim();
    return rowData[key] !== undefined ? String(rowData[key]) : match;
  });
  return resolved;
}
