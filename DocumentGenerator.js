/**
 * DocumentGenerator.gs
 * Core generation: duplicate template -> replace placeholders -> export -> log.
 */

/**
 * @param {string} templateName
 * @param {string} rowMode - 'selected' | 'all'
 * @param {number[]} selectedRowNumbers - 1-indexed sheet rows (only used if rowMode === 'selected')
 * @param {string} outputFormat - 'docx' | 'pdf' | 'both'
 * @param {string} nameTemplate - e.g. "Employment Contract - {{EmployeeName}}"
 */
function generateDocuments(
  templateName,
  rowMode,
  selectedRowNumbers,
  outputFormat,
  nameTemplate,
) {
  if (
    rowMode === "selected" &&
    (!selectedRowNumbers || selectedRowNumbers.length === 0)
  ) {
    throw new Error("No rows were selected.");
  }

  const template = getTemplateByName_(templateName);
  const { headers, rows } = getSheetDataAsObjects_(template["Source Sheet"]);
  const folderId = getFolderIdOrDefault_(template["Output Folder ID"]);

  // If the sheet is shared across templates (has a CONTRACT_TYPE_COLUMN),
  // "All Rows" should only mean "all rows for THIS template" — not every
  // row in the sheet regardless of which template it actually belongs to.
  const hasContractTypeColumn = headers.includes(CONTRACT_TYPE_COLUMN);

  const targetRows =
    rowMode === "all"
      ? hasContractTypeColumn
        ? rows.filter(
            (r) =>
              String(r[CONTRACT_TYPE_COLUMN] || "").trim() === templateName,
          )
        : rows
      : rows.filter((r) => selectedRowNumbers.includes(r.__row));

  if (targetRows.length === 0) {
    throw new Error(
      rowMode === "all" && hasContractTypeColumn
        ? 'No rows found where "' +
            CONTRACT_TYPE_COLUMN +
            '" is "' +
            templateName +
            '" in "' +
            template["Source Sheet"] +
            '".'
        : "No matching rows found. They may have been deleted or the sheet changed since selection.",
    );
  }

  const results = [];
  targetRows.forEach((row) => {
    try {
      // For Selected Row(s) / Filter by Column, the rows weren't already
      // narrowed to this template above — catch a mismatch here so a
      // manual mistake doesn't silently generate the wrong template.
      if (rowMode !== "all" && hasContractTypeColumn) {
        const rowContractType = String(row[CONTRACT_TYPE_COLUMN] || "").trim();
        if (rowContractType && rowContractType !== templateName) {
          throw new Error(
            "Row " +
              row.__row +
              ' is marked "' +
              rowContractType +
              '" in ' +
              CONTRACT_TYPE_COLUMN +
              ', not "' +
              templateName +
              '". Skipped to avoid generating the wrong template — use Auto (by Contract Type) mode if rows should each use their own template.',
          );
        }
      }

      const result = generateSingleDocument_(
        template,
        row,
        folderId,
        outputFormat,
        nameTemplate,
      );
      writeStatus_(
        template["Source Sheet"],
        row.__row,
        "Generate Status",
        "Generated",
        result.docUrl || result.pdfUrl,
      );
      appendLog_(templateName, result.fileName, "Success", "");
      results.push({
        row: row.__row,
        status: "success",
        fileName: result.fileName,
        docUrl: result.docUrl,
        pdfUrl: result.pdfUrl,
      });
    } catch (err) {
      writeStatus_(
        template["Source Sheet"],
        row.__row,
        "Generate Status",
        "Error",
      );
      appendLog_(templateName, "(row " + row.__row + ")", "Error", err.message);
      results.push({ row: row.__row, status: "error", message: err.message });
    }
  });
  return results;
}

// Column in the source sheet whose value is expected to exactly match a
// "Template Name" in the Template Registry. Used by generateDocumentsAuto.
const CONTRACT_TYPE_COLUMN = "CONTRACT TYPE";

/**
 * Auto mode: for a shared sheet (e.g. Employees) used by several templates,
 * resolve each row's template individually from its CONTRACT_TYPE_COLUMN
 * value instead of applying one template to the whole batch.
 * @param {string} sourceSheetName
 * @param {string} rowMode - 'selected' | 'all'
 * @param {number[]} selectedRowNumbers
 * @param {string} outputFormat - 'docx' | 'pdf' | 'both'
 * @param {string} nameTemplate
 */
function generateDocumentsAuto(
  sourceSheetName,
  rowMode,
  selectedRowNumbers,
  outputFormat,
  nameTemplate,
) {
  if (
    rowMode === "selected" &&
    (!selectedRowNumbers || selectedRowNumbers.length === 0)
  ) {
    throw new Error("No rows were selected.");
  }

  const { headers, rows } = getSheetDataAsObjects_(sourceSheetName);
  if (!headers.includes(CONTRACT_TYPE_COLUMN)) {
    throw new Error(
      'Auto mode requires a "' +
        CONTRACT_TYPE_COLUMN +
        '" column in "' +
        sourceSheetName +
        '", which was not found.',
    );
  }

  const targetRows =
    rowMode === "all"
      ? rows
      : rows.filter((r) => selectedRowNumbers.includes(r.__row));
  if (targetRows.length === 0)
    throw new Error(
      "No matching rows found. They may have been deleted or the sheet changed since selection.",
    );

  const results = [];
  targetRows.forEach((row) => {
    const contractType = String(row[CONTRACT_TYPE_COLUMN] || "").trim();
    if (!contractType) {
      writeStatus_(sourceSheetName, row.__row, "Generate Status", "Error");
      appendLog_(
        "(auto)",
        "(row " + row.__row + ")",
        "Error",
        '"' + CONTRACT_TYPE_COLUMN + '" is blank for this row.',
      );
      results.push({
        row: row.__row,
        status: "error",
        message: '"' + CONTRACT_TYPE_COLUMN + '" is blank for this row.',
      });
      return;
    }

    const template = findActiveTemplateByName_(contractType);
    if (!template) {
      const message =
        'No active template named "' +
        contractType +
        '" (from "' +
        CONTRACT_TYPE_COLUMN +
        '") is registered. Check the Template Registry sheet for a typo or an inactive row.';
      writeStatus_(sourceSheetName, row.__row, "Generate Status", "Error");
      appendLog_("(auto)", "(row " + row.__row + ")", "Error", message);
      results.push({ row: row.__row, status: "error", message: message });
      return;
    }
    if (template["Source Sheet"] !== sourceSheetName) {
      const message =
        'Template "' +
        contractType +
        '" is registered against a different sheet ("' +
        template["Source Sheet"] +
        '"), not "' +
        sourceSheetName +
        '".';
      writeStatus_(sourceSheetName, row.__row, "Generate Status", "Error");
      appendLog_("(auto)", "(row " + row.__row + ")", "Error", message);
      results.push({ row: row.__row, status: "error", message: message });
      return;
    }

    try {
      const folderId = getFolderIdOrDefault_(template["Output Folder ID"]);
      const result = generateSingleDocument_(
        template,
        row,
        folderId,
        outputFormat,
        nameTemplate,
      );
      writeStatus_(
        sourceSheetName,
        row.__row,
        "Generate Status",
        "Generated",
        result.docUrl || result.pdfUrl,
      );
      appendLog_(contractType, result.fileName, "Success", "");
      results.push({
        row: row.__row,
        status: "success",
        fileName: result.fileName,
        templateName: contractType,
        docUrl: result.docUrl,
        pdfUrl: result.pdfUrl,
      });
    } catch (err) {
      writeStatus_(sourceSheetName, row.__row, "Generate Status", "Error");
      appendLog_(contractType, "(row " + row.__row + ")", "Error", err.message);
      results.push({ row: row.__row, status: "error", message: err.message });
    }
  });
  return results;
}

function generateSingleDocument_(
  template,
  rowData,
  folderId,
  outputFormat,
  nameTemplate,
) {
  const derivedDates = deriveDateFields_(rowData);
  const derivedClauses = deriveClauseFields_(rowData);
  const mergedData = Object.assign(
    {},
    getAllSettings_(),
    rowData,
    derivedDates,
    derivedClauses,
  );

  const fileName = sanitizeFileName_(
    fillPlaceholders_(nameTemplate, mergedData),
  );

  const newDocId = duplicateDocToFolder_(
    template["Google Doc ID"],
    fileName,
    folderId,
  );

  const doc = DocumentApp.openById(newDocId);
  replaceAllPlaceholdersInDoc_(doc, mergedData);
  doc.saveAndClose();

  const result = { fileName: fileName, docUrl: null, pdfUrl: null };

  if (outputFormat === "pdf" || outputFormat === "both") {
    const pdfId = convertDocToPdf_(newDocId, folderId, fileName);
    result.pdfUrl = DriveApp.getFileById(pdfId).getUrl();
  }
  if (outputFormat === "docx") {
    result.docUrl = DriveApp.getFileById(newDocId).getUrl();
  } else if (outputFormat === "pdf") {
    DriveApp.getFileById(newDocId).setTrashed(true);
  } else {
    result.docUrl = DriveApp.getFileById(newDocId).getUrl();
  }

  return result;
}

// Maps a source-sheet "option" column to the Clause Key it resolves against
// in the Clause Library sheet. Add a line here whenever a new {{OPTION}}
// column + clause key pair is introduced.
const CLAUSE_OPTION_COLUMNS = {
  "ACCRUAL OPTION": "ACCRUAL",
  "CARRYOVER OPTION": "CARRYOVER",
  "LEAVE USAGE OPTION": "LEAVE_USAGE",
  "HAS HMO": "HMO",
  "HMO EFFECTIVE OPTION": "HMO_EFFECTIVE",
  "HMO COVERAGE OPTION": "HMO_COVERAGE",
  "BUSINESS TOOLS OPTION": "BUSINESS_TOOLS",
};

// Cached for the lifetime of one execution (i.e. one Generate click, however
// many rows it covers) so the 989-row Clause Library isn't re-fetched per row.
let clauseLibraryCache_ = null;
function getClauseLibrary_() {
  if (clauseLibraryCache_ === null) {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Clause Library");
    // No Clause Library sheet in this spreadsheet -> clause resolution is a
    // silent no-op rather than an error, so setups without one still work.
    clauseLibraryCache_ = sheet
      ? getSheetDataAsObjects_("Clause Library").rows
      : [];
  }
  return clauseLibraryCache_;
}

// Looks up Clause Key + Option in the Clause Library and returns its Text,
// with any placeholders inside that text (e.g. "{{CARRYOVER MAX DAYS}}")
// filled from the same row's data. Returns null if there's no matching
// option row, so the caller can leave that placeholder untouched.
function resolveClauseText_(clauseKey, optionValue, rowData) {
  if (optionValue === null || optionValue === undefined || optionValue === "")
    return null;
  const optionStr = String(optionValue).trim();
  const match = getClauseLibrary_().find(
    (r) =>
      String(r["Clause Key"]).trim() === clauseKey &&
      String(r["Option"]).trim() === optionStr,
  );
  if (!match) return null;
  return fillPlaceholders_(String(match["Text"] || ""), rowData);
}

function deriveClauseFields_(rowData) {
  const derived = {};
  Object.keys(CLAUSE_OPTION_COLUMNS).forEach((column) => {
    const clauseKey = CLAUSE_OPTION_COLUMNS[column];
    const text = resolveClauseText_(clauseKey, rowData[column], rowData);
    if (text !== null) derived[clauseKey] = text;
  });
  return derived;
}

function deriveDateFields_(rowData) {
  const raw = rowData["CREATE DATE"];
  if (!raw) return {};
  const dt = new Date(raw);
  if (isNaN(dt.getTime())) return {};
  return {
    day: dt.getDate(),
    month: Utilities.formatDate(dt, Session.getScriptTimeZone(), "MMMM"),
  };
}

function replaceAllPlaceholdersInDoc_(doc, rowData) {
  const body = doc.getBody();
  const header = doc.getHeader();
  const footer = doc.getFooter();

  Object.keys(rowData).forEach((key) => {
    if (key === "__row") return;
    const placeholder = "{{" + key + "}}";
    let value =
      rowData[key] === null || rowData[key] === undefined
        ? ""
        : String(rowData[key]);
    value = value.replace(/\n/g, "\v");
    body.replaceText(escapeRegex_(placeholder), value);
    if (header) header.replaceText(escapeRegex_(placeholder), value);
    if (footer) footer.replaceText(escapeRegex_(placeholder), value);
  });
}

// Used for filenames, e.g. "Employment Contract - {{EmployeeName}}"
function fillPlaceholders_(templateString, rowData) {
  return templateString.replace(/{{\s*([^{}]+?)\s*}}/g, (match, rawKey) => {
    const key = rawKey.trim();
    return rowData[key] !== undefined ? String(rowData[key]) : match;
  });
}

function escapeRegex_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
