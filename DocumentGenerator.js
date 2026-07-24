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
  const { rows } = getSheetDataAsObjects_(template["Source Sheet"]);
  const folderId = getFolderIdOrDefault_(template["Output Folder ID"]);

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
    try {
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

function generateSingleDocument_(
  template,
  rowData,
  folderId,
  outputFormat,
  nameTemplate,
) {
  const derived = deriveDateFields_(rowData);
  const clauses = deriveClauseFields_(rowData);
  const mergedData = Object.assign(
    {},
    getAllSettings_(),
    rowData,
    derived,
    clauses,
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

function deriveClauseFields_(rowData) {
  return {
    LEAVE_ACCRUAL_CLAUSE: resolveClauseTemplate_(
      getClauseText_("ACCRUAL", rowData["ACCRUAL OPTION"]),
      rowData,
    ),
    LEAVE_CARRYOVER_CLAUSE: resolveClauseTemplate_(
      getClauseText_("CARRYOVER", rowData["CARRYOVER OPTION"]),
      rowData,
    ),
    LEAVE_USAGE_CLAUSE: resolveClauseTemplate_(
      getClauseText_("LEAVE_USAGE", rowData["LEAVE USAGE OPTION"]),
      rowData,
    ),
    HMO_CLAUSE: buildHmoClause_(rowData),
    ANNUAL_SALARY: computeAnnualSalary_(rowData),
    BUSINESS_TOOLS_CLAUSE: resolveClauseTemplate_(
      getClauseText_("BUSINESS_TOOLS", rowData["BUSINESS TOOLS OPTION"]),
      rowData,
    ),
  };
}

function buildHmoClause_(row) {
  const hasHmo = String(row["HAS HMO"] || "")
    .trim()
    .toUpperCase();
  if (hasHmo !== "Y") {
    return getClauseText_("HMO", "NO");
  }

  const effectiveText = resolveClauseTemplate_(
    getClauseText_("HMO_EFFECTIVE", row["HMO EFFECTIVE OPTION"]),
    row,
  );
  const coverageText = getClauseText_(
    "HMO_COVERAGE",
    row["HMO COVERAGE OPTION"],
  );
  const mbl = row["HMO MBL"] || "XX,XXX";

  return (
    "The Employee shall be eligible for HMO enrollment effective " +
    effectiveText +
    ", subject to completion of the HMO enrollment and activation process. " +
    "Coverage shall apply to " +
    coverageText +
    ", with a Maximum Benefit Limit (MBL) of Php " +
    mbl +
    " per year, subject to Company policy and provider terms and conditions."
  );
}

function computeAnnualSalary_(row) {
  const raw = String(row["Basic Salary"] || "").replace(/,/g, "");
  const monthly = parseFloat(raw);
  if (isNaN(monthly)) return "";
  const annual = monthly * 12;
  return annual.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}
