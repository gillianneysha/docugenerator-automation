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
  "HMO OPTION": "HMO",
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
    const rawValue = normalizeInlineBullets_(
      rowData[key] === null || rowData[key] === undefined
        ? ""
        : String(rowData[key]).replace(/\r\n/g, "\n"),
    );

    if (rawValue.indexOf("\n") === -1) {
      // Fast path: single-line value, plain inline text swap.
      const value = escapeReplacement_(rawValue);
      body.replaceText(escapeRegex_(placeholder), value);
      if (header) header.replaceText(escapeRegex_(placeholder), value);
      if (footer) footer.replaceText(escapeRegex_(placeholder), value);
    } else {
      // Multi-line value (e.g. a Job Description with several lines/bullets).
      // replaceText can't create real paragraph breaks or bullets, so build
      // actual paragraphs / list items instead. Blank lines in the cell
      // become blank paragraphs (spacing preserved), and lines starting with
      // "- ", "* ", or "• " become real bulleted list items.
      replaceMultilinePlaceholder_(body, placeholder, rawValue);
      if (header) replaceMultilinePlaceholder_(header, placeholder, rawValue);
      if (footer) replaceMultilinePlaceholder_(footer, placeholder, rawValue);
    }
  });
}

// $ has special meaning in Docs' replaceText replacement string (regex
// backreferences); escape it so a literal "$" in a value (e.g. "$50,000")
// isn't swallowed.
function escapeReplacement_(str) {
  return str.replace(/\$/g, "$$$$");
}

// Some sheet cells use a bullet character typed inline ("...text.  • Next
// item") instead of an actual line break (Alt+Enter) before each bullet.
// Insert a real newline before every "•" so those still become real Docs
// bullets, even without a literal line break in the source cell.
function normalizeInlineBullets_(value) {
  return value.replace(/\s*•\s*/g, "\n• ").replace(/^\n/, "");
}

function replaceMultilinePlaceholder_(container, placeholder, value) {
  const MIN_SPACING_AFTER = 8; // safety net only, if the template paragraph has 0pt spacing set

  const lines = value.split("\n");
  const searchPattern = escapeRegex_(placeholder);

  let result = container.findText(searchPattern);
  while (result !== null) {
    const element = result.getElement();
    const startOffset = result.getStartOffset();
    const endOffsetInclusive = result.getEndOffsetInclusive();
    const text = element.editAsText();
    const fullText = text.getText();
    const before = fullText.substring(0, startOffset);
    const after = fullText.substring(endOffsetInclusive + 1);

    // Capture the placeholder's OWN character-level formatting (color,
    // highlight, bold, font) directly from the text run, before anything
    // gets modified. Paragraph-level attributes alone don't reliably carry
    // an explicit run override like a highlight color.
    const runAttributes = text.getAttributes(startOffset);
    Logger.log(
      "DEBUG placeholder='" +
        placeholder +
        "' runAttributes=" +
        JSON.stringify(runAttributes),
    );

    let para = element.getParent();
    while (
      para &&
      para.getType() !== DocumentApp.ElementType.PARAGRAPH &&
      para.getType() !== DocumentApp.ElementType.LIST_ITEM
    ) {
      para = para.getParent();
    }

    const insertionContainer = para ? para.getParent() : null;
    Logger.log(
      "DEBUG paraAttributes=" +
        JSON.stringify(para ? para.getAttributes() : null),
    );
    const canInsert =
      insertionContainer &&
      typeof insertionContainer.insertParagraph === "function";

    if (!para || !canInsert) {
      text.deleteText(startOffset, endOffsetInclusive);
      text.insertText(startOffset, escapeReplacement_(lines.join("\v")));
      result = container.findText(searchPattern);
      continue;
    }

    // Combine: paragraph-level spacing/alignment from the paragraph itself,
    // character-level styling (color, bold, highlight, font) from the run —
    // both sourced from the template, none hardcoded.
    const templateAttributes = Object.assign(
      {},
      para.getAttributes(),
      runAttributes,
    );
    const currentSpacing =
      templateAttributes[DocumentApp.Attribute.SPACING_AFTER];
    if (!currentSpacing || currentSpacing < MIN_SPACING_AFTER) {
      templateAttributes[DocumentApp.Attribute.SPACING_AFTER] =
        MIN_SPACING_AFTER;
    }

    const bodyIndex = insertionContainer.getChildIndex(para);
    const placeholderIsAloneOnLine =
      before.trim() === "" && after.trim() === "";

    const built = lines.map((line) => {
      const m = line.match(/^\s*[-*•]\s+(.*)$/);
      return { bullet: !!m, text: m ? m[1] : line };
    });

    let insertAt = bodyIndex;
    let lastPara;

    if (placeholderIsAloneOnLine) {
      if (built[0].bullet) {
        lastPara = insertionContainer.insertListItem(bodyIndex, built[0].text);
        lastPara.setGlyphType(DocumentApp.GlyphType.BULLET);
        lastPara.editAsText().setAttributes(templateAttributes);
        para.removeFromParent();
      } else {
        text.insertText(startOffset, built[0].text);
        const oldPlaceholderStart = startOffset + built[0].text.length;
        const oldPlaceholderEnd = oldPlaceholderStart + (endOffsetInclusive - startOffset);
        text.deleteText(oldPlaceholderStart, oldPlaceholderEnd);
        // Apply formatting via the TEXT object with an explicit range, not
        // Paragraph.setAttributes() — calling setAttributes() directly on a
        // Paragraph does not reliably cascade run-level formatting (bold,
        // highlight, font) into its text the way the Text object does.
        if (built[0].text.length > 0) {
          text.setAttributes(
            startOffset,
            startOffset + built[0].text.length - 1,
            templateAttributes,
          );
        }
        lastPara = para;
      }
    } else {
      text.deleteText(startOffset, endOffsetInclusive);
      text.insertText(startOffset, built[0].text);
      if (built[0].text.length > 0) {
        text.setAttributes(
          startOffset,
          startOffset + built[0].text.length - 1,
          templateAttributes,
        );
      }
      if (after) {
        const afterStart = startOffset + built[0].text.length;
        text.deleteText(afterStart, afterStart + after.length - 1);
      }
      lastPara = para;
    }

    for (let i = 1; i < built.length; i++) {
      insertAt++;
      if (built[i].bullet) {
        lastPara = insertionContainer.insertListItem(insertAt, built[i].text);
        lastPara.setGlyphType(DocumentApp.GlyphType.BULLET);
        lastPara.editAsText().setAttributes(templateAttributes);
      } else if (built[i].text.trim() === "") {
        lastPara = insertionContainer.insertParagraph(insertAt, "");
        lastPara.editAsText().setAttributes(templateAttributes);
      } else {
        lastPara = insertionContainer.insertParagraph(insertAt, built[i].text);
        lastPara.editAsText().setAttributes(templateAttributes);
      }
    }

    if (!placeholderIsAloneOnLine && after) {
      lastPara.asText().appendText(after);
    }

    result = container.findText(searchPattern);
  }
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