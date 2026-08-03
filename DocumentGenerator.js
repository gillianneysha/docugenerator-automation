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
// in the Clause Library sheet. Read from the "Clause Key Registry" sheet at
// runtime so HR can add new clause types (new column + new clause key) via
// the sheet, without a code change. Cached for the lifetime of one execution,
// same pattern as getClauseLibrary_().
let clauseOptionColumnsCache_ = null;
function getClauseOptionColumns_() {
  if (clauseOptionColumnsCache_ === null) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      "Clause Key Registry",
    );
    const map = {};
    if (sheet) {
      getSheetDataAsObjects_("Clause Key Registry").rows.forEach((r) => {
        const col = String(r["Option Column"] || "").trim();
        const key = String(r["Clause Key"] || "").trim();
        if (col && key) map[col] = key;
      });
    }
    clauseOptionColumnsCache_ = map;
  }
  return clauseOptionColumnsCache_;
}

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
  const clauseOptionColumns = getClauseOptionColumns_();
  Object.keys(clauseOptionColumns).forEach((column) => {
    const clauseKey = clauseOptionColumns[column];
    const text = resolveClauseText_(clauseKey, rowData[column], rowData);
    if (text !== null) derived[clauseKey] = text;
  });
  return derived;
}

// Remofirst-style HMO input: unlike CLAUSE_OPTION_COLUMNS (one column -> one
// Clause Library option code), this sheet spreads the HMO choice across
// four columns (HAS HMO / HMO EFFECTIVE OPTION / HMO COVERAGE OPTION /
// HMO EFFECTIVE MONTHS). This maps those four into the matching Clause
// Library "HMO" option code (NO, or 1-6), then resolves it the normal way.
// {{HMO MBL}} and {{HMO EFFECTIVE MONTHS}} inside the resolved clause text
// are filled automatically by resolveClauseText_ since they're already
// columns in rowData.
function resolveHmoOptionCode_(rowData) {
  const hasHmo = String(rowData["HAS HMO"] || "")
    .trim()
    .toUpperCase();
  if (hasHmo !== "YES") return "NO";

  const effective = String(rowData["HMO EFFECTIVE OPTION"] || "")
    .trim()
    .toUpperCase();
  const coverage = String(rowData["HMO COVERAGE OPTION"] || "")
    .trim()
    .toUpperCase();

  const principalOnly = coverage === "PRINCIPAL ONLY";
  const withDependents = coverage === "PRINCIPAL + DEPENDENTS";

  if (effective === "FIRST DAY") {
    if (principalOnly) return "1";
    if (withDependents) return "2";
  } else if (effective === "REGULARIZATION") {
    if (principalOnly) return "3";
    if (withDependents) return "4";
  } else if (effective === "AFTER MONTHS") {
    if (principalOnly) return "5";
    if (withDependents) return "6";
  }
  // Unrecognized combination -> caller gets null and leaves {{HMO_CLAUSE}}
  // untouched, rather than silently guessing.
  return null;
}

function deriveHmoClauseField_(rowData) {
  if (!("HAS HMO" in rowData)) return {}; // sheet doesn't use this pattern
  const optionCode = resolveHmoOptionCode_(rowData);
  if (optionCode === null) return {};
  const text = resolveClauseText_("HMO", optionCode, rowData);
  return text !== null ? { HMO_CLAUSE: text } : {};
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

// Attribute keys that belong to the paragraph itself (spacing, alignment,
// indentation) as opposed to character-level run formatting.
const PARAGRAPH_LEVEL_ATTRIBUTE_KEYS = [
  DocumentApp.Attribute.HEADING,
  DocumentApp.Attribute.HORIZONTAL_ALIGNMENT,
  DocumentApp.Attribute.INDENT_END,
  DocumentApp.Attribute.INDENT_FIRST_LINE,
  DocumentApp.Attribute.INDENT_START,
  DocumentApp.Attribute.LEFT_TO_RIGHT,
  DocumentApp.Attribute.LINE_SPACING,
  DocumentApp.Attribute.SPACING_AFTER,
  DocumentApp.Attribute.SPACING_BEFORE,
];

function replaceMultilinePlaceholder_(container, placeholder, value) {
  const MIN_SPACING_AFTER = 8;

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

    // Character-level formatting (bold, highlight, font, color) — read from
    // the placeholder's own text run.
    const runAttributes = text.getAttributes(startOffset);

    let para = element.getParent();
    while (
      para &&
      para.getType() !== DocumentApp.ElementType.PARAGRAPH &&
      para.getType() !== DocumentApp.ElementType.LIST_ITEM
    ) {
      para = para.getParent();
    }

    const insertionContainer = para ? para.getParent() : null;
    const canInsert =
      insertionContainer &&
      typeof insertionContainer.insertParagraph === "function";

    if (!para || !canInsert) {
      text.deleteText(startOffset, endOffsetInclusive);
      text.insertText(startOffset, escapeReplacement_(lines.join("\v")));
      result = container.findText(searchPattern);
      continue;
    }

    // Paragraph-level formatting (spacing, alignment, indent) — read
    // separately, and ONLY these keys, from the paragraph itself.
    const fullParaAttributes = para.getAttributes();
    const paragraphAttributes = {};
    PARAGRAPH_LEVEL_ATTRIBUTE_KEYS.forEach((key) => {
      paragraphAttributes[key] = fullParaAttributes[key];
    });
    if (
      !paragraphAttributes[DocumentApp.Attribute.SPACING_AFTER] ||
      paragraphAttributes[DocumentApp.Attribute.SPACING_AFTER] <
        MIN_SPACING_AFTER
    ) {
      paragraphAttributes[DocumentApp.Attribute.SPACING_AFTER] =
        MIN_SPACING_AFTER;
    }

    const bodyIndex = insertionContainer.getChildIndex(para);
    const placeholderIsAloneOnLine =
      before.trim() === "" && after.trim() === "";

    const built = lines.map((line) => {
      const m = line.match(/^\s*[-*•]\s+(.*)$/);
      return { bullet: !!m, text: m ? m[1] : line };
    });

    // Applies character-level formatting to a newly-created paragraph/list
    // item's full text range, plus paragraph-level formatting to the
    // element itself. Kept as two separate calls — mixing both categories
    // into one setAttributes() call on a Paragraph is unreliable.
    function applyTemplateFormatting(paraOrListItem) {
      paraOrListItem.setAttributes(paragraphAttributes);
      const t = paraOrListItem.editAsText();
      const len = t.getText().length;
      if (len > 0) t.setAttributes(0, len - 1, runAttributes);
    }

    let insertAt = bodyIndex;
    let lastPara;

    if (placeholderIsAloneOnLine) {
      if (built[0].bullet) {
        lastPara = insertionContainer.insertListItem(bodyIndex, built[0].text);
        lastPara.setGlyphType(DocumentApp.GlyphType.BULLET);
        applyTemplateFormatting(lastPara);
        para.removeFromParent();
      } else {
        text.insertText(startOffset, built[0].text);
        const oldPlaceholderStart = startOffset + built[0].text.length;
        const oldPlaceholderEnd =
          oldPlaceholderStart + (endOffsetInclusive - startOffset);
        text.deleteText(oldPlaceholderStart, oldPlaceholderEnd);
        // Order matters: paragraph-level attributes must be applied BEFORE
        // character-level ones, or setting paragraph attributes afterward
        // resets the run-level highlight/bold that was just applied.
        para.setAttributes(paragraphAttributes);
        if (built[0].text.length > 0) {
          text.setAttributes(
            startOffset,
            startOffset + built[0].text.length - 1,
            runAttributes,
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
          runAttributes,
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
      } else if (built[i].text.trim() === "") {
        lastPara = insertionContainer.insertParagraph(insertAt, "");
      } else {
        lastPara = insertionContainer.insertParagraph(insertAt, built[i].text);
      }
      applyTemplateFormatting(lastPara);
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
