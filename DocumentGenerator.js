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
function generateDocuments(templateName, rowMode, selectedRowNumbers, outputFormat, nameTemplate) {
  const template = getTemplateByName_(templateName);
  const { rows } = getSheetDataAsObjects_(template['Source Sheet']);
  const folderId = getFolderIdOrDefault_(template['Output Folder ID']);

  const targetRows = rowMode === 'all'
    ? rows
    : rows.filter(r => selectedRowNumbers.includes(r.__row));

  const results = [];

  targetRows.forEach(row => {
    try {
      const result = generateSingleDocument_(template, row, folderId, outputFormat, nameTemplate);
      writeStatus_(template['Source Sheet'], row.__row, 'Generate Status', '✅ Generated');
      appendLog_(templateName, result.fileName, 'Success', '');
      results.push({ row: row.__row, status: 'success', fileName: result.fileName, docUrl: result.docUrl, pdfUrl: result.pdfUrl });
    } catch (err) {
      writeStatus_(template['Source Sheet'], row.__row, 'Generate Status', '⚠️ Error');
      appendLog_(templateName, '(row ' + row.__row + ')', 'Error', err.message);
      results.push({ row: row.__row, status: 'error', message: err.message });
    }
  });

  return results;
}

function generateSingleDocument_(template, rowData, folderId, outputFormat, nameTemplate) {
  const fileName = fillPlaceholders_(nameTemplate, rowData);

  // 1. Duplicate the template doc
  const newDocId = duplicateDocToFolder_(template['Google Doc ID'], fileName, folderId);

  // 2. Replace placeholders in body/header/footer
  const doc = DocumentApp.openById(newDocId);
  replaceAllPlaceholdersInDoc_(doc, rowData);
  doc.saveAndClose();

  const result = { fileName: fileName, docUrl: null, pdfUrl: null };

  // 3. Handle output format
  if (outputFormat === 'pdf' || outputFormat === 'both') {
    const pdfId = convertDocToPdf_(newDocId, folderId, fileName);
    result.pdfUrl = DriveApp.getFileById(pdfId).getUrl();
  }
  if (outputFormat === 'docx') {
    // Keep as Google Doc (or convert to actual .docx if the client truly needs a Word file
    // rather than a Google Doc — see note in README about Drive export to docx).
    result.docUrl = DriveApp.getFileById(newDocId).getUrl();
  } else if (outputFormat === 'pdf') {
    // Delete the intermediate Google Doc copy, keep only the PDF
    DriveApp.getFileById(newDocId).setTrashed(true);
  } else {
    result.docUrl = DriveApp.getFileById(newDocId).getUrl();
  }

  return result;
}

function replaceAllPlaceholdersInDoc_(doc, rowData) {
  const body = doc.getBody();
  const header = doc.getHeader();
  const footer = doc.getFooter();

  Object.keys(rowData).forEach(key => {
    if (key === '__row') return;
    const placeholder = '{{' + key + '}}';
    const value = rowData[key] === null || rowData[key] === undefined ? '' : String(rowData[key]);
    body.replaceText(escapeRegex_(placeholder), value);
    if (header) header.replaceText(escapeRegex_(placeholder), value);
    if (footer) footer.replaceText(escapeRegex_(placeholder), value);
  });
}

// Used for filenames, e.g. "Employment Contract - {{EmployeeName}}"
function fillPlaceholders_(templateString, rowData) {
  return templateString.replace(/{{\s*([\w]+)\s*}}/g, (match, key) => {
    return rowData[key] !== undefined ? String(rowData[key]) : match;
  });
}

function escapeRegex_(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
