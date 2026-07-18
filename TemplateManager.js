/**
 * TemplateManager.gs
 * Scans {{placeholders}} in a Doc, validates them against a sheet's headers,
 * and handles registration.
 */

function scanPlaceholders_(docId) {
  const doc = DocumentApp.openById(docId);

  let fullText = doc.getBody().getText();
  const header = doc.getHeader();
  const footer = doc.getFooter();
  if (header) fullText += '\n' + header.getText();
  if (footer) fullText += '\n' + footer.getText();

  const matches = fullText.match(/{{\s*[\w]+\s*}}/g) || [];
  const unique = [...new Set(matches.map(m => m.replace(/[{}]/g, '').trim()))];
  return unique;
}

function validatePlaceholders_(docId, sourceSheetName) {
  const placeholders = scanPlaceholders_(docId);
  const { headers } = getSheetDataAsObjects_(sourceSheetName);

  const unknown = placeholders.filter(p => !headers.includes(p));
  return {
    valid: unknown.length === 0,
    placeholders,
    unknown,
    headers
  };
}

/**
 * Called from the Register Template sidebar.
 */
function registerTemplate(templateName, docId, sourceSheet, outputFolderId) {
  const validation = validatePlaceholders_(docId, sourceSheet);
  if (!validation.valid) {
    return {
      success: false,
      message: 'Unknown placeholder(s): ' + validation.unknown.join(', ') +
        '. These do not match any column header in "' + sourceSheet + '".'
    };
  }
  addTemplateToRegistry_(templateName, docId, sourceSheet, outputFolderId);
  return { success: true, message: 'Template "' + templateName + '" registered with placeholders: ' + validation.placeholders.join(', ') };
}

function listTemplateNames() {
  return getTemplateRegistry_().map(r => r['Template Name']);
}
