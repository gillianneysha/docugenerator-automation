/**
 * TemplateManager.gs
 * Scans {{placeholders}} in a Doc, validates them against a sheet's headers,
 * and handles registration.
 */

function scanPlaceholders_(docId) {
  let doc;
  try {
    doc = DocumentApp.openById(docId);
  } catch (e) {
    throw new Error(
      "Could not open the template document. It may have been deleted, moved, or you may not have access. Doc ID: " +
        docId,
    );
  }

  let fullText = doc.getBody().getText();
  const header = doc.getHeader();
  const footer = doc.getFooter();
  if (header) fullText += "\n" + header.getText();
  if (footer) fullText += "\n" + footer.getText();

  // Allow any character except braces, so placeholders can contain spaces,
  // e.g. {{Employee Name}} matching an "Employee Name" column header.
  const matches = fullText.match(/{{\s*[^{}]+?\s*}}/g) || [];
  return [...new Set(matches.map((m) => m.replace(/[{}]/g, "").trim()))];
}

const DERIVED_KEYS_ = ["day", "month"];

function validatePlaceholders_(docId, sourceSheetName) {
  const placeholders = scanPlaceholders_(docId);
  const { headers } = getSheetDataAsObjects_(sourceSheetName);
  const settingsKeys = Object.keys(getAllSettings_());

  const unknown = placeholders.filter(
    (p) =>
      !headers.includes(p) &&
      !settingsKeys.includes(p) &&
      !DERIVED_KEYS_.includes(p),
  );
  return {
    valid: unknown.length === 0,
    placeholders,
    unknown,
    headers,
    settingsKeys,
  };
}

/**
 * Called from the Register Template sidebar.
 */
function registerTemplate(templateName, docId, sourceSheet, outputFolderId) {
  const nameTrimmed = (templateName || "").trim();
  if (!nameTrimmed) {
    return { success: false, message: "Template name cannot be empty." };
  }

  const allTemplates = getAllTemplateRows_();

  const duplicateName = allTemplates.find(
    (r) =>
      String(r["Template Name"]).trim().toLowerCase() ===
      nameTrimmed.toLowerCase(),
  );
  if (duplicateName) {
    return {
      success: false,
      message:
        'A template named "' +
        nameTrimmed +
        '" already exists (Status: ' +
        duplicateName["Status"] +
        "). Choose a different name, or edit the existing row in the Template Registry sheet.",
    };
  }

  const duplicateDoc = allTemplates.find(
    (r) => String(r["Google Doc ID"]).trim() === String(docId).trim(),
  );
  if (duplicateDoc) {
    return {
      success: false,
      message:
        'This Google Doc is already registered as "' +
        duplicateDoc["Template Name"] +
        '" (Status: ' +
        duplicateDoc["Status"] +
        "). Each template document should only be registered once — " +
        "make a copy of the Doc first if you need a variant.",
    };
  }

  const validation = validatePlaceholders_(docId, sourceSheet);
  if (!validation.valid) {
    return {
      success: false,
      message:
        "Unknown placeholder(s): " +
        validation.unknown.join(", ") +
        '. These do not match any column header in "' +
        sourceSheet +
        '" or a Settings key.',
    };
  }
  addTemplateToRegistry_(nameTrimmed, docId, sourceSheet, outputFolderId);
  return {
    success: true,
    message:
      'Template "' +
      nameTrimmed +
      '" registered with placeholders: ' +
      validation.placeholders.join(", "),
  };
}

function listTemplateNames() {
  return getTemplateRegistry_().map((r) => r["Template Name"]);
}

function revalidateAllTemplates_UI() {
  const allTemplates = getAllTemplateRows_();
  const results = [];

  allTemplates.forEach((t) => {
    const templateName = t["Template Name"];
    const status = t["Status"];

    try {
      const validation = validatePlaceholders_(
        t["Google Doc ID"],
        t["Source Sheet"],
      );
      if (validation.valid) {
        results.push({ templateName, status, ok: true, message: "OK" });
      } else {
        results.push({
          templateName,
          status,
          ok: false,
          message: "Unknown placeholder(s): " + validation.unknown.join(", "),
        });
      }
    } catch (err) {
      results.push({
        templateName,
        status,
        ok: false,
        message: err.message,
      });
    }
  });

  return results;
}
