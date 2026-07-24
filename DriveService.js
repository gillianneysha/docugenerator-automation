/**
 * DriveService.gs
 */

function duplicateDocToFolder_(sourceDocId, newName, folderId) {
  let sourceFile;
  try {
    sourceFile = DriveApp.getFileById(sourceDocId);
  } catch (e) {
    throw new Error(
      "Template document could not be opened (it may have been deleted or moved). Doc ID: " +
        sourceDocId,
    );
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(
      "Output folder could not be opened (it may have been deleted or you may not have access). Folder ID: " +
        folderId,
    );
  }

  try {
    const copy = sourceFile.makeCopy(newName, folder);
    return copy.getId();
  } catch (e) {
    throw new Error(
      "Could not create a copy of the template. Check that you have edit access to the output folder. (" +
        e.message +
        ")",
    );
  }
}

function convertDocToPdf_(docId, folderId, fileName) {
  let doc, folder;
  try {
    doc = DriveApp.getFileById(docId);
  } catch (e) {
    throw new Error("Could not find the generated document to convert to PDF.");
  }
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(
      "Output folder could not be opened when saving the PDF. Folder ID: " +
        folderId,
    );
  }

  try {
    const pdfBlob = doc.getAs("application/pdf").setName(fileName + ".pdf");
    const pdfFile = folder.createFile(pdfBlob);
    return pdfFile.getId();
  } catch (e) {
    throw new Error("PDF conversion failed. (" + e.message + ")");
  }
}

function getFolderIdOrDefault_(explicitFolderId) {
  const folderId = explicitFolderId || getSetting_("Default Output Folder");
  if (!folderId) {
    throw new Error(
      'No output folder configured. Set one on the template in the Template Registry, or fill in "Default Output Folder" in the Settings sheet.',
    );
  }
  return folderId;
}

function initializeDriveFolders_() {
  const rootName = "HR Documents";
  const root = getOrCreateFolder_(DriveApp.getRootFolder(), rootName);

  const templatesFolder = getOrCreateFolder_(root, "Templates");
  const generatedDocsFolder = getOrCreateFolder_(root, "Generated Docs");
  const generatedPdfsFolder = getOrCreateFolder_(root, "Generated PDFs");
  const archiveFolder = getOrCreateFolder_(root, "Archive");

  return {
    rootId: root.getId(),
    rootUrl: root.getUrl(),
    templatesId: templatesFolder.getId(),
    generatedDocsId: generatedDocsFolder.getId(),
    generatedPdfsId: generatedPdfsFolder.getId(),
    archiveId: archiveFolder.getId(),
  };
}

function getOrCreateFolder_(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}
