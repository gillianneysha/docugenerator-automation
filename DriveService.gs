/**
 * DriveService.gs
 */

function duplicateDocToFolder_(sourceDocId, newName, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const sourceFile = DriveApp.getFileById(sourceDocId);
  const copy = sourceFile.makeCopy(newName, folder);
  return copy.getId();
}

function convertDocToPdf_(docId, folderId, fileName) {
  const doc = DriveApp.getFileById(docId);
  const pdfBlob = doc.getAs('application/pdf').setName(fileName + '.pdf');
  const folder = DriveApp.getFolderById(folderId);
  const pdfFile = folder.createFile(pdfBlob);
  return pdfFile.getId();
}

function getFolderIdOrDefault_(explicitFolderId) {
  return explicitFolderId || getSetting_('Default Output Folder');
}
