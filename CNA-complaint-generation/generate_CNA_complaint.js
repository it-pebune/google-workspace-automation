const complaintTemplateDocId = "1fikjAhpgH3QdAnVhssiqlZ1u5fDOvZKsmSiOdE_67HQ";
const complaintsFolderId = "1S0wA-5a66Glrg_a8sRUQIV3APm30Ym-_";

const complaintsDataSheet = SpreadsheetApp.getActiveSpreadsheet();

/**
 * A special function that runs when the spreadsheet is open.
 * Adds a custom menu to the spreadsheet.
 */
function onOpen() {
  complaintsDataSheet.addMenu("SESIZĂRI", [
    { name: "GENEREAZĂ SESIZĂRILE LIPSĂ", functionName: "addComplaints" },
  ]);
}

function addComplaints() {
  // Get the first sheet
  const sheet = complaintsDataSheet.getSheets()[0];
  // Get all data
  const data = sheet.getDataRange().getValues();
  // Iterate data rows
  data.forEach((row, index) => {
    if (hasNoComplaint(sheet, row, index)) {
      addComplaint(sheet, row, index);
    }
  });
}

function hasNoComplaint(sheet, row, index) {
  const isHeaderRow = index === 0;
  const isEmptyRow = row[0] === "";
  const complaintLinkCell = sheet.getRange("S" + (index + 1));
  const hasComplaint = complaintLinkCell.getValue() !== "";
  return !isHeaderRow && !isEmptyRow && !hasComplaint;
}

function addComplaint(sheet, row, index) {
  const complaintDoc = createComplaintDoc(row);

  fillInComplaintDoc(complaintDoc, row);

  addLinktoComplaintDoc(sheet, index, complaintDoc);
}

function createComplaintDoc(row) {
  // Use the time zone of the spreadsheet for reading the date and time, otherwise
  // you will get the date and time with an offset equal to the difference between
  // the time zone used for reading and the time zone of the spreadsheet
  const time = Utilities.formatDate(
    row[4],
    complaintsDataSheet.getSpreadsheetTimeZone(),
    "dd.MM.yyyy_HH:mm"
  );
  const channel = row[2];
  const broadcast = row[3];
  const createdAt = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd_HHmmss"
  );

  // Cannot be declared globally because the simple trigger 'onOpen()'
  // cannot access services that require authorization (simple triggers
  // get fired automatically, without asking the user for authorization),
  // so it has no permission to call 'DriveApp.getFileById'.
  //
  // See https://developers.google.com/apps-script/guides/triggers#restrictions
  const complaintTemplateDoc = DriveApp.getFileById(complaintTemplateDocId);
  const complaintsFolder = DriveApp.getFolderById(complaintsFolderId);

  const complaintDoc = complaintTemplateDoc.makeCopy(
    channel + "-" + broadcast + "-" + time + "-" + createdAt,
    complaintsFolder
  );

  Logger.log("CREATED COMPLAINT: " + complaintDoc.getName());

  return complaintDoc;
}

function fillInComplaintDoc(complaintDoc, row) {
  const complaintDocFile = DocumentApp.openById(complaintDoc.getId());
  const complaintDocContents = complaintDocFile.getBody();

  replacePlaceholder(complaintDocContents, "{Sursa 1}", row[2]);
  replacePlaceholder(complaintDocContents, "{Sursa 2}", row[3]);
  replacePlaceholder(
    complaintDocContents,
    "{Sursa 3}",
    // Use the time zone of the spreadsheet for reading the date and time, otherwise
    // you will get the date and time with an offset equal to the difference between
    // the time zone used for reading and the time zone of the spreadsheet
    Utilities.formatDate(
      row[4],
      complaintsDataSheet.getSpreadsheetTimeZone(),
      "dd/MM/yyyy HH:mm:ss"
    )
  );
  replacePlaceholder(complaintDocContents, "{Sursa 4}", row[5]);
  replacePlaceholder(complaintDocContents, "{Sursa 5}", row[13]);
  replacePlaceholder(complaintDocContents, "{Sursa 6}", row[7]);
  replacePlaceholder(complaintDocContents, "{Sursa 7}", row[14]);
  replacePlaceholder(complaintDocContents, "{Sursa 8}", row[9]);
  replacePlaceholder(complaintDocContents, "{Sursa 9}", row[15]);
  replacePlaceholder(complaintDocContents, "{Sursa 10}", row[11]);
  replacePlaceholder(complaintDocContents, "{Sursa 11}", row[16]);

  complaintDocFile.saveAndClose();

  Logger.log("FILLED IN COMPLAINT: " + complaintDoc.getName());
}

function replacePlaceholder(docContents, placeholder, replacement) {
  if (replacement !== "" && replacement !== "#N/A") {
    docContents.replaceText(placeholder, replacement);
  } else {
    removeEnclosingParagraph(docContents, placeholder);
  }
}

function removeEnclosingParagraph(docContents, placeholder) {
  docContents
    .findText(placeholder)
    .getElement()
    .getParent()
    .asParagraph()
    .removeFromParent();
}

function addLinktoComplaintDoc(sheet, index, complaintDoc) {
  const complaintLinkCell = sheet.getRange("S" + (index + 1));
  complaintLinkCell.setValue(complaintDoc.getUrl());

  Logger.log("ADDED LINK TO COMPLAINT: " + complaintDoc.getUrl());
}
