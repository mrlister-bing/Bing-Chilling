const TRACKER_ID = '18jiRHZfpedExNJSXAe-R9c6_qWTT-pSqQ-ewlJlzj8E';
const TRACKER_SHEET_NAME = 'Sheet1';
const DESTINATION_FOLDER_ID = '1C7aG1Y9mdZ4VmMeNSiw5lFDw2yXfFcpM';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('SUBMIT')
    .addItem('Send Request', 'submitSheet')
    .addToUi();
}

function addTrackerEntry(linkToFile, builderName, mmEmail, numberOfHomes, unitVolumeYearOne, totalUnitVolume, numberYears, blendedDCM) {
  try {
    const trackerSs = SpreadsheetApp.openById(TRACKER_ID);
    const trackerSheet = trackerSs.getSheetByName(TRACKER_SHEET_NAME);
    
    if (!trackerSheet) {
      throw new Error(`Tracker sheet "${TRACKER_SHEET_NAME}" not found in spreadsheet with ID "${TRACKER_ID}"`);
    }
    
    const data = ['', linkToFile, new Date(), '', builderName, mmEmail, '', numberOfHomes, unitVolumeYearOne, totalUnitVolume, numberYears, blendedDCM];
    trackerSheet.appendRow(data);
  } catch (error) {
    throw new Error(`Error in addTrackerEntry: ${error.message}`);
  }
}

function sendSubmissionEmail(mmEmail, builderEmail, linkToCopiedSheet) {
  const subject = `Regional Single Family RFP Request: ${builderEmail}`;
  const body = `
Below is a request for a Regional Single Family RFP from ${mmEmail}
${linkToCopiedSheet}
`;

  let email = builderEmail;
  if (mmEmail !== "") {
    email += "," + mmEmail;
  }
  
  GmailApp.sendEmail(email, subject, body);
}

function submitSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Submission View");
  const builder = sheet.getRange("D3").getValue();
  const mmEmail = sheet.getRange("G6").getValue();
  const builderEmail = sheet.getRange("G7").getValue();
  
  const numberOfHomes = sheet.getRange("F16").getValue();
  
  const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  const copiedFile = DriveApp.getFileById(ss.getId()).makeCopy(builder, folder);
  const copiedIRFP = copiedFile.getUrl();
  
  sendSubmissionEmail(mmEmail, builderEmail, copiedIRFP);
  
  // Add call to function that fills out / appends to tracker
  addTrackerEntry(copiedIRFP, builder, mmEmail, numberOfHomes, 0, 0, 0, 0);
}
