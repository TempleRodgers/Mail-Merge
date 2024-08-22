/**
 * Temple Rodgers - 17/1/24
 * Mail merge, getting data from a selected spreadsheet
 * which contains sender data on one tab and merge data
 * on another
 * 
 * this script belongs to the project:
 *  https://console.cloud.google.com/home/dashboard?project=hackneymailmergesingledocument&authuser=0
 * 
 * named hackneyMailMergeSingleDocument, project number 767490064715
 * 
 * this video is useful: https://www.youtube.com/watch?v=QNPPEB64QbI&t=1625s
 */
function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  // console.log("adding the pull-down menu");
  // need to use createAddonMenu because this will be an extension
  var menu = DocumentApp.getUi().createAddonMenu();
  menu.addItem('NEW Single Letter mail merge', 'showSheetPickerDialog');
  menu.addToUi(); 
}

// global variables for the selected sheet
var selectedSheetUrl = null;
var progress = {
  processed: 0,
  total: 0
};

function showSheetPickerDialog() {
  // Create a custom dialog box with a picker to select a Google Sheet from the current folder
  var htmlOutput = HtmlService.createHtmlOutputFromFile('SheetPickerDialog')
      .setWidth(400)
      .setHeight(320);
  
  DocumentApp.getUi().showModalDialog(htmlOutput, 'Select Google Sheet with mailing list');
}

function setSelectedSheetUrl(url) {
  // Set the selected Google Sheet URL to the global variable
  selectedSheetUrl = url;
  performBodyMerge(selectedSheetUrl);
}

function getFolderSpreadsheets() {
  // Find all the spreadsheets in the current folder so the pick list can be presented
  var folderId = DriveApp.getFileById(DocumentApp.getActiveDocument().getId()).getParents().next().getId();
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var sheets = [];

  while (files.hasNext()) {
    var file = files.next();
    sheets.push({ name: file.getName(), url: file.getUrl() });
  }

  return sheets;
}
