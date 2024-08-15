/**
 * Temple Rodgers - 21/2/24
 * Mail merge, getting data from a selected spreadsheet
 * which contains sender data on one tab and merge data
 * on another
 * this video is useful: https://www.youtube.com/watch?v=QNPPEB64QbI&t=1625s
 */
function showSheetPickerDialog() {
  // Create a custom dialog box with a picker to select a Google Sheet from the current folder
  var htmlOutput = HtmlService.createHtmlOutputFromFile('sheetPickerDialog')
      .setWidth(400)
      .setHeight(300);
  DocumentApp.getUi().showModalDialog(htmlOutput, 'Select Google Sheet with mailing list');
}

function setSelectedSheetUrl(url) {
  // Set the selected Google Sheet URL to the global variable
  selectedSheetUrl = url;
  Logger.log(selectedSheetUrl);
  performMailMerge(url);
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
