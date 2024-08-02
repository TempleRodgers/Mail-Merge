/**
 * Temple Rodgers - 17/1/24
 * Mail merge, getting data from a selected spreadsheet
 * which contains sender data on one tab and merge data
 * on another
 * this video is useful: https://www.youtube.com/watch?v=QNPPEB64QbI&t=1625s
 */
function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  // console.log("adding the pull-down menu");
  var menu = DocumentApp.getUi().createAddonMenu();
  menu.addItem('Single Letter mail merge', 'showSheetPickerDialog');
  menu.addToUi(); 
}

// global variable for the selected sheet
var selectedSheetUrl;

function showSheetPickerDialog() {
  // Create a custom dialog box with a picker to select a Google Sheet from the current folder
  var htmlOutput = HtmlService.createHtmlOutputFromFile('SheetPickerDialog')
      .setWidth(400)
      .setHeight(300);
  
  DocumentApp.getUi().showModalDialog(htmlOutput, 'Select Google Sheet with mailing list');
}

function setSelectedSheetUrl(url) {
  // Set the selected Google Sheet URL to the global variable
  selectedSheetUrl = url;
  performMailMerge(selectedSheetUrl);
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

function performMailMerge(spreadsheetURL) {
  // get the id of the current document, which is the template
  // there are two docs: the template and the merge document
  // i.e. template... and mergeDoc...
  const templateId = DocumentApp.getActiveDocument().getId();
  const templateName = DocumentApp.getActiveDocument().getName();
  const template = DocumentApp.openById(templateId);
  const templateParagraphs = Array.from(template.getBody().getParagraphs());

  var mergeDoc = [];
  var mergeDocId = "";
  const finishedFileName = "merged document"

  // set the mail merge spreadsheet variables
  // the script gathers merge data from two
  // tabs in the spreadsheet: Mail_Merge and
  // Sender_Details
  var sheet = null
    ,mailMergeTab = null
    ,senderDataTab = null;

  try {
    // Open the spreadsheet and get sheets
    const sheet = SpreadsheetApp.openByUrl(spreadsheetURL);
    const mailMergeTab = sheet.getSheetByName('Mail_Merge');
    if (!mailMergeTab) {
      throw new Error('Sheet named "Mail_Merge" not found.');
    }
    const senderDataTab = sheet.getSheetByName('Sender_Details');
    if (!senderDataTab) {
      throw new Error('Sheet named "Sender_Details" not found.');
    }

    try {
      // pull back the template file and get its information
      const mergeDocument = DriveApp.getFileById(templateId).makeCopy('TempMergeFile - delete');
      mergeDocId = mergeDocument.getId();
      // copy the template and give it a temporary name (to be replaced later)
      mergeDoc = DocumentApp.openById(mergeDocId);
      mergeDoc.getBody().clear(); // clear the template

    } catch (error) {
      console.error(`An error occurred: ${error}`);
    }

    // Retrieve sender data with flexible column names
    const senderData = senderDataTab.getDataRange().getValues();

    // Get merge data with flexible column names -
    // the script allows the user to put their own
    // column names in the spreadsheet and then to
    // reference them in the merge document
    const data = mailMergeTab.getDataRange().getValues();
    const columnHeaders = senderData[0].concat(data[0]); // Get headers for mapping

    // Filter out header row
    const mergeData = data.slice(1);

    // Process each merge record
    for (let i = 0; i < mergeData.length; i++) {
      const record = senderData[1].concat(mergeData[i]);
      const toMergeData = {};

      // Map merge fields dynamically based on headers
      for (let j = 0; j < columnHeaders.length; j++) {
        toMergeData[columnHeaders[j]] = record[j];
      }

      // Fill in additional data
      toMergeData["date"] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy MMMM dd");

      // Perform the merge
      mergeTemplate_(mergeDoc,toMergeData,templateParagraphs);

      console.log(`Merged letter ${i + 1}: docs.google.com/document/d/${mergeDocId}/edit`);
    }
      // Rename the file
    DriveApp.getFileById(mergeDocId).setName(templateName + ' - ' + finishedFileName);
  } catch (error) {
    console.error(`An error occurred: ${error}`);
  }
}
