/**
 * Temple Rodgers - 16/01/2024
 * Simple mail merge from one template,
 * creating many merged output files
 */
function onInstall(e) {
  onOpen(e);
}

function onOpen(e)
{
  // console.log("adding the pull-down menu");
  var menu = DocumentApp.getUi().createAddonMenu();
  menu.addItem('Run Mail merge', 'showSheetPickerDialog');
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
       sheets.push({name: file.getName(), url: file.getUrl()});
   }
   return sheets;
}

// set the mail merge spreadsheet variables
var sheet = null
   , mailMergeTab = null
   , senderDataTab = null;

function performMailMerge(spreadsheetURL) {
  spreadsheetURL = "https://docs.google.com/spreadsheets/d/1kzwtS7uywB-faZjZRRI5vMN1c7pyZoMwPX1zWBEGqWY/edit?usp=drivesdk";
   const templateName = DocumentApp.getActiveDocument().getName();

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

       // Retrieve sender data with flexibility for column names
       const senderData = senderDataTab.getDataRange().getValues();

       // Get merge data with flexible column names
       const data = mailMergeTab.getDataRange().getValues();
       // const columnHeaders = senderData[0].concat(data[0]).map(header => header.toUpperCase());
       const columnHeaders = senderData[0].concat(data[0]); // Get headers for mapping
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
           const copyId = mergeTemplate(toMergeData, senderData);

           // Rename the file
           DriveApp.getFileById(copyId).setName(templateName + ' - ' + toMergeData["to_name"]);
       }
   } catch (error) {
       console.error(`An error occurred: ${error}`);
       // Display an error message to the user (implementation omitted for brevity)
   }
}

function mergeTemplate(mergeData, senderData) {
   const templateName = DocumentApp.getActiveDocument().getName();

   try {
       const copyId = copyTemplate();
       const mergeCopy = DocumentApp.openById(copyId);
       const body = mergeCopy.getBody();

       // Create an array to hold the replacement pairs
       const replacements = [];

       // Collect replacement pairs
       //replacements = [];
       for (const [key, value] of Object.entries(mergeData)) {
           //replacements.push({placeholder: `{{${key.toUpperCase()}}}`, value});
           replacements.push({placeholder: `{{${key}}}`, value});
       }

       for (const [key, value] of Object.entries(senderData)) {
           //replacements.push({placeholder: `{{${key.toUpperCase()}}}`, value});
           replacements.push({placeholder: `{{${key}}}`, value});
       }

       // Batch replace text in the document
       for (const replacement of replacements) {
             // make the search & replace case insensitive using a REGEX
             var caseInsensitivePlaceholder = '(?i)' + replacement.placeholder;
             body.replaceText(caseInsensitivePlaceholder, replacement.value);
//           body.replaceText(replacement.placeholder, replacement.value);
       }

       // Rename the file
       DriveApp.getFileById(copyId).setName(templateName + ' - ' + mergeData["to_name"]);
       return copyId;
   } catch (error) {
       console.error(`An error occurred: ${error}`);
       return error;
   }
}

function copyTemplate() {
   const templateName = DocumentApp.getActiveDocument().getName();

   // get the id of the current document, which is the template
   const templateId = DocumentApp.getActiveDocument().getId();
   try {
       // pull back the template file
       const templateFile = DriveApp.getFileById(templateId);

       // copy the template and give it a temporary name (to be replaced later)
       const copy = templateFile.makeCopy(templateName + ' - blank');

       // return with the id of the copy
       return copy.getId();
   } catch (error) {
       console.error(`An error occurred: ${error}`);
       return error;
   }
}
