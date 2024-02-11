/**
 * given the URL of the merge spreadsheet, merge the template document
 * with the spreadsheet content into a single document
 */
function performMailMerge(spreadsheetURL) {
//function performMailMerge() {
  const spreadsheetURL = "https://docs.google.com/spreadsheets/d/158Md3meKiyZAO2aXj5qnQaosCRU4Fp7R_Ecss7gsrr0/edit?usp=drivesdk";
  try {
    // Open the spreadsheet and get sheets
    const sheet = SpreadsheetApp.openByUrl(spreadsheetURL);
    const mailMergeTab = sheet.getSheetByName('Mail_Merge');
    
    if (!mailMergeTab) {
      throw new Error('Sheet named "Mail_Merge" not found.');
    }
    // Check if Sender_Details sheet exists; handle its absence gracefully
    
    const senderDataTab = sheet.getSheetByName('Sender_Details');
    // Retrieve sender data with flexible column names
    // Define default sender data (e.g., empty object or set key-value pairs)
    const senderData = [];
    if (!senderDataTab) {
      Logger.log('Sheet named "Sender_Details" not found, not using sender data.');
    } else {
        senderData = senderDataTab.getDataRange().getValues();
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
    // Get merge data with flexible column names -
    // the script allows the user to put their own
    // column names in the spreadsheet and then to
    // reference them in the merge document
    const data = mailMergeTab.getDataRange().getValues();
    var columnHeaders = [];
    columnHeaders = senderData[0]?.concat(data[0]) ?? data[0];

  // Filter out header row
  const mergeData = data.slice(1);

  // Process each merge record
  for (let i = 0; i < mergeData.length; i++) {
    const record = senderData && senderData[1] ? senderData[1].concat(mergeData[i]) : mergeData[i];
    const toMergeData = {};

    // Map merge fields dynamically based on headers
    for (let j = 0; j < columnHeaders.length; j++) {
      toMergeData[columnHeaders[j]] = record[j] ?? ""; // Use ?? for default empty string
    }

    // Fill in additional data
    toMergeData["date"] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy MMMM dd");

    // Perform the merge
    mergeTemplate(mergeDoc, toMergeData);
  }

    // Display success message with link
    // substitute %s with strings for name and url
  const successMessage = Utilities.formatString(
        'Merged letter %s: %s',
        templateName + ' - ' + finishedFileName,
        DriveApp.getFileById(mergeDocId).getUrl()
      );
  Logger.log(successMessage);
  // Rename the file
  DriveApp.getFileById(mergeDocId).setName(templateName + ' - ' + finishedFileName);
  } catch (error) { // Outer catch block for overall errors
    console.error(`An error occurred during mail merge: ${error}`);
    // Implement appropriate error handling and user feedback
  }
}
