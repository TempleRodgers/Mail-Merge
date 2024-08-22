/**
 * Temple Rodgers - 16/8/24
 * Mail merge, getting data from a selected spreadsheet
 * which contains sender data on one tab and merge data
 * on another
 * this video is useful: https://www.youtube.com/watch?v=QNPPEB64QbI&t=1625s
 * also, the document classes and functions https://developers.google.com/apps-script/reference/document/
 *  
 */
function performBodyMerge(spreadsheetURL) {
//function performBodyMerge() {
//  const spreadsheetURL = "https://docs.google.com/spreadsheets/d/158Md3meKiyZAO2aXj5qnQaosCRU4Fp7R_Ecss7gsrr0/edit?usp=drivesdk";
//  const spreadsheetURL = "https://docs.google.com/spreadsheets/d/1UkipnRBM0xPMCAu8bbKYjAnIt1FBv__jxzhzB3hVbyk/edit?usp=drivesdk";
//
  const template = DocumentApp.getActiveDocument(),   // returns a type Document, 
                                                      // which is the current document 
                                                      // being used as a template
    templateId = template.getId(),                    // returns the ID of the document
    templateName = template.getName(),                // returns the name of the document
    templateFile = DriveApp.getFileById(templateId),  // load the templateFile info,
    // get the template body to use in the merge
    templateBody = template.getBody();

  var mergeDocFile = null, // destination file info
    mergeDocId = null, // destination doc ID
    mergeDoc = null, // destination document
    mergeDocBody = null,
    finishedFileName = "finished testing file",
    // set the mail merge spreadsheet variables
    // the script gathers merge data from two
    // tabs in the spreadsheet: Mail_Merge and
    // Sender_Details    
    sheet = null,
    mailMergeTab = null,
    senderDataTab = null;

  try {
    // Show the progress popup
    var progressPopup = HtmlService.createHtmlOutputFromFile('ProgressPopup')
      .setWidth(300)
      .setHeight(200);
    DocumentApp.getUi().showModalDialog(progressPopup, 'Mail Merge Progress');

    // Open the spreadsheet and get sheets
    sheet = SpreadsheetApp.openByUrl(spreadsheetURL); // use selectedSheetURL
//    Logger.log("sheet value = " + sheet);
    mailMergeTab = sheet.getSheetByName("Mail_Merge"); //get the tab

    if (!mailMergeTab) {
      throw new Error('Sheet named "Mail_Merge" not found.');
    }

    // Sender Details (optional)
    const senderDataTab = sheet.getSheetByName(
      "Sender_Details"
    );
    let senderData = [];
    if (senderDataTab) {
      senderData = senderDataTab.getDataRange().getValues();
    } else {
      Logger.log(
        'Sheet named "Sender_Details" not found, not using sender data.'
      );
    }

    // Get merge data and put it into the `data` array
    const mailMergeData = mailMergeTab.getDataRange().getValues();
    // Use the first row (row 0) as the column headers and add the sender data so it's in one array
    const columnHeaders = senderData[0]
      ? senderData[0].concat(mailMergeData[0])
      : mailMergeData[0];
    // Then slice of the chunk of data for the mail merge excluding headers
    const mergeData = mailMergeData.slice(1);

    // now construct a set of merge data substitutions for each row, one by one and
    // call mergeTemplate to add the merged data to the merge document
    try {
        // copy the template and give it a temporary name (to be replaced later)
        mergeDocFile = templateFile.makeCopy(`${templateName} - ${finishedFileName}`);

        // get the ID of the file just created
        mergeDocId = mergeDocFile.getId();
        // open the document just created using ID
        mergeDoc = DocumentApp.openById(mergeDocId);
        mergeDocBody = mergeDoc.getBody();
        mergeDocBody.clear();

        // Process each merge record
        mergeData.forEach((record, i) => {
          const recordData = senderData.length > 1 ? senderData[1].concat(record) : record;
          const toMergeData = {};
          // *** Update the progress popup ***
          updateProgress(i + 1, mergeData.length); 

          // Map merge fields
          columnHeaders.forEach((header, j) => {
            toMergeData[header] = recordData[j] || "";
          });

          // Add additional data
          toMergeData["date"] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy MMMM dd");
/** for troubleshooting
          // Log values before calling mergeTemplate
          console.log("templateBody:", templateBody.getText()); // Log the entire template body content
          console.log("toMergeData:", toMergeData);
          console.log("mergeDocBody:", mergeDocBody.getText()); // Log the current state of mergeDocBody
*/
          // Perform the merge
          // templateBody has the merge document information to be used in the merge
          // toMergeData is the row of data that has to be merged into the template
          // mergeDoc is the actual output merge document
          // Create a fresh copy of the temporaryBody for each record
          let temporaryBody = templateBody.copy(); 
          mergeTemplate(temporaryBody, mergeDocBody, toMergeData);

          // Add a page break after each record (except the last one)
          if (i < mergeData.length - 1) {
            mergeDocBody.appendPageBreak();
          }
        });

    } catch (error) {
        console.error(`An error occurred: ${error}`);
    }

    // Save the changes to the output document
    mergeDoc.saveAndClose();

    // Display success message
    Logger.log(`Merged letter ${templateName} - ${finishedFileName}: ${mergeDoc.getUrl()}`);
  } catch (error) {
    console.error("An error occurred during mail merge: " + error);
  }
}
