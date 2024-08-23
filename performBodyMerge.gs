/**
 * Temple Rodgers - 23/8/24
 * Mail merge, getting data from a selected spreadsheet
 * which contains sender data on one tab and merge data
 * on another
 * this video is useful: https://www.youtube.com/watch?v=QNPPEB64QbI&t=1625s
 * also, the document classes and functions https://developers.google.com/apps-script/reference/document/
 *  
 */
function performBodyMerge() {
//  const spreadsheetURL = "https://docs.google.com/spreadsheets/d/158Md3meKiyZAO2aXj5qnQaosCRU4Fp7R_Ecss7gsrr0/edit?usp=drivesdk";
  const spreadsheetURL = "https://docs.google.com/spreadsheets/d/1UkipnRBM0xPMCAu8bbKYjAnIt1FBv__jxzhzB3hVbyk/edit?usp=drivesdk";
//
// function performBodyMerge(spreadsheetURL) {
  resetProgress(); // Reset progress at the start
  // Update progress message for data gathering
  progress.total = -1;  // Mark as gathering data (pseudo-progress)
  updateProgress();      // Call the update immediately
  
  // Wait for a short delay to ensure the message is updated in the UI
  Utilities.sleep(500);

  const template = DocumentApp.getActiveDocument(),   // returns a type Document, 
                                                      // which is the current document 
                                                      // being used as a template
    templateId = template.getId(),                    // returns the ID of the document
    templateName = template.getName(),                // returns the name of the document
    templateFile = DriveApp.getFileById(templateId),  // load the templateFile info,
    // get the template body to use in the merge
    templateBody = template.getBody();

  let mergeDocFile, // destination file info
    mergeDocId, // destination doc ID
    mergeDoc, // destination document
    mergeDocBody;

    // set the mail merge spreadsheet variables
    // the script gathers merge data from two
    // tabs in the spreadsheet: Mail_Merge and
    // Sender_Details    

  try {
    // Open the spreadsheet and get sheets
    sheet = SpreadsheetApp.openByUrl(spreadsheetURL); // use selectedSheetURL
    //    Logger.log("sheet value = " + sheet);
    mailMergeTab = sheet.getSheetByName("Mail_Merge"); //get the tab
    if (!mailMergeTab) throw new Error('Sheet named "Mail_Merge" not found.');

    // Get merge data and put it into the `data` array
    const mailMergeData = mailMergeTab.getDataRange().getValues();

    // for progress tracker popup
    progress.total = mailMergeData.length - 1; // Exclude header row from total
    updateProgress();

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
    // Use the first row from senderData and mailMergeData (row 0) as the 
    // column headers and add the sender data so it's in one array
    const columnHeaders = mailMergeData[0] ? mailMergeData[0].concat(senderData[0]) : senderData[0];
    // Then slice of the chunk of data for the mail merge excluding headers
    const mergeData = mailMergeData.slice(1);

    // instantiate the global value of how many rows have been processed
    progress.total = mergeData.length; // Set total once data is gathered
    updateProgress(); // Update progress again

    // now construct a set of merge data substitutions for each row, one by one and
    // call mergeTemplate to add the merged data to the merge document
    // copy the template and give it a temporary name
        const date = new Date();
        const year = date.getFullYear();
        const month = ('0' + (date.getMonth() + 1)).slice(-2); // Add leading zero if needed
        const day = ('0' + date.getDate()).slice(-2); // Add leading zero if needed
        const hours = ('0' + date.getHours()).slice(-2); // Add leading zero if needed
        const minutes = ('0' + date.getMinutes()).slice(-2); // Add leading zero if needed

        const dateandtime = `${year}${month}${day} ${hours}:${minutes}`;

        let finishedFileName = `finished merge file - ${dateandtime}`;
        mergeDocFile = templateFile.makeCopy(`${templateName} - ${finishedFileName}`);

      // get the ID of the file just created
      mergeDocId = mergeDocFile.getId();
      // open the document just created using ID
      mergeDoc = DocumentApp.openById(mergeDocId);
      mergeDocBody = mergeDoc.getBody();
      mergeDocBody.clear();

      // Process each merge record
      mergeData.forEach((record, i) => {
        const recordData = senderData.length > 1 ? record.concat(senderData[1]) : record;
        // const recordData = senderData.length > 1 ? senderData[1].concat(record) : record;
        const toMergeData = {};
        // Map merge fields
        columnHeaders.forEach((header, j) => toMergeData[header] = recordData[j] || "");

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

          // Update global progress
          progress.processed++;
          updateProgress();

          // progress update
          SpreadsheetApp.flush(); // Ensure changes are saved to the spreadsheet

          // Add a page break after each record (except the last one)
          if (i < mergeData.length - 1) 
            mergeDocBody.appendPageBreak();
      });

    // Save the changes to the output document
    mergeDoc.saveAndClose();

    // Display success message
    Logger.log(`Merged letter ${templateName} - ${finishedFileName}: ${mergeDoc.getUrl()}`);
  } catch (error) {
    console.error("An error occurred during mail merge: " + error);
    // Signal completion of the merge to the client
  }
}
