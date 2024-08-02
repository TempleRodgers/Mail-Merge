/**
 * 12/1/2024: Temple Rodgers
 * Creates a table with multiple columns per row and populates it with data from a spreadsheet.
 */
function onInstall(e) {
  onOpen(e);
}

function onOpen(e)
{
  // need to use createAddonMenu because this will be an extension
  var menu = DocumentApp.getUi().createAddonMenu();
  menu.addItem('Run label address merge', 'showSheetPickerDialog');
  menu.addToUi(); 
}

// global variable for the selected sheet
var selectedSheetUrl;
var selectedColumns;

function showSheetPickerDialog() {
  // Create a custom dialog box with a picker to select a Google Sheet from the current folder
  var htmlOutput = HtmlService.createHtmlOutputFromFile('sheetPickerDialog')
      .setWidth(400)
      .setHeight(500);
  
  DocumentApp.getUi().showModalDialog(htmlOutput, 'Select Google Sheet with the addresses');
}
function setSelectedSheetUrlAndColumns(selectedSheetUrl,selectedColumns, selectedBorderColor, selectedBorderWidth, selectedRowHeight) {
  // Set the selected Google Sheet URL to the global variable
  createMultiColumnTable(selectedSheetUrl,selectedColumns, selectedBorderColor, selectedBorderWidth, selectedRowHeight);
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

function selectSheet() {
  // Display the spinner when the button is clicked
  document.querySelector('.spinner-border').classList.remove('d-none');

  // Your existing logic for selecting the sheet goes here

  // After completing the operation, hide the spinner
  document.querySelector('.spinner-border').classList.add('d-none');
}

function createMultiColumnTable(selectedSheetUrl,selectedColumns, selectedBorderColor, selectedBorderWidth, selectedRowHeight) {
  try {
    // Open the Google Docs template by its ID
    const document = DocumentApp.getActiveDocument();
    const tableColumns = selectedColumns; // number of columns in the table, will be an input variable
    const rowHeight = selectedRowHeight * 72 / 2.54; // row height in cm * 72 / 2.54 to get points
    const borderWidth = selectedBorderWidth; //table border width
    const borderColour = selectedBorderColor; //table border colour
    console.log("details: "+tableColumns+", "+rowHeight+", "+borderWidth+", "+borderColour);

    // Create a body section in the document
    var body = document.getBody();
    // clear the document
    body.clear();

    // Create a table with the specified number of columns
    const table = body.insertTable(0).setBorderWidth(borderWidth); // Insert table at the beginning of the document
    table.setBorderColor(borderColour);

    // Move the first paragraph (if it exists) after the table
    const firstElement = body.getChild(0);
    if (firstElement && firstElement.getType() !== DocumentApp.ElementType.TABLE) {
      firstElement.moveTo(table.getChild(table.getNumChildren()));
    }

    // Open the spreadsheet by its ID
    const sheet = SpreadsheetApp.openByUrl(selectedSheetUrl);
    const mergeTab = sheet.getSheetByName('Label_Merge');

    // Get the data from the spreadsheet, excluding the first row (titles)
    const data = mergeTab.getRange(2, 1, mergeTab.getLastRow() - 1, mergeTab.getLastColumn()).getValues();

    // Populate the table with data from the spreadsheet
    for (let i = 0; i < data.length; i++) {
      const columnIndex = i % tableColumns; // Calculate the column index based on the current row

      // If it's a new row in the table, add a new row of the specified minimum height
      if (columnIndex === 0) {
        var currentRow = table.appendTableRow();
        currentRow.setMinimumHeight(rowHeight);
      }

      // Add a cell to the current row and set its text
      const cell = currentRow.appendTableCell();
      
      // Concatenate the data from each column in the row into a single cell
      let currentCellText = "";
      for (let k = 0; k < data[i].length; k++) {
        currentCellText += `${data[i][k]} \n`;
      }
      
      cell.setText(currentCellText.trim()); // Use trim() to remove trailing whitespace
    }
  } catch (error) {
    console.error(`An error occurred: ${error}`);
  }
}
