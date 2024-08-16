/**
 * Temple Rodgers - 16/08/24
 * Mail merge, getting data from a selected spreadsheet
 * which contains sender data on one tab and merge data
 * on another
 * this video is useful: https://www.youtube.com/watch?v=QNPPEB64QbI&t=1625s
 * 
 * FUNCTION TO ADD A MENU TO THE TEMPLATE DOCUMENT
 */
function onOpen(e)
{
  addMenu(); 
}

function addMenu()
{
  var menu = DocumentApp.getUi().createMenu('TEST 7 Spreadsheet mail merge');
  menu.addItem('Run Mail merge 7', 'showSheetPickerDialog');
  menu.addToUi(); 
}
