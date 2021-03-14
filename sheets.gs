
function helloWorld() {
  Browser.msgBox('Hello world!');
}

//add a row of values to the bottom of the active google sheet
function addRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(['Cotton Sweatshirt XL', 'css004']);
}

//create a custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Bryden Menu')
    .addItem('Say Hello','helloWorld')
    .addItem('Clear Existing','clearsheet_prompt')
    .addToUi()
}

//clear contents for specific cells
function clearInvoice() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var invoiceNumber = sheet.getRange("B5").clearContent();
  var invoiceAmount = sheet.getRange("B8").clearContent();
  var invoiceTo = sheet.getRange("E5").clearContent();
  var invoiceFrom = sheet.getRange("E6").clearContent(); 
}

//populating column B with dummy values for testing
function populate_colb_dummy() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('Header');
  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().setValue('dummy_value');
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('B2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('C2').activate();
}

//testing - paste column B to A:G
function pasteAtoG() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:G').activate();
  spreadsheet.getRange('B:B').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

//clear contents for a specific sheet!range
function clearsheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var range = spreadsheet.getRange("Existing!D2:E");
  range.activate();
  range.clear();
}

//clear contents for specific sheet/range if user clicks 'YES' to prompt
function clearsheet_prompt() {
  // Prompt for user to verify before clearing cells
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Do you wish to clear contents on sheet ABC?', ui.ButtonSet.YES_NO);
 
  // Proceed only if 'OK' is selected by user
  if (response.getSelectedButton() == ui.Button.YES) {
    //then do this ;
  var spreadsheet = SpreadsheetApp.getActive();
  var range = spreadsheet.getRange("Existing!D2:E");
  range.activate();
  range.clear();
}
}

//testing - clear range B2:C on active sheet
function clear_colb() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B2:C').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, commentsOnly: true, skipFilteredRows: true});
};
