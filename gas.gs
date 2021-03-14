function addRow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(['Cotton Sweatshirt XL', 'css004']);
}

function helloWorld() {
  Browser.msgBox('Hello world!');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Bryden Menu')
    .addItem('Say Hello','helloWorld')
    .addItem('Clear Existing','clearsheet_prompt')
    .addToUi()
}

function clearInvoice() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var invoiceNumber = sheet.getRange("B5").clearContent();
  var invoiceAmount = sheet.getRange("B8").clearContent();
  var invoiceTo = sheet.getRange("E5").clearContent();
  var invoiceFrom = sheet.getRange("E6").clearContent(); 
}

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

function pasteAtoG() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:G').activate();
  spreadsheet.getRange('B:B').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

function clearsheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var range = spreadsheet.getRange("Existing!D2:E");
  range.activate();
  range.clear();
}

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

// function
// var ui = DocumentApp.getUi();
// var response = ui.prompt('Getting to know you', 'May I know your name?', ui.ButtonSet.YES_NO);

// // Process the user's response.
// if (response.getSelectedButton() == ui.Button.YES) {
//   Logger.log('The user\'s name is %s.', response.getResponseText());
// } else if (response.getSelectedButton() == ui.Button.NO) {
//   Logger.log('The user didn\'t want to provide a name.');
// } else {
//   Logger.log('The user clicked the close button in the dialog\'s title bar.');
// }



function clear_colb() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B2:C').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, commentsOnly: true, skipFilteredRows: true});
};
