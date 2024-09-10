// Array para definir que card/question de Metabase seran cargadas
var menuItems = [
  { sheetName: 'Sheet1', cardId: '181' },
  { sheetName: 'Sheet2', cardId: '240' }
];

// Function to create the menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Get Data');
  
  // Loop through the menuItems array to dynamically add menu items
  menuItems.forEach(function(item, index) {
    menu.addItem(`Load Question ${item.cardId} on Sheet ${item.sheetName}`, `runHardcodedSheetAndCard${index}`);
  });
  
  // Add the menu to the UI
  menu.addItem('Select Sheet & Question', 'showInputForm')
      .addToUi();
}

// Create dynamic functions for each sheet and card pair
menuItems.forEach(function(item, index) {
  this[`runHardcodedSheetAndCard${index}`] = function() {
    runHardcodedSheetAndCard(item.cardId, item.sheetName);
  };
});

// Function to handle loading data into sheets
function runHardcodedSheetAndCard(cardId, sheetName) {
  try {
    // Use the passed parameters for sheetName and cardId
    MetabaseGoogleSheetConnector.fetchDataAndFillSheet(sheetName, cardId);
    
    // Notify the user about the success
    SpreadsheetApp.getUi().alert(`Data successfully fetched and filled in the sheet: ${sheetName}`);
  } catch (error) {
    // Notify the user about any error
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

function showInputForm() {
  // Get the HTML content from the library
  var htmlContent = MetabaseGoogleSheetConnector.getHtmlForm();
  
  // Serve the HTML as a modal dialog
  var htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(400)
      .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enter Data');
}

// Template function for time-based triggers (with hardcoded sheetName and cardId)
function runTimeTriggeredSheetAndCard() {
  try {
    // Define the sheet name and card ID for this specific function
    var sheetName = 'Sheet1';  // Replace with your desired sheet name
    var cardId = '181';  // Replace with your desired card ID

    // Call the library function to fetch data and fill the sheet
    MetabaseGoogleSheetConnector.fetchDataAndFillSheet(sheetName, cardId);

    // Log the success (viewable in the "Execution Logs")
    Logger.log(`Data successfully fetched and filled in the sheet: ${sheetName}`);
  } catch (error) {
    // Log the error (viewable in the "Execution Logs")
    Logger.log(`Error: ${error.message}`);
  }
}
