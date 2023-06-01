// Select a range, then click "Find and Select" from the menu.
// It will select any cells within your range with the search value provided.
// It's similar to "Find and Replace All", but allows for cell formatting.
function findAndSelect() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    var ui = SpreadsheetApp.getUi();
    var searchValue = ui.prompt("Enter the search value").getResponseText();
    var searchRange = sheet.getActiveRange();

    var rangeValues = searchRange.getValues();
    var matchingCells = [];

    for (var i = 0; i < rangeValues.length; i++) {
        for (var j = 0; j < rangeValues[i].length; j++) {
            if (rangeValues[i][j] == searchValue) {
                matchingCells.push(searchRange.getCell(i + 1, j + 1).getA1Notation());
            }
        }
    }

    SpreadsheetApp.setActiveRangeList(sheet.getRangeList(matchingCells));
}


//Function to add findAndSelect to the Google Sheets menu.
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Tools')
      .addItem('Find and Select', 'findAndSelect')
      .addToUi();
  }