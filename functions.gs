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


// Select a range, then click "Color Rows" from the menu.
// It will color the rows in a banded mannner, but with user-defined columns that must be different
// to alternate colors.
function colorRows() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var selectedRange = sheet.getActiveRange();
    var identicalColumns = Browser.inputBox('Identical Columns', 'Enter comma-separated column numbers for identical comparison (e.g., 2,4,6):', Browser.Buttons.OK_CANCEL);
    selectedRange.setBackground('white')
    if (identicalColumns === 'cancel') {
      return; // Exit the function if the user clicked "Cancel"
    }
  
    var columnsToMatch = identicalColumns.split(',').map(Number);
  
    var numRows = selectedRange.getNumRows();
    var values = selectedRange.getValues();
  
    var currentColor = '#CCC';
  
    for (var col = 1; col <= selectedRange.getNumColumns(); col++) {
        selectedRange.getCell(1, col).setBackground(currentColor);
      }    

    for (var i = 1; i < numRows; i++) {
        var previousRow = values[i-1]; 
        var row = values[i];
  
      // Check if the specified columns have identical values
      var identical = true;
      for (var j = 0; j < columnsToMatch.length; j++) {
        if (row[columnsToMatch[j] - 1].toString() != previousRow[columnsToMatch[j] - 1].toString()) {
          identical = false;
          break;
        }
      }

      if(!identical){
        currentColor = (currentColor === '#CCC') ? '#EEE' : '#CCC';
      }
      for (var col = 1; col <= selectedRange.getNumColumns(); col++) {
          selectedRange.getCell(i+1, col).setBackground(currentColor);
      }      
        
    }
  }

//Function to add findAndSelect to the Google Sheets menu.
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Tools')
      .addItem('Find and Select', 'findAndSelect')
      .addItem('Color Rows', 'colorRows')
      .addToUi();
  }
