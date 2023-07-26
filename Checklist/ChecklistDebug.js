function resetCheckboxesAndDates() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataRange = sheet.getDataRange();
    var dataValues = dataRange.getValues();
    var now = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  
    for (var row = 0; row < dataValues.length; row++) {
      var checkboxColumns = [0, 4, 8];
      var dateColumns = [1, 5, 9];
  
      for (var col of checkboxColumns) {
        if (dataValues[row][col] === true) {
          sheet.getRange(row + 1, col + 1).setValue(false);
        }
      }
  
      for (var col of dateColumns) {
        var cellValue = dataValues[row][col];
        if (cellValue instanceof Date) {
          var cellDate = new Date(cellValue.getFullYear(), cellValue.getMonth(), cellValue.getDate(), 0, 0, 0);
          if (cellDate < today) {
            sheet.getRange(row + 1, col + 1).setValue(today);
          }
        }
      }
    }
  }
  
  function testResetCheckboxesAndDates() {
    resetCheckboxesAndDates();
  }
  
  function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
      .addItem('Test Reset Checkboxes and Dates', 'testResetCheckboxesAndDates')
      .addToUi();
  }