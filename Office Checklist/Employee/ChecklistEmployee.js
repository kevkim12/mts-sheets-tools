function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('MTS Tools')
      .addItem('Create Employee Sheets', 'createEmployeeSheets')
      .addItem('Upload Checklist', 'uploadChecklist')
      .addToUi();
  }
  
  function uploadChecklist() {
    var adminSpreadsheetId = '1KrA4rZhdOTNH_F9Pqil-RW9uNftF4UKvXLvUrQgdtgY'; // ID of the Admin checklist spreadsheet
    var adminSheetName = 'Sheet1'; // Name of the Admin checklist sheet
  
    var adminSpreadsheet = SpreadsheetApp.openById(adminSpreadsheetId);
    var adminSheet = adminSpreadsheet.getSheetByName(adminSheetName);
    var employeeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var employeeSheets = employeeSpreadsheet.getSheets();
  
    for (var i = 1; i < employeeSheets.length; i++) {
      var employeeSheet = employeeSheets[i];
      var employeeName = employeeSheet.getName();
  
      var employeeNameRange = adminSheet.getRange('B:B');
      var employeeNameValues = employeeNameRange.getValues();
      var employeeNameRowIndex = -1;
  
      for (var row = 0; row < employeeNameValues.length; row++) {
        if (employeeNameValues[row][0] === employeeName) {
          employeeNameRowIndex = row;
          break;
        }
      }
  
      if (employeeNameRowIndex !== -1) {
        var checkboxesRange = employeeSheet.getRange(2, 1, 1, employeeSheet.getLastColumn());
        var checkboxValues = checkboxesRange.getValues()[0];
  
        for (var col = 0; col < checkboxValues.length; col++) {
          var checkboxValue = checkboxValues[col];
          var taskCell = adminSheet.getRange(employeeNameRowIndex + 1, col + 3);
  
          if (checkboxValue === true) {
            taskCell.setBackground('#b6d7a8');
          } else {
            taskCell.setBackground('#ea9999');
          }
        }
      }
    }
  }
  
  function createEmployeeSheets() {
    var adminSpreadsheetId = '1KrA4rZhdOTNH_F9Pqil-RW9uNftF4UKvXLvUrQgdtgY'; // ID of the Admin checklist spreadsheet
    var adminSheetName = 'Sheet1'; // Name of the Admin checklist sheet
  
    var adminSpreadsheet = SpreadsheetApp.openById(adminSpreadsheetId);
    var adminSheet = adminSpreadsheet.getSheetByName(adminSheetName);
    var employeeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
    var employeeSheets = employeeSpreadsheet.getSheets();
    for (var i = employeeSheets.length - 1; i >= 1; i--) {
      employeeSpreadsheet.deleteSheet(employeeSheets[i]);
    }
  
    var employeeNames = adminSheet
      .getRange('B2:B' + adminSheet.getLastRow())
      .getValues()
      .flat()
      .filter(function (name) {
        return name !== '';
      });
  
    for (var i = 0; i < employeeNames.length; i++) {
      var employeeName = employeeNames[i];
      var employeeSheet = employeeSpreadsheet.insertSheet(employeeName);
  
      var filteredTasks = adminSheet.getRange(2, 1, adminSheet.getLastRow() - 1, adminSheet.getLastColumn())
        .getValues()
        .filter(function (row) {
          return row[1] === employeeName && row.slice(2).some(function (task) { return task !== ''; });
        });
  
      var taskNames = Array.from(new Set(filteredTasks.flatMap(function (row) {
        return row.slice(2);
      })));
  
      var headers = taskNames;
      employeeSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
      if (filteredTasks.length > 0) {
        var taskData = filteredTasks.map(function (row) {
          return row.slice(2);
        });
        employeeSheet.getRange(1, 1, taskData.length, taskData[0].length).setValues(taskData);
      }
  
      var range = employeeSheet.getRange(1, 1, filteredTasks.length, taskNames.length);
      range.setHorizontalAlignment('center');
      range.setVerticalAlignment('middle');
  
      var checkboxesRange = employeeSheet.getRange(2, 1, taskData.length, taskData[0].length);
      var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
      checkboxesRange.setDataValidation(rule);
  
      var checkboxValues = checkboxesRange.getValues();
      for (var row = 0; row < checkboxValues.length; row++) {
        for (var col = 0; col < checkboxValues[row].length; col++) {
          var cell = checkboxesRange.getCell(row + 1, col + 1);
          if (checkboxValues[row][col] === true) {
            cell.setBackground('#b6d7a8');
          } else {
           cell.setBackground('#ea9999');
          }
        }
      }
    }
  }