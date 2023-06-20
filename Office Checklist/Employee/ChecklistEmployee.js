function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('MTS Tools')
      .addItem('Create Employee Sheets', 'createEmployeeSheets')
      .addToUi();
  }
  
  function createEmployeeSheets() {
    var adminSpreadsheetId = '1KrA4rZhdOTNH_F9Pqil-RW9uNftF4UKvXLvUrQgdtgY'; // ID of the Admin checklist spreadsheet
    var adminSheetName = 'Sheet1'; // Name of the Admin checklist sheet
  
    var adminSpreadsheet = SpreadsheetApp.openById(adminSpreadsheetId);
    var adminSheet = adminSpreadsheet.getSheetByName(adminSheetName);
    var employeeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
    // Clear existing employee sheets
    var employeeSheets = employeeSpreadsheet.getSheets();
    for (var i = employeeSheets.length - 1; i >= 1; i--) {
      employeeSpreadsheet.deleteSheet(employeeSheets[i]);
    }
  
    // Get unique employee names from the admin sheet
    var employeeNames = [...new Set(adminSheet.getRange(3, 2, adminSheet.getLastRow() - 2, 1).getValues().flat())];
  
    // Create separate sheet for each employee
    for (var i = 0; i < employeeNames.length; i++) {
      var employeeName = employeeNames[i];
      var employeeSheet = employeeSpreadsheet.insertSheet(employeeName);
  
      // Filter tasks for the current employee
      var filteredTasks = adminSheet.getRange(3, 1, adminSheet.getLastRow() - 2, adminSheet.getLastColumn())
        .getValues()
        .filter(function (row) {
          return row[1] === employeeName && row.slice(2).some(function (task) { return task !== ''; }); // Exclude rows without any task
        });
  
      // Get task names from the filtered tasks
      var taskNames = filteredTasks.length > 0 ? filteredTasks[0].slice(2) : [];
  
      // Set headers in employee sheet
      var headers = [''].concat(taskNames);
      employeeSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
      // Copy filtered tasks to employee sheet
      if (filteredTasks.length > 0) {
        var taskData = filteredTasks.map(function (row) {
          return [''].concat(row.slice(2));
        });
        employeeSheet.getRange(2, 1, taskData.length, taskData[0].length).setValues(taskData);
      }
  
      // Apply formatting to employee sheet
      var range = employeeSheet.getRange(1, 1, filteredTasks.length + 1, taskNames.length + 1);
      range.setHorizontalAlignment('center');
      range.setVerticalAlignment('middle');
    }
  }