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
  
      // Get task names from the admin sheet
      var taskNames = adminSheet.getRange(3, 3, 1, adminSheet.getLastColumn() - 2).getValues()[0];
  
      // Set headers in employee sheet
      var headers = [taskNames];
      employeeSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  
      // Filter tasks for the current employee
      var tasks = adminSheet.getRange(3, 3, adminSheet.getLastRow() - 2, adminSheet.getLastColumn() - 2)
        .getValues()
        .filter(function (row) {
          return row[0] === employeeName && row.slice(2).some(function (task) { return task !== ''; }); // Exclude rows without any task
        });
  
      // Copy filtered tasks to employee sheet
      if (tasks.length > 0) {
        employeeSheet.getRange(2, 1, tasks.length, 1).setValues(tasks.map(function (row) {
          return [''];
        }));
      }
  
      // Apply formatting to employee sheet
      var range = employeeSheet.getRange('A1:A' + (tasks.length + 1));
      range.setHorizontalAlignment('center');
      range.setVerticalAlignment('middle');
    }
  }