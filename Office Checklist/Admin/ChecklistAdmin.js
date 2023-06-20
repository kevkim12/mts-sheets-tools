function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Add Task', 'showInputDialog')
    .addToUi();
}

function showInputDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(300)
    .setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Add Task');
}

function addTask(employeeName, taskName) {
  var adminSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var employeeSheet = SpreadsheetApp.openById('1AFKLl6KPPzNzpGtPCHvZxBhtGJQHfrugVpAy7iL_ORU').getActiveSheet();

  var employeeColumn = adminSheet.getRange('B:B').getValues().flat();
  var employeeIndices = [];
  for (var i = 0; i < employeeColumn.length; i++) {
    if (employeeColumn[i] === employeeName) {
      employeeIndices.push(i + 1); // Adding 1 to start from row 1 instead of 0
    }
  }

  if (employeeIndices.length === 0) {
    var lastRow = adminSheet.getLastRow();
    if (lastRow > 0 && adminSheet.getRange(lastRow, 2).isBlank()) {
      // If the last row is empty, overwrite the employee name
      adminSheet.getRange(lastRow, 2).setValue(employeeName);
      employeeIndices.push(lastRow);
    } else {
      // If the last row is not empty, add a new row with the employee name
      adminSheet.appendRow([null, employeeName]);
      employeeIndices.push(lastRow + 1);
    }
    employeeSheet.getRange(lastRow + 1, 1).setValue(employeeName);
  }

  var lastColumn = adminSheet.getLastColumn();

  // Add date on the left of the employee name
  var date = new Date();
  for (var i = 0; i < employeeIndices.length; i++) {
    adminSheet.getRange(employeeIndices[i], 1).setValue(date);
  }

  // Add task name on top of the employee name
  var taskNumber = lastColumn - 2 + 1; // Calculate task number based on column index, starting from 1
  adminSheet.getRange(1, lastColumn + 1).setValue('Task ' + taskNumber);
  for (var i = 0; i < employeeIndices.length; i++) {
    adminSheet.getRange(employeeIndices[i], lastColumn + 1).setValue(taskName);
  }
}