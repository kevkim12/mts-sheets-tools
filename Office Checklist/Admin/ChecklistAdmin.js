function addTask(employeeName, taskName) {
  var adminSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var employeeSheet = SpreadsheetApp.openById('1AFKLl6KPPzNzpGtPCHvZxBhtGJQHfrugVpAy7iL_ORU').getActiveSheet();

  var employeeColumn = adminSheet.getRange('B:B').getValues().flat();
  var employeeIndices = [];
  for (var i = 0; i < employeeColumn.length; i++) {
    if (employeeColumn[i] === employeeName) {
      employeeIndices.push(i + 1);
    }
  }

  if (employeeIndices.length === 0) {
    var lastRow = adminSheet.getLastRow();
    if (lastRow > 0 && adminSheet.getRange(lastRow, 2).isBlank()) {
      adminSheet.getRange(lastRow, 2).setValue(employeeName);
      employeeIndices.push(lastRow);
    } else {
      adminSheet.appendRow([null, employeeName]);
      employeeIndices.push(lastRow + 1);
    }
    employeeSheet.getRange(lastRow + 1, 1).setValue(employeeName);
  }

  var lastColumn = adminSheet.getLastColumn();

  var taskColumnIndex = findEmptyTaskColumn(adminSheet);
  if (taskColumnIndex === -1) {
    taskColumnIndex = lastColumn + 1;
    adminSheet.insertColumnAfter(lastColumn);
  }

  var date = new Date();
  for (var i = 0; i < employeeIndices.length; i++) {
    adminSheet.getRange(employeeIndices[i], 1).setValue(date);
  }

  var taskNumber = taskColumnIndex - 2;
  adminSheet.getRange(1, taskColumnIndex).setValue('Task ' + taskNumber);
  for (var i = 0; i < employeeIndices.length; i++) {
    adminSheet.getRange(employeeIndices[i], taskColumnIndex).setValue(taskName);
  }
}

function findEmptyTaskColumn(sheet) {
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 2; i <= headerRow.length; i++) {
    if (headerRow[i - 1] === '') {
      return i;
    }
  }
  return -1;
}