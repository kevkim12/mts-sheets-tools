function doGet() {
  var htmlTemplate = HtmlService.createTemplateFromFile('dialog');
  var htmlOutput = htmlTemplate.evaluate();
  htmlOutput.setTitle('Part Number Entry');
  return htmlOutput;
}

function showDialog() {
  SpreadsheetApp.getUi().showModalDialog(doGet(), 'Part Number Entry');
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('MTS Tools')
    .addItem('Inventory Scan', 'openDialog')
    .addToUi();
}

function openDialog() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('dialog')
    .setWidth(300)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Inventory Scan');
}

function processForm(form) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var partNumber = form.partNumber;
  var option = form.option;
  var message = '';

  if (option === 'add') {
    var existingRow = findPartNumber(partNumber);
    if (existingRow) {
      var quantity = sheet.getRange('B' + existingRow).getValue();
      sheet.getRange('B' + existingRow).setValue(quantity + 1);
      message = 'Part number ' + partNumber + ' added successfully.';
      scrollToRow(existingRow); // Scroll to existing row with updated quantity
    } else {
      var lastRow = sheet.getLastRow();
      sheet.getRange('A' + (lastRow + 1)).setValue(partNumber);
      sheet.getRange('B' + (lastRow + 1)).setValue(1);
      message = 'Part number ' + partNumber + ' added successfully.';
      scrollToRow(lastRow + 1); // Scroll to newly added row
    }
  } else if (option === 'remove') {
    var existingRow = findPartNumber(partNumber);
    if (existingRow) {
      var quantity = sheet.getRange('B' + existingRow).getValue();
      if (quantity > 1) {
        sheet.getRange('B' + existingRow).setValue(quantity - 1);
        message = 'Quantity for part number ' + partNumber + ' reduced by 1.';
        scrollToRow(existingRow); // Scroll to existing row with updated quantity
      } else {
        sheet.deleteRow(existingRow);
        message = 'Part number ' + partNumber + ' removed successfully.';
        scrollToRow(existingRow); // Scroll to the row above the deleted row
      }
    } else {
      message = 'Part number ' + partNumber + ' does not exist.';
    }
  }

  return message;
}

function scrollToRow(row) {
  var file = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = file.getActiveSheet();
  var range = sheet.getRange(row, 1);
  sheet.setActiveRange(range);
}

function findPartNumber(partNumber) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getRange('A:B').getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === partNumber) {
      return i + 1;
    }
  }
  return null;
}