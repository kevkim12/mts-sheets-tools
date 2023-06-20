function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var row = range.getRow();
  
    if (row == 2) {
      var checkboxCell = sheet.getRange(row, range.getColumn());
      var checkboxValue = checkboxCell.getValue();
  
      if (checkboxValue === true) {
        checkboxCell.setBackground("#b6d7a8");
      } else if (checkboxValue === false) {
        checkboxCell.setBackground("#ea9999");
      }
    }
  }