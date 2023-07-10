function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var checkbox = range.getDataValidation();
    
    if (checkbox != null && checkbox.getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      var isChecked = range.isChecked();
      var cellColor = isChecked ? 'green' : 'red';
      range.setBackground(cellColor);
    }
  }