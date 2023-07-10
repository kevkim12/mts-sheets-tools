function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var checkbox = range.getDataValidation();
    
    if (checkbox != null && checkbox.getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      var isChecked = range.isChecked();
      var cellColor = isChecked ? '#b6d7a8' : '#ea9999';
      range.setBackground(cellColor);
    }
  }