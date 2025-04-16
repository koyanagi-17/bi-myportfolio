function setFileNameToCell() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var fileName = spreadsheet.getName();
  spreadsheet.getSheetByName('data_imp').getRange('B2').setValue(fileName);
}
