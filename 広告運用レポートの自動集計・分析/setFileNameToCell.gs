function setFileNameToCell() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var fileName = spreadsheet.getName();
  spreadsheet.getSheetByName('01_DATA_FILTERD').getRange('B2').setValue(fileName);
}
