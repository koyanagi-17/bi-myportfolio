function copyFunctionCells() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("合計_媒体別");
  var range = sheet.getRange("H16:381");
  var formulas = range.getFormulas();

  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j] == "") {
        formulas[i][j] = "";
      }
    }
  }
  range.setFormulas(formulas);
}
