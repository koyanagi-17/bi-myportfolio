function clearCells() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var currentDate = new Date();
  var nextMonth = Utilities.formatDate(new Date(currentDate.getFullYear(), currentDate.getMonth()+1, currentDate.getDate()), "JST", "yyyy-MM");

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    if (sheetName.indexOf(nextMonth+"_") === 0) {
      var range = sheet.getRange("T10:40");
      var bgColors = range.getBackgrounds();
      
      for (var j = 0; j < bgColors.length; j++) {
        for (var k = 0; k < bgColors[j].length; k++) {
          if (bgColors[j][k] == "#fff2cc") {
            range.getCell(j + 1, k + 1).clearContent();
          }
        }
      }
    }
  }
}
