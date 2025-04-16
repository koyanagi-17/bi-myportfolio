function copyTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var currentDate = new Date();
  var currentMonth = Utilities.formatDate(currentDate, "JST", "yyyy-MM");
  var nextMonth = Utilities.formatDate(new Date(currentDate.getFullYear(), currentDate.getMonth()+1, currentDate.getDate()), "JST", "yyyy-MM")
  var nextMonthFirstDay = Utilities.formatDate(new Date(currentDate.getFullYear(), currentDate.getMonth()+1, 1), "JST", "yyyy-MM-dd");

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();

    if (sheetName.indexOf(currentMonth+"_") === 0) {
      var newName = nextMonth + "_" + sheetName.substring(8);
      var newSheet = sheet.copyTo(ss);
      newSheet.setName(newName);
      newSheet.getRange("A1").setValue(nextMonthFirstDay);
    }
  }
}
