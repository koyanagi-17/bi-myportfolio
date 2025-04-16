function getSheetName() { 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName("index");
  var url = ss.getUrl();
  var sheets = ss.getSheets();
  var links = [];
 
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name = sheet.getName();

    if (name !== "index") {
      var id = sheet.getSheetId();
      var sheetUrl = url + "#gid=" + id;
      var link = '=HYPERLINK("' + sheetUrl + '","' + name + '")';
      links.push([link]);
    }
  }
  var rows = sheets.length-1;
  outputSheet.getRange("A2:A").clearContent();
  outputSheet.getRange(2,1,rows,1).setValues(links);
  outputSheet.hideColumns(1);
}
