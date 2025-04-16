function MoveActiveSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mgws = ss.getSheetByName("index");
  var manageData = mgws.getRange("B1:B").getValues().filter(line => line[0] != "");
  var sheets = ss.getSheets();
  var sheetnames = sheets.map(sheet => {
    return [sheet.getSheetName(), sheet.isSheetHidden()];
  });
  
  for (let i = 0; i < manageData.length; i++) {
    if (sheetnames[i][0] != manageData[i][0]) {
      var ws = ss.getSheetByName(manageData[i][0]);

      if (ws) {
        ws.activate();
        ss.moveActiveSheet(i + 1);
      }
    }
  }
  mgws.activate();
  ss.moveActiveSheet(1);
}
