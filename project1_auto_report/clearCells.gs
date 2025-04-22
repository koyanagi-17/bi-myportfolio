/**
 * この関数は、翌月のシートにある特定のセル範囲の内容をクリアするためのものです。
 * 処理の流れは次の通りです：
 * 1. 次の月を計算し、次の月を含むシートを探す
 * 2. 各シート内で、指定されたセル範囲（T10:T40）の背景色を確認
 * 3. 背景色が指定の色（#fff2cc）であるセルの内容をクリアする
 */

function clearCells() {
  // 現在のスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // スプレッドシート内のすべてのシートを取得
  var sheets = ss.getSheets();
  
  // 現在の日付を取得
  var currentDate = new Date();
  
  // 翌月の日付（yyyy-MM形式）を取得
  var nextMonth = Utilities.formatDate(new Date(currentDate.getFullYear(), currentDate.getMonth()+1, currentDate.getDate()), "JST", "yyyy-MM");

  // すべてのシートをループして処理
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();

    // シート名が翌月（nextMonth_）で始まるシートを探す
    if (sheetName.indexOf(nextMonth + "_") === 0) {
      // T10:T40 の範囲を取得
      var range = sheet.getRange("T10:40");
      
      // 範囲内のセルの背景色を取得
      var bgColors = range.getBackgrounds();
      
      // セルの背景色が指定の色（#fff2cc）であれば、そのセルの内容をクリア
      for (var j = 0; j < bgColors.length; j++) {
        for (var k = 0; k < bgColors[j].length; k++) {
          if (bgColors[j][k] == "#fff2cc") {
            // 背景色が #fff2cc のセルの内容をクリア
            range.getCell(j + 1, k + 1).clearContent();
          }
        }
      }
    }
  }
}
