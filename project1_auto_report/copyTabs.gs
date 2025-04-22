/**
 * この関数は、スプレッドシート内のシートをコピーし、名前を次の月に更新するためのものです。
 * 処理の流れは次の通りです：
 * 1. 現在の日付を基に、現在の月と翌月を計算
 * 2. 現在の月を含むシート名を持つシートを探し、そのシートをコピー
 * 3. コピーしたシートに翌月の日付を設定
 */

function copyTabs() {
  // 現在のスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // スプレッドシート内のすべてのシートを取得
  var sheets = ss.getSheets();
  
  // 現在の日付を取得
  var currentDate = new Date();
  
  // 現在の月（yyyy-MM形式）を取得
  var currentMonth = Utilities.formatDate(currentDate, "JST", "yyyy-MM");
  
  // 翌月の月（yyyy-MM形式）を取得
  var nextMonth = Utilities.formatDate(new Date(currentDate.getFullYear(), currentDate.getMonth()+1, currentDate.getDate()), "JST", "yyyy-MM");
  
  // 翌月の1日の日付（yyyy-MM-dd形式）を取得
  var nextMonthFirstDay = Utilities.formatDate(new Date(currentDate.getFullYear(), currentDate.getMonth()+1, 1), "JST", "yyyy-MM-dd");

  // すべてのシートをループして処理
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();

    // シート名が現在の月（currentMonth_）で始まるシートを探す
    if (sheetName.indexOf(currentMonth + "_") === 0) {
      // 新しいシート名を作成（次の月に更新）
      var newName = nextMonth + "_" + sheetName.substring(8);
      
      // シートをコピーし、スプレッドシートに追加
      var newSheet = sheet.copyTo(ss);
      
      // コピーしたシートに新しい名前を付ける
      newSheet.setName(newName);
      
      // 新しいシートのA1セルに、翌月の1日の日付を設定
      newSheet.getRange("A1").setValue(nextMonthFirstDay);
    }
  }
}
