/**
 * この関数は、スプレッドシート内のすべてのシートの名前とリンクを取得し、「00_INDEX」シートにそのリンクを挿入します。
 * 
 * 処理の流れは次の通りです：
 * 1. 現在のスプレッドシートの URL を取得
 * 2. すべてのシートを取得
 * 3. 「00_INDEX」シート以外のシートの名前とリンクを作成
 * 4. 作成したリンクを「00_INDEX」シートの A列に設定
 * 5. 最後に、A列のリンクが表示されないようにその列を非表示に設定
 */

function getSheetName() { 
  // 現在のスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 「00_INDEX」シートを取得
  var outputSheet = ss.getSheetByName("00_INDEX");
  
  // スプレッドシートの URL を取得
  var url = ss.getUrl();
  
  // スプレッドシート内のすべてのシートを取得
  var sheets = ss.getSheets();
  
  // シートのリンクを格納する配列
  var links = [];
 
  // すべてのシートをループして処理
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name = sheet.getName();

    // 「00_INDEX」シートを除外して処理
    if (name !== "00_INDEX") {
      // シートの ID を取得
      var id = sheet.getSheetId();
      
      // シートの URL を作成
      var sheetUrl = url + "#gid=" + id;
      
      // HYPERLINK 関数を使ってリンクを作成
      var link = '=HYPERLINK("' + sheetUrl + '","' + name + '")';
      
      // 作成したリンクを配列に追加
      links.push([link]);
    }
  }

  // 「00_INDEX」シートの A2 以降の範囲をクリア
  var rows = sheets.length - 1;
  outputSheet.getRange("A2:A").clearContent();
  
  // 作成したリンクを「00_INDEX」シートに設定
  outputSheet.getRange(2, 1, rows, 1).setValues(links);
  
  // A列を非表示に設定
  outputSheet.hideColumns(1);
}
