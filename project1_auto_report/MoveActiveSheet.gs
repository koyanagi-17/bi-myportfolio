/**
 * この関数は、指定されたシート（"00_INDEX"）に記載されたシート名に従って、
 * スプレッドシート内のシートを並べ替えるためのものです。
 * 具体的には、シートの順番を管理し、指定のシートがアクティブになるように移動します。
 * 
 * 主な処理：
 * 1. "00_INDEX"シートからシート名を取得
 * 2. そのシート名に従って、スプレッドシート内のシートを並べ替え
 * 3. 最後に、"00_INDEX"シートを最前面に戻します
 */

function MoveActiveSheet() {
  // 現在開いているスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // "00_INDEX"という名前のシートを取得（シート管理用）
  var mgws = ss.getSheetByName("00_INDEX");
  
  // "00_INDEX"シートのB列（B1:B）に記載されたシート名を取得し、空のセルを除外
  var manageData = mgws.getRange("B1:B").getValues().filter(line => line[0] != "");
  
  // スプレッドシート内に存在するすべてのシートを取得
  var sheets = ss.getSheets();
  
  // 各シートの名前と、そのシートが非表示かどうかを取得
  var sheetnames = sheets.map(sheet => {
    return [sheet.getSheetName(), sheet.isSheetHidden()];
  });
  
  // "00_INDEX"シートに記載されたシート名に基づいて並べ替えを実行
  for (let i = 0; i < manageData.length; i++) {
    // もし、現在のシートの名前が "00_INDEX"のシート名と異なれば、並べ替え処理を実行
    if (sheetnames[i][0] != manageData[i][0]) {
      // "00_INDEX"に記載されたシート名を持つシートが存在するか確認
      var ws = ss.getSheetByName(manageData[i][0]);
      
      if (ws) {
        // 対象のシートをアクティブ（選択状態）にする
        ws.activate();
        
        // アクティブなシートを指定された位置（i + 1）に移動
        ss.moveActiveSheet(i + 1);
      }
    }
  }
  
  // 最後に、"00_INDEX"シートを再度アクティブにし、1番目の位置に戻す
  mgws.activate();
  ss.moveActiveSheet(1);
}
