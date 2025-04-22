/**
 * この関数は、現在のスプレッドシートの名前を取得し、その名前を「01_DATA_FILTERD」シートの B2 セルに設定します。
 * 
 * 処理の流れは次の通りです：
 * 1. 現在のスプレッドシートを取得
 * 2. スプレッドシートの名前を取得
 * 3. 「01_DATA_FILTERD」シートの B2 セルにその名前を設定
 */
function setFileNameToCell() {
  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // スプレッドシートの名前を取得
  var fileName = spreadsheet.getName();
  
  // 「01_DATA_FILTERD」シートの B2 セルにスプレッドシートの名前を設定
  spreadsheet.getSheetByName('01_DATA_FILTERD').getRange('B2').setValue(fileName);
}
