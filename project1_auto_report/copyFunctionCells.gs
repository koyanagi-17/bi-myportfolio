/**
 * この関数は、「02_MONTHLY_SUMMARY」シートの特定の範囲（H16:H381）の数式を取得し、
 * その数式が空である場合に空のままにして、新しい数式を設定するためのものです。
 * 
 * 処理の流れは次の通りです：
 * 1. 「02_MONTHLY_SUMMARY」シートを取得
 * 2. H16:H381 の範囲を取得し、その範囲内の数式を取得
 * 3. 取得した数式が空でないか確認し、空であればそのまま空のままにする
 * 4. 数式をシートに再設定
 */

function copyFunctionCells() {
  // 現在のスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 「02_MONTHLY_SUMMARY」シートを取得
  var sheet = ss.getSheetByName("02_MONTHLY_SUMMARY");
  
  // H16:H381 の範囲を取得
  var range = sheet.getRange("H16:381");
  
  // 範囲内の数式を取得
  var formulas = range.getFormulas();

  // 数式が空でないか確認し、空であればそのまま空にする
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      if (formulas[i][j] == "") {
        // 数式が空の場合、そのまま空にする
        formulas[i][j] = "";
      }
    }
  }

  // 変更した数式をシートに設定
  range.setFormulas(formulas);
}
