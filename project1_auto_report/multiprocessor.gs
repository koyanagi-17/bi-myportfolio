/**
 * この関数は、複数の処理を一度に実行するためのものです。
 * 実行する処理は、以下の関数を順番に呼び出すことです：
 * 1. copyTabs() - シートをコピーする
 * 2. clearCells() - セルの内容をクリアする
 * 3. copyFunctionCells() - 特定のセルの内容をコピーする
 * 4. getSheetName() - シート名を取得する
 * 5. MoveActiveSheet() - シートの順番を並べ替える
 */
function multiprocessor_1() {
  copyTabs();        // シートをコピー
  clearCells();      // セルの内容をクリア
  copyFunctionCells(); // 特定のセルの内容をコピー
  getSheetName();    // シート名を取得
  MoveActiveSheet(); // シートの順番を並べ替え
}

/**
 * この関数は、multiprocessor_1と似ていますが、処理内容が少し異なります。
 * 実行する処理は、以下の関数を順番に呼び出すことです：
 * 1. copyFunctionCells() - 特定のセルの内容をコピーする
 * 2. getSheetName() - シート名を取得する
 * 3. MoveActiveSheet() - シートの順番を並べ替える
 */
function multiprocessor_2() {
  copyFunctionCells(); // 特定のセルの内容をコピー
  getSheetName();      // シート名を取得
  MoveActiveSheet();   // シートの順番を並べ替え
}
