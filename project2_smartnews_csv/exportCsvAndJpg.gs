/**
 * このスクリプトは、Google スプレッドシートから特定のデータを取得し、 
 * それを元にSmartNews広告関連のCSVファイルと画像（.jpg）をGoogleドライブからダウンロードしてZIPファイルにまとめ、
 * 最後にそのZIPファイルを作成してダウンロードリンクを表示する機能を提供します。
 *
 * 主な処理：
 * 1. スプレッドシートのデータを読み取り、指定されたフォルダ内のCSVとJPGファイルを取得
 * 2. 取得したCSVとJPGファイルをZIP圧縮して、指定フォルダに保存
 * 3. ZIPファイルのダウンロードリンクをユーザーに提供
 */

function exportCsvAndJpg() {
  // スプレッドシートとシートの取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetA = ss.getSheetByName("07_SN_OUTPUT");  // 出力データを格納するシート
  var sheetB = ss.getSheetByName("01_BANNER_FOLDER_LIST");  // フォルダ情報を格納するシート
  
  // "SmartNews広告"という名前がついたフォルダIDを取得
  var folderId = getFolderIdFromSheet(sheetB, 'SmartNews広告');
  
  // フォルダIDが見つからない場合、アラートを表示して処理を終了
  if (!folderId) {
    SpreadsheetApp.getUi().alert('フォルダが見つかりません。');
    return;
  }
  
  // Googleドライブ内の指定されたフォルダを取得
  var destinationFolder = DriveApp.getFolderById(folderId);
  
  // シートAのデータを取得し、CSV形式に変換
  var data = sheetA.getDataRange().getValues();
  var csv = convertToCsv(data);
  
  // CSVデータをBlob形式に変換
  var csvBlob = Utilities.newBlob(csv, 'text/csv', 'data.csv').setBytes(Utilities.newBlob(csv).getBytes());

  // 指定されたフォルダ内のファイルを取得
  var sourceFolder = DriveApp.getFolderById(folderId); 
  
  // フォルダ内のファイルをファイル名をキーにしてマップに格納
  var files = sourceFolder.getFiles();
  var fileMap = {};
  
  while (files.hasNext()) {
    var file = files.next();
    fileMap[file.getName()] = file.getBlob();
  }

  // JPGファイルのBlobを格納する配列
  var jpgBlobs = [];
  
  // 同じファイルを重複して処理しないためのセット
  var addedFileNames = new Set();
  
  // シートAの各行を確認し、M列とN列に記載されたJPGファイル名を処理
  for (var i = 1; i < data.length; i++) {
    var mColumnJpg = data[i][12];  // M列（画像ファイル名）
    var nColumnJpg = data[i][13];  // N列（画像ファイル名）

    // M列とN列のファイルが存在し、かつまだ追加されていない場合に処理
    if (mColumnJpg && mColumnJpg.endsWith(".jpg") && !addedFileNames.has(mColumnJpg) && fileMap[mColumnJpg]) {
      jpgBlobs.push(fileMap[mColumnJpg].setName(mColumnJpg));
      addedFileNames.add(mColumnJpg);
    }
    
    if (nColumnJpg && nColumnJpg.endsWith(".jpg") && !addedFileNames.has(nColumnJpg) && fileMap[nColumnJpg]) {
      jpgBlobs.push(fileMap[nColumnJpg].setName(nColumnJpg));
      addedFileNames.add(nColumnJpg);
    }
  }

  // CSVとJPGファイルを1つの配列にまとめ、ZIPファイルを作成
  var allBlobs = [csvBlob].concat(jpgBlobs);
  var zipBlob = Utilities.zip(allBlobs, 'files.zip');
  
  // 作成したZIPファイルをGoogleドライブの指定フォルダに保存
  var zipFile = destinationFolder.createFile(zipBlob);
  
  // ZIPファイルのダウンロードURLを取得
  var downloadUrl = zipFile.getDownloadUrl();

  // ダウンロードリンクを表示するHTMLを作成
  var htmlOutput = HtmlService.createHtmlOutput('<a href="' + downloadUrl + '" target="_blank">こちらをクリックしてZIPファイルをダウンロード</a>');
  
  // スプレッドシートのUIにダウンロードリンクを表示
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'zipファイル生成中');
}

/**
 * スプレッドシートから指定した名前のフォルダIDを取得する関数
 * @param {Object} sheet フォルダ情報が格納されているシート
 * @param {string} searchText 検索するフォルダ名（例: "SmartNews広告"）
 * @returns {string|null} フォルダIDまたはnull（見つからない場合）
 */
function getFolderIdFromSheet(sheet, searchText) {
  var data = sheet.getDataRange().getValues();  // シートのデータを取得
  
  // フォルダ名が一致する行を検索
  for (var i = 0; i < data.length; i++) {
    if (data[i].includes(searchText)) {
      var folderUrl = data[i][data[i].length - 1];  // フォルダURLを取得
      return extractFolderIdFromUrl(folderUrl);  // URLからIDを抽出
    }
  }
  return null;  // フォルダが見つからない場合はnullを返す
}

/**
 * URLからフォルダIDを抽出する関数
 * @param {string} url フォルダのURL
 * @returns {string|null} フォルダIDまたはnull（URLが不正な場合）
 */
function extractFolderIdFromUrl(url) {
  var matches = url.match(/folders\/([a-zA-Z0-9_-]+)/);  // 正規表現でIDを抽出
  return matches ? matches[1] : null;  // IDが見つかれば返す
}

/**
 * 二次元配列をCSV形式の文字列に変換する関数
 * @param {Array} data 二次元配列のデータ
 * @returns {string} CSV形式の文字列
 */
function convertToCsv(data) {
  return data.map(function(row) {
    return row.join(',');  // 各行をカンマで結合
  }).join('\n');  // 行ごとに改行
}
