/**
 * このスクリプトは、Google ドライブ内で「SmartNews広告」という名前のフォルダを検索し、 
 * その中にある「files.zip」という名前のZIPファイルをゴミ箱に移動します。
 *
 * 主な処理：
 * 1. スプレッドシート内で「SmartNews広告」という名前のフォルダURLを探す
 * 2. そのフォルダ内にある「files.zip」という名前のファイルを見つけて削除
 */

function deleteZipFile() {
  // スプレッドシートの取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('01_BANNER_FOLDER_LIST');  // フォルダURLが格納されているシート

  // シート内のデータを取得
  var data = sheet.getDataRange().getValues();
  var folderUrl = '';  // フォルダURLを格納する変数

  // 「SmartNews広告」という名前のフォルダURLを探す
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === 'SmartNews広告') {
      folderUrl = data[i][1];  // 見つかった場合、そのURLを格納
      break;
    }
  }

  // フォルダURLが見つからない場合はログに記録して終了
  if (!folderUrl) {
    Logger.log('フォルダURLが見つかりませんでした。');
    return;
  }

  // フォルダURLからフォルダIDを抽出
  var folderId = getFolderIdFromUrl(folderUrl);

  // フォルダIDが無効な場合はログに記録して終了
  if (!folderId) {
    Logger.log('フォルダIDが無効です。');
    return;
  }

  // フォルダIDを使って実際のGoogleドライブフォルダを取得
  var folder = DriveApp.getFolderById(folderId);
  
  // フォルダ内のファイルを取得
  var files = folder.getFiles();
  
  // フォルダ内のファイルを順に確認
  while (files.hasNext()) {
    var file = files.next();
    
    // ファイル名が「files.zip」だった場合、そのファイルをゴミ箱に移動
    if (file.getName() === 'files.zip') {
      file.setTrashed(true);  // ゴミ箱に移動
      Logger.log('Deleted file: ' + file.getName());  // ログに削除されたファイル名を記録
    }
  }
}

/**
 * URLからフォルダIDを抽出する関数
 * @param {string} url フォルダのURL
 * @returns {string|null} フォルダIDまたはnull（URLが不正な場合）
 */
function getFolderIdFromUrl(url) {
  var regex = /folders\/([a-zA-Z0-9-_]+)/;  // 正規表現を使ってフォルダIDを抽出
  var match = url.match(regex);
  return match ? match[1] : null;  // IDが見つかれば返す、見つからなければnullを返す
}
