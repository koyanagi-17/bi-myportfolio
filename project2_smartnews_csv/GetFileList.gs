/**
 * この関数は「01_BANNER_FOLDER_LIST」シートに、指定された親フォルダ内のサブフォルダと、そのフォルダ内のファイルのリストを取得し、シートに表示します。
 * 
 * 処理の流れ：
 * 1. 「01_BANNER_FOLDER_LIST」シートを取得
 * 2. 親フォルダのURLをシートから取得し、そのフォルダ内のサブフォルダとファイルリストを取得
 * 3. サブフォルダとファイルの情報をシートに書き込む
 */
function GetFileList(){
  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // 「01_BANNER_FOLDER_LIST」シートを取得
  var drive_ss = spreadsheet.getSheetByName("01_BANNER_FOLDER_LIST");
  
  // 親フォルダ内のサブフォルダリストを取得
  listFoldersInParentFolder(drive_ss);
  
  // シート内のファイルリストを取得
  GetSheetList(drive_ss);
};

/**
 * この関数は指定された親フォルダ内のサブフォルダを取得し、その名前とURLを「01_BANNER_FOLDER_LIST」シートに表示します。
 * 
 * 処理の流れ：
 * 1. 親フォルダのURLをシートから取得
 * 2. 親フォルダ内のすべてのサブフォルダを取得
 * 3. サブフォルダ名とURLをシートに書き込む
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} drive_ss - 対象となるスプレッドシートのシートオブジェクト
 */
function listFoldersInParentFolder(drive_ss) {
  // 親フォルダのURLをシートから取得
  var parentFolderURL = drive_ss.getRange("B4").getValue();
  
  // URLから親フォルダのIDを抽出
  var parentFolderId = parentFolderURL.split('/folders/')[1];
  
  // DriveAppを使って親フォルダを取得
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  
  // 親フォルダ内のサブフォルダを取得
  var subFolders = parentFolder.getFolders();
  
  // サブフォルダの情報を格納するための配列
  var data = [];

  // サブフォルダをループして情報を収集
  while (subFolders.hasNext()) {
    var folder = subFolders.next();
    var folderName = folder.getName();  // フォルダ名
    var folderURL = folder.getUrl();    // フォルダURL
    data.push([folderName, folderURL]); // サブフォルダ情報を配列に追加
  }

  // 「01_BANNER_FOLDER_LIST」シートにサブフォルダ情報を書き込む
  drive_ss.getRange(8, 1, data.length, 2).setValues(data);
}

/**
 * この関数は、指定されたサブフォルダ内のファイルを取得し、その名前とURLを「01_BANNER_FOLDER_LIST」シートに表示します。
 * 
 * 処理の流れ：
 * 1. サブフォルダ内のファイルをリストアップ
 * 2. 各ファイルの名前とURLを取得
 * 3. 特定のファイル（ここではコメントアウト部分を参照）をフィルタリングしてシートに書き込む
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} drive_ss - 対象となるスプレッドシートのシートオブジェクト
 */
function GetSheetList(drive_ss){
  // シートのA8:C列からデータを取得（親フォルダURLとサブフォルダURL）
  var driveurl_List = drive_ss.getRange("A8:C").getValues().filter(function(row){
    return row[0] != "";  // 空白の行を除外
  });

  var list = []; // ファイル情報を格納する配列

  // 各サブフォルダURLに対して処理を実行
  for(var j = 8; j < driveurl_List.length + 1 ; j++){
    var folder_url = drive_ss.getRange(j, 2).getValue();  // フォルダURLを取得
    var folder_id = folder_url.split('/folders/')[1];    // フォルダIDを取得
    var folder = DriveApp.getFolderById(folder_id);      // フォルダを取得
    var files = folder.getFiles();                       // フォルダ内のファイルを取得
    
    // ファイルをループして名前とURLを取得
    while(files.hasNext()){
      var buff = files.next();
      list.push([buff.getName(), buff.getUrl()]);  // ファイル名とURLを配列に追加
    }
  }

  // フィルタリング後のファイルURL（コメントアウトされたコード部分）
  /*
  var ssurl = list.filter(function(url){
    return url[0].indexOf("CR一覧") != -1;  // "CR一覧" を含むファイル名をフィルタリング
  });
  */

  // 取得したファイルURLのみを新しい配列に変換（コメントアウト部分を参考に）
  var afterArray = ssurl.map(function(x){
    return x.splice(1, 1);  // URL部分だけを抽出
  });

  Logger.log(afterArray);  // ログに出力（デバッグ用）

  // フィルタリングされたURLがあればシートに書き込み
  if(ssurl.length !== 0){
    drive_ss.getRange(4, 3, ssurl.length, 1).setValues(afterArray);
  }
}
