/**
 * この関数は「01_BANNER_FOLDER_LIST」シート内のフォルダURLから、Google ドライブ内のファイルリストを取得し、必要な情報を「02_BANNER_LIST」シートに書き込みます。
 * 特に、動画ファイル（mp4、movなど）のURLは特別に処理され、画像URLに変換されます。
 * 
 * 処理の流れ：
 * 1. 「01_BANNER_FOLDER_LIST」シートからフォルダURLを取得
 * 2. 各フォルダ内のファイル（バナー画像など）を取得
 * 3. ファイル情報（作成日、メディア名、ファイル名、URL）を整形
 * 4. 既に「02_BANNER_LIST」シートに存在するURLを除外
 * 5. ファイルの拡張子が「mp4」「mov」の場合はURLを特別処理
 * 6. 新しいファイルリストを「02_BANNER_LIST」シートに書き込む
 */

function GetBannerList() {
  // 現在のスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 「01_BANNER_FOLDER_LIST」シートを取得
  var drive_ss = ss.getSheetByName("01_BANNER_FOLDER_LIST");
  
  // 「02_BANNER_LIST」シートを取得
  var banner_ss = ss.getSheetByName("02_BANNER_LIST");
  
  // 「01_BANNER_FOLDER_LIST」シートからフォルダURLのリストを取得（空白の行を除外）
  var driveurl_List = drive_ss.getRange("A1:C").getValues().filter(function(row) {
    return row[0] != "";
  });

  // 「02_BANNER_LIST」シートから既存のバナーリストを取得（空白の行を除外）
  var banner_List = banner_ss.getRange("A1:E").getValues().filter(function(row) {
    return row[0] != "";
  });

  // 新しいファイル情報を格納する配列
  var list = [];

  // 各フォルダURLについて処理を実行
  for (var j = 8; j < driveurl_List.length + 1; j++) {
    // フォルダURLを取得
    var folder_url = drive_ss.getRange(j, 2).getValue();
    
    // URLからフォルダIDを抽出
    var folder_id = folder_url.split('/folders/')[1];
    
    // DriveAppを使ってフォルダを取得
    var folder = DriveApp.getFolderById(folder_id);
    
    // フォルダ内のファイルを取得
    var files = folder.getFiles();

    // 各ファイルをループして情報を収集
    while (files.hasNext()) {
      var buff = files.next();  // ファイルオブジェクトを取得
      var medianame = drive_ss.getRange(j, 1).getValue();  // メディア名を取得
      var date = new Date(buff.getDateCreated());  // ファイルの作成日を取得
      var uploadDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');  // 作成日を指定フォーマットに変換
      list.push([uploadDate, medianame, buff.getName(), buff.getUrl(), buff.getUrl()]);  // ファイル情報をリストに追加
    }
  }

  // ファイルリストを作成日でソート
  list.sort(function(a, b) {
    var dateComparison = new Date(a[0]) - new Date(b[0]);  // 日付順に比較
    if (dateComparison != 0) {
      return dateComparison;  // 日付が異なればその順で並べ替え
    }
    return a[2].localeCompare(b[2]);  // ファイル名で比較（同じ日付の場合）
  });

  // 「02_BANNER_LIST」シートで既に取得されたファイルURLをリストとして取得
  var acquired = banner_ss.getRange("D2:D").getValues().filter(String);

  // 新しいファイルリストを、既存のリストと照合して重複を除外
  var res = list.filter(function(e) {
    return acquired.filter(function(f) {
      return e[4].toString() == f[0].toString();  // 既存URLがリストにない場合
    }).length == 0;
  }).filter(function(value) {
    return (value[2].indexOf("CR一覧") == -1) &&  // "CR一覧"が含まれている場合は除外
           (value[2].toLowerCase().indexOf(".zip") == -1);  // .zipファイルは除外
  });

  // 新しいリストのURLを画像として表示するために処理
  for (var i in res) {
    // Google DriveのURLからファイルIDを抽出
    var startId = res[i][4].indexOf("file/d/") + 7;
    var endId = res[i][4].indexOf("/view");
    
    // 画像として表示するURLを作成
    res[i][4] = "=IMAGE(\"https://drive.google.com/uc?export=download&id=" + res[i][4].slice(startId, endId) + "\")";

    // 動画ファイル（mp4, mov）の場合、URLではなくメッセージを表示
    if (res[i][2].indexOf("mp4") != -1 || res[i][2].indexOf("mov") != -1 || res[i][2].indexOf("MOV") != -1 || res[i][2].indexOf("MP4") != -1) {
      res[i][4] = "動画のため左のURLよりご確認ください";
    }
  }

  // 新しいリストがあれば「02_BANNER_LIST」シートに書き込む
  if (res.length !== 0) {
    var range = banner_ss.getRange(banner_List.length + 1, 1, res.length, res[0].length);
    range.setValues(res);  // ファイル情報をシートに書き込む
  }
}
