function GetBannerList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var drive_ss = ss.getSheetByName("バナーフォルダ一覧");
  var banner_ss = ss.getSheetByName("banner_list");
  var driveurl_List = drive_ss.getRange("A1:C").getValues().filter(function(row) {
    return row[0] != "";
  });

  var banner_List = banner_ss.getRange("A1:E").getValues().filter(function(row) {
    return row[0] != "";
  });

  var list = [];

  for (var j = 8; j < driveurl_List.length + 1; j++) {
    var folder_url = drive_ss.getRange(j, 2).getValue();
    var folder_id = folder_url.split('/folders/')[1];
    var folder = DriveApp.getFolderById(folder_id);
    var files = folder.getFiles();

    while (files.hasNext()) {
      var buff = files.next();
      var medianame = drive_ss.getRange(j, 1).getValue();
      var date = new Date(buff.getDateCreated());
      var uploadDate = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
      list.push([uploadDate, medianame, buff.getName(), buff.getUrl(), buff.getUrl()]);
    }
  }

  list.sort(function(a, b) {
    var dateComparison = new Date(a[0]) - new Date(b[0]);
    if (dateComparison != 0) {
      return dateComparison;
    }
    return a[2].localeCompare(b[2]);
  });

  var res = [];
  var acquired = banner_ss.getRange("D2:D").getValues().filter(String);

  var res = list.filter(function(e) {
    return acquired.filter(function(f) {
      return e[4].toString() == f[0].toString();
    }).length == 0;
  }).filter(function(value) {
    return (value[2].indexOf("CR一覧") == -1);
  });

  for (var i in res) {
    var startId = res[i][4].indexOf("file/d/") + 7;
    var endId = res[i][4].indexOf("/view");

    res[i][4] = "=IMAGE(\"https://drive.google.com/uc?export=download&id=" + res[i][4].slice(startId, endId) + "\")";

    if (res[i][2].indexOf("mp4") != -1) {
      res[i][4] = "動画のため左のURLよりご確認ください";
    };
    if (res[i][2].indexOf("mov") != -1) {
      res[i][4] = "動画のため左のURLよりご確認ください";
    };
    if (res[i][2].indexOf("MOV") != -1) {
      res[i][4] = "動画のため左のURLよりご確認ください";
    };
    if (res[i][2].indexOf("MP4") != -1) {
      res[i][4] = "動画のため左のURLよりご確認ください";
    }
  }

  if (res.length !== 0) {
    var range = banner_ss.getRange(banner_List.length + 1, 1, res.length, res[0].length);
    range.setValues(res);
  }
}
