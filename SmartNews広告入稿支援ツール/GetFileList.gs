function GetFileList(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var drive_ss = spreadsheet.getSheetByName("バナーフォルダ一覧");
  listFoldersInParentFolder(drive_ss);
  GetSheetList(drive_ss);
};

function listFoldersInParentFolder(drive_ss) {
  var parentFolderURL = drive_ss.getRange("B4").getValue();
  var parentFolderId = parentFolderURL.split('/folders/')[1];
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  var subFolders = parentFolder.getFolders();
  var data = [];

  while (subFolders.hasNext()) {
    var folder = subFolders.next();
    var folderName = folder.getName();
    var folderURL = folder.getUrl();
    data.push([folderName, folderURL]);
  }
  drive_ss.getRange(8, 1, data.length, 2).setValues(data);
}

function GetSheetList(drive_ss){
  var driveurl_List = drive_ss.getRange("A8:C").getValues().filter(function(row){
    return row[0] != "";
  });

  var list = []; 

  for(var j = 8; j < driveurl_List.length + 1 ; j++){
    var folder_url = drive_ss.getRange(j, 2).getValue();
    var folder_id = folder_url.split('/folders/')[1];
    var folder = DriveApp.getFolderById(folder_id);
    var files = folder.getFiles();
    
    while(files.hasNext()){
      var buff = files.next();
      list.push([buff.getName(), buff.getUrl()]);
    }
  }
  var ssurl = [];
  var ssurl = list.filter(function(url){
    return url[0].indexOf("CR一覧") != -1;
  });

  var afterArray = ssurl.map(function(x){
    return x.splice(1, 1);
  });
  Logger.log(afterArray);

  if(ssurl.length !== 0){
    drive_ss.getRange(4, 3, ssurl.length, 1).setValues(afterArray);
  }
}
