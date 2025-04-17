function deleteZipFile() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('01_BANNER_FOLDER_LIST');

  var data = sheet.getDataRange().getValues();
  var folderUrl = '';

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === 'SmartNews広告') {
      folderUrl = data[i][1];
      break;
    }
  }

  if (!folderUrl) {
    Logger.log('フォルダURLが見つかりませんでした。');
    return;
  }

  var folderId = getFolderIdFromUrl(folderUrl);

  if (!folderId) {
    Logger.log('フォルダIDが無効です。');
    return;
  }

  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  
  while (files.hasNext()) {
    var file = files.next();
    if (file.getName() === 'files.zip') {
      file.setTrashed(true);
      Logger.log('Deleted file: ' + file.getName());
    }
  }
}

function getFolderIdFromUrl(url) {
  var regex = /folders\/([a-zA-Z0-9-_]+)/;
  var match = url.match(regex);
  return match ? match[1] : null;
}
