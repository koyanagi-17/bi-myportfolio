function exportCsvAndJpg() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetA = ss.getSheetByName("07_SN_OUTPUT");
  var sheetB = ss.getSheetByName("01_BANNER_FOLDER_LIST");

  var folderId = getFolderIdFromSheet(sheetB, 'SmartNews広告');
  if (!folderId) {
    SpreadsheetApp.getUi().alert('フォルダが見つかりません。');
    return;
  }
  
  var destinationFolder = DriveApp.getFolderById(folderId);
  
  var data = sheetA.getDataRange().getValues();
  var csv = convertToCsv(data);
  
  var csvBlob = Utilities.newBlob(csv, 'text/csv', 'data.csv').setBytes(Utilities.newBlob(csv).getBytes());

  var sourceFolder = DriveApp.getFolderById(folderId); 
  
  var files = sourceFolder.getFiles();
  var fileMap = {};
  
  while (files.hasNext()) {
    var file = files.next();
    fileMap[file.getName()] = file.getBlob();
  }

  var jpgBlobs = [];
  var addedFileNames = new Set();
  
  for (var i = 1; i < data.length; i++) {
    var mColumnJpg = data[i][12];
    var nColumnJpg = data[i][13];

    if (mColumnJpg && mColumnJpg.endsWith(".jpg") && !addedFileNames.has(mColumnJpg) && fileMap[mColumnJpg]) {
      jpgBlobs.push(fileMap[mColumnJpg].setName(mColumnJpg));
      addedFileNames.add(mColumnJpg);
    }
    
    if (nColumnJpg && nColumnJpg.endsWith(".jpg") && !addedFileNames.has(nColumnJpg) && fileMap[nColumnJpg]) {
      jpgBlobs.push(fileMap[nColumnJpg].setName(nColumnJpg));
      addedFileNames.add(nColumnJpg);
    }
  }

  var allBlobs = [csvBlob].concat(jpgBlobs);
  var zipBlob = Utilities.zip(allBlobs, 'files.zip');
  
  var zipFile = destinationFolder.createFile(zipBlob);
  
  var downloadUrl = zipFile.getDownloadUrl();

  var htmlOutput = HtmlService.createHtmlOutput('<a href="' + downloadUrl + '" target="_blank">こちらをクリックしてZIPファイルをダウンロード</a>');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'zipファイル生成中');
}

function getFolderIdFromSheet(sheet, searchText) {
  var data = sheet.getDataRange().getValues();
  
  for (var i = 0; i < data.length; i++) {
    if (data[i].includes(searchText)) {
      var folderUrl = data[i][data[i].length - 1];
      return extractFolderIdFromUrl(folderUrl);
    }
  }
  return null;
}

function extractFolderIdFromUrl(url) {
  var matches = url.match(/folders\/([a-zA-Z0-9_-]+)/);
  return matches ? matches[1] : null;
}

function convertToCsv(data) {
  return data.map(function(row) {
    return row.join(',');
  }).join('\n');
}
