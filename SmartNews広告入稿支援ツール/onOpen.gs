function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GAS')
    .addItem("ドライブURLリスト更新", "GetFileList")
    .addItem("バナーリスト更新", "GetBannerList")
    .addToUi();
}
