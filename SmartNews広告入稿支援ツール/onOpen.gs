function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GAS')
    .addItem("01_BANNER_FOLDER_LIST更新", "GetFileList")
    .addItem("02_BANNER_LIST更新", "GetBannerList")
    .addToUi();
}
