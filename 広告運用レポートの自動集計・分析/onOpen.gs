function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GAS')
    .addItem("02_MONTHLY_SUMMARY更新", "multiprocessor_2")
    .addItem("翌月シート作成", "multiprocessor_1")
    .addToUi();
}
