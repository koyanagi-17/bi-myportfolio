/**
 * スプレッドシートを開いたときに実行される関数です。
 * スプレッドシートのUIにカスタムメニューを追加します。
 * メニュー項目をクリックすると、それぞれ対応する関数が実行されます。
 * 
 * 追加されるメニュー項目：
 * 1. "02_MONTHLY_SUMMARY更新" - 関数 "multiprocessor_2" を実行
 * 2. "翌月シート作成" - 関数 "multiprocessor_1" を実行
 */

function onOpen() {
  // スプレッドシートのUIに新しいメニューを追加
  SpreadsheetApp.getUi()
    // メニュー名を "GAS" に設定
    .createMenu('GAS')
    // "02_MONTHLY_SUMMARY更新" というメニュー項目を追加し、"multiprocessor_2" 関数を実行
    .addItem("02_MONTHLY_SUMMARY更新", "multiprocessor_2")
    // "翌月シート作成" というメニュー項目を追加し、"multiprocessor_1" 関数を実行
    .addItem("翌月シート作成", "multiprocessor_1")
    // メニューをスプレッドシートのUIに追加
    .addToUi();
}
