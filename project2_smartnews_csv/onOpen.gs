/**
 * この関数は、スプレッドシートを開いた際にカスタムメニューを作成します。
 * メニュー項目として「01_BANNER_FOLDER_LIST更新」と「02_BANNER_LIST更新」を追加し、それぞれの項目に対応する関数を紐づけます。
 * 
 * 処理の流れは次の通りです：
 * 1. スプレッドシートが開かれたときに自動的に実行される
 * 2. メニュー「GAS」を作成
 * 3. メニューに「01_BANNER_FOLDER_LIST更新」と「02_BANNER_LIST更新」の2つの項目を追加
 * 4. 各メニュー項目をクリックすると、それぞれ対応する関数（`GetFileList` と `GetBannerList`）が実行される
 */

function onOpen() {
  // スプレッドシートの UI（ユーザーインターフェース）を取得
  SpreadsheetApp.getUi()
    // 新しいカスタムメニュー「GAS」を作成
    .createMenu('GAS')
    // メニュー項目「01_BANNER_FOLDER_LIST更新」を作成し、クリック時に「GetFileList」関数を実行
    .addItem("01_BANNER_FOLDER_LIST更新", "GetFileList")
    // メニュー項目「02_BANNER_LIST更新」を作成し、クリック時に「GetBannerList」関数を実行
    .addItem("02_BANNER_LIST更新", "GetBannerList")
    // 作成したメニューをスプレッドシートに追加
    .addToUi();
}
