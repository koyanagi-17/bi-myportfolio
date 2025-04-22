/**
 * この関数は、現在の日付と1週間後の日付を比較し、
 * 月が変わっている場合に翌月シート作成の処理を実行するためのものです。
 * 
 * 具体的には、現在の日付から1週間後の日付を計算し、現在の月と比較します。
 * もし、月が変わっていた場合（翌月に移行していた場合）、"multiprocessor_1" 関数を実行します。
 */

function trigger() {
  // 現在の日付を取得
  date = new Date();
  
  // 現在の月を取得（0:1月, 1:2月, ..., 11:12月）
  var month = date.getMonth();

  // 現在の日付から7日後の日付を設定
  date.setDate(date.getDate() + 7);
  
  // 7日後の月を取得
  var month_next_week = date.getMonth();

  // 現在の月と7日後の月が異なっていた場合（つまり、月が変わっていた場合）
  if (month != month_next_week) {
    // 月が変わっていた場合、翌月シート作成の処理を実行
    multiprocessor_1();
  }
}

