/**
 * このスクリプトは、広告レポートのデータと手入力データを結合し、
 * 分析用に1つのシートにまとめて整形するものです。
 *
 * 主な処理内容：
 * 1. 2つのスプレッドシートからデータを取得（広告レポートと手入力の売上・利益情報）
 * 2. キャンペーンIDと日付をキーにしてデータを突合
 * 3. 指定の形式で新しいシートに出力（「整形済みデータ」シートを作成）
 */

function mergeCampaignDataFromURLs() {
  // 各スプレッドシートのID（実際は自分の環境に合わせて書き換え）
  const REPORT_SPREADSHEET_ID = '1jVwHuKIRYZ4c2xTy48O6b-2bLQklpmPJ5-T4epUCFss'; // 広告レポートデータ
  const MANUAL_SPREADSHEET_ID = '1sXkrQnvFnpgZZ59Nj7ZZx31jHl2xG4sHTe-hNbiy4ho'; // 手入力の売上データ

  // 各シートの取得
  const reportSheet = SpreadsheetApp.openById(REPORT_SPREADSHEET_ID).getSheetByName("campaign_report");
  const manualSheet = SpreadsheetApp.openById(MANUAL_SPREADSHEET_ID).getSheetByName("manual_input_sample");

  // データの取得
  const reportData = reportSheet.getDataRange().getValues(); // 広告レポートデータ
  const manualData = manualSheet.getDataRange().getValues(); // 売上などの手入力データ

  // 出力用のカラム名（見出し）
  const headers = [
    "Date", "ServiceNameJA", "AccountDescriptiveName", "PromotionName",
    "CampaignId_", "CampaignName", "Impressions", "Clicks", "Cost", "Conversions",
    "MCV", "CV", "売上", "粗利"
  ];

  // 手入力データの突合用マップを作成
  // キーは "日付#キャンペーンID"
  const manualMap = {};
  for (let i = 2; i < manualData.length; i++) { // 2行目以降（見出し除く）
    const row = manualData[i];
    const key = `${row[0]}#${row[4]}`; // 例: "2025-04-01#6e1d5xxxx"
    manualMap[key] = {
      MCV: row[6],     // 複数CV
      CV: row[7],      // 単一CV
      Sales: row[8],   // 売上
      Profit: row[9],  // 粗利
    };
  }

  // 出力データ（見出しをまず追加）
  const output = [headers];

  // 広告レポートデータの各行について処理
  for (let i = 2; i < reportData.length; i++) { // 2行目以降（見出し除く）
    const row = reportData[i];
    const date = row[7];         // Date
    const campaignId = row[5];   // CampaignId_
    const key = `${date}#${campaignId}`;
    const manual = manualMap[key]; // 手入力データがあれば取得

    // 1行分の出力データを作成
    output.push([
      date,                       // 日付
      row[8],                     // 媒体名（例: facebook広告）
      row[9],                     // アカウント名
      row[10],                    // プロモーション名
      campaignId,                 // キャンペーンID
      row[6],                     // キャンペーン名
      row[0],                     // インプレッション数
      row[1],                     // クリック数
      row[2],                     // コスト
      row[3],                     // CV（コンバージョン）
      manual ? manual.MCV : "",   // 複数CV（手入力があれば）
      manual ? manual.CV : "",    // 単一CV（手入力があれば）
      manual ? manual.Sales : "", // 売上（手入力があれば）
      manual ? manual.Profit : "" // 粗利（手入力があれば）
    ]);
  }

  // 出力先のシート（整形済みデータ）を準備
  const outputSheetName = "整形済みデータ";
  const thisSheet = SpreadsheetApp.getActiveSpreadsheet();
  let outputSheet = thisSheet.getSheetByName(outputSheetName);

  // 既に存在する場合は削除して作り直す
  if (outputSheet) {
    thisSheet.deleteSheet(outputSheet);
  }
  outputSheet = thisSheet.insertSheet(outputSheetName);

  // 整形されたデータをシートに出力（2行目から）
  outputSheet.getRange(2, 1, output.length, headers.length).setValues(output);
}

