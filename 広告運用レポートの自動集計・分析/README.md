# 広告運用レポートの自動集計・分析（GAS）

## 概要

BigQueryから取得した広告データをもとに、商材別の成果指標（CPA・CVRなど）を日次・週次・月次で自動集計するスプレッドシートと、それを管理・更新するGoogle Apps Scriptです。

## 特徴
- QUERY関数＋IMPORTRANGE＋ARRAYFORMULAでデータを整形
- GASとAIを活用して、毎月の更新・集計業務を完全自動化
- GASトリガーで定期実行をスケジュール
- 社内全員が利用できる形式で配布・共有

## 使用技術
- Google Apps Script
- BigQuery（スプレッドシートのデータコネクタ経由）
- Google スプレッドシート（QUERY, XLOOKUP, ARRAYFORMULA など）

## ファイル一覧
- `updateSheet.gs`：データ整形・集計
- `createReport.gs`：レポート生成
- `setTriggers.gs`：月初の自動実行設定
- `utils.gs`：補助的な共通処理

