# 📊 広告運用レポートの自動集計・分析（GAS）

## 🎯 目的・背景
広告運用チームで毎月発生する「成果データの集計」「月別レポートの更新作業」に多くの時間と手間がかかっていたため、  
**ミスの防止・属人化の解消・作業負担の軽減**を目的に、スプレッドシート＋GASを使って自動化しました。

BIツールではなくスプレッドシートを選定したのは、非エンジニアのメンバーでも扱いやすく、柔軟に加工・共有ができるためです。

---

## ✨ 概要
BigQueryから取得した広告データをもとに、商材別の成果指標（CPA・CVRなど）を日次・週次・月次で自動集計するスプレッドシートと、  
それを管理・更新するGoogle Apps Scriptのセットです。

---

## 🔧 特徴
- `QUERY` + `IMPORTRANGE` + `ARRAYFORMULA` を用いて、BigQueryから取得したデータを特定商材ごとに自動整形
- 手作業での入力・誤入力を防止する仕組み
- GASとAI（ChatGPTなど）を活用し、**毎月の更新・集計業務を完全自動化**
- GASのトリガーで月末処理を定期実行
- 誰が実行しても安定して処理ができる構成に

---

## 🛠 使用技術
- Google Apps Script（GAS）
- BigQuery（スプレッドシートのデータコネクタ経由）
- Google スプレッドシート（QUERY, XLOOKUP, ARRAYFORMULA など）

---

## 👤 利用対象者
- 社内の全社員（広告運用・営業・マネジメント層）

---

## 🗂 ファイル構成

| ファイル名 | 役割 |
|------------|------|
| `onOpen.gs` | メニューバーにカスタムメニューを追加 |
| `trigger.gs` | 月末の自動実行トリガー設定 |
| `multiprocessor.gs` | 複数処理をまとめて実行 |
| `copyTabs.gs` | タブの複製処理 |
| `clearCells.gs` | 複製したタブ内の手入力箇所を全てクリア |
| `copyFunctionCells.gs` | 月別集計のタブを更新 |
| `getSheetName.gs` | インデックス作成 |
| `MoveActiveSheet.gs` | タブの配置変更 |
| `setFileNameToCell.gs` | ファイル名をセルに取得 |

---

## 📎 実際のシート
👉 [広告レポート自動集計スプレッドシート（閲覧用）](https://docs.google.com/spreadsheets/d/1d1GF3-Cb-T35ocooElEliWLSJrUc1L1500kz25qQSRg/edit?usp=sharing)

