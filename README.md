# Blueprint ワークショップ - Google Apps Script Webアプリ

営業が顧客同席で「合意形成→次アクション→稟議書案」まで進めるためのワークショップアプリケーションです。

## 🎯 概要

**目的**: 営業と顧客が一緒に使える、Google Apps Script ベースの業務Webアプリ
**技術**: Apps Script HTML Service + Spreadsheet DB + Google Docs Export
**画面数**: 9画面 (Home / Projects / Module 1-7)

## ✨ 主な機能

- ✅ プロジェクト作成/一覧/検索/状態管理（下書き/確定/アーカイブ）
- ✅ 各モジュールの保存/読み込み/完了マーク
- ✅ 価値トラッカーで証跡リンク（Drive）管理
- ✅ Googleドキュメント自動生成（稟議書エクスポート）
- ✅ 権限管理（プロジェクト作成者/編集者のみアクセス可）
- ✅ 監査ログ（すべての更新操作を記録）

## 🚀 デプロイ手順

### 1. Apps Script プロジェクト作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 「新しいプロジェクト」をクリック
3. プロジェクト名を「Blueprint Workshop」に変更

### 2. ファイルのアップロード

以下のファイルを Apps Script エディタにコピー＆ペースト:

1. **Code.gs** - メインのサーバーサイドロジック
2. **Index.html** - メインレイアウト
3. **Stylesheet.html** - スタイルシート
4. **ClientJS.html** - クライアントサイドJavaScript
5. **Views/Module1_Vision.html** (ファイル名: Module1_Vision)
6. **Views/Module2_UseCase.html** (ファイル名: Module2_UseCase)
7. **Views/Module4_RACI.html** (ファイル名: Module4_RACI)
8. **Views/Module7_ValueTracker.html** (ファイル名: Module7_ValueTracker)

### 3. データベース初期化

1. Apps Script エディタで Code.gs を開く
2. メニューから **実行 > 関数を実行 > initDb** を選択
3. 初回実行時、権限の許可が求められるので許可
4. 実行ログに成功メッセージとSpreadsheet IDが表示されることを確認

### 4. Webアプリとしてデプロイ

1. 右上の「デプロイ」→「新しいデプロイ」をクリック
2. 「種類の選択」で **ウェブアプリ** を選択
3. 設定を行い「デプロイ」をクリック
4. **Webアプリ URL** をコピーして保存

## 📖 使用方法

### プロジェクト作成
1. 「新規プロジェクト」をクリック
2. 顧客名を入力
3. モジュールを編集

### 稟議書の生成
1. プロジェクト詳細画面で「稟議書作成」ボタンをクリック
2. Google Docsが自動生成される

## 🔧 技術スタック

- Google Apps Script
- HTML Service
- Tailwind CSS (CDN)
- Google Spreadsheet (DB)
- Google Docs (Export)

---
**バージョン**: 1.0
