# Tableau Blueprint v10 営業ワークショップアプリ

営業が顧客と画面を見ながら、合意形成から稟議書案までを一気通貫で進められるワークショップ型アプリです。

## 概要

Tableau Blueprint v10の「始動→強化→凌駕」思想に基づき、以下の7つのモジュールを提供します：

### MVPモジュール (優先実装)
1. **北極星ビジョン・メーカー** - ビジョン策定支援
2. **戦略的ユースケース選定 + 90日計画** - 始動段階の計画立案
3. **3本柱組織 + RACI設計** - 体制・責任の明確化
4. **価値実現トラッカー** - 成果と証跡の管理

### 拡張モジュール
5. **6コアプロセス ギャップ診断** - 現状診断とロードマップ生成
6. **最小限の統制設計キット** - データガバナンスの骨格
7. **運営支援設計 + 運用台帳** - 継続運用の仕組み

## 技術スタック

- **プラットフォーム**: Google Apps Script (Web App)
- **データストア**: Google Spreadsheet
- **ドキュメント生成**: Google Docs API
- **認証**: Google Account
- **UI**: HTML Service

## セットアップ手順

### 1. Google Apps Script プロジェクト作成

1. [Google Apps Script](https://script.google.com) にアクセス
2. 新しいプロジェクトを作成
3. プロジェクト名を「Tableau Blueprint Workshop」などに設定

### 2. スプレッドシート (データベース) 作成

1. 新しいGoogleスプレッドシートを作成
2. 以下の9つのシートを作成:
   - プロジェクト
   - ビジョン
   - ユースケース
   - 90日計画
   - 体制RACI
   - 統制
   - 運営支援
   - 価値
   - 監査ログ

3. スプレッドシートのIDをコピー (URLの `/d/` と `/edit` の間の文字列)

### 3. コードのデプロイ

#### 方法A: clasp を使用 (推奨)

```bash
# claspのインストール (初回のみ)
npm install -g @google/clasp

# Google アカウントでログイン
clasp login

# Apps Script プロジェクトと連携
# .clasp.json.template をコピーして .clasp.json を作成し、
# scriptId を設定してください

cp .clasp.json.template .clasp.json
# .clasp.json の scriptId を編集

# コードをプッシュ
clasp push
```

#### 方法B: 手動コピー

1. `src/` 配下のすべての `.gs` ファイルをApps Script エディタにコピー
2. `src/html/` 配下のHTMLファイルも同様にコピー

### 4. スクリプトプロパティの設定

Apps Script エディタで:
1. プロジェクト設定 (⚙️) → スクリプトプロパティ
2. `SPREADSHEET_ID` を追加し、手順2で作成したスプレッドシートのIDを設定

### 5. Web App としてデプロイ

1. Apps Script エディタで「デプロイ」→「新しいデプロイ」
2. 種類: Web アプリ
3. 実行ユーザー: **アプリにアクセスするユーザー**
4. アクセスできるユーザー: 自分だけ (または組織内)
5. デプロイ

### 6. 初回アクセス

1. デプロイされたURLにアクセス
2. 権限の承認を求められるので承認
3. アプリが起動します

## 開発

詳細な設計ドキュメントは [CLAUDE.md](./CLAUDE.md) を参照してください。

### ディレクトリ構造

```
.
├── README.md                      # このファイル
├── CLAUDE.md                      # 設計ドキュメント
├── appsscript.json               # Apps Script マニフェスト
├── .clasp.json.template          # clasp 設定テンプレート
└── src/
    ├── Code.gs                   # エントリーポイント
    ├── Config.gs                 # 設定・定数
    ├── services/                 # ビジネスロジック層
    │   ├── ProjectService.gs
    │   ├── VisionService.gs
    │   ├── UsecaseService.gs
    │   ├── OrganizationService.gs
    │   ├── ValueService.gs
    │   └── DocumentService.gs
    ├── repositories/             # データアクセス層
    │   ├── SpreadsheetRepository.gs
    │   └── DriveRepository.gs
    └── html/                     # UI
        ├── Index.html
        ├── Styles.html
        └── Scripts.html
```

## 使い方

1. **新規プロジェクト作成**: ホーム画面で顧客名を入力して作成
2. **ビジョン策定**: Module 1 で北極星ビジョンを入力
3. **ユースケース選定**: Module 2 で戦略的ユースケースを登録し、90日計画を作成
4. **体制確定**: Module 4 で3本柱組織とRACIマトリクスを設計
5. **価値トラッキング**: Module 7 で成果を記録し、証跡を添付
6. **稟議書生成**: プロジェクトを「確定」すると、稟議書案が自動生成されます

## ライセンス

MIT License

## サポート

Issues または Pull Requests をお寄せください。
