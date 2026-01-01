# デプロイメントガイド

このドキュメントでは、Tableau Blueprint ワークショップアプリを実際にGoogle Apps Scriptにデプロイする手順を説明します。

## 前提条件

- Googleアカウント
- Google Driveへのアクセス
- Google Apps Script エディタへのアクセス権限

## claspでデプロイ（推奨）

このリポジトリには、Node.js がPCに入っていなくても `clasp` を実行できるPowerShellスクリプトを同梱しています。

1. 初回のみ: Apps Script API を有効化（Google側のユーザー設定）
   - `https://script.google.com/home/usersettings` を開き、「Google Apps Script API」を ON
2. デプロイ（新規GASプロジェクトを作成してpush）
   - `powershell -ExecutionPolicy Bypass -File tools/deploy-gas.ps1 -Title "salesnavi_ne-2" -DeployDescription "initial deploy"`
3. （任意）スプレッドシートの紐付け＋シート初期化
   - 既存シートを使う: `powershell -ExecutionPolicy Bypass -File tools/deploy-gas.ps1 -SpreadsheetId "<SPREADSHEET_ID>"`
   - 新規に作る: `powershell -ExecutionPolicy Bypass -File tools/deploy-gas.ps1 -CreateNewSpreadsheet -SpreadsheetName "Tableau Blueprint DB"`
4. 日常運用（ターミナルからURLを固定して更新したい場合）
   - Apps Script エディタでWebアプリとして1回だけデプロイし、その「デプロイID（AKfy...）」を控える
   - 以後は `powershell -ExecutionPolicy Bypass -File tools/deploy-gas.ps1 -DeploymentId "<AKfy...>"` を実行（`.gas-deploy.json` に保存され、同じURLのまま更新できます）
5. コード反映のみ（deploy不要）
   - `powershell -ExecutionPolicy Bypass -File tools/push-gas.ps1`
   - 監視して自動反映: `powershell -ExecutionPolicy Bypass -File tools/watch-gas.ps1`

## デプロイ手順

### ステップ 1: Google Apps Script プロジェクトの作成

1. [Google Apps Script](https://script.google.com) にアクセス
2. 「新しいプロジェクト」をクリック
3. プロジェクト名を「Tableau Blueprint Workshop」に変更（左上の「無題のプロジェクト」をクリック）

### ステップ 2: スプレッドシート（データベース）の作成

1. [Google Spreadsheets](https://sheets.google.com) にアクセス
2. 「新しいスプレッドシート」を作成
3. スプレッドシート名を「Tableau Blueprint DB」に変更
4. **スプレッドシートIDをコピー**
   - URLの `/d/` と `/edit` の間の文字列をコピー
   - 例: `https://docs.google.com/spreadsheets/d/【ここがID】/edit`

### ステップ 3: コードのアップロード

Apps Script エディタで、以下のファイルを作成します：

#### 3-1. appsscript.json

1. エディタ左側の「プロジェクトの設定」（⚙️）をクリック
2. 「マニフェスト ファイルをエディタで表示する」にチェック
3. `appsscript.json` ファイルを開き、以下の内容で置き換え：

```json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Drive",
        "version": "v3",
        "serviceId": "drive"
      },
      {
        "userSymbol": "Docs",
        "version": "v1",
        "serviceId": "docs"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {
    "executeAs": "USER_ACCESSING",
    "access": "ANYONE"
  }
}
```

#### 3-2. スクリプトファイル（.gs）のアップロード

エディタ左側の「ファイル」(+) から以下のファイルを追加し、このリポジトリの `src/` 配下のコードをコピー＆ペースト：

- `src/` 配下の **すべての `.gs` ファイル**（例: `src/Code.gs`, `src/Config.gs`, `src/Auth.gs`, `src/repositories/SpreadsheetRepository.gs`, `src/services/*.gs`）を同名で作成して貼り付け

#### 3-3. HTMLファイルのアップロード

エディタ左側の「ファイル」(+) から「HTML」を選択し、以下のファイルを追加：

1. **Index** - `src/Index.html` の内容をコピー
2. **Styles** - `src/Styles.html` の内容をコピー
3. **Scripts** - `src/Scripts.html` の内容をコピー

### ステップ 4: スクリプトプロパティの設定

1. エディタ左側の「プロジェクトの設定」（⚙️）をクリック
2. 「スクリプト プロパティ」セクションまでスクロール
3. 「スクリプト プロパティを追加」をクリック
4. 以下を入力：
   - プロパティ: `SPREADSHEET_ID`
   - 値: ステップ2でコピーしたスプレッドシートID
5. 「スクリプト プロパティを保存」をクリック

補足: Apps Script エディタで `setupSpreadsheet('<SPREADSHEET_ID>')` を実行すると、スクリプトプロパティ設定とシート初期化をまとめて行えます。

### ステップ 5: Web Appとしてデプロイ

1. エディタ右上の「デプロイ」→「新しいデプロイ」をクリック
2. 「種類の選択」で「ウェブアプリ」を選択
3. 以下を設定：
   - **説明**: 「初回デプロイ」（任意）
   - **次のユーザーとして実行**: **アプリにアクセスするユーザー**（重要！）
   - **アクセスできるユーザー**: 自分のみ（または組織内のユーザー）
4. 「デプロイ」をクリック
5. 権限の承認を求められたら：
   - 「アクセスを承認」をクリック
   - Googleアカウントを選択
   - 「詳細」→「Tableau Blueprint Workshop（安全ではないページ）に移動」をクリック
   - 「許可」をクリック
6. **Web AppのURLをコピー**（このURLがアプリケーションのアクセスURLです）

### ステップ 6: データベースの初期化

1. コピーしたWeb App URLにアクセス
2. 画面左側のナビゲーションで「シート初期化」ボタンをクリック
3. 確認ダイアログで「OK」をクリック
4. スプレッドシートを確認し、9つのシートが作成されていることを確認：
   - プロジェクト
   - ビジョン
   - ユースケース
   - 90日計画
   - 体制RACI
   - 統制
   - 運営支援
   - 価値
   - 監査ログ

### ステップ 7: 動作確認

#### 7-1. プロジェクト作成

1. ホーム画面の「プロジェクト管理」カードをクリック
2. 「新規プロジェクト作成」ボタンをクリック
3. 顧客名を入力（例: 株式会社サンプル）
4. 「作成」をクリック
5. プロジェクトが作成され、一覧に表示されることを確認

#### 7-2. ビジョン入力

1. 作成したプロジェクトをクリック
2. 左側ナビで「1. 北極星ビジョン」をクリック
3. ビジョン本文を入力
4. 「保存」をクリック
5. 「✓ 保存しました」と表示されることを確認

#### 7-3. ユースケース追加

1. 左側ナビで「2. ユースケース選定」をクリック
2. 「ユースケースを追加」ボタンをクリック
3. 課題と狙いを入力
4. 「追加」をクリック
5. ユースケースが一覧に表示されることを確認

#### 7-4. 体制・RACI設定

1. 左側ナビで「4. 体制・RACI」をクリック
2. 「行を追加」ボタンで行を追加
3. 3本柱区分、タスク、担当者、RACIを入力
4. 「保存」をクリック
5. データが保存されることを確認

#### 7-5. 価値トラッキング

1. 左側ナビで「7. 価値トラッキング」をクリック
2. 定量効果・定性効果を入力
3. 「保存」ボタンをクリック
4. データが保存されることを確認

#### 7-6. 稟議書案生成

1. 左側ナビで「稟議書案を生成」ボタンをクリック
2. 確認ダイアログで「OK」をクリック
3. 新しいタブでGoogleドキュメントが開くことを確認
4. 稟議書案の内容が正しく生成されていることを確認

## コード更新時のデプロイ

コードを修正した場合：

1. Apps Script エディタで該当ファイルを編集
2. 自動保存されます（Ctrl+S でも保存可能）
3. 「デプロイ」→「デプロイを管理」
4. 既存のデプロイの「編集」（鉛筆アイコン）をクリック
5. 「バージョン」を「新バージョン」に変更
6. 「デプロイ」をクリック
7. 既存のWeb App URLはそのまま使用できます

## トラブルシューティング

### エラー: "SPREADSHEET_ID が設定されていません"

**原因**: スクリプトプロパティが未設定

**解決方法**: ステップ4を再実行し、SPREADSHEET_IDを設定

### エラー: "このプロジェクトへのアクセス権限がありません"

**原因**: プロジェクトの編集者リストに含まれていない

**解決方法**: プロジェクト作成者に編集者として追加してもらう

### 画面が真っ白

**原因**: HTMLファイルが正しくアップロードされていない、またはスクリプトエラー

**解決方法**:
1. ブラウザのコンソール（F12）でエラーを確認
2. Apps Script エディタの「実行ログ」でエラーを確認
3. HTMLファイル（Index, Styles, Scripts）が正しくアップロードされているか確認

### ドキュメント生成時のエラー

**原因**: Docs APIが有効化されていない

**解決方法**:
1. `appsscript.json` の dependencies が正しく設定されているか確認
2. プロジェクトの設定で「マニフェスト ファイルをエディタで表示する」が有効か確認

## セキュリティに関する注意事項

1. **アクセス制御**: Web Appの「アクセスできるユーザー」設定を適切に管理してください
2. **プロジェクト権限**: プロジェクトの編集者メールは慎重に管理してください
3. **監査ログ**: すべての操作が監査ログシートに記録されます
4. **スプレッドシートID**: Script Propertiesに保存されるため、コードには含まれません

## パフォーマンスに関する注意

- Apps Scriptには1日の実行時間制限があります（無料版: 90分/日）
- 大量のデータを扱う場合、スプレッドシートの読み書き回数を最適化してください
- 同時アクセスユーザー数が多い場合、Apps Scriptの同時実行制限に注意してください

## サポート

問題が発生した場合：
1. このリポジトリのIssuesに報告
2. CLAUDE.md の「意思決定ログ」を参照
3. Apps Script エディタの「実行ログ」を確認

## 次のステップ

MVP実装が完了したら、以下の拡張モジュールの実装を検討してください：

- **Module 3**: 6コアプロセス ギャップ診断
- **Module 5**: 最小限の統制設計キット
- **Module 6**: 運営支援設計＋運用台帳

これらのモジュールは、始動フェーズから強化フェーズへの移行をサポートします。
