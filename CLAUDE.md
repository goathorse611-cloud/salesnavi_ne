# Tableau Blueprint v10 ベース 営業ワークショップアプリ - 開発メモ

## プロジェクト概要

### 目的
Tableau Blueprint v10 の「始動→強化→凌駕」思想に基づき、営業が顧客と画面を見ながら「合意形成→次アクション→稟議書案」まで進められるワークショップ型アプリを実装する。

### 技術スタック
- **プラットフォーム**: Google Apps Script (Web App)
- **データストア**: Google Spreadsheet (簡易DB)
- **ドキュメント生成**: Google Docs API
- **認証**: Google Account (Apps Script標準)
- **UI**: HTML Service + google.script.run
- **開発ツール**: clasp (可能であれば)

### アーキテクチャ方針

#### レイヤー構成
```
┌─────────────────────────────────────┐
│ Presentation Layer (HTML Service)  │
│  - Index.html (SPA)                │
│  - CSS/JavaScript (Stitch生成案)   │
└─────────────────────────────────────┘
          ↓ google.script.run
┌─────────────────────────────────────┐
│ Controller Layer (Code.gs)         │
│  - doGet() - Web App Entry         │
│  - API Handlers                    │
│  - Access Control                  │
└─────────────────────────────────────┘
          ↓
┌─────────────────────────────────────┐
│ Service Layer                      │
│  - ProjectService.gs               │
│  - VisionService.gs                │
│  - UsecaseService.gs               │
│  - OrganizationService.gs          │
│  - ValueService.gs                 │
│  - DocumentService.gs              │
└─────────────────────────────────────┘
          ↓
┌─────────────────────────────────────┐
│ Repository Layer                   │
│  - SpreadsheetRepository.gs        │
│  - DriveRepository.gs              │
└─────────────────────────────────────┘
          ↓
┌─────────────────────────────────────┐
│ Data Store                         │
│  - Spreadsheet (9 sheets)          │
│  - Google Drive (Documents)        │
└─────────────────────────────────────┘
```

## データモデル設計

### スプレッドシート構成 (9シート)

#### 1. プロジェクト (Projects)
| 列名 | 型 | 説明 |
|------|-----|------|
| プロジェクトID | String | PRJ-YYYYMMDD-9999 |
| 顧客名 | String | 顧客企業名 |
| 作成日 | Date | 作成日時 |
| 作成者メール | String | 作成者のGoogleアカウント |
| 編集者メール | String | CSV形式の編集可能ユーザーリスト |
| 状態 | Enum | 下書き/確定/アーカイブ |

#### 2. ビジョン (Vision)
| 列名 | 型 | 説明 |
|------|-----|------|
| プロジェクトID | String | FK |
| ビジョン本文 | Text | 北極星ビジョン |
| 意思決定ルール | Text | 判断基準 |
| 成功指標 | Text | KPI定義 |
| 備考 | Text | 補足情報 |

#### 3. ユースケース (UseCases)
| 列名 | 型 | 説明 |
|------|-----|------|
| プロジェクトID | String | FK |
| ユースケースID | String | UC-YYYYMMDD-999 |
| 課題 | Text | 解決したい課題 |
| 狙い | Text | ビジネスゴール |
| 想定効果 | Text | 期待される成果 |
| 90日ゴール | Text | 始動段階の目標 |
| スコア | Number | 優先度スコア |
| 優先度 | Number | ランキング |

#### 4. 90日計画 (NinetyDayPlan)
| 列名 | 型 | 説明 |
|------|-----|------|
| プロジェクトID | String | FK |
| ユースケースID | String | FK |
| 体制 | Text | チーム構成 |
| 必要データ | Text | データソース |
| リスク | Text | リスク項目 |
| コミュニケーション計画 | Text | 報告体制 |
| 週次マイルストーン | JSON | 12週分のマイルストーン |

#### 5. 体制RACI (OrganizationRACI)
| 列名 | 型 | 説明 |
|------|-----|------|
| プロジェクトID | String | FK |
| 役割名 | String | CoE/ビジネスデータネットワーク/IT |
| タスク | String | タスク名 |
| 担当者 | String | 担当者名 |
| RACI | Enum | R/A/C/I |

#### 6. 統制 (Governance)
| 列名 | 型 | 説明 |
|------|-----|------|
| プロジェクトID | String | FK |
| 対象 | String | データ分類/領域 |
| 統制モデル | String | 統制方式 |
| 責任者 | String | オーナー |
| ルール | Text | 統制ルール |
| 例外プロセス | Text | 例外処理 |

#### 7. 運営支援 (OperationsSupport)
| 列名 | 型 | 説明 |
|------|-----|------|
| プロジェクトID | String | FK |
| 一次窓口 | String | L1サポート |
| 二次窓口 | String | L2サポート |
| 三次窓口 | String | L3サポート |
| 四次窓口 | String | L4サポート |
| FAQリンク | String | FAQ URL |
| エスカレーション条件 | Text | エスカレ基準 |
| コミュニティ運用 | Text | コミュニティ情報 |

#### 8. 価値 (ValueTracking)
| 列名 | 型 | 説明 |
|------|-----|------|
| プロジェクトID | String | FK |
| ユースケースID | String | FK |
| 定量効果 | Text | 数値効果 |
| 定性効果 | Text | 質的効果 |
| 証跡 | String | Driveリンク |
| 次の投資判断 | Text | 継続/拡大/縮小 |

#### 9. 監査ログ (AuditLog)
| 列名 | 型 | 説明 |
|------|-----|------|
| 日時 | Timestamp | 操作日時 |
| ユーザー | String | 操作者メール |
| 操作 | String | CREATE/UPDATE/DELETE |
| プロジェクトID | String | 対象プロジェクト |
| 詳細 | JSON | 操作詳細 |

## モジュール設計

### MVP フェーズ (優先実装)
1. **Module 1: 北極星ビジョン・メーカー**
   - 成果物: 1枚ビジョン (本文+意思決定ルール+成功指標)

2. **Module 2: 戦略的ユースケース選定 + 90日計画**
   - 成果物: 上位1〜2件の90日計画 (週次マイルストーン含む)
   - 始動段階要件: 2〜3か月で達成、経営に認知、成功指標、週次反復

3. **Module 4: 3本柱組織 + RACI設計**
   - 3本柱: CoE / ビジネスデータネットワーク / IT
   - 成果物: 体制図 + RACI マトリクス

4. **Module 7: 価値実現トラッカー**
   - 成果物: 効果 + 証跡 (Driveリンク) + 次の投資判断

### 拡張フェーズ
5. **Module 3: 6コアプロセス ギャップ診断**
   - 軽量スコア → ロードマップ自動生成 (始動/強化/凌駕のToDo)

6. **Module 5: 最小限の統制設計キット**
   - データ統制の骨格 (オーナー/品質/ライフサイクル/保護・コンプラ/ディスカバリ)

7. **Module 6: 運営支援設計 + 運用台帳**
   - 階層サポート設計 (L1〜L4)

## セキュリティ・アクセス制御

### 認証
- Google Account (Apps Script Web App 標準機能)
- `Session.getActiveUser().getEmail()` で認証済みユーザー取得

### 認可
- プロジェクト単位でアクセス制御
- 編集者メール欄に含まれるユーザーのみ閲覧/編集可能
- サーバーサイドで必ず検証 (クライアント側チェックは補助のみ)

### 監査ログ
- すべてのCRUD操作をログに記録
- ユーザー・日時・操作内容・対象データを保存

## UI/UX設計指針

### 画面構成
- **レイアウト**: 左ナビ + メインコンテンツ (SPA)
- **ナビゲーション**: 目的別入口 (直線フロー禁止)
- **デザイン**: シンプルなB2B SaaS、カード型
- **言語**: 日本語中心、カタカナ過多を避ける
- **レスポンシブ**: モバイルでも崩壊しない

### Google Stitch活用
- プロンプト駆動でUI案とフロントコード生成
- 生成コードをHTML Serviceに移植
- google.script.run で サーバーサイド連携

### ホーム画面 (目的別入口)
1. 北極星を決める (Module 1)
2. 90日計画を作る (Module 2)
3. 体制と責任を確定する (Module 4)
4. 成果を稟議に変換する (Module 7)
5. 追加: ギャップ診断 (Module 3)
6. 追加: 統制 (Module 5)
7. 追加: 運営支援 (Module 6)

## ドキュメント自動生成

### テンプレート (Google Docs)
- ビジョン1枚
- 90日計画
- 体制・RACI
- 統制パック
- 運営支援SOP
- 価値実現レポート
- **稟議書案** (上記を要約: 背景/目的/スコープ/体制/費用仮/リスク/期待効果)

### 生成タイミング
- プロジェクト「確定」時に自動生成
- Drive に保存してリンクをUIに表示

## 開発方針

### 実装順序
1. **Phase 1: 基盤構築**
   - プロジェクト構造セットアップ
   - データモデル実装 (Spreadsheet初期化)
   - 認証・権限制御実装

2. **Phase 2: MVP実装**
   - Module 1, 2, 4, 7 の実装
   - 基本UI (ホーム、プロジェクト一覧、各モジュール画面)
   - ドキュメント自動生成 (稟議書案含む)

3. **Phase 3: 拡張実装**
   - Module 3, 5, 6 の実装
   - UI/UX改善

### コーディング規約
- **ファイル命名**: PascalCase.gs (例: ProjectService.gs)
- **関数命名**: camelCase (例: createProject, getUserProjects)
- **定数**: UPPER_SNAKE_CASE (例: SHEET_NAME_PROJECTS)
- **コメント**: JSDoc形式で必須パラメータと戻り値を記載
- **エラーハンドリング**: try-catch で適切なエラーメッセージを返す

### テスト方針
- Apps Script Editor で手動テスト
- 各サービス層の主要関数をLogger.logで動作確認
- Web App デプロイ後のE2Eテスト

## デプロイ手順

### 初回セットアップ
1. Google Apps Script プロジェクト作成
2. clasp でローカル連携 (オプション)
3. スプレッドシート作成 + シート初期化
4. Script Properties に SPREADSHEET_ID 設定
5. Web App としてデプロイ (Execute as: User accessing the web app)

### 更新デプロイ
1. コード更新
2. clasp push (またはエディタで直接編集)
3. 新しいバージョンとしてデプロイ

## 意思決定ログ

### 2025-12-31: プロジェクト初期設計
- **決定**: スプレッドシート1ファイル = 9シート構成
  - **理由**: Apps Scriptの同時接続制限を考慮し、1ファイルに集約して管理を簡素化

- **決定**: 正規化を最小限にし、冗長性を許容
  - **理由**: 検索性と壊れにくさを優先。小規模データなので正規化のメリットは薄い

- **決定**: 週次マイルストーンはJSON文字列で保存
  - **理由**: Apps Scriptで配列をそのまま保存できないため、JSON.stringify/parse で対応

- **決定**: MVP は Module 1, 2, 4, 7 に絞る
  - **理由**: 最小限で「ビジョン→計画→体制→成果」の一連フローを完成させる

- **決定**: Google Stitch は設計参考に留め、手動でHTML実装
  - **理由**: Stitch生成コードの完全移植は時間がかかるため、まず動くものを優先

## 参考リンク
- [Apps Script Web Apps Guide](https://developers.google.com/apps-script/guides/web)
- [Apps Script HTML Service](https://developers.google.com/apps-script/guides/html)
- [Apps Script Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)
- [Tableau Blueprint v10](https://help.tableau.com/current/blueprint/ja-jp/bp_overview.htm)
