/**
 * Config.gs
 * Central configuration and shared utilities.
 */

// ========================================
// Spreadsheet configuration
// ========================================

/**
 * Returns the spreadsheet ID from Script Properties.
 * Script Properties: SPREADSHEET_ID
 */
function getSpreadsheetId() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('スクリプトプロパティにSPREADSHEET_IDが設定されていません。');
  }
  return id;
}

/**
 * Returns the Spreadsheet instance.
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(getSpreadsheetId());
}

/**
 * One-time setup helper: set Script Property `SPREADSHEET_ID`.
 * Run from the Apps Script editor (or call from another setup function).
 * @param {string} spreadsheetId
 */
function setSpreadsheetId(spreadsheetId) {
  if (!spreadsheetId) {
    throw new Error('spreadsheetId is required. Set Script Property `SPREADSHEET_ID` or pass an ID.');
  }
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
}

/**
 * One-time setup helper: set `SPREADSHEET_ID` and initialize all sheets.
 * @param {string} spreadsheetId
 * @return {{success: boolean, message: string}}
 */
function setupSpreadsheet(spreadsheetId) {
  var id = spreadsheetId || PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('spreadsheetId is required. Pass an ID to setupSpreadsheet(id) or set Script Property `SPREADSHEET_ID`.');
  }
  setSpreadsheetId(id);
  return initializeAllSheets();
}

/**
 * One-time setup helper (no-args): prompts for SPREADSHEET_ID then initializes sheets.
 * Intended to be run from the Apps Script editor.
 * @return {{success: boolean, message: string}}
 */
function setupSpreadsheetInteractive() {
  var id = Browser.inputBox(
    'Setup Spreadsheet',
    'Paste your Spreadsheet ID (URL between /d/ and /edit):',
    Browser.Buttons.OK_CANCEL
  );
  if (!id || id === 'cancel') {
    throw new Error('Cancelled.');
  }
  return setupSpreadsheet(String(id).trim());
}

/**
 * One-time setup helper: create a new spreadsheet, bind it, and initialize all sheets.
 * @param {string} spreadsheetName
 * @return {{success: boolean, spreadsheetId: string, url: string}}
 */
function createNewSpreadsheet(spreadsheetName) {
  var name = spreadsheetName || 'Tableau Blueprint DB';
  var ss = SpreadsheetApp.create(name);

  setSpreadsheetId(ss.getId());
  initializeAllSheets();

  return { success: true, spreadsheetId: ss.getId(), url: ss.getUrl() };
}

// ========================================
// Sheet names
// ========================================

var SHEET_NAMES = {
  PROJECTS: 'プロジェクト',
  VISION: 'ビジョン',
  USECASES: 'ユースケース',
  NINETY_DAY_PLAN: '90日計画',
  ORGANIZATION_RACI: '体制RACI',
  CORE_PROCESS: 'コアプロセス',
  GOVERNANCE: '統制',
  OPERATIONS_SUPPORT: '運営支援',
  VALUE: '価値',
  AUDIT_LOG: '監査ログ'
};

// ========================================
// Schema definitions
// ========================================

var SCHEMA_PROJECTS = {
  sheetName: SHEET_NAMES.PROJECTS,
  headers: [
    'プロジェクトID',
    '顧客名',
    '作成日',
    '作成者メール',
    '編集者メール',
    '状態',
    '更新日',
    '更新者メール'
  ],
  columns: {
    PROJECT_ID: 0,
    CUSTOMER_NAME: 1,
    CREATED_DATE: 2,
    CREATED_BY: 3,
    EDITORS: 4,
    STATUS: 5,
    UPDATED_DATE: 6,
    UPDATED_BY: 7
  }
};

var SCHEMA_VISION = {
  sheetName: SHEET_NAMES.VISION,
  headers: [
    'プロジェクトID',
    'ビジョン本文',
    '意思決定ルール',
    '成功指標',
    '備考',
    '更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    VISION_TEXT: 1,
    DECISION_RULES: 2,
    SUCCESS_METRICS: 3,
    NOTES: 4,
    UPDATED_DATE: 5
  }
};

var SCHEMA_USECASES = {
  sheetName: SHEET_NAMES.USECASES,
  headers: [
    'プロジェクトID',
    'ユースケースID',
    '課題',
    '目標',
    '想定効果',
    '90日ゴール',
    'スコア',
    '優先度',
    '更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    USECASE_ID: 1,
    CHALLENGE: 2,
    GOAL: 3,
    EXPECTED_IMPACT: 4,
    NINETY_DAY_GOAL: 5,
    SCORE: 6,
    PRIORITY: 7,
    UPDATED_DATE: 8
  }
};

var SCHEMA_NINETY_DAY_PLAN = {
  sheetName: SHEET_NAMES.NINETY_DAY_PLAN,
  headers: [
    'プロジェクトID',
    'ユースケースID',
    '体制',
    '必要データ',
    'リスク',
    'コミュニケーション計画',
    '週次マイルストーン',
    '更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    USECASE_ID: 1,
    TEAM_STRUCTURE: 2,
    REQUIRED_DATA: 3,
    RISKS: 4,
    COMMUNICATION_PLAN: 5,
    WEEKLY_MILESTONES: 6,
    UPDATED_DATE: 7
  }
};

var SCHEMA_ORGANIZATION_RACI = {
  sheetName: SHEET_NAMES.ORGANIZATION_RACI,
  headers: [
    'プロジェクトID',
    '3本柱',
    'タスク',
    '担当者',
    'RACI',
    '更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    PILLAR: 1,
    TASK: 2,
    ASSIGNEE: 3,
    RACI: 4,
    UPDATED_DATE: 5
  }
};

var SCHEMA_GOVERNANCE = {
  sheetName: SHEET_NAMES.GOVERNANCE,
  headers: [
    'プロジェクトID',
    '対象',
    'モデル',
    '責任者',
    'ルール',
    '例外プロセス',
    '更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    TARGET: 1,
    MODEL: 2,
    OWNER: 3,
    RULES: 4,
    EXCEPTION_PROCESS: 5,
    UPDATED_DATE: 6
  }
};

var SCHEMA_OPERATIONS_SUPPORT = {
  sheetName: SHEET_NAMES.OPERATIONS_SUPPORT,
  headers: [
    'プロジェクトID',
    'L1サポート',
    'L2サポート',
    'L3サポート',
    'L4サポート',
    'FAQリンク',
    'エスカレーション基準',
    'コミュニティ運用',
    '更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    L1_SUPPORT: 1,
    L2_SUPPORT: 2,
    L3_SUPPORT: 3,
    L4_SUPPORT: 4,
    FAQ_LINK: 5,
    ESCALATION_CRITERIA: 6,
    COMMUNITY_OPS: 7,
    UPDATED_DATE: 8
  }
};

var SCHEMA_VALUE = {
  sheetName: SHEET_NAMES.VALUE,
  headers: [
    'プロジェクトID',
    'ユースケースID',
    '定量効果',
    '定性効果',
    '証跡',
    '次の投資判断',
    '更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    USECASE_ID: 1,
    QUANTITATIVE_IMPACT: 2,
    QUALITATIVE_IMPACT: 3,
    EVIDENCE: 4,
    NEXT_INVESTMENT: 5,
    UPDATED_DATE: 6
  }
};

var SCHEMA_AUDIT_LOG = {
  sheetName: SHEET_NAMES.AUDIT_LOG,
  headers: [
    '日時',
    'ユーザー',
    '操作',
    'プロジェクトID',
    '詳細'
  ],
  columns: {
    TIMESTAMP: 0,
    USER: 1,
    OPERATION: 2,
    PROJECT_ID: 3,
    DETAILS: 4
  }
};

var SCHEMA_CORE_PROCESS = {
  sheetName: SHEET_NAMES.CORE_PROCESS,
  headers: [
    'プロジェクトID',
    'アジリティスコア',
    'スキルスコア',
    'データ品質スコア',
    '信頼スコア',
    '運用効率スコア',
    'コミュニティスコア',
    'アジリティコメント',
    'スキルコメント',
    'データ品質コメント',
    '信頼コメント',
    '運用効率コメント',
    'コミュニティコメント',
    '更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    AGILITY_SCORE: 1,
    SKILLS_SCORE: 2,
    DATA_QUALITY_SCORE: 3,
    TRUST_SCORE: 4,
    OPERATIONAL_EFFICIENCY_SCORE: 5,
    COMMUNITY_SCORE: 6,
    AGILITY_COMMENT: 7,
    SKILLS_COMMENT: 8,
    DATA_QUALITY_COMMENT: 9,
    TRUST_COMMENT: 10,
    OPERATIONAL_EFFICIENCY_COMMENT: 11,
    COMMUNITY_COMMENT: 12,
    UPDATED_DATE: 13
  }
};

// ========================================
// Enumerations
// ========================================

var PROJECT_STATUS = {
  DRAFT: 'Draft',
  CONFIRMED: 'Confirmed',
  ARCHIVED: 'Archived'
};

var RACI_TYPES = {
  RESPONSIBLE: 'R',
  ACCOUNTABLE: 'A',
  CONSULTED: 'C',
  INFORMED: 'I'
};

var THREE_PILLARS = {
  COE: 'CoE',
  BUSINESS_DATA_NETWORK: 'Biz Data Network',
  IT: 'IT'
};

var OPERATION_TYPES = {
  CREATE: 'CREATE',
  UPDATE: 'UPDATE',
  DELETE: 'DELETE',
  VIEW: 'VIEW'
};

// ========================================
// Maturity Levels (based on Tableau Blueprint)
// ========================================

var MATURITY_LEVELS = {
  IGNITE: 'ignite',      // 始動段階: 平均スコア 1-2
  STRENGTHEN: 'strengthen', // 強化段階: 平均スコア 2-3.5
  MASTER: 'master'       // 凌駕段階: 平均スコア 3.5-5
};

var MATURITY_LABELS = {
  ignite: '始動段階',
  strengthen: '強化段階',
  master: '凌駕段階'
};

var MATURITY_DESCRIPTIONS = {
  ignite: 'データ活用を始めたばかりの段階。基盤構築と小規模な成功体験が重要です。',
  strengthen: 'データ活用が定着しつつある段階。標準化とスケールアップが課題です。',
  master: 'データドリブンな文化が確立している段階。高度な分析と自律運用が目標です。'
};

/**
 * Calculate maturity level from core process scores.
 * @param {Object} coreProcess - Core process data with scores
 * @return {Object} { level: string, averageScore: number, label: string, description: string }
 */
function calculateMaturityLevel(coreProcess) {
  if (!coreProcess) {
    return {
      level: MATURITY_LEVELS.IGNITE,
      averageScore: 1,
      label: MATURITY_LABELS.ignite,
      description: MATURITY_DESCRIPTIONS.ignite
    };
  }

  var scores = [
    parseInt(coreProcess.agilityScore) || 1,
    parseInt(coreProcess.skillsScore) || 1,
    parseInt(coreProcess.dataQualityScore) || 1,
    parseInt(coreProcess.trustScore) || 1,
    parseInt(coreProcess.operationalEfficiencyScore) || 1,
    parseInt(coreProcess.communityScore) || 1
  ];

  var sum = scores.reduce(function(a, b) { return a + b; }, 0);
  var avg = sum / scores.length;

  var level;
  if (avg >= 3.5) {
    level = MATURITY_LEVELS.MASTER;
  } else if (avg >= 2) {
    level = MATURITY_LEVELS.STRENGTHEN;
  } else {
    level = MATURITY_LEVELS.IGNITE;
  }

  return {
    level: level,
    averageScore: Math.round(avg * 10) / 10,
    label: MATURITY_LABELS[level],
    description: MATURITY_DESCRIPTIONS[level]
  };
}

/**
 * Get 90-day plan templates based on maturity level.
 * @param {string} maturityLevel - ignite, strengthen, or master
 * @return {Object} Phase templates with objectives, milestones, and success criteria
 */
function get90DayPlanTemplates(maturityLevel) {
  var templates = {
    ignite: {
      phases: {
        ignite: {
          objective: '基盤構築と初期ダッシュボード作成。パイロットユーザーでの検証開始。',
          milestones: [
            'Week1: キックオフ、要件ヒアリング完了',
            'Week2: データ接続、品質検証',
            'Week3: ダッシュボード v0.5 レビュー',
            'Week4: パイロット開始、フィードバック収集'
          ],
          successCriteria: 'パイロットユーザー全員がダッシュボードにアクセスし、1回以上業務で活用'
        },
        strengthen: {
          objective: 'フィードバック反映と利用者拡大',
          milestones: [
            'Week5: v1.0 リリース、機能改善',
            'Week6: 追加ユーザー向けトレーニング',
            'Week7: 利用状況モニタリング開始',
            'Week8: 中間レビュー、効果測定'
          ],
          successCriteria: '拡大ユーザーが週3回以上アクセス、初期効果を確認'
        },
        establish: {
          objective: '展開拡大と運用体制の確立',
          milestones: [
            'Week9: 全社展開計画策定',
            'Week10: 残りユーザーへの展開',
            'Week11: 運用マニュアル・FAQ整備',
            'Week12: 効果測定、次期計画策定'
          ],
          successCriteria: '対象ユーザー全員が利用開始、目標効果の50%を達成'
        }
      }
    },
    strengthen: {
      phases: {
        ignite: {
          objective: 'ガバナンス基盤の整備とデータ品質向上',
          milestones: [
            'Week1: 既存ダッシュボード/データソース棚卸し',
            'Week2: データカタログ設計、品質ルール策定',
            'Week3: 認定データソースの定義',
            'Week4: ガバナンスポリシー初版リリース'
          ],
          successCriteria: 'データカタログに主要データソースが登録され、品質基準が明文化'
        },
        strengthen: {
          objective: '標準化とベストプラクティスの横展開',
          milestones: [
            'Week5: ダッシュボードテンプレート作成',
            'Week6: 部門横断KPIの統一定義',
            'Week7: パワーユーザー育成プログラム開始',
            'Week8: コミュニティ活動の活性化'
          ],
          successCriteria: 'テンプレート利用率50%以上、パワーユーザー10名以上認定'
        },
        establish: {
          objective: 'セルフサービス化と自律運用体制',
          milestones: [
            'Week9: セルフサービス分析ガイドライン整備',
            'Week10: CoEサポートモデル確立',
            'Week11: 効果測定フレームワーク導入',
            'Week12: 次期ロードマップ策定'
          ],
          successCriteria: 'セルフサービス分析件数が月10件以上、CoE負荷が削減'
        }
      }
    },
    master: {
      phases: {
        ignite: {
          objective: '高度分析/AI活用の企画と基盤整備',
          milestones: [
            'Week1: 高度分析ユースケースの特定',
            'Week2: AI/MLプラットフォーム評価',
            'Week3: データサイエンスチーム体制設計',
            'Week4: PoC計画策定、データ準備'
          ],
          successCriteria: '高度分析ユースケースが3件以上特定され、PoC計画が承認'
        },
        strengthen: {
          objective: '予測分析/機械学習モデルのPoC実施',
          milestones: [
            'Week5: 予測モデル開発開始',
            'Week6: モデル検証、精度チューニング',
            'Week7: ビジネスユーザーへのデモ',
            'Week8: モデル運用計画策定'
          ],
          successCriteria: '予測モデルの精度が目標値を達成、ビジネス価値が確認'
        },
        establish: {
          objective: 'AIインサイトの業務組み込みと拡大',
          milestones: [
            'Week9: 本番環境へのモデルデプロイ',
            'Week10: 業務プロセスへの組み込み',
            'Week11: 効果測定、モニタリング体制確立',
            'Week12: 横展開計画、次期AI施策策定'
          ],
          successCriteria: 'AIモデルが業務で活用され、定量効果が確認'
        }
      }
    }
  };

  return templates[maturityLevel] || templates.ignite;
}

// ========================================
// ID helpers
// ========================================

function generateProjectId() {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  var randomNum = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return 'PRJ-' + dateStr + '-' + randomNum;
}

function generateUsecaseId() {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  var randomNum = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return 'UC-' + dateStr + '-' + randomNum;
}

// ========================================
// Utilities
// ========================================

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getCurrentTimestamp() {
  return new Date();
}

function safeJsonParse(jsonString, defaultValue) {
  try {
    return JSON.parse(jsonString);
  } catch (e) {
    Logger.log('JSON parse error: ' + e.message);
    return defaultValue || null;
  }
}

function safeJsonStringify(obj) {
  try {
    return JSON.stringify(obj);
  } catch (e) {
    Logger.log('JSON stringify error: ' + e.message);
    return '';
  }
}
