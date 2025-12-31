/**
 * Config.gs
 * アプリケーション全体の設定と定数を管理
 */

// ========================================
// スプレッドシート設定
// ========================================

/**
 * スプレッドシートIDを取得
 * Script Properties の SPREADSHEET_ID から取得
 */
function getSpreadsheetId() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('SPREADSHEET_ID が設定されていません。スクリプトプロパティに設定してください。');
  }
  return id;
}

/**
 * スプレッドシートオブジェクトを取得
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(getSpreadsheetId());
}

// ========================================
// シート名定数
// ========================================

var SHEET_NAMES = {
  PROJECTS: 'プロジェクト',
  VISION: 'ビジョン',
  USECASES: 'ユースケース',
  NINETY_DAY_PLAN: '90日計画',
  ORGANIZATION_RACI: '体制RACI',
  GOVERNANCE: '統制',
  OPERATIONS_SUPPORT: '運営支援',
  VALUE: '価値',
  AUDIT_LOG: '監査ログ'
};

// ========================================
// データスキーマ定義
// ========================================

/**
 * 1. プロジェクト シートのスキーマ
 */
var SCHEMA_PROJECTS = {
  sheetName: SHEET_NAMES.PROJECTS,
  headers: [
    'プロジェクトID',
    '顧客名',
    '作成日',
    '作成者メール',
    '編集者メール',
    '状態',
    '最終更新日',
    '最終更新者'
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

/**
 * 2. ビジョン シートのスキーマ
 */
var SCHEMA_VISION = {
  sheetName: SHEET_NAMES.VISION,
  headers: [
    'プロジェクトID',
    'ビジョン本文',
    '意思決定ルール',
    '成功指標',
    '備考',
    '最終更新日'
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

/**
 * 3. ユースケース シートのスキーマ
 */
var SCHEMA_USECASES = {
  sheetName: SHEET_NAMES.USECASES,
  headers: [
    'プロジェクトID',
    'ユースケースID',
    '課題',
    '狙い',
    '想定効果',
    '90日ゴール',
    'スコア',
    '優先度',
    '最終更新日'
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

/**
 * 4. 90日計画 シートのスキーマ
 */
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
    '最終更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    USECASE_ID: 1,
    TEAM_STRUCTURE: 2,
    REQUIRED_DATA: 3,
    RISKS: 4,
    COMMUNICATION_PLAN: 5,
    WEEKLY_MILESTONES: 6, // JSON文字列
    UPDATED_DATE: 7
  }
};

/**
 * 5. 体制RACI シートのスキーマ
 */
var SCHEMA_ORGANIZATION_RACI = {
  sheetName: SHEET_NAMES.ORGANIZATION_RACI,
  headers: [
    'プロジェクトID',
    '3本柱区分',
    'タスク',
    '担当者',
    'RACI',
    '最終更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    PILLAR: 1, // CoE / ビジネスデータネットワーク / IT
    TASK: 2,
    ASSIGNEE: 3,
    RACI: 4, // R / A / C / I
    UPDATED_DATE: 5
  }
};

/**
 * 6. 統制 シートのスキーマ
 */
var SCHEMA_GOVERNANCE = {
  sheetName: SHEET_NAMES.GOVERNANCE,
  headers: [
    'プロジェクトID',
    '対象',
    '統制モデル',
    '責任者',
    'ルール',
    '例外プロセス',
    '最終更新日'
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

/**
 * 7. 運営支援 シートのスキーマ
 */
var SCHEMA_OPERATIONS_SUPPORT = {
  sheetName: SHEET_NAMES.OPERATIONS_SUPPORT,
  headers: [
    'プロジェクトID',
    '一次窓口',
    '二次窓口',
    '三次窓口',
    '四次窓口',
    'FAQリンク',
    'エスカレーション条件',
    'コミュニティ運用',
    '最終更新日'
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

/**
 * 8. 価値 シートのスキーマ
 */
var SCHEMA_VALUE = {
  sheetName: SHEET_NAMES.VALUE,
  headers: [
    'プロジェクトID',
    'ユースケースID',
    '定量効果',
    '定性効果',
    '証跡',
    '次の投資判断',
    '最終更新日'
  ],
  columns: {
    PROJECT_ID: 0,
    USECASE_ID: 1,
    QUANTITATIVE_IMPACT: 2,
    QUALITATIVE_IMPACT: 3,
    EVIDENCE: 4, // Driveリンク
    NEXT_INVESTMENT: 5,
    UPDATED_DATE: 6
  }
};

/**
 * 9. 監査ログ シートのスキーマ
 */
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
    DETAILS: 4 // JSON文字列
  }
};

// ========================================
// 列挙型定数
// ========================================

/**
 * プロジェクト状態
 */
var PROJECT_STATUS = {
  DRAFT: '下書き',
  CONFIRMED: '確定',
  ARCHIVED: 'アーカイブ'
};

/**
 * RACI 区分
 */
var RACI_TYPES = {
  RESPONSIBLE: 'R',
  ACCOUNTABLE: 'A',
  CONSULTED: 'C',
  INFORMED: 'I'
};

/**
 * 3本柱組織
 */
var THREE_PILLARS = {
  COE: 'CoE',
  BUSINESS_DATA_NETWORK: 'ビジネスデータネットワーク',
  IT: 'IT'
};

/**
 * 操作種別
 */
var OPERATION_TYPES = {
  CREATE: 'CREATE',
  UPDATE: 'UPDATE',
  DELETE: 'DELETE',
  VIEW: 'VIEW'
};

// ========================================
// ID生成関数
// ========================================

/**
 * プロジェクトIDを生成
 * フォーマット: PRJ-YYYYMMDD-9999
 */
function generateProjectId() {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  var randomNum = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return 'PRJ-' + dateStr + '-' + randomNum;
}

/**
 * ユースケースIDを生成
 * フォーマット: UC-YYYYMMDD-999
 */
function generateUsecaseId() {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  var randomNum = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  return 'UC-' + dateStr + '-' + randomNum;
}

// ========================================
// ユーティリティ関数
// ========================================

/**
 * 現在のユーザーメールアドレスを取得
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * 現在の日時を取得
 */
function getCurrentTimestamp() {
  return new Date();
}

/**
 * JSON文字列を安全にパース
 */
function safeJsonParse(jsonString, defaultValue) {
  try {
    return JSON.parse(jsonString);
  } catch (e) {
    Logger.log('JSON parse error: ' + e.message);
    return defaultValue || null;
  }
}

/**
 * オブジェクトをJSON文字列に変換
 */
function safeJsonStringify(obj) {
  try {
    return JSON.stringify(obj);
  } catch (e) {
    Logger.log('JSON stringify error: ' + e.message);
    return '';
  }
}
