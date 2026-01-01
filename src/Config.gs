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
  var uuid = Utilities.getUuid();
  return 'UC-' + dateStr + '-' + uuid;
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
