/**
 * SpreadsheetRepository.gs
 * スプレッドシート（簡易DB）へのアクセスを担当するリポジトリ層
 */

// ========================================
// シート初期化
// ========================================

/**
 * すべてのシートを初期化
 * シートが存在しない場合は作成し、ヘッダー行を設定
 */
function initializeAllSheets() {
  var ss = getSpreadsheet();
  var schemas = [
    SCHEMA_PROJECTS,
    SCHEMA_VISION,
    SCHEMA_USECASES,
    SCHEMA_NINETY_DAY_PLAN,
    SCHEMA_ORGANIZATION_RACI,
    SCHEMA_GOVERNANCE,
    SCHEMA_OPERATIONS_SUPPORT,
    SCHEMA_VALUE,
    SCHEMA_AUDIT_LOG
  ];

  schemas.forEach(function(schema) {
    initializeSheet(ss, schema);
  });

  Logger.log('すべてのシートを初期化しました');
  return { success: true, message: 'すべてのシートを初期化しました' };
}

/**
 * 個別シートを初期化
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Object} schema - シートスキーマ
 */
function initializeSheet(ss, schema) {
  var sheet = ss.getSheetByName(schema.sheetName);

  if (!sheet) {
    // シートが存在しない場合は作成
    sheet = ss.insertSheet(schema.sheetName);
    Logger.log('シート「' + schema.sheetName + '」を作成しました');
  }

  // ヘッダー行をチェック
  var headers = sheet.getRange(1, 1, 1, schema.headers.length).getValues()[0];
  var headersEmpty = headers.every(function(h) { return h === ''; });

  if (headersEmpty) {
    // ヘッダー行を設定
    sheet.getRange(1, 1, 1, schema.headers.length).setValues([schema.headers]);
    sheet.getRange(1, 1, 1, schema.headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('シート「' + schema.sheetName + '」にヘッダーを設定しました');
  }
}

/**
 * シートオブジェクトを取得
 * @param {string} sheetName - シート名
 * @return {Sheet} シートオブジェクト
 */
function getSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シート「' + sheetName + '」が見つかりません');
  }
  return sheet;
}

// ========================================
// 汎用CRUD操作
// ========================================

/**
 * データを追加
 * @param {string} sheetName - シート名
 * @param {Array} rowData - 追加する行データ
 * @return {number} 追加された行番号
 */
function appendRow(sheetName, rowData) {
  var sheet = getSheet(sheetName);
  sheet.appendRow(rowData);
  return sheet.getLastRow();
}

/**
 * データを更新
 * @param {string} sheetName - シート名
 * @param {number} rowNumber - 更新する行番号 (1-indexed)
 * @param {Array} rowData - 更新するデータ
 */
function updateRow(sheetName, rowNumber, rowData) {
  var sheet = getSheet(sheetName);
  sheet.getRange(rowNumber, 1, 1, rowData.length).setValues([rowData]);
}

/**
 * データを削除
 * @param {string} sheetName - シート名
 * @param {number} rowNumber - 削除する行番号 (1-indexed)
 */
function deleteRow(sheetName, rowNumber) {
  var sheet = getSheet(sheetName);
  sheet.deleteRow(rowNumber);
}

/**
 * すべてのデータを取得
 * @param {string} sheetName - シート名
 * @return {Array<Array>} すべてのデータ行 (ヘッダー除く)
 */
function getAllRows(sheetName) {
  var sheet = getSheet(sheetName);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return []; // ヘッダーのみの場合は空配列
  }

  var lastCol = sheet.getLastColumn();
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

/**
 * 条件に一致する行を検索
 * @param {string} sheetName - シート名
 * @param {number} columnIndex - 検索対象の列インデックス (0-indexed)
 * @param {*} value - 検索する値
 * @return {Array} { rowNumber: 行番号, data: データ配列 }
 */
function findRows(sheetName, columnIndex, value) {
  var rows = getAllRows(sheetName);
  var results = [];

  rows.forEach(function(row, index) {
    if (row[columnIndex] === value) {
      results.push({
        rowNumber: index + 2, // ヘッダー分+1、0-indexed→1-indexed
        data: row
      });
    }
  });

  return results;
}

/**
 * 条件に一致する最初の行を取得
 * @param {string} sheetName - シート名
 * @param {number} columnIndex - 検索対象の列インデックス (0-indexed)
 * @param {*} value - 検索する値
 * @return {Object|null} { rowNumber: 行番号, data: データ配列 } または null
 */
function findFirstRow(sheetName, columnIndex, value) {
  var results = findRows(sheetName, columnIndex, value);
  return results.length > 0 ? results[0] : null;
}

// ========================================
// プロジェクト操作
// ========================================

/**
 * プロジェクトを作成
 * @param {string} customerName - 顧客名
 * @param {string} userEmail - 作成者メール
 * @return {Object} 作成されたプロジェクト情報
 */
function createProject(customerName, userEmail) {
  var projectId = generateProjectId();
  var now = getCurrentTimestamp();

  var rowData = [];
  rowData[SCHEMA_PROJECTS.columns.PROJECT_ID] = projectId;
  rowData[SCHEMA_PROJECTS.columns.CUSTOMER_NAME] = customerName;
  rowData[SCHEMA_PROJECTS.columns.CREATED_DATE] = now;
  rowData[SCHEMA_PROJECTS.columns.CREATED_BY] = userEmail;
  rowData[SCHEMA_PROJECTS.columns.EDITORS] = userEmail; // 初期は作成者のみ
  rowData[SCHEMA_PROJECTS.columns.STATUS] = PROJECT_STATUS.DRAFT;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_DATE] = now;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_BY] = userEmail;

  appendRow(SCHEMA_PROJECTS.sheetName, rowData);

  // 監査ログ
  logAudit(userEmail, OPERATION_TYPES.CREATE, projectId, { customerName: customerName });

  return {
    projectId: projectId,
    customerName: customerName,
    status: PROJECT_STATUS.DRAFT,
    createdDate: now
  };
}

/**
 * プロジェクトを取得
 * @param {string} projectId - プロジェクトID
 * @return {Object|null} プロジェクト情報
 */
function getProject(projectId) {
  var result = findFirstRow(
    SCHEMA_PROJECTS.sheetName,
    SCHEMA_PROJECTS.columns.PROJECT_ID,
    projectId
  );

  if (!result) return null;

  var data = result.data;
  return {
    projectId: data[SCHEMA_PROJECTS.columns.PROJECT_ID],
    customerName: data[SCHEMA_PROJECTS.columns.CUSTOMER_NAME],
    createdDate: data[SCHEMA_PROJECTS.columns.CREATED_DATE],
    createdBy: data[SCHEMA_PROJECTS.columns.CREATED_BY],
    editors: data[SCHEMA_PROJECTS.columns.EDITORS],
    status: data[SCHEMA_PROJECTS.columns.STATUS],
    updatedDate: data[SCHEMA_PROJECTS.columns.UPDATED_DATE],
    updatedBy: data[SCHEMA_PROJECTS.columns.UPDATED_BY],
    _rowNumber: result.rowNumber
  };
}

/**
 * ユーザーがアクセス可能なプロジェクト一覧を取得
 * @param {string} userEmail - ユーザーメール
 * @return {Array<Object>} プロジェクト一覧
 */
function getUserProjects(userEmail) {
  var rows = getAllRows(SCHEMA_PROJECTS.sheetName);
  var projects = [];

  rows.forEach(function(row) {
    var editors = row[SCHEMA_PROJECTS.columns.EDITORS] || '';
    var editorList = editors.split(',').map(function(e) { return e.trim(); });

    // 作成者または編集者リストに含まれる場合のみ追加
    if (row[SCHEMA_PROJECTS.columns.CREATED_BY] === userEmail ||
        editorList.indexOf(userEmail) !== -1) {
      projects.push({
        projectId: row[SCHEMA_PROJECTS.columns.PROJECT_ID],
        customerName: row[SCHEMA_PROJECTS.columns.CUSTOMER_NAME],
        status: row[SCHEMA_PROJECTS.columns.STATUS],
        createdDate: row[SCHEMA_PROJECTS.columns.CREATED_DATE],
        updatedDate: row[SCHEMA_PROJECTS.columns.UPDATED_DATE]
      });
    }
  });

  return projects;
}

/**
 * プロジェクトの状態を更新
 * @param {string} projectId - プロジェクトID
 * @param {string} status - 新しい状態
 * @param {string} userEmail - 更新者メール
 */
function updateProjectStatus(projectId, status, userEmail) {
  var project = getProject(projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません: ' + projectId);
  }

  var rowData = [];
  rowData[SCHEMA_PROJECTS.columns.PROJECT_ID] = project.projectId;
  rowData[SCHEMA_PROJECTS.columns.CUSTOMER_NAME] = project.customerName;
  rowData[SCHEMA_PROJECTS.columns.CREATED_DATE] = project.createdDate;
  rowData[SCHEMA_PROJECTS.columns.CREATED_BY] = project.createdBy;
  rowData[SCHEMA_PROJECTS.columns.EDITORS] = project.editors;
  rowData[SCHEMA_PROJECTS.columns.STATUS] = status;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_DATE] = getCurrentTimestamp();
  rowData[SCHEMA_PROJECTS.columns.UPDATED_BY] = userEmail;

  updateRow(SCHEMA_PROJECTS.sheetName, project._rowNumber, rowData);

  // 監査ログ
  logAudit(userEmail, OPERATION_TYPES.UPDATE, projectId, {
    field: 'status',
    oldValue: project.status,
    newValue: status
  });
}

// ========================================
// ビジョン操作
// ========================================

/**
 * ビジョンを保存
 * @param {Object} visionData - ビジョンデータ
 */
function saveVision(visionData) {
  var existing = findFirstRow(
    SCHEMA_VISION.sheetName,
    SCHEMA_VISION.columns.PROJECT_ID,
    visionData.projectId
  );

  var rowData = [];
  rowData[SCHEMA_VISION.columns.PROJECT_ID] = visionData.projectId;
  rowData[SCHEMA_VISION.columns.VISION_TEXT] = visionData.visionText || '';
  rowData[SCHEMA_VISION.columns.DECISION_RULES] = visionData.decisionRules || '';
  rowData[SCHEMA_VISION.columns.SUCCESS_METRICS] = visionData.successMetrics || '';
  rowData[SCHEMA_VISION.columns.NOTES] = visionData.notes || '';
  rowData[SCHEMA_VISION.columns.UPDATED_DATE] = getCurrentTimestamp();

  if (existing) {
    updateRow(SCHEMA_VISION.sheetName, existing.rowNumber, rowData);
    logAudit(visionData.userEmail, OPERATION_TYPES.UPDATE, visionData.projectId, { module: 'vision' });
  } else {
    appendRow(SCHEMA_VISION.sheetName, rowData);
    logAudit(visionData.userEmail, OPERATION_TYPES.CREATE, visionData.projectId, { module: 'vision' });
  }
}

/**
 * ビジョンを取得
 * @param {string} projectId - プロジェクトID
 * @return {Object|null} ビジョンデータ
 */
function getVision(projectId) {
  var result = findFirstRow(
    SCHEMA_VISION.sheetName,
    SCHEMA_VISION.columns.PROJECT_ID,
    projectId
  );

  if (!result) return null;

  var data = result.data;
  return {
    projectId: data[SCHEMA_VISION.columns.PROJECT_ID],
    visionText: data[SCHEMA_VISION.columns.VISION_TEXT],
    decisionRules: data[SCHEMA_VISION.columns.DECISION_RULES],
    successMetrics: data[SCHEMA_VISION.columns.SUCCESS_METRICS],
    notes: data[SCHEMA_VISION.columns.NOTES],
    updatedDate: data[SCHEMA_VISION.columns.UPDATED_DATE]
  };
}

// ========================================
// ユースケース操作
// ========================================

/**
 * ユースケースを追加
 * @param {Object} usecaseData - ユースケースデータ
 * @return {string} 生成されたユースケースID
 */
function addUsecase(usecaseData) {
  var usecaseId = generateUsecaseId();

  var rowData = [];
  rowData[SCHEMA_USECASES.columns.PROJECT_ID] = usecaseData.projectId;
  rowData[SCHEMA_USECASES.columns.USECASE_ID] = usecaseId;
  rowData[SCHEMA_USECASES.columns.CHALLENGE] = usecaseData.challenge || '';
  rowData[SCHEMA_USECASES.columns.GOAL] = usecaseData.goal || '';
  rowData[SCHEMA_USECASES.columns.EXPECTED_IMPACT] = usecaseData.expectedImpact || '';
  rowData[SCHEMA_USECASES.columns.NINETY_DAY_GOAL] = usecaseData.ninetyDayGoal || '';
  rowData[SCHEMA_USECASES.columns.SCORE] = usecaseData.score || 0;
  rowData[SCHEMA_USECASES.columns.PRIORITY] = usecaseData.priority || 0;
  rowData[SCHEMA_USECASES.columns.UPDATED_DATE] = getCurrentTimestamp();

  appendRow(SCHEMA_USECASES.sheetName, rowData);

  logAudit(usecaseData.userEmail, OPERATION_TYPES.CREATE, usecaseData.projectId, {
    module: 'usecase',
    usecaseId: usecaseId
  });

  return usecaseId;
}

/**
 * プロジェクトのユースケース一覧を取得
 * @param {string} projectId - プロジェクトID
 * @return {Array<Object>} ユースケース一覧
 */
function getUsecases(projectId) {
  var results = findRows(
    SCHEMA_USECASES.sheetName,
    SCHEMA_USECASES.columns.PROJECT_ID,
    projectId
  );

  return results.map(function(result) {
    var data = result.data;
    return {
      usecaseId: data[SCHEMA_USECASES.columns.USECASE_ID],
      projectId: data[SCHEMA_USECASES.columns.PROJECT_ID],
      challenge: data[SCHEMA_USECASES.columns.CHALLENGE],
      goal: data[SCHEMA_USECASES.columns.GOAL],
      expectedImpact: data[SCHEMA_USECASES.columns.EXPECTED_IMPACT],
      ninetyDayGoal: data[SCHEMA_USECASES.columns.NINETY_DAY_GOAL],
      score: data[SCHEMA_USECASES.columns.SCORE],
      priority: data[SCHEMA_USECASES.columns.PRIORITY],
      updatedDate: data[SCHEMA_USECASES.columns.UPDATED_DATE],
      _rowNumber: result.rowNumber
    };
  });
}

// ========================================
// 90日計画操作
// ========================================

/**
 * 90日計画を保存
 * @param {Object} planData - 90日計画データ
 */
function saveNinetyDayPlan(planData) {
  var existing = findFirstRow(
    SCHEMA_NINETY_DAY_PLAN.sheetName,
    SCHEMA_NINETY_DAY_PLAN.columns.USECASE_ID,
    planData.usecaseId
  );

  var rowData = [];
  rowData[SCHEMA_NINETY_DAY_PLAN.columns.PROJECT_ID] = planData.projectId;
  rowData[SCHEMA_NINETY_DAY_PLAN.columns.USECASE_ID] = planData.usecaseId;
  rowData[SCHEMA_NINETY_DAY_PLAN.columns.TEAM_STRUCTURE] = planData.teamStructure || '';
  rowData[SCHEMA_NINETY_DAY_PLAN.columns.REQUIRED_DATA] = planData.requiredData || '';
  rowData[SCHEMA_NINETY_DAY_PLAN.columns.RISKS] = planData.risks || '';
  rowData[SCHEMA_NINETY_DAY_PLAN.columns.COMMUNICATION_PLAN] = planData.communicationPlan || '';
  rowData[SCHEMA_NINETY_DAY_PLAN.columns.WEEKLY_MILESTONES] = safeJsonStringify(planData.weeklyMilestones || []);
  rowData[SCHEMA_NINETY_DAY_PLAN.columns.UPDATED_DATE] = getCurrentTimestamp();

  if (existing) {
    updateRow(SCHEMA_NINETY_DAY_PLAN.sheetName, existing.rowNumber, rowData);
    logAudit(planData.userEmail, OPERATION_TYPES.UPDATE, planData.projectId, {
      module: '90day-plan',
      usecaseId: planData.usecaseId
    });
  } else {
    appendRow(SCHEMA_NINETY_DAY_PLAN.sheetName, rowData);
    logAudit(planData.userEmail, OPERATION_TYPES.CREATE, planData.projectId, {
      module: '90day-plan',
      usecaseId: planData.usecaseId
    });
  }
}

/**
 * 90日計画を取得
 * @param {string} usecaseId - ユースケースID
 * @return {Object|null} 90日計画データ
 */
function getNinetyDayPlan(usecaseId) {
  var result = findFirstRow(
    SCHEMA_NINETY_DAY_PLAN.sheetName,
    SCHEMA_NINETY_DAY_PLAN.columns.USECASE_ID,
    usecaseId
  );

  if (!result) return null;

  var data = result.data;
  return {
    projectId: data[SCHEMA_NINETY_DAY_PLAN.columns.PROJECT_ID],
    usecaseId: data[SCHEMA_NINETY_DAY_PLAN.columns.USECASE_ID],
    teamStructure: data[SCHEMA_NINETY_DAY_PLAN.columns.TEAM_STRUCTURE],
    requiredData: data[SCHEMA_NINETY_DAY_PLAN.columns.REQUIRED_DATA],
    risks: data[SCHEMA_NINETY_DAY_PLAN.columns.RISKS],
    communicationPlan: data[SCHEMA_NINETY_DAY_PLAN.columns.COMMUNICATION_PLAN],
    weeklyMilestones: safeJsonParse(data[SCHEMA_NINETY_DAY_PLAN.columns.WEEKLY_MILESTONES], []),
    updatedDate: data[SCHEMA_NINETY_DAY_PLAN.columns.UPDATED_DATE]
  };
}

// ========================================
// 体制RACI操作
// ========================================

/**
 * RACIエントリを保存（一括）
 * @param {string} projectId - プロジェクトID
 * @param {Array<Object>} raciEntries - RACIエントリ配列
 * @param {string} userEmail - 更新者メール
 */
function saveRACIEntries(projectId, raciEntries, userEmail) {
  // 既存のエントリを削除
  var existingRows = findRows(
    SCHEMA_ORGANIZATION_RACI.sheetName,
    SCHEMA_ORGANIZATION_RACI.columns.PROJECT_ID,
    projectId
  );

  // 逆順で削除（行番号がずれないように）
  existingRows.reverse().forEach(function(row) {
    deleteRow(SCHEMA_ORGANIZATION_RACI.sheetName, row.rowNumber);
  });

  // 新しいエントリを追加
  raciEntries.forEach(function(entry) {
    var rowData = [];
    rowData[SCHEMA_ORGANIZATION_RACI.columns.PROJECT_ID] = projectId;
    rowData[SCHEMA_ORGANIZATION_RACI.columns.PILLAR] = entry.pillar;
    rowData[SCHEMA_ORGANIZATION_RACI.columns.TASK] = entry.task;
    rowData[SCHEMA_ORGANIZATION_RACI.columns.ASSIGNEE] = entry.assignee;
    rowData[SCHEMA_ORGANIZATION_RACI.columns.RACI] = entry.raci;
    rowData[SCHEMA_ORGANIZATION_RACI.columns.UPDATED_DATE] = getCurrentTimestamp();

    appendRow(SCHEMA_ORGANIZATION_RACI.sheetName, rowData);
  });

  logAudit(userEmail, OPERATION_TYPES.UPDATE, projectId, {
    module: 'organization-raci',
    count: raciEntries.length
  });
}

/**
 * RACIエントリを取得
 * @param {string} projectId - プロジェクトID
 * @return {Array<Object>} RACIエントリ一覧
 */
function getRACIEntries(projectId) {
  var results = findRows(
    SCHEMA_ORGANIZATION_RACI.sheetName,
    SCHEMA_ORGANIZATION_RACI.columns.PROJECT_ID,
    projectId
  );

  return results.map(function(result) {
    var data = result.data;
    return {
      pillar: data[SCHEMA_ORGANIZATION_RACI.columns.PILLAR],
      task: data[SCHEMA_ORGANIZATION_RACI.columns.TASK],
      assignee: data[SCHEMA_ORGANIZATION_RACI.columns.ASSIGNEE],
      raci: data[SCHEMA_ORGANIZATION_RACI.columns.RACI],
      updatedDate: data[SCHEMA_ORGANIZATION_RACI.columns.UPDATED_DATE]
    };
  });
}

// ========================================
// 価値トラッキング操作
// ========================================

/**
 * 価値トラッキングを保存
 * @param {Object} valueData - 価値データ
 */
function saveValue(valueData) {
  var existing = findFirstRow(
    SCHEMA_VALUE.sheetName,
    SCHEMA_VALUE.columns.USECASE_ID,
    valueData.usecaseId
  );

  var rowData = [];
  rowData[SCHEMA_VALUE.columns.PROJECT_ID] = valueData.projectId;
  rowData[SCHEMA_VALUE.columns.USECASE_ID] = valueData.usecaseId;
  rowData[SCHEMA_VALUE.columns.QUANTITATIVE_IMPACT] = valueData.quantitativeImpact || '';
  rowData[SCHEMA_VALUE.columns.QUALITATIVE_IMPACT] = valueData.qualitativeImpact || '';
  rowData[SCHEMA_VALUE.columns.EVIDENCE] = valueData.evidence || '';
  rowData[SCHEMA_VALUE.columns.NEXT_INVESTMENT] = valueData.nextInvestment || '';
  rowData[SCHEMA_VALUE.columns.UPDATED_DATE] = getCurrentTimestamp();

  if (existing) {
    updateRow(SCHEMA_VALUE.sheetName, existing.rowNumber, rowData);
    logAudit(valueData.userEmail, OPERATION_TYPES.UPDATE, valueData.projectId, {
      module: 'value',
      usecaseId: valueData.usecaseId
    });
  } else {
    appendRow(SCHEMA_VALUE.sheetName, rowData);
    logAudit(valueData.userEmail, OPERATION_TYPES.CREATE, valueData.projectId, {
      module: 'value',
      usecaseId: valueData.usecaseId
    });
  }
}

/**
 * 価値トラッキングを取得
 * @param {string} projectId - プロジェクトID
 * @return {Array<Object>} 価値データ一覧
 */
function getValues(projectId) {
  var results = findRows(
    SCHEMA_VALUE.sheetName,
    SCHEMA_VALUE.columns.PROJECT_ID,
    projectId
  );

  return results.map(function(result) {
    var data = result.data;
    return {
      usecaseId: data[SCHEMA_VALUE.columns.USECASE_ID],
      quantitativeImpact: data[SCHEMA_VALUE.columns.QUANTITATIVE_IMPACT],
      qualitativeImpact: data[SCHEMA_VALUE.columns.QUALITATIVE_IMPACT],
      evidence: data[SCHEMA_VALUE.columns.EVIDENCE],
      nextInvestment: data[SCHEMA_VALUE.columns.NEXT_INVESTMENT],
      updatedDate: data[SCHEMA_VALUE.columns.UPDATED_DATE]
    };
  });
}

// ========================================
// 監査ログ操作
// ========================================

/**
 * 監査ログを記録
 * @param {string} userEmail - 操作ユーザー
 * @param {string} operation - 操作種別
 * @param {string} projectId - プロジェクトID
 * @param {Object} details - 詳細情報
 */
function logAudit(userEmail, operation, projectId, details) {
  var rowData = [];
  rowData[SCHEMA_AUDIT_LOG.columns.TIMESTAMP] = getCurrentTimestamp();
  rowData[SCHEMA_AUDIT_LOG.columns.USER] = userEmail;
  rowData[SCHEMA_AUDIT_LOG.columns.OPERATION] = operation;
  rowData[SCHEMA_AUDIT_LOG.columns.PROJECT_ID] = projectId;
  rowData[SCHEMA_AUDIT_LOG.columns.DETAILS] = safeJsonStringify(details);

  appendRow(SCHEMA_AUDIT_LOG.sheetName, rowData);
}

/**
 * 監査ログを取得
 * @param {string} projectId - プロジェクトID
 * @return {Array<Object>} 監査ログ一覧
 */
function getAuditLogs(projectId) {
  var results = findRows(
    SCHEMA_AUDIT_LOG.sheetName,
    SCHEMA_AUDIT_LOG.columns.PROJECT_ID,
    projectId
  );

  return results.map(function(result) {
    var data = result.data;
    return {
      timestamp: data[SCHEMA_AUDIT_LOG.columns.TIMESTAMP],
      user: data[SCHEMA_AUDIT_LOG.columns.USER],
      operation: data[SCHEMA_AUDIT_LOG.columns.OPERATION],
      projectId: data[SCHEMA_AUDIT_LOG.columns.PROJECT_ID],
      details: safeJsonParse(data[SCHEMA_AUDIT_LOG.columns.DETAILS], {})
    };
  });
}
