/**
 * Blueprint ワークショップ - Google Apps Script Web Application
 *
 * 営業が顧客同席で「合意形成→次アクション→稟議書案（Doc生成）」まで進めるWebアプリ
 */

// ===========================================
// 設定
// ===========================================
const CONFIG = {
  SPREADSHEET_ID: '', // デプロイ時に設定（空の場合は自動作成）
  SPREADSHEET_NAME: 'Blueprint_Workshop_DB',
  SHEETS: {
    PROJECTS: 'プロジェクト',
    MODULE_DATA: 'モジュールデータ',
    VALUES: '価値',
    AUDIT_LOG: '監査ログ'
  }
};

// ===========================================
// Web App エントリーポイント
// ===========================================

/**
 * GET リクエストハンドラ - Web Appのエントリーポイント
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.user = getCurrentUser_();

  return template.evaluate()
    .setTitle('Blueprint ワークショップ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * HTMLファイルをインクルードする
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===========================================
// ユーザー・権限関連
// ===========================================

/**
 * 現在のユーザー情報を取得
 */
function getCurrentUser_() {
  const email = Session.getActiveUser().getEmail();
  return {
    email: email,
    name: email.split('@')[0]
  };
}

/**
 * プロジェクトへのアクセス権限をチェック
 */
function checkProjectAccess_(projectId, requireEdit = false) {
  const userEmail = Session.getActiveUser().getEmail();
  const project = getProjectById_(projectId);

  if (!project) {
    throw new Error('プロジェクトが見つかりません');
  }

  const isCreator = project.creatorEmail === userEmail;
  const editors = (project.editorsCSV || '').split(',').map(e => e.trim().toLowerCase());
  const isEditor = editors.includes(userEmail.toLowerCase());

  if (!isCreator && !isEditor) {
    throw new Error('このプロジェクトへのアクセス権限がありません');
  }

  return project;
}

// ===========================================
// データベース初期化
// ===========================================

/**
 * データベース（スプレッドシート）を初期化
 */
function initDb() {
  let ss;

  if (CONFIG.SPREADSHEET_ID) {
    try {
      ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } catch (e) {
      ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
      Logger.log('新しいスプレッドシートを作成しました: ' + ss.getId());
    }
  } else {
    ss = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
    Logger.log('新しいスプレッドシートを作成しました: ' + ss.getId());
  }

  // プロジェクトシート
  createSheetIfNotExists_(ss, CONFIG.SHEETS.PROJECTS, [
    'プロジェクトID', '顧客名', '作成日', '作成者メール', '編集者メールCSV', '状態', '最終更新'
  ]);

  // モジュールデータシート
  createSheetIfNotExists_(ss, CONFIG.SHEETS.MODULE_DATA, [
    'プロジェクトID', 'モジュールID', 'JSON', '最終更新', '更新者'
  ]);

  // 価値シート
  createSheetIfNotExists_(ss, CONFIG.SHEETS.VALUES, [
    'プロジェクトID', 'ユースケースID', '定量効果', '定性効果', '証跡リンク', '次の投資判断', '最終更新'
  ]);

  // 監査ログシート
  createSheetIfNotExists_(ss, CONFIG.SHEETS.AUDIT_LOG, [
    '日時', 'ユーザー', '操作', 'プロジェクトID', '詳細JSON'
  ]);

  return {
    success: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl()
  };
}

/**
 * シートが存在しない場合に作成
 */
function createSheetIfNotExists_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * データベーススプレッドシートを取得
 */
function getDatabase_() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }

  // 名前で検索
  const files = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }

  // 見つからない場合は初期化
  const result = initDb();
  return SpreadsheetApp.openById(result.spreadsheetId);
}

// ===========================================
// プロジェクト CRUD
// ===========================================

/**
 * プロジェクトを作成
 */
function createProject(customerName, editorsCsv) {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const userEmail = Session.getActiveUser().getEmail();

  const projectId = Utilities.getUuid();
  const now = new Date();

  sheet.appendRow([
    projectId,
    customerName,
    now,
    userEmail,
    editorsCsv || '',
    '下書き',
    now
  ]);

  writeAuditLog_('プロジェクト作成', projectId, { customerName, editorsCsv });

  return {
    success: true,
    projectId: projectId,
    customerName: customerName
  };
}

/**
 * プロジェクト一覧を取得
 */
function listProjects(filter, search) {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const projects = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const project = {
      projectId: row[0],
      customerName: row[1],
      createdAt: row[2],
      creatorEmail: row[3],
      editorsCSV: row[4],
      status: row[5],
      updatedAt: row[6]
    };

    // アクセス権限チェック
    const isCreator = project.creatorEmail.toLowerCase() === userEmail;
    const editors = (project.editorsCSV || '').split(',').map(e => e.trim().toLowerCase());
    const isEditor = editors.includes(userEmail);

    if (!isCreator && !isEditor) continue;

    // フィルター
    if (filter && filter !== 'all' && project.status !== filter) continue;

    // 検索
    if (search) {
      const searchLower = search.toLowerCase();
      if (!project.customerName.toLowerCase().includes(searchLower)) continue;
    }

    // 完了度を計算
    project.completion = getProjectCompletion_(project.projectId);

    projects.push(project);
  }

  // 最終更新日でソート（新しい順）
  projects.sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));

  return projects;
}

/**
 * プロジェクトを取得（権限チェック付き）
 */
function getProject(projectId) {
  const project = checkProjectAccess_(projectId);
  project.completion = getProjectCompletion_(projectId);
  project.modules = getProjectModules_(projectId);
  return project;
}

/**
 * プロジェクトをIDで取得（内部用）
 */
function getProjectById_(projectId) {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      return {
        projectId: data[i][0],
        customerName: data[i][1],
        createdAt: data[i][2],
        creatorEmail: data[i][3],
        editorsCSV: data[i][4],
        status: data[i][5],
        updatedAt: data[i][6],
        rowIndex: i + 1
      };
    }
  }

  return null;
}

/**
 * プロジェクトの完了度を計算
 */
function getProjectCompletion_(projectId) {
  const modules = getProjectModules_(projectId);
  const totalModules = 7;
  let completedModules = 0;

  Object.values(modules).forEach(m => {
    if (m && m.isCompleted) completedModules++;
  });

  return {
    completed: completedModules,
    total: totalModules,
    percentage: Math.round((completedModules / totalModules) * 100)
  };
}

/**
 * プロジェクトのモジュール一覧を取得
 */
function getProjectModules_(projectId) {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MODULE_DATA);
  const data = sheet.getDataRange().getValues();
  const modules = {};

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      const moduleId = data[i][1];
      try {
        modules[moduleId] = JSON.parse(data[i][2] || '{}');
        modules[moduleId].updatedAt = data[i][3];
        modules[moduleId].updatedBy = data[i][4];
      } catch (e) {
        modules[moduleId] = { error: 'JSON parse error' };
      }
    }
  }

  return modules;
}

/**
 * プロジェクトの状態を更新
 */
function updateProjectStatus(projectId, status) {
  checkProjectAccess_(projectId, true);

  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      sheet.getRange(i + 1, 6).setValue(status);
      sheet.getRange(i + 1, 7).setValue(new Date());
      break;
    }
  }

  writeAuditLog_('プロジェクト状態更新', projectId, { status });

  return { success: true };
}

// ===========================================
// モジュールデータ CRUD
// ===========================================

/**
 * モジュールデータを読み込む
 */
function loadModule(projectId, moduleId) {
  checkProjectAccess_(projectId);

  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MODULE_DATA);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId && data[i][1] === moduleId) {
      try {
        return {
          success: true,
          data: JSON.parse(data[i][2] || '{}'),
          updatedAt: data[i][3],
          updatedBy: data[i][4]
        };
      } catch (e) {
        return { success: false, error: 'JSON parse error' };
      }
    }
  }

  return { success: true, data: {}, isNew: true };
}

/**
 * モジュールデータを保存
 */
function saveModule(projectId, moduleId, payloadJson) {
  checkProjectAccess_(projectId, true);

  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MODULE_DATA);
  const data = sheet.getDataRange().getValues();
  const userEmail = Session.getActiveUser().getEmail();
  const now = new Date();

  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId && data[i][1] === moduleId) {
      sheet.getRange(i + 1, 3).setValue(payloadJson);
      sheet.getRange(i + 1, 4).setValue(now);
      sheet.getRange(i + 1, 5).setValue(userEmail);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow([projectId, moduleId, payloadJson, now, userEmail]);
  }

  // プロジェクトの最終更新日を更新
  updateProjectLastModified_(projectId);

  writeAuditLog_('モジュール保存', projectId, { moduleId });

  return {
    success: true,
    savedAt: now.toISOString()
  };
}

/**
 * モジュールを完了としてマーク
 */
function markComplete(projectId, moduleId) {
  checkProjectAccess_(projectId, true);

  // 現在のデータを取得
  const result = loadModule(projectId, moduleId);
  const moduleData = result.data || {};

  // isCompletedフラグを設定
  moduleData.isCompleted = true;
  moduleData.completedAt = new Date().toISOString();
  moduleData.completedBy = Session.getActiveUser().getEmail();

  // 保存
  saveModule(projectId, moduleId, JSON.stringify(moduleData));

  writeAuditLog_('モジュール完了', projectId, { moduleId });

  return { success: true };
}

/**
 * モジュールの完了を解除
 */
function unmarkComplete(projectId, moduleId) {
  checkProjectAccess_(projectId, true);

  const result = loadModule(projectId, moduleId);
  const moduleData = result.data || {};

  moduleData.isCompleted = false;
  delete moduleData.completedAt;
  delete moduleData.completedBy;

  saveModule(projectId, moduleId, JSON.stringify(moduleData));

  writeAuditLog_('モジュール完了解除', projectId, { moduleId });

  return { success: true };
}

/**
 * プロジェクトの最終更新日を更新
 */
function updateProjectLastModified_(projectId) {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      sheet.getRange(i + 1, 7).setValue(new Date());
      break;
    }
  }
}

// ===========================================
// 価値トラッカー CRUD
// ===========================================

/**
 * 価値データを保存
 */
function saveValue(projectId, useCaseId, quantitative, qualitative, evidenceLink, nextAction) {
  checkProjectAccess_(projectId, true);

  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.VALUES);
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId && data[i][1] === useCaseId) {
      sheet.getRange(i + 1, 3, 1, 4).setValues([[quantitative, qualitative, evidenceLink, nextAction]]);
      sheet.getRange(i + 1, 7).setValue(now);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow([projectId, useCaseId, quantitative, qualitative, evidenceLink, nextAction, now]);
  }

  writeAuditLog_('価値データ保存', projectId, { useCaseId });

  return { success: true };
}

/**
 * 価値データ一覧を取得
 */
function listValues(projectId) {
  checkProjectAccess_(projectId);

  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.VALUES);
  const data = sheet.getDataRange().getValues();
  const values = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      values.push({
        useCaseId: data[i][1],
        quantitative: data[i][2],
        qualitative: data[i][3],
        evidenceLink: data[i][4],
        nextAction: data[i][5],
        updatedAt: data[i][6]
      });
    }
  }

  return values;
}

// ===========================================
// ドキュメントエクスポート
// ===========================================

/**
 * プロジェクトをGoogleドキュメントにエクスポート
 */
function exportProjectDoc(projectId) {
  const project = checkProjectAccess_(projectId);
  const modules = getProjectModules_(projectId);
  const values = listValues(projectId);

  // ドキュメントを作成
  const doc = DocumentApp.create('Blueprint_' + project.customerName + '_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd'));
  const body = doc.getBody();

  // ヘッダー
  body.appendParagraph('Blueprint ワークショップ 稟議書').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('');

  // プロジェクト情報
  body.appendParagraph('プロジェクト概要').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('顧客名: ' + project.customerName);
  body.appendParagraph('作成日: ' + Utilities.formatDate(new Date(project.createdAt), 'Asia/Tokyo', 'yyyy/MM/dd'));
  body.appendParagraph('作成者: ' + project.creatorEmail);
  body.appendParagraph('状態: ' + project.status);
  body.appendParagraph('');

  // モジュール1: ビジョン
  if (modules.module1) {
    body.appendParagraph('1. 北極星ビジョン').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const m1 = modules.module1;
    if (m1.visionStatement) {
      body.appendParagraph('ビジョン本文').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(m1.visionStatement);
    }
    if (m1.decisionRules) {
      body.appendParagraph('意思決定ルール').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(m1.decisionRules);
    }
    if (m1.successMetrics) {
      body.appendParagraph('成功指標').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(m1.successMetrics);
    }
    body.appendParagraph('');
  }

  // モジュール2: ユースケース選定
  if (modules.module2) {
    body.appendParagraph('2. ユースケース選定').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const m2 = modules.module2;
    if (m2.useCases && m2.useCases.length > 0) {
      m2.useCases.forEach((uc, idx) => {
        body.appendParagraph((idx + 1) + '. ' + (uc.name || 'ユースケース'));
        body.appendParagraph('  インパクト: ' + (uc.impact || '-') + ' / 実現性: ' + (uc.feasibility || '-') + ' / 能力構築: ' + (uc.capability || '-'));
      });
    }
    body.appendParagraph('');
  }

  // モジュール3: 90日計画
  if (modules.module3) {
    body.appendParagraph('3. 90日計画').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const m3 = modules.module3;
    if (m3.coreObjective) {
      body.appendParagraph('目標: ' + m3.coreObjective);
    }
    if (m3.kpiTarget1) {
      body.appendParagraph('KPI 1: ' + m3.kpiTarget1);
    }
    if (m3.kpiTarget2) {
      body.appendParagraph('KPI 2: ' + m3.kpiTarget2);
    }
    body.appendParagraph('');
  }

  // モジュール4: 体制
  if (modules.module4) {
    body.appendParagraph('4. 体制と責任').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const m4 = modules.module4;
    if (m4.roles && m4.roles.length > 0) {
      m4.roles.forEach(role => {
        body.appendParagraph(role.title + ': ' + (role.assignee || '未定'));
      });
    }
    body.appendParagraph('');
  }

  // 価値トラッカー
  if (values.length > 0) {
    body.appendParagraph('価値実現トラッカー').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    values.forEach(v => {
      body.appendParagraph('ユースケース: ' + v.useCaseId).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph('定量効果: ' + (v.quantitative || '-'));
      body.appendParagraph('定性効果: ' + (v.qualitative || '-'));
      body.appendParagraph('証跡: ' + (v.evidenceLink || '-'));
      body.appendParagraph('次の投資判断: ' + (v.nextAction || '-'));
      body.appendParagraph('');
    });
  }

  // フッター
  body.appendParagraph('');
  body.appendParagraph('---');
  body.appendParagraph('生成日時: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
  body.appendParagraph('Blueprint ワークショップにて自動生成');

  doc.saveAndClose();

  writeAuditLog_('ドキュメントエクスポート', projectId, { docId: doc.getId() });

  return {
    success: true,
    docId: doc.getId(),
    docUrl: doc.getUrl()
  };
}

// ===========================================
// 監査ログ
// ===========================================

/**
 * 監査ログを書き込む
 */
function writeAuditLog_(operation, projectId, details) {
  try {
    const ss = getDatabase_();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT_LOG);
    const userEmail = Session.getActiveUser().getEmail();

    sheet.appendRow([
      new Date(),
      userEmail,
      operation,
      projectId || '',
      JSON.stringify(details || {})
    ]);
  } catch (e) {
    // 監査ログの書き込み失敗は握りつぶさない
    Logger.log('監査ログ書き込みエラー: ' + e.message);
  }
}

/**
 * 監査ログを取得（管理者用）
 */
function getAuditLogs(projectId, limit) {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT_LOG);
  const data = sheet.getDataRange().getValues();
  const logs = [];

  for (let i = data.length - 1; i >= 1; i--) {
    if (projectId && data[i][3] !== projectId) continue;

    logs.push({
      timestamp: data[i][0],
      user: data[i][1],
      operation: data[i][2],
      projectId: data[i][3],
      details: data[i][4]
    });

    if (limit && logs.length >= limit) break;
  }

  return logs;
}

// ===========================================
// ユーティリティ
// ===========================================

/**
 * HTML特殊文字をエスケープ（XSS対策）
 */
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * 現在のユーザー情報を取得（クライアント用）
 */
function getCurrentUserInfo() {
  return getCurrentUser_();
}
