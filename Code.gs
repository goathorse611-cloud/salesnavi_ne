// ================================================================================
// Blueprint Workshop - Google Apps Script Web Application
// ================================================================================

// Configuration
const CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID'),
  APP_NAME: 'Blueprint ワークショップ',
  SHEETS: {
    PROJECTS: 'プロジェクト',
    MODULES: 'モジュールデータ',
    VALUES: '価値',
    AUDIT: '監査ログ'
  }
};

// ================================================================================
// Main Entry Point
// ================================================================================

/**
 * Serves the main web application
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = Session.getActiveUser().getEmail();

  return template.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Include HTML partials (for <?!= include('filename') ?> syntax)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ================================================================================
// Database Initialization
// ================================================================================

/**
 * Initialize database spreadsheet with required sheets
 */
function initDb() {
  try {
    let ss;
    const existingId = CONFIG.SPREADSHEET_ID;

    if (existingId) {
      try {
        ss = SpreadsheetApp.openById(existingId);
      } catch (e) {
        // If can't open, create new
        ss = SpreadsheetApp.create(CONFIG.APP_NAME + ' - Database');
        PropertiesService.getScriptProperties().setProperty('DB_SPREADSHEET_ID', ss.getId());
      }
    } else {
      ss = SpreadsheetApp.create(CONFIG.APP_NAME + ' - Database');
      PropertiesService.getScriptProperties().setProperty('DB_SPREADSHEET_ID', ss.getId());
    }

    // Create Projects sheet
    createOrUpdateSheet(ss, CONFIG.SHEETS.PROJECTS, [
      'プロジェクトID', '顧客名', '作成日', '作成者メール', '編集者メールCSV',
      '状態', '最終更新', '完了モジュール数'
    ]);

    // Create Modules sheet
    createOrUpdateSheet(ss, CONFIG.SHEETS.MODULES, [
      'プロジェクトID', 'モジュールID', 'JSON', '完了フラグ', '最終更新', '更新者'
    ]);

    // Create Values sheet
    createOrUpdateSheet(ss, CONFIG.SHEETS.VALUES, [
      'プロジェクトID', 'ユースケースID', '定量効果', '定性効果',
      '証跡リンク', '次の投資判断', '最終更新'
    ]);

    // Create Audit sheet
    createOrUpdateSheet(ss, CONFIG.SHEETS.AUDIT, [
      '日時', 'ユーザー', '操作', 'プロジェクトID', '詳細JSON'
    ]);

    return {
      success: true,
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl()
    };
  } catch (error) {
    Logger.log('Error in initDb: ' + error.toString());
    throw new Error('DB初期化エラー: ' + error.message);
  }
}

/**
 * Helper: Create or update sheet with headers
 */
function createOrUpdateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0176d3').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

// ================================================================================
// Project CRUD Operations
// ================================================================================

/**
 * Create new project
 */
function createProject(customerName, editorsCsv) {
  try {
    checkDbInitialized();
    const user = Session.getActiveUser().getEmail();
    const projectId = 'P' + Utilities.getUuid().substring(0, 8);
    const now = new Date();

    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);

    sheet.appendRow([
      projectId,
      customerName,
      now,
      user,
      editorsCsv || user,
      '下書き',
      now,
      0
    ]);

    logAudit('CREATE_PROJECT', projectId, { customerName: customerName });

    return {
      success: true,
      projectId: projectId,
      message: 'プロジェクトを作成しました'
    };
  } catch (error) {
    Logger.log('Error in createProject: ' + error.toString());
    throw new Error('プロジェクト作成エラー: ' + error.message);
  }
}

/**
 * List projects with optional filters
 */
function listProjects(filter, search) {
  try {
    checkDbInitialized();
    const user = Session.getActiveUser().getEmail();

    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return { success: true, projects: [] };
    }

    const headers = data[0];
    const projects = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const projectId = row[0];
      const creatorEmail = row[3];
      const editorsEmail = row[4];

      // Permission check
      if (!hasProjectPermission(user, creatorEmail, editorsEmail)) {
        continue;
      }

      const project = {
        projectId: row[0],
        customerName: row[1],
        createdDate: row[2],
        creatorEmail: row[3],
        editorsEmail: row[4],
        status: row[5],
        lastUpdated: row[6],
        completedModules: row[7] || 0
      };

      // Apply filters
      if (filter && filter !== 'すべて' && project.status !== filter) {
        continue;
      }

      if (search && !project.customerName.includes(search)) {
        continue;
      }

      projects.push(project);
    }

    return {
      success: true,
      projects: projects
    };
  } catch (error) {
    Logger.log('Error in listProjects: ' + error.toString());
    throw new Error('プロジェクト一覧取得エラー: ' + error.message);
  }
}

/**
 * Get single project by ID
 */
function getProject(projectId) {
  try {
    checkDbInitialized();
    const user = Session.getActiveUser().getEmail();

    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === projectId) {
        const creatorEmail = data[i][3];
        const editorsEmail = data[i][4];

        // Permission check
        if (!hasProjectPermission(user, creatorEmail, editorsEmail)) {
          throw new Error('このプロジェクトへのアクセス権限がありません');
        }

        return {
          success: true,
          project: {
            projectId: data[i][0],
            customerName: data[i][1],
            createdDate: data[i][2],
            creatorEmail: data[i][3],
            editorsEmail: data[i][4],
            status: data[i][5],
            lastUpdated: data[i][6],
            completedModules: data[i][7] || 0
          }
        };
      }
    }

    throw new Error('プロジェクトが見つかりません');
  } catch (error) {
    Logger.log('Error in getProject: ' + error.toString());
    throw new Error('プロジェクト取得エラー: ' + error.message);
  }
}

// ================================================================================
// Module CRUD Operations
// ================================================================================

/**
 * Load module data for a project
 */
function loadModule(projectId, moduleId) {
  try {
    checkDbInitialized();
    checkProjectPermission(projectId);

    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MODULES);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === projectId && data[i][1] === moduleId) {
        return {
          success: true,
          data: JSON.parse(data[i][2] || '{}'),
          completed: data[i][3] || false,
          lastUpdated: data[i][4]
        };
      }
    }

    // Return empty if not found
    return {
      success: true,
      data: {},
      completed: false,
      lastUpdated: null
    };
  } catch (error) {
    Logger.log('Error in loadModule: ' + error.toString());
    throw new Error('モジュール読込エラー: ' + error.message);
  }
}

/**
 * Save module data
 */
function saveModule(projectId, moduleId, payloadJson) {
  try {
    checkDbInitialized();
    checkProjectPermission(projectId);

    const user = Session.getActiveUser().getEmail();
    const now = new Date();

    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MODULES);
    const data = sheet.getDataRange().getValues();

    // Find existing row
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === projectId && data[i][1] === moduleId) {
        rowIndex = i + 1;
        break;
      }
    }

    const jsonString = JSON.stringify(payloadJson);
    const completed = payloadJson._completed || false;

    if (rowIndex > 0) {
      // Update existing
      sheet.getRange(rowIndex, 3).setValue(jsonString);
      sheet.getRange(rowIndex, 4).setValue(completed);
      sheet.getRange(rowIndex, 5).setValue(now);
      sheet.getRange(rowIndex, 6).setValue(user);
    } else {
      // Insert new
      sheet.appendRow([
        projectId,
        moduleId,
        jsonString,
        completed,
        now,
        user
      ]);
    }

    // Update project last updated
    updateProjectTimestamp(projectId);

    logAudit('SAVE_MODULE', projectId, { moduleId: moduleId });

    return {
      success: true,
      message: '保存しました',
      timestamp: now
    };
  } catch (error) {
    Logger.log('Error in saveModule: ' + error.toString());
    throw new Error('モジュール保存エラー: ' + error.message);
  }
}

/**
 * Mark module as complete
 */
function markComplete(projectId, moduleId) {
  try {
    checkDbInitialized();
    checkProjectPermission(projectId);

    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.MODULES);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === projectId && data[i][1] === moduleId) {
        sheet.getRange(i + 1, 4).setValue(true);
        sheet.getRange(i + 1, 5).setValue(new Date());

        logAudit('MARK_COMPLETE', projectId, { moduleId: moduleId });

        return {
          success: true,
          message: 'モジュールを完了にしました'
        };
      }
    }

    throw new Error('モジュールが見つかりません');
  } catch (error) {
    Logger.log('Error in markComplete: ' + error.toString());
    throw new Error('完了マークエラー: ' + error.message);
  }
}

// ================================================================================
// Document Export
// ================================================================================

/**
 * Export project to Google Doc
 */
function exportProjectDoc(projectId) {
  try {
    checkDbInitialized();
    checkProjectPermission(projectId);

    const project = getProject(projectId).project;
    const modules = getAllModules(projectId);

    // Create new Google Doc
    const doc = DocumentApp.create(project.customerName + ' - Blueprint 稟議書 - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    const body = doc.getBody();

    // Title
    body.appendParagraph(project.customerName + ' プロジェクト').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Blueprint ワークショップ成果物').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendHorizontalRule();

    // Project info
    body.appendParagraph('プロジェクト情報').setHeading(DocumentApp.ParagraphHeading.HEADING3);
    body.appendParagraph('顧客名: ' + project.customerName);
    body.appendParagraph('作成日: ' + Utilities.formatDate(project.createdDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
    body.appendParagraph('ステータス: ' + project.status);
    body.appendParagraph('');

    // Modules
    body.appendParagraph('モジュール詳細').setHeading(DocumentApp.ParagraphHeading.HEADING3);

    modules.forEach(function(module) {
      body.appendParagraph('Module: ' + module.moduleId).setHeading(DocumentApp.ParagraphHeading.HEADING4);
      body.appendParagraph(JSON.stringify(module.data, null, 2));
      body.appendParagraph('');
    });

    doc.saveAndClose();

    logAudit('EXPORT_DOC', projectId, { docId: doc.getId() });

    return {
      success: true,
      docUrl: doc.getUrl(),
      docId: doc.getId(),
      message: 'ドキュメントを生成しました'
    };
  } catch (error) {
    Logger.log('Error in exportProjectDoc: ' + error.toString());
    throw new Error('ドキュメント生成エラー: ' + error.message);
  }
}

// ================================================================================
// Permission & Audit Helpers
// ================================================================================

/**
 * Check if user has permission to access project
 */
function hasProjectPermission(userEmail, creatorEmail, editorsEmail) {
  if (userEmail === creatorEmail) return true;

  const editors = editorsEmail.split(',').map(e => e.trim());
  return editors.includes(userEmail);
}

/**
 * Check project permission and throw if denied
 */
function checkProjectPermission(projectId) {
  const user = Session.getActiveUser().getEmail();
  const ss = getDbSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      const creatorEmail = data[i][3];
      const editorsEmail = data[i][4];

      if (!hasProjectPermission(user, creatorEmail, editorsEmail)) {
        throw new Error('このプロジェクトへのアクセス権限がありません');
      }
      return true;
    }
  }

  throw new Error('プロジェクトが見つかりません');
}

/**
 * Log audit trail
 */
function logAudit(operation, projectId, detailsObj) {
  try {
    const ss = getDbSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT);
    const user = Session.getActiveUser().getEmail();

    sheet.appendRow([
      new Date(),
      user,
      operation,
      projectId || '',
      JSON.stringify(detailsObj || {})
    ]);
  } catch (error) {
    Logger.log('Error in logAudit: ' + error.toString());
  }
}

// ================================================================================
// Utility Functions
// ================================================================================

/**
 * Get DB spreadsheet
 */
function getDbSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (!id) {
    throw new Error('データベースが初期化されていません。initDb()を実行してください。');
  }
  return SpreadsheetApp.openById(id);
}

/**
 * Check if DB is initialized
 */
function checkDbInitialized() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_SPREADSHEET_ID');
  if (!id) {
    throw new Error('データベースが初期化されていません。管理者に連絡してください。');
  }
}

/**
 * Update project timestamp
 */
function updateProjectTimestamp(projectId) {
  const ss = getDbSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PROJECTS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      sheet.getRange(i + 1, 7).setValue(new Date());
      break;
    }
  }
}

/**
 * Get all modules for a project
 */
function getAllModules(projectId) {
  const ss = getDbSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.MODULES);
  const data = sheet.getDataRange().getValues();

  const modules = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      modules.push({
        moduleId: data[i][1],
        data: JSON.parse(data[i][2] || '{}'),
        completed: data[i][3],
        lastUpdated: data[i][4]
      });
    }
  }

  return modules;
}

/**
 * Get current user email
 */
function getCurrentUser() {
  return {
    email: Session.getActiveUser().getEmail()
  };
}
