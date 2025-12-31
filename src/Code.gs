/**
 * Code.gs
 * Web App のエントリーポイントとコントローラー層
 */

/**
 * Web App の GET リクエストハンドラー
 * @param {Object} e - イベントオブジェクト
 * @return {HtmlOutput} HTML出力
 */
function doGet(e) {
  // ユーザー認証チェック
  var userEmail = getCurrentUserEmail();
  if (!userEmail) {
    return HtmlService.createHtmlOutput('認証が必要です。Googleアカウントでログインしてください。');
  }

  // メインHTMLを返却
  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = userEmail;

  return template.evaluate()
    .setTitle('Tableau Blueprint ワークショップ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTMLファイル内でインクルードするためのヘルパー関数
 * @param {string} filename - ファイル名
 * @return {string} ファイルの内容
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ========================================
// API: プロジェクト管理
// ========================================

/**
 * 新規プロジェクトを作成
 * @param {string} customerName - 顧客名
 * @return {Object} 作成されたプロジェクト情報
 */
function apiCreateProject(customerName) {
  try {
    var userEmail = getCurrentUserEmail();
    var project = createProject(customerName, userEmail);
    return { success: true, data: project };
  } catch (error) {
    Logger.log('apiCreateProject error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * ユーザーのプロジェクト一覧を取得
 * @return {Object} プロジェクト一覧
 */
function apiGetUserProjects() {
  try {
    var userEmail = getCurrentUserEmail();
    var projects = getUserProjects(userEmail);
    return { success: true, data: projects };
  } catch (error) {
    Logger.log('apiGetUserProjects error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * プロジェクト詳細を取得
 * @param {string} projectId - プロジェクトID
 * @return {Object} プロジェクト詳細
 */
function apiGetProject(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var project = getProject(projectId);
    return { success: true, data: project };
  } catch (error) {
    Logger.log('apiGetProject error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * プロジェクトの状態を更新
 * @param {string} projectId - プロジェクトID
 * @param {string} status - 新しい状態
 * @return {Object} 更新結果
 */
function apiUpdateProjectStatus(projectId, status) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    updateProjectStatus(projectId, status, userEmail);
    return { success: true, message: 'プロジェクトの状態を更新しました' };
  } catch (error) {
    Logger.log('apiUpdateProjectStatus error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========================================
// API: Module 1 - ビジョン
// ========================================

/**
 * ビジョンを保存
 * @param {Object} visionData - ビジョンデータ
 * @return {Object} 保存結果
 */
function apiSaveVision(visionData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(visionData.projectId, userEmail);

    visionData.userEmail = userEmail;
    saveVision(visionData);

    return { success: true, message: 'ビジョンを保存しました' };
  } catch (error) {
    Logger.log('apiSaveVision error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * ビジョンを取得
 * @param {string} projectId - プロジェクトID
 * @return {Object} ビジョンデータ
 */
function apiGetVision(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var vision = getVision(projectId);
    return { success: true, data: vision };
  } catch (error) {
    Logger.log('apiGetVision error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========================================
// API: Module 2 - ユースケース & 90日計画
// ========================================

/**
 * ユースケースを追加
 * @param {Object} usecaseData - ユースケースデータ
 * @return {Object} 追加結果
 */
function apiAddUsecase(usecaseData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(usecaseData.projectId, userEmail);

    usecaseData.userEmail = userEmail;
    var usecaseId = addUsecase(usecaseData);

    return { success: true, data: { usecaseId: usecaseId } };
  } catch (error) {
    Logger.log('apiAddUsecase error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * ユースケース一覧を取得
 * @param {string} projectId - プロジェクトID
 * @return {Object} ユースケース一覧
 */
function apiGetUsecases(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var usecases = getUsecases(projectId);
    return { success: true, data: usecases };
  } catch (error) {
    Logger.log('apiGetUsecases error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 90日計画を保存
 * @param {Object} planData - 90日計画データ
 * @return {Object} 保存結果
 */
function apiSaveNinetyDayPlan(planData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(planData.projectId, userEmail);

    planData.userEmail = userEmail;
    saveNinetyDayPlan(planData);

    return { success: true, message: '90日計画を保存しました' };
  } catch (error) {
    Logger.log('apiSaveNinetyDayPlan error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 90日計画を取得
 * @param {string} usecaseId - ユースケースID
 * @param {string} projectId - プロジェクトID（権限チェック用）
 * @return {Object} 90日計画データ
 */
function apiGetNinetyDayPlan(usecaseId, projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var plan = getNinetyDayPlan(usecaseId);
    return { success: true, data: plan };
  } catch (error) {
    Logger.log('apiGetNinetyDayPlan error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========================================
// API: Module 4 - 体制・RACI
// ========================================

/**
 * RACIエントリを保存
 * @param {string} projectId - プロジェクトID
 * @param {Array<Object>} raciEntries - RACIエントリ配列
 * @return {Object} 保存結果
 */
function apiSaveRACIEntries(projectId, raciEntries) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    saveRACIEntries(projectId, raciEntries, userEmail);

    return { success: true, message: 'RACI設計を保存しました' };
  } catch (error) {
    Logger.log('apiSaveRACIEntries error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * RACIエントリを取得
 * @param {string} projectId - プロジェクトID
 * @return {Object} RACIエントリ一覧
 */
function apiGetRACIEntries(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var entries = getRACIEntries(projectId);
    return { success: true, data: entries };
  } catch (error) {
    Logger.log('apiGetRACIEntries error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========================================
// API: Module 7 - 価値トラッキング
// ========================================

/**
 * 価値トラッキングを保存
 * @param {Object} valueData - 価値データ
 * @return {Object} 保存結果
 */
function apiSaveValue(valueData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(valueData.projectId, userEmail);

    valueData.userEmail = userEmail;
    saveValue(valueData);

    return { success: true, message: '価値トラッキングを保存しました' };
  } catch (error) {
    Logger.log('apiSaveValue error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 価値トラッキング一覧を取得
 * @param {string} projectId - プロジェクトID
 * @return {Object} 価値データ一覧
 */
function apiGetValues(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var values = getValues(projectId);
    return { success: true, data: values };
  } catch (error) {
    Logger.log('apiGetValues error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========================================
// API: ドキュメント生成
// ========================================

/**
 * 稟議書案を生成
 * @param {string} projectId - プロジェクトID
 * @return {Object} 生成結果（ドキュメントURL含む）
 */
function apiGenerateProposal(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var docUrl = generateProposalDocument(projectId, userEmail);

    return { success: true, data: { documentUrl: docUrl } };
  } catch (error) {
    Logger.log('apiGenerateProposal error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========================================
// API: ユーティリティ
// ========================================

/**
 * シートを初期化（管理者用）
 * @return {Object} 初期化結果
 */
function apiInitializeSheets() {
  try {
    var result = initializeAllSheets();
    return { success: true, message: result.message };
  } catch (error) {
    Logger.log('apiInitializeSheets error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * 現在のユーザー情報を取得
 * @return {Object} ユーザー情報
 */
function apiGetCurrentUser() {
  try {
    var user = getCurrentUser();
    return { success: true, data: user };
  } catch (error) {
    Logger.log('apiGetCurrentUser error: ' + error.message);
    return { success: false, error: error.message };
  }
}
