/**
 * Auth.gs
 * 認証・権限制御を管理
 */

/**
 * ユーザーがプロジェクトにアクセス権限があるかチェック
 * @param {string} projectId - プロジェクトID
 * @param {string} userEmail - ユーザーメール（省略時は現在のユーザー）
 * @return {boolean} アクセス権限の有無
 */
function hasProjectAccess(projectId, userEmail) {
  userEmail = userEmail || getCurrentUserEmail();
  var project = getProject(projectId);

  if (!project) {
    return false;
  }

  // 作成者の場合はアクセス可能
  if (project.createdBy === userEmail) {
    return true;
  }

  // 編集者リストに含まれているかチェック
  var editors = project.editors || '';
  var editorList = editors.split(',').map(function(e) { return e.trim(); });

  return editorList.indexOf(userEmail) !== -1;
}

/**
 * プロジェクトへのアクセス権限を検証（権限がない場合はエラーをスロー）
 * @param {string} projectId - プロジェクトID
 * @param {string} userEmail - ユーザーメール（省略時は現在のユーザー）
 * @throws {Error} アクセス権限がない場合
 */
function requireProjectAccess(projectId, userEmail) {
  userEmail = userEmail || getCurrentUserEmail();

  if (!hasProjectAccess(projectId, userEmail)) {
    throw new Error('このプロジェクトへのアクセス権限がありません: ' + projectId);
  }
}

/**
 * ユーザーがプロジェクトを編集可能かチェック
 * （現在の実装では、アクセス可能 = 編集可能）
 * @param {string} projectId - プロジェクトID
 * @param {string} userEmail - ユーザーメール（省略時は現在のユーザー）
 * @return {boolean} 編集権限の有無
 */
function canEditProject(projectId, userEmail) {
  return hasProjectAccess(projectId, userEmail);
}

/**
 * プロジェクトに編集者を追加
 * @param {string} projectId - プロジェクトID
 * @param {string} newEditorEmail - 追加する編集者のメール
 * @param {string} currentUserEmail - 現在のユーザーメール
 */
function addProjectEditor(projectId, newEditorEmail, currentUserEmail) {
  requireProjectAccess(projectId, currentUserEmail);

  var project = getProject(projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません: ' + projectId);
  }

  var editors = project.editors || '';
  var editorList = editors.split(',').map(function(e) { return e.trim(); });

  // すでに含まれている場合はスキップ
  if (editorList.indexOf(newEditorEmail) !== -1) {
    return;
  }

  editorList.push(newEditorEmail);

  var rowData = [];
  rowData[SCHEMA_PROJECTS.columns.PROJECT_ID] = project.projectId;
  rowData[SCHEMA_PROJECTS.columns.CUSTOMER_NAME] = project.customerName;
  rowData[SCHEMA_PROJECTS.columns.CREATED_DATE] = project.createdDate;
  rowData[SCHEMA_PROJECTS.columns.CREATED_BY] = project.createdBy;
  rowData[SCHEMA_PROJECTS.columns.EDITORS] = editorList.join(', ');
  rowData[SCHEMA_PROJECTS.columns.STATUS] = project.status;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_DATE] = getCurrentTimestamp();
  rowData[SCHEMA_PROJECTS.columns.UPDATED_BY] = currentUserEmail;

  updateRow(SCHEMA_PROJECTS.sheetName, project._rowNumber, rowData);

  logAudit(currentUserEmail, OPERATION_TYPES.UPDATE, projectId, {
    action: 'add_editor',
    newEditor: newEditorEmail
  });
}

/**
 * プロジェクトから編集者を削除
 * @param {string} projectId - プロジェクトID
 * @param {string} editorEmailToRemove - 削除する編集者のメール
 * @param {string} currentUserEmail - 現在のユーザーメール
 */
function removeProjectEditor(projectId, editorEmailToRemove, currentUserEmail) {
  requireProjectAccess(projectId, currentUserEmail);

  var project = getProject(projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません: ' + projectId);
  }

  // 作成者は削除できない
  if (project.createdBy === editorEmailToRemove) {
    throw new Error('作成者は編集者リストから削除できません');
  }

  var editors = project.editors || '';
  var editorList = editors.split(',').map(function(e) { return e.trim(); });

  editorList = editorList.filter(function(e) {
    return e !== editorEmailToRemove;
  });

  var rowData = [];
  rowData[SCHEMA_PROJECTS.columns.PROJECT_ID] = project.projectId;
  rowData[SCHEMA_PROJECTS.columns.CUSTOMER_NAME] = project.customerName;
  rowData[SCHEMA_PROJECTS.columns.CREATED_DATE] = project.createdDate;
  rowData[SCHEMA_PROJECTS.columns.CREATED_BY] = project.createdBy;
  rowData[SCHEMA_PROJECTS.columns.EDITORS] = editorList.join(', ');
  rowData[SCHEMA_PROJECTS.columns.STATUS] = project.status;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_DATE] = getCurrentTimestamp();
  rowData[SCHEMA_PROJECTS.columns.UPDATED_BY] = currentUserEmail;

  updateRow(SCHEMA_PROJECTS.sheetName, project._rowNumber, rowData);

  logAudit(currentUserEmail, OPERATION_TYPES.UPDATE, projectId, {
    action: 'remove_editor',
    removedEditor: editorEmailToRemove
  });
}

/**
 * 現在のユーザー情報を取得
 * @return {Object} ユーザー情報
 */
function getCurrentUser() {
  var userEmail = getCurrentUserEmail();
  return {
    email: userEmail,
    isAuthenticated: userEmail !== ''
  };
}
