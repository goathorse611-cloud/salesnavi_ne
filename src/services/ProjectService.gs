/**
 * ProjectService.gs
 * プロジェクト管理のビジネスロジック層
 */

// ========================================
// プロジェクト作成・更新
// ========================================

/**
 * 新規プロジェクトを作成（バリデーション付き）
 * @param {string} customerName - 顧客名
 * @param {string} userEmail - 作成者メール
 * @return {Object} 作成されたプロジェクト情報
 */
function createProjectWithValidation(customerName, userEmail) {
  // バリデーション
  if (!customerName || customerName.trim().length === 0) {
    throw new Error('顧客名は必須です');
  }

  if (customerName.length > 100) {
    throw new Error('顧客名は100文字以内で入力してください');
  }

  if (!userEmail) {
    throw new Error('ユーザー認証が必要です');
  }

  // プロジェクト作成
  return createProject(customerName.trim(), userEmail);
}

/**
 * プロジェクトの顧客名を更新
 * @param {string} projectId - プロジェクトID
 * @param {string} customerName - 新しい顧客名
 * @param {string} userEmail - 更新者メール
 */
function updateProjectCustomerName(projectId, customerName, userEmail) {
  requireProjectAccess(projectId, userEmail);

  if (!customerName || customerName.trim().length === 0) {
    throw new Error('顧客名は必須です');
  }

  var project = getProject(projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません');
  }

  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('アーカイブされたプロジェクトは編集できません');
  }

  var rowData = [];
  rowData[SCHEMA_PROJECTS.columns.PROJECT_ID] = project.projectId;
  rowData[SCHEMA_PROJECTS.columns.CUSTOMER_NAME] = customerName.trim();
  rowData[SCHEMA_PROJECTS.columns.CREATED_DATE] = project.createdDate;
  rowData[SCHEMA_PROJECTS.columns.CREATED_BY] = project.createdBy;
  rowData[SCHEMA_PROJECTS.columns.EDITORS] = project.editors;
  rowData[SCHEMA_PROJECTS.columns.STATUS] = project.status;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_DATE] = getCurrentTimestamp();
  rowData[SCHEMA_PROJECTS.columns.UPDATED_BY] = userEmail;

  updateRow(SCHEMA_PROJECTS.sheetName, project._rowNumber, rowData);

  logAudit(userEmail, OPERATION_TYPES.UPDATE, projectId, {
    field: 'customerName',
    oldValue: project.customerName,
    newValue: customerName.trim()
  });
}

/**
 * プロジェクト状態を更新（遷移ルール付き）
 * @param {string} projectId - プロジェクトID
 * @param {string} newStatus - 新しい状態
 * @param {string} userEmail - 更新者メール
 */
function updateProjectStatusWithValidation(projectId, newStatus, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var project = getProject(projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません');
  }

  // 状態遷移ルールの検証
  var currentStatus = project.status;
  var validTransitions = {
    '下書き': ['確定', 'アーカイブ'],
    '確定': ['アーカイブ'],
    'アーカイブ': [] // アーカイブからは遷移不可
  };

  var allowedTransitions = validTransitions[currentStatus] || [];
  if (allowedTransitions.indexOf(newStatus) === -1) {
    throw new Error('「' + currentStatus + '」から「' + newStatus + '」への状態変更はできません');
  }

  // 確定時の必須チェック
  if (newStatus === PROJECT_STATUS.CONFIRMED) {
    var vision = getVision(projectId);
    if (!vision || !vision.visionText) {
      throw new Error('確定するにはビジョンの設定が必要です');
    }
  }

  updateProjectStatus(projectId, newStatus, userEmail);
}

// ========================================
// プロジェクト検索・取得
// ========================================

/**
 * プロジェクトを検索
 * @param {string} userEmail - ユーザーメール
 * @param {Object} filters - フィルター条件
 * @return {Array<Object>} プロジェクト一覧
 */
function searchProjects(userEmail, filters) {
  var projects = getUserProjects(userEmail);

  if (!filters) {
    return projects;
  }

  // ステータスでフィルタ
  if (filters.status) {
    projects = projects.filter(function(p) {
      return p.status === filters.status;
    });
  }

  // 顧客名で検索
  if (filters.customerName) {
    var searchTerm = filters.customerName.toLowerCase();
    projects = projects.filter(function(p) {
      return p.customerName.toLowerCase().indexOf(searchTerm) !== -1;
    });
  }

  // ソート
  if (filters.sortBy) {
    var sortField = filters.sortBy;
    var sortOrder = filters.sortOrder === 'asc' ? 1 : -1;

    projects.sort(function(a, b) {
      var aVal = a[sortField];
      var bVal = b[sortField];

      if (aVal < bVal) return -1 * sortOrder;
      if (aVal > bVal) return 1 * sortOrder;
      return 0;
    });
  }

  return projects;
}

/**
 * プロジェクトの完成度を計算
 * @param {string} projectId - プロジェクトID
 * @return {Object} 完成度情報
 */
function getProjectCompleteness(projectId) {
  var result = {
    vision: false,
    usecases: false,
    ninetyDayPlan: false,
    organization: false,
    value: false,
    overallPercent: 0
  };

  // ビジョン
  var vision = getVision(projectId);
  if (vision && vision.visionText) {
    result.vision = true;
  }

  // ユースケース
  var usecases = getUsecases(projectId);
  if (usecases && usecases.length > 0) {
    result.usecases = true;

    // 90日計画
    var hasPlans = usecases.some(function(uc) {
      var plan = getNinetyDayPlan(uc.usecaseId);
      return plan && plan.teamStructure;
    });
    result.ninetyDayPlan = hasPlans;
  }

  // 体制RACI
  var raciEntries = getRACIEntries(projectId);
  if (raciEntries && raciEntries.length > 0) {
    result.organization = true;
  }

  // 価値トラッキング
  var values = getValues(projectId);
  if (values && values.length > 0) {
    var hasValue = values.some(function(v) {
      return v.quantitativeImpact || v.qualitativeImpact;
    });
    result.value = hasValue;
  }

  // 完成度パーセント計算
  var completed = 0;
  if (result.vision) completed++;
  if (result.usecases) completed++;
  if (result.ninetyDayPlan) completed++;
  if (result.organization) completed++;
  if (result.value) completed++;

  result.overallPercent = Math.round((completed / 5) * 100);

  return result;
}

/**
 * プロジェクトの詳細情報を取得（関連データ含む）
 * @param {string} projectId - プロジェクトID
 * @param {string} userEmail - ユーザーメール
 * @return {Object} プロジェクト詳細
 */
function getProjectDetails(projectId, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var project = getProject(projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません');
  }

  return {
    project: project,
    vision: getVision(projectId),
    usecases: getUsecases(projectId),
    raciEntries: getRACIEntries(projectId),
    values: getValues(projectId),
    completeness: getProjectCompleteness(projectId)
  };
}

// ========================================
// プロジェクト削除・アーカイブ
// ========================================

/**
 * プロジェクトをアーカイブ
 * @param {string} projectId - プロジェクトID
 * @param {string} userEmail - ユーザーメール
 */
function archiveProject(projectId, userEmail) {
  updateProjectStatusWithValidation(projectId, PROJECT_STATUS.ARCHIVED, userEmail);
}

/**
 * プロジェクトを複製
 * @param {string} projectId - 複製元プロジェクトID
 * @param {string} newCustomerName - 新しい顧客名
 * @param {string} userEmail - ユーザーメール
 * @return {Object} 新しいプロジェクト情報
 */
function duplicateProject(projectId, newCustomerName, userEmail) {
  requireProjectAccess(projectId, userEmail);

  // 元のプロジェクトデータを取得
  var original = getProjectDetails(projectId, userEmail);

  // 新しいプロジェクトを作成
  var newProject = createProjectWithValidation(
    newCustomerName || original.project.customerName + ' (コピー)',
    userEmail
  );

  // ビジョンをコピー
  if (original.vision) {
    saveVision({
      projectId: newProject.projectId,
      visionText: original.vision.visionText,
      decisionRules: original.vision.decisionRules,
      successMetrics: original.vision.successMetrics,
      notes: original.vision.notes,
      userEmail: userEmail
    });
  }

  // ユースケースをコピー
  if (original.usecases) {
    original.usecases.forEach(function(uc) {
      addUsecase({
        projectId: newProject.projectId,
        challenge: uc.challenge,
        goal: uc.goal,
        expectedImpact: uc.expectedImpact,
        ninetyDayGoal: uc.ninetyDayGoal,
        score: uc.score,
        priority: uc.priority,
        userEmail: userEmail
      });
    });
  }

  // RACIエントリをコピー
  if (original.raciEntries && original.raciEntries.length > 0) {
    saveRACIEntries(newProject.projectId, original.raciEntries, userEmail);
  }

  return newProject;
}
