/**
 * UsecaseService.gs
 * ユースケース選定・90日計画（Module 2）のビジネスロジック層
 */

// ========================================
// ユースケース管理
// ========================================

/**
 * ユースケースを追加（バリデーション付き）
 * @param {Object} usecaseData - ユースケースデータ
 * @param {string} userEmail - ユーザーメール
 * @return {string} 生成されたユースケースID
 */
function addUsecaseWithValidation(usecaseData, userEmail) {
  requireProjectAccess(usecaseData.projectId, userEmail);

  var project = getProject(usecaseData.projectId);
  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('アーカイブされたプロジェクトは編集できません');
  }

  // バリデーション
  if (!usecaseData.challenge || usecaseData.challenge.trim().length === 0) {
    throw new Error('課題は必須です');
  }

  if (!usecaseData.goal || usecaseData.goal.trim().length === 0) {
    throw new Error('狙い（ゴール）は必須です');
  }

  if (usecaseData.challenge.length > 1000) {
    throw new Error('課題は1000文字以内で入力してください');
  }

  if (usecaseData.goal.length > 500) {
    throw new Error('狙いは500文字以内で入力してください');
  }

  // スコアの範囲チェック
  var score = parseInt(usecaseData.score) || 50;
  if (score < 0 || score > 100) {
    throw new Error('スコアは0〜100の範囲で入力してください');
  }

  // 既存のユースケース数をチェック
  var existingUsecases = getUsecases(usecaseData.projectId);
  if (existingUsecases.length >= 20) {
    throw new Error('ユースケースは最大20件までです');
  }

  // データを整形
  var cleanData = {
    projectId: usecaseData.projectId,
    challenge: usecaseData.challenge.trim(),
    goal: usecaseData.goal.trim(),
    expectedImpact: (usecaseData.expectedImpact || '').trim(),
    ninetyDayGoal: (usecaseData.ninetyDayGoal || '').trim(),
    score: score,
    priority: existingUsecases.length + 1,
    userEmail: userEmail
  };

  return addUsecase(cleanData);
}

/**
 * ユースケースを更新
 * @param {Object} usecaseData - 更新データ
 * @param {string} userEmail - ユーザーメール
 */
function updateUsecaseWithValidation(usecaseData, userEmail) {
  requireProjectAccess(usecaseData.projectId, userEmail);

  var usecases = getUsecases(usecaseData.projectId);
  var existing = usecases.find(function(uc) {
    return uc.usecaseId === usecaseData.usecaseId;
  });

  if (!existing) {
    throw new Error('ユースケースが見つかりません');
  }

  // バリデーション
  if (!usecaseData.challenge || usecaseData.challenge.trim().length === 0) {
    throw new Error('課題は必須です');
  }

  if (!usecaseData.goal || usecaseData.goal.trim().length === 0) {
    throw new Error('狙い（ゴール）は必須です');
  }

  var rowData = [];
  rowData[SCHEMA_USECASES.columns.PROJECT_ID] = usecaseData.projectId;
  rowData[SCHEMA_USECASES.columns.USECASE_ID] = usecaseData.usecaseId;
  rowData[SCHEMA_USECASES.columns.CHALLENGE] = usecaseData.challenge.trim();
  rowData[SCHEMA_USECASES.columns.GOAL] = usecaseData.goal.trim();
  rowData[SCHEMA_USECASES.columns.EXPECTED_IMPACT] = (usecaseData.expectedImpact || '').trim();
  rowData[SCHEMA_USECASES.columns.NINETY_DAY_GOAL] = (usecaseData.ninetyDayGoal || '').trim();
  rowData[SCHEMA_USECASES.columns.SCORE] = parseInt(usecaseData.score) || existing.score;
  rowData[SCHEMA_USECASES.columns.PRIORITY] = usecaseData.priority || existing.priority;
  rowData[SCHEMA_USECASES.columns.UPDATED_DATE] = getCurrentTimestamp();

  updateRow(SCHEMA_USECASES.sheetName, existing._rowNumber, rowData);

  logAudit(userEmail, OPERATION_TYPES.UPDATE, usecaseData.projectId, {
    module: 'usecase',
    usecaseId: usecaseData.usecaseId
  });
}

/**
 * ユースケースを削除
 * @param {string} projectId - プロジェクトID
 * @param {string} usecaseId - ユースケースID
 * @param {string} userEmail - ユーザーメール
 */
function deleteUsecaseWithValidation(projectId, usecaseId, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var usecases = getUsecases(projectId);
  var target = usecases.find(function(uc) {
    return uc.usecaseId === usecaseId;
  });

  if (!target) {
    throw new Error('ユースケースが見つかりません');
  }

  // 関連データ（90日計画、価値トラッキング）も削除
  var plan = getNinetyDayPlan(usecaseId);
  if (plan) {
    var planResult = findFirstRow(
      SCHEMA_NINETY_DAY_PLAN.sheetName,
      SCHEMA_NINETY_DAY_PLAN.columns.USECASE_ID,
      usecaseId
    );
    if (planResult) {
      deleteRow(SCHEMA_NINETY_DAY_PLAN.sheetName, planResult.rowNumber);
    }
  }

  var valueResult = findFirstRow(
    SCHEMA_VALUE.sheetName,
    SCHEMA_VALUE.columns.USECASE_ID,
    usecaseId
  );
  if (valueResult) {
    deleteRow(SCHEMA_VALUE.sheetName, valueResult.rowNumber);
  }

  // ユースケース本体を削除
  deleteRow(SCHEMA_USECASES.sheetName, target._rowNumber);

  logAudit(userEmail, OPERATION_TYPES.DELETE, projectId, {
    module: 'usecase',
    usecaseId: usecaseId
  });
}

// ========================================
// 優先度スコアリング
// ========================================

/**
 * ユースケースの優先度を再計算
 * @param {string} projectId - プロジェクトID
 * @param {string} userEmail - ユーザーメール
 */
function recalculatePriorities(projectId, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var usecases = getUsecases(projectId);

  // スコア順にソート
  usecases.sort(function(a, b) {
    return (b.score || 0) - (a.score || 0);
  });

  // 優先度を更新
  usecases.forEach(function(uc, index) {
    var rowData = [];
    rowData[SCHEMA_USECASES.columns.PROJECT_ID] = uc.projectId;
    rowData[SCHEMA_USECASES.columns.USECASE_ID] = uc.usecaseId;
    rowData[SCHEMA_USECASES.columns.CHALLENGE] = uc.challenge;
    rowData[SCHEMA_USECASES.columns.GOAL] = uc.goal;
    rowData[SCHEMA_USECASES.columns.EXPECTED_IMPACT] = uc.expectedImpact;
    rowData[SCHEMA_USECASES.columns.NINETY_DAY_GOAL] = uc.ninetyDayGoal;
    rowData[SCHEMA_USECASES.columns.SCORE] = uc.score;
    rowData[SCHEMA_USECASES.columns.PRIORITY] = index + 1;
    rowData[SCHEMA_USECASES.columns.UPDATED_DATE] = getCurrentTimestamp();

    updateRow(SCHEMA_USECASES.sheetName, uc._rowNumber, rowData);
  });

  logAudit(userEmail, OPERATION_TYPES.UPDATE, projectId, {
    action: 'recalculate_priorities',
    count: usecases.length
  });
}

/**
 * ユースケースのスコアリング基準を取得
 * @return {Object} スコアリング基準
 */
function getScoringCriteria() {
  return {
    criteria: [
      {
        name: 'ビジネスインパクト',
        description: 'このユースケースが事業に与える影響の大きさ',
        weight: 30,
        levels: [
          { score: 10, label: '低い' },
          { score: 20, label: '中程度' },
          { score: 30, label: '高い' }
        ]
      },
      {
        name: '実現可能性',
        description: '90日以内に成果を出せる可能性',
        weight: 25,
        levels: [
          { score: 8, label: '困難' },
          { score: 16, label: '可能' },
          { score: 25, label: '容易' }
        ]
      },
      {
        name: 'データ準備度',
        description: '必要なデータが利用可能か',
        weight: 20,
        levels: [
          { score: 5, label: '未整備' },
          { score: 12, label: '部分的' },
          { score: 20, label: '準備完了' }
        ]
      },
      {
        name: '経営の関心度',
        description: '経営層がこのテーマに関心を持っているか',
        weight: 15,
        levels: [
          { score: 5, label: '低い' },
          { score: 10, label: '中程度' },
          { score: 15, label: '高い' }
        ]
      },
      {
        name: 'ステークホルダーの巻き込み',
        description: '関係者の協力が得られそうか',
        weight: 10,
        levels: [
          { score: 3, label: '困難' },
          { score: 6, label: '可能' },
          { score: 10, label: '積極的' }
        ]
      }
    ],
    maxScore: 100
  };
}

// ========================================
// 90日計画管理
// ========================================

/**
 * 90日計画を保存（バリデーション付き）
 * @param {Object} planData - 計画データ
 * @param {string} userEmail - ユーザーメール
 */
function saveNinetyDayPlanWithValidation(planData, userEmail) {
  requireProjectAccess(planData.projectId, userEmail);

  // ユースケースの存在確認
  var usecases = getUsecases(planData.projectId);
  var exists = usecases.some(function(uc) {
    return uc.usecaseId === planData.usecaseId;
  });

  if (!exists) {
    throw new Error('指定されたユースケースが見つかりません');
  }

  // 週次マイルストーンのバリデーション
  var milestones = planData.weeklyMilestones || [];
  if (milestones.length > 12) {
    throw new Error('週次マイルストーンは最大12週分までです');
  }

  // データを整形
  var cleanData = {
    projectId: planData.projectId,
    usecaseId: planData.usecaseId,
    teamStructure: (planData.teamStructure || '').trim(),
    requiredData: (planData.requiredData || '').trim(),
    risks: (planData.risks || '').trim(),
    communicationPlan: (planData.communicationPlan || '').trim(),
    weeklyMilestones: milestones.map(function(m) { return m.trim(); }),
    userEmail: userEmail
  };

  saveNinetyDayPlan(cleanData);
}

/**
 * 90日計画のテンプレートを取得
 * @return {Object} 計画テンプレート
 */
function getNinetyDayPlanTemplate() {
  return {
    teamStructure: '【必須役割】\n- プロジェクトリーダー: 1名\n- データ分析担当: 1名\n- 業務担当: 1-2名\n\n【推奨体制】\n- CoE: 技術支援・品質管理\n- IT: インフラ・セキュリティ\n- ビジネス: ユーザー代表・要件定義',
    requiredData: '【データソース】\n- 売上データ（期間: 過去2年分）\n- 顧客マスタ\n- 商品マスタ\n\n【取得方法】\n- 基幹システムからのエクスポート\n- API連携（可否を確認）',
    risks: '【技術リスク】\n- データ品質の課題\n- システム連携の遅延\n\n【組織リスク】\n- キーパーソンの異動\n- 優先度の変更\n\n【対策】\n- 早期のPoCで技術検証\n- エグゼクティブスポンサーの確保',
    communicationPlan: '【定例会議】\n- 週次進捗会議: 毎週金曜 30分\n- 月次報告: 毎月最終週\n\n【コミュニケーションチャネル】\n- チャット: 日常連絡\n- メール: 正式な依頼・報告',
    weeklyMilestones: [
      'Week 1-2: キックオフ、ゴール・スコープ合意',
      'Week 3-4: データ接続、初期ダッシュボード作成',
      'Week 5-6: ユーザーレビュー、フィードバック反映',
      'Week 7-8: トレーニング、並行運用開始',
      'Week 9-10: 改善・チューニング',
      'Week 11-12: 成果測定、次フェーズ計画策定'
    ]
  };
}

/**
 * 90日計画の進捗を計算
 * @param {string} usecaseId - ユースケースID
 * @return {Object} 進捗情報
 */
function calculatePlanProgress(usecaseId) {
  var plan = getNinetyDayPlan(usecaseId);

  if (!plan) {
    return {
      hasTeamStructure: false,
      hasRequiredData: false,
      hasRisks: false,
      hasCommunicationPlan: false,
      hasMilestones: false,
      completenessPercent: 0
    };
  }

  var result = {
    hasTeamStructure: plan.teamStructure && plan.teamStructure.length > 0,
    hasRequiredData: plan.requiredData && plan.requiredData.length > 0,
    hasRisks: plan.risks && plan.risks.length > 0,
    hasCommunicationPlan: plan.communicationPlan && plan.communicationPlan.length > 0,
    hasMilestones: plan.weeklyMilestones && plan.weeklyMilestones.length > 0,
    completenessPercent: 0
  };

  var completed = 0;
  if (result.hasTeamStructure) completed++;
  if (result.hasRequiredData) completed++;
  if (result.hasRisks) completed++;
  if (result.hasCommunicationPlan) completed++;
  if (result.hasMilestones) completed++;

  result.completenessPercent = Math.round((completed / 5) * 100);

  return result;
}

// ========================================
// ユースケース分析
// ========================================

/**
 * ユースケースのサマリーを取得
 * @param {string} projectId - プロジェクトID
 * @return {Object} サマリー情報
 */
function getUsecaseSummary(projectId) {
  var usecases = getUsecases(projectId);

  if (usecases.length === 0) {
    return {
      totalCount: 0,
      topUsecases: [],
      averageScore: 0,
      plansCreated: 0
    };
  }

  // スコア順にソート
  var sorted = usecases.slice().sort(function(a, b) {
    return (b.score || 0) - (a.score || 0);
  });

  // 平均スコア
  var totalScore = usecases.reduce(function(sum, uc) {
    return sum + (uc.score || 0);
  }, 0);
  var averageScore = Math.round(totalScore / usecases.length);

  // 90日計画の作成数
  var plansCreated = 0;
  usecases.forEach(function(uc) {
    var plan = getNinetyDayPlan(uc.usecaseId);
    if (plan && plan.teamStructure) {
      plansCreated++;
    }
  });

  return {
    totalCount: usecases.length,
    topUsecases: sorted.slice(0, 3),
    averageScore: averageScore,
    plansCreated: plansCreated
  };
}
