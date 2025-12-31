/**
 * OrganizationService.gs
 * 3本柱組織・RACI設計（Module 4）のビジネスロジック層
 */

// ========================================
// RACI管理
// ========================================

/**
 * RACIエントリを保存（バリデーション付き）
 * @param {string} projectId - プロジェクトID
 * @param {Array<Object>} entries - RACIエントリ配列
 * @param {string} userEmail - ユーザーメール
 */
function saveRACIEntriesWithValidation(projectId, entries, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var project = getProject(projectId);
  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('アーカイブされたプロジェクトは編集できません');
  }

  // バリデーション
  if (!entries || !Array.isArray(entries)) {
    throw new Error('RACIエントリが不正です');
  }

  if (entries.length > 100) {
    throw new Error('RACIエントリは最大100件までです');
  }

  var validPillars = [THREE_PILLARS.COE, THREE_PILLARS.BUSINESS_DATA_NETWORK, THREE_PILLARS.IT];
  var validRaciTypes = [RACI_TYPES.RESPONSIBLE, RACI_TYPES.ACCOUNTABLE, RACI_TYPES.CONSULTED, RACI_TYPES.INFORMED];

  var cleanedEntries = [];

  entries.forEach(function(entry, index) {
    if (!entry.pillar || validPillars.indexOf(entry.pillar) === -1) {
      throw new Error('エントリ ' + (index + 1) + ': 3本柱区分が不正です');
    }

    if (!entry.task || entry.task.trim().length === 0) {
      throw new Error('エントリ ' + (index + 1) + ': タスク名は必須です');
    }

    if (!entry.assignee || entry.assignee.trim().length === 0) {
      throw new Error('エントリ ' + (index + 1) + ': 担当者は必須です');
    }

    if (!entry.raci || validRaciTypes.indexOf(entry.raci) === -1) {
      throw new Error('エントリ ' + (index + 1) + ': RACI区分が不正です');
    }

    cleanedEntries.push({
      pillar: entry.pillar,
      task: entry.task.trim(),
      assignee: entry.assignee.trim(),
      raci: entry.raci
    });
  });

  saveRACIEntries(projectId, cleanedEntries, userEmail);
}

/**
 * RACIマトリクスのバリデーション結果を取得
 * @param {string} projectId - プロジェクトID
 * @return {Object} バリデーション結果
 */
function validateRACIMatrix(projectId) {
  var entries = getRACIEntries(projectId);

  var result = {
    isValid: true,
    warnings: [],
    errors: [],
    summary: {
      totalTasks: 0,
      tasksByPillar: {},
      raciDistribution: {
        R: 0,
        A: 0,
        C: 0,
        I: 0
      }
    }
  };

  if (entries.length === 0) {
    result.warnings.push('RACIマトリクスが設定されていません');
    return result;
  }

  // タスクごとにグループ化
  var taskMap = {};

  entries.forEach(function(entry) {
    if (!taskMap[entry.task]) {
      taskMap[entry.task] = [];
    }
    taskMap[entry.task].push(entry);

    // 分布をカウント
    result.summary.raciDistribution[entry.raci]++;

    // 柱ごとにカウント
    if (!result.summary.tasksByPillar[entry.pillar]) {
      result.summary.tasksByPillar[entry.pillar] = 0;
    }
    result.summary.tasksByPillar[entry.pillar]++;
  });

  var tasks = Object.keys(taskMap);
  result.summary.totalTasks = tasks.length;

  // 各タスクのRACIルールをチェック
  tasks.forEach(function(taskName) {
    var taskEntries = taskMap[taskName];

    var hasResponsible = taskEntries.some(function(e) { return e.raci === 'R'; });
    var hasAccountable = taskEntries.some(function(e) { return e.raci === 'A'; });
    var accountableCount = taskEntries.filter(function(e) { return e.raci === 'A'; }).length;

    if (!hasResponsible) {
      result.warnings.push('タスク「' + taskName + '」にR（実行責任者）がいません');
    }

    if (!hasAccountable) {
      result.warnings.push('タスク「' + taskName + '」にA（説明責任者）がいません');
    }

    if (accountableCount > 1) {
      result.errors.push('タスク「' + taskName + '」にA（説明責任者）が複数います（1名にすべき）');
      result.isValid = false;
    }
  });

  // 3本柱のバランスをチェック
  var pillars = Object.keys(result.summary.tasksByPillar);
  if (pillars.length < 2) {
    result.warnings.push('複数の組織柱（CoE/ビジネスデータネットワーク/IT）を巻き込むことを推奨します');
  }

  return result;
}

// ========================================
// 3本柱組織テンプレート
// ========================================

/**
 * 3本柱組織の推奨体制テンプレートを取得
 * @return {Object} 体制テンプレート
 */
function getThreePillarTemplate() {
  return {
    pillars: [
      {
        name: THREE_PILLARS.COE,
        nameEn: 'Center of Excellence',
        description: 'データ活用の推進・支援・品質管理を担う中核組織',
        recommendedRoles: [
          { role: 'CoEリーダー', raci: 'A', description: '全体統括、経営との橋渡し' },
          { role: 'データアナリスト', raci: 'R', description: 'ダッシュボード開発、データ分析' },
          { role: 'トレーナー', raci: 'R', description: 'ユーザー教育、スキル向上支援' },
          { role: 'データスチュワード', raci: 'C', description: 'データ品質管理、マスタ管理' }
        ],
        typicalTasks: [
          'ダッシュボード開発・保守',
          'ユーザートレーニング',
          'ベストプラクティス策定',
          '利用状況モニタリング'
        ]
      },
      {
        name: THREE_PILLARS.BUSINESS_DATA_NETWORK,
        nameEn: 'Business Data Network',
        description: '各事業部門のデータ活用推進者ネットワーク',
        recommendedRoles: [
          { role: 'ビジネスオーナー', raci: 'A', description: '業務要件の最終承認' },
          { role: 'データチャンピオン', raci: 'R', description: '部門内のデータ活用推進' },
          { role: 'サブジェクトマターエキスパート', raci: 'C', description: '業務知識の提供' }
        ],
        typicalTasks: [
          '業務要件の定義・伝達',
          'ダッシュボードのレビュー・承認',
          '部門内へのデータ活用浸透',
          '効果測定・フィードバック'
        ]
      },
      {
        name: THREE_PILLARS.IT,
        nameEn: 'IT',
        description: 'インフラ・セキュリティ・システム統合を担う',
        recommendedRoles: [
          { role: 'IT責任者', raci: 'A', description: '技術基盤の最終承認' },
          { role: 'インフラエンジニア', raci: 'R', description: 'サーバー・ネットワーク管理' },
          { role: 'セキュリティ担当', raci: 'C', description: 'セキュリティ要件の検証' },
          { role: 'データエンジニア', raci: 'R', description: 'データパイプライン構築' }
        ],
        typicalTasks: [
          'サーバー環境の構築・保守',
          'データソース接続・ETL',
          'セキュリティ・アクセス管理',
          'バックアップ・災害対策'
        ]
      }
    ],
    raciExplanation: {
      R: {
        name: 'Responsible（実行責任）',
        description: '実際にタスクを実行する人。1つのタスクに複数のRがいてもよい。'
      },
      A: {
        name: 'Accountable（説明責任）',
        description: '最終的な承認者・責任者。1つのタスクにつき必ず1人のみ。'
      },
      C: {
        name: 'Consulted（相談）',
        description: 'タスク実行前に意見を求める人。双方向のコミュニケーション。'
      },
      I: {
        name: 'Informed（情報共有）',
        description: 'タスクの進捗や結果を通知される人。一方向のコミュニケーション。'
      }
    }
  };
}

/**
 * デフォルトのRACIエントリを生成
 * @param {string} projectId - プロジェクトID
 * @return {Array<Object>} デフォルトエントリ
 */
function generateDefaultRACIEntries(projectId) {
  return [
    // CoE タスク
    { pillar: THREE_PILLARS.COE, task: 'ダッシュボード開発', assignee: '（担当者名）', raci: 'R' },
    { pillar: THREE_PILLARS.COE, task: 'ユーザートレーニング', assignee: '（担当者名）', raci: 'R' },
    { pillar: THREE_PILLARS.COE, task: '品質レビュー', assignee: '（担当者名）', raci: 'A' },

    // ビジネスデータネットワーク タスク
    { pillar: THREE_PILLARS.BUSINESS_DATA_NETWORK, task: '業務要件定義', assignee: '（担当者名）', raci: 'R' },
    { pillar: THREE_PILLARS.BUSINESS_DATA_NETWORK, task: '成果物承認', assignee: '（担当者名）', raci: 'A' },
    { pillar: THREE_PILLARS.BUSINESS_DATA_NETWORK, task: '効果測定', assignee: '（担当者名）', raci: 'R' },

    // IT タスク
    { pillar: THREE_PILLARS.IT, task: '環境構築', assignee: '（担当者名）', raci: 'R' },
    { pillar: THREE_PILLARS.IT, task: 'データ接続設定', assignee: '（担当者名）', raci: 'R' },
    { pillar: THREE_PILLARS.IT, task: 'セキュリティ審査', assignee: '（担当者名）', raci: 'A' }
  ];
}

// ========================================
// 体制分析
// ========================================

/**
 * 体制のサマリーを取得
 * @param {string} projectId - プロジェクトID
 * @return {Object} サマリー情報
 */
function getOrganizationSummary(projectId) {
  var entries = getRACIEntries(projectId);

  if (entries.length === 0) {
    return {
      configured: false,
      totalEntries: 0,
      uniqueAssignees: 0,
      pillarCoverage: [],
      validation: null
    };
  }

  // ユニークな担当者をカウント
  var assignees = {};
  entries.forEach(function(e) {
    assignees[e.assignee] = true;
  });

  // 柱ごとのエントリ数
  var pillarCoverage = [];
  var pillarCounts = {};

  entries.forEach(function(e) {
    if (!pillarCounts[e.pillar]) {
      pillarCounts[e.pillar] = 0;
    }
    pillarCounts[e.pillar]++;
  });

  Object.keys(pillarCounts).forEach(function(pillar) {
    pillarCoverage.push({
      pillar: pillar,
      count: pillarCounts[pillar]
    });
  });

  return {
    configured: true,
    totalEntries: entries.length,
    uniqueAssignees: Object.keys(assignees).length,
    pillarCoverage: pillarCoverage,
    validation: validateRACIMatrix(projectId)
  };
}

/**
 * RACI完成度をチェック
 * @param {string} projectId - プロジェクトID
 * @return {Object} 完成度情報
 */
function checkRACICompleteness(projectId) {
  var entries = getRACIEntries(projectId);

  var result = {
    hasCoE: false,
    hasBusinessDataNetwork: false,
    hasIT: false,
    hasAccountable: false,
    hasResponsible: false,
    completenessPercent: 0,
    suggestions: []
  };

  if (entries.length === 0) {
    result.suggestions.push('RACIマトリクスを設定してください');
    return result;
  }

  entries.forEach(function(e) {
    if (e.pillar === THREE_PILLARS.COE) result.hasCoE = true;
    if (e.pillar === THREE_PILLARS.BUSINESS_DATA_NETWORK) result.hasBusinessDataNetwork = true;
    if (e.pillar === THREE_PILLARS.IT) result.hasIT = true;
    if (e.raci === 'A') result.hasAccountable = true;
    if (e.raci === 'R') result.hasResponsible = true;
  });

  var score = 0;
  if (result.hasCoE) score += 20;
  if (result.hasBusinessDataNetwork) score += 20;
  if (result.hasIT) score += 20;
  if (result.hasAccountable) score += 20;
  if (result.hasResponsible) score += 20;

  result.completenessPercent = score;

  if (!result.hasCoE) {
    result.suggestions.push('CoE（Center of Excellence）のタスクを追加してください');
  }
  if (!result.hasBusinessDataNetwork) {
    result.suggestions.push('ビジネスデータネットワークのタスクを追加してください');
  }
  if (!result.hasIT) {
    result.suggestions.push('IT部門のタスクを追加してください');
  }
  if (!result.hasAccountable) {
    result.suggestions.push('説明責任者（A）を設定してください');
  }
  if (!result.hasResponsible) {
    result.suggestions.push('実行責任者（R）を設定してください');
  }

  return result;
}
