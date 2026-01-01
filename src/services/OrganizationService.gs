/**
 * OrganizationService.gs
 * Org structure and RACI logic.
 */

function saveRACIEntriesWithValidation(projectId, entries, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var project = getProject(projectId);
  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('アーカイブ済みのプロジェクトは編集できません。');
  }

  if (!entries || !Array.isArray(entries)) {
    throw new Error('RACIエントリーは配列である必要があります。');
  }

  if (entries.length > 100) {
    throw new Error('RACIエントリーは100件以内にしてください。');
  }

  var validPillars = [
    THREE_PILLARS.COE,
    THREE_PILLARS.BUSINESS_DATA_NETWORK,
    THREE_PILLARS.IT
  ];
  var validRaciTypes = [
    RACI_TYPES.RESPONSIBLE,
    RACI_TYPES.ACCOUNTABLE,
    RACI_TYPES.CONSULTED,
    RACI_TYPES.INFORMED
  ];

  var cleanedEntries = [];

  entries.forEach(function(entry, index) {
    if (!entry.pillar || validPillars.indexOf(entry.pillar) === -1) {
      throw new Error('エントリー' + (index + 1) + '：柱が無効です。');
    }

    if (!entry.task || entry.task.trim().length === 0) {
      throw new Error('エントリー' + (index + 1) + '：タスクは必須です。');
    }

    if (!entry.assignee || entry.assignee.trim().length === 0) {
      throw new Error('エントリー' + (index + 1) + '：担当者は必須です。');
    }

    if (!entry.raci || validRaciTypes.indexOf(entry.raci) === -1) {
      throw new Error('エントリー' + (index + 1) + '：RACI種別が無効です。');
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
    result.warnings.push('RACIがまだ設定されていません。');
    return result;
  }

  var taskMap = {};

  entries.forEach(function(entry) {
    if (!taskMap[entry.task]) {
      taskMap[entry.task] = [];
    }
    taskMap[entry.task].push(entry);

    result.summary.raciDistribution[entry.raci]++;

    if (!result.summary.tasksByPillar[entry.pillar]) {
      result.summary.tasksByPillar[entry.pillar] = 0;
    }
    result.summary.tasksByPillar[entry.pillar]++;
  });

  var tasks = Object.keys(taskMap);
  result.summary.totalTasks = tasks.length;

  tasks.forEach(function(taskName) {
    var taskEntries = taskMap[taskName];

    var hasResponsible = taskEntries.some(function(e) { return e.raci === 'R'; });
    var hasAccountable = taskEntries.some(function(e) { return e.raci === 'A'; });
    var accountableCount = taskEntries.filter(function(e) { return e.raci === 'A'; }).length;

    if (!hasResponsible) {
      result.warnings.push('タスク「' + taskName + '」にR（実行責任）が設定されていません。');
    }

    if (!hasAccountable) {
      result.warnings.push('タスク「' + taskName + '」にA（説明責任）が設定されていません。');
    }

    if (accountableCount > 1) {
      result.errors.push('タスク「' + taskName + '」にA（説明責任）が複数設定されています。');
      result.isValid = false;
    }
  });

  var pillars = Object.keys(result.summary.tasksByPillar);
  if (pillars.length < 2) {
    result.warnings.push('複数の柱（CoE、ビジネス・データネットワーク、IT）の関与を検討してください。');
  }

  return result;
}

function getThreePillarTemplate() {
  return {
    pillars: [
      {
        name: THREE_PILLARS.COE,
        nameEn: 'Center of Excellence',
        description: 'イネーブルメント、標準化、品質。',
        recommendedRoles: [
          { role: 'CoEリード', raci: 'A', description: '全体ガバナンスとアラインメント。' },
          { role: 'データアナリスト', raci: 'R', description: 'ダッシュボード開発と洞察提供。' }
        ],
        typicalTasks: [
          'ダッシュボード開発',
          'ベストプラクティスの展開'
        ]
      },
      {
        name: THREE_PILLARS.BUSINESS_DATA_NETWORK,
        nameEn: 'Business Data Network',
        description: 'ビジネスオーナーシップと定着。',
        recommendedRoles: [
          { role: 'ビジネスオーナー', raci: 'A', description: '最終承認と成果責任。' },
          { role: 'データチャンピオン', raci: 'R', description: 'チーム内の定着を推進。' }
        ],
        typicalTasks: [
          '要件定義',
          '定着・チェンジマネジメント'
        ]
      },
      {
        name: THREE_PILLARS.IT,
        nameEn: 'IT',
        description: 'インフラ、セキュリティ、連携。',
        recommendedRoles: [
          { role: 'ITリード', raci: 'A', description: 'プラットフォーム判断とセキュリティ。' },
          { role: 'データエンジニア', raci: 'R', description: 'データパイプラインとアクセス。' }
        ],
        typicalTasks: [
          'データソース連携',
          'アクセス制御とセキュリティ'
        ]
      }
    ],
    raciExplanation: {
      R: {
        name: '実行責任',
        description: 'タスクを実行する。'
      },
      A: {
        name: '説明責任',
        description: '最終責任者（1タスクにつき1名）。'
      },
      C: {
        name: '相談先',
        description: '作業開始前に意見を提供する。'
      },
      I: {
        name: '共有先',
        description: '進捗を共有される。'
      }
    }
  };
}

function generateDefaultRACIEntries(projectId) {
  return [
    { pillar: THREE_PILLARS.COE, task: 'ダッシュボード提供', assignee: '未設定', raci: 'R' },
    { pillar: THREE_PILLARS.COE, task: '品質レビュー', assignee: '未設定', raci: 'A' },
    { pillar: THREE_PILLARS.BUSINESS_DATA_NETWORK, task: '要件定義', assignee: '未設定', raci: 'R' },
    { pillar: THREE_PILLARS.BUSINESS_DATA_NETWORK, task: 'ビジネス承認', assignee: '未設定', raci: 'A' },
    { pillar: THREE_PILLARS.IT, task: 'データアクセス', assignee: '未設定', raci: 'R' },
    { pillar: THREE_PILLARS.IT, task: 'セキュリティレビュー', assignee: '未設定', raci: 'A' }
  ];
}

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

  var assignees = {};
  entries.forEach(function(e) {
    assignees[e.assignee] = true;
  });

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
    result.suggestions.push('責任を定義するためにRACIエントリーを追加してください。');
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
    result.suggestions.push('CoEのタスクを追加してください。');
  }
  if (!result.hasBusinessDataNetwork) {
    result.suggestions.push('ビジネス・データネットワークのタスクを追加してください。');
  }
  if (!result.hasIT) {
    result.suggestions.push('ITのタスクを追加してください。');
  }
  if (!result.hasAccountable) {
    result.suggestions.push('A（説明責任）の役割を少なくとも1つ追加してください。');
  }
  if (!result.hasResponsible) {
    result.suggestions.push('R（実行責任）の役割を少なくとも1つ追加してください。');
  }

  return result;
}
