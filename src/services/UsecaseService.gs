/**
 * UsecaseService.gs
 * Use case selection and 90-day plan logic.
 */

function addUsecaseWithValidation(usecaseData, userEmail) {
  requireProjectAccess(usecaseData.projectId, userEmail);

  var project = getProject(usecaseData.projectId);
  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('アーカイブ済みのプロジェクトは編集できません。');
  }

  if (!usecaseData.challenge || usecaseData.challenge.trim().length === 0) {
    throw new Error('課題は必須です。');
  }

  if (!usecaseData.goal || usecaseData.goal.trim().length === 0) {
    throw new Error('目的は必須です。');
  }

  if (usecaseData.challenge.length > 1000) {
    throw new Error('課題は1000文字以内で入力してください。');
  }

  if (usecaseData.goal.length > 500) {
    throw new Error('目的は500文字以内で入力してください。');
  }

  var score = parseInt(usecaseData.score, 10) || 50;
  if (score < 0 || score > 100) {
    throw new Error('スコアは0〜100の範囲で入力してください。');
  }

  var existingUsecases = getUsecases(usecaseData.projectId);
  if (existingUsecases.length >= 20) {
    throw new Error('1つのプロジェクトに登録できるユースケースは最大20件です。');
  }

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

function updateUsecaseWithValidation(usecaseData, userEmail) {
  requireProjectAccess(usecaseData.projectId, userEmail);

  var usecases = getUsecases(usecaseData.projectId);
  var existing = usecases.find(function(uc) {
    return uc.usecaseId === usecaseData.usecaseId;
  });

  if (!existing) {
    throw new Error('ユースケースが見つかりません。');
  }

  if (!usecaseData.challenge || usecaseData.challenge.trim().length === 0) {
    throw new Error('課題は必須です。');
  }

  if (!usecaseData.goal || usecaseData.goal.trim().length === 0) {
    throw new Error('目的は必須です。');
  }

  var rowData = [];
  rowData[SCHEMA_USECASES.columns.PROJECT_ID] = usecaseData.projectId;
  rowData[SCHEMA_USECASES.columns.USECASE_ID] = usecaseData.usecaseId;
  rowData[SCHEMA_USECASES.columns.CHALLENGE] = usecaseData.challenge.trim();
  rowData[SCHEMA_USECASES.columns.GOAL] = usecaseData.goal.trim();
  rowData[SCHEMA_USECASES.columns.EXPECTED_IMPACT] = (usecaseData.expectedImpact || '').trim();
  rowData[SCHEMA_USECASES.columns.NINETY_DAY_GOAL] = (usecaseData.ninetyDayGoal || '').trim();
  rowData[SCHEMA_USECASES.columns.SCORE] = parseInt(usecaseData.score, 10) || existing.score;
  rowData[SCHEMA_USECASES.columns.PRIORITY] = usecaseData.priority || existing.priority;
  rowData[SCHEMA_USECASES.columns.UPDATED_DATE] = getCurrentTimestamp();

  updateRow(SCHEMA_USECASES.sheetName, existing._rowNumber, rowData);

  logAudit(userEmail, OPERATION_TYPES.UPDATE, usecaseData.projectId, {
    module: 'usecase',
    usecaseId: usecaseData.usecaseId
  });
}

function deleteUsecaseWithValidation(projectId, usecaseId, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var usecases = getUsecases(projectId);
  var target = usecases.find(function(uc) {
    return uc.usecaseId === usecaseId;
  });

  if (!target) {
    throw new Error('ユースケースが見つかりません。');
  }

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

  deleteRow(SCHEMA_USECASES.sheetName, target._rowNumber);

  logAudit(userEmail, OPERATION_TYPES.DELETE, projectId, {
    module: 'usecase',
    usecaseId: usecaseId
  });
}

function recalculatePriorities(projectId, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var usecases = getUsecases(projectId);

  usecases.sort(function(a, b) {
    return (b.score || 0) - (a.score || 0);
  });

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

function getScoringCriteria() {
  return {
    criteria: [
      {
        name: 'Business impact',
        description: 'Expected business value and visibility.',
        weight: 30,
        levels: [
          { score: 10, label: 'Low' },
          { score: 20, label: 'Medium' },
          { score: 30, label: 'High' }
        ]
      },
      {
        name: 'Feasibility',
        description: 'Likelihood of delivering results in 90 days.',
        weight: 25,
        levels: [
          { score: 8, label: 'Hard' },
          { score: 16, label: 'Possible' },
          { score: 25, label: 'Easy' }
        ]
      },
      {
        name: 'Data readiness',
        description: 'Availability and quality of data.',
        weight: 20,
        levels: [
          { score: 5, label: 'Not ready' },
          { score: 12, label: 'Partial' },
          { score: 20, label: 'Ready' }
        ]
      },
      {
        name: 'Executive support',
        description: 'Level of leadership sponsorship.',
        weight: 15,
        levels: [
          { score: 5, label: 'Low' },
          { score: 10, label: 'Medium' },
          { score: 15, label: 'High' }
        ]
      },
      {
        name: 'Stakeholder alignment',
        description: 'Ability to align key stakeholders.',
        weight: 10,
        levels: [
          { score: 3, label: 'Hard' },
          { score: 6, label: 'Possible' },
          { score: 10, label: 'Strong' }
        ]
      }
    ],
    maxScore: 100
  };
}

function saveNinetyDayPlanWithValidation(planData, userEmail) {
  requireProjectAccess(planData.projectId, userEmail);

  var usecases = getUsecases(planData.projectId);
  var exists = usecases.some(function(uc) {
    return uc.usecaseId === planData.usecaseId;
  });

  if (!exists) {
    throw new Error('ユースケースが見つかりません。');
  }

  var milestones = planData.weeklyMilestones || [];
  if (milestones.length > 12) {
    throw new Error('週次マイルストーンは12週以内で入力してください。');
  }

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

function getNinetyDayPlanTemplate() {
  return {
    teamStructure: 'Roles:\\n- Project lead\\n- Data analyst\\n- Business owner\\n\\nRecommended:\\n- CoE: quality and enablement\\n- IT: security and infrastructure\\n- Business: requirements and adoption',
    requiredData: 'Data sources:\\n- Sales history\\n- Customer master\\n- Product catalog\\n\\nAccess methods:\\n- Export or API integration',
    risks: 'Risks:\\n- Data quality issues\\n- Integration delays\\n- Stakeholder changes\\n\\nMitigations:\\n- Early PoC\\n- Weekly steering check-ins',
    communicationPlan: 'Cadence:\\n- Weekly status\\n- Monthly executive review\\n\\nChannels:\\n- Chat for daily updates\\n- Email for approvals',
    weeklyMilestones: [
      'Week 1-2: Kickoff, scope, success criteria',
      'Week 3-4: Data access and first dashboard',
      'Week 5-6: User review and iteration',
      'Week 7-8: Training and rollout',
      'Week 9-10: Optimization',
      'Week 11-12: Value measurement and next plan'
    ]
  };
}

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

  var sorted = usecases.slice().sort(function(a, b) {
    return (b.score || 0) - (a.score || 0);
  });

  var totalScore = usecases.reduce(function(sum, uc) {
    return sum + (uc.score || 0);
  }, 0);
  var averageScore = Math.round(totalScore / usecases.length);

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
