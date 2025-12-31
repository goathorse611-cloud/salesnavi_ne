/**
 * SpreadsheetRepository.gs
 * Spreadsheet-backed storage helpers.
 */

// ========================================
// Sheet initialization
// ========================================

function initializeAllSheets() {
  var ss = getSpreadsheet();
  var schemas = [
    SCHEMA_PROJECTS,
    SCHEMA_VISION,
    SCHEMA_USECASES,
    SCHEMA_NINETY_DAY_PLAN,
    SCHEMA_ORGANIZATION_RACI,
    SCHEMA_CORE_PROCESS,
    SCHEMA_GOVERNANCE,
    SCHEMA_OPERATIONS_SUPPORT,
    SCHEMA_VALUE,
    SCHEMA_AUDIT_LOG
  ];

  schemas.forEach(function(schema) {
    initializeSheet(ss, schema);
  });

  Logger.log('All sheets initialized.');
  return { success: true, message: 'All sheets initialized.' };
}

function initializeSheet(ss, schema) {
  var sheet = ss.getSheetByName(schema.sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(schema.sheetName);
    Logger.log('Created sheet: ' + schema.sheetName);
  }

  var headers = sheet.getRange(1, 1, 1, schema.headers.length).getValues()[0];
  var headersEmpty = headers.every(function(h) { return h === ''; });

  if (headersEmpty) {
    sheet.getRange(1, 1, 1, schema.headers.length).setValues([schema.headers]);
    sheet.getRange(1, 1, 1, schema.headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    Logger.log('Headers set for sheet: ' + schema.sheetName);
  }
}

function getSheet(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName);
  }
  return sheet;
}

// ========================================
// Generic CRUD
// ========================================

function appendRow(sheetName, rowData) {
  var sheet = getSheet(sheetName);
  sheet.appendRow(rowData);
  return sheet.getLastRow();
}

function updateRow(sheetName, rowNumber, rowData) {
  var sheet = getSheet(sheetName);
  sheet.getRange(rowNumber, 1, 1, rowData.length).setValues([rowData]);
}

function deleteRow(sheetName, rowNumber) {
  var sheet = getSheet(sheetName);
  sheet.deleteRow(rowNumber);
}

function getAllRows(sheetName) {
  var sheet = getSheet(sheetName);
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  var lastCol = sheet.getLastColumn();
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function findRows(sheetName, columnIndex, value) {
  var rows = getAllRows(sheetName);
  var results = [];

  rows.forEach(function(row, index) {
    if (row[columnIndex] === value) {
      results.push({
        rowNumber: index + 2,
        data: row
      });
    }
  });

  return results;
}

function findFirstRow(sheetName, columnIndex, value) {
  var results = findRows(sheetName, columnIndex, value);
  return results.length > 0 ? results[0] : null;
}

// ========================================
// Projects
// ========================================

function createProject(customerName, userEmail) {
  var projectId = generateProjectId();
  var now = getCurrentTimestamp();

  var rowData = [];
  rowData[SCHEMA_PROJECTS.columns.PROJECT_ID] = projectId;
  rowData[SCHEMA_PROJECTS.columns.CUSTOMER_NAME] = customerName;
  rowData[SCHEMA_PROJECTS.columns.CREATED_DATE] = now;
  rowData[SCHEMA_PROJECTS.columns.CREATED_BY] = userEmail;
  rowData[SCHEMA_PROJECTS.columns.EDITORS] = userEmail;
  rowData[SCHEMA_PROJECTS.columns.STATUS] = PROJECT_STATUS.DRAFT;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_DATE] = now;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_BY] = userEmail;

  appendRow(SCHEMA_PROJECTS.sheetName, rowData);

  logAudit(userEmail, OPERATION_TYPES.CREATE, projectId, { customerName: customerName });

  return {
    projectId: projectId,
    customerName: customerName,
    status: PROJECT_STATUS.DRAFT,
    createdDate: now
  };
}

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

function getUserProjects(userEmail) {
  var rows = getAllRows(SCHEMA_PROJECTS.sheetName);
  var projects = [];

  rows.forEach(function(row) {
    var editors = row[SCHEMA_PROJECTS.columns.EDITORS] || '';
    var editorList = editors.split(',').map(function(e) { return e.trim(); });

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

function updateProjectStatus(projectId, status, userEmail) {
  var project = getProject(projectId);
  if (!project) {
    throw new Error('Project not found: ' + projectId);
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

  logAudit(userEmail, OPERATION_TYPES.UPDATE, projectId, {
    field: 'status',
    oldValue: project.status,
    newValue: status
  });
}

// ========================================
// Vision
// ========================================

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
// Use cases
// ========================================

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
// 90-day plan
// ========================================

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
// RACI
// ========================================

function saveRACIEntries(projectId, raciEntries, userEmail) {
  var existingRows = findRows(
    SCHEMA_ORGANIZATION_RACI.sheetName,
    SCHEMA_ORGANIZATION_RACI.columns.PROJECT_ID,
    projectId
  );

  existingRows.reverse().forEach(function(row) {
    deleteRow(SCHEMA_ORGANIZATION_RACI.sheetName, row.rowNumber);
  });

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
// Value tracking
// ========================================

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
// Audit logs
// ========================================

function logAudit(userEmail, operation, projectId, details) {
  var rowData = [];
  rowData[SCHEMA_AUDIT_LOG.columns.TIMESTAMP] = getCurrentTimestamp();
  rowData[SCHEMA_AUDIT_LOG.columns.USER] = userEmail;
  rowData[SCHEMA_AUDIT_LOG.columns.OPERATION] = operation;
  rowData[SCHEMA_AUDIT_LOG.columns.PROJECT_ID] = projectId;
  rowData[SCHEMA_AUDIT_LOG.columns.DETAILS] = safeJsonStringify(details);

  appendRow(SCHEMA_AUDIT_LOG.sheetName, rowData);
}

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

// ========================================
// Governance (Module 5)
// ========================================

function saveGovernanceEntries(projectId, governanceEntries, userEmail) {
  var existingRows = findRows(
    SCHEMA_GOVERNANCE.sheetName,
    SCHEMA_GOVERNANCE.columns.PROJECT_ID,
    projectId
  );

  existingRows.reverse().forEach(function(row) {
    deleteRow(SCHEMA_GOVERNANCE.sheetName, row.rowNumber);
  });

  governanceEntries.forEach(function(entry) {
    var rowData = [];
    rowData[SCHEMA_GOVERNANCE.columns.PROJECT_ID] = projectId;
    rowData[SCHEMA_GOVERNANCE.columns.TARGET] = entry.target || '';
    rowData[SCHEMA_GOVERNANCE.columns.MODEL] = entry.model || '';
    rowData[SCHEMA_GOVERNANCE.columns.OWNER] = entry.owner || '';
    rowData[SCHEMA_GOVERNANCE.columns.RULES] = entry.rules || '';
    rowData[SCHEMA_GOVERNANCE.columns.EXCEPTION_PROCESS] = entry.exceptionProcess || '';
    rowData[SCHEMA_GOVERNANCE.columns.UPDATED_DATE] = getCurrentTimestamp();

    appendRow(SCHEMA_GOVERNANCE.sheetName, rowData);
  });

  logAudit(userEmail, OPERATION_TYPES.UPDATE, projectId, {
    module: 'governance',
    count: governanceEntries.length
  });
}

function getGovernanceEntries(projectId) {
  var results = findRows(
    SCHEMA_GOVERNANCE.sheetName,
    SCHEMA_GOVERNANCE.columns.PROJECT_ID,
    projectId
  );

  return results.map(function(result) {
    var data = result.data;
    return {
      target: data[SCHEMA_GOVERNANCE.columns.TARGET],
      model: data[SCHEMA_GOVERNANCE.columns.MODEL],
      owner: data[SCHEMA_GOVERNANCE.columns.OWNER],
      rules: data[SCHEMA_GOVERNANCE.columns.RULES],
      exceptionProcess: data[SCHEMA_GOVERNANCE.columns.EXCEPTION_PROCESS],
      updatedDate: data[SCHEMA_GOVERNANCE.columns.UPDATED_DATE]
    };
  });
}

function addGovernanceEntry(governanceData, userEmail) {
  var rowData = [];
  rowData[SCHEMA_GOVERNANCE.columns.PROJECT_ID] = governanceData.projectId;
  rowData[SCHEMA_GOVERNANCE.columns.TARGET] = governanceData.target || '';
  rowData[SCHEMA_GOVERNANCE.columns.MODEL] = governanceData.model || '';
  rowData[SCHEMA_GOVERNANCE.columns.OWNER] = governanceData.owner || '';
  rowData[SCHEMA_GOVERNANCE.columns.RULES] = governanceData.rules || '';
  rowData[SCHEMA_GOVERNANCE.columns.EXCEPTION_PROCESS] = governanceData.exceptionProcess || '';
  rowData[SCHEMA_GOVERNANCE.columns.UPDATED_DATE] = getCurrentTimestamp();

  appendRow(SCHEMA_GOVERNANCE.sheetName, rowData);

  logAudit(userEmail, OPERATION_TYPES.CREATE, governanceData.projectId, {
    module: 'governance',
    target: governanceData.target
  });
}

// ========================================
// Operations support (Module 6)
// ========================================

function saveOperationsSupport(supportData, userEmail) {
  var existing = findFirstRow(
    SCHEMA_OPERATIONS_SUPPORT.sheetName,
    SCHEMA_OPERATIONS_SUPPORT.columns.PROJECT_ID,
    supportData.projectId
  );

  var rowData = [];
  rowData[SCHEMA_OPERATIONS_SUPPORT.columns.PROJECT_ID] = supportData.projectId;
  rowData[SCHEMA_OPERATIONS_SUPPORT.columns.L1_SUPPORT] = supportData.l1Support || '';
  rowData[SCHEMA_OPERATIONS_SUPPORT.columns.L2_SUPPORT] = supportData.l2Support || '';
  rowData[SCHEMA_OPERATIONS_SUPPORT.columns.L3_SUPPORT] = supportData.l3Support || '';
  rowData[SCHEMA_OPERATIONS_SUPPORT.columns.L4_SUPPORT] = supportData.l4Support || '';
  rowData[SCHEMA_OPERATIONS_SUPPORT.columns.FAQ_LINK] = supportData.faqLink || '';
  rowData[SCHEMA_OPERATIONS_SUPPORT.columns.ESCALATION_CRITERIA] = supportData.escalationCriteria || '';
  rowData[SCHEMA_OPERATIONS_SUPPORT.columns.COMMUNITY_OPS] = supportData.communityOps || '';
  rowData[SCHEMA_OPERATIONS_SUPPORT.columns.UPDATED_DATE] = getCurrentTimestamp();

  if (existing) {
    updateRow(SCHEMA_OPERATIONS_SUPPORT.sheetName, existing.rowNumber, rowData);
    logAudit(userEmail, OPERATION_TYPES.UPDATE, supportData.projectId, { module: 'operations-support' });
  } else {
    appendRow(SCHEMA_OPERATIONS_SUPPORT.sheetName, rowData);
    logAudit(userEmail, OPERATION_TYPES.CREATE, supportData.projectId, { module: 'operations-support' });
  }
}

function getOperationsSupport(projectId) {
  var result = findFirstRow(
    SCHEMA_OPERATIONS_SUPPORT.sheetName,
    SCHEMA_OPERATIONS_SUPPORT.columns.PROJECT_ID,
    projectId
  );

  if (!result) return null;

  var data = result.data;
  return {
    projectId: data[SCHEMA_OPERATIONS_SUPPORT.columns.PROJECT_ID],
    l1Support: data[SCHEMA_OPERATIONS_SUPPORT.columns.L1_SUPPORT],
    l2Support: data[SCHEMA_OPERATIONS_SUPPORT.columns.L2_SUPPORT],
    l3Support: data[SCHEMA_OPERATIONS_SUPPORT.columns.L3_SUPPORT],
    l4Support: data[SCHEMA_OPERATIONS_SUPPORT.columns.L4_SUPPORT],
    faqLink: data[SCHEMA_OPERATIONS_SUPPORT.columns.FAQ_LINK],
    escalationCriteria: data[SCHEMA_OPERATIONS_SUPPORT.columns.ESCALATION_CRITERIA],
    communityOps: data[SCHEMA_OPERATIONS_SUPPORT.columns.COMMUNITY_OPS],
    updatedDate: data[SCHEMA_OPERATIONS_SUPPORT.columns.UPDATED_DATE]
  };
}

// ========================================
// Usecase lookup
// ========================================

function getUsecaseById(usecaseId) {
  var result = findFirstRow(
    SCHEMA_USECASES.sheetName,
    SCHEMA_USECASES.columns.USECASE_ID,
    usecaseId
  );

  if (!result) return null;

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
}

// ========================================
// Core Process (Module 3)
// ========================================

function saveCoreProcess(processData, userEmail) {
  var existing = findFirstRow(
    SCHEMA_CORE_PROCESS.sheetName,
    SCHEMA_CORE_PROCESS.columns.PROJECT_ID,
    processData.projectId
  );

  var rowData = [];
  rowData[SCHEMA_CORE_PROCESS.columns.PROJECT_ID] = processData.projectId;
  rowData[SCHEMA_CORE_PROCESS.columns.AGILITY_SCORE] = processData.agilityScore || 0;
  rowData[SCHEMA_CORE_PROCESS.columns.SKILLS_SCORE] = processData.skillsScore || 0;
  rowData[SCHEMA_CORE_PROCESS.columns.DATA_QUALITY_SCORE] = processData.dataQualityScore || 0;
  rowData[SCHEMA_CORE_PROCESS.columns.TRUST_SCORE] = processData.trustScore || 0;
  rowData[SCHEMA_CORE_PROCESS.columns.OPERATIONAL_EFFICIENCY_SCORE] = processData.operationalEfficiencyScore || 0;
  rowData[SCHEMA_CORE_PROCESS.columns.COMMUNITY_SCORE] = processData.communityScore || 0;
  rowData[SCHEMA_CORE_PROCESS.columns.AGILITY_COMMENT] = processData.agilityComment || '';
  rowData[SCHEMA_CORE_PROCESS.columns.SKILLS_COMMENT] = processData.skillsComment || '';
  rowData[SCHEMA_CORE_PROCESS.columns.DATA_QUALITY_COMMENT] = processData.dataQualityComment || '';
  rowData[SCHEMA_CORE_PROCESS.columns.TRUST_COMMENT] = processData.trustComment || '';
  rowData[SCHEMA_CORE_PROCESS.columns.OPERATIONAL_EFFICIENCY_COMMENT] = processData.operationalEfficiencyComment || '';
  rowData[SCHEMA_CORE_PROCESS.columns.COMMUNITY_COMMENT] = processData.communityComment || '';
  rowData[SCHEMA_CORE_PROCESS.columns.UPDATED_DATE] = getCurrentTimestamp();

  if (existing) {
    updateRow(SCHEMA_CORE_PROCESS.sheetName, existing.rowNumber, rowData);
    logAudit(userEmail, OPERATION_TYPES.UPDATE, processData.projectId, { module: 'core-process' });
  } else {
    appendRow(SCHEMA_CORE_PROCESS.sheetName, rowData);
    logAudit(userEmail, OPERATION_TYPES.CREATE, processData.projectId, { module: 'core-process' });
  }
}

function getCoreProcess(projectId) {
  var result = findFirstRow(
    SCHEMA_CORE_PROCESS.sheetName,
    SCHEMA_CORE_PROCESS.columns.PROJECT_ID,
    projectId
  );

  if (!result) return null;

  var data = result.data;
  return {
    projectId: data[SCHEMA_CORE_PROCESS.columns.PROJECT_ID],
    agilityScore: data[SCHEMA_CORE_PROCESS.columns.AGILITY_SCORE] || 0,
    skillsScore: data[SCHEMA_CORE_PROCESS.columns.SKILLS_SCORE] || 0,
    dataQualityScore: data[SCHEMA_CORE_PROCESS.columns.DATA_QUALITY_SCORE] || 0,
    trustScore: data[SCHEMA_CORE_PROCESS.columns.TRUST_SCORE] || 0,
    operationalEfficiencyScore: data[SCHEMA_CORE_PROCESS.columns.OPERATIONAL_EFFICIENCY_SCORE] || 0,
    communityScore: data[SCHEMA_CORE_PROCESS.columns.COMMUNITY_SCORE] || 0,
    agilityComment: data[SCHEMA_CORE_PROCESS.columns.AGILITY_COMMENT],
    skillsComment: data[SCHEMA_CORE_PROCESS.columns.SKILLS_COMMENT],
    dataQualityComment: data[SCHEMA_CORE_PROCESS.columns.DATA_QUALITY_COMMENT],
    trustComment: data[SCHEMA_CORE_PROCESS.columns.TRUST_COMMENT],
    operationalEfficiencyComment: data[SCHEMA_CORE_PROCESS.columns.OPERATIONAL_EFFICIENCY_COMMENT],
    communityComment: data[SCHEMA_CORE_PROCESS.columns.COMMUNITY_COMMENT],
    updatedDate: data[SCHEMA_CORE_PROCESS.columns.UPDATED_DATE]
  };
}
