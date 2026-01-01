/**
 * Code.gs
 * Web App entry point and API handlers.
 */

/**
 * Web App GET handler.
 * @param {Object} e
 * @return {HtmlOutput}
 */
function doGet(e) {
  var userEmail = getCurrentUserEmail();
  if (!userEmail) {
    return HtmlService.createHtmlOutput(
      '認証が必要です。Googleアカウントでログインしてください。'
    );
  }

  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = userEmail;

  return template.evaluate()
    .setTitle('Tableau Blueprint ワークショップ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include helper for HTML templates.
 * @param {string} filename
 * @return {string}
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ========================================
// API: Projects
// ========================================

function apiCreateProject(customerName) {
  try {
    var userEmail = getCurrentUserEmail();
    var project = createProjectWithValidation(customerName, userEmail);
    return { success: true, data: project };
  } catch (error) {
    Logger.log('apiCreateProject error: ' + error.message);
    return { success: false, error: error.message };
  }
}

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

function apiGetProject(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var project = getProject(projectId);
    if (project) {
      project.creatorEmail = project.createdBy;
    }
    return { success: true, data: project };
  } catch (error) {
    Logger.log('apiGetProject error: ' + error.message);
    return { success: false, error: error.message };
  }
}

function apiUpdateProjectStatus(projectId, status) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    updateProjectStatusWithValidation(projectId, status, userEmail);
    return { success: true, message: 'プロジェクトのステータスを更新しました。' };
  } catch (error) {
    Logger.log('apiUpdateProjectStatus error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========================================
// API: Module 1 - Vision
// ========================================

function apiSaveVision(visionData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(visionData.projectId, userEmail);

    saveVisionWithValidation(visionData, userEmail);

    return { success: true, message: 'ビジョンを保存しました。' };
  } catch (error) {
    Logger.log('apiSaveVision error: ' + error.message);
    return { success: false, error: error.message };
  }
}

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
// API: Module 2 - Use Cases & 90-Day Plan
// ========================================

function apiAddUsecase(usecaseData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(usecaseData.projectId, userEmail);

    var usecaseId = addUsecaseWithValidation(usecaseData, userEmail);
    return { success: true, data: { usecaseId: usecaseId } };
  } catch (error) {
    Logger.log('apiAddUsecase error: ' + error.message);
    return { success: false, error: error.message };
  }
}

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

function apiSaveNinetyDayPlan(planData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(planData.projectId, userEmail);

    saveNinetyDayPlanWithValidation(planData, userEmail);
    return { success: true, message: '90日計画を保存しました。' };
  } catch (error) {
    Logger.log('apiSaveNinetyDayPlan error: ' + error.message);
    return { success: false, error: error.message };
  }
}

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
// API: Module 4 - Org & RACI
// ========================================

function apiSaveRACIEntries(projectId, raciEntries) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    saveRACIEntriesWithValidation(projectId, raciEntries, userEmail);
    return { success: true, message: 'RACIを保存しました。' };
  } catch (error) {
    Logger.log('apiSaveRACIEntries error: ' + error.message);
    return { success: false, error: error.message };
  }
}

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
// API: Module 7 - Value Tracking
// ========================================

function apiSaveValue(valueData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(valueData.projectId, userEmail);

    saveValueWithValidation(valueData, userEmail);
    return { success: true, message: '価値トラッキングを保存しました。' };
  } catch (error) {
    Logger.log('apiSaveValue error: ' + error.message);
    return { success: false, error: error.message };
  }
}

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
// API: Document generation
// ========================================

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
// API: Module 3 - Core Process Gap Analysis
// ========================================

function apiSaveCoreProcess(processData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(processData.projectId, userEmail);

    saveCoreProcess(processData, userEmail);
    return { success: true, message: 'コアプロセスデータを保存しました。' };
  } catch (error) {
    Logger.log('apiSaveCoreProcess error: ' + error.message);
    return { success: false, error: error.message };
  }
}

function apiGetCoreProcess(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var data = getCoreProcess(projectId);
    return { success: true, data: data };
  } catch (error) {
    Logger.log('apiGetCoreProcess error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========================================
// API: Module 5 - Governance
// ========================================

function apiSaveGovernanceEntries(projectId, governanceEntries) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    saveGovernanceEntries(projectId, governanceEntries, userEmail);
    return { success: true, message: 'ガバナンス項目を保存しました。' };
  } catch (error) {
    Logger.log('apiSaveGovernanceEntries error: ' + error.message);
    return { success: false, error: error.message };
  }
}

function apiGetGovernanceEntries(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var entries = getGovernanceEntries(projectId);
    return { success: true, data: entries };
  } catch (error) {
    Logger.log('apiGetGovernanceEntries error: ' + error.message);
    return { success: false, error: error.message };
  }
}

function apiAddGovernanceEntry(governanceData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(governanceData.projectId, userEmail);

    addGovernanceEntry(governanceData, userEmail);
    return { success: true, message: 'Governance entry added.' };
  } catch (error) {
    Logger.log('apiAddGovernanceEntry error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========================================
// API: Module 6 - Operations Support
// ========================================

function apiSaveOperationsSupport(supportData) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(supportData.projectId, userEmail);

    saveOperationsSupport(supportData, userEmail);
    return { success: true, message: '運用サポートデータを保存しました。' };
  } catch (error) {
    Logger.log('apiSaveOperationsSupport error: ' + error.message);
    return { success: false, error: error.message };
  }
}

function apiGetOperationsSupport(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var data = getOperationsSupport(projectId);
    return { success: true, data: data };
  } catch (error) {
    Logger.log('apiGetOperationsSupport error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// ========================================
// API: Utilities
// ========================================

function apiInitializeSheets() {
  try {
    var result = initializeAllSheets();
    return { success: true, message: result.message };
  } catch (error) {
    Logger.log('apiInitializeSheets error: ' + error.message);
    return { success: false, error: error.message };
  }
}

function apiGetCurrentUser() {
  try {
    var user = getCurrentUser();
    return { success: true, data: user };
  } catch (error) {
    Logger.log('apiGetCurrentUser error: ' + error.message);
    return { success: false, error: error.message };
  }
}
