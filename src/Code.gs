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

function apiGetNinetyDayPlan(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var plan = getNinetyDayPlanByProjectId(projectId);
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

// ========================================
// API: Demo Data Management
// ========================================

/**
 * Generate demo data for a project to help users understand the app.
 */
function apiGenerateDemoData(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    generateDemoDataForProject(projectId, userEmail);
    return { success: true, message: 'デモデータを生成しました。' };
  } catch (error) {
    Logger.log('apiGenerateDemoData error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Clear all data for a project (reset to empty state).
 */
function apiClearProjectData(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    clearAllProjectData(projectId, userEmail);
    return { success: true, message: 'プロジェクトデータをクリアしました。' };
  } catch (error) {
    Logger.log('apiClearProjectData error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Generate demo data for all modules of a project.
 */
function generateDemoDataForProject(projectId, userEmail) {
  // 1. Vision demo data
  var visionData = {
    projectId: projectId,
    visionText: '2026年度末までに、全営業担当者がデータドリブンな意思決定を行い、商談成約率を現状比30%向上させる。顧客インサイトに基づく提案により、顧客満足度業界トップを実現する。',
    decisionRules: '1. 短期売上より長期顧客価値を優先する\n2. 主観的判断よりデータに基づく判断を優先する\n3. 個人の経験より組織の知見を活用する\n4. 迷ったときは顧客視点で判断する',
    successMetrics: '• 商談成約率: 現状25% → 目標32.5%（30%向上）\n• 営業生産性: 1人あたり月間商談数 15件 → 20件\n• 顧客満足度NPS: 現状+20 → 目標+40\n• データ活用率: ダッシュボード週次アクセス率 80%以上',
    notes: '経営会議で承認済み（2025年12月）。四半期ごとにKPIレビューを実施。'
  };
  saveVisionWithValidation(visionData, userEmail);

  // 2. Use cases demo data
  var usecasesDemo = [
    {
      projectId: projectId,
      challenge: '営業担当者が商談前に顧客情報を収集するのに平均2時間かかっている。複数システムへのアクセスが必要で、情報の鮮度も担保できない。',
      goal: '商談準備時間を2時間から30分に短縮',
      expectedImpact: '年間効果: ¥18,000,000（対象50人 × 90分/日削減 × 時給3,000円 × 20日/月 × 12ヶ月）\n営業活動時間の増加により、月間商談数が1.5倍に増加見込み',
      ninetyDayGoal: '営業ダッシュボード v1.0 を10名のパイロットユーザーに展開し、商談準備時間の削減効果を検証する',
      score: 85,
      priority: 1
    },
    {
      projectId: projectId,
      challenge: '月次の売上予測が営業マネージャーの勘と経験に依存しており、予測精度が±30%と低い。四半期末の追い込みが常態化している。',
      goal: '売上予測精度を±30%から±10%に改善',
      expectedImpact: '年間効果: ¥45,000,000（在庫最適化による機会損失削減、リソース配置の最適化）\n経営判断の迅速化により、市場変化への対応力向上',
      ninetyDayGoal: 'AIベースの売上予測モデルを構築し、過去データで予測精度を検証。±15%以内を達成する',
      score: 78,
      priority: 2
    },
    {
      projectId: projectId,
      challenge: '顧客の解約リスクを事前に把握できず、解約通知を受けてから対応している。解約率が年間15%と高止まり。',
      goal: '解約率を15%から10%に削減',
      expectedImpact: '年間効果: ¥30,000,000（顧客維持による売上確保、新規獲得コスト削減）\n顧客生涯価値（LTV）の向上',
      ninetyDayGoal: '解約リスクスコアリングモデルを構築し、ハイリスク顧客の早期フォロー体制を確立する',
      score: 72,
      priority: 3
    }
  ];

  usecasesDemo.forEach(function(uc) {
    addUsecaseWithValidation(uc, userEmail);
  });

  // 3. 90-Day Plan demo data
  var ninetyDayPlanData = {
    projectId: projectId,
    usecaseId: null,
    teamStructure: JSON.stringify({
      roles: [
        { pillar: 'CoE', title: 'プロジェクトオーナー', name: '田中 太郎（営業本部長）' },
        { pillar: 'CoE', title: 'データスチュワード', name: '佐藤 花子（経営企画）' },
        { pillar: 'Biz', title: 'ビジネスアナリスト', name: '鈴木 一郎（営業企画）' },
        { pillar: 'Biz', title: 'パワーユーザー', name: '高橋 美咲（営業1課）' },
        { pillar: 'IT', title: 'データエンジニア', name: '山田 健太（IT部）' },
        { pillar: 'IT', title: 'インフラ担当', name: '伊藤 誠（IT部）' }
      ]
    }),
    requiredData: JSON.stringify({
      items: [
        { name: '顧客マスタ', source: 'CRM（Salesforce）', status: 'ready' },
        { name: '商談履歴', source: 'CRM（Salesforce）', status: 'ready' },
        { name: '売上実績', source: 'ERP（SAP）', status: 'partial' },
        { name: '顧客行動ログ', source: 'Webサイト・メール', status: 'not_ready' }
      ]
    }),
    risks: JSON.stringify({
      items: [
        { description: 'データ品質が不十分で、ダッシュボードの信頼性が低下する可能性', impact: 'high', mitigation: '初期フェーズでデータプロファイリングを実施し、品質基準を設定' },
        { description: '営業現場の抵抗により、ダッシュボード利用が定着しない可能性', impact: 'medium', mitigation: 'パイロットユーザーを巻き込んだ設計、成功事例の横展開' },
        { description: 'IT部門のリソース不足により、開発が遅延する可能性', impact: 'medium', mitigation: '外部パートナーの活用を検討、スコープの優先順位付け' }
      ]
    }),
    communicationPlan: JSON.stringify({
      meetingFrequency: 'weekly',
      stakeholders: '営業本部長、IT部長、経営企画部長',
      reporting: '経営会議（月次）、部門長会議（隔週）'
    }),
    weeklyMilestones: JSON.stringify({
      phases: {
        ignite: {
          objective: '環境構築と初期ダッシュボード作成。パイロットユーザー10名での検証開始。',
          milestones: [
            'Week1: キックオフ、要件ヒアリング完了',
            'Week2: データ接続、品質検証',
            'Week3: ダッシュボード v0.5 レビュー',
            'Week4: パイロット開始、フィードバック収集'
          ],
          successCriteria: 'パイロットユーザー全員がダッシュボードにアクセスし、1回以上の商談準備に活用'
        },
        strengthen: {
          objective: 'フィードバックを反映した改善と、ユーザー拡大（10名→30名）',
          milestones: [
            'Week5: v1.0 リリース、機能改善',
            'Week6: 追加ユーザー向けトレーニング',
            'Week7: 利用状況モニタリング開始',
            'Week8: 中間レビュー、効果測定'
          ],
          successCriteria: '30名のユーザーが週3回以上アクセス、商談準備時間50%削減を達成'
        },
        establish: {
          objective: '全営業部門への展開と運用体制の確立',
          milestones: [
            'Week9: 全社展開計画策定',
            'Week10: 残りユーザーへの展開',
            'Week11: 運用マニュアル・FAQ整備',
            'Week12: 効果測定、次期計画策定'
          ],
          successCriteria: '全営業担当者（50名）が利用開始、商談準備時間75%削減を達成'
        }
      }
    })
  };
  saveNinetyDayPlanWithValidation(ninetyDayPlanData, userEmail);

  // 4. RACI demo data
  var raciEntries = [
    { pillar: 'CoE', task: 'プロジェクト全体統括', assignee: '田中 太郎', raci: 'A' },
    { pillar: 'CoE', task: 'データガバナンス策定', assignee: '佐藤 花子', raci: 'R' },
    { pillar: 'CoE', task: '経営層への報告', assignee: '田中 太郎', raci: 'R' },
    { pillar: 'Biz Data Network', task: '要件定義・優先順位付け', assignee: '鈴木 一郎', raci: 'R' },
    { pillar: 'Biz Data Network', task: 'ユーザートレーニング', assignee: '高橋 美咲', raci: 'R' },
    { pillar: 'Biz Data Network', task: '現場フィードバック収集', assignee: '高橋 美咲', raci: 'R' },
    { pillar: 'IT', task: 'データパイプライン構築', assignee: '山田 健太', raci: 'R' },
    { pillar: 'IT', task: 'ダッシュボード開発', assignee: '山田 健太', raci: 'R' },
    { pillar: 'IT', task: 'インフラ・セキュリティ', assignee: '伊藤 誠', raci: 'R' }
  ];
  saveRACIEntriesWithValidation(projectId, raciEntries, userEmail);

  // 5. Value tracking demo data
  var usecases = getUsecases(projectId);
  if (usecases.length > 0) {
    var valueData = {
      projectId: projectId,
      usecaseId: usecases[0].usecaseId,
      quantitativeImpact: '商談準備時間: 2時間 → 35分（71%削減）\n月間商談数: 15件 → 22件（47%増加）\n初月コスト削減: ¥1,500,000',
      qualitativeImpact: '• 営業担当者のモチベーション向上\n• 顧客との対話品質向上（データに基づく提案）\n• マネージャーの状況把握が容易に',
      evidence: '',
      nextInvestment: 'Expand'
    };
    saveValueWithValidation(valueData, userEmail);
  }

  logAudit(userEmail, OPERATION_TYPES.CREATE, projectId, {
    action: 'generate_demo_data'
  });
}

/**
 * Clear all data for a project.
 */
function clearAllProjectData(projectId, userEmail) {
  var ss = getSpreadsheet();

  // Clear Vision
  clearSheetRowsByProjectId(ss, SHEET_NAMES.VISION, 0, projectId);

  // Clear Use Cases
  clearSheetRowsByProjectId(ss, SHEET_NAMES.USECASES, 0, projectId);

  // Clear 90-Day Plan
  clearSheetRowsByProjectId(ss, SHEET_NAMES.NINETY_DAY_PLAN, 0, projectId);

  // Clear RACI
  clearSheetRowsByProjectId(ss, SHEET_NAMES.ORGANIZATION_RACI, 0, projectId);

  // Clear Governance
  clearSheetRowsByProjectId(ss, SHEET_NAMES.GOVERNANCE, 0, projectId);

  // Clear Operations Support
  clearSheetRowsByProjectId(ss, SHEET_NAMES.OPERATIONS_SUPPORT, 0, projectId);

  // Clear Value Tracking
  clearSheetRowsByProjectId(ss, SHEET_NAMES.VALUE, 0, projectId);

  // Clear Core Process
  clearSheetRowsByProjectId(ss, SHEET_NAMES.CORE_PROCESS, 0, projectId);

  logAudit(userEmail, OPERATION_TYPES.DELETE, projectId, {
    action: 'clear_all_project_data'
  });
}

/**
 * Helper to clear rows by project ID.
 */
function clearSheetRowsByProjectId(ss, sheetName, projectIdColumn, projectId) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];

  // Start from 1 to skip header
  for (var i = 1; i < data.length; i++) {
    if (data[i][projectIdColumn] === projectId) {
      rowsToDelete.push(i + 1); // 1-indexed
    }
  }

  // Delete from bottom to top to avoid index shifting
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }
}

// ========================================
// API: Maturity Level
// ========================================

/**
 * Get maturity level for a project based on core process assessment.
 */
function apiGetMaturityLevel(projectId) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var coreProcess = getCoreProcess(projectId);
    var maturity = calculateMaturityLevel(coreProcess);
    var templates = get90DayPlanTemplates(maturity.level);

    return {
      success: true,
      data: {
        maturity: maturity,
        templates: templates
      }
    };
  } catch (error) {
    Logger.log('apiGetMaturityLevel error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Get 90-day plan templates for a specific maturity level.
 */
function apiGet90DayPlanTemplates(maturityLevel) {
  try {
    var templates = get90DayPlanTemplates(maturityLevel || 'ignite');
    return { success: true, data: templates };
  } catch (error) {
    Logger.log('apiGet90DayPlanTemplates error: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * Apply maturity-based template to 90-day plan.
 */
function apiApplyMaturityTemplate(projectId, maturityLevel) {
  try {
    var userEmail = getCurrentUserEmail();
    requireProjectAccess(projectId, userEmail);

    var templates = get90DayPlanTemplates(maturityLevel || 'ignite');

    // Get existing plan or create new one
    var existingPlan = getNinetyDayPlanByProjectId(projectId);

    var planData = {
      projectId: projectId,
      usecaseId: existingPlan ? existingPlan.usecaseId : null,
      teamStructure: existingPlan ? existingPlan.teamStructure : '{}',
      requiredData: existingPlan ? existingPlan.requiredData : '{}',
      risks: existingPlan ? existingPlan.risks : '{}',
      communicationPlan: existingPlan ? existingPlan.communicationPlan : '{}',
      weeklyMilestones: JSON.stringify(templates)
    };

    saveNinetyDayPlanWithValidation(planData, userEmail);

    return {
      success: true,
      message: MATURITY_LABELS[maturityLevel] + 'のテンプレートを適用しました。',
      data: templates
    };
  } catch (error) {
    Logger.log('apiApplyMaturityTemplate error: ' + error.message);
    return { success: false, error: error.message };
  }
}
