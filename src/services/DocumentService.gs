/**
 * DocumentService.gs
 * Google Docs ドキュメント自動生成サービス
 */

/**
 * 稟議書案を生成
 * @param {string} projectId - プロジェクトID
 * @param {string} userEmail - ユーザーメール
 * @return {string} 生成されたドキュメントのURL
 */
function generateProposalDocument(projectId, userEmail) {
  // プロジェクト情報を取得
  var project = getProject(projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません');
  }

  // 各モジュールのデータを取得
  var vision = getVision(projectId);
  var usecases = getUsecases(projectId);
  var raciEntries = getRACIEntries(projectId);
  var values = getValues(projectId);

  // ドキュメントを作成
  var docName = '【稟議書案】' + project.customerName + ' - Tableau Blueprint 導入提案';
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  // タイトル
  body.appendParagraph(docName)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph('作成日: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy年MM月dd日'));
  body.appendParagraph('');

  // 1. 背景・目的
  appendSection(body, '1. 背景・目的', DocumentApp.ParagraphHeading.HEADING2);
  if (vision && vision.visionText) {
    body.appendParagraph(vision.visionText);
  } else {
    body.appendParagraph('（ビジョンが未設定です）');
  }
  body.appendParagraph('');

  // 2. スコープ（ユースケース）
  appendSection(body, '2. スコープ（戦略的ユースケース）', DocumentApp.ParagraphHeading.HEADING2);
  if (usecases && usecases.length > 0) {
    // 優先度順にソート
    usecases.sort(function(a, b) {
      return (a.priority || 999) - (b.priority || 999);
    });

    // 上位3件まで
    var topUsecases = usecases.slice(0, 3);
    topUsecases.forEach(function(uc, index) {
      body.appendParagraph((index + 1) + '. ' + uc.goal)
        .setBold(true);
      body.appendParagraph('課題: ' + uc.challenge);
      body.appendParagraph('想定効果: ' + uc.expectedImpact);
      body.appendParagraph('90日ゴール: ' + uc.ninetyDayGoal);
      body.appendParagraph('');
    });
  } else {
    body.appendParagraph('（ユースケースが未設定です）');
    body.appendParagraph('');
  }

  // 3. 体制
  appendSection(body, '3. 推進体制', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Tableau Blueprintの3本柱組織モデルに基づく体制:');
  body.appendParagraph('');

  if (raciEntries && raciEntries.length > 0) {
    // 3本柱ごとにグループ化
    var pillars = {};
    raciEntries.forEach(function(entry) {
      if (!pillars[entry.pillar]) {
        pillars[entry.pillar] = [];
      }
      pillars[entry.pillar].push(entry);
    });

    // CoE
    if (pillars[THREE_PILLARS.COE]) {
      body.appendParagraph('【CoE (Center of Excellence)】')
        .setBold(true);
      pillars[THREE_PILLARS.COE].forEach(function(entry) {
        body.appendParagraph('  • ' + entry.task + ': ' + entry.assignee + ' (' + entry.raci + ')');
      });
      body.appendParagraph('');
    }

    // ビジネスデータネットワーク
    if (pillars[THREE_PILLARS.BUSINESS_DATA_NETWORK]) {
      body.appendParagraph('【ビジネスデータネットワーク】')
        .setBold(true);
      pillars[THREE_PILLARS.BUSINESS_DATA_NETWORK].forEach(function(entry) {
        body.appendParagraph('  • ' + entry.task + ': ' + entry.assignee + ' (' + entry.raci + ')');
      });
      body.appendParagraph('');
    }

    // IT
    if (pillars[THREE_PILLARS.IT]) {
      body.appendParagraph('【IT】')
        .setBold(true);
      pillars[THREE_PILLARS.IT].forEach(function(entry) {
        body.appendParagraph('  • ' + entry.task + ': ' + entry.assignee + ' (' + entry.raci + ')');
      });
      body.appendParagraph('');
    }
  } else {
    body.appendParagraph('（体制が未設定です）');
    body.appendParagraph('');
  }

  // 4. 期待効果
  appendSection(body, '4. 期待効果', DocumentApp.ParagraphHeading.HEADING2);

  if (vision && vision.successMetrics) {
    body.appendParagraph('【成功指標】');
    body.appendParagraph(vision.successMetrics);
    body.appendParagraph('');
  }

  if (values && values.length > 0) {
    body.appendParagraph('【実現価値】');
    values.forEach(function(val) {
      if (val.quantitativeImpact || val.qualitativeImpact) {
        body.appendParagraph('• 定量効果: ' + (val.quantitativeImpact || '未設定'));
        body.appendParagraph('  定性効果: ' + (val.qualitativeImpact || '未設定'));
        body.appendParagraph('');
      }
    });
  }

  if (!vision || (!vision.successMetrics && (!values || values.length === 0))) {
    body.appendParagraph('（期待効果が未設定です）');
    body.appendParagraph('');
  }

  // 5. リスクと対策
  appendSection(body, '5. リスクと対策', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('【想定リスク】');
  body.appendParagraph('• データ品質の課題: データクレンジングと品質向上プロセスを並行実施');
  body.appendParagraph('• ユーザー定着の遅れ: 階層型サポート体制とトレーニング計画で対応');
  body.appendParagraph('• 技術的な統合課題: 段階的な導入とプロトタイプ検証で早期発見');
  body.appendParagraph('');

  if (vision && vision.decisionRules) {
    body.appendParagraph('【意思決定ルール】');
    body.appendParagraph(vision.decisionRules);
    body.appendParagraph('');
  }

  // 6. 投資判断（概算）
  appendSection(body, '6. 投資判断', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('【始動フェーズ（1〜3ヶ月）】');
  body.appendParagraph('• Tableau ライセンス: Creator x名、Explorer x名');
  body.appendParagraph('• 導入支援: ワークショップ、トレーニング、環境構築支援');
  body.appendParagraph('• 推進体制: CoE 専任x名、兼任x名');
  body.appendParagraph('');
  body.appendParagraph('【強化フェーズ以降】');
  body.appendParagraph('• 90日計画の成果をもとに、継続/拡大/縮小を判断');
  body.appendParagraph('');

  if (values && values.length > 0) {
    var hasInvestmentDecision = values.some(function(val) {
      return val.nextInvestment;
    });

    if (hasInvestmentDecision) {
      body.appendParagraph('【次フェーズの投資判断材料】');
      values.forEach(function(val) {
        if (val.nextInvestment) {
          body.appendParagraph('• ' + val.nextInvestment);
        }
      });
      body.appendParagraph('');
    }
  }

  // 7. スケジュール
  appendSection(body, '7. スケジュール', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('【始動フェーズ（1〜3ヶ月）】');
  body.appendParagraph('Week 1-2: キックオフ、ビジョン策定、環境構築');
  body.appendParagraph('Week 3-4: データ接続、初期ダッシュボード開発');
  body.appendParagraph('Week 5-8: ユーザートレーニング、フィードバック反映');
  body.appendParagraph('Week 9-12: 成果測定、経営報告、次フェーズ計画');
  body.appendParagraph('');

  // 8. 承認欄
  appendSection(body, '8. 承認', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('起案者: ___________________   日付: ___/___/___');
  body.appendParagraph('');
  body.appendParagraph('承認者: ___________________   日付: ___/___/___');
  body.appendParagraph('');

  // ドキュメントを保存
  doc.saveAndClose();

  // 監査ログ
  logAudit(userEmail, OPERATION_TYPES.CREATE, projectId, {
    action: 'generate_proposal',
    documentId: doc.getId()
  });

  return doc.getUrl();
}

/**
 * セクション見出しを追加
 * @param {Body} body - ドキュメント本文
 * @param {string} title - セクションタイトル
 * @param {DocumentApp.ParagraphHeading} heading - 見出しレベル
 */
function appendSection(body, title, heading) {
  body.appendParagraph(title)
    .setHeading(heading);
}

/**
 * ビジョン1枚シートを生成
 * @param {string} projectId - プロジェクトID
 * @param {string} userEmail - ユーザーメール
 * @return {string} 生成されたドキュメントのURL
 */
function generateVisionDocument(projectId, userEmail) {
  var project = getProject(projectId);
  var vision = getVision(projectId);

  if (!vision) {
    throw new Error('ビジョンが未設定です');
  }

  var docName = '【ビジョン1枚】' + project.customerName;
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  // タイトル
  body.appendParagraph('北極星ビジョン')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph(project.customerName)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph('');

  // ビジョン本文
  body.appendParagraph('【ビジョン】')
    .setBold(true);
  body.appendParagraph(vision.visionText || '');
  body.appendParagraph('');

  // 意思決定ルール
  body.appendParagraph('【意思決定ルール】')
    .setBold(true);
  body.appendParagraph(vision.decisionRules || '');
  body.appendParagraph('');

  // 成功指標
  body.appendParagraph('【成功指標】')
    .setBold(true);
  body.appendParagraph(vision.successMetrics || '');
  body.appendParagraph('');

  // 備考
  if (vision.notes) {
    body.appendParagraph('【備考】')
      .setBold(true);
    body.appendParagraph(vision.notes);
  }

  doc.saveAndClose();

  logAudit(userEmail, OPERATION_TYPES.CREATE, projectId, {
    action: 'generate_vision_doc',
    documentId: doc.getId()
  });

  return doc.getUrl();
}

/**
 * 90日計画ドキュメントを生成
 * @param {string} projectId - プロジェクトID
 * @param {string} usecaseId - ユースケースID
 * @param {string} userEmail - ユーザーメール
 * @return {string} 生成されたドキュメントのURL
 */
function generateNinetyDayPlanDocument(projectId, usecaseId, userEmail) {
  var project = getProject(projectId);
  var usecases = getUsecases(projectId);
  var usecase = usecases.find(function(uc) { return uc.usecaseId === usecaseId; });
  var plan = getNinetyDayPlan(usecaseId);

  if (!plan) {
    throw new Error('90日計画が未設定です');
  }

  var docName = '【90日計画】' + project.customerName + ' - ' + (usecase ? usecase.goal : usecaseId);
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  // タイトル
  body.appendParagraph('90日計画（始動フェーズ）')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph(project.customerName)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph('');

  // ユースケース概要
  if (usecase) {
    body.appendParagraph('【ユースケース】')
      .setBold(true);
    body.appendParagraph('狙い: ' + usecase.goal);
    body.appendParagraph('課題: ' + usecase.challenge);
    body.appendParagraph('90日ゴール: ' + usecase.ninetyDayGoal);
    body.appendParagraph('');
  }

  // 体制
  body.appendParagraph('【体制】')
    .setBold(true);
  body.appendParagraph(plan.teamStructure || '');
  body.appendParagraph('');

  // 必要データ
  body.appendParagraph('【必要データ】')
    .setBold(true);
  body.appendParagraph(plan.requiredData || '');
  body.appendParagraph('');

  // リスク
  body.appendParagraph('【リスク】')
    .setBold(true);
  body.appendParagraph(plan.risks || '');
  body.appendParagraph('');

  // コミュニケーション計画
  body.appendParagraph('【コミュニケーション計画】')
    .setBold(true);
  body.appendParagraph(plan.communicationPlan || '');
  body.appendParagraph('');

  // 週次マイルストーン
  body.appendParagraph('【週次マイルストーン】')
    .setBold(true);

  if (plan.weeklyMilestones && plan.weeklyMilestones.length > 0) {
    plan.weeklyMilestones.forEach(function(milestone, index) {
      body.appendParagraph('Week ' + (index + 1) + ': ' + milestone);
    });
  } else {
    body.appendParagraph('（未設定）');
  }

  doc.saveAndClose();

  logAudit(userEmail, OPERATION_TYPES.CREATE, projectId, {
    action: 'generate_90day_plan_doc',
    documentId: doc.getId(),
    usecaseId: usecaseId
  });

  return doc.getUrl();
}
