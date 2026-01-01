/**
 * DocumentService.gs
 * Google Docs generation helpers.
 */

function generateProposalDocument(projectId, userEmail) {
  var project = getProject(projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません。');
  }

  var vision = getVision(projectId);
  var usecases = getUsecases(projectId);
  var raciEntries = getRACIEntries(projectId);
  var values = getValues(projectId);

  var docName = '[提案書ドラフト] ' + project.customerName + ' - Tableau Blueprint';
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  body.appendParagraph(docName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('作成日: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  body.appendParagraph('');

  appendSection(body, '1. 背景と目的', DocumentApp.ParagraphHeading.HEADING2);
  if (vision && vision.visionText) {
    body.appendParagraph(vision.visionText);
  } else {
    body.appendParagraph('ビジョンがまだ設定されていません。');
  }
  body.appendParagraph('');

  appendSection(body, '2. 範囲（ユースケース）', DocumentApp.ParagraphHeading.HEADING2);
  if (usecases && usecases.length > 0) {
    usecases.sort(function(a, b) {
      return (a.priority || 999) - (b.priority || 999);
    });

    usecases.slice(0, 3).forEach(function(uc, index) {
      body.appendParagraph((index + 1) + '. ' + uc.goal).setBold(true);
      body.appendParagraph('課題: ' + uc.challenge);
      body.appendParagraph('期待効果: ' + uc.expectedImpact);
      body.appendParagraph('90日目標: ' + uc.ninetyDayGoal);
      body.appendParagraph('');
    });
  } else {
    body.appendParagraph('ユースケースがまだありません。');
    body.appendParagraph('');
  }

  appendSection(body, '3. 組織 & RACI', DocumentApp.ParagraphHeading.HEADING2);
  if (raciEntries && raciEntries.length > 0) {
    var pillars = {};
    raciEntries.forEach(function(entry) {
      if (!pillars[entry.pillar]) {
        pillars[entry.pillar] = [];
      }
      pillars[entry.pillar].push(entry);
    });

    [THREE_PILLARS.COE, THREE_PILLARS.BUSINESS_DATA_NETWORK, THREE_PILLARS.IT].forEach(function(pillar) {
      if (pillars[pillar]) {
        body.appendParagraph('[' + pillar + ']').setBold(true);
        pillars[pillar].forEach(function(entry) {
          body.appendParagraph('  - ' + entry.task + ': ' + entry.assignee + ' (' + entry.raci + ')');
        });
        body.appendParagraph('');
      }
    });
  } else {
    body.appendParagraph('RACIがまだ設定されていません。');
    body.appendParagraph('');
  }

  appendSection(body, '4. 期待成果', DocumentApp.ParagraphHeading.HEADING2);
  if (vision && vision.successMetrics) {
    body.appendParagraph('成功指標:');
    body.appendParagraph(vision.successMetrics);
    body.appendParagraph('');
  }

  if (values && values.length > 0) {
    body.appendParagraph('価値のエビデンス:');
    values.forEach(function(val) {
      if (val.quantitativeImpact || val.qualitativeImpact) {
        body.appendParagraph('- 定量: ' + (val.quantitativeImpact || '未設定'));
        body.appendParagraph('  定性: ' + (val.qualitativeImpact || '未設定'));
        body.appendParagraph('');
      }
    });
  }

  if (!vision || (!vision.successMetrics && (!values || values.length === 0))) {
    body.appendParagraph('期待成果がまだ設定されていません。');
    body.appendParagraph('');
  }

  appendSection(body, '5. リスクと対策', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('- データ品質のリスク: クレンジングやプロファイリングで対応。');
  body.appendParagraph('- 利用定着のリスク: トレーニングと支援を実施。');
  body.appendParagraph('- 統合のリスク: 早期PoCで検証。');
  body.appendParagraph('');

  if (vision && vision.decisionRules) {
    body.appendParagraph('意思決定ルール:');
    body.appendParagraph(vision.decisionRules);
    body.appendParagraph('');
  }

  appendSection(body, '6. 投資判断', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('初期フェーズ: ライセンス、イネーブルメント、人員の見積もり。');
  body.appendParagraph('次フェーズ: 90日成果に基づき拡大可否を判断。');
  body.appendParagraph('');

  if (values && values.length > 0) {
    var hasInvestmentDecision = values.some(function(val) {
      return val.nextInvestment;
    });

    if (hasInvestmentDecision) {
      body.appendParagraph('投資判断:');
      var investmentLabels = {
        Continue: '継続',
        Expand: '拡大',
        Reduce: '縮小'
      };
      values.forEach(function(val) {
        if (val.nextInvestment) {
          body.appendParagraph('- ' + (investmentLabels[val.nextInvestment] || val.nextInvestment));
        }
      });
      body.appendParagraph('');
    }
  }

  appendSection(body, '7. スケジュール', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('週1-2: キックオフ、ビジョン、データアクセス。');
  body.appendParagraph('週3-4: 初期ダッシュボード。');
  body.appendParagraph('週5-8: レビューと定着。');
  body.appendParagraph('週9-12: 価値測定と次フェーズ計画。');
  body.appendParagraph('');

  appendSection(body, '8. 承認', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('起案者: ___________________   日付: ____/____/____');
  body.appendParagraph('');
  body.appendParagraph('承認者:  ___________________   日付: ____/____/____');
  body.appendParagraph('');

  doc.saveAndClose();

  logAudit(userEmail, OPERATION_TYPES.CREATE, projectId, {
    action: 'generate_proposal',
    documentId: doc.getId()
  });

  return doc.getUrl();
}

function appendSection(body, title, heading) {
  body.appendParagraph(title).setHeading(heading);
}

function generateVisionDocument(projectId, userEmail) {
  var project = getProject(projectId);
  var vision = getVision(projectId);

  if (!vision) {
    throw new Error('ビジョンが設定されていません。');
  }

  var docName = '[ビジョンシート] ' + project.customerName;
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  body.appendParagraph('ビジョン').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(project.customerName).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('');

  body.appendParagraph('ビジョンステートメント').setBold(true);
  body.appendParagraph(vision.visionText || '');
  body.appendParagraph('');

  body.appendParagraph('意思決定ルール').setBold(true);
  body.appendParagraph(vision.decisionRules || '');
  body.appendParagraph('');

  body.appendParagraph('成功指標').setBold(true);
  body.appendParagraph(vision.successMetrics || '');
  body.appendParagraph('');

  if (vision.notes) {
    body.appendParagraph('メモ').setBold(true);
    body.appendParagraph(vision.notes);
  }

  doc.saveAndClose();

  logAudit(userEmail, OPERATION_TYPES.CREATE, projectId, {
    action: 'generate_vision_doc',
    documentId: doc.getId()
  });

  return doc.getUrl();
}

function generateNinetyDayPlanDocument(projectId, usecaseId, userEmail) {
  var project = getProject(projectId);
  var usecases = getUsecases(projectId);
  var usecase = usecases.find(function(uc) { return uc.usecaseId === usecaseId; });
  var plan = getNinetyDayPlan(usecaseId, projectId);

  if (!plan) {
    throw new Error('90日計画が設定されていません。');
  }

  var docName = '90日計画 - ' + project.customerName + ' - ' + (usecase ? usecase.goal : usecaseId);
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  body.appendParagraph('90日計画').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(project.customerName).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('');

  if (usecase) {
    body.appendParagraph('ユースケース').setBold(true);
    body.appendParagraph('目的: ' + usecase.goal);
    body.appendParagraph('課題: ' + usecase.challenge);
    body.appendParagraph('90日目標: ' + usecase.ninetyDayGoal);
    body.appendParagraph('');
  }

  body.appendParagraph('体制').setBold(true);
  body.appendParagraph(plan.teamStructure || '');
  body.appendParagraph('');

  body.appendParagraph('必要データ').setBold(true);
  body.appendParagraph(plan.requiredData || '');
  body.appendParagraph('');

  body.appendParagraph('リスク').setBold(true);
  body.appendParagraph(plan.risks || '');
  body.appendParagraph('');

  body.appendParagraph('コミュニケーション計画').setBold(true);
  body.appendParagraph(plan.communicationPlan || '');
  body.appendParagraph('');

  body.appendParagraph('週次マイルストーン').setBold(true);

  if (plan.weeklyMilestones && plan.weeklyMilestones.length > 0) {
    plan.weeklyMilestones.forEach(function(milestone, index) {
      if (!milestone) return;
      body.appendParagraph('週' + (index + 1) + ': ' + milestone);
    });
  } else {
    body.appendParagraph('未設定');
  }

  doc.saveAndClose();

  logAudit(userEmail, OPERATION_TYPES.CREATE, projectId, {
    action: 'generate_90day_plan_doc',
    documentId: doc.getId(),
    usecaseId: usecaseId
  });

  return doc.getUrl();
}
