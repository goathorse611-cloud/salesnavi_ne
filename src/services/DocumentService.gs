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
  var ninetyDayPlan = getNinetyDayPlanByProjectId(projectId);

  var docName = '[稟議書案] ' + project.customerName + ' データ分析基盤導入';
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  // Title and metadata
  body.appendParagraph(docName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('');

  var metaTable = body.appendTable();
  var metaRow1 = metaTable.appendTableRow();
  metaRow1.appendTableCell('作成日').setBackgroundColor('#f3f4f6');
  metaRow1.appendTableCell(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy年MM月dd日'));
  var metaRow2 = metaTable.appendTableRow();
  metaRow2.appendTableCell('顧客名').setBackgroundColor('#f3f4f6');
  metaRow2.appendTableCell(project.customerName);
  var metaRow3 = metaTable.appendTableRow();
  metaRow3.appendTableCell('担当者').setBackgroundColor('#f3f4f6');
  metaRow3.appendTableCell(userEmail);
  body.appendParagraph('');

  // ========================================
  // Executive Summary
  // ========================================
  appendSection(body, 'エグゼクティブサマリー', DocumentApp.ParagraphHeading.HEADING2);

  var summaryText = '';
  if (vision && vision.visionText) {
    summaryText += '【目指す姿】' + vision.visionText.substring(0, 100) + '...\n\n';
  }
  if (usecases && usecases.length > 0) {
    var topUsecase = usecases.sort(function(a, b) { return (a.priority || 999) - (b.priority || 999); })[0];
    summaryText += '【最優先課題】' + (topUsecase.challenge || '').substring(0, 80) + '\n';
    summaryText += '【期待効果】' + (topUsecase.expectedImpact || '').substring(0, 80) + '\n\n';
  }
  summaryText += '本提案は、90日間の初期フェーズで成果を実証し、段階的に拡大することを提案します。';
  body.appendParagraph(summaryText);
  body.appendParagraph('');

  // ========================================
  // 1. Background - Current Challenges
  // ========================================
  appendSection(body, '1. 背景：現状の課題', DocumentApp.ParagraphHeading.HEADING2);

  if (usecases && usecases.length > 0) {
    body.appendParagraph('現在、以下の課題が顕在化しています：').setBold(true);
    body.appendParagraph('');

    usecases.slice(0, 3).forEach(function(uc, index) {
      body.appendParagraph('課題' + (index + 1) + ': ' + (uc.challenge || '未定義')).setBold(true);
      if (uc.goal) {
        body.appendParagraph('→ 解決により: ' + uc.goal);
      }
      body.appendParagraph('');
    });

    body.appendParagraph('これらの課題を放置した場合のリスク：').setItalic(true);
    body.appendParagraph('・ 競合他社との差別化が困難になる');
    body.appendParagraph('・ 意思決定の遅延による機会損失');
    body.appendParagraph('・ 属人化によるノウハウ喪失リスク');
  } else {
    body.appendParagraph('ユースケースがまだ定義されていません。');
  }
  body.appendParagraph('');

  // ========================================
  // 2. Vision and Goals
  // ========================================
  appendSection(body, '2. 目指す姿（ビジョン）', DocumentApp.ParagraphHeading.HEADING2);

  if (vision && vision.visionText) {
    body.appendParagraph(vision.visionText);
    body.appendParagraph('');

    if (vision.successMetrics) {
      body.appendParagraph('【成功指標（KPI）】').setBold(true);
      body.appendParagraph(vision.successMetrics);
      body.appendParagraph('');
    }

    if (vision.decisionRules) {
      body.appendParagraph('【意思決定の原則】').setBold(true);
      body.appendParagraph(vision.decisionRules);
    }
  } else {
    body.appendParagraph('ビジョンがまだ設定されていません。');
  }
  body.appendParagraph('');

  // ========================================
  // 3. Investment Effect (ROI)
  // ========================================
  appendSection(body, '3. 投資対効果', DocumentApp.ParagraphHeading.HEADING2);

  if (usecases && usecases.length > 0) {
    var totalEffect = 0;
    body.appendParagraph('【定量効果（年間想定）】').setBold(true);

    var effectTable = body.appendTable();
    var headerRow = effectTable.appendTableRow();
    headerRow.appendTableCell('ユースケース').setBackgroundColor('#0176D3').getChild(0).asParagraph().editAsText().setForegroundColor('#ffffff');
    headerRow.appendTableCell('想定効果').setBackgroundColor('#0176D3').getChild(0).asParagraph().editAsText().setForegroundColor('#ffffff');

    usecases.slice(0, 3).forEach(function(uc) {
      var row = effectTable.appendTableRow();
      row.appendTableCell(uc.goal || '未定義');
      row.appendTableCell(uc.expectedImpact || '未定義');
    });

    body.appendParagraph('');

    if (values && values.length > 0) {
      body.appendParagraph('【定性効果】').setBold(true);
      values.forEach(function(val) {
        if (val.qualitativeImpact) {
          body.appendParagraph(val.qualitativeImpact);
        }
      });
      body.appendParagraph('');
    }

    body.appendParagraph('【投資判断】').setBold(true);
    body.appendParagraph('・ 初期投資：90日間のPoC（概念実証）として最小限のリソースで開始');
    body.appendParagraph('・ 拡大判断：90日後の成果に基づき、次フェーズの投資規模を決定');
    body.appendParagraph('・ 撤退基準：KPI達成率50%未満の場合は計画見直し');
  } else {
    body.appendParagraph('ユースケースが定義されていないため、効果試算ができません。');
  }
  body.appendParagraph('');

  // ========================================
  // 4. Implementation Plan (90 days)
  // ========================================
  appendSection(body, '4. 実行計画（90日）', DocumentApp.ParagraphHeading.HEADING2);

  if (ninetyDayPlan && ninetyDayPlan.weeklyMilestones) {
    var phasesData = parseJsonField(ninetyDayPlan.weeklyMilestones);
    if (phasesData && phasesData.phases) {
      var phaseInfo = [
        { id: 'ignite', name: '始動フェーズ', period: '1-4週目', color: '#f59e0b' },
        { id: 'strengthen', name: '推進フェーズ', period: '5-8週目', color: '#3b82f6' },
        { id: 'establish', name: '定着フェーズ', period: '9-12週目', color: '#10b981' }
      ];

      phaseInfo.forEach(function(info) {
        var phase = phasesData.phases[info.id];
        if (phase) {
          body.appendParagraph('【' + info.name + '（' + info.period + '）】').setBold(true);
          if (phase.objective) {
            body.appendParagraph('目標: ' + phase.objective);
          }
          if (phase.milestones && phase.milestones.length > 0) {
            phase.milestones.forEach(function(ms) {
              if (ms) body.appendParagraph('  ・ ' + ms);
            });
          }
          if (phase.successCriteria) {
            body.appendParagraph('成功基準: ' + phase.successCriteria).setItalic(true);
          }
          body.appendParagraph('');
        }
      });
    }
  } else {
    body.appendParagraph('90日計画がまだ設定されていません。');
    body.appendParagraph('');
  }

  // ========================================
  // 5. Team Structure
  // ========================================
  appendSection(body, '5. 推進体制', DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph('本プロジェクトは「3本柱」体制で推進します：').setBold(true);
  body.appendParagraph('');

  // 3 Pillars explanation
  var pillarTable = body.appendTable();
  var pillarHeader = pillarTable.appendTableRow();
  pillarHeader.appendTableCell('役割').setBackgroundColor('#0176D3').getChild(0).asParagraph().editAsText().setForegroundColor('#ffffff');
  pillarHeader.appendTableCell('責任範囲').setBackgroundColor('#0176D3').getChild(0).asParagraph().editAsText().setForegroundColor('#ffffff');
  pillarHeader.appendTableCell('担当タスク例').setBackgroundColor('#0176D3').getChild(0).asParagraph().editAsText().setForegroundColor('#ffffff');

  var coeRow = pillarTable.appendTableRow();
  coeRow.appendTableCell('CoE\n（全社推進）');
  coeRow.appendTableCell('標準化・ガバナンス・品質管理');

  var bizRow = pillarTable.appendTableRow();
  bizRow.appendTableCell('Biz Data Network\n（部門推進）');
  bizRow.appendTableCell('要件定義・現場定着・フィードバック');

  var itRow = pillarTable.appendTableRow();
  itRow.appendTableCell('IT\n（技術基盤）');
  itRow.appendTableCell('インフラ・セキュリティ・開発');

  // Fill in from RACI
  if (raciEntries && raciEntries.length > 0) {
    var pillars = {};
    raciEntries.forEach(function(entry) {
      if (!pillars[entry.pillar]) {
        pillars[entry.pillar] = [];
      }
      pillars[entry.pillar].push(entry.task + ': ' + entry.assignee);
    });

    if (pillars['CoE']) {
      coeRow.getCell(2).setText(pillars['CoE'].join('\n'));
    }
    if (pillars['Biz Data Network']) {
      bizRow.getCell(2).setText(pillars['Biz Data Network'].join('\n'));
    }
    if (pillars['IT']) {
      itRow.getCell(2).setText(pillars['IT'].join('\n'));
    }
  }

  body.appendParagraph('');

  // ========================================
  // 6. Risks and Mitigation
  // ========================================
  appendSection(body, '6. リスクと対策', DocumentApp.ParagraphHeading.HEADING2);

  if (ninetyDayPlan && ninetyDayPlan.risks) {
    var risksData = parseJsonField(ninetyDayPlan.risks);
    if (risksData && risksData.items && risksData.items.length > 0) {
      var riskTable = body.appendTable();
      var riskHeader = riskTable.appendTableRow();
      riskHeader.appendTableCell('リスク').setBackgroundColor('#dc2626').getChild(0).asParagraph().editAsText().setForegroundColor('#ffffff');
      riskHeader.appendTableCell('影響度').setBackgroundColor('#dc2626').getChild(0).asParagraph().editAsText().setForegroundColor('#ffffff');
      riskHeader.appendTableCell('対策').setBackgroundColor('#dc2626').getChild(0).asParagraph().editAsText().setForegroundColor('#ffffff');

      var impactLabels = { high: '高', medium: '中', low: '低' };
      risksData.items.forEach(function(risk) {
        var row = riskTable.appendTableRow();
        row.appendTableCell(risk.description || '');
        row.appendTableCell(impactLabels[risk.impact] || risk.impact || '');
        row.appendTableCell(risk.mitigation || '');
      });
    } else {
      body.appendParagraph('リスクがまだ設定されていません。');
    }
  } else {
    body.appendParagraph('【想定されるリスク】');
    body.appendParagraph('・ データ品質リスク: 初期フェーズでデータプロファイリングを実施');
    body.appendParagraph('・ 利用定着リスク: パイロットユーザーを巻き込んだ設計、トレーニング実施');
    body.appendParagraph('・ 技術リスク: 早期のPoC（概念実証）で検証');
  }
  body.appendParagraph('');

  // ========================================
  // 7. Next Steps
  // ========================================
  appendSection(body, '7. 次のステップ', DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph('本稟議承認後、以下のアクションを実施します：').setBold(true);
  body.appendParagraph('');
  body.appendParagraph('□ Week 0: キックオフミーティング開催');
  body.appendParagraph('□ Week 1: 詳細要件定義、データアクセス確認');
  body.appendParagraph('□ Week 2: 環境構築、初期開発開始');
  body.appendParagraph('□ Week 4: 初期レビュー、パイロット開始');
  body.appendParagraph('');

  // ========================================
  // 8. Approval Section
  // ========================================
  appendSection(body, '8. 承認欄', DocumentApp.ParagraphHeading.HEADING2);

  var approvalTable = body.appendTable();

  var approvalHeader = approvalTable.appendTableRow();
  approvalHeader.appendTableCell('').setBackgroundColor('#f3f4f6');
  approvalHeader.appendTableCell('氏名').setBackgroundColor('#f3f4f6');
  approvalHeader.appendTableCell('署名').setBackgroundColor('#f3f4f6');
  approvalHeader.appendTableCell('日付').setBackgroundColor('#f3f4f6');

  var proposerRow = approvalTable.appendTableRow();
  proposerRow.appendTableCell('起案者');
  proposerRow.appendTableCell('');
  proposerRow.appendTableCell('');
  proposerRow.appendTableCell('    /    /    ');

  var approver1Row = approvalTable.appendTableRow();
  approver1Row.appendTableCell('承認者1');
  approver1Row.appendTableCell('');
  approver1Row.appendTableCell('');
  approver1Row.appendTableCell('    /    /    ');

  var approver2Row = approvalTable.appendTableRow();
  approver2Row.appendTableCell('承認者2');
  approver2Row.appendTableCell('');
  approver2Row.appendTableCell('');
  approver2Row.appendTableCell('    /    /    ');

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

function parseJsonField(value) {
  if (!value) return null;
  if (typeof value === 'object') return value;
  try {
    return JSON.parse(value);
  } catch (e) {
    return null;
  }
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
  var plan = getNinetyDayPlan(usecaseId);

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
