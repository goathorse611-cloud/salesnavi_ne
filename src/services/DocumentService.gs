/**
 * DocumentService.gs
 * Google Docs generation helpers.
 */

function generateProposalDocument(projectId, userEmail) {
  var project = getProject(projectId);
  if (!project) {
    throw new Error('Project not found.');
  }

  var vision = getVision(projectId);
  var usecases = getUsecases(projectId);
  var raciEntries = getRACIEntries(projectId);
  var values = getValues(projectId);

  var docName = '[Proposal Draft] ' + project.customerName + ' - Tableau Blueprint';
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  body.appendParagraph(docName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Created: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  body.appendParagraph('');

  appendSection(body, '1. Background and Objectives', DocumentApp.ParagraphHeading.HEADING2);
  if (vision && vision.visionText) {
    body.appendParagraph(vision.visionText);
  } else {
    body.appendParagraph('Vision is not set yet.');
  }
  body.appendParagraph('');

  appendSection(body, '2. Scope (Use Cases)', DocumentApp.ParagraphHeading.HEADING2);
  if (usecases && usecases.length > 0) {
    usecases.sort(function(a, b) {
      return (a.priority || 999) - (b.priority || 999);
    });

    usecases.slice(0, 3).forEach(function(uc, index) {
      body.appendParagraph((index + 1) + '. ' + uc.goal).setBold(true);
      body.appendParagraph('Challenge: ' + uc.challenge);
      body.appendParagraph('Expected impact: ' + uc.expectedImpact);
      body.appendParagraph('90-day goal: ' + uc.ninetyDayGoal);
      body.appendParagraph('');
    });
  } else {
    body.appendParagraph('No use cases defined yet.');
    body.appendParagraph('');
  }

  appendSection(body, '3. Organization & RACI', DocumentApp.ParagraphHeading.HEADING2);
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
    body.appendParagraph('RACI is not set yet.');
    body.appendParagraph('');
  }

  appendSection(body, '4. Expected Outcomes', DocumentApp.ParagraphHeading.HEADING2);
  if (vision && vision.successMetrics) {
    body.appendParagraph('Success metrics:');
    body.appendParagraph(vision.successMetrics);
    body.appendParagraph('');
  }

  if (values && values.length > 0) {
    body.appendParagraph('Value evidence:');
    values.forEach(function(val) {
      if (val.quantitativeImpact || val.qualitativeImpact) {
        body.appendParagraph('- Quantitative: ' + (val.quantitativeImpact || 'Not set'));
        body.appendParagraph('  Qualitative: ' + (val.qualitativeImpact || 'Not set'));
        body.appendParagraph('');
      }
    });
  }

  if (!vision || (!vision.successMetrics && (!values || values.length === 0))) {
    body.appendParagraph('Expected outcomes are not set yet.');
    body.appendParagraph('');
  }

  appendSection(body, '5. Risks and Mitigations', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('- Data quality risks: address with cleansing and profiling.');
  body.appendParagraph('- Adoption risks: provide training and enablement.');
  body.appendParagraph('- Integration risks: validate with early PoC.');
  body.appendParagraph('');

  if (vision && vision.decisionRules) {
    body.appendParagraph('Decision rules:');
    body.appendParagraph(vision.decisionRules);
    body.appendParagraph('');
  }

  appendSection(body, '6. Investment Decision', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Initial phase: estimate licenses, enablement, and staffing.');
  body.appendParagraph('Next phase: decide expansion based on 90-day results.');
  body.appendParagraph('');

  if (values && values.length > 0) {
    var hasInvestmentDecision = values.some(function(val) {
      return val.nextInvestment;
    });

    if (hasInvestmentDecision) {
      body.appendParagraph('Investment decisions:');
      values.forEach(function(val) {
        if (val.nextInvestment) {
          body.appendParagraph('- ' + val.nextInvestment);
        }
      });
      body.appendParagraph('');
    }
  }

  appendSection(body, '7. Schedule', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Week 1-2: Kickoff, vision, data access.');
  body.appendParagraph('Week 3-4: Initial dashboard.');
  body.appendParagraph('Week 5-8: Review and adoption.');
  body.appendParagraph('Week 9-12: Measure value and plan next phase.');
  body.appendParagraph('');

  appendSection(body, '8. Approvals', DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Initiator: ___________________   Date: ____/____/____');
  body.appendParagraph('');
  body.appendParagraph('Approver:  ___________________   Date: ____/____/____');
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
    throw new Error('Vision is not set.');
  }

  var docName = '[Vision Sheet] ' + project.customerName;
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  body.appendParagraph('Vision').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(project.customerName).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('');

  body.appendParagraph('Vision statement').setBold(true);
  body.appendParagraph(vision.visionText || '');
  body.appendParagraph('');

  body.appendParagraph('Decision rules').setBold(true);
  body.appendParagraph(vision.decisionRules || '');
  body.appendParagraph('');

  body.appendParagraph('Success metrics').setBold(true);
  body.appendParagraph(vision.successMetrics || '');
  body.appendParagraph('');

  if (vision.notes) {
    body.appendParagraph('Notes').setBold(true);
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
    throw new Error('90-day plan is not set.');
  }

  var docName = '90-Day Plan - ' + project.customerName + ' - ' + (usecase ? usecase.goal : usecaseId);
  var doc = DocumentApp.create(docName);
  var body = doc.getBody();

  body.appendParagraph('90-Day Plan').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(project.customerName).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('');

  if (usecase) {
    body.appendParagraph('Use case').setBold(true);
    body.appendParagraph('Goal: ' + usecase.goal);
    body.appendParagraph('Challenge: ' + usecase.challenge);
    body.appendParagraph('90-day goal: ' + usecase.ninetyDayGoal);
    body.appendParagraph('');
  }

  body.appendParagraph('Team structure').setBold(true);
  body.appendParagraph(plan.teamStructure || '');
  body.appendParagraph('');

  body.appendParagraph('Required data').setBold(true);
  body.appendParagraph(plan.requiredData || '');
  body.appendParagraph('');

  body.appendParagraph('Risks').setBold(true);
  body.appendParagraph(plan.risks || '');
  body.appendParagraph('');

  body.appendParagraph('Communication plan').setBold(true);
  body.appendParagraph(plan.communicationPlan || '');
  body.appendParagraph('');

  body.appendParagraph('Weekly milestones').setBold(true);

  if (plan.weeklyMilestones && plan.weeklyMilestones.length > 0) {
    plan.weeklyMilestones.forEach(function(milestone, index) {
      body.appendParagraph('Week ' + (index + 1) + ': ' + milestone);
    });
  } else {
    body.appendParagraph('Not set.');
  }

  doc.saveAndClose();

  logAudit(userEmail, OPERATION_TYPES.CREATE, projectId, {
    action: 'generate_90day_plan_doc',
    documentId: doc.getId(),
    usecaseId: usecaseId
  });

  return doc.getUrl();
}
