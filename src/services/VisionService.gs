/**
 * VisionService.gs
 * Vision module logic.
 */

function saveVisionWithValidation(visionData, userEmail) {
  requireProjectAccess(visionData.projectId, userEmail);

  var project = getProject(visionData.projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません。');
  }

  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('アーカイブ済みのプロジェクトは編集できません。');
  }

  if (!visionData.visionText || visionData.visionText.trim().length === 0) {
    throw new Error('ビジョンは必須です。');
  }

  if (visionData.visionText.length > 2000) {
    throw new Error('ビジョンは2000文字以内で入力してください。');
  }

  if (visionData.decisionRules && visionData.decisionRules.length > 3000) {
    throw new Error('意思決定ルールは3000文字以内で入力してください。');
  }

  if (visionData.successMetrics && visionData.successMetrics.length > 1500) {
    throw new Error('成功指標は1500文字以内で入力してください。');
  }

  var cleanData = {
    projectId: visionData.projectId,
    visionText: visionData.visionText.trim(),
    decisionRules: (visionData.decisionRules || '').trim(),
    successMetrics: (visionData.successMetrics || '').trim(),
    notes: (visionData.notes || '').trim(),
    userEmail: userEmail
  };

  saveVision(cleanData);

  return {
    success: true,
    message: 'ビジョンを保存しました。'
  };
}

function checkVisionCompleteness(projectId) {
  var vision = getVision(projectId);

  var result = {
    hasVisionText: false,
    hasDecisionRules: false,
    hasSuccessMetrics: false,
    hasNotes: false,
    completenessPercent: 0,
    suggestions: []
  };

  if (!vision) {
    result.suggestions.push('Vision has not been set yet.');
    return result;
  }

  if (vision.visionText && vision.visionText.length > 0) {
    result.hasVisionText = true;
  } else {
    result.suggestions.push('Add a vision statement.');
  }

  if (vision.decisionRules && vision.decisionRules.length > 0) {
    result.hasDecisionRules = true;
  } else {
    result.suggestions.push('Add decision rules to clarify the team criteria.');
  }

  if (vision.successMetrics && vision.successMetrics.length > 0) {
    result.hasSuccessMetrics = true;
  } else {
    result.suggestions.push('Add success metrics to track progress.');
  }

  if (vision.notes && vision.notes.length > 0) {
    result.hasNotes = true;
  }

  var score = 0;
  if (result.hasVisionText) score += 40;
  if (result.hasDecisionRules) score += 25;
  if (result.hasSuccessMetrics) score += 25;
  if (result.hasNotes) score += 10;

  result.completenessPercent = score;

  return result;
}

function suggestSuccessMetrics(visionText) {
  var suggestions = [];

  if (!visionText) {
    return suggestions;
  }

  var text = visionText.toLowerCase();

  if (text.indexOf('efficien') !== -1 || text.indexOf('productivity') !== -1) {
    suggestions.push('Reduce cycle time by XX%.');
    suggestions.push('Increase monthly throughput by XX%.');
  }

  if (text.indexOf('cost') !== -1 || text.indexOf('expense') !== -1) {
    suggestions.push('Annual cost reduction of $XX.');
    suggestions.push('Operating cost reduction rate of XX%.');
  }

  if (text.indexOf('customer') !== -1 || text.indexOf('satisfaction') !== -1) {
    suggestions.push('NPS target of XX or higher.');
    suggestions.push('Customer response time reduced by XX%.');
  }

  if (text.indexOf('data') !== -1 || text.indexOf('insight') !== -1) {
    suggestions.push('Data-driven decisions in XX% of key meetings.');
    suggestions.push('Reporting time reduced from XX hours to XX hours.');
  }

  if (text.indexOf('revenue') !== -1 || text.indexOf('profit') !== -1) {
    suggestions.push('Year-over-year revenue growth of XX%.');
    suggestions.push('Margin improvement of XX%.');
  }

  if (suggestions.length === 0) {
    suggestions.push('Monthly active users increase by XX%.');
    suggestions.push('ROI target of XX%.');
    suggestions.push('Project milestones completed at 90%+.');
  }

  return suggestions;
}

function getDecisionRuleTemplates() {
  return [
    {
      category: 'Prioritization',
      templates: [
        'Customer value over short-term gains.',
        'Prefer data-driven decisions.',
        'Balance speed and quality.'
      ]
    },
    {
      category: 'Technology',
      templates: [
        'Prefer alignment with existing systems.',
        'Security requirements are non-negotiable.',
        'Design for scalability from day one.'
      ]
    },
    {
      category: 'Organization',
      templates: [
        'Cross-functional collaboration over silos.',
        'Frontline feedback must be reflected.',
        'Weekly reviews to adjust direction.'
      ]
    },
    {
      category: 'Risk',
      templates: [
        'Share risks early and escalate quickly.',
        'Treat failures as learning opportunities.',
        'Start small when uncertainty is high.'
      ]
    }
  ];
}

function generateVisionPreview(projectId) {
  var project = getProject(projectId);
  var vision = getVision(projectId);

  if (!project) {
    throw new Error('プロジェクトが見つかりません。');
  }

  return {
    customerName: project.customerName,
    visionText: vision ? vision.visionText : '',
    decisionRules: vision ? parseMultilineText(vision.decisionRules) : [],
    successMetrics: vision ? parseMultilineText(vision.successMetrics) : [],
    notes: vision ? vision.notes : '',
    lastUpdated: vision ? vision.updatedDate : null,
    completeness: checkVisionCompleteness(projectId)
  };
}

function parseMultilineText(text) {
  if (!text) return [];

  return text.split('\n')
    .map(function(line) { return line.trim(); })
    .filter(function(line) { return line.length > 0; });
}

function getVisionHistory(projectId) {
  var logs = getAuditLogs(projectId);

  return logs
    .filter(function(log) {
      return log.details && log.details.module === 'vision';
    })
    .map(function(log) {
      return {
        timestamp: log.timestamp,
        user: log.user,
        operation: log.operation
      };
    });
}
