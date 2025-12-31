/**
 * ValueService.gs
 * Value tracking logic.
 */

function saveValueWithValidation(valueData, userEmail) {
  requireProjectAccess(valueData.projectId, userEmail);

  var project = getProject(valueData.projectId);
  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('Archived projects cannot be edited.');
  }

  var usecases = getUsecases(valueData.projectId);
  var exists = usecases.some(function(uc) {
    return uc.usecaseId === valueData.usecaseId;
  });

  if (!exists) {
    throw new Error('Use case not found.');
  }

  if (valueData.quantitativeImpact && valueData.quantitativeImpact.length > 2000) {
    throw new Error('Quantitative impact must be 2000 characters or less.');
  }

  if (valueData.qualitativeImpact && valueData.qualitativeImpact.length > 2000) {
    throw new Error('Qualitative impact must be 2000 characters or less.');
  }

  if (valueData.evidence && valueData.evidence.length > 500) {
    throw new Error('Evidence URL must be 500 characters or less.');
  }

  if (valueData.evidence && valueData.evidence.trim().length > 0) {
    var url = valueData.evidence.trim();
    if (url.indexOf('http://') !== 0 &&
        url.indexOf('https://') !== 0 &&
        url.indexOf('drive.google.com') === -1) {
      Logger.log('Evidence URL may be invalid: ' + url);
    }
  }

  var cleanData = {
    projectId: valueData.projectId,
    usecaseId: valueData.usecaseId,
    quantitativeImpact: (valueData.quantitativeImpact || '').trim(),
    qualitativeImpact: (valueData.qualitativeImpact || '').trim(),
    evidence: (valueData.evidence || '').trim(),
    nextInvestment: (valueData.nextInvestment || '').trim(),
    userEmail: userEmail
  };

  saveValue(cleanData);
}

function getValueTrackingStatus(projectId) {
  var usecases = getUsecases(projectId);
  var values = getValues(projectId);

  return usecases.map(function(uc) {
    var value = values.find(function(v) {
      return v.usecaseId === uc.usecaseId;
    });

    return {
      usecaseId: uc.usecaseId,
      goal: uc.goal,
      challenge: uc.challenge,
      hasValue: value !== undefined,
      quantitativeImpact: value ? value.quantitativeImpact : null,
      qualitativeImpact: value ? value.qualitativeImpact : null,
      evidence: value ? value.evidence : null,
      nextInvestment: value ? value.nextInvestment : null,
      completeness: value ? calculateValueCompleteness(value) : 0
    };
  });
}

function calculateValueCompleteness(value) {
  if (!value) return 0;

  var score = 0;
  if (value.quantitativeImpact && value.quantitativeImpact.length > 0) score += 30;
  if (value.qualitativeImpact && value.qualitativeImpact.length > 0) score += 25;
  if (value.evidence && value.evidence.length > 0) score += 25;
  if (value.nextInvestment && value.nextInvestment.length > 0) score += 20;

  return score;
}

function getValueSummary(projectId) {
  var usecases = getUsecases(projectId);
  var values = getValues(projectId);

  var result = {
    totalUsecases: usecases.length,
    trackedUsecases: 0,
    completelyTracked: 0,
    hasEvidence: 0,
    averageCompleteness: 0,
    investmentDecisions: {
      continue: 0,
      expand: 0,
      reduce: 0,
      undecided: 0
    },
    quantitativeEffects: [],
    qualitativeEffects: []
  };

  if (usecases.length === 0) {
    return result;
  }

  var totalCompleteness = 0;

  values.forEach(function(val) {
    result.trackedUsecases++;

    var completeness = calculateValueCompleteness(val);
    totalCompleteness += completeness;

    if (completeness >= 80) {
      result.completelyTracked++;
    }

    if (val.evidence && val.evidence.length > 0) {
      result.hasEvidence++;
    }

    if (val.quantitativeImpact && val.quantitativeImpact.length > 0) {
      result.quantitativeEffects.push({
        usecaseId: val.usecaseId,
        effect: val.quantitativeImpact
      });
    }

    if (val.qualitativeImpact && val.qualitativeImpact.length > 0) {
      result.qualitativeEffects.push({
        usecaseId: val.usecaseId,
        effect: val.qualitativeImpact
      });
    }

    if (val.nextInvestment) {
      var decision = val.nextInvestment.toLowerCase();
      if (decision.indexOf('continue') !== -1) {
        result.investmentDecisions.continue++;
      } else if (decision.indexOf('expand') !== -1) {
        result.investmentDecisions.expand++;
      } else if (decision.indexOf('reduce') !== -1) {
        result.investmentDecisions.reduce++;
      } else {
        result.investmentDecisions.undecided++;
      }
    } else {
      result.investmentDecisions.undecided++;
    }
  });

  if (result.trackedUsecases > 0) {
    result.averageCompleteness = Math.round(totalCompleteness / result.trackedUsecases);
  }

  return result;
}

function generateInvestmentRecommendation(projectId, usecaseId) {
  var usecases = getUsecases(projectId);
  var usecase = usecases.find(function(uc) {
    return uc.usecaseId === usecaseId;
  });

  if (!usecase) {
    throw new Error('Use case not found.');
  }

  var values = getValues(projectId);
  var value = values.find(function(v) {
    return v.usecaseId === usecaseId;
  });

  var result = {
    usecaseId: usecaseId,
    goal: usecase.goal,
    recommendation: 'undecided',
    reasons: [],
    nextSteps: []
  };

  if (!value || calculateValueCompleteness(value) < 50) {
    result.recommendation = 'needs_tracking';
    result.reasons.push('Value tracking data is incomplete.');
    result.nextSteps.push('Record quantitative and qualitative impacts.');
    result.nextSteps.push('Attach evidence (screenshots or reports).');
    return result;
  }

  var quantitative = value.quantitativeImpact || '';
  var hasPositiveNumbers = /[0-9]+%/.test(quantitative) || /increase|improve|reduce|gain/i.test(quantitative);

  var hasEvidence = value.evidence && value.evidence.length > 0;

  if (hasPositiveNumbers && hasEvidence) {
    result.recommendation = 'expand';
    result.reasons.push('Quantitative impact is evident.');
    result.reasons.push('Evidence is recorded.');
    result.nextSteps.push('Consider scaling to other teams.');
    result.nextSteps.push('Plan additional investment.');
  } else if (hasPositiveNumbers) {
    result.recommendation = 'continue';
    result.reasons.push('Quantitative impact is visible.');
    result.reasons.push('Additional evidence is recommended.');
    result.nextSteps.push('Capture evidence artifacts.');
    result.nextSteps.push('Verify sustainability of impact.');
  } else {
    result.recommendation = 'review';
    result.reasons.push('Quantitative impact is unclear.');
    result.nextSteps.push('Define KPIs and measure results.');
    result.nextSteps.push('Re-evaluate after 90 days.');
  }

  return result;
}

function getValueTrackingTemplates() {
  return {
    quantitativeTemplates: [
      {
        category: 'Time savings',
        examples: [
          'Reporting time reduced from XX hrs to XX hrs.',
          'Data preparation time reduced by XX%.',
          'Decision cycle reduced by XX%.'
        ]
      },
      {
        category: 'Cost reduction',
        examples: [
          'Annual operating cost reduced by $XX.',
          'Vendor spend reduced by $XX per month.',
          'Labor hours reduced by XX person-days.'
        ]
      },
      {
        category: 'Revenue impact',
        examples: [
          'Revenue increased by XX% YoY.',
          'Lost opportunities reduced by $XX.',
          'New deals won: XX per quarter.'
        ]
      },
      {
        category: 'Quality',
        examples: [
          'Data accuracy improved from XX% to XX%.',
          'Error rate reduced by XX%.',
          'Forecast accuracy improved by XX%.'
        ]
      }
    ],
    qualitativeTemplates: [
      {
        category: 'Process improvement',
        examples: [
          'Faster cross-team alignment.',
          'Clearer decision rationale.',
          'Real-time operational visibility.'
        ]
      },
      {
        category: 'Culture',
        examples: [
          'More data-driven culture.',
          'Higher analytics literacy.',
          'Stronger executive engagement.'
        ]
      },
      {
        category: 'Customer experience',
        examples: [
          'Improved response quality.',
          'More personalized proposals.',
          'Higher satisfaction scores.'
        ]
      }
    ],
    investmentDecisionOptions: [
      {
        value: 'Continue',
        description: 'Maintain current scope.',
        nextAction: 'Continue measuring value.'
      },
      {
        value: 'Expand',
        description: 'Scale to more teams or use cases.',
        nextAction: 'Plan broader rollout.'
      },
      {
        value: 'Reduce',
        description: 'Reduce scope based on outcomes.',
        nextAction: 'Refocus on highest impact.'
      }
    ]
  };
}

function checkValueCompleteness(projectId) {
  var status = getValueTrackingStatus(projectId);

  var result = {
    hasAnyValue: false,
    hasAllUsecasesTracked: false,
    hasEvidence: false,
    hasInvestmentDecisions: false,
    completenessPercent: 0,
    suggestions: []
  };

  if (status.length === 0) {
    result.suggestions.push('Create use cases before tracking value.');
    return result;
  }

  var trackedCount = 0;
  var evidenceCount = 0;
  var decisionCount = 0;

  status.forEach(function(s) {
    if (s.hasValue) {
      result.hasAnyValue = true;
      trackedCount++;

      if (s.evidence) evidenceCount++;
      if (s.nextInvestment) decisionCount++;
    }
  });

  result.hasAllUsecasesTracked = trackedCount === status.length;
  result.hasEvidence = evidenceCount > 0;
  result.hasInvestmentDecisions = decisionCount > 0;

  var score = 0;
  if (result.hasAnyValue) score += 25;
  if (result.hasAllUsecasesTracked) score += 25;
  if (result.hasEvidence) score += 25;
  if (result.hasInvestmentDecisions) score += 25;

  result.completenessPercent = score;

  if (!result.hasAnyValue) {
    result.suggestions.push('Record value for at least one use case.');
  }
  if (!result.hasAllUsecasesTracked) {
    result.suggestions.push('Track value for all use cases for stronger proposals.');
  }
  if (!result.hasEvidence) {
    result.suggestions.push('Add evidence links or artifacts.');
  }
  if (!result.hasInvestmentDecisions) {
    result.suggestions.push('Record the next investment decision.');
  }

  return result;
}
