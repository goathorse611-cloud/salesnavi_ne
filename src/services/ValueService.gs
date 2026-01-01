/**
 * ValueService.gs
 * Value tracking logic.
 */

function saveValueWithValidation(valueData, userEmail) {
  requireProjectAccess(valueData.projectId, userEmail);

  var project = getProject(valueData.projectId);
  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('アーカイブ済みのプロジェクトは編集できません。');
  }

  var usecases = getUsecases(valueData.projectId);
  var exists = usecases.some(function(uc) {
    return uc.usecaseId === valueData.usecaseId;
  });

  if (!exists) {
    throw new Error('ユースケースが見つかりません。');
  }

  if (valueData.quantitativeImpact && valueData.quantitativeImpact.length > 2000) {
    throw new Error('定量的インパクトは2000文字以内で入力してください。');
  }

  if (valueData.qualitativeImpact && valueData.qualitativeImpact.length > 2000) {
    throw new Error('定性的インパクトは2000文字以内で入力してください。');
  }

  if (valueData.evidence && valueData.evidence.length > 500) {
    throw new Error('エビデンスURLは500文字以内で入力してください。');
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
    throw new Error('ユースケースが見つかりません。');
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
    result.reasons.push('価値トラッキングのデータが不足しています。');
    result.nextSteps.push('定量・定性的インパクトを記録してください。');
    result.nextSteps.push('エビデンス（スクリーンショットやレポート）を添付してください。');
    return result;
  }

  var quantitative = value.quantitativeImpact || '';
  var hasPositiveNumbers = /[0-9]+%/.test(quantitative) || /increase|improve|reduce|gain|増加|改善|削減|向上/i.test(quantitative);

  var hasEvidence = value.evidence && value.evidence.length > 0;

  if (hasPositiveNumbers && hasEvidence) {
    result.recommendation = 'expand';
    result.reasons.push('定量的インパクトが確認できます。');
    result.reasons.push('エビデンスが記録されています。');
    result.nextSteps.push('他チームへの展開を検討してください。');
    result.nextSteps.push('追加投資の計画を立ててください。');
  } else if (hasPositiveNumbers) {
    result.recommendation = 'continue';
    result.reasons.push('定量的インパクトが見えています。');
    result.reasons.push('追加のエビデンスが推奨されます。');
    result.nextSteps.push('エビデンスを記録してください。');
    result.nextSteps.push('効果の持続性を確認してください。');
  } else {
    result.recommendation = 'review';
    result.reasons.push('定量的インパクトが不明確です。');
    result.nextSteps.push('KPIを定義し、結果を測定してください。');
    result.nextSteps.push('90日後に再評価してください。');
  }

  return result;
}

function getValueTrackingTemplates() {
  return {
    quantitativeTemplates: [
      {
        category: '時間削減',
        examples: [
          'レポート作成時間がXX時間からXX時間に短縮。',
          'データ準備時間をXX%削減。',
          '意思決定サイクルをXX%短縮。'
        ]
      },
      {
        category: 'コスト削減',
        examples: [
          '年間運用コストをXX円削減。',
          'ベンダー費用を月額XX円削減。',
          '工数をXX人日削減。'
        ]
      },
      {
        category: '売上インパクト',
        examples: [
          '売上が前年比XX%増加。',
          '機会損失をXX円削減。',
          '新規成約: 四半期あたりXX件。'
        ]
      },
      {
        category: '品質',
        examples: [
          'データ精度がXX%からXX%に改善。',
          'エラー率をXX%削減。',
          '予測精度をXX%改善。'
        ]
      }
    ],
    qualitativeTemplates: [
      {
        category: 'プロセス改善',
        examples: [
          '部門横断の合意形成が迅速化。',
          '意思決定の根拠が明確化。',
          'リアルタイムの運用可視化。'
        ]
      },
      {
        category: '文化',
        examples: [
          'データドリブン文化の醸成。',
          '分析リテラシーの向上。',
          '経営層の関与強化。'
        ]
      },
      {
        category: '顧客体験',
        examples: [
          '対応品質の向上。',
          '提案のパーソナライズ。',
          '満足度スコアの向上。'
        ]
      }
    ],
    investmentDecisionOptions: [
      {
        value: 'Continue',
        description: '現在のスコープを維持。',
        nextAction: '価値測定を継続。'
      },
      {
        value: 'Expand',
        description: '対象チームやユースケースを拡大。',
        nextAction: '広範な展開計画を策定。'
      },
      {
        value: 'Reduce',
        description: '成果に基づきスコープを縮小。',
        nextAction: '最も効果の高い領域に集中。'
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
    result.suggestions.push('価値を追跡する前にユースケースを作成してください。');
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
    result.suggestions.push('少なくとも1件のユースケースで価値を記録してください。');
  }
  if (!result.hasAllUsecasesTracked) {
    result.suggestions.push('提案力を高めるため、すべてのユースケースで価値を追跡してください。');
  }
  if (!result.hasEvidence) {
    result.suggestions.push('エビデンスのリンクや資料を追加してください。');
  }
  if (!result.hasInvestmentDecisions) {
    result.suggestions.push('次の投資判断を記録してください。');
  }

  return result;
}
