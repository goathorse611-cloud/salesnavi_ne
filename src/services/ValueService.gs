/**
 * ValueService.gs
 * 価値実現トラッカー（Module 7）のビジネスロジック層
 */

// ========================================
// 価値トラッキング管理
// ========================================

/**
 * 価値トラッキングを保存（バリデーション付き）
 * @param {Object} valueData - 価値データ
 * @param {string} userEmail - ユーザーメール
 */
function saveValueWithValidation(valueData, userEmail) {
  requireProjectAccess(valueData.projectId, userEmail);

  var project = getProject(valueData.projectId);
  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('アーカイブされたプロジェクトは編集できません');
  }

  // ユースケースの存在確認
  var usecases = getUsecases(valueData.projectId);
  var exists = usecases.some(function(uc) {
    return uc.usecaseId === valueData.usecaseId;
  });

  if (!exists) {
    throw new Error('指定されたユースケースが見つかりません');
  }

  // バリデーション
  if (valueData.quantitativeImpact && valueData.quantitativeImpact.length > 2000) {
    throw new Error('定量効果は2000文字以内で入力してください');
  }

  if (valueData.qualitativeImpact && valueData.qualitativeImpact.length > 2000) {
    throw new Error('定性効果は2000文字以内で入力してください');
  }

  if (valueData.evidence && valueData.evidence.length > 500) {
    throw new Error('証跡URLは500文字以内で入力してください');
  }

  // 証跡URLの形式チェック（任意）
  if (valueData.evidence && valueData.evidence.trim().length > 0) {
    var url = valueData.evidence.trim();
    if (url.indexOf('http://') !== 0 && url.indexOf('https://') !== 0 && url.indexOf('drive.google.com') === -1) {
      // 警告のみ、エラーにはしない
      Logger.log('証跡URL形式が不正な可能性: ' + url);
    }
  }

  // データを整形
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

/**
 * 全ユースケースの価値トラッキング状況を取得
 * @param {string} projectId - プロジェクトID
 * @return {Array<Object>} ユースケースごとの価値データ
 */
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

/**
 * 価値データの完成度を計算
 * @param {Object} value - 価値データ
 * @return {number} 完成度パーセント
 */
function calculateValueCompleteness(value) {
  if (!value) return 0;

  var score = 0;
  if (value.quantitativeImpact && value.quantitativeImpact.length > 0) score += 30;
  if (value.qualitativeImpact && value.qualitativeImpact.length > 0) score += 25;
  if (value.evidence && value.evidence.length > 0) score += 25;
  if (value.nextInvestment && value.nextInvestment.length > 0) score += 20;

  return score;
}

// ========================================
// 価値分析・レポート
// ========================================

/**
 * プロジェクト全体の価値サマリーを取得
 * @param {string} projectId - プロジェクトID
 * @return {Object} 価値サマリー
 */
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

    // 定量効果を収集
    if (val.quantitativeImpact && val.quantitativeImpact.length > 0) {
      result.quantitativeEffects.push({
        usecaseId: val.usecaseId,
        effect: val.quantitativeImpact
      });
    }

    // 定性効果を収集
    if (val.qualitativeImpact && val.qualitativeImpact.length > 0) {
      result.qualitativeEffects.push({
        usecaseId: val.usecaseId,
        effect: val.qualitativeImpact
      });
    }

    // 投資判断を分類
    if (val.nextInvestment) {
      var decision = val.nextInvestment.toLowerCase();
      if (decision.indexOf('継続') !== -1) {
        result.investmentDecisions.continue++;
      } else if (decision.indexOf('拡大') !== -1 || decision.indexOf('拡張') !== -1) {
        result.investmentDecisions.expand++;
      } else if (decision.indexOf('縮小') !== -1 || decision.indexOf('終了') !== -1) {
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

/**
 * 投資判断の推奨を生成
 * @param {string} projectId - プロジェクトID
 * @param {string} usecaseId - ユースケースID
 * @return {Object} 推奨情報
 */
function generateInvestmentRecommendation(projectId, usecaseId) {
  var usecases = getUsecases(projectId);
  var usecase = usecases.find(function(uc) {
    return uc.usecaseId === usecaseId;
  });

  if (!usecase) {
    throw new Error('ユースケースが見つかりません');
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

  // 価値データがない場合
  if (!value || calculateValueCompleteness(value) < 50) {
    result.recommendation = 'needs_tracking';
    result.reasons.push('価値トラッキングデータが不十分です');
    result.nextSteps.push('定量効果と定性効果を記録してください');
    result.nextSteps.push('証跡（スクリーンショット、レポート等）を添付してください');
    return result;
  }

  // 定量効果の分析
  var quantitative = value.quantitativeImpact || '';
  var hasPositiveNumbers = /[0-9]+%/.test(quantitative) || /向上|改善|削減|増加/.test(quantitative);

  // 証跡の有無
  var hasEvidence = value.evidence && value.evidence.length > 0;

  // 推奨を生成
  if (hasPositiveNumbers && hasEvidence) {
    result.recommendation = 'expand';
    result.reasons.push('定量的な効果が確認されています');
    result.reasons.push('証跡が記録されています');
    result.nextSteps.push('他部門への横展開を検討');
    result.nextSteps.push('効果をさらに拡大する追加投資を検討');
  } else if (hasPositiveNumbers) {
    result.recommendation = 'continue';
    result.reasons.push('定量的な効果が確認されています');
    result.reasons.push('証跡の追加が推奨されます');
    result.nextSteps.push('証跡を収集して記録');
    result.nextSteps.push('効果の持続性を確認');
  } else {
    result.recommendation = 'review';
    result.reasons.push('効果の定量化が必要です');
    result.nextSteps.push('KPIを設定して測定開始');
    result.nextSteps.push('90日後に再評価');
  }

  return result;
}

// ========================================
// 価値テンプレート
// ========================================

/**
 * 価値トラッキングのテンプレートを取得
 * @return {Object} テンプレート
 */
function getValueTrackingTemplates() {
  return {
    quantitativeTemplates: [
      {
        category: '時間削減',
        examples: [
          'レポート作成時間: XX時間/月 → XX時間/月（XX%削減）',
          'データ収集時間: XX日 → XX時間（XX%削減）',
          '意思決定サイクル: XX日 → XX日（XX%短縮）'
        ]
      },
      {
        category: 'コスト削減',
        examples: [
          '運用コスト: 年間XX万円削減',
          '外注費: XX万円/月 → XX万円/月',
          '人件費（工数換算）: XX人月削減'
        ]
      },
      {
        category: '売上・収益',
        examples: [
          '売上向上: 前年比XX%増',
          '機会損失削減: 推定XX万円/月',
          '新規案件獲得: XX件/四半期'
        ]
      },
      {
        category: '品質・精度',
        examples: [
          'データ精度: XX% → XX%',
          'エラー率: XX% → XX%（XX%削減）',
          '予測精度: XX%向上'
        ]
      }
    ],
    qualitativeTemplates: [
      {
        category: '業務改善',
        examples: [
          '意思決定の迅速化・根拠の明確化',
          '部門間の情報共有促進',
          'ダッシュボードによるリアルタイム状況把握'
        ]
      },
      {
        category: '組織・文化',
        examples: [
          'データドリブン文化の醸成',
          '社員のデータリテラシー向上',
          '経営層のデータ活用への理解深化'
        ]
      },
      {
        category: '顧客体験',
        examples: [
          '顧客対応品質の向上',
          'パーソナライズされた提案の実現',
          '顧客満足度の向上'
        ]
      }
    ],
    investmentDecisionOptions: [
      {
        value: '継続',
        description: '現状維持で価値が持続している',
        nextAction: '定期的な効果測定を継続'
      },
      {
        value: '拡大',
        description: '効果が確認され、さらなる投資価値がある',
        nextAction: '他部門への展開、機能拡張を検討'
      },
      {
        value: '縮小',
        description: '期待した効果が得られていない',
        nextAction: '原因分析、スコープ見直しを実施'
      },
      {
        value: '終了',
        description: '目的達成または効果なし',
        nextAction: '振り返りを実施し、学びを記録'
      }
    ]
  };
}

/**
 * 価値完成度チェック
 * @param {string} projectId - プロジェクトID
 * @return {Object} 完成度情報
 */
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
    result.suggestions.push('先にユースケースを作成してください');
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

  // 完成度計算
  var score = 0;
  if (result.hasAnyValue) score += 25;
  if (result.hasAllUsecasesTracked) score += 25;
  if (result.hasEvidence) score += 25;
  if (result.hasInvestmentDecisions) score += 25;

  result.completenessPercent = score;

  // 提案
  if (!result.hasAnyValue) {
    result.suggestions.push('少なくとも1つのユースケースの価値を記録してください');
  }
  if (!result.hasAllUsecasesTracked) {
    result.suggestions.push('すべてのユースケースの価値を記録すると、稟議書の説得力が上がります');
  }
  if (!result.hasEvidence) {
    result.suggestions.push('証跡（Driveリンク、スクリーンショット等）を追加してください');
  }
  if (!result.hasInvestmentDecisions) {
    result.suggestions.push('次の投資判断を記録してください');
  }

  return result;
}
