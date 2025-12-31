/**
 * VisionService.gs
 * 北極星ビジョン（Module 1）のビジネスロジック層
 */

// ========================================
// ビジョン保存・更新
// ========================================

/**
 * ビジョンを保存（バリデーション付き）
 * @param {Object} visionData - ビジョンデータ
 * @param {string} userEmail - ユーザーメール
 * @return {Object} 保存結果
 */
function saveVisionWithValidation(visionData, userEmail) {
  requireProjectAccess(visionData.projectId, userEmail);

  var project = getProject(visionData.projectId);
  if (!project) {
    throw new Error('プロジェクトが見つかりません');
  }

  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('アーカイブされたプロジェクトは編集できません');
  }

  // バリデーション
  if (!visionData.visionText || visionData.visionText.trim().length === 0) {
    throw new Error('ビジョン本文は必須です');
  }

  if (visionData.visionText.length > 2000) {
    throw new Error('ビジョン本文は2000文字以内で入力してください');
  }

  if (visionData.decisionRules && visionData.decisionRules.length > 3000) {
    throw new Error('意思決定ルールは3000文字以内で入力してください');
  }

  if (visionData.successMetrics && visionData.successMetrics.length > 1500) {
    throw new Error('成功指標は1500文字以内で入力してください');
  }

  // データを整形
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
    message: 'ビジョンを保存しました'
  };
}

/**
 * ビジョンの完成度をチェック
 * @param {string} projectId - プロジェクトID
 * @return {Object} 完成度情報
 */
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
    result.suggestions.push('ビジョンがまだ設定されていません');
    return result;
  }

  // 各項目のチェック
  if (vision.visionText && vision.visionText.length > 0) {
    result.hasVisionText = true;
  } else {
    result.suggestions.push('ビジョン本文を入力してください（必須）');
  }

  if (vision.decisionRules && vision.decisionRules.length > 0) {
    result.hasDecisionRules = true;
  } else {
    result.suggestions.push('意思決定ルールを追加すると、チームの判断基準が明確になります');
  }

  if (vision.successMetrics && vision.successMetrics.length > 0) {
    result.hasSuccessMetrics = true;
  } else {
    result.suggestions.push('成功指標を設定すると、進捗の測定が可能になります');
  }

  if (vision.notes && vision.notes.length > 0) {
    result.hasNotes = true;
  }

  // 完成度計算（ビジョン本文は必須なので重み付け）
  var score = 0;
  if (result.hasVisionText) score += 40;
  if (result.hasDecisionRules) score += 25;
  if (result.hasSuccessMetrics) score += 25;
  if (result.hasNotes) score += 10;

  result.completenessPercent = score;

  return result;
}

// ========================================
// ビジョン分析・提案
// ========================================

/**
 * ビジョン本文から成功指標候補を提案
 * @param {string} visionText - ビジョン本文
 * @return {Array<string>} 提案された成功指標
 */
function suggestSuccessMetrics(visionText) {
  var suggestions = [];

  if (!visionText) {
    return suggestions;
  }

  var text = visionText.toLowerCase();

  // キーワードに基づく提案
  if (text.indexOf('効率') !== -1 || text.indexOf('生産性') !== -1) {
    suggestions.push('業務処理時間の短縮率（目標: XX%削減）');
    suggestions.push('月間処理件数の増加（目標: 現状比XX%増）');
  }

  if (text.indexOf('コスト') !== -1 || text.indexOf('費用') !== -1) {
    suggestions.push('年間コスト削減額（目標: XX万円/年）');
    suggestions.push('運用コスト削減率（目標: XX%）');
  }

  if (text.indexOf('顧客') !== -1 || text.indexOf('満足') !== -1) {
    suggestions.push('顧客満足度スコア（目標: NPS XX以上）');
    suggestions.push('顧客対応時間の短縮（目標: XX%削減）');
  }

  if (text.indexOf('データ') !== -1 || text.indexOf('分析') !== -1) {
    suggestions.push('データ活用率（目標: XX%の意思決定でデータを活用）');
    suggestions.push('レポート作成時間の短縮（目標: XX時間→XX分）');
  }

  if (text.indexOf('売上') !== -1 || text.indexOf('収益') !== -1) {
    suggestions.push('売上増加率（目標: 前年比XX%増）');
    suggestions.push('利益率改善（目標: XX%向上）');
  }

  // 汎用的な提案
  if (suggestions.length === 0) {
    suggestions.push('ユーザー定着率（目標: 月間アクティブユーザーXX%）');
    suggestions.push('導入効果の定量化（目標: ROI XX%）');
    suggestions.push('プロジェクト完了率（目標: 計画の90%以上達成）');
  }

  return suggestions;
}

/**
 * 意思決定ルールのテンプレートを取得
 * @return {Array<Object>} ルールテンプレート
 */
function getDecisionRuleTemplates() {
  return [
    {
      category: '優先順位',
      templates: [
        '顧客価値 > 短期利益',
        'データに基づく意思決定を優先',
        'スピード重視 vs 品質重視のバランス基準'
      ]
    },
    {
      category: '技術選定',
      templates: [
        '既存システムとの整合性を最優先',
        'セキュリティ要件は非妥協',
        'スケーラビリティを考慮した設計'
      ]
    },
    {
      category: '組織・プロセス',
      templates: [
        'サイロ化よりも部門横断の協業を優先',
        '現場の声を必ず反映',
        '週次レビューで方向修正を判断'
      ]
    },
    {
      category: 'リスク管理',
      templates: [
        'リスクは早期に共有・エスカレーション',
        '失敗は学習機会として記録・共有',
        '不確実性が高い場合は小さく始める'
      ]
    }
  ];
}

// ========================================
// ビジョンドキュメント生成
// ========================================

/**
 * ビジョン1枚シートのプレビューデータを生成
 * @param {string} projectId - プロジェクトID
 * @return {Object} プレビューデータ
 */
function generateVisionPreview(projectId) {
  var project = getProject(projectId);
  var vision = getVision(projectId);

  if (!project) {
    throw new Error('プロジェクトが見つかりません');
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

/**
 * 複数行テキストを配列にパース
 * @param {string} text - テキスト
 * @return {Array<string>} 行の配列
 */
function parseMultilineText(text) {
  if (!text) return [];

  return text.split('\n')
    .map(function(line) { return line.trim(); })
    .filter(function(line) { return line.length > 0; });
}

// ========================================
// ビジョン履歴管理（オプション）
// ========================================

/**
 * ビジョンの変更履歴を取得
 * @param {string} projectId - プロジェクトID
 * @return {Array<Object>} 変更履歴
 */
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
