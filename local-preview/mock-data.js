/**
 * Mock Data for Local Preview
 * テスト用のサンプルデータ
 */

const MOCK_USER_EMAIL = 'demo@example.com';

// プロジェクトデータ
const MOCK_PROJECTS = [
  {
    projectId: 'PRJ-20251231-0001',
    customerName: 'Acme Corporation',
    createdDate: '2025-12-20T10:00:00Z',
    updatedDate: '2025-12-31T14:30:00Z',
    creatorEmail: 'demo@example.com',
    editorEmails: 'demo@example.com',
    status: '下書き'
  },
  {
    projectId: 'PRJ-20251231-0002',
    customerName: 'Global Tech Inc.',
    createdDate: '2025-12-15T09:00:00Z',
    updatedDate: '2025-12-28T11:00:00Z',
    creatorEmail: 'demo@example.com',
    editorEmails: 'demo@example.com,team@example.com',
    status: '確定'
  },
  {
    projectId: 'PRJ-20251231-0003',
    customerName: 'StartUp Labs',
    createdDate: '2025-11-01T08:00:00Z',
    updatedDate: '2025-11-30T17:00:00Z',
    creatorEmail: 'demo@example.com',
    editorEmails: 'demo@example.com',
    status: 'アーカイブ'
  }
];

// ビジョンデータ
const MOCK_VISIONS = {
  'PRJ-20251231-0001': {
    projectId: 'PRJ-20251231-0001',
    visionText: 'To become the leading data-driven organization in our industry, empowering every team member to make informed decisions through accessible, real-time insights.',
    decisionRules: '1. Data quality over speed of delivery\n2. User adoption is the key metric\n3. Security and governance cannot be compromised',
    successMetrics: '- 80% of business users actively using dashboards weekly\n- 50% reduction in report generation time\n- NPS score > 60 from internal users',
    notes: 'Focus on quick wins in Q1 to build momentum'
  },
  'PRJ-20251231-0002': {
    projectId: 'PRJ-20251231-0002',
    visionText: 'Enable seamless collaboration across global teams through unified analytics platform.',
    decisionRules: '1. Mobile-first approach\n2. Real-time data sync',
    successMetrics: '- 100% team onboarding within 90 days\n- 30% productivity improvement',
    notes: ''
  }
};

// ユースケースデータ
const MOCK_USECASES = {
  'PRJ-20251231-0001': [
    {
      usecaseId: 'UC-20251231-001',
      projectId: 'PRJ-20251231-0001',
      challenge: 'Sales team spends 4+ hours weekly creating manual reports',
      goal: 'Automate sales reporting with self-service dashboards',
      expectedImpact: 'Save 200+ hours monthly across sales team, faster decision making',
      ninetyDayGoal: 'Launch automated sales dashboard with top 5 KPIs',
      score: 85,
      priority: 1
    },
    {
      usecaseId: 'UC-20251231-002',
      projectId: 'PRJ-20251231-0001',
      challenge: 'Marketing ROI is difficult to measure across channels',
      goal: 'Create unified marketing attribution dashboard',
      expectedImpact: 'Better budget allocation, improved campaign performance',
      ninetyDayGoal: 'Integrate 3 major marketing platforms into single view',
      score: 72,
      priority: 2
    },
    {
      usecaseId: 'UC-20251231-003',
      projectId: 'PRJ-20251231-0001',
      challenge: 'Customer churn prediction is currently reactive',
      goal: 'Build predictive churn model with early warning system',
      expectedImpact: 'Reduce churn by 15%, increase customer lifetime value',
      ninetyDayGoal: 'Deploy initial churn prediction model for top 100 accounts',
      score: 68,
      priority: 3
    }
  ],
  'PRJ-20251231-0002': [
    {
      usecaseId: 'UC-20251231-004',
      projectId: 'PRJ-20251231-0002',
      challenge: 'Global inventory visibility is fragmented',
      goal: 'Real-time inventory tracking across all regions',
      expectedImpact: 'Reduce stockouts by 30%, optimize working capital',
      ninetyDayGoal: 'Connect 5 major warehouses to central dashboard',
      score: 90,
      priority: 1
    }
  ]
};

// RACIデータ
const MOCK_RACI = {
  'PRJ-20251231-0001': [
    { pillar: 'CoE', task: 'Define data governance policies', assignee: 'Sarah Chen', raci: 'A' },
    { pillar: 'CoE', task: 'Establish best practices', assignee: 'Sarah Chen', raci: 'R' },
    { pillar: 'ビジネスデータネットワーク', task: 'Identify key business metrics', assignee: 'Mike Johnson', raci: 'R' },
    { pillar: 'ビジネスデータネットワーク', task: 'User training and adoption', assignee: 'Lisa Park', raci: 'R' },
    { pillar: 'IT', task: 'Data pipeline development', assignee: 'Tom Wilson', raci: 'R' },
    { pillar: 'IT', task: 'Security implementation', assignee: 'Tom Wilson', raci: 'A' },
    { pillar: 'IT', task: 'Infrastructure setup', assignee: 'Alex Kim', raci: 'R' }
  ],
  'PRJ-20251231-0002': [
    { pillar: 'CoE', task: 'Global standards alignment', assignee: 'Emma Davis', raci: 'A' },
    { pillar: 'IT', task: 'API integration', assignee: 'James Lee', raci: 'R' }
  ]
};

// 価値追跡データ
const MOCK_VALUES = {
  'PRJ-20251231-0001': [
    {
      usecaseId: 'UC-20251231-001',
      projectId: 'PRJ-20251231-0001',
      quantitativeImpact: '150 hours saved per month, $45,000 annual cost reduction',
      qualitativeImpact: 'Improved sales team morale, faster response to market changes',
      evidence: 'https://drive.google.com/example-evidence-001',
      nextInvestment: 'Expand'
    },
    {
      usecaseId: 'UC-20251231-002',
      projectId: 'PRJ-20251231-0001',
      quantitativeImpact: '',
      qualitativeImpact: 'In progress - initial feedback positive',
      evidence: '',
      nextInvestment: 'Continue'
    }
  ],
  'PRJ-20251231-0002': []
};

// 90日計画データ
const MOCK_90DAY_PLANS = {
  'UC-20251231-001': {
    usecaseId: 'UC-20251231-001',
    projectId: 'PRJ-20251231-0001',
    structure: 'Project Lead: Mike Johnson, Technical Lead: Tom Wilson',
    requiredData: 'CRM data, Sales transactions, Product catalog',
    risks: 'Data quality issues, User adoption resistance',
    communicationPlan: 'Weekly standup, Monthly steering committee',
    milestones: JSON.stringify([
      { week: 1, task: 'Requirements gathering' },
      { week: 2, task: 'Data source mapping' },
      { week: 3, task: 'Pipeline development' },
      { week: 4, task: 'Initial dashboard prototype' },
      { week: 5, task: 'User feedback round 1' },
      { week: 6, task: 'Iteration and refinement' },
      { week: 7, task: 'Security review' },
      { week: 8, task: 'UAT preparation' },
      { week: 9, task: 'User acceptance testing' },
      { week: 10, task: 'Training materials' },
      { week: 11, task: 'Pilot launch' },
      { week: 12, task: 'Full rollout' }
    ])
  }
};
