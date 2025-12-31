/**
 * OrganizationService.gs
 * Org structure and RACI logic.
 */

function saveRACIEntriesWithValidation(projectId, entries, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var project = getProject(projectId);
  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('Archived projects cannot be edited.');
  }

  if (!entries || !Array.isArray(entries)) {
    throw new Error('RACI entries must be an array.');
  }

  if (entries.length > 100) {
    throw new Error('RACI entries must be 100 or fewer.');
  }

  var validPillars = [
    THREE_PILLARS.COE,
    THREE_PILLARS.BUSINESS_DATA_NETWORK,
    THREE_PILLARS.IT
  ];
  var validRaciTypes = [
    RACI_TYPES.RESPONSIBLE,
    RACI_TYPES.ACCOUNTABLE,
    RACI_TYPES.CONSULTED,
    RACI_TYPES.INFORMED
  ];

  var cleanedEntries = [];

  entries.forEach(function(entry, index) {
    if (!entry.pillar || validPillars.indexOf(entry.pillar) === -1) {
      throw new Error('Entry ' + (index + 1) + ': invalid pillar.');
    }

    if (!entry.task || entry.task.trim().length === 0) {
      throw new Error('Entry ' + (index + 1) + ': task is required.');
    }

    if (!entry.assignee || entry.assignee.trim().length === 0) {
      throw new Error('Entry ' + (index + 1) + ': assignee is required.');
    }

    if (!entry.raci || validRaciTypes.indexOf(entry.raci) === -1) {
      throw new Error('Entry ' + (index + 1) + ': invalid RACI type.');
    }

    cleanedEntries.push({
      pillar: entry.pillar,
      task: entry.task.trim(),
      assignee: entry.assignee.trim(),
      raci: entry.raci
    });
  });

  saveRACIEntries(projectId, cleanedEntries, userEmail);
}

function validateRACIMatrix(projectId) {
  var entries = getRACIEntries(projectId);

  var result = {
    isValid: true,
    warnings: [],
    errors: [],
    summary: {
      totalTasks: 0,
      tasksByPillar: {},
      raciDistribution: {
        R: 0,
        A: 0,
        C: 0,
        I: 0
      }
    }
  };

  if (entries.length === 0) {
    result.warnings.push('No RACI entries configured.');
    return result;
  }

  var taskMap = {};

  entries.forEach(function(entry) {
    if (!taskMap[entry.task]) {
      taskMap[entry.task] = [];
    }
    taskMap[entry.task].push(entry);

    result.summary.raciDistribution[entry.raci]++;

    if (!result.summary.tasksByPillar[entry.pillar]) {
      result.summary.tasksByPillar[entry.pillar] = 0;
    }
    result.summary.tasksByPillar[entry.pillar]++;
  });

  var tasks = Object.keys(taskMap);
  result.summary.totalTasks = tasks.length;

  tasks.forEach(function(taskName) {
    var taskEntries = taskMap[taskName];

    var hasResponsible = taskEntries.some(function(e) { return e.raci === 'R'; });
    var hasAccountable = taskEntries.some(function(e) { return e.raci === 'A'; });
    var accountableCount = taskEntries.filter(function(e) { return e.raci === 'A'; }).length;

    if (!hasResponsible) {
      result.warnings.push('Task "' + taskName + '" has no R (Responsible).');
    }

    if (!hasAccountable) {
      result.warnings.push('Task "' + taskName + '" has no A (Accountable).');
    }

    if (accountableCount > 1) {
      result.errors.push('Task "' + taskName + '" has multiple A (Accountable).');
      result.isValid = false;
    }
  });

  var pillars = Object.keys(result.summary.tasksByPillar);
  if (pillars.length < 2) {
    result.warnings.push('Consider involving multiple pillars (CoE, Biz Data Network, IT).');
  }

  return result;
}

function getThreePillarTemplate() {
  return {
    pillars: [
      {
        name: THREE_PILLARS.COE,
        nameEn: 'Center of Excellence',
        description: 'Enablement, standards, and quality.',
        recommendedRoles: [
          { role: 'CoE Lead', raci: 'A', description: 'Overall governance and alignment.' },
          { role: 'Data Analyst', raci: 'R', description: 'Dashboard development and insights.' }
        ],
        typicalTasks: [
          'Dashboard development',
          'Best practice enablement'
        ]
      },
      {
        name: THREE_PILLARS.BUSINESS_DATA_NETWORK,
        nameEn: 'Business Data Network',
        description: 'Business ownership and adoption.',
        recommendedRoles: [
          { role: 'Business Owner', raci: 'A', description: 'Final approval and outcomes.' },
          { role: 'Data Champion', raci: 'R', description: 'Drive adoption within the team.' }
        ],
        typicalTasks: [
          'Requirements definition',
          'Adoption and change management'
        ]
      },
      {
        name: THREE_PILLARS.IT,
        nameEn: 'IT',
        description: 'Infrastructure, security, and integration.',
        recommendedRoles: [
          { role: 'IT Lead', raci: 'A', description: 'Platform decisions and security.' },
          { role: 'Data Engineer', raci: 'R', description: 'Data pipelines and access.' }
        ],
        typicalTasks: [
          'Data source integration',
          'Access control and security'
        ]
      }
    ],
    raciExplanation: {
      R: {
        name: 'Responsible',
        description: 'Executes the task.'
      },
      A: {
        name: 'Accountable',
        description: 'Ultimately accountable; one per task.'
      },
      C: {
        name: 'Consulted',
        description: 'Provides input before work proceeds.'
      },
      I: {
        name: 'Informed',
        description: 'Kept informed of progress.'
      }
    }
  };
}

function generateDefaultRACIEntries(projectId) {
  return [
    { pillar: THREE_PILLARS.COE, task: 'Dashboard delivery', assignee: 'TBD', raci: 'R' },
    { pillar: THREE_PILLARS.COE, task: 'Quality review', assignee: 'TBD', raci: 'A' },
    { pillar: THREE_PILLARS.BUSINESS_DATA_NETWORK, task: 'Requirements', assignee: 'TBD', raci: 'R' },
    { pillar: THREE_PILLARS.BUSINESS_DATA_NETWORK, task: 'Business approval', assignee: 'TBD', raci: 'A' },
    { pillar: THREE_PILLARS.IT, task: 'Data access', assignee: 'TBD', raci: 'R' },
    { pillar: THREE_PILLARS.IT, task: 'Security review', assignee: 'TBD', raci: 'A' }
  ];
}

function getOrganizationSummary(projectId) {
  var entries = getRACIEntries(projectId);

  if (entries.length === 0) {
    return {
      configured: false,
      totalEntries: 0,
      uniqueAssignees: 0,
      pillarCoverage: [],
      validation: null
    };
  }

  var assignees = {};
  entries.forEach(function(e) {
    assignees[e.assignee] = true;
  });

  var pillarCoverage = [];
  var pillarCounts = {};

  entries.forEach(function(e) {
    if (!pillarCounts[e.pillar]) {
      pillarCounts[e.pillar] = 0;
    }
    pillarCounts[e.pillar]++;
  });

  Object.keys(pillarCounts).forEach(function(pillar) {
    pillarCoverage.push({
      pillar: pillar,
      count: pillarCounts[pillar]
    });
  });

  return {
    configured: true,
    totalEntries: entries.length,
    uniqueAssignees: Object.keys(assignees).length,
    pillarCoverage: pillarCoverage,
    validation: validateRACIMatrix(projectId)
  };
}

function checkRACICompleteness(projectId) {
  var entries = getRACIEntries(projectId);

  var result = {
    hasCoE: false,
    hasBusinessDataNetwork: false,
    hasIT: false,
    hasAccountable: false,
    hasResponsible: false,
    completenessPercent: 0,
    suggestions: []
  };

  if (entries.length === 0) {
    result.suggestions.push('Add RACI entries to define responsibilities.');
    return result;
  }

  entries.forEach(function(e) {
    if (e.pillar === THREE_PILLARS.COE) result.hasCoE = true;
    if (e.pillar === THREE_PILLARS.BUSINESS_DATA_NETWORK) result.hasBusinessDataNetwork = true;
    if (e.pillar === THREE_PILLARS.IT) result.hasIT = true;
    if (e.raci === 'A') result.hasAccountable = true;
    if (e.raci === 'R') result.hasResponsible = true;
  });

  var score = 0;
  if (result.hasCoE) score += 20;
  if (result.hasBusinessDataNetwork) score += 20;
  if (result.hasIT) score += 20;
  if (result.hasAccountable) score += 20;
  if (result.hasResponsible) score += 20;

  result.completenessPercent = score;

  if (!result.hasCoE) {
    result.suggestions.push('Add tasks for the CoE pillar.');
  }
  if (!result.hasBusinessDataNetwork) {
    result.suggestions.push('Add tasks for Biz Data Network.');
  }
  if (!result.hasIT) {
    result.suggestions.push('Add tasks for IT.');
  }
  if (!result.hasAccountable) {
    result.suggestions.push('Add at least one Accountable (A) role.');
  }
  if (!result.hasResponsible) {
    result.suggestions.push('Add at least one Responsible (R) role.');
  }

  return result;
}
