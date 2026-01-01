/**
 * ProjectService.gs
 * Project domain logic.
 */

function createProjectWithValidation(customerName, userEmail) {
  if (!customerName || customerName.trim().length === 0) {
    throw new Error('Customer name is required.');
  }

  if (customerName.length > 100) {
    throw new Error('Customer name must be 100 characters or less.');
  }

  if (!userEmail) {
    throw new Error('User authentication is required.');
  }

  return createProject(customerName.trim(), userEmail);
}

function updateProjectCustomerName(projectId, customerName, userEmail) {
  requireProjectAccess(projectId, userEmail);

  if (!customerName || customerName.trim().length === 0) {
    throw new Error('Customer name is required.');
  }

  var project = getProject(projectId);
  if (!project) {
    throw new Error('Project not found.');
  }

  if (project.status === PROJECT_STATUS.ARCHIVED) {
    throw new Error('Archived projects cannot be edited.');
  }

  var rowData = [];
  rowData[SCHEMA_PROJECTS.columns.PROJECT_ID] = project.projectId;
  rowData[SCHEMA_PROJECTS.columns.CUSTOMER_NAME] = customerName.trim();
  rowData[SCHEMA_PROJECTS.columns.CREATED_DATE] = project.createdDate;
  rowData[SCHEMA_PROJECTS.columns.CREATED_BY] = project.createdBy;
  rowData[SCHEMA_PROJECTS.columns.EDITORS] = project.editors;
  rowData[SCHEMA_PROJECTS.columns.STATUS] = project.status;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_DATE] = getCurrentTimestamp();
  rowData[SCHEMA_PROJECTS.columns.UPDATED_BY] = userEmail;

  updateRow(SCHEMA_PROJECTS.sheetName, project._rowNumber, rowData);

  logAudit(userEmail, OPERATION_TYPES.UPDATE, projectId, {
    field: 'customerName',
    oldValue: project.customerName,
    newValue: customerName.trim()
  });
}

function updateProjectStatusWithValidation(projectId, newStatus, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var project = getProject(projectId);
  if (!project) {
    throw new Error('Project not found.');
  }

  var currentStatus = project.status;
  var validTransitions = {};
  validTransitions[PROJECT_STATUS.DRAFT] = [PROJECT_STATUS.CONFIRMED, PROJECT_STATUS.ARCHIVED];
  validTransitions[PROJECT_STATUS.CONFIRMED] = [PROJECT_STATUS.ARCHIVED];
  validTransitions[PROJECT_STATUS.ARCHIVED] = [];

  var allowedTransitions = validTransitions[currentStatus] || [];
  if (allowedTransitions.indexOf(newStatus) === -1) {
    throw new Error('Invalid status transition: ' + currentStatus + ' -> ' + newStatus);
  }

  if (newStatus === PROJECT_STATUS.CONFIRMED) {
    var vision = getVision(projectId);
    if (!vision || !vision.visionText) {
      throw new Error('Vision is required before confirming the project.');
    }
  }

  updateProjectStatus(projectId, newStatus, userEmail);
}

function searchProjects(userEmail, filters) {
  var projects = getUserProjects(userEmail);

  if (!filters) {
    return projects;
  }

  if (filters.status) {
    projects = projects.filter(function(p) {
      return p.status === filters.status;
    });
  }

  if (filters.customerName) {
    var searchTerm = filters.customerName.toLowerCase();
    projects = projects.filter(function(p) {
      return p.customerName.toLowerCase().indexOf(searchTerm) !== -1;
    });
  }

  if (filters.sortBy) {
    var sortField = filters.sortBy;
    var sortOrder = filters.sortOrder === 'asc' ? 1 : -1;

    projects.sort(function(a, b) {
      var aVal = a[sortField];
      var bVal = b[sortField];

      if (aVal < bVal) return -1 * sortOrder;
      if (aVal > bVal) return 1 * sortOrder;
      return 0;
    });
  }

  return projects;
}

function getProjectCompleteness(projectId) {
  var result = {
    vision: false,
    usecases: false,
    ninetyDayPlan: false,
    organization: false,
    value: false,
    overallPercent: 0
  };

  var vision = getVision(projectId);
  if (vision && vision.visionText) {
    result.vision = true;
  }

  var usecases = getUsecases(projectId);
  if (usecases && usecases.length > 0) {
    result.usecases = true;

    var hasPlans = usecases.some(function(uc) {
      var plan = getNinetyDayPlan(uc.usecaseId, projectId);
      return plan && plan.teamStructure;
    });
    result.ninetyDayPlan = hasPlans;
  }

  var raciEntries = getRACIEntries(projectId);
  if (raciEntries && raciEntries.length > 0) {
    result.organization = true;
  }

  var values = getValues(projectId);
  if (values && values.length > 0) {
    var hasValue = values.some(function(v) {
      return v.quantitativeImpact || v.qualitativeImpact;
    });
    result.value = hasValue;
  }

  var completed = 0;
  if (result.vision) completed++;
  if (result.usecases) completed++;
  if (result.ninetyDayPlan) completed++;
  if (result.organization) completed++;
  if (result.value) completed++;

  result.overallPercent = Math.round((completed / 5) * 100);

  return result;
}

function getProjectDetails(projectId, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var project = getProject(projectId);
  if (!project) {
    throw new Error('Project not found.');
  }

  return {
    project: project,
    vision: getVision(projectId),
    usecases: getUsecases(projectId),
    raciEntries: getRACIEntries(projectId),
    values: getValues(projectId),
    completeness: getProjectCompleteness(projectId)
  };
}

function archiveProject(projectId, userEmail) {
  updateProjectStatusWithValidation(projectId, PROJECT_STATUS.ARCHIVED, userEmail);
}

function duplicateProject(projectId, newCustomerName, userEmail) {
  requireProjectAccess(projectId, userEmail);

  var original = getProjectDetails(projectId, userEmail);

  var newProject = createProjectWithValidation(
    newCustomerName || original.project.customerName + ' (Copy)',
    userEmail
  );

  if (original.vision) {
    saveVision({
      projectId: newProject.projectId,
      visionText: original.vision.visionText,
      decisionRules: original.vision.decisionRules,
      successMetrics: original.vision.successMetrics,
      notes: original.vision.notes,
      userEmail: userEmail
    });
  }

  if (original.usecases) {
    original.usecases.forEach(function(uc) {
      addUsecase({
        projectId: newProject.projectId,
        challenge: uc.challenge,
        goal: uc.goal,
        expectedImpact: uc.expectedImpact,
        ninetyDayGoal: uc.ninetyDayGoal,
        score: uc.score,
        priority: uc.priority,
        userEmail: userEmail
      });
    });
  }

  if (original.raciEntries && original.raciEntries.length > 0) {
    saveRACIEntries(newProject.projectId, original.raciEntries, userEmail);
  }

  return newProject;
}
