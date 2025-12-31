/**
 * Auth.gs
 * Authentication and access control helpers.
 */

function hasProjectAccess(projectId, userEmail) {
  userEmail = userEmail || getCurrentUserEmail();
  var project = getProject(projectId);

  if (!project) {
    return false;
  }

  if (project.createdBy === userEmail) {
    return true;
  }

  var editors = project.editors || '';
  var editorList = editors.split(',').map(function(e) { return e.trim(); });

  return editorList.indexOf(userEmail) !== -1;
}

function requireProjectAccess(projectId, userEmail) {
  userEmail = userEmail || getCurrentUserEmail();

  if (!hasProjectAccess(projectId, userEmail)) {
    throw new Error('Access denied for project: ' + projectId);
  }
}

function canEditProject(projectId, userEmail) {
  return hasProjectAccess(projectId, userEmail);
}

function addProjectEditor(projectId, newEditorEmail, currentUserEmail) {
  requireProjectAccess(projectId, currentUserEmail);

  var project = getProject(projectId);
  if (!project) {
    throw new Error('Project not found: ' + projectId);
  }

  var editors = project.editors || '';
  var editorList = editors.split(',').map(function(e) { return e.trim(); });

  if (editorList.indexOf(newEditorEmail) !== -1) {
    return;
  }

  editorList.push(newEditorEmail);

  var rowData = [];
  rowData[SCHEMA_PROJECTS.columns.PROJECT_ID] = project.projectId;
  rowData[SCHEMA_PROJECTS.columns.CUSTOMER_NAME] = project.customerName;
  rowData[SCHEMA_PROJECTS.columns.CREATED_DATE] = project.createdDate;
  rowData[SCHEMA_PROJECTS.columns.CREATED_BY] = project.createdBy;
  rowData[SCHEMA_PROJECTS.columns.EDITORS] = editorList.join(', ');
  rowData[SCHEMA_PROJECTS.columns.STATUS] = project.status;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_DATE] = getCurrentTimestamp();
  rowData[SCHEMA_PROJECTS.columns.UPDATED_BY] = currentUserEmail;

  updateRow(SCHEMA_PROJECTS.sheetName, project._rowNumber, rowData);

  logAudit(currentUserEmail, OPERATION_TYPES.UPDATE, projectId, {
    action: 'add_editor',
    newEditor: newEditorEmail
  });
}

function removeProjectEditor(projectId, editorEmailToRemove, currentUserEmail) {
  requireProjectAccess(projectId, currentUserEmail);

  var project = getProject(projectId);
  if (!project) {
    throw new Error('Project not found: ' + projectId);
  }

  if (project.createdBy === editorEmailToRemove) {
    throw new Error('Cannot remove the project owner from editors.');
  }

  var editors = project.editors || '';
  var editorList = editors.split(',').map(function(e) { return e.trim(); });

  editorList = editorList.filter(function(e) {
    return e !== editorEmailToRemove;
  });

  var rowData = [];
  rowData[SCHEMA_PROJECTS.columns.PROJECT_ID] = project.projectId;
  rowData[SCHEMA_PROJECTS.columns.CUSTOMER_NAME] = project.customerName;
  rowData[SCHEMA_PROJECTS.columns.CREATED_DATE] = project.createdDate;
  rowData[SCHEMA_PROJECTS.columns.CREATED_BY] = project.createdBy;
  rowData[SCHEMA_PROJECTS.columns.EDITORS] = editorList.join(', ');
  rowData[SCHEMA_PROJECTS.columns.STATUS] = project.status;
  rowData[SCHEMA_PROJECTS.columns.UPDATED_DATE] = getCurrentTimestamp();
  rowData[SCHEMA_PROJECTS.columns.UPDATED_BY] = currentUserEmail;

  updateRow(SCHEMA_PROJECTS.sheetName, project._rowNumber, rowData);

  logAudit(currentUserEmail, OPERATION_TYPES.UPDATE, projectId, {
    action: 'remove_editor',
    removedEditor: editorEmailToRemove
  });
}

function getCurrentUser() {
  var userEmail = getCurrentUserEmail();
  return {
    email: userEmail,
    isAuthenticated: userEmail !== ''
  };
}
