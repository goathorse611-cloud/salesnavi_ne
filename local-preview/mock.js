/**
 * Mock for google.script.run
 * ローカルプレビュー用のAPIモック
 */

// LocalStorage keys
const STORAGE_KEYS = {
  PROJECTS: 'mock_projects',
  VISIONS: 'mock_visions',
  USECASES: 'mock_usecases',
  RACI: 'mock_raci',
  VALUES: 'mock_values',
  PLANS: 'mock_plans'
};

// Initialize storage with mock data
function initializeStorage() {
  if (!localStorage.getItem(STORAGE_KEYS.PROJECTS)) {
    localStorage.setItem(STORAGE_KEYS.PROJECTS, JSON.stringify(MOCK_PROJECTS));
  }
  if (!localStorage.getItem(STORAGE_KEYS.VISIONS)) {
    localStorage.setItem(STORAGE_KEYS.VISIONS, JSON.stringify(MOCK_VISIONS));
  }
  if (!localStorage.getItem(STORAGE_KEYS.USECASES)) {
    localStorage.setItem(STORAGE_KEYS.USECASES, JSON.stringify(MOCK_USECASES));
  }
  if (!localStorage.getItem(STORAGE_KEYS.RACI)) {
    localStorage.setItem(STORAGE_KEYS.RACI, JSON.stringify(MOCK_RACI));
  }
  if (!localStorage.getItem(STORAGE_KEYS.VALUES)) {
    localStorage.setItem(STORAGE_KEYS.VALUES, JSON.stringify(MOCK_VALUES));
  }
  if (!localStorage.getItem(STORAGE_KEYS.PLANS)) {
    localStorage.setItem(STORAGE_KEYS.PLANS, JSON.stringify(MOCK_90DAY_PLANS));
  }
}

// Storage helpers
function getStorage(key) {
  const data = localStorage.getItem(key);
  return data ? JSON.parse(data) : null;
}

function setStorage(key, data) {
  localStorage.setItem(key, JSON.stringify(data));
}

// Generate unique ID
function generateId(prefix) {
  const now = new Date();
  const dateStr = now.toISOString().slice(0, 10).replace(/-/g, '');
  const random = Math.floor(Math.random() * 10000).toString().padStart(4, '0');
  return `${prefix}-${dateStr}-${random}`;
}

// Simulate async delay
function delay(ms = 100) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Mock google.script.run
const google = {
  script: {
    run: new Proxy({}, {
      get: function(target, prop) {
        if (prop === 'withSuccessHandler') {
          return function(successCallback) {
            return new Proxy({}, {
              get: function(t, p) {
                if (p === 'withFailureHandler') {
                  return function(failureCallback) {
                    return new Proxy({}, {
                      get: function(tt, apiMethod) {
                        return async function(...args) {
                          await delay(150); // Simulate network delay
                          try {
                            const result = await MockAPI[apiMethod](...args);
                            successCallback(result);
                          } catch (error) {
                            failureCallback(error);
                          }
                        };
                      }
                    });
                  };
                }
                // Direct method call without failure handler
                return async function(...args) {
                  await delay(150);
                  try {
                    const result = await MockAPI[prop](...args);
                    successCallback(result);
                  } catch (error) {
                    console.error('API Error:', error);
                  }
                };
              }
            });
          };
        }
        return undefined;
      }
    })
  }
};

// Mock API implementations
const MockAPI = {
  // ========================================
  // Project APIs
  // ========================================
  apiGetUserProjects: function() {
    const projects = getStorage(STORAGE_KEYS.PROJECTS) || [];
    return { success: true, data: projects };
  },

  apiGetProject: function(projectId) {
    const projects = getStorage(STORAGE_KEYS.PROJECTS) || [];
    const project = projects.find(p => p.projectId === projectId);
    if (!project) {
      return { success: false, error: 'Project not found' };
    }
    return { success: true, data: project };
  },

  apiCreateProject: function(customerName) {
    const projects = getStorage(STORAGE_KEYS.PROJECTS) || [];
    const newProject = {
      projectId: generateId('PRJ'),
      customerName: customerName,
      createdDate: new Date().toISOString(),
      updatedDate: new Date().toISOString(),
      creatorEmail: MOCK_USER_EMAIL,
      editorEmails: MOCK_USER_EMAIL,
      status: '下書き'
    };
    projects.unshift(newProject);
    setStorage(STORAGE_KEYS.PROJECTS, projects);
    return { success: true, data: newProject };
  },

  apiUpdateProjectStatus: function(projectId, status) {
    const projects = getStorage(STORAGE_KEYS.PROJECTS) || [];
    const idx = projects.findIndex(p => p.projectId === projectId);
    if (idx === -1) {
      return { success: false, error: 'Project not found' };
    }
    projects[idx].status = status;
    projects[idx].updatedDate = new Date().toISOString();
    setStorage(STORAGE_KEYS.PROJECTS, projects);
    return { success: true, message: 'Status updated' };
  },

  // ========================================
  // Vision APIs (Module 1)
  // ========================================
  apiGetVision: function(projectId) {
    const visions = getStorage(STORAGE_KEYS.VISIONS) || {};
    return { success: true, data: visions[projectId] || null };
  },

  apiSaveVision: function(visionData) {
    const visions = getStorage(STORAGE_KEYS.VISIONS) || {};
    visions[visionData.projectId] = visionData;
    setStorage(STORAGE_KEYS.VISIONS, visions);

    // Update project timestamp
    const projects = getStorage(STORAGE_KEYS.PROJECTS) || [];
    const idx = projects.findIndex(p => p.projectId === visionData.projectId);
    if (idx !== -1) {
      projects[idx].updatedDate = new Date().toISOString();
      setStorage(STORAGE_KEYS.PROJECTS, projects);
    }

    return { success: true, message: 'Vision saved' };
  },

  // ========================================
  // Use Case APIs (Module 2A)
  // ========================================
  apiGetUsecases: function(projectId) {
    const usecases = getStorage(STORAGE_KEYS.USECASES) || {};
    return { success: true, data: usecases[projectId] || [] };
  },

  apiAddUsecase: function(usecaseData) {
    const usecases = getStorage(STORAGE_KEYS.USECASES) || {};
    if (!usecases[usecaseData.projectId]) {
      usecases[usecaseData.projectId] = [];
    }

    const newUsecase = {
      ...usecaseData,
      usecaseId: generateId('UC')
    };
    usecases[usecaseData.projectId].push(newUsecase);
    setStorage(STORAGE_KEYS.USECASES, usecases);

    return { success: true, data: { usecaseId: newUsecase.usecaseId } };
  },

  // ========================================
  // 90-Day Plan APIs (Module 2B)
  // ========================================
  apiGetNinetyDayPlan: function(usecaseId, projectId) {
    const plans = getStorage(STORAGE_KEYS.PLANS) || {};
    return { success: true, data: plans[usecaseId] || null };
  },

  apiSaveNinetyDayPlan: function(planData) {
    const plans = getStorage(STORAGE_KEYS.PLANS) || {};
    plans[planData.usecaseId] = planData;
    setStorage(STORAGE_KEYS.PLANS, plans);
    return { success: true, message: '90-day plan saved' };
  },

  // ========================================
  // RACI APIs (Module 4)
  // ========================================
  apiGetRACIEntries: function(projectId) {
    const raci = getStorage(STORAGE_KEYS.RACI) || {};
    return { success: true, data: raci[projectId] || [] };
  },

  apiSaveRACIEntries: function(projectId, entries) {
    const raci = getStorage(STORAGE_KEYS.RACI) || {};
    raci[projectId] = entries;
    setStorage(STORAGE_KEYS.RACI, raci);
    return { success: true, message: 'RACI entries saved' };
  },

  // ========================================
  // Value Tracking APIs (Module 7)
  // ========================================
  apiGetValues: function(projectId) {
    const values = getStorage(STORAGE_KEYS.VALUES) || {};
    return { success: true, data: values[projectId] || [] };
  },

  apiSaveValue: function(valueData) {
    const values = getStorage(STORAGE_KEYS.VALUES) || {};
    if (!values[valueData.projectId]) {
      values[valueData.projectId] = [];
    }

    const idx = values[valueData.projectId].findIndex(
      v => v.usecaseId === valueData.usecaseId
    );

    if (idx === -1) {
      values[valueData.projectId].push(valueData);
    } else {
      values[valueData.projectId][idx] = valueData;
    }

    setStorage(STORAGE_KEYS.VALUES, values);
    return { success: true, message: 'Value saved' };
  },

  // ========================================
  // Document Generation APIs
  // ========================================
  apiGenerateProposal: function(projectId) {
    // Mock - just return a fake URL
    const fakeUrl = `https://docs.google.com/document/d/mock-proposal-${projectId}`;
    alert(`[Mock] Proposal generated!\nIn production, this would create a Google Doc.\n\nMock URL: ${fakeUrl}`);
    return {
      success: true,
      data: { documentUrl: fakeUrl }
    };
  },

  // ========================================
  // Utility APIs
  // ========================================
  apiInitializeSheets: function() {
    // Reset to initial mock data
    localStorage.clear();
    initializeStorage();
    return { success: true, message: 'Storage initialized' };
  },

  apiGetCurrentUser: function() {
    return {
      success: true,
      data: { email: MOCK_USER_EMAIL }
    };
  }
};

// Initialize on load
initializeStorage();

console.log('%c[Mock API] Local preview mode active', 'color: #0176D3; font-weight: bold;');
console.log('%c[Mock API] Data is stored in localStorage. Use localStorage.clear() to reset.', 'color: #666;');
