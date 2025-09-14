// ✅ FIREBASE COMPAT SDK — NO IMPORT STATEMENTS
const firebaseConfig = {
  apiKey: "AIzaSyDtAZFSvPRVaJzVGVd7xHxdIRfM1KEPruE",
  authDomain: "access-coordinator-portal.firebaseapp.com",
  projectId: "access-coordinator-portal",
  storageBucket: "access-coordinator-portal.firebasestorage.app",
  messagingSenderId: "1095681155011",
  appId: "1:1095681155011:web:3b52fc93641d8153f56776"
};

// ✅ INITIALIZE FIREBASE COMPAT
const firebaseApp = firebase.initializeApp(firebaseConfig);
const db = firebase.firestore(firebaseApp);

// 🌐 OFFLINE DETECTION
let isOnline = navigator.onLine;
const connectionStatus = document.getElementById('connectionStatus');

function updateConnectionStatus() {
  isOnline = navigator.onLine;
  if (connectionStatus) {
    if (isOnline) {
      connectionStatus.innerHTML = '<i class="fas fa-cloud"></i> Online • Syncing...';
      connectionStatus.style.color = '#28a745';
      setTimeout(syncPendingOperations, 1000);
    } else {
      connectionStatus.innerHTML = '<i class="fas fa-wifi-slash"></i> Offline • Working locally';
      connectionStatus.style.color = '#dc3545';
    }
  }
}

window.addEventListener('online', updateConnectionStatus);
window.addEventListener('offline', updateConnectionStatus);

// 🗃️ INDEXEDDB SETUP
const DB_NAME = 'AccessHubDB';
const DB_VERSION = 1;
let dbInstance = null;

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    
    request.onerror = () => reject(request.error);
    request.onsuccess = () => {
      dbInstance = request.result;
      resolve(dbInstance);
    };
    
    request.onupgradeneeded = (event) => {
      const db = event.target.result;
      if (!db.objectStoreNames.contains('applications')) {
        db.createObjectStore('applications', { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains('excel_files')) {
        db.createObjectStore('excel_files', { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains('pending_operations')) {
        db.createObjectStore('pending_operations', { keyPath: 'id', autoIncrement: true });
      }
    };
  });
}

async function saveToIndexedDB(storeName, data) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([storeName], 'readwrite');
    const store = transaction.objectStore(storeName);
    const request = store.put(data);
    
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getAllFromIndexedDB(storeName) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([storeName], 'readonly');
    const store = transaction.objectStore(storeName);
    const request = store.getAll();
    
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function deleteFromIndexedDB(storeName, id) {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction([storeName], 'readwrite');
    const store = transaction.objectStore(storeName);
    const request = store.delete(id);
    
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

// 🧠 GLOBAL STATE
let apps = [];
let currentFilter = 'all';
let searchTerm = '';
let deleteTarget = null;

// 🎯 DOM ELEMENT REFERENCES
const elements = {
  searchInput: null,
  filterAll: null,
  filterFrequent: null,
  filterWithExcel: null,
  filterRecent: null,
  btnAddApp: null,
  btnAllExcel: null,
  closeAddAppModal: null,
  closeAddExcelModal: null,
  closeAllExcelModal: null,
  btnCancelAddApp: null,
  btnCancelAddExcel: null,
  btnCancelDelete: null,
  btnSaveApp: null,
  btnSaveExcel: null,
  confirmYesBtn: null,
  btnAddExcelInput: null
};

function initElements() {
  elements.searchInput = document.getElementById('searchInput');
  elements.filterAll = document.getElementById('filterAll');
  elements.filterFrequent = document.getElementById('filterFrequent');
  elements.filterWithExcel = document.getElementById('filterWithExcel');
  elements.filterRecent = document.getElementById('filterRecent');
  elements.btnAddApp = document.getElementById('btnAddApp');
  elements.btnAllExcel = document.getElementById('btnAllExcel');
  elements.closeAddAppModal = document.getElementById('closeAddAppModal');
  elements.closeAddExcelModal = document.getElementById('closeAddExcelModal');
  elements.closeAllExcelModal = document.getElementById('closeAllExcelModal');
  elements.btnCancelAddApp = document.getElementById('btnCancelAddApp');
  elements.btnCancelAddExcel = document.getElementById('btnCancelAddExcel');
  elements.btnCancelDelete = document.getElementById('btnCancelDelete');
  elements.btnSaveApp = document.getElementById('btnSaveApp');
  elements.btnSaveExcel = document.getElementById('btnSaveExcel');
  elements.confirmYesBtn = document.getElementById('confirmYesBtn');
  elements.btnAddExcelInput = document.getElementById('btnAddExcelInput');
}

// 🎛️ ATTACH EVENT LISTENERS
function attachEventListeners() {
  initElements();

  if (!elements.btnAddApp) {
    setTimeout(attachEventListeners, 100);
    return;
  }

  // Filters
  elements.filterAll?.addEventListener('click', () => filterApps('all'));
  elements.filterFrequent?.addEventListener('click', () => filterApps('frequent'));
  elements.filterWithExcel?.addEventListener('click', () => filterApps('withExcel'));
  elements.filterRecent?.addEventListener('click', () => filterApps('recent'));

  // Modals
  elements.btnAddApp?.addEventListener('click', showAddAppModal);
  elements.btnAllExcel?.addEventListener('click', showAllExcelModal);

  // Close modals
  elements.closeAddAppModal?.addEventListener('click', () => closeModal('addAppModal'));
  elements.closeAddExcelModal?.addEventListener('click', () => closeModal('addExcelModal'));
  elements.closeAllExcelModal?.addEventListener('click', () => closeModal('allExcelModal'));
  elements.btnCancelAddApp?.addEventListener('click', () => closeModal('addAppModal'));
  elements.btnCancelAddExcel?.addEventListener('click', () => closeModal('addExcelModal'));
  elements.btnCancelDelete?.addEventListener('click', () => closeModal('confirmModal'));

  // Save & Delete
  elements.btnSaveApp?.addEventListener('click', saveApp);
  elements.btnSaveExcel?.addEventListener('click', saveExcel);
  elements.confirmYesBtn?.addEventListener('click', confirmDelete);
  elements.btnAddExcelInput?.addEventListener('click', addExcelInput);

  // Search
  elements.searchInput?.addEventListener('input', (e) => {
    searchTerm = e.target.value;
    renderApps();
  });

  // Close modals on outside click
  window.addEventListener('click', (event) => {
    if (event.target.classList.contains('modal-overlay')) {
      event.target.style.display = 'none';
    }
  });

  updateConnectionStatus();
}

// 📥 LOAD APPS — FROM FIREBASE OR INDEXEDDB
async function loadApps() {
  try {
    if (isOnline) {
      const appsSnapshot = await db.collection('applications').get();
      const appsData = [];

      for (const appDoc of appsSnapshot.docs) {
        const appData = {
          id: appDoc.id,
          ...appDoc.data()
        };

        const excelSnapshot = await db.collection('excel_files').where('appId', '==', appDoc.id).get();
        appData.excelFiles = excelSnapshot.docs.map(doc => ({
          id: doc.id,
          ...doc.data()
        }));

        await saveToIndexedDB('applications', appData);
        for (const excel of appData.excelFiles) {
          await saveToIndexedDB('excel_files', excel);
        }

        appsData.push(appData);
      }

      apps = appsData;
      renderApps();
      console.log('✅ Loaded from Firebase');
    } else {
      const appsFromDB = await getAllFromIndexedDB('applications');
      const excelFilesFromDB = await getAllFromIndexedDB('excel_files');

      const excelMap = {};
      excelFilesFromDB.forEach(excel => {
        if (!excelMap[excel.appId]) {
          excelMap[excel.appId] = [];
        }
        excelMap[excel.appId].push(excel);
      });

      apps = appsFromDB.map(app => ({
        ...app,
        excelFiles: excelMap[app.id] || []
      }));

      renderApps();
      console.log('✅ Loaded from IndexedDB (offline)');
    }
  } catch (error) {
    console.error('❌ Error loading apps:', error);
    const container = document.getElementById('appsContainer');
    if (container) {
      container.innerHTML = `
        <div class="empty-state">
          <i class="fas fa-exclamation-triangle"></i>
          <h3>Error Loading Data</h3>
          <p>Please check your connection and refresh.</p>
        </div>
      `;
    }
  }
}

// 🔄 SYNC PENDING OPERATIONS
async function syncPendingOperations() {
  if (!isOnline) return;

  const pendingOps = await getAllFromIndexedDB('pending_operations');
  if (pendingOps.length === 0) {
    if (connectionStatus) {
      connectionStatus.innerHTML = '<i class="fas fa-cloud"></i> Online • All synced';
    }
    return;
  }

  for (const op of pendingOps) {
    try {
      if (op.type === 'add_app') {
        const docRef = await db.collection('applications').add(op.data);
        await deleteFromIndexedDB('applications', op.data.id);
        op.data.id = docRef.id;
        await saveToIndexedDB('applications', op.data);
      } else if (op.type === 'update_app') {
        await db.collection('applications').doc(op.appId).update(op.data);
      } else if (op.type === 'delete_app') {
        await db.collection('applications').doc(op.appId).delete();
        const excelFiles = await getAllFromIndexedDB('excel_files');
        const appExcelFiles = excelFiles.filter(excel => excel.appId === op.appId);
        for (const excel of appExcelFiles) {
          await deleteFromIndexedDB('excel_files', excel.id);
        }
      } else if (op.type === 'add_excel') {
        const docRef = await db.collection('excel_files').add(op.data);
        await deleteFromIndexedDB('excel_files', op.data.id);
        op.data.id = docRef.id;
        await saveToIndexedDB('excel_files', op.data);
      } else if (op.type === 'update_excel') {
        await db.collection('excel_files').doc(op.excelId).update(op.data);
      } else if (op.type === 'delete_excel') {
        await db.collection('excel_files').doc(op.excelId).delete();
        await deleteFromIndexedDB('excel_files', op.excelId);
      }

      await deleteFromIndexedDB('pending_operations', op.id);
    } catch (error) {
      console.error('Error syncing:', error);
    }
  }

  await loadApps();
  
  if (connectionStatus) {
    connectionStatus.innerHTML = '<i class="fas fa-cloud"></i> Online • All synced';
  }
}

// 🖼️ RENDER APPLICATIONS (same logic as before — omitted for brevity)
// ... (keep your existing renderApps function)

function renderApps() {
  const container = document.getElementById('appsContainer');
  if (!container) return;

  container.innerHTML = '';

  let filteredApps = [...apps];

  if (searchTerm) {
    filteredApps = filteredApps.filter(app =>
      app.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      (app.desc && app.desc.toLowerCase().includes(searchTerm.toLowerCase())) ||
      app.excelFiles.some(excel =>
        excel.name.toLowerCase().includes(searchTerm.toLowerCase())
      )
    );
  }

  if (currentFilter === 'frequent') {
    const weekAgo = new Date();
    weekAgo.setDate(weekAgo.getDate() - 7);
    filteredApps = filteredApps.filter(app => {
      const lastAccess = app.accessedAt?.toDate ? app.accessedAt.toDate() : new Date(app.accessedAt || 0);
      return (app.accessCount || 0) > 3 && lastAccess > weekAgo;
    });
  } else if (currentFilter === 'withExcel') {
    filteredApps = filteredApps.filter(app => app.excelFiles.length > 0);
  } else if (currentFilter === 'recent') {
    const weekAgo = new Date();
    weekAgo.setDate(weekAgo.getDate() - 7);
    filteredApps = filteredApps.filter(app => {
      const created = app.createdAt?.toDate ? app.createdAt.toDate() : new Date(app.createdAt || 0);
      return created > weekAgo;
    });
  }

  if (filteredApps.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <i class="fas fa-folder-open"></i>
        <h3>No applications found</h3>
        <p>Start by adding your first application</p>
        <button class="btn btn-primary" id="emptyStateAddBtn">
          <i class="fas fa-plus"></i> Add New Application
        </button>
      </div>
    `;
    setTimeout(() => {
      const btn = document.getElementById('emptyStateAddBtn');
      if (btn) btn.addEventListener('click', showAddAppModal);
    }, 100);
    return;
  }

  filteredApps.forEach(app => {
    const lastAccess = app.accessedAt?.toDate ? app.accessedAt.toDate() : new Date(app.accessedAt || 0);
    const now = new Date();
    const diffHours = (now - lastAccess) / (1000 * 60 * 60);
    const isRecent = diffHours < 24;

    let excelSection = '';
    if (app.excelFiles.length > 0) {
      const excelItems = app.excelFiles.map(excel => `
        <div class="excel-item">
          <div class="excel-name">${excel.name}</div>
          <div class="excel-actions">
            <button class="btn btn-open-excel" style="padding: 8px 16px; font-size: 0.85rem;" data-url="${excel.url}">
              <i class="fas fa-external-link-alt"></i> Open
            </button>
            <button class="btn btn-outline" style="padding: 8px 16px; font-size: 0.85rem;" data-app-id="${app.id}" data-excel-id="${excel.id}">
              <i class="fas fa-edit"></i> Edit
            </button>
            <button class="btn btn-danger" style="padding: 8px 16px; font-size: 0.85rem;" data-app-id="${app.id}" data-excel-id="${excel.id}">
              <i class="fas fa-trash"></i> Delete
            </button>
            <span class="excel-status status-${excel.status}">${excel.status}</span>
          </div>
        </div>
      `).join('');

      excelSection = `
        <div class="excel-section">
          <div class="excel-header">
            <div class="excel-title">
              <i class="fas fa-file-excel"></i> Excel Files
              <span class="excel-count">${app.excelFiles.length}</span>
            </div>
            <button class="btn btn-outline" style="padding: 8px 16px; font-size: 0.9rem;" data-app-id="${app.id}" id="addExcelBtn_${app.id}">
              <i class="fas fa-plus"></i> Add
            </button>
          </div>
          <div class="excel-list">
            ${excelItems}
          </div>
        </div>
      `;
    } else {
      excelSection = `
        <div class="excel-section">
          <div class="excel-header">
            <div class="excel-title">
              <i class="fas fa-file-excel"></i> Excel Files
            </div>
          </div>
          <p style="color: var(--gray); text-align: center; padding: 20px 0; font-style: italic;">
            No Excel files associated with this application yet
          </p>
          <button class="btn btn-add-excel" data-app-id="${app.id}" id="addExcelBtn_${app.id}">
            <i class="fas fa-plus"></i> Add Excel File
          </button>
        </div>
      `;
    }

    const card = document.createElement('div');
    card.className = `app-card ${isRecent ? 'recent-pulse' : ''}`;
    card.innerHTML = `
      <div class="app-header">
        <h3 class="app-name">${app.name}</h3>
      </div>
      ${app.desc ? `<p class="app-meta">${app.desc}</p>` : ''}
      <div class="access-info-inline">
        <i class="fas fa-chart-bar"></i> Used ${app.accessCount || 0} times • Last accessed: ${formatDate(lastAccess)}
      </div>
      <div class="action-container">
        <button class="btn btn-open-app" data-app-id="${app.id}">
          <i class="fas fa-external-link-alt"></i> Open Application
        </button>
        ${excelSection}
        <div style="display: flex; gap: 12px; margin-top: 16px;">
          <button class="btn btn-outline" style="flex: 1;" data-app-id="${app.id}" id="editAppBtn_${app.id}">
            <i class="fas fa-edit"></i> Edit
          </button>
          <button class="btn btn-danger" style="flex: 1;" data-app-id="${app.id}" id="deleteAppBtn_${app.id}">
            <i class="fas fa-trash"></i> Delete
          </button>
        </div>
      </div>
    `;
    container.appendChild(card);

    setTimeout(() => {
      const openAppBtn = card.querySelector('.btn-open-app');
      if (openAppBtn) {
        openAppBtn.addEventListener('click', (e) => {
          openApp(e.currentTarget.getAttribute('data-app-id'));
        });
      }

      card.querySelectorAll('.btn-open-excel').forEach(btn => {
        btn.addEventListener('click', (e) => {
          openExcel(e.currentTarget.getAttribute('data-url'));
        });
      });

      card.querySelectorAll('.btn-outline[data-excel-id]').forEach(btn => {
        btn.addEventListener('click', (e) => {
          const appId = e.currentTarget.getAttribute('data-app-id');
          const excelId = e.currentTarget.getAttribute('data-excel-id');
          showEditExcelModal(appId, excelId);
        });
      });

      card.querySelectorAll('.btn-danger[data-excel-id]').forEach(btn => {
        btn.addEventListener('click', (e) => {
          const appId = e.currentTarget.getAttribute('data-app-id');
          const excelId = e.currentTarget.getAttribute('data-excel-id');
          requestDeleteExcel(appId, excelId);
        });
      });

      const addExcelBtn = card.querySelector(`#addExcelBtn_${app.id}`);
      if (addExcelBtn) {
        addExcelBtn.addEventListener('click', (e) => {
          showAddExcelModal(e.currentTarget.getAttribute('data-app-id'));
        });
      }

      const editAppBtn = card.querySelector(`#editAppBtn_${app.id}`);
      if (editAppBtn) {
        editAppBtn.addEventListener('click', (e) => {
          showEditAppModal(e.currentTarget.getAttribute('data-app-id'));
        });
      }

      const deleteAppBtn = card.querySelector(`#deleteAppBtn_${app.id}`);
      if (deleteAppBtn) {
        deleteAppBtn.addEventListener('click', (e) => {
          requestDeleteApp(e.currentTarget.getAttribute('data-app-id'));
        });
      }
    }, 0);
  });
}

// 🗓️ FORMAT DATE
function formatDate(date) {
  if (!date) return 'Never';
  const now = new Date();
  const diffMs = now - date;
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  const diffHours = Math.floor(diffMs / (1000 * 60 * 60));

  if (diffHours < 1) return 'Just now';
  if (diffHours < 24) return `${diffHours} hours ago`;
  if (diffDays === 1) return 'Yesterday';
  if (diffDays < 7) return `${diffDays} days ago`;
  return date.toLocaleDateString();
}

// 🌐 OPEN APPLICATION
async function openApp(appId) {
  const app = apps.find(a => a.id === appId);
  if (app) {
    if (isOnline) {
      await db.collection('applications').doc(appId).update({
        accessedAt: firebase.firestore.FieldValue.serverTimestamp(),
        accessCount: (app.accessCount || 0) + 1
      });
    } else {
      app.accessedAt = new Date().toISOString();
      app.accessCount = (app.accessCount || 0) + 1;
      await saveToIndexedDB('applications', app);
    }
    window.open(app.url, '_blank');
    if (isOnline) {
      loadApps();
    } else {
      renderApps();
    }
  }
}

// 📊 OPEN EXCEL FILE
function openExcel(url) {
  window.open(url, '_blank');
}

// 🎚️ FILTER APPLICATIONS
function filterApps(filter) {
  currentFilter = filter;
  renderApps();
}

// 🚪 CLOSE MODAL
function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) modal.style.display = 'none';
}

// 💾 SAVE APPLICATION — OFFLINE-FIRST
async function saveApp() {
  const editAppId = document.getElementById('editAppId').value;
  const name = document.getElementById('appName').value.trim();
  const url = document.getElementById('appUrl').value.trim();
  const desc = document.getElementById('appDesc').value.trim();

  if (!name || !url) {
    alert('Application name and URL are required!');
    return;
  }

  try {
    const newAppData = {
      name,
      url,
      desc,
      createdAt: isOnline ? firebase.firestore.FieldValue.serverTimestamp() : new Date().toISOString(),
      accessedAt: isOnline ? firebase.firestore.FieldValue.serverTimestamp() : new Date().toISOString(),
      accessCount: 0
    };

    if (isOnline) {
      if (editAppId) {
        await db.collection('applications').doc(editAppId).update({
          name,
          url,
          desc
        });

        const oldExcelSnapshot = await db.collection('excel_files').where('appId', '==', editAppId).get();
        const deletePromises = oldExcelSnapshot.docs.map(doc => doc.ref.delete());
        await Promise.all(deletePromises);

        const container = document.getElementById('excelFilesContainer');
        const excelDivs = container.querySelectorAll('div > div');

        for (const div of excelDivs) {
          const inputs = div.querySelectorAll('input, select');
          if (inputs.length >= 3) {
            const nameInput = inputs[0];
            const urlInput = inputs[1];
            const statusInput = inputs[2];

            if (nameInput.value.trim() && urlInput.value.trim()) {
              await db.collection('excel_files').add({
                appId: editAppId,
                name: nameInput.value.trim(),
                url: urlInput.value.trim(),
                status: statusInput.value,
                addedAt: firebase.firestore.FieldValue.serverTimestamp()
              });
            }
          }
        }
      } else {
        const docRef = await db.collection('applications').add(newAppData);

        const container = document.getElementById('excelFilesContainer');
        const excelDivs = container.querySelectorAll('div > div');

        for (const div of excelDivs) {
          const inputs = div.querySelectorAll('input, select');
          if (inputs.length >= 3) {
            const nameInput = inputs[0];
            const urlInput = inputs[1];
            const statusInput = inputs[2];

            if (nameInput.value.trim() && urlInput.value.trim()) {
              await db.collection('excel_files').add({
                appId: docRef.id,
                name: nameInput.value.trim(),
                url: urlInput.value.trim(),
                status: statusInput.value,
                addedAt: firebase.firestore.FieldValue.serverTimestamp()
              });
            }
          }
        }
      }

      alert('✅ Saved to cloud!');
    } else {
      if (editAppId) {
        const existingApp = apps.find(a => a.id === editAppId);
        if (existingApp) {
          existingApp.name = name;
          existingApp.url = url;
          existingApp.desc = desc;
          await saveToIndexedDB('applications', existingApp);

          const excelFiles = await getAllFromIndexedDB('excel_files');
          const appExcelFiles = excelFiles.filter(excel => excel.appId === editAppId);
          for (const excel of appExcelFiles) {
            await deleteFromIndexedDB('excel_files', excel.id);
          }

          const container = document.getElementById('excelFilesContainer');
          const excelDivs = container.querySelectorAll('div > div');

          for (const div of excelDivs) {
            const inputs = div.querySelectorAll('input, select');
            if (inputs.length >= 3) {
              const nameInput = inputs[0];
              const urlInput = inputs[1];
              const statusInput = inputs[2];

              if (nameInput.value.trim() && urlInput.value.trim()) {
                const newExcel = {
                  id: 'excel_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
                  appId: editAppId,
                  name: nameInput.value.trim(),
                  url: urlInput.value.trim(),
                  status: statusInput.value,
                  addedAt: new Date().toISOString()
                };
                await saveToIndexedDB('excel_files', newExcel);
              }
            }
          }
        }
      } else {
        newAppData.id = 'app_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
        await saveToIndexedDB('applications', newAppData);

        const container = document.getElementById('excelFilesContainer');
        const excelDivs = container.querySelectorAll('div > div');

        for (const div of excelDivs) {
          const inputs = div.querySelectorAll('input, select');
          if (inputs.length >= 3) {
            const nameInput = inputs[0];
            const urlInput = inputs[1];
            const statusInput = inputs[2];

            if (nameInput.value.trim() && urlInput.value.trim()) {
              const newExcel = {
                id: 'excel_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
                appId: newAppData.id,
                name: nameInput.value.trim(),
                url: urlInput.value.trim(),
                status: statusInput.value,
                addedAt: new Date().toISOString()
              };
              await saveToIndexedDB('excel_files', newExcel);
            }
          }
        }
      }

      await saveToIndexedDB('pending_operations', {
        id: 'op_' + Date.now(),
        type: editAppId ? 'update_app' : 'add_app',
        appId: editAppId || null,
        data: newAppData
      });

      alert('✅ Saved locally. Will sync when online.');
    }

    closeModal('addAppModal');
    loadApps();
  } catch (error) {
    console.error('Error saving app:', error);
    alert('Failed to save application.');
  }
}

// 💾 SAVE EXCEL FILE
async function saveExcel() {
  const appId = document.getElementById('currentAppId').value;
  const excelId = document.getElementById('editExcelId').value;
  const name = document.getElementById('excelName').value.trim();
  const url = document.getElementById('excelUrl').value.trim();
  const status = document.getElementById('excelStatus').value;

  if (!name || !url) {
    alert('Excel file name and URL are required!');
    return;
  }

  try {
    if (isOnline) {
      if (excelId) {
        await db.collection('excel_files').doc(excelId).update({
          name,
          url,
          status
        });
      } else {
        await db.collection('excel_files').add({
          appId,
          name,
          url,
          status,
          addedAt: firebase.firestore.FieldValue.serverTimestamp()
        });
      }
      alert('✅ Excel saved to cloud!');
    } else {
      if (excelId) {
        const excelFiles = await getAllFromIndexedDB('excel_files');
        const excel = excelFiles.find(e => e.id === excelId);
        if (excel) {
          excel.name = name;
          excel.url = url;
          excel.status = status;
          await saveToIndexedDB('excel_files', excel);
        }
      } else {
        const newExcel = {
          id: 'excel_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9),
          appId,
          name,
          url,
          status,
          addedAt: new Date().toISOString()
        };
        await saveToIndexedDB('excel_files', newExcel);
      }

      await saveToIndexedDB('pending_operations', {
        id: 'op_' + Date.now(),
        type: excelId ? 'update_excel' : 'add_excel',
        excelId: excelId || null,
        data: { name, url, status, appId }
      });

      alert('✅ Excel saved locally. Will sync when online.');
    }

    closeModal('addExcelModal');
    loadApps();
  } catch (error) {
    console.error('Error saving Excel:', error);
    alert('Failed to save Excel file.');
  }
}

// 🗑️ DELETE APP
async function requestDeleteApp(appId) {
  deleteTarget = { type: 'app', id: appId };
  document.getElementById('confirmMessage').textContent = 'Are you sure you want to delete this application and all its Excel files?';
  document.getElementById('confirmYesBtn').onclick = async () => {
    try {
      if (isOnline) {
        const excelSnapshot = await db.collection('excel_files').where('appId', '==', appId).get();
        const deletePromises = excelSnapshot.docs.map(doc => doc.ref.delete());
        await Promise.all(deletePromises);
        await db.collection('applications').doc(appId).delete();
      } else {
        await deleteFromIndexedDB('applications', appId);
        await saveToIndexedDB('pending_operations', {
          id: 'op_' + Date.now(),
          type: 'delete_app',
          appId: appId
        });
      }
      alert('🗑️ Application deleted!');
      loadApps();
    } catch (error) {
      console.error('Error deleting app:', error);
      alert('Failed to delete application.');
    }
    closeModal('confirmModal');
  };
  document.getElementById('confirmModal').style.display = 'flex';
}

// 🗑️ DELETE EXCEL
async function requestDeleteExcel(appId, excelId) {
  deleteTarget = { type: 'excel', appId, excelId };
  document.getElementById('confirmMessage').textContent = 'Are you sure you want to delete this Excel file?';
  document.getElementById('confirmYesBtn').onclick = async () => {
    try {
      if (isOnline) {
        await db.collection('excel_files').doc(excelId).delete();
      } else {
        await deleteFromIndexedDB('excel_files', excelId);
        await saveToIndexedDB('pending_operations', {
          id: 'op_' + Date.now(),
          type: 'delete_excel',
          excelId: excelId
        });
      }
      alert('🗑️ Excel file deleted!');
      loadApps();
    } catch (error) {
      console.error('Error deleting Excel:', error);
      alert('Failed to delete Excel file.');
    }
    closeModal('confirmModal');
  };
  document.getElementById('confirmModal').style.display = 'flex';
}

// ➕ MODAL FUNCTIONS
function showAddAppModal() {
  document.getElementById('modalTitle').textContent = 'Add New Application';
  document.getElementById('saveBtnText').textContent = 'Save Application';
  document.getElementById('editAppId').value = '';
  document.getElementById('appName').value = '';
  document.getElementById('appUrl').value = '';
  document.getElementById('appDesc').value = '';
  document.getElementById('excelFilesContainer').innerHTML = '';
  addExcelInput();
  document.getElementById('addAppModal').style.display = 'flex';
}

async function showEditAppModal(appId) {
  const app = apps.find(a => a.id === appId);
  if (!app) return;

  document.getElementById('modalTitle').textContent = 'Edit Application';
  document.getElementById('saveBtnText').textContent = 'Update Application';
  document.getElementById('editAppId').value = app.id;
  document.getElementById('appName').value = app.name;
  document.getElementById('appUrl').value = app.url;
  document.getElementById('appDesc').value = app.desc || '';
  document.getElementById('excelFilesContainer').innerHTML = '';

  app.excelFiles.forEach(excel => {
    addExcelInput(excel);
  });

  document.getElementById('addAppModal').style.display = 'flex';
}

function addExcelInput(existingExcel = null) {
  const container = document.getElementById('excelFilesContainer');
  const index = container.children.length;
  const div = document.createElement('div');
  div.style.marginBottom = '16px';

  div.innerHTML = `
    <div style="display: flex; gap: 12px; align-items: center; background: var(--light); padding: 16px; border-radius: 12px;">
      <div style="flex: 2;">
        <input type="text" placeholder="Excel file name" class="form-control" value="${existingExcel ? existingExcel.name : ''}" data-index="${index}" style="margin-bottom: 0;" />
      </div>
      <div style="flex: 3;">
        <input type="url" placeholder="Excel file URL" class="form-control" value="${existingExcel ? existingExcel.url : ''}" data-index="${index}" style="margin-bottom: 0;" />
      </div>
      <div style="flex: 1;">
        <select class="form-control" data-index="${index}" style="margin-bottom: 0;">
          <option value="active" ${existingExcel && existingExcel.status === 'active' ? 'selected' : ''}>Active</option>
          <option value="pending" ${existingExcel && existingExcel.status === 'pending' ? 'selected' : ''}>Pending</option>
          <option value="archived" ${existingExcel && existingExcel.status === 'archived' ? 'selected' : ''}>Archived</option>
        </select>
      </div>
      <button type="button" class="btn btn-danger" style="padding: 10px; width: 40px; height: 40px;">
        <i class="fas fa-trash"></i>
      </button>
    </div>
  `;
  container.appendChild(div);

  setTimeout(() => {
    const deleteBtn = div.querySelector('.btn-danger');
    if (deleteBtn) {
      deleteBtn.addEventListener('click', () => {
        div.remove();
      });
    }
  }, 0);
}

function showAddExcelModal(appId) {
  document.getElementById('excelModalTitle').textContent = 'Add Excel File';
  document.getElementById('saveExcelBtnText').textContent = 'Add Excel File';
  document.getElementById('currentAppId').value = appId;
  document.getElementById('editExcelId').value = '';
  document.getElementById('excelName').value = '';
  document.getElementById('excelUrl').value = '';
  document.getElementById('excelStatus').value = 'active';
  document.getElementById('addExcelModal').style.display = 'flex';
}

async function showEditExcelModal(appId, excelId) {
  const app = apps.find(a => a.id === appId);
  if (!app) return;
  
  const excel = app.excelFiles.find(e => e.id === excelId);
  if (!excel) return;

  document.getElementById('excelModalTitle').textContent = 'Edit Excel File';
  document.getElementById('saveExcelBtnText').textContent = 'Update Excel File';
  document.getElementById('currentAppId').value = appId;
  document.getElementById('editExcelId').value = excelId;
  document.getElementById('excelName').value = excel.name;
  document.getElementById('excelUrl').value = excel.url;
  document.getElementById('excelStatus').value = excel.status;
  document.getElementById('addExcelModal').style.display = 'flex';
}

async function showAllExcelModal() {
  const container = document.getElementById('allExcelContainer');
  if (!container) return;

  container.innerHTML = '';

  const allExcel = [];
  apps.forEach(app => {
    app.excelFiles.forEach(excel => {
      allExcel.push({
        ...excel,
        appName: app.name,
        appId: app.id
      });
    });
  });

  allExcel.sort((a, b) => {
    const dateA = a.addedAt?.toDate ? a.addedAt.toDate() : new Date(a.addedAt || 0);
    const dateB = b.addedAt?.toDate ? b.addedAt.toDate() : new Date(b.addedAt || 0);
    return dateB - dateA;
  });

  if (allExcel.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <i class="fas fa-file-excel"></i>
        <h3>No Excel files found</h3>
        <p>Add Excel files to your applications to see them here.</p>
      </div>
    `;
  } else {
    const list = allExcel.map(excel => {
      const addedDate = excel.addedAt?.toDate ? excel.addedAt.toDate() : new Date(excel.addedAt || 0);
      return `
        <div style="margin-bottom: 24px; padding: 24px; background: var(--light); border-radius: 16px; border-left: 4px solid var(--secondary);">
          <div style="display: flex; justify-content: space-between; align-items: start; flex-wrap: wrap; gap: 16px; margin-bottom: 16px;">
            <div>
              <h4 style="color: var(--dark); margin: 0 0 8px 0;">${excel.name}</h4>
              <p style="color: var(--gray); margin: 0;">
                <i class="fas fa-folder"></i> Belongs to: <strong>${excel.appName}</strong>
              </p>
              <p style="color: var(--gray); margin: 8px 0 0 0;">
                <i class="fas fa-calendar"></i> Added: ${formatDate(addedDate)}
              </p>
            </div>
            <div style="display: flex; flex-direction: column; gap: 12px; min-width: 200px;">
              <button class="btn btn-open-excel" data-url="${excel.url}">
                <i class="fas fa-external-link-alt"></i> Open File
              </button>
              <span class="excel-status status-${excel.status}">${excel.status}</span>
            </div>
          </div>
        </div>
      `;
    }).join('');

    container.innerHTML = list;

    setTimeout(() => {
      container.querySelectorAll('.btn-open-excel').forEach(btn => {
        btn.addEventListener('click', (e) => {
          openExcel(e.currentTarget.getAttribute('data-url'));
        });
      });
    }, 0);
  }

  document.getElementById('allExcelModal').style.display = 'flex';
}

function confirmDelete() {
  // Handled by requestDeleteApp / requestDeleteExcel
}

// 🚀 INITIALIZE
document.addEventListener('DOMContentLoaded', () => {
  attachEventListeners();
  loadApps();
  updateConnectionStatus();
});