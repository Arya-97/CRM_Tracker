/* ═══════════════════════════════════════════════════════
   COLLECTION TRACKER - PERFORMANCE OPTIMIZED
   Built for 80+ concurrent users | Smart caching
   ═══════════════════════════════════════════════════════ */

// ═══ THEME MANAGEMENT ═══
(function initTheme(){
  const theme = localStorage.getItem('crm_theme') || 'dark';
  if(theme === 'light') document.body.classList.add('light');
})();

function toggleTheme(){
  const isLight = document.body.classList.toggle('light');
  localStorage.setItem('crm_theme', isLight ? 'light' : 'dark');
  document.getElementById('theme-btn').textContent = isLight ? '☀️ Light' : '🌙 Dark';
}

document.addEventListener('DOMContentLoaded', () => {
  const theme = localStorage.getItem('crm_theme') || 'dark';
  document.getElementById('theme-btn').textContent = theme === 'light' ? '☀️ Light' : '🌙 Dark';
  
  const sidebarCollapsed = localStorage.getItem('crm_sidebar_collapsed') === 'true';
  if(sidebarCollapsed) document.querySelector('.sidebar').classList.add('collapsed');
});

// ═══ CONFIGURATION ═══
const CFG = {
  get clientId() { return HARDCODED_CLIENT_ID; },
  get sheetId()  { return HARDCODED_SHEET_ID;  },
};

// ═══ SHEET STRUCTURE ═══
const SHEETS = {
  main: 'Main Tracker',
  audit: 'Audit Log',
  users: 'Users'
};

const COLUMNS = {
  regno: 0, student_name: 1, center: 2, region: 3, newpayment_checks: 4, fees_amt: 5,
  fees_paid: 6, form_status: 7, emi1_paid_date: 8, emi2_paid_date: 9, emi3_paid_date: 10,
  emi4_paid_date: 11, total_paid_date: 12, attendance_15days: 13, last_punch_date: 14,
  ptp: 15, connected_status: 16, dialled_status: 17, yesterday_disposition: 18,
  todays_disposition: 19, other_remarks: 20
};

const NUM_COLS = 21;
const MAIN_HEADERS = ['Reg No', 'Student Name', 'Center', 'Region', 'Payment Check', 'Fees Amt', 'Fees Paid', '1st EMI Paid Date', '2nd EMI Paid Date', '3rd EMI Paid Date', '4th EMI Paid Date', 'Total Paid Date', 'Form Status', '% 15 Days Attendance', 'Last Punch Date', 'PTP', 'Connected Status', 'Dialled Status', 'Yesterday Disposition', 'Today\'s Disposition', 'Other Remarks'];

// ═══ STATE MANAGEMENT ═══
let accessToken = null;
let currentUser = null;
let allData = [];
let filteredData = [];
let auditLog = [];
let users = [];
let selectedIdx = null;
let allRegions = [];
let allCenters = [];
let regionCenterMap = {};

// ═══ PERFORMANCE OPTIMIZATION - LOCAL CACHE ═══
const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes cache
const cache = {
  data: null,
  timestamp: null,
  isValid() {
    return this.data && this.timestamp && (Date.now() - this.timestamp < CACHE_DURATION);
  },
  set(data) {
    this.data = data;
    this.timestamp = Date.now();
    // Store in sessionStorage for cross-tab sync
    try {
      sessionStorage.setItem('tracker_cache', JSON.stringify({data, timestamp: this.timestamp}));
    } catch(e) { console.warn('Cache storage failed'); }
  },
  get() {
    if(this.isValid()) return this.data;
    // Try sessionStorage
    try {
      const stored = sessionStorage.getItem('tracker_cache');
      if(stored) {
        const parsed = JSON.parse(stored);
        if(Date.now() - parsed.timestamp < CACHE_DURATION) {
          this.data = parsed.data;
          this.timestamp = parsed.timestamp;
          return this.data;
        }
      }
    } catch(e) { console.warn('Cache retrieval failed'); }
    return null;
  },
  clear() {
    this.data = null;
    this.timestamp = null;
    try { sessionStorage.removeItem('tracker_cache'); } catch(e) {}
  }
};

// ═══ AUTHENTICATION ═══
function startGoogleAuth(){
  if(!HARDCODED_CLIENT_ID || HARDCODED_CLIENT_ID === 'YOUR_CLIENT_ID_HERE'){
    showAuthErr('⚙️ Admin configuration required: Please set HARDCODED_CLIENT_ID in index.html');
    return;
  }
  if(!HARDCODED_SHEET_ID || HARDCODED_SHEET_ID === 'YOUR_SHEET_ID_HERE'){
    showAuthErr('⚙️ Admin configuration required: Please set HARDCODED_SHEET_ID in index.html');
    return;
  }
  if(!window.google){
    showAuthErr('Unable to load Google authentication. Please check your internet connection.');
    return;
  }

  google.accounts.oauth2.initTokenClient({
    client_id: HARDCODED_CLIENT_ID,
    scope: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email',
    callback: onTokenReceived,
    error_callback: err => showAuthErr('Authentication error: ' + (err.message || JSON.stringify(err))),
  }).requestAccessToken();
}

async function onTokenReceived(response){
  if(response.error){
    showAuthErr(response.error);
    return;
  }

  accessToken = response.access_token;
  showLoader('Verifying your credentials…');

  try {
    const userInfo = await gFetch('https://www.googleapis.com/oauth2/v2/userinfo').then(r => r.json());
    await ensureSheets();
    await loadUsers();
    
    const foundUser = users.find(u => u.email.toLowerCase() === userInfo.email.toLowerCase());
    
    if(!foundUser){
      if(users.length === 0){
        // First user becomes admin
        await sheetsAppend(SHEETS.users, [[userInfo.email, 'admin', 'system', nowStr()]]);
        await loadUsers();
        currentUser = { ...userInfo, role: 'admin' };
        bootApp();
      } else {
        hideLoader();
        showAuthErr('❌ Access Denied\n\nYour email (' + userInfo.email + ') is not registered in the system.\n\nPlease contact your administrator to request access.');
        accessToken = null;
      }
    } else {
      currentUser = { ...userInfo, role: foundUser.role };
      bootApp();
    }
  } catch(err) {
    hideLoader();
    showAuthErr('System error: ' + err.message);
  }
}

function bootApp(){
  hideLoader();
  document.getElementById('auth-screen').style.display = 'none';
  document.getElementById('app').style.display = 'flex';
  
  document.getElementById('user-name').textContent = currentUser.name || currentUser.email;
  document.getElementById('user-role').textContent = currentUser.role.toUpperCase();
  
  const avatar = document.getElementById('user-avatar');
  if(currentUser.picture){
    avatar.innerHTML = `<img src="${currentUser.picture}" style="width:32px;height:32px;border-radius:50%;object-fit:cover"/>`;
  } else {
    avatar.textContent = (currentUser.name || currentUser.email)[0].toUpperCase();
  }
  
  if(currentUser.role === 'admin'){
    document.querySelectorAll('.admin-nav').forEach(el => el.style.display = '');
    document.querySelectorAll('.agent-hide').forEach(el => el.style.display = '');
    
    const clientIdEl = document.getElementById('info-client-id');
    const sheetIdEl = document.getElementById('info-sheet-id');
    if(clientIdEl) clientIdEl.textContent = HARDCODED_CLIENT_ID.substring(0, 30) + '…';
    if(sheetIdEl) sheetIdEl.textContent = HARDCODED_SHEET_ID.substring(0, 30) + '…';
  } else {
    document.querySelectorAll('.admin-nav').forEach(el => el.style.display = 'none');
    document.querySelectorAll('.agent-hide').forEach(el => el.style.display = 'none');
  }
  
  loadAll();
}

function doLogout(){
  cache.clear();
  accessToken = null;
  currentUser = null;
  document.getElementById('app').style.display = 'none';
  document.getElementById('auth-screen').style.display = 'flex';
  document.getElementById('auth-err').classList.remove('show');
}

// ═══ GOOGLE SHEETS API ═══
const API_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

function gFetch(url, options = {}){
  return fetch(url, {
    ...options,
    headers: {
      'Authorization': 'Bearer ' + accessToken,
      'Content-Type': 'application/json',
      ...(options.headers || {})
    }
  });
}

async function sheetsGet(range){
  const response = await gFetch(`${API_BASE}/${CFG.sheetId}/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`);
  if(!response.ok){
    const error = await response.json();
    throw new Error(error.error?.message || 'Failed to read sheet');
  }
  return response.json();
}

async function sheetsUpdate(range, values){
  const response = await gFetch(`${API_BASE}/${CFG.sheetId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, {
    method: 'PUT',
    body: JSON.stringify({ range, majorDimension: 'ROWS', values })
  });
  if(!response.ok){
    const error = await response.json();
    throw new Error(error.error?.message || 'Failed to update sheet');
  }
  return response.json();
}

async function sheetsAppend(sheet, values){
  const response = await gFetch(`${API_BASE}/${CFG.sheetId}/values/${encodeURIComponent(sheet + '!A1')}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`, {
    method: 'POST',
    body: JSON.stringify({ majorDimension: 'ROWS', values })
  });
  if(!response.ok){
    const error = await response.json();
    throw new Error(error.error?.message || 'Failed to append to sheet');
  }
  return response.json();
}

async function sheetsBatch(requests){
  const response = await gFetch(`${API_BASE}/${CFG.sheetId}:batchUpdate`, {
    method: 'POST',
    body: JSON.stringify({ requests })
  });
  if(!response.ok){
    const error = await response.json();
    throw new Error(error.error?.message || 'Batch operation failed');
  }
  return response.json();
}

async function getSheetInfo(){
  const response = await gFetch(`${API_BASE}/${CFG.sheetId}`);
  if(!response.ok) throw new Error('Cannot read sheet. Please verify Sheet ID.');
  return response.json();
}

async function ensureSheets(){
  const info = await getSheetInfo();
  const existingSheets = info.sheets.map(s => s.properties.title);
  const sheetsToCreate = [SHEETS.main, SHEETS.audit, SHEETS.users].filter(name => !existingSheets.includes(name));
  
  if(sheetsToCreate.length){
    await sheetsBatch(sheetsToCreate.map(title => ({ addSheet: { properties: { title } } })));
    
    if(sheetsToCreate.includes(SHEETS.main))  await sheetsUpdate(`${SHEETS.main}!A1`, [MAIN_HEADERS]);
    if(sheetsToCreate.includes(SHEETS.audit)) await sheetsUpdate(`${SHEETS.audit}!A1`, [['timestamp', 'agent_email', 'agent_name', 'regno', 'student_name', 'field', 'old_value', 'new_value']]);
    if(sheetsToCreate.includes(SHEETS.users)) await sheetsUpdate(`${SHEETS.users}!A1`, [['email', 'role', 'added_by', 'added_on']]);
  }
}

// ═══ DATA LOADING WITH CACHE ═══
async function loadAll(){
  const syncBtn = document.getElementById('sync-btn');
  syncBtn.classList.add('spinning');
  showLoader('Syncing data from Google Sheets…');

  try {
    // Check cache first
    const cached = cache.get();
    if(cached && !syncBtn.classList.contains('force-refresh')){
      allData = cached.main || [];
      auditLog = cached.audit || [];
      users = cached.users || [];
      buildFilters();
      applyFilters();
      renderAudit();
      renderUsersList();
      hideLoader();
      showToast('✅ Loaded from cache (auto-syncs every 5 min)', 'info');
      syncBtn.classList.remove('spinning');
      return;
    }

    await Promise.all([
      loadMainData(),
      loadAuditData(),
      loadUsers()
    ]);
    
    // Cache the data
    cache.set({ main: allData, audit: auditLog, users: users });
    
    buildFilters();
    applyFilters();
    renderAudit();
    renderUsersList();
    
    hideLoader();
    showToast('✅ Data synced successfully', 'ok');
  } catch(err) {
    hideLoader();
    showToast('❌ Sync failed: ' + err.message, 'err');
  }

  syncBtn.classList.remove('spinning');
  syncBtn.classList.remove('force-refresh');
}

async function loadMainData(){
  const result = await sheetsGet(`${SHEETS.main}!A1:U2000`);
  const rows = result.values || [];
  
  if(rows.length <= 1){
    allData = [];
    return;
  }
  
  allData = rows.slice(1).map((row, index) => {
    const record = {};
    Object.entries(COLUMNS).forEach(([key, colIndex]) => {
      record[key] = (row[colIndex] || '').toString().trim();
    });
    record._rowIndex = index + 2;
    return record;
  }).filter(r => r.regno);
}

async function loadAuditData(){
  const result = await sheetsGet(`${SHEETS.audit}!A1:H5000`);
  const rows = result.values || [];
  
  if(rows.length <= 1){
    auditLog = [];
    return;
  }
  
  auditLog = rows.slice(1).map(r => ({
    timestamp: r[0] || '', agent_email: r[1] || '', agent_name: r[2] || '',
    regno: r[3] || '', student_name: r[4] || '', field: r[5] || '',
    old_value: r[6] || '', new_value: r[7] || ''
  })).reverse();
  
  document.getElementById('audit-count').textContent = auditLog.length;
  
  const uniqueEmails = [...new Set(auditLog.map(a => a.agent_email).filter(Boolean))];
  const auditFilterSelect = document.getElementById('audit-filter-user');
  auditFilterSelect.innerHTML = '<option value="">All Users</option>' + 
    uniqueEmails.map(email => `<option value="${email}">${email}</option>`).join('');
}

async function loadUsers(){
  const result = await sheetsGet(`${SHEETS.users}!A1:D500`);
  const rows = result.values || [];
  
  if(rows.length <= 1){
    users = [];
    return;
  }
  
  users = rows.slice(1).map(r => ({
    email: r[0] || '', role: r[1] || 'agent',
    added_by: r[2] || '', added_on: r[3] || ''
  })).filter(u => u.email);
}

// ═══ DATA WRITING ═══
async function writeRowToSheet(rowIndex, record, changes){
  const row = Array(NUM_COLS).fill('');
  Object.entries(COLUMNS).forEach(([key, colIndex]) => {
    row[colIndex] = record[key] || '';
  });
  
  await sheetsUpdate(`${SHEETS.main}!A${rowIndex}:U${rowIndex}`, [row]);
  
  if(changes.length){
    const timestamp = nowStr();
    const auditEntries = changes.map(change => [
      timestamp, currentUser.email, currentUser.name || currentUser.email,
      record.regno, record.student_name, change.field, change.old, change.new
    ]);
    await sheetsAppend(SHEETS.audit, auditEntries);
  }
  
  // Clear cache after write
  cache.clear();
}

// ═══ FILTERS & SEARCH ═══
function buildFilters(){
  allRegions = unique(allData.map(r => r.region));
  allCenters = unique(allData.map(r => r.center));
  
  regionCenterMap = {};
  allData.forEach(record => {
    if(record.region && record.center){
      if(!regionCenterMap[record.region]) regionCenterMap[record.region] = new Set();
      regionCenterMap[record.region].add(record.center);
    }
  });
  
  fillSelect('f-region', allRegions, 'All Regions');
  fillSelect('f-center', allCenters, 'All Centers');
  fillSelect('f-payment', unique(allData.map(r => r.newpayment_checks)), 'All');
}

function fillSelect(id, options, defaultLabel){
  const element = document.getElementById(id);
  const currentValue = element.value;
  
  element.innerHTML = `<option value="">${defaultLabel}</option>` + 
    options.filter(Boolean).map(opt => `<option value="${opt}">${opt}</option>`).join('');
  
  if(currentValue) element.value = currentValue;
}

function onRegionChange(){
  const selectedRegion = getValue('f-region');
  const centerSelect = document.getElementById('f-center');
  
  if(!selectedRegion){
    fillSelect('f-center', allCenters, 'All Centers');
  } else {
    const centersInRegion = regionCenterMap[selectedRegion] ? Array.from(regionCenterMap[selectedRegion]) : [];
    fillSelect('f-center', centersInRegion, 'All Centers');
  }
  
  applyFilters();
}

// Debounce for search input
let filterTimeout;
function applyFilters(){
  clearTimeout(filterTimeout);
  filterTimeout = setTimeout(() => {
    const region = getValue('f-region');
    const center = getValue('f-center');
    const payment = getValue('f-payment');
    const form = getValue('f-form');
    const connected = getValue('f-connected');
    const dialled = getValue('f-dialled');
    const search = getValue('f-search').toLowerCase();
    
    filteredData = allData.filter(record => {
      if(region && record.region !== region) return false;
      if(center && record.center !== center) return false;
      if(payment && record.newpayment_checks !== payment) return false;
      if(form && record.form_status !== form) return false;
      if(connected && record.connected_status !== connected) return false;
      if(dialled && record.dialled_status !== dialled) return false;
      if(search && !record.regno.toLowerCase().includes(search) && !record.student_name.toLowerCase().includes(search)) return false;
      return true;
    });
    
    renderTable();
    updateStats();
  }, 150); // Debounce 150ms
}

function clearFilters(){
  ['f-region', 'f-center', 'f-payment', 'f-form', 'f-connected', 'f-dialled'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('f-search').value = '';
  fillSelect('f-center', allCenters, 'All Centers');
  applyFilters();
}

function updateStats(){
  const connected = filteredData.filter(r => r.connected_status === 'Connected').length;
  const ptp = filteredData.filter(r => r.ptp).length;
  const dialled = filteredData.filter(r => r.dialled_status === 'Dialled').length;
  
  document.getElementById('s-total').textContent = filteredData.length;
  document.getElementById('s-conn').textContent = connected;
  document.getElementById('s-ptp').textContent = ptp;
  document.getElementById('s-dialled').textContent = dialled;
  document.getElementById('tw-meta').textContent = `${filteredData.length} / ${allData.length} records`;
}

// ═══ TABLE RENDERING (Optimized) ═══
function renderTable(){
  const tbody = document.getElementById('tbl-body');
  
  if(!filteredData.length){
    tbody.innerHTML = `<tr><td colspan="21"><div class="empty"><div class="empty-icon">🔍</div><div class="empty-t">No records match your filters</div></div></td></tr>`;
    return;
  }
  
  // Use DocumentFragment for better performance
  const fragment = document.createDocumentFragment();
  const tempDiv = document.createElement('div');
  
  const html = filteredData.map(record => {
    const index = allData.indexOf(record);
    const isSelected = selectedIdx === index ? 'selected' : '';
    
    return `<tr onclick="openPanel(${index})" class="${isSelected}">
      <td class="regno-cell">${record.regno}</td>
      <td class="name-cell">${record.student_name}</td>
      <td>${record.center || '—'}</td>
      <td>${badgeBuilder(record.region, 'blue')}</td>
      <td>${paymentBadge(record.newpayment_checks)}</td>
      <td>${record.fees_amt || '—'}</td>
      <td>${record.fees_paid || '—'}</td>
      <td>${record.emi1_paid_date || '—'}</td>
      <td>${record.emi2_paid_date || '—'}</td>
      <td>${record.emi3_paid_date || '—'}</td>
      <td>${record.emi4_paid_date || '—'}</td>
      <td>${record.total_paid_date || '—'}</td>
      <td>${formBadge(record.form_status)}</td>
      <td>${record.attendance_15days || '—'}</td>
      <td>${record.last_punch_date || '—'}</td>
      <td>${record.ptp ? badgeBuilder('✓', 'green') : '—'}</td>
      <td>${connectedBadge(record.connected_status)}</td>
      <td>${dialledBadge(record.dialled_status)}</td>
      <td>${dispositionBadge(record.yesterday_disposition)}</td>
      <td>${dispositionBadge(record.todays_disposition)}</td>
      <td><button class="btn btn-ghost btn-sm" onclick="event.stopPropagation();openPanel(${index})">Edit</button></td>
    </tr>`;
  }).join('');
  
  tbody.innerHTML = html;
}

const badgeBuilder = (value, color) => value ? `<span class="badge b-${color}">${value}</span>` : `<span style="color:var(--ink3)">—</span>`;

function paymentBadge(value){
  const colorMap = {'Full Paid': 'green', '50% Paid': 'yellow', 'Token Paid': 'blue', 'Less than Token': 'orange'};
  return badgeBuilder(value, colorMap[value] || 'gray');
}

function connectedBadge(value){
  const colorMap = {'Connected': 'green', 'Not Connected': 'red', 'Busy': 'yellow', 'No Answer': 'orange', 'Switched Off': 'red', 'Invalid Number': 'red'};
  return badgeBuilder(value, colorMap[value] || 'gray');
}

function dialledBadge(value){
  return badgeBuilder(value, value === 'Dialled' ? 'green' : 'gray');
}

function formBadge(value){
  const colorMap = {'Stage 1': 'blue', 'Stage 2': 'yellow', 'Stage 3': 'orange', 'Completed': 'green'};
  return badgeBuilder(value, colorMap[value] || 'gray');
}

function dispositionBadge(value){
  const colorMap = {'Interested': 'green', 'Enrolled': 'green', 'Not Interested': 'red', 'Drop': 'red', 'Follow Up': 'yellow', 'Callback Requested': 'yellow', 'RNR': 'orange', 'Wrong Number': 'red', 'Future Prospect': 'violet'};
  return badgeBuilder(value, colorMap[value] || 'gray');
}

// ═══ EDIT PANEL ═══
function openPanel(index){
  selectedIdx = index;
  const record = allData[index];
  
  document.getElementById('p-name').textContent = record.student_name;
  document.getElementById('p-regno').textContent = '#' + record.regno;
  
  document.getElementById('p-info').innerHTML = `
    <div class="info-cell"><div class="info-k">Reg No</div><div class="info-v">${record.regno}</div></div>
    <div class="info-cell"><div class="info-k">Center</div><div class="info-v">${record.center}</div></div>
    <div class="info-cell"><div class="info-k">Region</div><div class="info-v">${record.region}</div></div>
    <div class="info-cell"><div class="info-k">Payment</div><div class="info-v">${record.newpayment_checks || '—'}</div></div>
    <div class="info-cell"><div class="info-k">Fees Amount</div><div class="info-v">₹${record.fees_amt || '0'}</div></div>
    <div class="info-cell"><div class="info-k">Fees Paid</div><div class="info-v">₹${record.fees_paid || '0'}</div></div>
    <div class="info-cell"><div class="info-k">Form Status</div><div class="info-v">${formBadge(record.form_status)}</div></div>
    <div class="info-cell"><div class="info-k">Attendance (15 Days)</div><div class="info-v">${record.attendance_15days || '—'}</div></div>
    <div class="info-cell"><div class="info-k">Last Punch Date</div><div class="info-v">${record.last_punch_date || '—'}</div></div>
    <div class="info-cell"><div class="info-k">1st EMI Date</div><div class="info-v">${record.emi1_paid_date || '—'}</div></div>
    <div class="info-cell"><div class="info-k">2nd EMI Date</div><div class="info-v">${record.emi2_paid_date || '—'}</div></div>
    <div class="info-cell"><div class="info-k">3rd EMI Date</div><div class="info-v">${record.emi3_paid_date || '—'}</div></div>
    <div class="info-cell"><div class="info-k">4th EMI Date</div><div class="info-v">${record.emi4_paid_date || '—'}</div></div>
    <div class="info-cell"><div class="info-k">Total Paid Date</div><div class="info-v">${record.total_paid_date || '—'}</div></div>
  `;
  
  setValue('e-ptp', record.ptp);
  setValue('e-connected', record.connected_status);
  setValue('e-dialled', record.dialled_status);
  setValue('e-today', record.todays_disposition);
  document.getElementById('e-remarks').value = record.other_remarks || '';
  
  const history = auditLog.filter(a => a.regno === record.regno);
  const historyHtml = !history.length 
    ? '<div style="font-size:12px;color:var(--ink3);padding:8px 0">No history yet</div>'
    : history.slice(0, 10).map(audit => `
        <div class="audit-item">
          <div class="audit-head">
            <span class="audit-who">📧 ${audit.agent_name || audit.agent_email}</span>
            <span class="audit-when">${audit.timestamp}</span>
          </div>
          <div class="audit-row">
            <span class="audit-field">${audit.field}:</span>
            <span class="audit-from">${audit.old_value || '(empty)'}</span>
            <span class="audit-arr">→</span>
            <span class="audit-to">${audit.new_value}</span>
          </div>
        </div>
      `).join('');
  document.getElementById('p-history').innerHTML = historyHtml;
  
  document.getElementById('panel-ov').classList.add('open');
  document.getElementById('panel').classList.add('open');
  renderTable();
}

function closePanel(){
  document.getElementById('panel-ov').classList.remove('open');
  document.getElementById('panel').classList.remove('open');
  selectedIdx = null;
}

// ═══ SAVE CHANGES ═══
async function saveChanges(){
  if(selectedIdx === null) return;
  
  const record = allData[selectedIdx];
  const newValues = {
    ptp: document.getElementById('e-ptp').value.trim(),
    connected_status: getValue('e-connected'),
    dialled_status: getValue('e-dialled'),
    todays_disposition: getValue('e-today'),
    other_remarks: document.getElementById('e-remarks').value.trim()
  };
  
  const fieldLabels = {
    ptp: 'PTP', connected_status: 'Connected Status', dialled_status: 'Dialled Status',
    todays_disposition: "Today's Disposition", other_remarks: 'Other Remarks'
  };
  
  const changes = Object.keys(newValues)
    .filter(key => (record[key] || '') !== (newValues[key] || ''))
    .map(key => ({ key: key, field: fieldLabels[key], old: record[key] || '', new: newValues[key] }));
  
  if(!changes.length){
    showToast('No changes to save', 'info');
    return;
  }
  
  const updatedRecord = { ...record };
  
  if(newValues.todays_disposition && newValues.todays_disposition !== record.todays_disposition && record.todays_disposition){
    updatedRecord.yesterday_disposition = record.todays_disposition;
    changes.push({
      key: 'yesterday_disposition', field: "Yesterday's Disposition (auto)",
      old: record.yesterday_disposition || '', new: record.todays_disposition
    });
  }
  
  Object.assign(updatedRecord, newValues);
  
  const saveBtn = document.getElementById('save-btn');
  saveBtn.textContent = 'Saving…';
  saveBtn.disabled = true;
  showLoader('Writing to Google Sheet…');
  
  try {
    await writeRowToSheet(record._rowIndex, updatedRecord, changes);
    Object.assign(allData[selectedIdx], updatedRecord);
    
    const timestamp = nowStr();
    changes.forEach(change => {
      auditLog.unshift({
        timestamp: timestamp, agent_email: currentUser.email,
        agent_name: currentUser.name || currentUser.email, regno: record.regno,
        student_name: record.student_name, field: change.field,
        old_value: change.old, new_value: change.new
      });
    });
    
    document.getElementById('audit-count').textContent = auditLog.length;
    renderTable();
    renderAudit();
    openPanel(selectedIdx);
    hideLoader();
    showToast('✅ Changes saved successfully!', 'ok');
  } catch(err) {
    hideLoader();
    showToast('❌ Save failed: ' + err.message, 'err');
  }
  
  saveBtn.textContent = '💾 Save to Sheet';
  saveBtn.disabled = false;
}

// ═══ AUDIT LOG ═══
function renderAudit(){
  const filterUser = getValue('audit-filter-user');
  const data = filterUser ? auditLog.filter(a => a.agent_email === filterUser) : auditLog;
  const auditListEl = document.getElementById('audit-list');
  
  if(!data.length){
    auditListEl.innerHTML = '<div class="empty"><div class="empty-icon">📭</div><div class="empty-t">No audit entries yet</div></div>';
    return;
  }
  
  auditListEl.innerHTML = data.slice(0, 200).map(audit => `
    <div class="audit-item">
      <div class="audit-head">
        <span class="audit-who">📧 ${audit.agent_name || audit.agent_email}</span>
        <span class="audit-when">${audit.timestamp}</span>
      </div>
      <div class="audit-regno">Reg: ${audit.regno} — ${audit.student_name}</div>
      <div class="audit-row">
        <span class="audit-field">${audit.field}:</span>
        <span class="audit-from">${audit.old_value || '(empty)'}</span>
        <span class="audit-arr">→</span>
        <span class="audit-to">${audit.new_value}</span>
      </div>
    </div>
  `).join('');
}

// ═══ USER MANAGEMENT ═══
function renderUsersList(){
  const usersListEl = document.getElementById('users-list');
  if(!usersListEl) return;
  
  if(!users.length){
    usersListEl.innerHTML = '<div style="font-size:13px;color:var(--ink3);padding:8px 0">No users registered yet</div>';
    return;
  }
  
  usersListEl.innerHTML = users.map(user => `
    <div class="u-row">
      <span class="u-email">${user.email}</span>
      <span class="u-role-tag ur-${user.role}">${user.role.toUpperCase()}</span>
      ${currentUser?.role === 'admin' ? `<button class="btn btn-danger btn-sm" onclick="removeUser('${user.email}')">Remove</button>` : ''}
    </div>
  `).join('');
}

async function addUser(){
  const email = document.getElementById('new-email').value.trim().toLowerCase();
  const role = document.getElementById('new-role').value;
  
  if(!email || !email.includes('@')){
    showToast('Please enter a valid email address', 'err');
    return;
  }
  
  if(users.find(u => u.email === email)){
    showToast('User already exists', 'err');
    return;
  }
  
  showLoader('Adding user…');
  
  try {
    await sheetsAppend(SHEETS.users, [[email, role, currentUser.email, nowStr()]]);
    await loadUsers();
    renderUsersList();
    document.getElementById('new-email').value = '';
    cache.clear();
    hideLoader();
    showToast('✅ User added: ' + email, 'ok');
  } catch(err) {
    hideLoader();
    showToast('Error: ' + err.message, 'err');
  }
}

async function removeUser(email){
  if(!confirm('Remove user: ' + email + '?')) return;
  
  const result = await sheetsGet(`${SHEETS.users}!A:A`);
  const rows = result.values || [];
  let rowNumber = -1;
  
  for(let i = 1; i < rows.length; i++){
    if((rows[i][0] || '').toLowerCase() === email.toLowerCase()){
      rowNumber = i + 1;
      break;
    }
  }
  
  if(rowNumber < 0){
    showToast('User not found', 'err');
    return;
  }
  
  const info = await getSheetInfo();
  const sheet = info.sheets.find(s => s.properties.title === SHEETS.users);
  
  showLoader('Removing user…');
  
  try {
    await sheetsBatch([{
      deleteDimension: {
        range: { sheetId: sheet.properties.sheetId, dimension: 'ROWS', startIndex: rowNumber - 1, endIndex: rowNumber }
      }
    }]);
    
    await loadUsers();
    renderUsersList();
    cache.clear();
    hideLoader();
    showToast('User removed', 'ok');
  } catch(err) {
    hideLoader();
    showToast('Error: ' + err.message, 'err');
  }
}

// ═══ NAVIGATION ═══
function toggleSidebar(){
  const sidebar = document.querySelector('.sidebar');
  const isCollapsed = sidebar.classList.toggle('collapsed');
  localStorage.setItem('crm_sidebar_collapsed', isCollapsed ? 'true' : 'false');
}

function showPage(pageName, buttonElement){
  ['tracker', 'audit', 'settings'].forEach(page => {
    document.getElementById('page-' + page).style.display = page === pageName ? '' : 'none';
  });
  
  document.querySelectorAll('.nav-item').forEach(btn => btn.classList.remove('active'));
  if(buttonElement) buttonElement.classList.add('active');
  
  const titles = { tracker: '📊 Tracker', audit: '🕐 Audit Log', settings: '⚙️ Settings' };
  document.getElementById('page-title').textContent = titles[pageName] || pageName;
  
  if(pageName === 'audit') renderAudit();
  if(pageName === 'settings') renderUsersList();
}

// ═══ UTILITY FUNCTIONS ═══
function getValue(id){ return document.getElementById(id)?.value || ''; }
function setValue(id, value){ const el = document.getElementById(id); if(el) el.value = value || ''; }
function unique(array){ return [...new Set(array.filter(Boolean))]; }
function nowStr(){
  const date = new Date();
  return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0') + ' ' + String(date.getHours()).padStart(2, '0') + ':' + String(date.getMinutes()).padStart(2, '0') + ':' + String(date.getSeconds()).padStart(2, '0');
}

function showLoader(message){
  document.getElementById('loader-msg').textContent = message || 'Loading…';
  document.getElementById('loader').classList.add('on');
}

function hideLoader(){
  document.getElementById('loader').classList.remove('on');
}

function showAuthErr(message){
  const errorEl = document.getElementById('auth-err');
  errorEl.textContent = message;
  errorEl.classList.add('show');
}

let toastTimeout;
function showToast(message, type = 'info'){
  const toastEl = document.getElementById('toast');
  toastEl.textContent = message;
  toastEl.className = 'toast on ' + (type || '');
  
  clearTimeout(toastTimeout);
  toastTimeout = setTimeout(() => {
    toastEl.classList.remove('on');
  }, 3200);
}

// ═══ AUTO-REFRESH CACHE (Every 5 minutes) ═══
setInterval(() => {
  if(currentUser && document.getElementById('app').style.display === 'flex'){
    console.log('Auto-refreshing data cache...');
    const syncBtn = document.getElementById('sync-btn');
    syncBtn.classList.add('force-refresh');
    loadAll();
  }
}, 5 * 60 * 1000); // 5 minutes
