/* ============================================================
   FOLIO — APP.JS v1.0.0
   Moody's Financial Reporting Tracker
   ============================================================
   SETUP: Fill in CONFIG values before first deployment.
   See Section 4 of the Build Guide for details.
   ============================================================ */

'use strict';

// ============================================================
// CONFIGURATION — FILL IN YOUR VALUES HERE
// ============================================================
const CONFIG = {
  // Azure App Registration
  clientId:    'bb00291f-d451-4e74-b8cf-10c334efb0ed',
  tenantId:    '1061a8b8-b1ee-4249-bb84-9a2cd2792fae',
  redirectUri: 'https://reneebrynczka.github.io/financial-reporting-tracker/',  // Full URL to index.html

  // SharePoint Site
  siteUrl: 'https://moodys.sharepoint.com/sites/finance_home_finrptg',

  // SharePoint List Names — must match exactly
  lists: {
    taskTemplates:        'TaskTemplates',
    quarterlyAssignments: 'QuarterlyAssignments',
    closeCalendar:        'CloseCalendar',
    appSettings:          'AppSettings',
    users:                'Users',
    auditLog:             'AuditLog',
    taskSuggestions:      'TaskSuggestions',
    matrixStatus:         'MatrixStatus',
    reviewComments:       'ReviewComments',
    reviewCommentReplies: 'ReviewCommentReplies',
  },

  // App Settings
  version:         '1.0.0',
  pollIntervalMs:  60000,           // 60 seconds — balances freshness vs API call volume
  timezone:        'America/New_York',
  verboseLogging:  false,           // Set true temporarily to debug — logs all API calls to browser console

  // Matrix checkpoints (order matters — defines column order)
  matrixCheckpoints: [
    'Prepared in Workiva',
    '1st Review Workiva',
    'Tie-out',
    '1st Review Tie-out',
    'XBRL',
    '1st Review XBRL',
    'SP Preparer',
    'SP 1st Reviewer',
    'Loaded to Clara',
    'Final Review',
  ],

  // Matrix-only columns (not tied to tasks)
  matrixOnlyColumns: ['SP Preparer', 'SP 1st Reviewer', 'Loaded to Clara', 'Final Review'],

  // User emoji options
  emojiOptions: ['🦊','⭐','💜','🌊','🦋','🔥','🎯','🚀','🎨','🌙','☀️','🐬','🦅','💎','🎵','🌺','🦁','🐋','🌻','🦄','🎸','🔮','🍀','🐝','🦉','🌴','🎲','⚡','🐧','🐶','🧁','🍓','🍦','🎈','🪅','✈️','🧸','🧢'],

  // User color options
  colorOptions: [
    { hex: '#F5A623', label: 'Amber' },
    { hex: '#00897B', label: 'Teal' },
    { hex: '#7B61FF', label: 'Purple' },
    { hex: '#29ABE2', label: 'Sky' },
    { hex: '#3AB54A', label: 'Green' },
    { hex: '#E91E8C', label: 'Rose' },
    { hex: '#75787B', label: 'Slate' },
    { hex: '#D4537E', label: 'Pink' },
    { hex: '#FF7043', label: 'Tangerine' },
    { hex: '#26C6DA', label: 'Cyan' },
    { hex: '#5C6BC0', label: 'Indigo' },
    { hex: '#8D6E63', label: 'Brown' },
    { hex: '#EC407A', label: 'Fuchsia' },
    { hex: '#66BB6A', label: 'Mint' },
    { hex: '#AB47BC', label: 'Violet' },
    { hex: '#EF5350', label: 'Red' },
    { hex: '#26A69A', label: 'Sea' },
    { hex: '#BDBDBD', label: 'Silver' },
    { hex: '#F5BCDD', label: 'Light Pink' },
    { hex: '#D0BCF5', label: 'Lilac' },
  ],
};

// ============================================================
// CONSTANTS
// ============================================================

// Single source of truth for matrix-only column -> SharePoint field mapping.
// Used by performMatrixUpdate, renderMatrixView, and exportMatrixExcel.
const MATRIX_FIELD_MAP = {
  'SP Preparer':     { status: 'SPPreparer',    date: 'SPPreparerDate',    by: 'SPPreparerBy'    },
  'SP 1st Reviewer': { status: 'SP1stReviewer', date: 'SP1stReviewerDate', by: 'SP1stReviewerBy' },
  'Loaded to Clara': { status: 'LoadedToClara', date: 'LoadedToClaraDate', by: 'LoadedToClaraBy' },
  'Final Review':    { status: 'FinalReview',   date: 'FinalReviewDate',   by: 'FinalReviewBy'   },
};

// ============================================================
// MSAL SETUP
// ============================================================
const msalConfig = {
  auth: {
    clientId:    CONFIG.clientId,
    authority:   `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri,
  },
  cache: {
    cacheLocation:          'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

const loginRequest = { scopes: ['User.Read', 'Sites.ReadWrite.All'] };

let msalInstance = null;
let currentAccount = null;

// ============================================================
// APP STATE
// ============================================================
const STATE = {
  currentUser:    null,       // User object from Users list
  activeQuarter:  null,       // e.g. "Q2 2026" — set by admin, never changed by users
  viewingQuarter: null,       // Quarter currently browsed — equals activeQuarter unless viewing history
  workingQuarter: null,       // Admin staging quarter
  assignments:    [],         // QuarterlyAssignments for active quarter
  templates:      [],         // TaskTemplates (cached)
  users:          [],         // Users list (cached)
  calendar:       [],         // CloseCalendar for active quarter
  matrixStatus:   [],         // MatrixStatus for active quarter
  reviewComments: [],         // ReviewComments for active quarter
  rcReplies:      [],         // ReviewCommentReplies for active quarter
  currentView:    'my-tasks',
  pollTimer:      null,
  siteId:         null,       // SharePoint site ID (auto-populated)
  isAdmin:        false,
  isFinalReviewer: false,
  taskDetailId:   null,       // Currently open task panel assignment ID
  filters: {
    status:    'all',
    category:  'all',
    assignee:  'all',
    search:    '',
    rcStatus:  'all',
    rcPriority:'all',
    rcQuarter: 'all',
    sort:      'overdue', // 'overdue' | 'category' | 'prepWD' | 'revWD' | 'status' | 'task'
    sortDir:   'asc',     // 'asc' | 'desc'
  },
  pendingMatrixAction: null,  // {item, column, quarter} — pending matrix update confirmation
  pendingSignoff: null,       // {assignmentId, role}
  pendingReversal: null,      // {assignmentId, role}
  pendingActivation: null,    // quarter name
  pendingRCResolve:       null,   // review comment ID
  pendingSuggestionReject: null,  // suggestion ID pending rejection
  pendingTemplateEdit:    null,   // template ID being edited
  pendingTemplateRetire:  null,   // template ID pending retire confirm
  pendingReassign:        null,   // {assignmentId, role} pending reassignment
  _stagingItems:          [],     // Cached staging assignments for rollforward grid
  _stagingLoading:        false,  // True while staging items are being fetched
  _addUserEmoji:          null,   // Emoji selected in Add User modal
  _addUserColor:          null,   // Color selected in Add User modal
  pendingCalendarEdit:    null,   // calendar row ID being edited
  pendingUserEdit:        null,   // user email being edited
  suggestions:            [],         // TaskSuggestions (loaded when admin panel opens)
  pendingCascade:      null,  // {quarter, wdNumber, shift, newDate}
  pendingRollforward:  null,  // quarter name awaiting rollforward confirm
  _auditEntries:       [],    // Loaded on-demand when audit log panel opens
  _auditFilter:        { type: 'All', person: '', quarter: '' }, // Audit log filter state
};

// ============================================================
// LOGGING
// ============================================================
function log(...args) {
  if (CONFIG.verboseLogging) console.log('[Folio]', ...args);
}
function logError(...args) {
  console.error('[Folio ERROR]', ...args);
}

// ============================================================
// UTILITY — EASTERN TIME
// ============================================================
function nowET() {
  return new Date(new Date().toLocaleString('en-US', { timeZone: CONFIG.timezone }));
}

function todayET() {
  const d = nowET();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function formatDateET(isoString) {
  if (!isoString) return '—';
  try {
    const d = new Date(isoString);
    return d.toLocaleString('en-US', {
      timeZone: CONFIG.timezone,
      month: 'short', day: 'numeric', year: 'numeric',
      hour: 'numeric', minute: '2-digit', hour12: true
    }) + ' ET';
  } catch { return isoString; }
}

function formatDateShort(isoString) {
  if (!isoString) return '—';
  try {
    const d = new Date(isoString + 'T12:00:00');
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  } catch { return isoString; }
}

// Returns the CSS class for a milestone pill based on MilestoneType.
// Falls back gracefully for legacy IsCustomMilestone boolean rows.
function milestoneClass(calRow) {
  const t = calRow.MilestoneType;
  if (t === 'SVP')      return 'milestone-svp';
  if (t === 'MD')       return 'milestone-md';
  if (t === 'CFO')      return 'milestone-cfo';
  if (t === 'Standard') return 'milestone-std';
  // Legacy fallback: IsCustomMilestone = true → SVP style
  if (calRow.IsCustomMilestone) return 'milestone-svp';
  return 'milestone-std';
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function isQuarterQ4(quarter) {
  return quarter && quarter.trim().toUpperCase().startsWith('Q4');
}

function getMaxWorkday(quarter) {
  return isQuarterQ4(quarter) ? 35 : 20;
}

// Returns which sign-off role a matrix checkpoint represents.
// '1st Review *' checkpoints map to the reviewer; everything else is the preparer.
function getCheckpointRole(checkpoint) {
  return checkpoint.startsWith('1st Review') ? 'reviewer' : 'preparer';
}

// Returns the sign-off field names for a given role on an assignment.
function getSignOffFields(role) {
  const isPreparer = role === 'preparer';
  return {
    signOff:    isPreparer ? 'PreparerSignOff'     : 'ReviewerSignOff',
    signOffDate:isPreparer ? 'PreparerSignOffDate'  : 'ReviewerSignOffDate',
    signOffBy:  isPreparer ? 'PreparerSignOffBy'    : 'ReviewerSignOffBy',
    assignee:   isPreparer ? 'Preparer'             : 'Reviewer',
    workday:    isPreparer ? 'PreparerWorkday'       : 'ReviewerWorkday',
  };
}

// ============================================================
// UTILITY — WORKDAY RESOLUTION (single source of truth)
// ============================================================
function resolveWorkday(quarter, wdNumber) {
  const entry = STATE.calendar.find(c =>
    c.Quarter === quarter && Number(c.WorkdayNumber) === Number(wdNumber)
  );
  return entry ? entry.ActualDate : null;
}

function getTodaysWorkday(quarter) {
  const today = todayET();
  const sorted = [...STATE.calendar]
    .filter(c => c.Quarter === quarter)
    .sort((a,b) => Number(a.WorkdayNumber) - Number(b.WorkdayNumber));
  if (!sorted.length) return null;
  const match = sorted.find(c => c.ActualDate === today);
  if (match) return Number(match.WorkdayNumber);
  if (today < sorted[0].ActualDate) return 'pre-close';
  if (today > sorted[sorted.length-1].ActualDate) return 'post-close';
  const prev = sorted.filter(c => c.ActualDate < today).pop();
  const next = sorted.find(c => c.ActualDate > today);
  if (prev && next) return `Between WD${prev.WorkdayNumber} and WD${next.WorkdayNumber}`;
  return null;
}

// Returns the workday number for the next calendar workday after today,
// or null if today is the last workday or the calendar has no future entries.
function getTomorrowWorkday(quarter) {
  const today = todayET();
  const sorted = [...STATE.calendar]
    .filter(c => c.Quarter === quarter)
    .sort((a, b) => Number(a.WorkdayNumber) - Number(b.WorkdayNumber));
  const next = sorted.find(c => c.ActualDate > today);
  return next ? Number(next.WorkdayNumber) : null;
}

function getWDIndicatorText(quarter) {
  if (!quarter) return '—';
  const wd = getTodaysWorkday(quarter);
  if (wd === null) return quarter;
  if (wd === 'pre-close') return `Pre-close · ${quarter}`;
  if (wd === 'post-close') return `Post-close · ${quarter}`;
  if (typeof wd === 'string') return wd; // "Between WD3 and WD4"
  const date = resolveWorkday(quarter, wd);
  const dateStr = date ? formatDateShort(date) : '';
  return `WD${wd}${dateStr ? ' · ' + dateStr : ''}`;
}

function isTaskOverdue(assignment) {
  const wd = getTodaysWorkday(STATE.activeQuarter);
  if (!wd || typeof wd !== 'number') return false;
  const role = assignment.SignOffMode === 'Preparer Only' ? 'preparer' :
    !assignment.PreparerSignOff ? 'preparer' : 'reviewer';
  const dueWD = role === 'preparer'
    ? Number(assignment.PreparerWorkday)
    : Number(assignment.ReviewerWorkday);
  return wd > dueWD && assignment.Status !== 'Complete';
}

// ============================================================
// GRAPH API — CENTRAL REQUEST HANDLER
// ============================================================
async function getToken() {
  if (!msalInstance || !currentAccount) throw new Error('Not authenticated');
  try {
    const result = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account: currentAccount,
    });
    return result.accessToken;
  } catch (err) {
    log('Silent token failed, redirecting...', err);
    await msalInstance.acquireTokenRedirect(loginRequest);
    throw err;
  }
}

async function graphRequest(method, endpoint, body = null, retries = 3) {
  const url = endpoint.startsWith('https://')
    ? endpoint
    : `https://graph.microsoft.com/v1.0${endpoint}`;

  log(`${method} ${endpoint}`);

  const bodyStr = body ? JSON.stringify(body) : null;

  for (let attempt = 1; attempt <= retries; attempt++) {
    // Re-fetch token on every attempt so a long backoff sleep never uses a stale token.
    const token = await getToken();
    const options = {
      method,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    };
    if (bodyStr) options.body = bodyStr;

    try {
      const res = await fetch(url, options);
      if (res.status === 429) {
        const retryAfter = parseInt(res.headers.get('Retry-After') || '5', 10);
        log(`Throttled. Retrying after ${retryAfter}s...`);
        await sleep(retryAfter * 1000 * attempt);
        continue;
      }
      if (!res.ok) {
        const errText = await res.text();
        throw new Error(`Graph API ${res.status}: ${errText}`);
      }
      if (res.status === 204) return null;
      return await res.json();
    } catch (err) {
      if (attempt === retries) throw err;
      await sleep(1000 * attempt);
    }
  }
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ============================================================
// SHAREPOINT — SITE ID
// ============================================================
async function getSiteId() {
  if (STATE.siteId) return STATE.siteId;
  // Strip trailing slash so the Graph API path is always well-formed
  const url = CONFIG.siteUrl.replace(/\/$/, '').replace('https://', '').replace('.sharepoint.com', '');
  const parts = url.split('/sites/');
  const hostname = parts[0] + '.sharepoint.com';
  const sitePath = 'sites/' + parts[1];
  const data = await graphRequest('GET', `/sites/${hostname}:/${sitePath}`);
  STATE.siteId = data.id;
  log('Site ID:', STATE.siteId);
  return STATE.siteId;
}

// ============================================================
// SHAREPOINT — LIST OPERATIONS
// ============================================================
// Fetches all items from a SharePoint list, following @odata.nextLink pages until
// the full result set is returned. Handles lists of any size including AuditLog.
async function getListItems(listName, filter = '', select = '', expand = '') {
  const siteId = await getSiteId();
  let url = `/sites/${siteId}/lists/${listName}/items?$top=500&$expand=fields`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;
  if (select) url += `&$select=${encodeURIComponent(select)}`;
  if (expand) url += `&$expand=${encodeURIComponent(expand)}`;

  const allItems = [];
  let pageCount = 0;

  while (url) {
    const data = await graphRequest('GET', url);
    const page = data.value || [];
    allItems.push(...page);
    pageCount++;

    // Follow the next page link if present, otherwise stop.
    url = data['@odata.nextLink'] || null;

    log(`${listName}: fetched page ${pageCount} (${page.length} items, ${allItems.length} total)`);
  }

  return allItems;
}

async function createListItem(listName, fields) {
  const siteId = await getSiteId();
  return graphRequest('POST', `/sites/${siteId}/lists/${listName}/items`, { fields });
}

async function updateListItem(listName, itemId, fields) {
  const siteId = await getSiteId();
  return graphRequest('PATCH', `/sites/${siteId}/lists/${listName}/items/${itemId}/fields`, fields);
}

async function getAppSetting(key) {
  const items = await getListItems(CONFIG.lists.appSettings, `fields/Title eq '${key}'`);
  if (items.length) return items[0].fields.SettingValue;
  return null;
}

async function setAppSetting(key, value) {
  const items = await getListItems(CONFIG.lists.appSettings, `fields/Title eq '${key}'`);
  if (items.length) {
    await updateListItem(CONFIG.lists.appSettings, items[0].id, { SettingValue: value });
  } else {
    await createListItem(CONFIG.lists.appSettings, { Title: key, SettingValue: value });
  }
}

// ============================================================
// AUDIT LOG
// ============================================================
async function writeAuditLog(actionType, details) {
  try {
    await createListItem(CONFIG.lists.auditLog, {
      Title: `${actionType}: ${details.taskName || details.description || ''}`,
      Quarter:       STATE.activeQuarter || '',
      ActionType:    actionType,
      ActionBy:      STATE.currentUser?.Email || '',
      ActionDate:    new Date().toISOString(),
      WorkdayNumber: (() => { const w = getTodaysWorkday(STATE.activeQuarter); return typeof w === 'number' ? w : 0; })(),
      TaskName:      details.taskName || '',
      AssignmentID:  details.assignmentId || null,
      PreviousValue: details.previousValue || '',
      NewValue:      details.newValue || '',
      ReasonNote:    details.reason || '',
    });
  } catch (err) {
    logError('Failed to write audit log:', err);
  }
}

// ============================================================
// DATA LOADING
// ============================================================
async function loadActiveQuarter() {
  STATE.activeQuarter  = await getAppSetting('ActiveQuarter');
  STATE.workingQuarter = await getAppSetting('WorkingQuarter');
  // viewingQuarter starts equal to activeQuarter on every login.
  // It diverges only when the user browses a historical quarter.
  STATE.viewingQuarter = STATE.activeQuarter;
  log('Active quarter:', STATE.activeQuarter);
}

async function loadCurrentUser(email) {
  const items = await getListItems(CONFIG.lists.users, `fields/Email eq '${email}'`);
  if (items.length) {
    STATE.currentUser = items[0].fields;
    STATE.currentUser._id = items[0].id;
  } else {
    // First login — create user record
    const created = await createListItem(CONFIG.lists.users, {
      Title: email.split('@')[0],
      Email: email,
      Role: 'TeamMember',
      IsActive: true,
      NotifyOnAssignment:       false,
      NotifyOnReviewUnlock:     false,
      NotifyOnOverdue:          false,
      NotifyOnReassignment:     false,
      NotifyOnSuggestionUpdate: false,
    });
    STATE.currentUser = { ...created.fields, _id: created.id };
    // Set role flags for first-login users too (Role defaults to TeamMember on creation)
    STATE.isAdmin = STATE.currentUser.Role === 'Admin';
    STATE.isFinalReviewer = STATE.currentUser.Role === 'FinalReviewer' || STATE.isAdmin;
    return false; // First login
  }
  STATE.isAdmin = STATE.currentUser.Role === 'Admin';
  STATE.isFinalReviewer = STATE.currentUser.Role === 'FinalReviewer' || STATE.isAdmin;
  return true; // Returning user
}

async function loadUsers() {
  const items = await getListItems(CONFIG.lists.users, `fields/IsActive eq true`);
  STATE.users = items.map(i => ({ ...i.fields, _id: i.id }));
}

async function loadTemplates() {
  const items = await getListItems(CONFIG.lists.taskTemplates, `fields/IsActive eq true`);
  STATE.templates = items.map(i => ({ ...i.fields, _id: i.id }));
  log('Templates loaded:', STATE.templates.length);
}

async function loadAssignments(quarter) {
  const items = await getListItems(
    CONFIG.lists.quarterlyAssignments,
    `fields/Quarter eq '${quarter}' and fields/IsStaging eq false`
  );
  STATE.assignments = items.map(i => ({ ...i.fields, _id: i.id }));
  log('Assignments loaded:', STATE.assignments.length);
}

async function loadCalendar(quarter) {
  const items = await getListItems(CONFIG.lists.closeCalendar, `fields/Quarter eq '${quarter}'`);
  STATE.calendar = items.map(i => ({ ...i.fields, _id: i.id }))
    .sort((a,b) => Number(a.WorkdayNumber) - Number(b.WorkdayNumber));
}

async function loadMatrixStatus(quarter) {
  const items = await getListItems(CONFIG.lists.matrixStatus, `fields/Quarter eq '${quarter}'`);
  STATE.matrixStatus = items.map(i => ({ ...i.fields, _id: i.id }));
}

async function loadSuggestions() {
  const items = await getListItems(CONFIG.lists.taskSuggestions);
  STATE.suggestions = items.map(i => ({ ...i.fields, _id: i.id }));
  log('Suggestions loaded:', STATE.suggestions.length);
}

async function loadReviewComments(quarter) {
  const items = await getListItems(CONFIG.lists.reviewComments, `fields/Quarter eq '${quarter}'`);
  STATE.reviewComments = items.map(i => ({ ...i.fields, _id: i.id }));
}

async function loadRCReplies() { // quarter param removed — filtering done client-side against loaded comment IDs
  // Replies don't have a Quarter field — we filter client-side by matching against
  // parent comment IDs already loaded for this quarter.
  // KNOWN SCALING ISSUE: this fetches ALL replies across all quarters then discards
  // those not belonging to the current quarter. This is safe at current scale
  // (replies are rare and small) but will become expensive once reply counts grow.
  // Future fix: add a Quarter column to ReviewCommentReplies and filter server-side.
  const rcIds = new Set(STATE.reviewComments.map(rc => rc._id));
  if (!rcIds.size) { STATE.rcReplies = []; return; }
  const items = await getListItems(CONFIG.lists.reviewCommentReplies);
  STATE.rcReplies = items
    .map(i => ({ ...i.fields, _id: i.id }))
    .filter(r => rcIds.has(r.ReviewCommentLookupId));
}

async function loadAllData() {
  if (!STATE.activeQuarter) return;
  // Load review comments first so loadRCReplies can filter by the loaded comment IDs.
  await Promise.all([
    loadAssignments(STATE.activeQuarter),
    loadCalendar(STATE.activeQuarter),
    loadMatrixStatus(STATE.activeQuarter),
    loadReviewComments(STATE.activeQuarter),
    loadUsers(),
  ]);
  // Replies depend on STATE.reviewComments being populated, so load sequentially after.
  await loadRCReplies();
}

// Returns true when the user is browsing a historical quarter.
function isViewingHistory() {
  return STATE.viewingQuarter && STATE.viewingQuarter !== STATE.activeQuarter;
}

// Use these helpers everywhere quarter context matters:
// getReadQuarter()  — the quarter currently being displayed (may be historical)
// getWriteQuarter() — the live quarter all writes must target (never historical)
function getReadQuarter()  { return getReadQuarter(); }
function getWriteQuarter() { return STATE.activeQuarter; }

// Loads all data for the viewing quarter (historical or live).
// Does NOT touch STATE.activeQuarter — write operations always use the live quarter.
async function loadViewingQuarterData(quarter) {
  if (!quarter) return;
  await Promise.all([
    loadAssignments(quarter),
    loadCalendar(quarter),
    loadMatrixStatus(quarter),
    loadReviewComments(quarter),
    loadUsers(),
  ]);
  await loadRCReplies();
}

// Switches the viewing context to a different quarter and re-renders.
async function switchToQuarter(quarter) {
  if (quarter === STATE.viewingQuarter) return;

  showLoading(`Loading ${quarter}...`);
  try {
    STATE.viewingQuarter = quarter;
    await loadViewingQuarterData(quarter);
    updateHistoryBanner();
    updateWDIndicator();
    refreshCurrentView();
  } catch (err) {
    showToast(`Failed to load ${quarter}`, 'error');
    logError('switchToQuarter failed:', err);
    // Revert to active quarter on failure
    STATE.viewingQuarter = STATE.activeQuarter;
    updateHistoryBanner();
  }
  hideLoading();
}

// Shows or hides the history banner and updates its text.
function updateHistoryBanner() {
  const banner = document.getElementById('history-banner');
  if (!banner) return;
  if (isViewingHistory()) {
    banner.classList.remove('hidden');
    const label = banner.querySelector('#history-banner-label');
    if (label) label.textContent = `Viewing ${STATE.viewingQuarter} — read only`;
  } else {
    banner.classList.add('hidden');
  }
}

// Populates the quarter picker dropdown with all quarters that have assignment data.
async function populateQuarterPicker() {
  const sel = document.getElementById('quarter-picker');
  if (!sel) return;

  // Fetch distinct quarters from QuarterlyAssignments.
  // We use a small select to avoid loading all items — just fetch the field.
  try {
    const items = await getListItems(CONFIG.lists.quarterlyAssignments, '', 'fields/Quarter', '');
    const quarters = [...new Set(
      items.map(i => i.fields?.Quarter).filter(Boolean)
    )].sort().reverse(); // newest first

    const current = getReadQuarter();
    sel.innerHTML = quarters.map(q =>
      `<option value="${escapeHtml(q)}" ${q === current ? 'selected' : ''}>${escapeHtml(q)}${q === STATE.activeQuarter ? ' (live)' : ''}</option>`
    ).join('');

    if (!sel.dataset.listenerAttached) {
      sel.dataset.listenerAttached = 'true';
      sel.addEventListener('change', () => switchToQuarter(sel.value));
    }
  } catch (err) {
    logError('populateQuarterPicker failed:', err);
  }
}

// ============================================================
// POLL — SILENT BACKGROUND REFRESH
// ============================================================
function startPolling() {
  stopPolling();
  STATE.pollTimer = setInterval(async () => {
    // Skip poll when the tab is hidden — no point refreshing data nobody can see,
    // and it avoids unnecessary Graph API calls when users leave Folio open overnight.
    if (document.hidden) return;

    // Skip poll when there's no active quarter — nothing to refresh.
    if (!STATE.activeQuarter) return;

    try {
      await loadAllData();
      refreshCurrentView();
      updateWDIndicator();
      populateQuarterPicker();
      showStaleBanner(false);
    } catch (err) {
      logError('Poll failed:', err);
      showStaleBanner(true);
    }
  }, CONFIG.pollIntervalMs);

  // When the tab becomes visible again after being hidden, do an immediate
  // refresh so data is never stale when the user returns to Folio.
  if (!document._folioVisibilityListenerAdded) {
    document._folioVisibilityListenerAdded = true;
    document.addEventListener('visibilitychange', () => {
      if (!document.hidden && STATE.activeQuarter) {
        loadAllData()
          .then(() => { refreshCurrentView(); updateWDIndicator(); showStaleBanner(false); })
          .catch(() => showStaleBanner(true));
      }
    });
  }
}

function stopPolling() {
  if (STATE.pollTimer) clearInterval(STATE.pollTimer);
  STATE.pollTimer = null;
}

// ============================================================
// SIGN-OFF
// ============================================================
async function performSignOff(assignmentId, role) {
  const assignment = STATE.assignments.find(a => a._id === assignmentId);
  if (!assignment) return;

  const f = getSignOffFields(role);
  const now = new Date().toISOString();
  const userEmail = STATE.currentUser.Email;

  const fields = {
    [f.signOff]:    true,
    [f.signOffDate]:now,
    [f.signOffBy]:  userEmail,
    Status: role === 'preparer' && assignment.SignOffMode !== 'Preparer Only' ? 'Prepared' : 'Complete',
  };

  // Snapshot the fields we are about to overwrite so we can restore them exactly on failure.
  const snapshot = {};
  Object.keys(fields).forEach(k => { snapshot[k] = assignment[k]; });

  // Optimistic update
  Object.assign(assignment, fields);
  refreshCurrentView();

  try {
    await updateListItem(CONFIG.lists.quarterlyAssignments, assignmentId, fields);
    const assignedEmail = role === 'preparer' ? assignment.Preparer : assignment.Reviewer;
    const onBehalf = assignedEmail && assignedEmail !== userEmail;
    await writeAuditLog('SignOff', {
      taskName: assignment.Title || assignment.TaskTemplateLookupId,
      assignmentId,
      newValue: onBehalf
        ? `${role} signed off by ${userEmail} ON BEHALF OF ${assignedEmail}`
        : `${role} signed off by ${userEmail}`,
    });
    showToast('✓ Signed off', 'success');
  } catch (err) {
    // Restore full snapshot — covers all fields set above, not just a subset.
    Object.assign(assignment, snapshot);
    refreshCurrentView();
    showToast('Sign-off failed — please try again', 'error');
    logError('Sign-off failed:', err);
  }
}

// ============================================================
// REVERSAL
// ============================================================
async function performReversal(assignmentId, role, reason) {
  const assignment = STATE.assignments.find(a => a._id === assignmentId);
  if (!assignment) return;

  const fields = {};
  let prevValue = '';

  if (role === 'preparer') {
    // Capture both sign-offs in prevValue before clearing — reviewer data would
    // otherwise be lost from the audit trail if it was also cleared.
    const reviewerNote = assignment.ReviewerSignOff
      ? ` | Reviewer sign-off by ${assignment.ReviewerSignOffBy} on ${assignment.ReviewerSignOffDate} also cleared`
      : '';
    prevValue = `Preparer signed off by ${assignment.PreparerSignOffBy} on ${assignment.PreparerSignOffDate}${reviewerNote}`;

    fields.PreparerSignOff     = false;
    fields.PreparerSignOffDate = null;
    fields.PreparerSignOffBy   = null;
    if (assignment.ReviewerSignOff) {
      fields.ReviewerSignOff     = false;
      fields.ReviewerSignOffDate = null;
      fields.ReviewerSignOffBy   = null;
    }
    fields.Status = 'Not Started';
  } else {
    prevValue = `Reviewer signed off by ${assignment.ReviewerSignOffBy} on ${assignment.ReviewerSignOffDate}`;
    fields.ReviewerSignOff     = false;
    fields.ReviewerSignOffDate = null;
    fields.ReviewerSignOffBy   = null;
    fields.Status = 'Prepared';
  }

  // Snapshot before optimistic update so we can restore on failure.
  const snapshot = {};
  Object.keys(fields).forEach(k => { snapshot[k] = assignment[k]; });

  Object.assign(assignment, fields);
  refreshCurrentView();

  try {
    await updateListItem(CONFIG.lists.quarterlyAssignments, assignmentId, fields);
    await writeAuditLog('Reversal', {
      taskName:      assignment.Title,
      assignmentId,
      previousValue: prevValue,
      reason,
    });
    showToast('Sign-off reversed', 'success');
  } catch (err) {
    // Restore full snapshot so UI reflects actual SharePoint state.
    Object.assign(assignment, snapshot);
    refreshCurrentView();
    showToast('Reversal failed — please try again', 'error');
    logError('Reversal failed:', err);
  }
}

// ============================================================
// MATRIX STATUS UPDATE
// ============================================================
// Guards against double-clicks while a matrix write is in flight.
let _matrixUpdateInFlight = false;

async function performMatrixUpdate(matrixItem, column, newStatus) {
  if (_matrixUpdateInFlight) return;
  _matrixUpdateInFlight = true;

  const existing = STATE.matrixStatus.find(
    m => m.MatrixItem === matrixItem && m.Quarter === STATE.activeQuarter
  );

  const now = new Date().toISOString();
  const userEmail = STATE.currentUser.Email;

  const fm = MATRIX_FIELD_MAP[column];
  if (!fm) { _matrixUpdateInFlight = false; return; }

  const fields = {
    [fm.status]: newStatus,
    [fm.date]:   now,
    [fm.by]:     userEmail,
  };

  // Optimistic update — apply to STATE immediately so the matrix re-renders
  // without waiting for the SharePoint round-trip.
  const snapshot = {};
  if (existing) {
    Object.keys(fields).forEach(k => { snapshot[k] = existing[k]; });
    Object.assign(existing, fields);
  }
  renderMatrixView();

  try {
    if (existing) {
      await updateListItem(CONFIG.lists.matrixStatus, existing._id, fields);
    } else {
      // Look up MatrixSection from templates — required field on MatrixStatus list.
      const sectionTemplate = STATE.templates.find(t => t.MatrixItem === matrixItem);
      const matrixSection = sectionTemplate?.MatrixSection || null;
      const created = await createListItem(CONFIG.lists.matrixStatus, {
        Title:         `${STATE.activeQuarter}-${matrixItem}`,
        Quarter:       STATE.activeQuarter,
        MatrixItem:    matrixItem,
        MatrixSection: matrixSection,
        ...fields,
      });
      STATE.matrixStatus.push({ ...created.fields, _id: created.id });
    }
    await writeAuditLog('MatrixStatusChange', {
      taskName: `${matrixItem} — ${column}`,
      newValue: newStatus,
    });
    showToast(`✓ ${column} updated to ${newStatus}`, 'success');
  } catch (err) {
    // Revert optimistic update on failure.
    if (existing && Object.keys(snapshot).length) {
      Object.assign(existing, snapshot);
      renderMatrixView();
    }
    showToast('Update failed — please try again', 'error');
    logError('Matrix update failed:', err);
  } finally {
    _matrixUpdateInFlight = false;
  }
}

// ============================================================
// USER HELPERS
// ============================================================
function getUserByEmail(email) {
  return STATE.users.find(u => u.Email === email);
}

function renderBadge(email) {
  const user = getUserByEmail(email);
  // Centralised display name: prefer Title, fall back to email prefix, then '?'.
  // Always escape to prevent XSS from SharePoint-sourced display names.
  const displayName = escapeHtml(
    (user?.Title) || (email ? email.split('@')[0] : null) || '?'
  );
  if (!user?.Emoji) {
    return `<span class="person-badge" style="background:var(--light-gray);color:var(--dark-slate)">${displayName}</span>`;
  }
  const hex = user.Color || '#75787B';
  return `<span class="person-badge" style="background:${hex}22;color:${hex}">${escapeHtml(user.Emoji)} ${displayName}</span>`;
}

// renderBadgeEl removed — unused. Use renderBadge() which returns an HTML string.

// ============================================================
// VIEW ROUTING
// ============================================================
function showView(viewName) {
  // My Tasks always shows the live quarter — snap back if viewing history.
  if (viewName === 'my-tasks' && isViewingHistory()) {
    switchToQuarter(STATE.activeQuarter);
    return;
  }
  STATE.currentView = viewName;

  // Update nav
  document.querySelectorAll('.nav-link').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.view === viewName);
  });

  // Hide all views
  document.querySelectorAll('.view').forEach(v => {
    v.classList.remove('active');
    v.style.display = 'none';
  });

  // Show target view
  const target = document.getElementById(`view-${viewName}`);
  if (target) {
    target.classList.add('active');
    target.style.display = 'block';
  }

  // Close side panel
  closeTaskPanel();

  // Render the view
  renderCurrentView();
}

function refreshCurrentView() {
  renderCurrentView();
  updateWDIndicator();
}

function renderCurrentView() {
  switch (STATE.currentView) {
    case 'my-tasks':       renderMyTasks();       break;
    case 'all-tasks':      renderAllTasks();      break;
    case 'review-comments':renderReviewComments();break;
    case 'matrix':         renderMatrixView();    break;
    case 'dashboard':      renderDashboard();     break;
    case 'calendar':       renderCalendarView();  break;
    case 'admin':          renderAdminView();     break;
    case 'profile':        renderProfileView();   break;
  }
}

// ============================================================
// WD INDICATOR
// ============================================================
function updateWDIndicator() {
  const pill = document.getElementById('wd-indicator');
  if (!pill) return;
  if (isViewingHistory()) {
    // Show the historical quarter name instead of a live workday
    pill.textContent = STATE.viewingQuarter;
    pill.style.background = 'var(--amber)';
  } else {
    pill.textContent = getWDIndicatorText(STATE.activeQuarter);
    pill.style.background = '';
  }
}

// ============================================================
// MY TASKS VIEW
// ============================================================
function filterMyAssignments() {
  const email = STATE.currentUser?.Email;
  if (!email) return [];
  return STATE.assignments.filter(a =>
    a.Preparer === email || a.Reviewer === email
  );
}

// Renders a named task section: toggles visibility, updates count, and fills cards.
function renderTaskSection({ sectionId, dividerId, countId, labelId, cardContainerId, tasks, email, isOverdue = false, isWaiting = false, labelText = null }) {
  const section = document.getElementById(sectionId);
  const divider = dividerId ? document.getElementById(dividerId) : null;
  const visible = tasks.length > 0;

  if (section) section.classList.toggle('hidden', !visible);
  if (divider) divider.style.display = visible ? '' : 'none';

  const countEl = countId ? document.getElementById(countId) : null;
  if (countEl) countEl.textContent = tasks.length;

  if (labelId && labelText) {
    const labelEl = document.getElementById(labelId);
    if (labelEl) labelEl.textContent = labelText;
  }

  const container = document.getElementById(cardContainerId);
  if (container) container.innerHTML = tasks.map(t => renderTaskCard(t, email, isOverdue, isWaiting)).join('');
}

function renderMyTasks() {
  const email = STATE.currentUser?.Email;
  const quarter = STATE.activeQuarter;

  const sub = document.getElementById('my-tasks-sub');
  if (sub) sub.textContent = `${quarter || 'No active quarter'} · Your assigned tasks`;

  if (!quarter) {
    // Show the no-quarter placeholder directly without going through the view router.
    // We do NOT set STATE.currentView here — no-quarter is not a routed view and the
    // router has no case for it. Keeping STATE.currentView = 'my-tasks' means the next
    // refreshCurrentView() call will re-run renderMyTasks(), which will re-check the
    // quarter and show this placeholder again if still needed.
    document.querySelectorAll('.view').forEach(v => { v.classList.remove('active'); v.style.display = 'none'; });
    const noQ = document.getElementById('view-no-quarter');
    if (noQ) { noQ.classList.add('active'); noQ.style.display = 'block'; }
    return;
  }

  const tasks = filterMyAssignments();
  const wd          = getTodaysWorkday(quarter);
  const tomorrowWD  = getTomorrowWorkday(quarter);
  const todayWD     = typeof wd === 'number' ? wd : -1;

  const overdue    = tasks.filter(t => isTaskOverdue(t) && t.Status !== 'Complete');
  const waiting    = tasks.filter(t => !isTaskOverdue(t) && t.Status !== 'Complete' && isLocked(t, email));
  const active     = tasks.filter(t => !isTaskOverdue(t) && t.Status !== 'Complete' && !isLocked(t, email));
  const dueToday   = active.filter(t => getDueWD(t, email) === todayWD);
  // Second condition (getDueWD !== todayWD) is always true when first is true since tomorrowWD !== todayWD.
  const dueTomorrow = tomorrowWD !== null
    ? active.filter(t => getDueWD(t, email) === tomorrowWD)
    : [];
  const upcoming   = active.filter(t =>
    getDueWD(t, email) !== todayWD &&
    (tomorrowWD === null || getDueWD(t, email) !== tomorrowWD)
  );

  renderTaskSection({ sectionId: 'my-tasks-overdue',   dividerId: 'div-overdue',   countId: 'overdue-count',   cardContainerId: 'overdue-cards',   tasks: overdue,      email, isOverdue: true });
  renderTaskSection({ sectionId: 'my-tasks-today',     dividerId: 'div-today',     countId: null,              cardContainerId: 'today-cards',     tasks: dueToday,     email, labelId: 'today-section-label',    labelText: typeof wd === 'number' ? `DUE TODAY — WD${wd}` : 'DUE TODAY' });
  renderTaskSection({ sectionId: 'my-tasks-tomorrow',  dividerId: 'div-tomorrow',  countId: 'tomorrow-count',  cardContainerId: 'tomorrow-cards',  tasks: dueTomorrow,  email, labelId: 'tomorrow-section-label', labelText: tomorrowWD !== null ? `DUE TOMORROW — WD${tomorrowWD}` : 'DUE TOMORROW' });
  renderTaskSection({ sectionId: 'my-tasks-upcoming',  dividerId: 'div-upcoming',  countId: 'upcoming-count',  cardContainerId: 'upcoming-cards',  tasks: upcoming,     email });

  // Waiting section (always visible, collapsed by default)
  const waitingCountEl = document.getElementById('waiting-count');
  if (waitingCountEl) waitingCountEl.textContent = waiting.length;
  const waitingCards = document.getElementById('waiting-cards');
  if (waitingCards) waitingCards.innerHTML = waiting.map(t => renderTaskCard(t, email, false, true)).join('');

  attachCardEvents();
}

function isLocked(assignment, email) {
  if (assignment.SignOffMode === 'Preparer Only') return false;
  if (assignment.Reviewer === email && !assignment.PreparerSignOff) return true;
  return false;
}

function getDueWD(assignment, email) {
  if (assignment.Preparer === email && !assignment.PreparerSignOff) {
    return Number(assignment.PreparerWorkday);
  }
  if (assignment.Reviewer === email && !assignment.ReviewerSignOff) {
    return Number(assignment.ReviewerWorkday);
  }
  // Fallback: called only from renderMyTasks where tasks are pre-filtered to the
  // current user's assignments, so this branch fires only for completed tasks.
  // Returning PreparerWorkday is safe — completed tasks don't appear in active buckets.
  return Number(assignment.PreparerWorkday);
}

function renderTaskCard(assignment, currentEmail, isOverdue = false, isWaiting = false) {
  const overdueCls = isOverdue ? 'overdue' : '';
  const waitingCls = isWaiting ? 'waiting' : '';
  const isPreparer       = assignment.Preparer === currentEmail;
  const isReviewer       = assignment.Reviewer === currentEmail;
  const isFinalReviewer  = STATE.isFinalReviewer;
  const isAdmin          = STATE.isAdmin;

  // Rule 3: anyone can sign preparer step
  // Rule 4: reviewer step restricted to assigned reviewer, admin, FinalReviewer
  const canSignPreparer  = !assignment.PreparerSignOff;
  const canSignReviewer  = !assignment.ReviewerSignOff && (isReviewer || isAdmin || isFinalReviewer);
  const role = canSignPreparer ? 'preparer'
    : canSignReviewer ? 'reviewer' : null;
  const locked = isLocked(assignment, currentEmail);
  const dueWD = getDueWD(assignment, currentEmail);
  const dueDate = resolveWorkday(getReadQuarter(), dueWD);

  // Check for urgent review comments
  const hasUrgentRC = STATE.reviewComments.some(
    rc => rc.TaskTemplateLookupId === assignment.TaskTemplateLookupId &&
          rc.Priority === 'Urgent' && rc.Status === 'Open'
  );
  const rcCount = STATE.reviewComments.filter(
    rc => rc.TaskTemplateLookupId === assignment.TaskTemplateLookupId
  ).length;

  const prepBadge = renderBadge(assignment.Preparer);
  const revBadge = assignment.Reviewer ? renderBadge(assignment.Reviewer) : '';

  let signoffBtn = '';
  if (isViewingHistory()) {
    // No sign-off actions available when browsing a historical quarter.
    signoffBtn = '';
  } else if (isWaiting || locked) {
    signoffBtn = `<button class="btn-secondary btn-sm" disabled>🔒 Awaiting preparer sign-off</button>`;
  } else if (role) {
    const label = role === 'preparer' ? 'Sign Off as Preparer' : 'Sign Off as Reviewer';
    signoffBtn = `<button class="btn-primary btn-sm" data-action="signoff" data-id="${assignment._id}" data-role="${role}">✓ ${label}</button>`;
  }

  const commentBtn = `<button class="btn-icon" data-action="open-task" data-id="${assignment._id}">💬 ${rcCount}</button>`;
  const linkBtn = assignment.HasDocumentLink && assignment.DocumentLink
    ? `<a class="btn-icon" href="${assignment.DocumentLink}" target="_blank">🔗</a>`
    : '';

  return `
    <div class="task-card ${overdueCls} ${waitingCls}" data-action="open-task" data-id="${assignment._id}" tabindex="0" role="button" aria-label="${escapeHtml(assignment.Title || 'Task')}">
      <div class="task-card-top">
        <div class="task-card-title">${escapeHtml(assignment.Title || '')}</div>
        ${isOverdue ? `<span class="overdue-badge">Overdue · WD${dueWD}${dueDate ? ' · ' + formatDateShort(dueDate) : ''}</span>` : ''}
        ${hasUrgentRC ? `<span class="urgent-rc-badge">💬 Urgent comment</span>` : ''}
      </div>
      <div class="task-card-meta">
        <span class="cat-tag">${escapeHtml(assignment.Category || '')}</span>
        <span class="due-tag ${isOverdue ? 'overdue' : ''}">Due WD${dueWD}${dueDate ? ' · ' + formatDateShort(dueDate) : ''}</span>
      </div>
      <div class="task-card-people">
        ${prepBadge}
        ${revBadge ? `<span style="font-size:10px;color:var(--slate)">Reviewer:</span>${revBadge}` : '<span style="font-size:10px;color:var(--slate)">Preparer only</span>'}
      </div>
      <div class="task-card-actions">
        ${signoffBtn}
        ${commentBtn}
        ${linkBtn}
      </div>
    </div>`;
}

// ============================================================
// ALL TASKS VIEW
// ============================================================
// Renders the active filter chips bar above the All Tasks table.
// Shows one chip per non-default filter with an × to remove it, plus
// a "Clear all" link when more than one filter is active.
function renderActiveFilterChips() {
  const bar = document.getElementById('active-filters-bar');
  if (!bar) return;

  const f = STATE.filters;
  const chips = [];

  if (f.status !== 'all') {
    const labels = { unsigned: 'Unsigned only', overdue: 'Overdue', complete: 'Complete' };
    chips.push({ label: `Status: ${labels[f.status] || f.status}`, clear: () => { f.status = 'all'; saveFilters(); } });
  }
  if (f.category !== 'all') {
    chips.push({ label: `Category: ${f.category}`, clear: () => { f.category = 'all'; saveFilters(); } });
  }
  if (f.assignee !== 'all') {
    const u = getUserByEmail(f.assignee);
    const name = u ? `${u.Emoji || ''} ${u.Title || f.assignee.split('@')[0]}`.trim() : f.assignee.split('@')[0];
    chips.push({ label: `Assignee: ${name}`, clear: () => { f.assignee = 'all'; saveFilters(); } });
  }
  if (f.search) {
    chips.push({ label: `Search: "${f.search}"`, clear: () => { f.search = ''; const el = document.getElementById('filter-search'); if (el) el.value = ''; } });
  }

  if (!chips.length) {
    bar.classList.add('hidden');
    bar.innerHTML = '';
    return;
  }

  bar.classList.remove('hidden');
  bar.innerHTML = chips.map((chip, i) => `
    <span class="filter-chip">
      ${escapeHtml(chip.label)}
      <button class="filter-chip-remove" data-chip="${i}" aria-label="Remove filter: ${escapeHtml(chip.label)}">×</button>
    </span>`).join('') +
    (chips.length > 1 ? '<button class="filter-chip-clear-all" id="btn-clear-all-filters">Clear all</button>' : '');

  // Wire chip remove buttons
  bar.querySelectorAll('.filter-chip-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      chips[Number(btn.dataset.chip)].clear();
      syncFilterUI();
      renderAllTasks();
    });
  });

  // Wire clear all
  document.getElementById('btn-clear-all-filters')?.addEventListener('click', () => {
    f.status = 'all'; f.category = 'all'; f.assignee = 'all'; f.search = '';
    saveFilters();
    syncFilterUI();
    renderAllTasks();
  });
}

// Syncs the toolbar UI controls to match STATE.filters after a programmatic reset.
function syncFilterUI() {
  // Status buttons
  document.querySelectorAll('[data-filter="status"]').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.value === STATE.filters.status);
  });
  // Category select
  const catSel = document.getElementById('filter-category');
  if (catSel) catSel.value = STATE.filters.category;
  // Assignee select
  const asnSel = document.getElementById('filter-assignee');
  if (asnSel) asnSel.value = STATE.filters.assignee;
  // Search input
  const searchEl = document.getElementById('filter-search');
  if (searchEl) searchEl.value = STATE.filters.search || '';
}

function renderAllTasks() {
  const quarter = getReadQuarter();
  const sub = document.getElementById('all-tasks-sub');
  if (sub) sub.textContent = `${quarter || '—'} · ${STATE.assignments.length} tasks · ${getCompletionPct()}% complete`;

  populateCategoryFilter();
  populateAssigneeFilter();
  renderSortHeaders();
  renderActiveFilterChips();

  const filtered = getFilteredAssignments();
  const tbody = document.getElementById('all-tasks-tbody');
  if (!tbody) return;

  tbody.innerHTML = '';
  // Category group headers only make sense when sorting by category
  const groupByCategory = STATE.filters.sort === 'category';
  let lastCategory = null;

  filtered.forEach(a => {
    if (groupByCategory && a.Category !== lastCategory) {
      const headerRow = tbody.insertRow();
      headerRow.className = 'category-header';
      headerRow.insertCell().colSpan = 8;
      headerRow.cells[0].textContent = a.Category || '—';
      lastCategory = a.Category;
    }
    const row = tbody.insertRow();
    if (isTaskOverdue(a)) row.classList.add('overdue-row');
    row.dataset.id = a._id;
    row.addEventListener('click', () => openTaskPanel(a._id));
    row.innerHTML = `
      <td style="font-weight:500;font-size:12px">${escapeHtml(a.Title || '')}</td>
      <td><span class="cat-tag">${escapeHtml(a.Category || '')}</span></td>
      <td>${renderBadge(a.Preparer)}</td>
      <td>${a.Reviewer ? renderBadge(a.Reviewer) : '<span style="font-size:10px;color:var(--slate)">Preparer only</span>'}</td>
      <td style="font-size:11px;color:var(--slate)">${a.PreparerWorkday ? 'WD' + a.PreparerWorkday : '—'}</td>
      <td style="font-size:11px;color:var(--slate)">${a.ReviewerWorkday ? 'WD' + a.ReviewerWorkday : '—'}</td>
      <td>${renderStatusBadge(a)}</td>
      <td style="font-size:10px;color:var(--slate)">${getTaskRCCount(a) || '—'}</td>`;
  });
}

// Updates sort indicator arrows on the table header row.
function renderSortHeaders() {
  const headers = document.querySelectorAll('#all-tasks-thead th[data-sort]');
  headers.forEach(th => {
    const isActive = th.dataset.sort === STATE.filters.sort;
    th.setAttribute('aria-sort', isActive ? (STATE.filters.sortDir === 'asc' ? 'ascending' : 'descending') : 'none');
    // Update the visible arrow character
    const arrow = th.querySelector('.sort-arrow');
    if (arrow) arrow.textContent = isActive ? (STATE.filters.sortDir === 'asc' ? ' ▲' : ' ▼') : ' ⇅';
    th.classList.toggle('sort-active', isActive);
  });
}

function renderStatusBadge(assignment) {
  const s = assignment.Status || 'Not Started';

  // Overdue takes priority over all other states.
  if (isTaskOverdue(assignment)) {
    return `<span class="status-badge status-overdue">⚠ Overdue</span>`;
  }

  // Reviewer step locked — preparer has not signed off yet.
  // Surfaces as Locked in RC cards so reviewers know they cannot act yet.
  if (assignment.SignOffMode !== 'Preparer Only' &&
      !assignment.PreparerSignOff &&
      !assignment.ReviewerSignOff) {
    return `<span class="status-badge status-notstarted">Locked</span>`;
  }

  const map = {
    'Complete':    ['status-complete',   '✓ Complete'],
    'Prepared':    ['status-prepared',   '→ Ready for review'],
    'In Progress': ['status-progress',   'In progress'],
    'Not Started': ['status-notstarted', 'Not started'],
  };
  const [cls, label] = map[s] || map['Not Started'];
  return `<span class="status-badge ${cls}">${label}</span>`;
}

// Status severity order used when sorting by status or overdue-first.
const STATUS_ORDER = { 'Overdue': 0, 'In Progress': 1, 'Not Started': 2, 'Prepared': 3, 'Complete': 4 };

function getEffectiveStatus(a) {
  return isTaskOverdue(a) ? 'Overdue' : (a.Status || 'Not Started');
}

function getFilteredAssignments() {
  const f = STATE.filters;

  const filtered = STATE.assignments.filter(a => {
    if (f.status === 'unsigned' && a.Status === 'Complete') return false;
    if (f.status === 'overdue' && !isTaskOverdue(a)) return false;
    if (f.status === 'complete' && a.Status !== 'Complete') return false;
    if (f.category !== 'all' && a.Category !== f.category) return false;
    if (f.assignee !== 'all' && a.Preparer !== f.assignee && a.Reviewer !== f.assignee) return false;
    if (f.search && !a.Title?.toLowerCase().includes(f.search.toLowerCase())) return false;
    return true;
  });

  const dir = f.sortDir === 'desc' ? -1 : 1;

  filtered.sort((a, b) => {
    let cmp = 0;
    switch (f.sort) {
      case 'overdue':
        // Primary: overdue severity (Overdue first, Complete last)
        // Secondary: prep workday ascending so soonest-due overdue tasks are first
        cmp = (STATUS_ORDER[getEffectiveStatus(a)] ?? 99) - (STATUS_ORDER[getEffectiveStatus(b)] ?? 99);
        if (cmp === 0) cmp = (Number(a.PreparerWorkday) || 0) - (Number(b.PreparerWorkday) || 0);
        break;
      case 'category':
        cmp = (a.Category || '').localeCompare(b.Category || '');
        if (cmp === 0) cmp = (Number(a.PreparerWorkday) || 0) - (Number(b.PreparerWorkday) || 0);
        break;
      case 'prepWD':
        cmp = (Number(a.PreparerWorkday) || 0) - (Number(b.PreparerWorkday) || 0);
        break;
      case 'revWD':
        cmp = (Number(a.ReviewerWorkday) || 0) - (Number(b.ReviewerWorkday) || 0);
        break;
      case 'status':
        cmp = (STATUS_ORDER[getEffectiveStatus(a)] ?? 99) - (STATUS_ORDER[getEffectiveStatus(b)] ?? 99);
        break;
      case 'task':
        cmp = (a.Title || '').localeCompare(b.Title || '');
        break;
      default:
        cmp = (a.Category || '').localeCompare(b.Category || '');
    }
    return cmp * dir;
  });

  return filtered;
}

function getCompletionPct() {
  if (!STATE.assignments.length) return 0;
  const complete = STATE.assignments.filter(a => a.Status === 'Complete').length;
  return Math.round((complete / STATE.assignments.length) * 100);
}

function getTaskRCCount(assignment) {
  return STATE.reviewComments.filter(rc => rc.TaskTemplateLookupId === assignment.TaskTemplateLookupId).length || 0;
}

function populateCategoryFilter() {
  const sel = document.getElementById('filter-category');
  if (!sel) return;
  const current = sel.value;
  const cats = [...new Set(STATE.assignments.map(a => a.Category).filter(Boolean))].sort();
  sel.innerHTML = '<option value="all">All categories</option>' +
    cats.map(c => `<option value="${escapeHtml(c)}" ${c === current ? 'selected' : ''}>${escapeHtml(c)}</option>`).join('');
  // Attach listener only once
  if (!sel.dataset.listenerAttached) {
    sel.dataset.listenerAttached = 'true';
    sel.addEventListener('change', () => {
      STATE.filters.category = sel.value;
      saveFilters();
      renderAllTasks();
    });
  }
}

function populateAssigneeFilter() {
  const sel = document.getElementById('filter-assignee');
  if (!sel) return;
  const current = sel.value;
  const emails = [...new Set(
    STATE.assignments.flatMap(a => [a.Preparer, a.Reviewer].filter(Boolean))
  )];
  sel.innerHTML = '<option value="all">All team members</option>' +
    emails.map(e => {
      const u = getUserByEmail(e);
      const name = u ? `${u.Emoji || ''} ${u.Title || e.split('@')[0]}` : e.split('@')[0];
      return `<option value="${escapeHtml(e)}" ${e === current ? 'selected' : ''}>${escapeHtml(name)}</option>`;
    }).join('');
  if (!sel.dataset.listenerAttached) {
    sel.dataset.listenerAttached = 'true';
    sel.addEventListener('change', () => {
      STATE.filters.assignee = sel.value;
      saveFilters();
      renderAllTasks();
    });
  }
}

// ============================================================
// TASK DETAIL SIDE PANEL
// ============================================================
function openTaskPanel(assignmentId) {
  const assignment = STATE.assignments.find(a => a._id === assignmentId);
  if (!assignment) return;
  STATE.taskDetailId = assignmentId;

  document.getElementById('panel-title').textContent = assignment.Title || '—';
  document.getElementById('panel-meta').textContent =
    `${assignment.Category || ''} · Due WD${assignment.PreparerWorkday} · ${resolveWorkday(getReadQuarter(), assignment.PreparerWorkday) ? formatDateShort(resolveWorkday(getReadQuarter(), assignment.PreparerWorkday)) : ''}`;

  // Assignment section
  const email = STATE.currentUser?.Email;
  const prepBadge = renderBadge(assignment.Preparer);
  const revBadge = assignment.Reviewer ? renderBadge(assignment.Reviewer) : '—';
  const docLink = assignment.HasDocumentLink && assignment.DocumentLink
    ? `<a class="panel-doc-link" href="${escapeHtml(assignment.DocumentLink)}" target="_blank">🔗 Open document</a>`
    : '';
  const canReassign = STATE.isAdmin && !isViewingHistory();
  document.getElementById('panel-assignment').innerHTML = `
    <div class="panel-meta-row"><span class="panel-meta-label">Preparer</span>${prepBadge}${canReassign ? `<button class="btn-icon btn-sm" style="margin-left:6px" data-action="reassign" data-id="${assignment._id}" data-role="preparer">Reassign</button>` : ''}</div>
    <div class="panel-meta-row"><span class="panel-meta-label">Reviewer</span>${revBadge}${canReassign && assignment.Reviewer ? `<button class="btn-icon btn-sm" style="margin-left:6px" data-action="reassign" data-id="${assignment._id}" data-role="reviewer">Reassign</button>` : ''}</div>
    <div class="panel-meta-row"><span class="panel-meta-label">Sign-off mode</span><span style="font-size:11px">${assignment.SignOffMode || '—'}</span></div>
    ${docLink ? `<div class="panel-meta-row" style="border-bottom:none"><span class="panel-meta-label">Document</span>${docLink}</div>` : ''}`;

  // Status chain
  renderPanelStatusChain(assignment, email);

  // Action
  renderPanelAction(assignment, email);

  // Review comments preview
  const rcs = STATE.reviewComments.filter(rc => rc.TaskTemplateLookupId === assignment.TaskTemplateLookupId);
  const rcPreview = document.getElementById('panel-rc-preview');
  if (rcPreview) {
    if (rcs.length) {
      rcPreview.innerHTML = rcs.slice(0,2).map(rc => `
        <div class="rc-card ${rc.Priority === 'Urgent' ? 'urgent' : ''}" style="cursor:default">
          <div class="rc-meta">
            ${renderBadge(rc.CreatedBy)}
            <span class="rc-meta-text">${formatDateShort(rc.CreatedDate) || '—'}</span>
            <span class="${rc.Priority === 'Urgent' ? 'badge-urgent' : 'badge-normal'}">${rc.Priority}</span>
          </div>
          <div class="rc-comment-text">"${escapeHtml((rc.CommentText || '').substring(0, 120))}${rc.CommentText?.length > 120 ? '...' : ''}"</div>
        </div>`).join('');
    } else {
      rcPreview.innerHTML = '<p style="font-size:11px;color:var(--slate)">No review comments on this task.</p>';
    }
  }

  // Audit trail (simplified — from assignments data)
  const auditEl = document.getElementById('panel-audit');
  if (auditEl) {
    const entries = [];
    if (assignment.PreparerSignOff) {
      entries.push({ action: `${renderBadge(assignment.PreparerSignOffBy || assignment.Preparer)} signed off as preparer`, date: assignment.PreparerSignOffDate });
    }
    if (assignment.ReviewerSignOff) {
      entries.push({ action: `${renderBadge(assignment.ReviewerSignOffBy || assignment.Reviewer)} signed off as reviewer`, date: assignment.ReviewerSignOffDate });
    }
    if (!entries.length) {
      auditEl.innerHTML = '<p style="font-size:11px;color:var(--slate)">No activity yet.</p>';
    } else {
      auditEl.innerHTML = entries.map(e => `
        <div class="audit-entry">
          <div class="audit-action">${e.action}</div>
          <div class="audit-meta">${formatDateET(e.date)}</div>
        </div>`).join('');
    }
  }

  // Show panel
  document.getElementById('task-panel').classList.remove('hidden');
  document.getElementById('panel-overlay').classList.remove('hidden');
}

function renderPanelStatusChain(assignment, email) {
  const chain = document.getElementById('panel-status-chain');
  if (!chain) return;
  const isPrepOnly = assignment.SignOffMode === 'Preparer Only';
  const prepDone = assignment.PreparerSignOff;
  const revDone = assignment.ReviewerSignOff;

  chain.innerHTML = `
    <div class="status-step ${prepDone ? 'complete' : ''}">
      <div class="status-step-dot ${prepDone ? 'dot-complete' : 'dot-pending'}"></div>
      <div>
        <div class="status-step-text">Preparer sign-off</div>
        <div class="status-step-sub">${prepDone ? renderBadge(assignment.PreparerSignOffBy || assignment.Preparer) + ' · ' + formatDateET(assignment.PreparerSignOffDate) : renderBadge(assignment.Preparer) + ' · Pending'}</div>
      </div>
    </div>
    ${!isPrepOnly ? `
    <div class="status-step ${revDone ? 'complete' : !prepDone ? 'locked' : ''}">
      <div class="status-step-dot ${revDone ? 'dot-complete' : !prepDone ? 'dot-locked' : 'dot-pending'}"></div>
      <div>
        <div class="status-step-text">Reviewer sign-off</div>
        <div class="status-step-sub">${!prepDone ? '🔒 Locked until preparer signs' : revDone ? renderBadge(assignment.ReviewerSignOffBy || assignment.Reviewer) + ' · ' + formatDateET(assignment.ReviewerSignOffDate) : renderBadge(assignment.Reviewer) + ' · Pending'}</div>
      </div>
    </div>` : ''}`;
}

function renderPanelAction(assignment, email) {
  const actionDiv = document.getElementById('panel-action');
  if (!actionDiv) return;

  if (isViewingHistory()) {
    actionDiv.innerHTML = '<p style="font-size:11px;color:var(--slate)">Read-only — historical quarter.</p>';
    return;
  }

  const isPrepOnly       = assignment.SignOffMode === 'Preparer Only';
  const prepDone         = assignment.PreparerSignOff;
  const revDone          = assignment.ReviewerSignOff;
  const isPreparer       = assignment.Preparer === email;
  const isReviewer       = assignment.Reviewer === email;
  const isAdmin          = STATE.isAdmin;
  const isFinalReviewer  = STATE.isFinalReviewer;

  // RULE 3: Preparer steps — any team member can sign off (always shown).
  // RULE 4: Reviewer steps — restricted to assigned reviewer, admin, FinalReviewer.
  //         Everyone else sees an "on behalf" override button that logs the actual signer.
  const canSignPreparer  = true;  // open to all
  const canSignReviewer  = isReviewer || isAdmin || isFinalReviewer;

  // Reversals stay restricted — only assigned person or admin can reverse.
  const canReversePreparer = isPreparer || isAdmin;
  const canReverseReviewer = isReviewer || isAdmin || isFinalReviewer;

  const et = formatDateET(new Date().toISOString());
  let html = '';

  if (!prepDone) {
    const onBehalf = !isPreparer;
    html = `
      <div class="confirm-box">
        <div class="confirm-text">Sign off preparer step?${onBehalf ? ` <span style="font-size:10px;color:var(--amber);font-weight:500">On behalf of ${renderBadge(assignment.Preparer)}</span>` : ''}</div>
        <div class="confirm-sub">Recorded as ${renderBadge(email)} · ${et}</div>
        <div class="confirm-btns">
          <button class="btn-primary btn-sm" data-action="signoff" data-id="${assignment._id}" data-role="preparer">✓ ${onBehalf ? 'Sign Off on Behalf' : 'Sign Off as Preparer'}</button>
        </div>
      </div>`;
  } else if (!isPrepOnly && !revDone) {
    if (canSignReviewer) {
      const onBehalf = !isReviewer;
      html = `
        <div class="confirm-box">
          <div class="confirm-text">Sign off reviewer step?${onBehalf ? ` <span style="font-size:10px;color:var(--amber);font-weight:500">On behalf of ${renderBadge(assignment.Reviewer)}</span>` : ''}</div>
          <div class="confirm-sub">Recorded as ${renderBadge(email)} · ${et}</div>
          <div class="confirm-btns">
            <button class="btn-primary btn-sm" data-action="signoff" data-id="${assignment._id}" data-role="reviewer">✓ ${onBehalf ? 'Sign Off on Behalf' : 'Sign Off as Reviewer'}</button>
          </div>
        </div>`;
    } else {
      // Not authorised to sign reviewer step — show override button
      html = `
        <div style="font-size:11px;color:var(--slate);margin-bottom:8px">
          Awaiting reviewer sign-off by ${renderBadge(assignment.Reviewer)}.
        </div>
        <button class="btn-secondary btn-sm" data-action="signoff-behalf" data-id="${assignment._id}" data-role="reviewer">
          Sign Off on Behalf…
        </button>`;
    }
  } else {
    // All signed off — show reverse options
    if (prepDone && canReversePreparer) {
      html += `<button class="btn-danger btn-sm" data-action="reverse" data-id="${assignment._id}" data-role="preparer" style="margin-right:6px">Reverse preparer sign-off</button>`;
    }
    if (revDone && canReverseReviewer) {
      html += `<button class="btn-danger btn-sm" data-action="reverse" data-id="${assignment._id}" data-role="reviewer">Reverse reviewer sign-off</button>`;
    }
    if (!html) {
      html = `<p style="font-size:11px;color:var(--slate)">Task complete.</p>`;
    }
  }

  actionDiv.innerHTML = html;
  attachCardEvents();
}

function closeTaskPanel() {
  document.getElementById('task-panel')?.classList.add('hidden');
  document.getElementById('panel-overlay')?.classList.add('hidden');
  STATE.taskDetailId = null;
}

// ============================================================
// REVIEW COMMENTS VIEW
// ============================================================
function renderReviewComments() {
  const quarter = getReadQuarter();

  // Populate the quarter filter dropdown — repopulated on every render so new quarters appear.
  const quarterFilterSel = document.getElementById('rc-quarter-filter');
  if (quarterFilterSel) {
    const currentQ = quarterFilterSel.value || 'all';
    const quarters = [...new Set(STATE.reviewComments.map(rc => rc.Quarter).filter(Boolean))].sort().reverse();
    quarterFilterSel.innerHTML = '<option value="all">All quarters</option>' +
      quarters.map(q => `<option value="${escapeHtml(q)}" ${q === currentQ ? 'selected' : ''}>${escapeHtml(q)}</option>`).join('');
    if (!quarterFilterSel.dataset.listenerAttached) {
      quarterFilterSel.dataset.listenerAttached = 'true';
      quarterFilterSel.addEventListener('change', () => {
        STATE.filters.rcQuarter = quarterFilterSel.value;
        renderReviewComments();
      });
    }
  }

  // Apply quarter filter if set
  const rcQuarter = STATE.filters.rcQuarter && STATE.filters.rcQuarter !== 'all'
    ? STATE.filters.rcQuarter
    : quarter;
  const sub = document.getElementById('rc-sub');
  if (sub) {
    const urgent  = STATE.reviewComments.filter(rc => rc.Priority === 'Urgent' && rc.Status === 'Open').length;
    const open    = STATE.reviewComments.filter(rc => rc.Status === 'Open').length;
    const resolved = STATE.reviewComments.filter(rc => rc.Status === 'Resolved').length;
    sub.textContent = `${quarter || '—'} · ${urgent} urgent · ${open} open · ${resolved} resolved`;
    document.getElementById('rc-urgent-count').textContent = urgent;
    document.getElementById('rc-open-count').textContent = STATE.reviewComments.filter(rc => rc.Priority === 'Normal' && rc.Status === 'Open').length;
    document.getElementById('rc-resolved-count').textContent = resolved;
  }

  const urgentList = document.getElementById('rc-urgent-list');
  const openList   = document.getElementById('rc-open-list');
  const resolvedList = document.getElementById('rc-resolved-list');

  const allRCs  = rcQuarter ? STATE.reviewComments.filter(rc => rc.Quarter === rcQuarter) : STATE.reviewComments;
  const urgent  = allRCs.filter(rc => rc.Priority === 'Urgent' && rc.Status === 'Open');
  const normal  = allRCs.filter(rc => rc.Priority === 'Normal' && rc.Status === 'Open');
  const resolved = allRCs.filter(rc => rc.Status === 'Resolved');

  const urgentSection = document.getElementById('rc-urgent-section');
  if (urgentSection) urgentSection.classList.toggle('hidden', urgent.length === 0);
  const openSection = document.getElementById('rc-open-section');
  if (openSection) openSection.classList.toggle('hidden', normal.length === 0);

  if (urgentList) urgentList.innerHTML = urgent.map(rc => renderRCCard(rc)).join('');
  if (openList)   openList.innerHTML   = normal.map(rc => renderRCCard(rc)).join('');
  if (resolvedList) resolvedList.innerHTML = resolved.map(rc => renderRCCard(rc, true)).join('');
}

function renderRCCard(rc, isResolved = false) {
  const template   = STATE.templates.find(t => t._id === rc.TaskTemplateLookupId);
  const taskName   = template?.TaskName || rc.Title || '—';
  const taggedBadges = (rc.TaggedUsers || '').split(';').filter(Boolean).map(e => renderBadge(e.trim())).join('');
  const resNote = rc.ResolutionNote
    ? `<div class="resolution-note">✓ Resolved by ${renderBadge(rc.ResolvedBy)} · ${formatDateET(rc.ResolvedDate)}${rc.ResolutionNote ? ' · "' + escapeHtml(rc.ResolutionNote) + '"' : ''}</div>`
    : '';
  const canResolve = !isResolved && (rc.CreatedBy === STATE.currentUser?.Email || STATE.isAdmin);

  // Find the assignment for this task so we can show key metadata without opening the panel.
  const assignment = STATE.assignments.find(a => a.TaskTemplateLookupId === rc.TaskTemplateLookupId);
  const assignmentId = assignment?._id || null;
  const taskMeta = assignment
    ? `<div class="rc-task-meta">
        ${renderBadge(assignment.Preparer)}
        ${assignment.Reviewer ? `<span class="rc-meta-text">→</span>${renderBadge(assignment.Reviewer)}` : ''}
        <span class="rc-meta-text">Due WD${assignment.PreparerWorkday}${assignment.ReviewerWorkday ? ' / WD' + assignment.ReviewerWorkday : ''}</span>
        ${renderStatusBadge(assignment)}
       </div>`
    : '';

  // Reply count badge — gives a heads-up that there's a thread without scrolling.
  const replyCount = (STATE.rcReplies || []).filter(r => r.ReviewCommentLookupId === rc._id).length;
  const replyBadge = replyCount > 0
    ? `<span class="rc-reply-count">${replyCount} repl${replyCount === 1 ? 'y' : 'ies'}</span>`
    : '';

  return `
    <div class="rc-card ${rc.Priority === 'Urgent' ? 'urgent' : ''} ${isResolved ? 'resolved' : ''}">
      <div class="rc-card-header">
        <div>
          <div class="rc-task-link ${assignmentId ? 'rc-task-link-active' : ''}"
               ${assignmentId ? `data-action="rc-open-task" data-id="${assignmentId}"` : ''}
               role="${assignmentId ? 'button' : ''}"
               ${assignmentId ? 'tabindex="0"' : ''}
               title="${assignmentId ? 'Click to open task' : ''}"
          >${escapeHtml(taskName)}</div>
          ${taskMeta}
        </div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px;flex-shrink:0">
          <div class="rc-badges">
            <span class="${rc.Priority === 'Urgent' ? 'badge-urgent' : 'badge-normal'}">${rc.Priority}</span>
            <span class="${isResolved ? 'badge-resolved' : 'badge-open'}">${isResolved ? '✓ Resolved' : 'Open'}</span>
          </div>
          ${replyBadge}
        </div>
      </div>
      <div class="rc-comment-text">"${escapeHtml(rc.CommentText || '')}"</div>
      <div class="rc-meta">
        ${renderBadge(rc.CreatedBy)}
        <span class="rc-meta-text">${formatDateShort(rc.CreatedDate) || '—'}</span>
        ${taggedBadges}
      </div>
      ${resNote}
      ${renderRCReplies(rc._id)}
      ${!isResolved ? `
      <div class="rc-actions">
        ${!isViewingHistory() ? `<button class="btn-icon" data-action="rc-reply" data-id="${rc._id}">Reply</button>` : ''}
        ${canResolve && !isViewingHistory() ? `<button class="btn-success btn-sm" data-action="rc-resolve" data-id="${rc._id}">✓ Mark Resolved</button>` : ''}
      </div>` : ''}
    </div>`;
}

function renderRCReplies(rcId) {
  const replies = (STATE.rcReplies || []).filter(r => r.ReviewCommentLookupId === rcId);
  if (!replies.length) return '';
  return replies.map(r => `
    <div class="rc-reply">
      <div class="rc-reply-text">${escapeHtml(r.ReplyText || '')}</div>
      <div class="rc-reply-meta">${renderBadge(r.CreatedByEmail)} · ${formatDateShort(r.CreatedDate)}</div>
    </div>`).join('');
}

// ============================================================
// MATRIX VIEW
// ============================================================
function renderMatrixView() {
  const container = document.getElementById('matrix-container');
  if (!container) return;

  const quarter = getReadQuarter();
  const sub = document.getElementById('matrix-sub');
  if (sub) sub.textContent = `${quarter || '—'} · Final reviewer summary`;

  // Get matrix items grouped by section
  const sections = {
    'Form 10-Q': [],
    'MD&A': [],
  };

  // Get unique matrix items from templates
  const filingType = isQuarterQ4(quarter) ? '10-K' : '10-Q';
  STATE.templates
    .filter(t => t.MatrixItem && t.MatrixSection && (t.FilingType === filingType || t.FilingType === 'Both'))
    .forEach(t => {
      const section = t.MatrixSection;
      if (!sections[section]) sections[section] = [];
      if (!sections[section].find(i => i.name === t.MatrixItem)) {
        sections[section].push({ name: t.MatrixItem });
      }
    });

  // Build matrix table
  const checkpoints = CONFIG.matrixCheckpoints;
  let html = `<table class="matrix-table">
    <thead><tr>
      <th class="left-align" style="min-width:160px">Item</th>
      <th class="left-align" style="min-width:80px">Preparer</th>
      <th class="left-align" style="min-width:80px">1st Reviewer</th>
      ${checkpoints.map(cp => `<th style="min-width:60px">${escapeHtml(cp)}</th>`).join('')}
    </tr></thead>
    <tbody>`;

  Object.entries(sections).forEach(([sectionName, items]) => {
    if (!items.length) return;
    html += `<tr class="section-header"><td colspan="${3 + checkpoints.length}">${escapeHtml(sectionName)}</td></tr>`;

    items.forEach(item => {
      // Get preparer and reviewer from assignments
      const assignments = STATE.assignments.filter(a => a.MatrixItem === item.name);
      const preparers = [...new Set(assignments.map(a => a.Preparer).filter(Boolean))];
      const reviewers = [...new Set(assignments.map(a => a.Reviewer).filter(Boolean))];

      html += `<tr>
        <td class="item-cell">${escapeHtml(item.name)}</td>
        <td class="person-cell">${preparers.map(e => renderBadge(e)).join('')}</td>
        <td class="person-cell">${reviewers.map(e => renderBadge(e)).join('')}</td>`;

      checkpoints.forEach(cp => {
        const isMatrixOnly = CONFIG.matrixOnlyColumns.includes(cp);

        if (isMatrixOnly) {
          // Matrix-only column — use module-level MATRIX_FIELD_MAP
          const ms = STATE.matrixStatus.find(m => m.MatrixItem === item.name && m.Quarter === quarter);
          const fm = MATRIX_FIELD_MAP[cp];
          const status = ms?.[fm.status] || 'Not Started';
          const isFinalReview = cp === 'Final Review';
          const canAct = isFinalReview ? STATE.isFinalReviewer : true;

          if (status === 'Complete') {
            const tooltip = `Signed off by ${ms?.[fm.by] || '—'} · ${formatDateET(ms?.[fm.date])}`;
            html += `<td class="cell-done" title="${escapeHtml(tooltip)}">
              <svg width="12" height="12" viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
            </td>`;
          } else if (status === 'N/A') {
            html += `<td class="cell-na" title="Not applicable"></td>`;
          } else if (canAct && !isViewingHistory()) {
            html += `<td class="cell-actionable" data-action="matrix-update" data-item="${escapeHtml(item.name)}" data-col="${escapeHtml(cp)}" title="Click to update">
              <svg width="12" height="12" viewBox="0 0 12 12"><circle cx="6" cy="6" r="4.5" fill="none" stroke="#005EFF" stroke-width="1.5"/></svg>
            </td>`;
          } else {
            html += `<td class="cell-empty"></td>`;
          }
        } else {
          // Task-linked column — use getCheckpointRole/getSignOffFields for consistent field access
          const linkedAssignment = STATE.assignments.find(
            a => a.MatrixItem === item.name && a.MatrixCheckpoint === cp
          );

          if (!linkedAssignment) {
            html += `<td class="cell-na" title="Not applicable"></td>`;
          } else {
            const cpRole = getCheckpointRole(cp);
            const cpFields = getSignOffFields(cpRole);
            const done = linkedAssignment[cpFields.signOff];

            if (done) {
              const tooltip = `Signed off by ${linkedAssignment[cpFields.signOffBy] || '—'} · ${formatDateET(linkedAssignment[cpFields.signOffDate])}`;
              html += `<td class="cell-done" title="${escapeHtml(tooltip)}" data-action="open-task" data-id="${linkedAssignment._id}">
                <svg width="12" height="12" viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/></svg>
              </td>`;
            } else {
              const overdue = isTaskOverdue(linkedAssignment);
              const tooltip = `Assigned to ${linkedAssignment[cpFields.assignee] || '—'} · Due WD${linkedAssignment[cpFields.workday]}${overdue ? ' · Overdue' : ''}`;
              html += `<td class="cell-empty" title="${escapeHtml(tooltip)}" data-action="open-task" data-id="${linkedAssignment._id}"></td>`;
            }
          }
        }
      });

      html += '</tr>';
    });
  });

  html += '</tbody></table>';
  container.innerHTML = html;

  // Attach matrix cell events
  container.querySelectorAll('[data-action="matrix-update"]').forEach(cell => {
    cell.setAttribute('tabindex', '0');
    const activateCell = () => {
      STATE.pendingMatrixAction = {
        item: cell.dataset.item,
        col:  cell.dataset.col,
      };
      const isNA = cell.dataset.col === 'Final Review' && !STATE.isFinalReviewer;
      const titleEl = document.getElementById('matrix-modal-title');
      const descEl = document.getElementById('matrix-modal-desc');
      const optsEl = document.getElementById('matrix-modal-options');
      if (titleEl) titleEl.textContent = `Update: ${cell.dataset.item} — ${cell.dataset.col}`;
      if (descEl) descEl.textContent = `Choose the new status for this item.`;
      if (optsEl) optsEl.innerHTML = `
        <label class="radio-opt"><input type="radio" name="matrix-action" value="Complete" checked/> ✓ Mark as Complete</label>
        <label class="radio-opt"><input type="radio" name="matrix-action" value="N/A"/> — Mark as N/A</label>`;
      showModal('modal-matrix-action');
    };
    cell.addEventListener('click', activateCell);
    cell.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); activateCell(); }
    });
  });

  container.querySelectorAll('[data-action="open-task"]').forEach(cell => {
    cell.setAttribute('tabindex', '0');
    cell.addEventListener('click', () => openTaskPanel(cell.dataset.id));
    cell.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openTaskPanel(cell.dataset.id); }
    });
  });
}

// ============================================================
// DASHBOARD VIEW
// ============================================================
function renderDashboard() {
  const quarter = getReadQuarter();
  const sub = document.getElementById('dashboard-sub');
  if (sub) sub.textContent = `${quarter || '—'} · ${STATE.isAdmin ? 'Admin view' : 'Read-only'}`;

  // Metrics
  const total    = STATE.assignments.length;
  const complete = STATE.assignments.filter(a => a.Status === 'Complete').length;
  const overdue  = STATE.assignments.filter(a => isTaskOverdue(a)).length;
  const urgentRC = STATE.reviewComments.filter(rc => rc.Priority === 'Urgent' && rc.Status === 'Open').length;
  const pct = total ? Math.round((complete / total) * 100) : 0;

  const metricGrid = document.getElementById('metric-grid');
  if (metricGrid) {
    metricGrid.innerHTML = `
      <div class="metric-card"><div class="metric-label">Overall complete</div><div class="metric-value ${pct > 75 ? 'success' : ''}">${pct}%</div><div class="metric-sub">${complete} of ${total} tasks</div></div>
      <div class="metric-card"><div class="metric-label">Overdue tasks</div><div class="metric-value ${overdue > 0 ? 'danger' : ''}">${overdue}</div><div class="metric-sub">Across all categories</div></div>
      <div class="metric-card"><div class="metric-label">Urgent comments</div><div class="metric-value ${urgentRC > 0 ? 'danger' : ''}">${urgentRC}</div><div class="metric-sub">${STATE.reviewComments.filter(rc => rc.Status === 'Open').length} open total</div></div>
      <div class="metric-card"><div class="metric-label">Active quarter</div><div class="metric-value" style="font-size:18px;padding-top:4px">${quarter || '—'}</div><div class="metric-sub">${isQuarterQ4(quarter) ? '10-K · WD1–35' : '10-Q · WD1–20'}</div></div>`;
  }

  // Category progress
  const catBars = document.getElementById('category-bars');
  if (catBars) {
    const cats = [...new Set(STATE.assignments.map(a => a.Category).filter(Boolean))].sort();
    catBars.innerHTML = cats.map(cat => {
      const catTasks = STATE.assignments.filter(a => a.Category === cat);
      const catComplete = catTasks.filter(a => a.Status === 'Complete').length;
      const catPct = catTasks.length ? Math.round((catComplete / catTasks.length) * 100) : 0;
      const danger = catPct < 30;
      return `<div class="prog-row">
        <div class="prog-label">${escapeHtml(cat)}</div>
        <div class="prog-bar-wrap"><div class="prog-bar ${danger ? 'danger' : ''}" style="width:${catPct}%"></div></div>
        <div class="prog-pct ${danger ? 'danger' : ''}">${catPct}%</div>
      </div>`;
    }).join('');
  }

  // Person progress
  const personBars = document.getElementById('person-bars');
  if (personBars) {
    personBars.innerHTML = STATE.users.map(user => {
      const myTasks = STATE.assignments.filter(a => a.Preparer === user.Email || a.Reviewer === user.Email);
      if (!myTasks.length) return '';
      const done = myTasks.filter(a => a.Status === 'Complete').length;
      const pctUser = Math.round((done / myTasks.length) * 100);
      const danger = pctUser < 30;
      const hex = user.Color || '#75787B';
      return `<div class="prog-row">
        <span class="person-badge" style="background:${hex}22;color:${hex};width:90px;flex-shrink:0">${user.Emoji || ''} ${user.Title || ''}</span>
        <div class="prog-bar-wrap"><div class="prog-bar ${danger ? 'danger' : ''}" style="width:${pctUser}%"></div></div>
        <div class="prog-pct">${pctUser}%</div>
      </div>`;
    }).filter(Boolean).join('');
  }

  // Upcoming milestones
  const milestoneList = document.getElementById('milestone-list');
  if (milestoneList) {
    const today = todayET();
    const upcoming = STATE.calendar
      .filter(c => c.MilestoneLabel && c.ActualDate >= today)
      .slice(0, 5);
    milestoneList.innerHTML = upcoming.map(m => `
      <div class="milestone-row">
        <span class="milestone-wd">WD${m.WorkdayNumber}</span>
        <span class="milestone-date">${formatDateShort(m.ActualDate)}</span>
        <span class="milestone-name">${escapeHtml(m.MilestoneLabel)}${m.ActualDate === today ? ' <span class="milestone-today">Today</span>' : ''}</span>
      </div>`).join('') || '<p style="font-size:11px;color:var(--slate)">No upcoming milestones.</p>';
  }

  // Overdue detail
  const overdueTitle = document.getElementById('overdue-summary-title');
  const overdueSub   = document.getElementById('overdue-summary-sub');
  if (overdueTitle) overdueTitle.textContent = `${overdue} overdue task${overdue !== 1 ? 's' : ''}`;
  const overdueTasks = STATE.assignments.filter(a => isTaskOverdue(a));
  const cats2 = [...new Set(overdueTasks.map(a => a.Category).filter(Boolean))];
  if (overdueSub) overdueSub.textContent = cats2.length ? `Across ${cats2.join(', ')}` : '';

  const overdueList = document.getElementById('overdue-detail-list');
  if (overdueList) {
    const wd = getTodaysWorkday(STATE.activeQuarter);
    overdueList.innerHTML = overdueTasks.map(a => {
      const dueWD = a.PreparerSignOff ? Number(a.ReviewerWorkday) : Number(a.PreparerWorkday);
      const daysOver = typeof wd === 'number' ? wd - dueWD : 0;
      return `<div class="overdue-detail-row">
        <div>
          <div style="font-size:12px;font-weight:500">${escapeHtml(a.Title || '')}</div>
          <div style="font-size:10px;color:var(--slate)">${escapeHtml(a.Category || '')} · Preparer: ${renderBadge(a.Preparer)} · Due WD${dueWD}${resolveWorkday(STATE.activeQuarter, dueWD) ? ' · ' + formatDateShort(resolveWorkday(STATE.activeQuarter, dueWD)) : ''}</div>
        </div>
        <span class="overdue-days">${daysOver > 0 ? daysOver + ' day' + (daysOver !== 1 ? 's' : '') + ' overdue' : 'Overdue'}</span>
      </div>`;
    }).join('');
  }
}

// ============================================================
// CALENDAR VIEW
// ============================================================
function renderCalendarView() {
  const container = document.getElementById('view-calendar');
  if (!container) return;

  const quarter = getReadQuarter();
  const sub = document.getElementById('calendar-sub');
  if (sub) sub.textContent = `${quarter || '—'} · Close calendar`;

  const calBody = document.getElementById('cal-view-body');
  if (!calBody) return;

  if (!STATE.calendar.length) {
    calBody.innerHTML = '<p style="font-size:13px;color:var(--slate)">No calendar rows set up yet. Go to Admin → Close Calendar → Setup Calendar.</p>';
    return;
  }

  // Build a map from date string → calendar row for fast lookup
  const byDate = {};
  STATE.calendar.forEach(c => { if (c.ActualDate) byDate[c.ActualDate] = c; });

  // Find the Monday of the week containing the first workday
  const firstDate = new Date(STATE.calendar[0].ActualDate + 'T12:00:00');
  const lastDate  = new Date(STATE.calendar[STATE.calendar.length - 1].ActualDate + 'T12:00:00');
  const today     = todayET();

  // Rewind to Monday of the first week
  const startMonday = new Date(firstDate);
  const dow = startMonday.getDay(); // 0=Sun,1=Mon,...
  const daysBack = dow === 0 ? 6 : dow - 1;
  startMonday.setDate(startMonday.getDate() - daysBack);

  // Forward to Sunday of the last week
  const endSunday = new Date(lastDate);
  const dowLast = endSunday.getDay();
  const daysForward = dowLast === 0 ? 0 : 7 - dowLast;
  endSunday.setDate(endSunday.getDate() + daysForward);

  let html = `
    <div class="cal-view-legend">
      <div class="cal-view-legend-item"><span class="milestone-std" style="padding:2px 8px;border-radius:8px">Standard</span>&nbsp;Meetings &amp; filings</div>
      <div class="cal-view-legend-item"><span class="milestone-svp" style="padding:2px 8px;border-radius:8px">SVP</span>&nbsp;SVP deliverables</div>
      <div class="cal-view-legend-item"><span class="milestone-md" style="padding:2px 8px;border-radius:8px">MD</span>&nbsp;MD deliverables</div>
      <div class="cal-view-legend-item"><span class="milestone-cfo" style="padding:2px 8px;border-radius:8px">CFO</span>&nbsp;CFO deliverables</div>
    </div>
    <div class="cal-dow-header">
      <div class="cal-dow-label">Mon</div>
      <div class="cal-dow-label">Tue</div>
      <div class="cal-dow-label">Wed</div>
      <div class="cal-dow-label">Thu</div>
      <div class="cal-dow-label">Fri</div>
      <div class="cal-dow-label wknd">Sat</div>
      <div class="cal-dow-label wknd">Sun</div>
    </div>`;

  // Walk week by week
  // Convert a Date to YYYY-MM-DD in Eastern Time — consistent with todayET()
  // and all other date strings stored in SharePoint.
  const toETDateStr = (d) => {
    const et = new Date(d.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
    const y   = et.getFullYear();
    const mo  = String(et.getMonth() + 1).padStart(2, '0');
    const day = String(et.getDate()).padStart(2, '0');
    return `${y}-${mo}-${day}`;
  };

  let cursor = new Date(startMonday);
  while (cursor <= endSunday) {
    html += '<div class="cal-week-row">';
    for (let d = 0; d < 7; d++) {
      const dateStr = toETDateStr(cursor);
      const calRow  = byDate[dateStr];
      const isToday = dateStr === today;
      const isPast  = dateStr < today;

      if (!calRow) {
        // Non-workday — empty cell
        html += '<div class="cal-day empty"></div>';
      } else {
        const cls = [
          'cal-day',
          isPast && !isToday ? 'past' : '',
          isToday ? 'today' : '',
          calRow.IsWeekend ? 'wknd-wd' : '',
        ].filter(Boolean).join(' ');

        html += `<div class="${cls}">
          <div class="cal-day-top">
            <span class="cal-day-wd">WD${calRow.WorkdayNumber}${isToday ? '<span class="cal-today-dot"></span>' : ''}</span>
            <span class="cal-day-date">${formatDateShort(dateStr + 'T12:00:00')}</span>
          </div>
          ${calRow.IsWeekend ? '<span class="cal-wknd-flag">Weekend</span>' : ''}
          ${calRow.MilestoneLabel
            ? `<span class="cal-ms ${milestoneClass(calRow)}">${escapeHtml(calRow.MilestoneLabel)}</span>`
            : ''}
        </div>`;
      }

      cursor.setDate(cursor.getDate() + 1);
    }
    html += '</div>';
  }

  calBody.innerHTML = html;
}

// ============================================================
// ADMIN VIEW
// ============================================================
function renderAdminView() {
  if (!STATE.isAdmin) {
    const content = document.getElementById('admin-content');
    if (content) content.innerHTML = '<p style="color:var(--red)">Access denied.</p>';
    return;
  }
  renderAdminPanel('overview');
}

function renderAdminPanel(panelName) {
  const content = document.getElementById('admin-content');
  if (!content) return;

  // Update sidebar active state
  document.querySelectorAll('.sidebar-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.panel === panelName);
  });

  switch (panelName) {
    case 'overview':    content.innerHTML = renderAdminOverview();    break;
    case 'calendar':    content.innerHTML = renderAdminCalendar();    break;
    case 'rollforward': content.innerHTML = renderAdminRollforward(); break;
    case 'templates':   content.innerHTML = renderAdminTemplates();   break;
    case 'suggestions':
      content.innerHTML = '<p style="font-size:12px;color:var(--slate);padding:12px">Loading...</p>';
      loadSuggestions().then(() => {
        // Guard: only write if the user hasn't navigated to a different panel while loading.
        const activeBtn = document.querySelector('.sidebar-btn.active');
        if (activeBtn?.dataset.panel === 'suggestions') {
          content.innerHTML = renderAdminSuggestions();
          attachAdminEvents('suggestions');
        }
      });
      return; // early return — attachAdminEvents called in callback above
    case 'users':       content.innerHTML = renderAdminUsers();       break;
    case 'auditlog':
      content.innerHTML = '<p style="font-size:12px;color:var(--slate);padding:12px">Loading audit log...</p>';
      loadAuditLogEntries().then(() => {
        const activeBtn = document.querySelector('.sidebar-btn.active');
        if (activeBtn?.dataset.panel === 'auditlog') {
          content.innerHTML = renderAdminAuditLog();
          attachAdminEvents('auditlog');
        }
      });
      return;
    case 'import':      content.innerHTML = renderAdminImport();      break;
    default: content.innerHTML = '';
  }
  attachAdminEvents(panelName);
}

function renderAdminOverview() {
  return `
    <div class="admin-section-title">Admin Overview</div>
    <div class="admin-section-sub">${STATE.activeQuarter || 'No active quarter'} · Folio v${CONFIG.version}</div>
    <div class="quarter-status-bar">
      <div class="quarter-pills">
        <div>
          <div class="quarter-pill-label">Live quarter</div>
          <span class="pill-live">${STATE.activeQuarter || 'None'}</span>
        </div>
        <div class="quarter-divider"></div>
        <div>
          <div class="quarter-pill-label">Staging quarter</div>
          <span class="pill-staging">${STATE.workingQuarter || 'None'}</span>
        </div>
      </div>
      <div style="display:flex;gap:6px">
        ${STATE.workingQuarter ? `<button class="btn-secondary btn-sm" id="btn-edit-staging">Edit staging</button>
        <button class="btn-success btn-sm" id="btn-activate-quarter">Activate ${STATE.workingQuarter}</button>` : ''}
      </div>
    </div>
    <div class="card" style="margin-bottom:12px">
      <div class="card-title" style="display:flex;align-items:center;justify-content:space-between">
        System diagnostics
        <button class="btn-secondary btn-sm" id="btn-run-diagnostics">Run diagnostics</button>
      </div>
      <div class="diag-grid" id="diag-results">
        <div class="diag-item"><div class="diag-dot dot-amber"></div><div class="diag-name">Run diagnostics to check all connections</div></div>
      </div>
    </div>
    <div class="card">
      <div class="card-title">Close calendar — ${STATE.activeQuarter || '—'}</div>
      ${renderCalendarPreview()}
    </div>`;
}

function renderCalendarPreview() {
  const today = todayET();
  const items = STATE.calendar.filter(c => c.MilestoneLabel).slice(0, 8);
  if (!items.length) return '<p style="font-size:11px;color:var(--slate)">No milestones set. Go to Close Calendar to configure.</p>';
  return `<table class="cal-table">
    <thead><tr><th>WD</th><th>Date</th><th>Milestone</th></tr></thead>
    <tbody>${items.map(m => `
      <tr ${m.ActualDate === today ? 'class="today-row"' : ''}>
        <td style="font-weight:500">WD${m.WorkdayNumber}</td>
        <td style="color:var(--slate)">${formatDateShort(m.ActualDate)}</td>
        <td>
          <span class="${milestoneClass(m)}">${escapeHtml(m.MilestoneLabel)}</span>
          ${m.ActualDate === today ? '<span class="today-marker" style="margin-left:4px">Today</span>' : ''}
          ${m.IsWeekend ? '<span class="weekend-marker" style="margin-left:4px">Weekend</span>' : ''}
        </td>
      </tr>`).join('')}
    </tbody></table>`;
}

function renderAdminCalendar() {
  const quarter = STATE.activeQuarter;
  const hasRows = STATE.calendar.length > 0;
  return `
    <div class="admin-section-title">Close Calendar</div>
    <div class="admin-section-sub">${quarter || 'No active quarter'}</div>
    <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center">
      <button class="btn-primary btn-sm" id="btn-setup-calendar">Setup Calendar…</button>
      <span style="font-size:11px;color:var(--slate)">${hasRows ? `${STATE.calendar.length} workdays configured` : 'No workdays set up yet — click Setup Calendar to create them'}</span>
    </div>
    <div class="card">
      ${!hasRows ? `<p style="font-size:12px;color:var(--slate);padding:8px 0">No calendar rows yet. Click <strong>Setup Calendar</strong> to create workday rows for this quarter.</p>` : `
      <table class="cal-table">
        <thead><tr><th>WD</th><th>Date</th><th>Milestone</th><th>Actions</th></tr></thead>
        <tbody>
          ${STATE.calendar.map(c => `
            <tr ${c.ActualDate === todayET() ? 'class="today-row"' : ''}>
              <td style="font-weight:500">WD${c.WorkdayNumber}</td>
              <td style="color:var(--slate)">${formatDateShort(c.ActualDate)}</td>
              <td>
                ${c.MilestoneLabel ? `<span class="${milestoneClass(c)}">${escapeHtml(c.MilestoneLabel)}</span>` : '—'}
                ${c.ActualDate === todayET() ? '<span class="today-marker" style="margin-left:4px">Today</span>' : ''}
                ${c.IsWeekend ? '<span class="weekend-marker" style="margin-left:4px">Weekend</span>' : ''}
              </td>
              <td><button class="btn-secondary btn-sm" data-action="edit-cal-row" data-id="${c._id}">Edit</button></td>
            </tr>`).join('')}
        </tbody>
      </table>`}
    </div>`;
}

function renderAdminRollforward() {
  return `
    <div class="admin-section-title">Quarterly Rollforward</div>
    <div class="admin-section-sub">Stage and activate a new quarter</div>
    <div class="card">
      <div class="card-title">Current status</div>
      <p style="font-size:13px;margin-bottom:12px">Live quarter: <strong>${STATE.activeQuarter || 'None'}</strong> &nbsp;·&nbsp; Staging: <strong>${STATE.workingQuarter || 'None'}</strong></p>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button class="btn-primary btn-sm" id="btn-start-new-quarter">Start New Quarter</button>
        ${STATE.workingQuarter ? `<button class="btn-secondary btn-sm" id="btn-rollforward">Roll Forward from ${STATE.activeQuarter || 'previous'}</button>` : ''}
        ${STATE.workingQuarter ? `<button class="btn-success btn-sm" id="btn-activate-quarter-rf">Activate ${STATE.workingQuarter}</button>` : ''}
      </div>
    </div>
    ${STATE.workingQuarter ? renderStagingGrid() : ''}`;
}

function renderStagingGrid() {
  const stagingItems = STATE._stagingItems.filter(
    a => a.Quarter === STATE.workingQuarter
  );

  // Load staging items from SharePoint if not already in STATE
  // (STATE.assignments only holds active quarter — staging is a different quarter)
  if (!stagingItems.length) {
    // Trigger async load and re-render
    if (!STATE._stagingLoading) {
      STATE._stagingLoading = true;
      getListItems(CONFIG.lists.quarterlyAssignments,
        `fields/Quarter eq '${STATE.workingQuarter}' and fields/IsStaging eq true`
      ).then(items => {
        STATE._stagingItems = items.map(i => ({ ...i.fields, _id: i.id }));
        STATE._stagingLoading = false;
        renderAdminPanel('rollforward');
      }).catch(() => { STATE._stagingLoading = false; });
    }
    return `<div class="card">
      <div class="card-title">Staging grid — ${STATE.workingQuarter}</div>
      <p style="font-size:12px;color:var(--slate)">
        ${STATE._stagingLoading ? 'Loading staging assignments...' : 'Click "Roll Forward" to populate staging assignments, then review them here.'}
      </p>
    </div>`;
  }

  // Build user dropdown options
  const userOpts = STATE.users
    .filter(u => u.IsActive !== false)
    .sort((a, b) => (a.Title || '').localeCompare(b.Title || ''))
    .map(u => `<option value="${escapeHtml(u.Email)}">${escapeHtml((u.Emoji || '') + ' ' + (u.Title || u.Email.split('@')[0]))}</option>`)
    .join('');
  const blankOpt = '<option value="">— Unassigned —</option>';

  const rows = stagingItems
    .sort((a, b) => (a.Category || '').localeCompare(b.Category || '') || (a.Title || '').localeCompare(b.Title || ''))
    .map(item => `
      <tr>
        <td style="font-size:11px;max-width:160px">${escapeHtml(item.Title || '')}</td>
        <td><span class="cat-tag">${escapeHtml(item.Category || '')}</span></td>
        <td>
          <input type="number" class="staging-select staging-wd" data-id="${item._id}" data-field="PreparerWorkday"
            value="${item.PreparerWorkday || ''}" min="1" max="35"
            style="font-size:11px;width:52px;text-align:center" title="Preparer workday" />
        </td>
        <td>
          ${item.SignOffMode === 'Preparer Only'
            ? '<span style="font-size:10px;color:var(--slate)">—</span>'
            : `<input type="number" class="staging-select staging-wd" data-id="${item._id}" data-field="ReviewerWorkday"
                value="${item.ReviewerWorkday || ''}" min="1" max="35"
                style="font-size:11px;width:52px;text-align:center" title="Reviewer workday" />`}
        </td>
        <td>
          <select class="staging-select" data-id="${item._id}" data-field="Preparer" style="font-size:11px;max-width:130px">
            ${blankOpt}${userOpts.replace(`value="${escapeHtml(item.Preparer)}"`, `value="${escapeHtml(item.Preparer)}" selected`)}
          </select>
        </td>
        <td>
          ${item.SignOffMode === 'Preparer Only'
            ? '<span style="font-size:10px;color:var(--slate)">Prep only</span>'
            : `<select class="staging-select" data-id="${item._id}" data-field="Reviewer" style="font-size:11px;max-width:130px">
                ${blankOpt}${userOpts.replace(`value="${escapeHtml(item.Reviewer)}"`, `value="${escapeHtml(item.Reviewer)}" selected`)}
              </select>`}
        </td>
      </tr>`).join('');

  return `
    <div class="card">
      <div class="card-title" style="display:flex;align-items:center;justify-content:space-between">
        Staging grid — ${STATE.workingQuarter}
        <span style="font-size:11px;font-weight:400;color:var(--slate)">${stagingItems.length} assignments · changes save instantly</span>
      </div>
      <p style="font-size:12px;color:var(--slate);margin-bottom:10px">Review and adjust workday numbers, preparers, and reviewers before activating. Changes here only affect the staging quarter.</p>
      <div class="table-wrap">
        <table class="data-table" style="table-layout:fixed;width:100%">
          <colgroup>
            <col style="width:22%"/><col style="width:12%"/><col style="width:7%"/>
            <col style="width:7%"/><col style="width:26%"/><col style="width:26%"/>
          </colgroup>
          <thead><tr>
            <th>Task</th><th>Category</th>
            <th title="Preparer Workday">Prep WD</th>
            <th title="Reviewer Workday">Rev WD</th>
            <th>Preparer</th><th>Reviewer</th>
          </tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    </div>`;
}

function renderAdminTemplates() {
  return `
    <div class="admin-section-title">Task Templates</div>
    <div class="admin-section-sub">${STATE.templates.length} active templates</div>
    <div style="display:flex;gap:8px;margin-bottom:12px">
      <button class="btn-primary btn-sm" id="btn-new-template">+ New Template</button>
      <input type="text" class="filter-search" id="template-search" placeholder="Search templates..." style="width:200px"/>
    </div>
    <div class="table-wrap">
      <table class="data-table">
        <thead><tr>
          <th>Task Name</th><th>Category</th><th>Filing</th><th>Sign-off</th>
          <th title="Standard / 10-Q preparer workday">Prep WD</th>
          <th title="Standard / 10-Q reviewer workday">Rev WD</th>
          <th title="10-K preparer workday (Q4 only)">Prep WD <span style="font-size:9px;opacity:0.7">10-K</span></th>
          <th title="10-K reviewer workday (Q4 only)">Rev WD <span style="font-size:9px;opacity:0.7">10-K</span></th>
          <th>Actions</th>
        </tr></thead>
        <tbody>
          ${STATE.templates.map(t => `
            <tr>
              <td style="font-size:12px">${escapeHtml(t.TaskName || t.Title || '')}</td>
              <td><span class="cat-tag">${escapeHtml(t.Category || '')}</span></td>
              <td style="font-size:11px">${escapeHtml(t.FilingType || '')}</td>
              <td style="font-size:11px">${escapeHtml(t.SignOffMode || '')}</td>
              <td style="font-size:11px">WD${t.PreparerWorkday || '—'}</td>
              <td style="font-size:11px">${t.ReviewerWorkday ? 'WD' + t.ReviewerWorkday : '—'}</td>
              <td style="font-size:11px">${t.PreparerWorkday10K ? 'WD' + t.PreparerWorkday10K : '<span style="color:var(--slate)">—</span>'}</td>
              <td style="font-size:11px">${t.ReviewerWorkday10K ? 'WD' + t.ReviewerWorkday10K : '<span style="color:var(--slate)">—</span>'}</td>
              <td style="font-size:11px">
                <button class="btn-icon btn-sm" data-action="edit-template" data-id="${t._id}">Edit</button>
                <button class="btn-danger btn-sm" data-action="retire-template" data-id="${t._id}" style="margin-left:4px">Retire</button>
              </td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
}

function renderAdminSuggestions() {
  // Suggestions are loaded via loadSuggestions() when Admin panel is opened.
  // STATE.suggestions is populated by that call.
  const pending  = (STATE.suggestions || []).filter(s => s.Status === 'Pending');
  const approved = (STATE.suggestions || []).filter(s => s.Status === 'Approved');
  const rejected = (STATE.suggestions || []).filter(s => s.Status === 'Rejected');

  const renderSuggestionRow = (s) => `
    <div class="suggest-item">
      <div>
        <span class="suggest-type-${(s.SuggestionType || '').toLowerCase()}">${escapeHtml(s.SuggestionType || '')}</span>
        <span style="font-size:12px;margin-left:6px;font-weight:500">${escapeHtml(s.Title || '')}</span>
        <div style="font-size:11px;color:var(--slate);margin-top:3px">${escapeHtml(s.ProposedChanges || '')}</div>
        <div style="font-size:10px;color:var(--slate);margin-top:2px">Submitted by ${renderBadge(s.SuggestedBy)}</div>
      </div>
      ${s.Status === 'Pending' ? `
        <div style="display:flex;gap:4px;flex-shrink:0">
          <button class="btn-success btn-sm" data-action="approve-suggestion" data-id="${s._id}">Approve</button>
          <button class="btn-danger btn-sm" data-action="reject-suggestion" data-id="${s._id}">Reject</button>
        </div>` : `<span class="cat-tag">${escapeHtml(s.Status)}</span>`}
    </div>`;

  return `
    <div class="admin-section-title">Task Suggestions</div>
    <div class="admin-section-sub">${pending.length} pending · ${approved.length} approved · ${rejected.length} rejected</div>
    <div id="suggestions-list">
      ${pending.length
        ? pending.map(renderSuggestionRow).join('')
        : '<p style="font-size:12px;color:var(--slate)">No pending suggestions.</p>'}
      ${(approved.length || rejected.length) ? `
        <hr style="margin:14px 0;border:none;border-top:1px solid var(--mid-gray)"/>
        <div style="font-size:11px;font-weight:600;color:var(--slate);margin-bottom:8px">RECENT</div>
        ${[...approved, ...rejected].sort((a,b) => new Date(b.ReviewDate||0) - new Date(a.ReviewDate||0)).slice(0,5).map(renderSuggestionRow).join('')}` : ''}
    </div>`;
}

function renderAdminUsers() {
  return `
    <div class="admin-section-title">Users</div>
    <div class="admin-section-sub">${STATE.users.length} active users</div>
    <div style="margin-bottom:12px">
      <button class="btn-primary btn-sm" id="btn-add-user">+ Add User</button>
    </div>
    <div class="table-wrap">
      <table class="data-table">
        <thead><tr><th>Name</th><th>Email</th><th>Role</th><th>Last Login</th><th>Actions</th></tr></thead>
        <tbody>
          ${STATE.users.map(u => `
            <tr>
              <td>${renderBadge(u.Email)}</td>
              <td style="font-size:11px">${escapeHtml(u.Email || '')}</td>
              <td><span class="cat-tag">${escapeHtml(u.Role || 'TeamMember')}</span></td>
              <td style="font-size:11px">${u.LastLogin ? formatDateShort(u.LastLogin) : '—'}</td>
              <td><button class="btn-secondary btn-sm" data-action="edit-user" data-email="${escapeHtml(u.Email)}">Edit role</button></td>
            </tr>`).join('')}
        </tbody>
      </table>
    </div>`;
}

async function loadAuditLogEntries() {
  // Load all audit entries across all quarters for the viewer.
  // Sorted by ActionDate descending (most recent first).
  const items = await getListItems(CONFIG.lists.auditLog);
  STATE._auditEntries = items
    .map(i => ({ ...i.fields, _id: i.id }))
    .sort((a, b) => new Date(b.ActionDate) - new Date(a.ActionDate));
}

function renderAdminAuditLog() {
  const entries = STATE._auditEntries || [];
  const f = STATE._auditFilter;

  const TYPE_STYLE = {
    SignOff:               'background:#EAF3DE;color:#27500A',
    Reversal:              'background:#FCEBEB;color:#791F1F',
    Reassignment:          'background:#FAEEDA;color:#633806',
    MatrixStatusChange:    'background:#EEEDFE;color:#3C3489',
    CalendarEdit:          'background:#E1F5EE;color:#085041',
    Rollforward:           'background:#E6F1FB;color:#0C447C',
    QuarterActivation:     'background:#E6F1FB;color:#0C447C',
    QuarterCreated:        'background:#E6F1FB;color:#0C447C',
    TaskEdit:              'background:#F1EFE8;color:#444441',
    UserEdit:              'background:#F1EFE8;color:#444441',
    ReviewCommentCreated:  'background:#FBEAF0;color:#72243E',
    ReviewCommentResolved: 'background:#EAF3DE;color:#085041',
    SuggestionApproved:    'background:#EAF3DE;color:#27500A',
    SuggestionRejected:    'background:#FCEBEB;color:#791F1F',
  };

  const allTypes    = ['All','SignOff','Reversal','Reassignment','ReviewCommentCreated',
    'ReviewCommentResolved','MatrixStatusChange','CalendarEdit','Rollforward',
    'QuarterActivation','TaskEdit','UserEdit'];
  const allPeople   = [...new Set(entries.map(e => e.ActionBy).filter(Boolean))].sort();
  const allQuarters = [...new Set(entries.map(e => e.Quarter).filter(Boolean))].sort().reverse();

  let filtered = entries;
  if (f.type && f.type !== 'All') filtered = filtered.filter(e => e.ActionType === f.type);
  if (f.person)  filtered = filtered.filter(e => e.ActionBy === f.person);
  if (f.quarter) filtered = filtered.filter(e => e.Quarter === f.quarter);

  const rows = filtered.slice(0, 200).map(e => {
    const style = TYPE_STYLE[e.ActionType] || 'background:#F1EFE8;color:#444441';
    const label = e.ActionType?.replace(/([A-Z])/g, ' $1').trim() || '';
    const badge = renderBadge(e.ActionBy);
    const detail = [
      e.NewValue,
      e.PreviousValue ? `← ${e.PreviousValue}` : '',
      e.ReasonNote ? `Reason: ${e.ReasonNote}` : '',
    ].filter(Boolean).join('  ·  ');
    return `<tr>
      <td style="font-size:11px;white-space:nowrap">
        <div>${formatDateShort(e.ActionDate)}</div>
        <div style="font-size:10px;color:var(--slate)">${formatDateET(e.ActionDate).split(',')[1]?.trim() || ''}</div>
        ${e.WorkdayNumber ? `<div style="font-size:10px;color:var(--slate)">WD${e.WorkdayNumber}</div>` : ''}
      </td>
      <td><span style="display:inline-block;font-size:10px;font-weight:500;padding:2px 6px;border-radius:99px;white-space:nowrap;${style}">${escapeHtml(label)}</span></td>
      <td style="font-size:11px;max-width:180px;word-break:break-word">${escapeHtml(e.TaskName || '—')}</td>
      <td>${badge}</td>
      <td style="font-size:11px;color:var(--slate);max-width:220px;word-break:break-word">${escapeHtml(detail || '—')}</td>
    </tr>`;
  }).join('');

  return `
    <div class="admin-section-title">Audit Log</div>
    <div class="admin-section-sub">${entries.length} total entries · ${filtered.length} matching${filtered.length > 200 ? ' · showing first 200 — export for full list' : ''}</div>

    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;align-items:center">
      <select class="field-input" id="audit-filter-type" style="width:auto;font-size:11px">
        ${allTypes.map(t => `<option value="${t}" ${f.type===t?'selected':''}>${t==='All'?'All types':t.replace(/([A-Z])/g,' $1').trim()}</option>`).join('')}
      </select>
      <select class="field-input" id="audit-filter-person" style="width:auto;font-size:11px">
        <option value="">All people</option>
        ${allPeople.map(p => `<option value="${escapeHtml(p)}" ${f.person===p?'selected':''}>${escapeHtml(p.split('@')[0])}</option>`).join('')}
      </select>
      <select class="field-input" id="audit-filter-quarter" style="width:auto;font-size:11px">
        <option value="">All quarters</option>
        ${allQuarters.map(q => `<option value="${escapeHtml(q)}" ${f.quarter===q?'selected':''}>${escapeHtml(q)}</option>`).join('')}
      </select>
      <button class="btn-secondary btn-sm" id="btn-export-audit-excel">Export CSV</button>
      <button class="btn-primary btn-sm" id="btn-export-sox">Audit Log Export…</button>
    </div>

    <div class="table-wrap">
      <table class="data-table" style="table-layout:fixed;width:100%">
        <colgroup>
          <col style="width:13%"/><col style="width:16%"/><col style="width:22%"/>
          <col style="width:13%"/><col style="width:36%"/>
        </colgroup>
        <thead><tr>
          <th>Date / WD</th><th>Action</th><th>Task / Subject</th>
          <th>By</th><th>Detail</th>
        </tr></thead>
        <tbody>${rows || '<tr><td colspan="5" style="font-size:12px;color:var(--slate);padding:12px 0">No entries match the current filters.</td></tr>'}</tbody>
      </table>
    </div>`;
}

function renderAdminImport() {
  return `
    <div class="admin-section-title">Bulk Import</div>
    <div class="admin-section-sub">One-time CSV import for TaskTemplates</div>
    <div class="card">
      <div class="card-title">Import TaskTemplates from CSV</div>
      <p style="font-size:12px;color:var(--slate);margin-bottom:12px">Upload a CSV file with your task templates. See the Build Guide Section 8 for the required column format.</p>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <input type="file" id="import-file" accept=".csv" class="field-input" style="width:auto"/>
        <button class="btn-secondary btn-sm" id="btn-validate-import">Validate</button>
        <button class="btn-primary btn-sm" id="btn-run-import" disabled>Import</button>
      </div>
      <div id="import-status" style="margin-top:12px;font-size:12px;color:var(--slate)"></div>
      <div id="import-progress" style="margin-top:8px"></div>
    </div>`;
}

function attachAdminEvents(panelName) {
  // Overview events
  const btnRunDiag = document.getElementById('btn-run-diagnostics');
  if (btnRunDiag) btnRunDiag.addEventListener('click', runDiagnostics);

  // Edit staging button navigates to the rollforward panel where the staging grid lives
  document.getElementById('btn-edit-staging')?.addEventListener('click', () => {
    renderAdminPanel('rollforward');
  });

  const btnActivate = document.getElementById('btn-activate-quarter') || document.getElementById('btn-activate-quarter-rf');
  if (btnActivate) btnActivate.addEventListener('click', () => {
    STATE.pendingActivation = STATE.workingQuarter;
    const titleEl = document.getElementById('activate-modal-title');
    const descEl  = document.getElementById('activate-modal-desc');
    if (titleEl) titleEl.textContent = `Activate ${STATE.workingQuarter}?`;
    if (descEl) descEl.textContent = `This will immediately make ${STATE.workingQuarter} visible to all ${STATE.users.length} team members.`;
    showModal('modal-activate');
  });

  // All admin-content delegated actions (edit-template, retire-template, edit-cal-row,
  // edit-user, rc-reply, approve-suggestion, reject-suggestion) are handled by the
  // unified adminActionsAttached listener below.

  // Rollforward events
  document.getElementById('btn-start-new-quarter')?.addEventListener('click', startNewQuarter);
  document.getElementById('btn-rollforward')?.addEventListener('click', performRollforward);

  // Calendar setup button
  document.getElementById('btn-setup-calendar')?.addEventListener('click', () => {
    const quarterEl = document.getElementById('cal-setup-quarter');
    const maxWDEl   = document.getElementById('cal-setup-maxwd');
    const errEl     = document.getElementById('cal-setup-error');
    if (quarterEl) quarterEl.value = STATE.activeQuarter || '';
    if (maxWDEl)   maxWDEl.value   = isQuarterQ4(STATE.activeQuarter) ? '35' : '20';
    if (errEl)     errEl.classList.add('hidden');
    showModal('modal-cal-setup');
  });

  // Template search
  document.getElementById('template-search')?.addEventListener('input', e => {
    filterTemplateTable(e.target.value);
  });

  // New template button — opens edit modal in create mode (no templateId)
  document.getElementById('btn-new-template')?.addEventListener('click', () => {
    openEditTemplateModal(null);
  });

  // Add user button
  document.getElementById('btn-add-user')?.addEventListener('click', () => {
    const emailEl = document.getElementById('add-user-email');
    const nameEl  = document.getElementById('add-user-name');
    const roleEl  = document.getElementById('add-user-role');
    const errEl   = document.getElementById('add-user-error');
    const customEl = document.getElementById('add-user-emoji-custom');
    if (emailEl)  emailEl.value  = '';
    if (nameEl)   nameEl.value   = '';
    if (roleEl)   roleEl.value   = 'TeamMember';
    if (errEl)    errEl.classList.add('hidden');
    if (customEl) customEl.value = '';

    // Reset preview
    const previewWrap = document.getElementById('add-user-preview-wrap');
    if (previewWrap) previewWrap.style.display = 'none';

    // Init emoji + color pickers — store selections in closure vars
    STATE._addUserEmoji = null;
    STATE._addUserColor = null;

    renderEmojiPicker('add-user-emoji-grid', null, (emoji) => {
      STATE._addUserEmoji = emoji;
      const customEl = document.getElementById('add-user-emoji-custom');
      if (customEl) customEl.value = '';
      updateAddUserPreview();
    });
    renderColorPicker('add-user-color-grid', null, (color) => {
      STATE._addUserColor = color;
      updateAddUserPreview();
    });

    // Custom emoji overrides grid selection
    document.getElementById('add-user-emoji-custom')?.addEventListener('input', function() {
      const val = this.value.trim();
      STATE._addUserEmoji = val || null;
      if (val) {
        document.querySelectorAll('#add-user-emoji-grid .emoji-option')
          .forEach(el => el.classList.remove('selected'));
      }
      updateAddUserPreview();
    });

    // Live preview updates as name/email is typed
    ['add-user-name', 'add-user-email'].forEach(id => {
      document.getElementById(id)?.addEventListener('input', updateAddUserPreview);
    });

    showModal('modal-add-user');
  });

  // Template edit/retire (delegated from admin-content)
  // Suggestion approve/reject already uses a delegated listener on admin-content.
  // We extend the same listener rather than adding another — handled below by
  // checking additional action values in the existing admin-content click handler.
  // Staging grid — save preparer/reviewer on dropdown change
  const adminContent2 = document.getElementById('admin-content');
  if (adminContent2 && !adminContent2.dataset.stagingEventsAttached) {
    adminContent2.dataset.stagingEventsAttached = 'true';
    adminContent2.addEventListener('change', async e => {
      const sel = e.target.closest('.staging-select');
      if (!sel) return;
      const { id, field } = sel.dataset;

      // WD fields are numbers; person fields are strings or null.
      const isWD = field === 'PreparerWorkday' || field === 'ReviewerWorkday';
      const raw  = sel.value;
      const value = isWD
        ? (raw ? Number(raw) : null)
        : (raw || null);

      // Validate WD range
      if (isWD && value !== null && (value < 1 || value > 35)) {
        showToast('Workday must be between 1 and 35', 'error');
        return;
      }

      try {
        await updateListItem(CONFIG.lists.quarterlyAssignments, id, { [field]: value });
        const item = STATE._stagingItems.find(i => i._id === id);
        if (item) item[field] = value;
        showToast(`✓ ${isWD ? field.replace('Workday','') + ' WD' : field} updated`, 'success');
      } catch (err) {
        showToast(`Failed to update ${field}`, 'error');
        logError('Staging grid update failed:', err);
      }
    });
  }

  const adminContentEl = document.getElementById('admin-content');
  if (adminContentEl && !adminContentEl.dataset.adminActionsAttached) {
    adminContentEl.dataset.adminActionsAttached = 'true';
    adminContentEl.addEventListener('click', async e => {
      const btn = e.target.closest('[data-action]');
      if (!btn) return;
      const { action, id, email } = btn.dataset;

      if (action === 'edit-template')      openEditTemplateModal(id);
      if (action === 'retire-template')   await retireTemplate(id);
      if (action === 'edit-cal-row')      openEditCalendarRowModal(id);
      if (action === 'edit-user')         openEditUserRoleModal(email);
      if (action === 'rc-reply')          openRCReplyInput(id);
      if (action === 'approve-suggestion') await approveSuggestion(id);
      if (action === 'reject-suggestion') {
        STATE.pendingSuggestionReject = id;
        const noteEl = document.getElementById('reject-suggestion-note');
        if (noteEl) noteEl.value = '';
        showModal('modal-reject-suggestion');
      }
    });
  }

  // Template edit modal confirm/cancel
  document.getElementById('btn-edit-tpl-save')?.addEventListener('click', saveTemplateEdit);
  document.getElementById('btn-edit-tpl-cancel')?.addEventListener('click', () => {
    hideModal('modal-edit-template');
    STATE.pendingTemplateEdit = null;
  });

  // Calendar edit modal confirm/cancel
  document.getElementById('btn-edit-cal-save')?.addEventListener('click', saveCalendarRowEdit);
  document.getElementById('btn-edit-cal-cancel')?.addEventListener('click', () => {
    hideModal('modal-edit-calendar');
    STATE.pendingCalendarEdit = null;
  });

  // User role modal confirm/cancel
  document.getElementById('btn-edit-user-save')?.addEventListener('click', saveUserRoleEdit);
  document.getElementById('btn-edit-user-cancel')?.addEventListener('click', () => {
    hideModal('modal-edit-user');
    STATE.pendingUserEdit = null;
  });

  // Audit log exports
  document.getElementById('btn-export-audit-excel')?.addEventListener('click', exportAuditLog);
  document.getElementById('btn-export-sox')?.addEventListener('click', () => openSOXExportModal());
  document.getElementById('btn-sox-confirm')?.addEventListener('click', confirmSOXExport);
  document.getElementById('btn-sox-cancel')?.addEventListener('click', () => hideModal('modal-sox-export'));

  // Audit log filter dropdowns — re-render on change
  ['audit-filter-type','audit-filter-person','audit-filter-quarter'].forEach(id => {
    document.getElementById(id)?.addEventListener('change', e => {
      const field = id.replace('audit-filter-', '');
      STATE._auditFilter[field] = e.target.value;
      document.getElementById('admin-content').innerHTML = renderAdminAuditLog();
      attachAdminEvents('auditlog');
    });
  });


  // Import events
  const btnValidate = document.getElementById('btn-validate-import');
  const btnImport   = document.getElementById('btn-run-import');
  if (btnValidate) btnValidate.addEventListener('click', validateImport);
  if (btnImport)   btnImport.addEventListener('click', runImport);
}

// ============================================================
// DIAGNOSTICS
// ============================================================
async function runDiagnostics() {
  const results = document.getElementById('diag-results');
  if (!results) return;
  results.innerHTML = '<div class="diag-item"><div class="diag-dot dot-amber"></div><div class="diag-name">Running diagnostics...</div></div>';

  const rows = [];

  // ── List connectivity ──────────────────────────────────────
  for (const [key, listName] of Object.entries(CONFIG.lists)) {
    try {
      const items = await getListItems(listName);
      rows.push({ name: listName, status: `${items.length} items`, ok: true });
    } catch {
      rows.push({ name: listName, status: 'Error — list not found or no access', ok: false });
    }
  }

  // ── Auth ──────────────────────────────────────────────────
  try {
    await getToken();
    rows.push({ name: 'MSAL auth', status: 'Token valid', ok: true });
  } catch {
    rows.push({ name: 'MSAL auth', status: 'Auth error', ok: false });
  }

  // ── Missing assignments check ─────────────────────────────
  // Every active template should have a QuarterlyAssignment for the active quarter.
  if (STATE.activeQuarter && STATE.templates.length && STATE.assignments.length) {
    const activeTemplateIds = STATE.templates
      .filter(t => t.IsActive !== false)
      .map(t => t._id);
    const assignedTemplateIds = new Set(STATE.assignments.map(a => a.TaskTemplateLookupId));
    const missing = activeTemplateIds.filter(id => !assignedTemplateIds.has(id));
    rows.push({
      name: 'Assignment coverage',
      status: missing.length === 0
        ? `All ${activeTemplateIds.length} active templates have assignments`
        : `${missing.length} active template${missing.length !== 1 ? 's' : ''} have no assignment for ${STATE.activeQuarter}`,
      ok: missing.length === 0,
    });
  }

  // ── Orphaned review comments check ────────────────────────
  // Review comments whose TaskTemplateLookupId no longer matches any known template.
  if (STATE.reviewComments.length && STATE.templates.length) {
    const templateIds = new Set(STATE.templates.map(t => t._id));
    const orphaned = STATE.reviewComments.filter(
      rc => rc.TaskTemplateLookupId && !templateIds.has(rc.TaskTemplateLookupId)
    );
    rows.push({
      name: 'Review comment integrity',
      status: orphaned.length === 0
        ? `All ${STATE.reviewComments.length} review comments reference valid tasks`
        : `${orphaned.length} review comment${orphaned.length !== 1 ? 's' : ''} reference retired or missing tasks`,
      ok: orphaned.length === 0,
    });
  }

  // ── Quarter mismatch check ────────────────────────────────
  // Confirms assignments and calendar entries all belong to the active quarter.
  if (STATE.activeQuarter) {
    const wrongQuarterAssignments = STATE.assignments.filter(
      a => a.Quarter && a.Quarter !== STATE.activeQuarter
    );
    const wrongQuarterCalendar = STATE.calendar.filter(
      c => c.Quarter && c.Quarter !== STATE.activeQuarter
    );
    const mismatch = wrongQuarterAssignments.length + wrongQuarterCalendar.length;
    rows.push({
      name: 'Quarter consistency',
      status: mismatch === 0
        ? `All loaded data matches active quarter (${STATE.activeQuarter})`
        : `${mismatch} record${mismatch !== 1 ? 's' : ''} have a quarter mismatch — reload may be needed`,
      ok: mismatch === 0,
    });
  }

  results.innerHTML = rows.map(r => `
    <div class="diag-item">
      <div class="diag-dot ${r.ok ? 'dot-green' : 'dot-red'}"></div>
      <div class="diag-name">${escapeHtml(r.name)}</div>
      <div class="diag-status">${escapeHtml(r.status)}</div>
    </div>`).join('');
}

// ============================================================
// BULK IMPORT
// ============================================================
// Parses a CSV string into an array of objects keyed by header row.
// Handles Windows (CRLF) and Unix (LF) line endings, quoted fields containing
// commas, and escaped double-quotes inside quoted fields.
function parseCSV(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').filter(l => l.trim());
  if (!lines.length) return [];

  function parseRow(line) {
    const values = [];
    let cur = '';
    let inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQuotes) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') { inQuotes = false; }
        else { cur += ch; }
      } else {
        if (ch === '"') { inQuotes = true; }
        else if (ch === ',') { values.push(cur.trim()); cur = ''; }
        else { cur += ch; }
      }
    }
    values.push(cur.trim());
    return values;
  }

  const headers = parseRow(lines[0]);
  return lines.slice(1).map(line => {
    const vals = parseRow(line);
    const obj = {};
    headers.forEach((h, i) => { obj[h] = vals[i] || ''; });
    return obj;
  });
}

function validateImport() {
  const fileInput = document.getElementById('import-file');
  const status    = document.getElementById('import-status');
  const btnImport = document.getElementById('btn-run-import');
  if (!fileInput?.files?.[0]) {
    if (status) status.textContent = 'Please select a CSV file first.';
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    const rows = parseCSV(e.target.result);
    const required = ['TaskName', 'Category', 'FilingType', 'SignOffMode', 'PreparerWorkday', 'IsActive'];
    const missing = required.filter(r => !rows[0] || !(r in rows[0]));
    if (missing.length) {
      if (status) status.textContent = `❌ Missing required columns: ${missing.join(', ')}`;
      return;
    }
    if (status) status.textContent = `✓ Validation passed. ${rows.length} tasks ready to import.`;
    if (btnImport) { btnImport.disabled = false; btnImport.dataset.rows = JSON.stringify(rows); }
  };
  reader.readAsText(fileInput.files[0]);
}

async function runImport() {
  const btnImport = document.getElementById('btn-run-import');
  const status    = document.getElementById('import-status');
  const progress  = document.getElementById('import-progress');
  if (!btnImport?.dataset.rows) return;

  const rows = JSON.parse(btnImport.dataset.rows);
  btnImport.disabled = true;
  let imported = 0, failed = 0;

  const batchSize = 20;
  for (let i = 0; i < rows.length; i += batchSize) {
    const batch = rows.slice(i, i + batchSize);
    for (const row of batch) {
      try {
        await createListItem(CONFIG.lists.taskTemplates, {
          Title:           row.TaskName || row.Title || '',
          TaskName:        row.TaskName || row.Title || '',
          Category:        row.Category || '',
          MatrixItem:      row.MatrixItem || null,
          MatrixCheckpoint:row.MatrixCheckpoint || null,
          MatrixSection:   row.MatrixSection || null,
          FilingType:      row.FilingType || 'Both',
          SignOffMode:     row.SignOffMode || 'Sequential',
          PreparerWorkday: Number(row.PreparerWorkday) || 1,
          ReviewerWorkday: row.ReviewerWorkday ? Number(row.ReviewerWorkday) : null,
          DefaultPreparer:     row.DefaultPreparer || null,
          DefaultReviewer:     row.DefaultReviewer || null,
          PreparerWorkday10K:  row.PreparerWorkday10K ? Number(row.PreparerWorkday10K) : null,
          ReviewerWorkday10K:  row.ReviewerWorkday10K ? Number(row.ReviewerWorkday10K) : null,
          HasDocumentLink:     row.HasDocumentLink === 'Yes',
          IsActive:        row.IsActive !== 'No',
        });
        imported++;
      } catch (err) {
        logError('Import failed for row:', row, err);
        failed++;
      }
    }
    const pct = Math.round(((i + batchSize) / rows.length) * 100);
    if (progress) progress.innerHTML = `
      <div class="prog-row">
        <div class="prog-bar-wrap"><div class="prog-bar" style="width:${Math.min(pct,100)}%"></div></div>
        <div class="prog-pct">${Math.min(pct,100)}%</div>
      </div>`;
    if (status) status.textContent = `Imported ${imported} tasks...${failed ? ` (${failed} failed)` : ''}`;
    await sleep(200);
  }

  if (status) status.textContent = `✓ Import complete. ${imported} tasks imported.${failed ? ` ${failed} failed.` : ''}`;
  try {
    await loadTemplates();
  } catch (err) {
    logError('Failed to refresh template cache after import:', err);
    showToast('Import complete but template list may be stale — refresh the page to update', '');
  }
}

// ============================================================
// PROFILE VIEW
// ============================================================
function renderProfileView() {
  const u = STATE.currentUser;
  if (!u) return;

  const nameEl = document.getElementById('profile-name');
  if (nameEl) nameEl.value = u.Title || '';

  renderEmojiPicker('profile-emoji-grid', u.Emoji, (emoji) => {
    STATE.currentUser.Emoji = emoji;
    updateProfilePreview();
  });
  renderColorPicker('profile-color-grid', u.Color, (color) => {
    STATE.currentUser.Color = color;
    updateProfilePreview();
  });
  updateProfilePreview();

  // Notification prefs
  const notifList = document.getElementById('notif-prefs-list');
  if (notifList) {
    const prefs = [
      { key: 'NotifyOnAssignment', label: 'Task assigned to me (quarter activation)' },
      { key: 'NotifyOnReviewUnlock', label: 'Task ready for my review' },
      { key: 'NotifyOnOverdue', label: 'Task overdue' },
      { key: 'NotifyOnReassignment', label: 'Task reassigned to me' },
      { key: 'NotifyOnSuggestionUpdate', label: 'My suggestion approved/rejected' },
    ];
    notifList.innerHTML = prefs.map(p => `
      <div class="notif-row">
        <span>${escapeHtml(p.label)}</span>
        <input type="checkbox" ${u[p.key] === true ? 'checked' : ''} data-pref="${p.key}"/>
      </div>`).join('');
  }

  // Quiet hours
  const qStart = document.getElementById('quiet-start');
  const qEnd   = document.getElementById('quiet-end');
  if (qStart && u.QuietHoursStart) qStart.value = u.QuietHoursStart;
  if (qEnd   && u.QuietHoursEnd)   qEnd.value   = u.QuietHoursEnd;
}

function updateProfilePreview() {
  const badge = document.getElementById('profile-preview-badge');
  const u = STATE.currentUser;
  if (!badge || !u) return;
  const hex = u.Color || '#75787B';
  badge.style.background = hex + '22';
  badge.style.color = hex;
  badge.textContent = `${u.Emoji || '?'} ${u.Title || ''}`;
}

async function saveProfile() {
  const u = STATE.currentUser;
  if (!u) return;

  const nameEl = document.getElementById('profile-name');
  if (nameEl) u.Title = nameEl.value.trim() || u.Title;

  const customEmoji = document.getElementById('profile-emoji-custom');
  if (customEmoji?.value?.trim()) u.Emoji = customEmoji.value.trim();

  const quietStart = document.getElementById('quiet-start');
  const quietEnd   = document.getElementById('quiet-end');
  if (quietStart) u.QuietHoursStart = quietStart.value;
  if (quietEnd)   u.QuietHoursEnd   = quietEnd.value;

  const notifCheckboxes = document.querySelectorAll('[data-pref]');
  notifCheckboxes.forEach(cb => { u[cb.dataset.pref] = cb.checked; });

  try {
    await updateListItem(CONFIG.lists.users, u._id, {
      Title:                    u.Title,
      Emoji:                    u.Emoji,
      Color:                    u.Color,
      QuietHoursStart:          u.QuietHoursStart || null,
      QuietHoursEnd:            u.QuietHoursEnd || null,
      NotifyOnAssignment:       u.NotifyOnAssignment === true,
      NotifyOnReviewUnlock:     u.NotifyOnReviewUnlock === true,
      NotifyOnOverdue:          u.NotifyOnOverdue === true,
      NotifyOnReassignment:     u.NotifyOnReassignment === true,
      NotifyOnSuggestionUpdate: u.NotifyOnSuggestionUpdate === true,
    });
    updateNavAvatar();
    showToast('✓ Profile saved', 'success');
  } catch (err) {
    showToast('Failed to save profile', 'error');
    logError('Profile save failed:', err);
  }
}

// ============================================================
// EMOJI & COLOR PICKERS
// ============================================================
function renderEmojiPicker(containerId, selected, onChange) {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.innerHTML = CONFIG.emojiOptions.map(e => `
    <div class="emoji-option ${e === selected ? 'selected' : ''}" data-emoji="${e}">${e}</div>`).join('');
  container.querySelectorAll('.emoji-option').forEach(el => {
    el.addEventListener('click', () => {
      container.querySelectorAll('.emoji-option').forEach(e => e.classList.remove('selected'));
      el.classList.add('selected');
      onChange(el.dataset.emoji);
    });
  });
}

function renderColorPicker(containerId, selected, onChange) {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.innerHTML = CONFIG.colorOptions.map(c => `
    <div class="color-option ${c.hex === selected ? 'selected' : ''}" data-hex="${c.hex}" style="background:${c.hex}" title="${c.label}"></div>`).join('');
  container.querySelectorAll('.color-option').forEach(el => {
    el.addEventListener('click', () => {
      container.querySelectorAll('.color-option').forEach(e => e.classList.remove('selected'));
      el.classList.add('selected');
      onChange(el.dataset.hex);
    });
  });
}

// ============================================================
// MODALS
// ============================================================
// ── Modal focus management ───────────────────────────────────
// Tracks the element that triggered the modal so focus can be restored on close.
let _modalTrigger = null;
let _modalKeyHandler = null;

// Focusable element selector — covers all interactive elements inside a modal.
const FOCUSABLE = 'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';

function trapFocus(modalEl) {
  const focusable = Array.from(modalEl.querySelectorAll(FOCUSABLE));
  if (!focusable.length) return;
  const first = focusable[0];
  const last  = focusable[focusable.length - 1];

  // Move focus into the modal — prefer the first interactive element.
  first.focus();

  // Remove any previous key handler before adding a new one.
  if (_modalKeyHandler) document.removeEventListener('keydown', _modalKeyHandler);

  _modalKeyHandler = (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      hideAllModals();
      return;
    }
    if (e.key !== 'Tab') return;

    // Keep Tab cycling inside the modal.
    if (e.shiftKey) {
      if (document.activeElement === first) { e.preventDefault(); last.focus(); }
    } else {
      if (document.activeElement === last)  { e.preventDefault(); first.focus(); }
    }
  };

  document.addEventListener('keydown', _modalKeyHandler);
}

function releaseFocus() {
  if (_modalKeyHandler) {
    document.removeEventListener('keydown', _modalKeyHandler);
    _modalKeyHandler = null;
  }
  // Return focus to the element that opened the modal.
  if (_modalTrigger && typeof _modalTrigger.focus === 'function') {
    _modalTrigger.focus();
  }
  _modalTrigger = null;
}

function showModal(modalId) {
  // Record what triggered the modal so we can restore focus on close.
  _modalTrigger = document.activeElement;

  const modal = document.getElementById(modalId);
  if (!modal) return;
  modal.classList.remove('hidden');
  document.getElementById('modal-backdrop')?.classList.remove('hidden');

  // Trap focus inside the modal box.
  trapFocus(modal);
}

function hideModal(modalId) {
  document.getElementById(modalId)?.classList.add('hidden');
  document.getElementById('modal-backdrop')?.classList.add('hidden');
  releaseFocus();
}

function hideAllModals() {
  document.querySelectorAll('.modal').forEach(m => m.classList.add('hidden'));
  document.getElementById('modal-backdrop')?.classList.add('hidden');
  releaseFocus();
}

// ============================================================
// TOAST
// ============================================================
function showToast(message, type = '') {
  const toast = document.getElementById('toast');
  if (!toast) return;
  toast.textContent = message;
  toast.className = `toast ${type}`;
  toast.classList.remove('hidden');
  setTimeout(() => toast.classList.add('hidden'), 3000);
}

// ============================================================
// LOADING
// ============================================================
function showLoading(text = 'Loading...') {
  document.getElementById('loading-text').textContent = text;
  document.getElementById('loading-overlay')?.classList.remove('hidden');
}
function hideLoading() {
  document.getElementById('loading-overlay')?.classList.add('hidden');
}

// ============================================================
// STALE DATA BANNER
// ============================================================
function showStaleBanner(show) {
  document.getElementById('stale-banner')?.classList.toggle('hidden', !show);
}

// ============================================================
// NAV AVATAR
// ============================================================
function updateNavAvatar() {
  const btn = document.getElementById('nav-user-avatar');
  if (!btn || !STATE.currentUser) return;
  const u = STATE.currentUser;
  btn.textContent = u.Emoji || u.Title?.[0] || '?';
  const hex = u.Color || '#75787B';
  btn.style.background = hex + '33';
  btn.style.color = hex;
}

// ============================================================
// EVENTS — GLOBAL
// ============================================================
function attachGlobalEvents() {
  // Nav links
  document.querySelectorAll('.nav-link').forEach(btn => {
    btn.addEventListener('click', () => showView(btn.dataset.view));
  });

  // Admin sidebar
  document.addEventListener('click', e => {
    const btn = e.target.closest('[data-panel]');
    if (btn) renderAdminPanel(btn.dataset.panel);
  });

  // Refresh button
  document.getElementById('btn-refresh')?.addEventListener('click', async () => {
    const btn = document.getElementById('btn-refresh');
    btn?.classList.add('spinning');
    try {
      // Refresh the currently viewed quarter, not necessarily the live one.
      await loadViewingQuarterData(getReadQuarter());
      refreshCurrentView();
      updateHistoryBanner();
      showStaleBanner(false);
    } catch { showStaleBanner(true); }
    btn?.classList.remove('spinning');
  });

  // Return to live quarter button
  document.getElementById('btn-return-live')?.addEventListener('click', () => {
    switchToQuarter(STATE.activeQuarter);
    const sel = document.getElementById('quarter-picker');
    if (sel) sel.value = STATE.activeQuarter;
  });

  // Stale retry
  document.getElementById('btn-stale-retry')?.addEventListener('click', async () => {
    try { await loadAllData(); refreshCurrentView(); showStaleBanner(false); }
    catch { /* stay stale */ }
  });

  // Profile save
  document.getElementById('btn-save-profile')?.addEventListener('click', saveProfile);

  // Nav user avatar → profile
  document.getElementById('nav-user-avatar')?.addEventListener('click', () => showView('profile'));

  // Panel close
  document.getElementById('panel-close')?.addEventListener('click', closeTaskPanel);
  document.getElementById('panel-overlay')?.addEventListener('click', closeTaskPanel);

  // Panel review comments link
  document.getElementById('panel-rc-link')?.addEventListener('click', () => {
    closeTaskPanel();
    showView('review-comments');
  });

  // Modal backdrop
  document.getElementById('modal-backdrop')?.addEventListener('click', hideAllModals);

  // Sign-off modal
  document.getElementById('btn-signoff-confirm')?.addEventListener('click', async () => {
    if (!STATE.pendingSignoff) return;
    hideModal('modal-signoff');
    await performSignOff(STATE.pendingSignoff.assignmentId, STATE.pendingSignoff.role);
    STATE.pendingSignoff = null;
    if (STATE.taskDetailId) openTaskPanel(STATE.taskDetailId);
  });
  document.getElementById('btn-signoff-cancel')?.addEventListener('click', () => {
    hideModal('modal-signoff');
    STATE.pendingSignoff = null;
  });

  // Reversal modal
  document.getElementById('btn-reversal-confirm')?.addEventListener('click', async () => {
    const reason = document.getElementById('reversal-reason')?.value?.trim();
    if (!reason) {
      document.getElementById('reversal-error')?.classList.remove('hidden');
      return;
    }
    if (!STATE.pendingReversal) return;
    hideModal('modal-reversal');
    await performReversal(STATE.pendingReversal.assignmentId, STATE.pendingReversal.role, reason);
    STATE.pendingReversal = null;
    if (STATE.taskDetailId) openTaskPanel(STATE.taskDetailId);
  });
  document.getElementById('btn-reversal-cancel')?.addEventListener('click', () => {
    hideModal('modal-reversal');
    STATE.pendingReversal = null;
  });

  // Review comment modal
  document.getElementById('btn-new-rc')?.addEventListener('click', () => {
    // Only reviewers and admins can create review comments.
    if (!STATE.isFinalReviewer && !STATE.isAdmin) {
      showToast('Only reviewers and admins can post review comments', 'error');
      return;
    }
    const sel = document.getElementById('rc-task-select');
    if (sel) sel.innerHTML = STATE.templates.map(t =>
      `<option value="${escapeHtml(t._id)}">${escapeHtml(t.TaskName || t.Title || '')}</option>`
    ).join('');
    showModal('modal-new-rc');
  });
  document.getElementById('btn-rc-save')?.addEventListener('click', saveReviewComment);
  document.getElementById('btn-rc-cancel')?.addEventListener('click', () => hideModal('modal-new-rc'));

  // Suggest modal
  document.getElementById('btn-suggest-change')?.addEventListener('click', () => {
    const sel = document.getElementById('suggest-task-select');
    if (sel) sel.innerHTML = STATE.templates.map(t =>
      `<option value="${escapeHtml(t._id)}">${escapeHtml(t.TaskName || t.Title || '')}</option>`
    ).join('');
    showModal('modal-suggest');
  });
  document.getElementById('btn-suggest-save')?.addEventListener('click', saveSuggestion);
  document.getElementById('btn-suggest-cancel')?.addEventListener('click', () => hideModal('modal-suggest'));

  // Matrix modal
  document.getElementById('btn-matrix-confirm')?.addEventListener('click', async () => {
    if (!STATE.pendingMatrixAction) return;
    const selected = document.querySelector('input[name="matrix-action"]:checked')?.value;
    hideModal('modal-matrix-action');
    await performMatrixUpdate(STATE.pendingMatrixAction.item, STATE.pendingMatrixAction.col, selected);
    STATE.pendingMatrixAction = null;
  });
  document.getElementById('btn-matrix-cancel')?.addEventListener('click', () => {
    hideModal('modal-matrix-action');
    STATE.pendingMatrixAction = null;
  });

  // Resolve RC modal
  document.getElementById('btn-resolve-rc-confirm')?.addEventListener('click', async () => {
    if (!STATE.pendingRCResolve) return;
    const note = document.getElementById('resolve-rc-note')?.value?.trim() || '';
    hideModal('modal-resolve-rc');
    await confirmResolveReviewComment(STATE.pendingRCResolve, note);
    STATE.pendingRCResolve = null;
  });
  document.getElementById('btn-resolve-rc-cancel')?.addEventListener('click', () => {
    hideModal('modal-resolve-rc');
    STATE.pendingRCResolve = null;
  });

  // Reject suggestion modal
  document.getElementById('btn-reject-suggestion-confirm')?.addEventListener('click', async () => {
    if (!STATE.pendingSuggestionReject) return;
    const note = document.getElementById('reject-suggestion-note')?.value?.trim() || '';
    hideModal('modal-reject-suggestion');
    await rejectSuggestion(STATE.pendingSuggestionReject, note);
    STATE.pendingSuggestionReject = null;
  });
  document.getElementById('btn-reject-suggestion-cancel')?.addEventListener('click', () => {
    hideModal('modal-reject-suggestion');
    STATE.pendingSuggestionReject = null;
  });

  // New quarter modal
  document.getElementById('btn-new-quarter-confirm')?.addEventListener('click', confirmNewQuarter);
  document.getElementById('btn-new-quarter-cancel')?.addEventListener('click', () => hideModal('modal-new-quarter'));
  document.getElementById('new-quarter-name')?.addEventListener('keydown', e => {
    if (e.key === 'Enter') confirmNewQuarter();
  });

  // Rollforward confirm modal
  document.getElementById('btn-rollforward-confirm')?.addEventListener('click', confirmRollforward);
  document.getElementById('btn-rollforward-cancel')?.addEventListener('click', () => {
    hideModal('modal-rollforward-confirm');
    STATE.pendingRollforward = null;
  });

  // Reassign modal
  document.getElementById('btn-reassign-confirm')?.addEventListener('click', confirmReassign);
  document.getElementById('btn-reassign-cancel')?.addEventListener('click', () => {
    hideModal('modal-reassign');
    STATE.pendingReassign = null;
  });

  // Calendar bulk setup modal
  document.getElementById('btn-cal-setup-confirm')?.addEventListener('click', setupCalendarBulk);
  document.getElementById('btn-cal-setup-cancel')?.addEventListener('click', () => hideModal('modal-cal-setup'));

  // Cascade modal
  document.getElementById('btn-cascade-confirm')?.addEventListener('click', confirmCascade);
  document.getElementById('btn-cascade-no')?.addEventListener('click', () => {
    hideModal('modal-cascade');
    STATE.pendingCascade = null;
    showToast('✓ Calendar row updated', 'success');
    renderAdminPanel('calendar');
  });

  // Add user modal
  document.getElementById('btn-add-user-confirm')?.addEventListener('click', createUser);
  document.getElementById('btn-add-user-cancel')?.addEventListener('click', () => hideModal('modal-add-user'));
  document.getElementById('add-user-email')?.addEventListener('keydown', e => {
    if (e.key === 'Enter') createUser();
  });

  // Retire template modal
  document.getElementById('btn-retire-template-confirm')?.addEventListener('click', () => {
    hideModal('modal-retire-template');
    confirmRetireTemplate();
  });
  document.getElementById('btn-retire-template-cancel')?.addEventListener('click', () => {
    hideModal('modal-retire-template');
    STATE.pendingTemplateRetire = null;
  });

  // Activation modal
  document.getElementById('btn-activate-confirm')?.addEventListener('click', async () => {
    if (!STATE.pendingActivation) return;
    hideModal('modal-activate');
    await activateQuarter(STATE.pendingActivation);
    STATE.pendingActivation = null;
  });
  document.getElementById('btn-activate-cancel')?.addEventListener('click', () => {
    hideModal('modal-activate');
    STATE.pendingActivation = null;
  });

  // Waiting toggle
  document.getElementById('waiting-toggle-header')?.addEventListener('click', () => {
    const cards = document.getElementById('waiting-cards');
    const btn   = document.getElementById('waiting-toggle');
    if (!cards || !btn) return;
    cards.classList.toggle('hidden');
    btn.textContent = cards.classList.contains('hidden') ? '▼ Show' : '▲ Hide';
  });

  // RC resolved toggle
  document.getElementById('rc-resolved-header')?.addEventListener('click', () => {
    const list = document.getElementById('rc-resolved-list');
    const btn  = document.getElementById('rc-resolved-toggle');
    if (!list || !btn) return;
    list.classList.toggle('hidden');
    btn.textContent = list.classList.contains('hidden') ? '▼ Show' : '▲ Hide';
  });

  // All tasks filters
  document.querySelectorAll('[data-filter="status"]').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('[data-filter="status"]').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      STATE.filters.status = btn.dataset.value;
      saveFilters();
      renderAllTasks();
    });
  });

  // Sort column headers (delegated from thead)
  document.getElementById('all-tasks-thead')?.addEventListener('click', e => {
    const th = e.target.closest('th[data-sort]');
    if (!th) return;
    const col = th.dataset.sort;
    if (STATE.filters.sort === col) {
      // Same column — toggle direction
      STATE.filters.sortDir = STATE.filters.sortDir === 'asc' ? 'desc' : 'asc';
    } else {
      STATE.filters.sort    = col;
      // Overdue sort defaults to asc (worst first); others default to asc too
      STATE.filters.sortDir = 'asc';
    }
    saveFilters();
    renderAllTasks();
  });

  // Search
  document.getElementById('filter-search')?.addEventListener('input', (e) => {
    STATE.filters.search = e.target.value;
    renderAllTasks();
  });

  // Table/card view toggle
  document.getElementById('btn-table-view')?.addEventListener('click', () => {
    document.getElementById('btn-table-view').classList.add('active');
    document.getElementById('btn-card-view').classList.remove('active');
    document.getElementById('all-tasks-table-wrap')?.classList.remove('hidden');
    document.getElementById('all-tasks-cards-wrap')?.classList.add('hidden');
  });
  document.getElementById('btn-card-view')?.addEventListener('click', () => {
    document.getElementById('btn-card-view').classList.add('active');
    document.getElementById('btn-table-view').classList.remove('active');
    document.getElementById('all-tasks-table-wrap')?.classList.add('hidden');
    document.getElementById('all-tasks-cards-wrap')?.classList.remove('hidden');
    renderAllTasksCards();
  });

  // Export sign-off log
  document.getElementById('btn-export-log')?.addEventListener('click', exportSignOffLog);

  // Export matrix
  document.getElementById('btn-export-matrix-excel')?.addEventListener('click', exportMatrixExcel);

  // Dashboard overdue expand
  document.getElementById('overdue-expand-toggle')?.addEventListener('click', () => {
    const list = document.getElementById('overdue-detail-list');
    const btn  = document.getElementById('overdue-expand-toggle');
    if (!list || !btn) return;
    list.classList.toggle('hidden');
    btn.textContent = list.classList.contains('hidden') ? '▼ Show all' : '▲ Hide';
  });
}

// attachCardEvents uses event delegation on stable container elements rather than
// per-card listeners. Cards are rebuilt on every poll; delegation means we never
// need to re-attach listeners after a re-render.
function attachCardEvents() {
  // Delegate from each view container AND the task panel so we cover cards in
  // all views as well as sign-off / reverse / reassign buttons inside the panel.
  const containers = [
    document.getElementById('view-my-tasks'),
    document.getElementById('view-all-tasks'),
    document.getElementById('view-review-comments'),
    document.getElementById('task-panel'),
  ].filter(Boolean);

  containers.forEach(container => {
    if (container.dataset.delegationAttached) return;
    container.dataset.delegationAttached = 'true';

    container.addEventListener('click', (e) => {
      const el = e.target.closest('[data-action]');
      if (!el) return;
      e.stopPropagation();
      const { action, id, role } = el.dataset;

      if (action === 'open-task') {
        openTaskPanel(id);
      }

      if (action === 'signoff') {
        const assignment = STATE.assignments.find(a => a._id === id);
        if (!assignment) return;

        // If fired from inside the task panel the confirm-box is the confirmation —
        // execute directly. If fired from a task card open the modal first.
        const fromPanel = !!e.target.closest('#task-panel');
        if (fromPanel) {
          performSignOff(id, role);
        } else {
          STATE.pendingSignoff = { assignmentId: id, role };
          const titleEl = document.getElementById('modal-signoff-title');
          const bodyEl  = document.getElementById('modal-signoff-body');
          if (titleEl) titleEl.textContent = `Sign off as ${role}?`;
          if (bodyEl) bodyEl.innerHTML = `
            <p style="font-size:13px;margin-bottom:8px">${escapeHtml(assignment.Title || '')}</p>
            <p style="font-size:12px;color:var(--slate)">Recorded as ${renderBadge(STATE.currentUser?.Email)} · ${formatDateET(new Date().toISOString())}</p>`;
          showModal('modal-signoff');
        }
      }

      if (action === 'reverse') {
        STATE.pendingReversal = { assignmentId: id, role };
        const desc = document.getElementById('reversal-desc');
        if (desc) desc.textContent = `You are reversing the ${role} sign-off. This action will be logged.`;
        const reasonEl = document.getElementById('reversal-reason');
        if (reasonEl) reasonEl.value = '';
        document.getElementById('reversal-error')?.classList.add('hidden');
        showModal('modal-reversal');
      }

      if (action === 'rc-resolve') {
        resolveReviewComment(id);
      }

      if (action === 'rc-open-task') {
        openTaskPanel(id);
      }

      if (action === 'reassign') {
        openReassignModal(id, el.dataset.role);
      }

      if (action === 'signoff-behalf') {
        openSignOffBehalfModal(id, el.dataset.role);
      }
    });

    // Keyboard activation for non-button interactive elements (cards, RC task links)
    container.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' && e.key !== ' ') return;
      const el = e.target.closest('[data-action="open-task"], [data-action="rc-open-task"]');
      if (!el) return;
      e.preventDefault();
      e.stopPropagation();
      openTaskPanel(el.dataset.id);
    });
  });
}

// ============================================================
// REVIEW COMMENT SAVE
// ============================================================
async function saveReviewComment() {
  const taskId = document.getElementById('rc-task-select')?.value;
  const text   = document.getElementById('rc-comment-text')?.value?.trim();
  const priority = document.querySelector('input[name="rc-priority"]:checked')?.value || 'Normal';

  if (!text) { showToast('Please enter a comment', 'error'); return; }

  hideModal('modal-new-rc');

  try {
    const created = await createListItem(CONFIG.lists.reviewComments, {
      Title:               `RC: ${STATE.templates.find(t => t._id === taskId)?.TaskName || taskId}`,
      Quarter:             STATE.activeQuarter,
      TaskTemplateLookupId: taskId,
      CommentText:         text,
      CreatedBy:           STATE.currentUser.Email,
      CreatedDate:         new Date().toISOString(),
      Priority:            priority,
      Status:              'Open',
    });
    STATE.reviewComments.push({ ...created.fields, _id: created.id });
    await writeAuditLog('ReviewCommentCreated', {
      taskName:    `RC: ${STATE.templates.find(t => t._id === taskId)?.TaskName || taskId}`,
      newValue:    `Priority: ${priority} — ${text.substring(0, 100)}${text.length > 100 ? '…' : ''}`,
      assignmentId: taskId,
    });
    renderReviewComments();
    showToast('✓ Review comment posted', 'success');
  } catch (err) {
    showToast('Failed to post comment', 'error');
    logError('RC save failed:', err);
  }
}

async function resolveReviewComment(rcId) {
  const rc = STATE.reviewComments.find(r => r._id === rcId);
  if (!rc) return;

  // Store pending resolution and show the dedicated modal (modal-resolve-rc in index.html).
  STATE.pendingRCResolve = rcId;
  const noteEl = document.getElementById('resolve-rc-note');
  if (noteEl) noteEl.value = '';
  showModal('modal-resolve-rc');
}

async function confirmResolveReviewComment(rcId, note) {
  const rc = STATE.reviewComments.find(r => r._id === rcId);
  if (!rc) return;
  try {
    const now = new Date().toISOString();
    await updateListItem(CONFIG.lists.reviewComments, rcId, {
      Status:         'Resolved',
      ResolvedBy:     STATE.currentUser.Email,
      ResolvedDate:   now,
      ResolutionNote: note,
    });
    rc.Status = 'Resolved';
    rc.ResolvedBy = STATE.currentUser.Email;
    rc.ResolvedDate = now;
    rc.ResolutionNote = note;
    await writeAuditLog('ReviewCommentResolved', {
      taskName:    rc.Title || '',
      newValue:    note ? `Resolution: ${note.substring(0, 100)}${note.length > 100 ? '…' : ''}` : 'Resolved — no note',
      assignmentId: rcId,
    });
    renderReviewComments();
    showToast('✓ Comment resolved', 'success');
  } catch (err) {
    showToast('Failed to resolve', 'error');
    logError('RC resolve failed:', err);
  }
}

// ============================================================
// SUGGESTION SAVE / APPROVE / REJECT
// ============================================================
async function saveSuggestion() {
  const type   = document.querySelector('input[name="suggest-type"]:checked')?.value || 'Edit';
  const taskId = document.getElementById('suggest-task-select')?.value;
  const desc   = document.getElementById('suggest-desc')?.value?.trim();
  if (!desc) { showToast('Please describe the change', 'error'); return; }
  hideModal('modal-suggest');
  try {
    await createListItem(CONFIG.lists.taskSuggestions, {
      Title:               `${type}: ${STATE.templates.find(t => t._id === taskId)?.TaskName || 'New task'}`,
      SuggestionType:      type,
      SuggestedBy:         STATE.currentUser.Email,
      SuggestionDate:      new Date().toISOString(),
      TaskTemplateLookupId: taskId || null,
      ProposedChanges:     desc,
      Status:              'Pending',
    });
    showToast('✓ Suggestion submitted', 'success');
  } catch (err) {
    showToast('Failed to submit suggestion', 'error');
    logError('Suggestion save failed:', err);
  }
}

// Approves a suggestion and, for Edit/Retire types, automatically applies the
// change to the TaskTemplates list so admin does not need to edit it manually.
async function approveSuggestion(suggestionId) {
  const suggestion = STATE.suggestions.find(s => s._id === suggestionId);
  if (!suggestion) return;

  showLoading('Approving suggestion...');
  try {
    // Mark approved first
    await updateListItem(CONFIG.lists.taskSuggestions, suggestionId, {
      Status:     'Approved',
      ReviewedBy: STATE.currentUser.Email,
      ReviewDate: new Date().toISOString(),
    });
    suggestion.Status = 'Approved';

    // Auto-apply template mutation where possible
    if (suggestion.SuggestionType === 'Retire' && suggestion.TaskTemplateLookupId) {
      await applySuggestionToTemplate(suggestion);
      showToast('✓ Suggestion approved — template retired', 'success');
    } else if (suggestion.SuggestionType === 'Edit' && suggestion.TaskTemplateLookupId) {
      // Edit suggestions are free-form text — cannot auto-apply, flag for manual update.
      showToast('✓ Suggestion approved — update the template manually to reflect the change', 'success');
    } else {
      // Add suggestions — template must be created manually via Task Templates panel.
      showToast('✓ Suggestion approved — add the new task via Task Templates if needed', 'success');
    }

    await writeAuditLog('SuggestionApproved', { taskName: suggestion.Title });
    renderAdminPanel('suggestions');
  } catch (err) {
    showToast('Failed to approve suggestion', 'error');
    logError('Suggestion approval failed:', err);
  } finally {
    // Always clear the loading overlay regardless of success, failure, or early return.
    hideLoading();
  }
}

// Rejects a suggestion with an admin note.
async function rejectSuggestion(suggestionId, adminNote) {
  const suggestion = STATE.suggestions.find(s => s._id === suggestionId);
  if (!suggestion) return;

  try {
    await updateListItem(CONFIG.lists.taskSuggestions, suggestionId, {
      Status:     'Rejected',
      ReviewedBy: STATE.currentUser.Email,
      ReviewDate: new Date().toISOString(),
      AdminNote:  adminNote || '',
    });
    suggestion.Status = 'Rejected';
    suggestion.AdminNote = adminNote;
    await writeAuditLog('SuggestionRejected', { taskName: suggestion.Title });
    showToast('Suggestion rejected', '');
    renderAdminPanel('suggestions');
  } catch (err) {
    showToast('Failed to reject suggestion', 'error');
    logError('Suggestion rejection failed:', err);
  }
}

// Applies an approved suggestion directly to the TaskTemplates list.
// Currently handles Retire (sets IsActive = false).
// Edit suggestions are free-form and require manual template updates.
async function applySuggestionToTemplate(suggestion) {
  if (!suggestion.TaskTemplateLookupId) return;

  if (suggestion.SuggestionType === 'Retire') {
    await updateListItem(CONFIG.lists.taskTemplates, suggestion.TaskTemplateLookupId, {
      IsActive: false,
    });
    // Reflect change in cached templates so the UI updates without a reload
    const template = STATE.templates.find(t => t._id === suggestion.TaskTemplateLookupId);
    if (template) template.IsActive = false;
    log('Template retired:', suggestion.TaskTemplateLookupId);
  }
}

// ============================================================
// BULK CALENDAR SETUP
// ============================================================
// Creates all workday rows for a quarter from a single start date.
// Skips weekends automatically, marks any resulting weekend workdays with IsWeekend = true.
// Rows are created sequentially: WD1 = start date, WD2 = next business day, etc.
// If rows already exist for the quarter they are replaced (deleted then recreated).

async function setupCalendarBulk() {
  const quarterEl   = document.getElementById('cal-setup-quarter');
  const startEl     = document.getElementById('cal-setup-start');
  const maxWDEl     = document.getElementById('cal-setup-maxwd');
  const errEl       = document.getElementById('cal-setup-error');

  const quarter = quarterEl?.value?.trim();
  const startDate = startEl?.value;
  const maxWD = Number(maxWDEl?.value) || 20; // select always has a value; fallback is a safety net only

  if (!quarter || !/^Q[1-4]\s+\d{4}$/.test(quarter)) {
    if (errEl) { errEl.textContent = 'Enter a valid quarter — e.g. Q2 2026'; errEl.classList.remove('hidden'); }
    return;
  }
  if (!startDate) {
    if (errEl) { errEl.textContent = 'Select a start date for WD1'; errEl.classList.remove('hidden'); }
    return;
  }

  hideModal('modal-cal-setup');
  showLoading(`Setting up ${quarter} calendar...`);

  try {
    // Delete existing rows for this quarter first
    const existing = await getListItems(CONFIG.lists.closeCalendar, `fields/Quarter eq '${quarter}'`);
    for (const item of existing) {
      await graphRequest('DELETE',
        `/sites/${await getSiteId()}/lists/${CONFIG.lists.closeCalendar}/items/${item.id}`
      );
    }

    // Generate workday dates — each WD is one calendar day after the previous,
    // skipping nothing (admins sometimes have weekend workdays, so we don't auto-skip).
    // We mark any weekend dates with IsWeekend = true as a warning flag.
    let current = new Date(startDate + 'T12:00:00');
    const created = [];

    for (let wd = 1; wd <= maxWD; wd++) {
      // Use ET for date string and weekend detection — consistent with all other date handling.
      const currentET  = new Date(current.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
      const dateStr    = `${currentET.getFullYear()}-${String(currentET.getMonth()+1).padStart(2,'0')}-${String(currentET.getDate()).padStart(2,'0')}`;
      const dayOfWeek  = currentET.getDay(); // 0 = Sun, 6 = Sat
      const isWeekend  = dayOfWeek === 0 || dayOfWeek === 6;

      await createListItem(CONFIG.lists.closeCalendar, {
        Title:         `${quarter}-WD${wd}`,
        Quarter:       quarter,
        WorkdayNumber: wd,
        ActualDate:    dateStr,
        IsWeekend:     isWeekend,
        MilestoneType: 'Standard',
      });
      created.push({ WorkdayNumber: wd, ActualDate: dateStr, IsWeekend: isWeekend,
                     MilestoneLabel: null, MilestoneType: 'Standard', Quarter: quarter });

      // Advance by one calendar day for next workday
      current = new Date(current.getTime() + 86400000);

      // Update progress every 5 rows
      if (wd % 5 === 0) {
        const loadingText = document.getElementById('loading-text');
        if (loadingText) loadingText.textContent = `Creating calendar... WD${wd} of ${maxWD}`;
      }
    }

    // Update STATE.calendar if this is the active or viewing quarter
    if (quarter === STATE.activeQuarter || quarter === STATE.viewingQuarter) {
      STATE.calendar = created;
    }

    await writeAuditLog('CalendarEdit', {
      taskName: `${quarter} calendar setup`,
      newValue: `Created ${maxWD} workday rows starting ${startDate}`,
    });

    showToast(`✓ ${quarter} calendar created — ${maxWD} workdays from ${formatDateShort(startDate + 'T12:00:00')}`, 'success');
    renderAdminPanel('calendar');
  } catch (err) {
    showToast('Calendar setup failed — check SharePoint and try again', 'error');
    logError('setupCalendarBulk failed:', err);
  } finally {
    hideLoading();
  }
}

// ============================================================
// CALENDAR ROW EDIT
// ============================================================
function openEditCalendarRowModal(calRowId) {
  const row = STATE.calendar.find(c => c._id === calRowId);
  if (!row) return;
  STATE.pendingCalendarEdit = calRowId;

  const dateEl       = document.getElementById('edit-cal-date');
  const milestoneEl  = document.getElementById('edit-cal-milestone');
  const typeEl       = document.getElementById('edit-cal-milestone-type');
  const weekendEl    = document.getElementById('edit-cal-weekend');
  if (dateEl)      dateEl.value      = row.ActualDate || '';
  if (milestoneEl) milestoneEl.value = row.MilestoneLabel || '';
  if (typeEl)      typeEl.value      = row.MilestoneType || 'Standard';
  if (weekendEl)   weekendEl.checked = !!row.IsWeekend;

  const titleEl = document.getElementById('modal-edit-calendar-title');
  if (titleEl) titleEl.textContent = `Edit WD${row.WorkdayNumber}`;
  showModal('modal-edit-calendar');
}

async function saveCalendarRowEdit() {
  const calRowId = STATE.pendingCalendarEdit;
  if (!calRowId) return;
  const row = STATE.calendar.find(c => c._id === calRowId);
  if (!row) return;

  const newDate      = document.getElementById('edit-cal-date')?.value;
  const newMilestone = document.getElementById('edit-cal-milestone')?.value?.trim() || null;
  const newType      = document.getElementById('edit-cal-milestone-type')?.value || 'Standard';
  const newWeekend   = document.getElementById('edit-cal-weekend')?.checked || false;

  if (!newDate) { showToast('Date is required', 'error'); return; }

  const prevDate = row.ActualDate;
  const quarter  = row.Quarter || STATE.activeQuarter;

  // Snapshot current values for rollback on failure.
  const snapshot = {
    ActualDate:     row.ActualDate,
    MilestoneLabel: row.MilestoneLabel,
    MilestoneType:  row.MilestoneType,
    IsWeekend:      row.IsWeekend,
  };

  // Optimistic update — apply immediately so the admin panel reflects the change.
  row.ActualDate     = newDate;
  row.MilestoneLabel = newMilestone;
  row.MilestoneType  = newMilestone ? newType : null;
  row.IsWeekend      = newWeekend;

  try {
    await updateListItem(CONFIG.lists.closeCalendar, calRowId, {
      ActualDate:     newDate,
      MilestoneLabel: newMilestone,
      MilestoneType:  newMilestone ? newType : null,
      IsWeekend:      newWeekend,
    });
    await writeAuditLog('CalendarEdit', {
      taskName:      `WD${row.WorkdayNumber}`,
      previousValue: prevDate,
      newValue:      newDate,
    });
    hideModal('modal-edit-calendar');
    STATE.pendingCalendarEdit = null;

    // If the date changed, calculate the shift in days and offer cascade.
    if (prevDate && newDate !== prevDate) {
      // Use T12:00:00 to avoid DST boundary issues in date arithmetic.
      const shiftDays = Math.round(
        (new Date(newDate + 'T12:00:00') - new Date(prevDate + 'T12:00:00')) / (1000 * 60 * 60 * 24)
      );

      if (shiftDays !== 0) {
        // Count subsequent workdays that would be affected.
        const subsequent = STATE.calendar.filter(
          c => c.Quarter === quarter && Number(c.WorkdayNumber) > Number(row.WorkdayNumber)
        );

        if (subsequent.length > 0) {
          STATE.pendingCascade = {
            quarter,
            fromWD:    Number(row.WorkdayNumber),
            shiftDays,
            subsequent,
          };

          const descEl = document.getElementById('cascade-modal-desc');
          if (descEl) descEl.textContent =
            `WD${row.WorkdayNumber} moved ${Math.abs(shiftDays)} day${Math.abs(shiftDays) !== 1 ? 's' : ''} ` +
            `${shiftDays > 0 ? 'later' : 'earlier'}. ` +
            `Apply the same shift to all ${subsequent.length} subsequent workdays (WD${subsequent[0].WorkdayNumber}–WD${subsequent[subsequent.length-1].WorkdayNumber})?`;

          // Warn about any resulting weekends
          const warnEl = document.getElementById('cascade-warnings');
          const weekendWarnings = subsequent
            .map(c => {
              const shifted = new Date(new Date(c.ActualDate + 'T12:00:00').getTime() + shiftDays * 86400000);
              const shiftedLocalET = new Date(shifted.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
              const day = shiftedLocalET.getDay();
              return (day === 0 || day === 6)
                ? `WD${c.WorkdayNumber} would land on a ${day === 6 ? 'Saturday' : 'Sunday'}`
                : null;
            })
            .filter(Boolean);

          if (warnEl) {
            if (weekendWarnings.length) {
              warnEl.textContent = '⚠ ' + weekendWarnings.join(' · ');
              warnEl.classList.remove('hidden');
            } else {
              warnEl.classList.add('hidden');
            }
          }

          showModal('modal-cascade');
          return; // Cascade modal takes over from here.
        }
      }
    }

    showToast('✓ Calendar row updated', 'success');
    renderAdminPanel('calendar');
  } catch (err) {
    // Revert optimistic update so the calendar reflects actual SharePoint state.
    Object.assign(row, snapshot);
    showToast('Failed to update calendar row', 'error');
    logError('saveCalendarRowEdit failed:', err);
  }
}

// Applies the pending cascade shift to all subsequent workday rows.
async function confirmCascade() {
  const { quarter, fromWD, shiftDays, subsequent } = STATE.pendingCascade || {};
  if (!subsequent?.length) return;
  hideModal('modal-cascade');
  STATE.pendingCascade = null;

  showLoading(`Cascading ${Math.abs(shiftDays)}-day shift to ${subsequent.length} workdays...`);
  let updated = 0;
  try {
    for (const c of subsequent) {
      // Shift the date in ET to stay consistent with all other date handling.
      const oldDate  = new Date(c.ActualDate + 'T12:00:00');
      const shifted  = new Date(oldDate.getTime() + shiftDays * 86400000);
      const shiftedET = new Date(shifted.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
      const newDateStr = `${shiftedET.getFullYear()}-${String(shiftedET.getMonth()+1).padStart(2,'0')}-${String(shiftedET.getDate()).padStart(2,'0')}`;
      const isWeekend  = shiftedET.getDay() === 0 || shiftedET.getDay() === 6;

      await updateListItem(CONFIG.lists.closeCalendar, c._id, {
        ActualDate: newDateStr,
        IsWeekend:  isWeekend,
      });
      c.ActualDate = newDateStr;
      c.IsWeekend  = isWeekend;
      updated++;
    }
    await writeAuditLog('CalendarEdit', {
      taskName:  `Cascade from WD${fromWD}`,
      newValue:  `Shifted ${updated} workdays by ${shiftDays > 0 ? '+' : ''}${shiftDays} days`,
    });
    showToast(`✓ Cascaded to ${updated} workdays`, 'success');
  } catch (err) {
    showToast(`Cascade failed after ${updated} workdays — remaining rows unchanged`, 'error');
    logError('confirmCascade failed:', err);
  } finally {
    hideLoading();
    renderAdminPanel('calendar');
  }
}

// ============================================================
// SIGN OFF ON BEHALF
// ============================================================
// Used when a non-reviewer needs to sign the reviewer step during a tight close.
// The actual signer's email is recorded in the SignOffBy field — full audit trail.

function openSignOffBehalfModal(assignmentId, role) {
  const assignment = STATE.assignments.find(a => a._id === assignmentId);
  if (!assignment) return;

  STATE.pendingSignoff = { assignmentId, role };

  const titleEl = document.getElementById('modal-signoff-title');
  const bodyEl  = document.getElementById('modal-signoff-body');

  const assignedEmail = role === 'preparer' ? assignment.Preparer : assignment.Reviewer;
  const et = formatDateET(new Date().toISOString());

  if (titleEl) titleEl.textContent = `Sign off on behalf?`;
  if (bodyEl) bodyEl.innerHTML = `
    <p style="font-size:13px;margin-bottom:6px">${escapeHtml(assignment.Title || '')}</p>
    <p style="font-size:12px;color:var(--slate);margin-bottom:6px">
      Assigned ${role}: ${renderBadge(assignedEmail)}
    </p>
    <p style="font-size:12px;color:var(--slate);margin-bottom:6px">
      Signing as: ${renderBadge(STATE.currentUser?.Email)} · ${et}
    </p>
    <p style="font-size:11px;color:var(--amber);font-weight:500">
      ⚠ This will be recorded in the audit log as signed on behalf of the assigned ${role}.
    </p>`;

  showModal('modal-signoff');
}

// ============================================================
// REASSIGN TASK
// ============================================================
function openReassignModal(assignmentId, role) {
  const assignment = STATE.assignments.find(a => a._id === assignmentId);
  if (!assignment) return;

  STATE.pendingReassign = { assignmentId, role };

  const titleEl = document.getElementById('reassign-modal-title');
  const currentEl = document.getElementById('reassign-current');
  const selectEl = document.getElementById('reassign-user-select');

  if (titleEl) titleEl.textContent = `Reassign ${role === 'preparer' ? 'Preparer' : 'Reviewer'}`;

  const currentEmail = role === 'preparer' ? assignment.Preparer : assignment.Reviewer;
  if (currentEl) currentEl.innerHTML = `Current: ${renderBadge(currentEmail || '—')}`;

  if (selectEl) {
    selectEl.innerHTML = `<option value="">— No change —</option>` +
      STATE.users
        .filter(u => u.IsActive !== false)
        .sort((a, b) => (a.Title || '').localeCompare(b.Title || ''))
        .map(u => `<option value="${escapeHtml(u.Email)}" ${u.Email === currentEmail ? 'selected' : ''}>${escapeHtml((u.Emoji || '') + ' ' + (u.Title || u.Email.split('@')[0]))}</option>`)
        .join('');
  }

  showModal('modal-reassign');
}

async function confirmReassign() {
  const { assignmentId, role } = STATE.pendingReassign || {};
  if (!assignmentId) return;

  const assignment = STATE.assignments.find(a => a._id === assignmentId);
  if (!assignment) return;

  const selectEl = document.getElementById('reassign-user-select');
  const newEmail = selectEl?.value;
  if (!newEmail) return;

  const field = role === 'preparer' ? 'Preparer' : 'Reviewer';
  const prevEmail = assignment[field];
  if (newEmail === prevEmail) { hideModal('modal-reassign'); return; }

  hideModal('modal-reassign');

  // Optimistic update — apply immediately so badge updates without waiting for SharePoint.
  const snapshot = assignment[field];
  assignment[field] = newEmail;
  openTaskPanel(assignmentId);
  refreshCurrentView();

  try {
    await updateListItem(CONFIG.lists.quarterlyAssignments, assignmentId, { [field]: newEmail });
    await writeAuditLog('Reassignment', {
      taskName:      assignment.Title,
      assignmentId,
      previousValue: `${role}: ${prevEmail}`,
      newValue:      `${role}: ${newEmail}`,
    });
    showToast(`✓ ${role === 'preparer' ? 'Preparer' : 'Reviewer'} reassigned`, 'success');
  } catch (err) {
    // Revert optimistic update so STATE reflects actual SharePoint state.
    assignment[field] = snapshot;
    openTaskPanel(assignmentId);
    refreshCurrentView();
    showToast('Reassignment failed — please try again', 'error');
    logError('confirmReassign failed:', err);
  }
  STATE.pendingReassign = null;
}

// ============================================================
// USER ROLE EDIT
// ============================================================
function openEditUserRoleModal(email) {
  const user = STATE.users.find(u => u.Email === email);
  if (!user) return;
  STATE.pendingUserEdit = email;

  const nameEl = document.getElementById('edit-user-name');
  const roleEl = document.getElementById('edit-user-role');
  const activeEl = document.getElementById('edit-user-active');
  if (nameEl) nameEl.textContent = `${user.Emoji || ''} ${user.Title || email}`;
  if (roleEl) roleEl.value = user.Role || 'TeamMember';
  if (activeEl) activeEl.checked = user.IsActive !== false;
  showModal('modal-edit-user');
}

async function saveUserRoleEdit() {
  const email = STATE.pendingUserEdit;
  if (!email) return;
  const user = STATE.users.find(u => u.Email === email);
  if (!user) return;

  const newRole   = document.getElementById('edit-user-role')?.value || 'TeamMember';
  const newActive = document.getElementById('edit-user-active')?.checked !== false;

  try {
    await updateListItem(CONFIG.lists.users, user._id, {
      Role:     newRole,
      IsActive: newActive,
    });
    const prevRole = user.Role;
    user.Role     = newRole;
    user.IsActive = newActive;
    await writeAuditLog('UserEdit', {
      taskName:      email,
      previousValue: `Role: ${prevRole}`,
      newValue:      `Role: ${newRole}, IsActive: ${newActive}`,
    });
    // Update current user's role flags if they edited themselves
    if (email === STATE.currentUser?.Email) {
      STATE.isAdmin = newRole === 'Admin';
      STATE.isFinalReviewer = newRole === 'FinalReviewer' || STATE.isAdmin;
    }
    hideModal('modal-edit-user');
    STATE.pendingUserEdit = null;
    showToast('✓ User updated', 'success');
    renderAdminPanel('users');
  } catch (err) {
    showToast('Failed to update user', 'error');
    logError('saveUserRoleEdit failed:', err);
  }
}

// ============================================================
// ADD USER
// ============================================================
function updateAddUserPreview() {
  const nameEl  = document.getElementById('add-user-name');
  const emailEl = document.getElementById('add-user-email');
  const wrap    = document.getElementById('add-user-preview-wrap');
  const badge   = document.getElementById('add-user-preview-badge');
  if (!wrap || !badge) return;

  const emoji = STATE._addUserEmoji;
  const color = STATE._addUserColor;
  const name  = nameEl?.value?.trim() ||
    emailEl?.value?.trim()?.split('@')[0] || '?';

  if (!emoji && !color) { wrap.style.display = 'none'; return; }
  wrap.style.display = 'block';

  if (emoji && color) {
    badge.style.background = color + '22';
    badge.style.color = color;
    badge.textContent = `${emoji} ${name}`;
  } else if (emoji) {
    badge.style.background = 'var(--light-gray)';
    badge.style.color = 'var(--dark-slate)';
    badge.textContent = `${emoji} ${name}`;
  } else {
    badge.style.background = color + '22';
    badge.style.color = color;
    badge.textContent = name;
  }
}

async function createUser() {
  const emailEl = document.getElementById('add-user-email');
  const nameEl  = document.getElementById('add-user-name');
  const roleEl  = document.getElementById('add-user-role');
  const errEl   = document.getElementById('add-user-error');

  const email = emailEl?.value?.trim().toLowerCase();
  const role  = roleEl?.value || 'TeamMember';
  const name  = nameEl?.value?.trim() || email.split('@')[0];
  const emoji = STATE._addUserEmoji || null;
  const color = STATE._addUserColor || null;

  // Basic validation
  if (!email || !email.includes('@')) {
    if (errEl) { errEl.textContent = 'Please enter a valid email address.'; errEl.classList.remove('hidden'); }
    return;
  }
  if (STATE.users.find(u => u.Email.toLowerCase() === email)) {
    if (errEl) { errEl.textContent = 'A user with that email already exists.'; errEl.classList.remove('hidden'); }
    return;
  }

  hideModal('modal-add-user');
  showLoading('Adding user...');
  try {
    const created = await createListItem(CONFIG.lists.users, {
      Title:                    name,
      Email:                    email,
      Role:                     role,
      Emoji:                    emoji,
      Color:                    color,
      IsActive:                 true,
      NotifyOnAssignment:       false,
      NotifyOnReviewUnlock:     false,
      NotifyOnOverdue:          false,
      NotifyOnReassignment:     false,
      NotifyOnSuggestionUpdate: false,
    });
    STATE.users.push({ ...created.fields, _id: created.id });
    await writeAuditLog('UserEdit', { taskName: email, newValue: `Pre-added with role: ${role}` });
    const badgeNote = emoji && color ? ` with badge ${emoji}` : '';
    showToast(`✓ ${name} added${badgeNote}`, 'success');
    renderAdminPanel('users');
  } catch (err) {
    showToast('Failed to add user', 'error');
    logError('createUser failed:', err);
  } finally {
    hideLoading();
    STATE._addUserEmoji = null;
    STATE._addUserColor = null;
  }
}

// ============================================================
// SOX EXPORT
// ============================================================
function openSOXExportModal() {
  // Populate quarter picker with all available quarters from audit entries + assignments
  const quarters = [...new Set([
    ...STATE._auditEntries.map(e => e.Quarter).filter(Boolean),
    STATE.activeQuarter,
  ].filter(Boolean))].sort().reverse();

  const sel = document.getElementById('sox-export-quarter');
  if (sel) {
    sel.innerHTML = quarters.map(q =>
      `<option value="${escapeHtml(q)}" ${q === STATE.activeQuarter ? 'selected' : ''}>${escapeHtml(q)}</option>`
    ).join('');
  }
  showModal('modal-sox-export');
}

async function confirmSOXExport() {
  const quarter = document.getElementById('sox-export-quarter')?.value;
  if (!quarter) return;
  hideModal('modal-sox-export');
  showLoading(`Building audit log export for ${quarter}...`);

  try {
    // ── 1. Sign-off log ──────────────────────────────────────
    // Always fetch assignments fresh from SharePoint for the export quarter.
    // STATE.assignments only holds the active quarter — historical exports
    // would produce empty sign-off and unsigned tabs without this fetch.
    let assignments = STATE.assignments.filter(a => a.Quarter === quarter);
    if (!assignments.length || quarter !== STATE.activeQuarter) {
      const items = await getListItems(CONFIG.lists.quarterlyAssignments,
        `fields/Quarter eq '${quarter}' and fields/IsStaging eq false`);
      assignments = items.map(i => ({ ...i.fields, _id: i.id }));
    }

    function getSignOffWD(isoDate) {
      if (!isoDate) return '';
      const dateStr = isoDate.substring(0, 10);
      const match = STATE.calendar.find(c => c.Quarter === quarter && c.ActualDate === dateStr);
      return match ? match.WorkdayNumber : '';
    }

    const signOffRows = [
      ['Quarter','Task Name','Category','Sign-Off Type','Assigned To',
       'Signed Off By','On Behalf','Date & Time ET','Sign-Off WD',
       'Due WD','Timeliness','Reversed','Reversal Reason'],
    ];

    assignments.forEach(a => {
      if (a.PreparerSignOff) {
        const signWD = getSignOffWD(a.PreparerSignOffDate);
        const dueWD  = Number(a.PreparerWorkday);
        const onBehalf = a.PreparerSignOffBy && a.PreparerSignOffBy !== a.Preparer;
        signOffRows.push([
          quarter, a.Title, a.Category, 'Preparer',
          a.Preparer, a.PreparerSignOffBy || a.Preparer,
          onBehalf ? 'Yes' : 'No',
          formatDateET(a.PreparerSignOffDate),
          signWD, dueWD,
          typeof signWD === 'number' ? (signWD <= dueWD ? 'On Time' : 'Late') : 'Unknown',
          'No', '',
        ]);
      }
      if (a.ReviewerSignOff && a.SignOffMode !== 'Preparer Only') {
        const signWD = getSignOffWD(a.ReviewerSignOffDate);
        const dueWD  = Number(a.ReviewerWorkday);
        const onBehalf = a.ReviewerSignOffBy && a.ReviewerSignOffBy !== a.Reviewer;
        signOffRows.push([
          quarter, a.Title, a.Category, 'Reviewer',
          a.Reviewer, a.ReviewerSignOffBy || a.Reviewer,
          onBehalf ? 'Yes' : 'No',
          formatDateET(a.ReviewerSignOffDate),
          signWD, dueWD,
          typeof signWD === 'number' ? (signWD <= dueWD ? 'On Time' : 'Late') : 'Unknown',
          'No', '',
        ]);
      }
    });

    // ── 2. Unsigned tasks ────────────────────────────────────
    const unsignedRows = [['Quarter','Task Name','Category','Sign-Off Type','Assigned To','Due WD','Status']];
    assignments.forEach(a => {
      if (!a.PreparerSignOff) {
        unsignedRows.push([quarter, a.Title, a.Category, 'Preparer', a.Preparer || 'Unassigned', a.PreparerWorkday || '', a.Status || '']);
      }
      if (!a.ReviewerSignOff && a.SignOffMode !== 'Preparer Only') {
        unsignedRows.push([quarter, a.Title, a.Category, 'Reviewer', a.Reviewer || 'Unassigned', a.ReviewerWorkday || '', a.Status || '']);
      }
    });

    // ── 3. Reversals ─────────────────────────────────────────
    const auditQ = STATE._auditEntries.filter(e => e.Quarter === quarter);
    const reversalRows = [['Quarter','Date ET','WD','Task Name','Reversed By','Detail','Reason']];
    auditQ.filter(e => e.ActionType === 'Reversal').forEach(e => {
      reversalRows.push([quarter, formatDateET(e.ActionDate), e.WorkdayNumber || '', e.TaskName, e.ActionBy, e.NewValue || '', e.ReasonNote || '']);
    });

    // ── 4. Review comments ───────────────────────────────────
    // Fetch fresh for the export quarter — STATE.reviewComments only holds active quarter.
    let exportRCs = STATE.reviewComments.filter(rc => rc.Quarter === quarter);
    if (!exportRCs.length || quarter !== STATE.activeQuarter) {
      const rcItems = await getListItems(CONFIG.lists.reviewComments, `fields/Quarter eq '${quarter}'`);
      exportRCs = rcItems.map(i => ({ ...i.fields, _id: i.id }));
    }
    const rcRows = [['Quarter','Task Name','Posted By','Posted Date ET','Priority','Status','Resolved By','Resolved Date ET','Resolution Note']];
    exportRCs.forEach(rc => {
      rcRows.push([
        quarter, rc.Title, rc.CreatedBy,
        formatDateET(rc.CreatedDate), rc.Priority || 'Normal', rc.Status,
        rc.ResolvedBy || '', rc.ResolvedDate ? formatDateET(rc.ResolvedDate) : '',
        rc.ResolutionNote || '',
      ]);
    });

    // ── 5. Reassignments ─────────────────────────────────────
    const reassignRows = [['Quarter','Date ET','WD','Task Name','Changed By','Change Detail']];
    auditQ.filter(e => e.ActionType === 'Reassignment').forEach(e => {
      reassignRows.push([quarter, formatDateET(e.ActionDate), e.WorkdayNumber || '', e.TaskName, e.ActionBy, e.NewValue || '']);
    });

    // ── 6. Admin actions ─────────────────────────────────────
    const adminTypes = ['QuarterActivation','QuarterCreated','Rollforward','TaskEdit','UserEdit','CalendarEdit'];
    const adminRows = [['Quarter','Date ET','WD','Action Type','Subject','By','Detail']];
    auditQ.filter(e => adminTypes.includes(e.ActionType)).forEach(e => {
      adminRows.push([quarter, formatDateET(e.ActionDate), e.WorkdayNumber || '', e.ActionType, e.TaskName, e.ActionBy, e.NewValue || '']);
    });

    // ── 7. Summary ───────────────────────────────────────────
    const totalTasks     = assignments.length;
    const totalPrepDone  = assignments.filter(a => a.PreparerSignOff).length;
    const totalRevDone   = assignments.filter(a => a.ReviewerSignOff).length;
    const totalOnTime    = signOffRows.slice(1).filter(r => r[10] === 'On Time').length;
    const totalLate      = signOffRows.slice(1).filter(r => r[10] === 'Late').length;
    const totalReversals = reversalRows.length - 1;
    const totalRCs       = rcRows.length - 1;
    const totalRCOpen    = STATE.reviewComments.filter(rc => rc.Quarter === quarter && rc.Status === 'Open').length;

    const summaryRows = [
      ['Folio Audit Log Export', '', '', ''],
      ['Quarter', quarter, '', ''],
      ['Generated', formatDateET(new Date().toISOString()), '', ''],
      ['Generated By', STATE.currentUser?.Email || '', '', ''],
      ['', '', '', ''],
      ['SUMMARY', '', '', ''],
      ['Total assignments', totalTasks, '', ''],
      ['Preparer sign-offs complete', totalPrepDone, '', ''],
      ['Reviewer sign-offs complete', totalRevDone, '', ''],
      ['Sign-offs on time', totalOnTime, '', ''],
      ['Sign-offs late', totalLate, '', ''],
      ['Reversals', totalReversals, '', ''],
      ['Review comments posted', totalRCs, '', ''],
      ['Review comments open at export', totalRCOpen, '', ''],
      ['', '', '', ''],
      ['TABS IN THIS WORKBOOK', '', '', ''],
      ['1. Summary', 'This tab', '', ''],
      ['2. Sign-Offs', 'All completed preparer and reviewer sign-offs with timeliness', '', ''],
      ['3. Unsigned', 'Tasks not yet signed off at time of export', '', ''],
      ['4. Reversals', 'All sign-off reversals with reasons', '', ''],
      ['5. Review Comments', 'All review comments and their resolution status', '', ''],
      ['6. Reassignments', 'All mid-quarter reassignments', '', ''],
      ['7. Admin Actions', 'Quarter lifecycle and template changes', '', ''],
    ];

    // ── Build Excel workbook (SheetJS) ──────────────────────
    // Single .xlsx file with one named tab per section — proper auditor deliverable.
    const XLSX = window.XLSX;
    if (!XLSX) throw new Error('SheetJS not loaded — check network connection');

    const wb = XLSX.utils.book_new();

    function addSheet(name, rows, headerColor) {
      const ws = XLSX.utils.aoa_to_sheet(rows);

      // Column widths — set all to reasonable auto-width approximation
      const maxCols = Math.max(...rows.map(r => r.length));
      ws['!cols'] = Array.from({ length: maxCols }, (_, i) => ({
        wch: Math.min(50, Math.max(10,
          ...rows.map(r => String(r[i] ?? '').length)
        ))
      }));

      // Freeze header row
      ws['!freeze'] = { xSplit: 0, ySplit: 1 };

      XLSX.utils.book_append_sheet(wb, ws, name);
    }

    addSheet('Summary',          summaryRows);
    addSheet('Sign-Offs',        signOffRows);
    addSheet('Unsigned Tasks',   unsignedRows);
    addSheet('Reversals',        reversalRows);
    addSheet('Review Comments',  rcRows);
    addSheet('Reassignments',    reassignRows);
    addSheet('Admin Actions',    adminRows);

    // Write and download
    const wbBuf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob  = new Blob([wbBuf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url   = URL.createObjectURL(blob);
    const a     = document.createElement('a');
    a.href     = url;
    a.download = `Folio-AuditLog-${quarter}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    await writeAuditLog('SOXExport', {
      description: `SOX report exported for ${quarter} — ${signOffRows.length - 1} sign-offs, ${reversalRows.length - 1} reversals, ${rcRows.length - 1} RCs`,
      taskName: `SOX Export: ${quarter}`,
      newValue: `Exported by ${STATE.currentUser?.Email}`,
    });

    showToast(`✓ Audit log export ready — Folio-AuditLog-${quarter}.xlsx`, 'success');
  } catch (err) {
    showToast('SOX export failed', 'error');
    logError('confirmSOXExport failed:', err);
  } finally {
    hideLoading();
  }
}

// ============================================================
// AUDIT LOG EXPORT
// ============================================================
async function exportAuditLog() {
  showLoading('Loading audit log...');
  try {
    const items = await getListItems(CONFIG.lists.auditLog,
      STATE.activeQuarter ? `fields/Quarter eq '${STATE.activeQuarter}'` : ''
    );
    const rows = [['Quarter','Action Type','Action By','Date ET','Workday','Task Name','Assignment ID','Previous Value','New Value','Reason']];
    items.forEach(i => {
      const f = i.fields;
      rows.push([
        f.Quarter || '',
        f.ActionType || '',
        f.ActionBy || '',
        formatDateET(f.ActionDate),
        f.WorkdayNumber || '',
        f.TaskName || '',
        f.AssignmentID || '',
        f.PreviousValue || '',
        f.NewValue || '',
        f.ReasonNote || '',
      ]);
    });
    downloadCSV(rows, `Folio-AuditLog-${STATE.activeQuarter || 'all'}.csv`);
    showToast(`✓ Exported ${items.length} audit entries`, 'success');
  } catch (err) {
    showToast('Failed to export audit log', 'error');
    logError('exportAuditLog failed:', err);
  } finally {
    hideLoading();
  }
}

// ============================================================
// RC REPLY
// ============================================================
function openRCReplyInput(rcId) {
  // Find the rc-actions div for this comment and inject an inline reply form
  const btn = document.querySelector(`[data-action="rc-reply"][data-id="${rcId}"]`);
  if (!btn) return;
  const actionsDiv = btn.closest('.rc-actions');
  if (!actionsDiv) return;
  if (actionsDiv.querySelector('.rc-reply-form')) return; // already open

  actionsDiv.insertAdjacentHTML('beforeend', `
    <div class="rc-reply-form" style="display:flex;gap:6px;margin-top:8px;width:100%">
      <textarea id="reply-text-${rcId}" class="field-textarea" rows="2"
        placeholder="Type your reply..." style="flex:1;font-size:12px"></textarea>
      <div style="display:flex;flex-direction:column;gap:4px">
        <button class="btn-primary btn-sm" onclick="submitRCReply('${rcId}')">Post</button>
        <button class="btn-secondary btn-sm" onclick="this.closest('.rc-reply-form').remove()">Cancel</button>
      </div>
    </div>`);
  document.getElementById(`reply-text-${rcId}`)?.focus();
}

async function submitRCReply(rcId) {
  const text = document.getElementById(`reply-text-${rcId}`)?.value?.trim();
  if (!text) { showToast('Please enter a reply', 'error'); return; }

  const now = new Date().toISOString();
  try {
    const created = await createListItem(CONFIG.lists.reviewCommentReplies, {
      Title:                 `Reply to RC ${rcId}`,
      ReviewCommentLookupId: rcId,
      ReplyText:             text,
      CreatedByEmail:        STATE.currentUser.Email,
      CreatedDate:           now,
    });
    // Push into STATE immediately so the reply renders without waiting for next poll.
    // Set ReviewCommentLookupId explicitly since SharePoint may return a numeric lookup ID
    // in created.fields rather than the string rcId we need for client-side filtering.
    STATE.rcReplies.push({ ...created.fields, _id: created.id, ReviewCommentLookupId: rcId });
    showToast('✓ Reply posted', 'success');
    renderReviewComments();
  } catch (err) {
    showToast('Failed to post reply', 'error');
    logError('submitRCReply failed:', err);
  }
}

// ============================================================
// TEMPLATE MANAGEMENT
// ============================================================

// Filters the template table by the search term
function filterTemplateTable(search) {
  const term = search.toLowerCase();
  document.querySelectorAll('#admin-content .data-table tbody tr').forEach(row => {
    const text = row.textContent.toLowerCase();
    row.style.display = !term || text.includes(term) ? '' : 'none';
  });
}

async function retireTemplate(templateId) {
  const template = STATE.templates.find(t => t._id === templateId);
  if (!template) return;
  const name = template.TaskName || template.Title || templateId;
  // Use modal instead of window.confirm for consistency with the rest of the app.
  STATE.pendingTemplateRetire = templateId;
  const retireDetail = document.getElementById('retire-template-detail');
  if (retireDetail) retireDetail.textContent =
    `Retire "${name}"? This will set IsActive = No. The task will no longer appear in future rollforwards but existing assignments are unaffected.`;
  showModal('modal-retire-template');
}

async function confirmRetireTemplate() {
  const templateId = STATE.pendingTemplateRetire;
  if (!templateId) return;
  STATE.pendingTemplateRetire = null;
  const template = STATE.templates.find(t => t._id === templateId);
  if (!template) return;
  const name = template.TaskName || template.Title || templateId;
  try {
    await updateListItem(CONFIG.lists.taskTemplates, templateId, { IsActive: false });
    template.IsActive = false;
    await writeAuditLog('TaskEdit', { taskName: name, newValue: 'IsActive set to false (retired)' });
    showToast(`✓ "${name}" retired`, 'success');
    renderAdminPanel('templates');
  } catch (err) {
    showToast('Failed to retire template', 'error');
    logError('confirmRetireTemplate failed:', err);
  }
}

// Opens a simple inline editor for a template row's most common fields.
// Full edit capability; saves to SharePoint on confirm.
function openEditTemplateModal(templateId) {
  // templateId === null means create mode; a valid ID means edit mode.
  STATE.pendingTemplateEdit = templateId;

  const t = templateId ? STATE.templates.find(t => t._id === templateId) : null;

  const titleEl = document.querySelector('#modal-edit-template .modal-title');
  if (titleEl) titleEl.textContent = t ? 'Edit Template' : 'New Template';

  // Populate modal fields — empty defaults for create mode.
  const fields = {
    'edit-tpl-name':          t?.TaskName || t?.Title || '',
    'edit-tpl-category':      t?.Category || '',
    'edit-tpl-filingtype':    t?.FilingType || 'Both',
    'edit-tpl-signoffmode':   t?.SignOffMode || 'Sequential',
    'edit-tpl-prepwd':        t?.PreparerWorkday || '',
    'edit-tpl-revwd':         t?.ReviewerWorkday || '',
    'edit-tpl-prepwd-10k':    t?.PreparerWorkday10K || '',
    'edit-tpl-revwd-10k':     t?.ReviewerWorkday10K || '',
  };
  Object.entries(fields).forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (el) el.value = val;
  });
  showModal('modal-edit-template');
}

async function saveTemplateEdit() {
  const templateId = STATE.pendingTemplateEdit; // null = create mode, string = edit mode
  const t = templateId ? STATE.templates.find(t => t._id === templateId) : null;

  const name = document.getElementById('edit-tpl-name')?.value?.trim();
  if (!name) { showToast('Task name is required', 'error'); return; }

  const prepWD10K = document.getElementById('edit-tpl-prepwd-10k')?.value;
  const revWD10K  = document.getElementById('edit-tpl-revwd-10k')?.value;

  const updates = {
    Title:               name,
    Category:            document.getElementById('edit-tpl-category')?.value || 'Other',
    FilingType:          document.getElementById('edit-tpl-filingtype')?.value || 'Both',
    SignOffMode:         document.getElementById('edit-tpl-signoffmode')?.value || 'Sequential',
    PreparerWorkday:     Number(document.getElementById('edit-tpl-prepwd')?.value) || 1,
    ReviewerWorkday:     document.getElementById('edit-tpl-revwd')?.value
      ? Number(document.getElementById('edit-tpl-revwd').value) : null,
    PreparerWorkday10K:  prepWD10K ? Number(prepWD10K) : null,
    ReviewerWorkday10K:  revWD10K  ? Number(revWD10K)  : null,
    IsActive:            true,
  };

  try {
    if (t) {
      // Edit mode — update existing template
      await updateListItem(CONFIG.lists.taskTemplates, templateId, updates);
      Object.assign(t, updates);
      await writeAuditLog('TaskEdit', {
        taskName: updates.Title,
        newValue: `Category: ${updates.Category}, FilingType: ${updates.FilingType}, SignOffMode: ${updates.SignOffMode}, PrepWD: ${updates.PreparerWorkday}`,
      });
      showToast('✓ Template saved', 'success');
    } else {
      // Create mode — new template
      const created = await createListItem(CONFIG.lists.taskTemplates, {
        ...updates,
        TaskName: name, // TaskName mirrors Title for the app's display logic
      });
      STATE.templates.push({ ...created.fields, _id: created.id });
      await writeAuditLog('TaskEdit', { taskName: name, newValue: 'New template created' });
      showToast('✓ Template created', 'success');
    }
    hideModal('modal-edit-template');
    STATE.pendingTemplateEdit = null;
    renderAdminPanel('templates');
  } catch (err) {
    showToast('Failed to save template', 'error');
    logError('saveTemplateEdit failed:', err);
  }
}

// ============================================================
// ROLLFORWARD
// ============================================================

// Prompts for a new quarter name and sets it as the WorkingQuarter in AppSettings.
async function startNewQuarter() {
  // Show the new-quarter modal instead of window.prompt.
  const input = document.getElementById('new-quarter-name');
  const err   = document.getElementById('new-quarter-error');
  if (input) input.value = '';
  if (err)   err.classList.add('hidden');
  showModal('modal-new-quarter');
}

async function confirmNewQuarter() {
  const input = document.getElementById('new-quarter-name');
  const err   = document.getElementById('new-quarter-error');
  const quarter = (input?.value || '').trim();

  if (!/^Q[1-4]\s+\d{4}$/.test(quarter)) {
    if (err) { err.textContent = 'Use format Q1/Q2/Q3/Q4 YYYY — e.g. Q2 2026'; err.classList.remove('hidden'); }
    return;
  }

  hideModal('modal-new-quarter');
  showLoading(`Creating ${quarter}...`);
  try {
    await setAppSetting('WorkingQuarter', quarter);
    // Clear cached staging items so the grid reloads for the new quarter.
    STATE._stagingItems   = [];
    STATE._stagingLoading = false;
    STATE._auditEntries   = [];  // Force reload next time audit log opens
    STATE._auditFilter    = { type: 'All', person: '', quarter: '' };
    STATE.workingQuarter  = quarter;
    await writeAuditLog('QuarterCreated', { description: `Staging quarter set to ${quarter}` });
    showToast(`✓ ${quarter} created as staging quarter`, 'success');
    renderAdminPanel('rollforward');
  } catch (err) {
    showToast('Failed to create quarter', 'error');
    logError('startNewQuarter failed:', err);
  } finally {
    hideLoading();
  }
}

// Copies all active TaskTemplates into QuarterlyAssignments for the working quarter
// with IsStaging = true. All-or-nothing: if any item fails the batch is halted.
async function performRollforward() {
  const quarter = STATE.workingQuarter;
  const fromQuarter = STATE.activeQuarter;
  if (!quarter) { showToast('No staging quarter set', 'error'); return; }

  // Show a proper confirmation modal instead of window.confirm.
  STATE.pendingRollforward = quarter;
  const rfDetail = document.getElementById('rollforward-confirm-detail');
  if (rfDetail) rfDetail.textContent =
    `This will create ~${STATE.templates.length} staging assignments for ${quarter} copied from templates. ` +
    `Existing staging assignments for ${quarter} will be replaced. You can review before activating.`;
  showModal('modal-rollforward-confirm');
}

// Called by the rollforward confirmation modal confirm button.
async function confirmRollforward() {
  const quarter = STATE.pendingRollforward;
  if (!quarter) return;
  STATE.pendingRollforward = null;
  const fromQuarter = STATE.activeQuarter;

  showLoading(`Rolling forward to ${quarter}...`);
  let created = 0;
  try {
    // Remove any existing staging rows for this quarter first (clean slate)
    const existing = await getListItems(
      CONFIG.lists.quarterlyAssignments,
      `fields/Quarter eq '${quarter}' and fields/IsStaging eq true`
    );
    for (const item of existing) {
      await graphRequest('DELETE',
        `/sites/${await getSiteId()}/lists/${CONFIG.lists.quarterlyAssignments}/items/${item.id}`
      );
    }

    // Determine filing type for this quarter
    const filingType = isQuarterQ4(quarter) ? '10-K' : '10-Q';
    const eligible = STATE.templates.filter(t =>
      t.IsActive !== false &&
      (t.FilingType === filingType || t.FilingType === 'Both')
    );

    // Carry forward assignments from previous quarter if one exists
    const prevAssignments = fromQuarter
      ? await getListItems(CONFIG.lists.quarterlyAssignments, `fields/Quarter eq '${fromQuarter}'`)
      : [];
    const prevMap = {};
    prevAssignments.forEach(i => {
      if (i.fields.TaskTemplateLookupId) prevMap[i.fields.TaskTemplateLookupId] = i.fields;
    });

    for (const template of eligible) {
      const prev = prevMap[template._id];
      // Use 10-K workday numbers for Q4 quarters if they exist on the template,
      // otherwise fall back to the standard workday numbers.
      const isQ4 = filingType === '10-K';
      const prepWD = (isQ4 && template.PreparerWorkday10K)
        ? template.PreparerWorkday10K
        : template.PreparerWorkday || null;
      const revWD  = (isQ4 && template.ReviewerWorkday10K)
        ? template.ReviewerWorkday10K
        : template.ReviewerWorkday || null;

      await createListItem(CONFIG.lists.quarterlyAssignments, {
        Title:        `${quarter} - ${template.TaskName || template.Title || ''}`,
        Quarter:      quarter,
        TaskTemplateLookupId: template._id,
        Preparer:     prev?.Preparer || template.DefaultPreparer || null,
        Reviewer:     prev?.Reviewer || template.DefaultReviewer || null,
        SignOffMode:  template.SignOffMode || 'Sequential',
        Category:     template.Category || '',
        MatrixItem:   template.MatrixItem || null,
        MatrixCheckpoint: template.MatrixCheckpoint || null,
        PreparerWorkday:  prepWD,
        ReviewerWorkday:  revWD,
        HasDocumentLink:  template.HasDocumentLink || false,
        PreparerSignOff:  false,
        ReviewerSignOff:  false,
        Status:       'Not Started',
        IsStaging:    true,
      });
      created++;

      // Update progress every 10 items
      if (created % 10 === 0) {
        const pct = Math.round((created / eligible.length) * 100);
        const loadingText = document.getElementById('loading-text');
        if (loadingText) loadingText.textContent =
          `Rolling forward... ${created} of ${eligible.length} tasks (${pct}%)`;
      }
    }

    await writeAuditLog('Rollforward', {
      description: `Rolled forward ${created} assignments to ${quarter} from ${fromQuarter || 'templates'}`,
    });
    STATE._stagingItems = [];
    STATE._stagingLoading = false;
    showToast(`✓ Rolled forward ${created} tasks to ${quarter}`, 'success');
    renderAdminPanel('rollforward');
  } catch (err) {
    showToast(`Rollforward failed after ${created} tasks — check staging assignments in SharePoint`, 'error');
    logError('confirmRollforward failed:', err);
  } finally {
    hideLoading();
  }
}

// ============================================================
// QUARTER ACTIVATION
// ============================================================
async function activateQuarter(quarter) {
  showLoading(`Activating ${quarter}...`);
  // Declared outside try so catch can safely reference it for the error message.
  let stagingItems = [];
  try {
    stagingItems = await getListItems(
      CONFIG.lists.quarterlyAssignments,
      `fields/Quarter eq '${quarter}' and fields/IsStaging eq true`
    );
    for (const item of stagingItems) {
      await updateListItem(CONFIG.lists.quarterlyAssignments, item.id, { IsStaging: false });
    }
    await setAppSetting('ActiveQuarter', quarter);
    STATE.activeQuarter = quarter;
    STATE.workingQuarter = '';
    await setAppSetting('WorkingQuarter', '');
    await writeAuditLog('QuarterActivation', { description: `Activated ${quarter}` });
    // Reset filters when a new quarter goes live — stale filters from the previous
    // quarter would hide tasks and cause confusion in the new quarter.
    STATE.filters.status   = 'all';
    STATE.filters.category = 'all';
    STATE.filters.assignee = 'all';
    clearSavedFilters();
    await loadAllData();
    refreshCurrentView();
    showToast(`✓ ${quarter} is now live`, 'success');
  } catch (err) {
    const partialMsg = stagingItems.length
      ? 'partial update occurred — check QuarterlyAssignments list in SharePoint'
      : 'no changes were made';
    showToast(`Activation failed — ${partialMsg}`, 'error');
    logError('Activation failed:', err);
  }
  hideLoading();
}

// ============================================================
// EXPORTS
// ============================================================
function exportSignOffLog() {
  const quarter = getReadQuarter();
  const rows = [
    ['Quarter','Task Name','Category','Sign-Off Type','Signed Off By','Assigned To','Date & Time ET','Sign-Off Workday','Due Workday','On Time / Overdue','Reversal','Reversal Reason'],
  ];

  // Resolves which workday a given ISO date fell on by matching against the close calendar.
  function getSignOffWorkday(isoDate) {
    if (!isoDate) return '';
    const dateStr = isoDate.substring(0, 10); // YYYY-MM-DD
    const match = STATE.calendar.find(c => c.Quarter === quarter && c.ActualDate === dateStr);
    return match ? match.WorkdayNumber : '';
  }

  STATE.assignments.forEach(a => {
    if (a.PreparerSignOff) {
      const dueWD      = Number(a.PreparerWorkday);
      const signOffWD  = getSignOffWorkday(a.PreparerSignOffDate);
      const timeliness = typeof signOffWD === 'number' ? (signOffWD <= dueWD ? 'On Time' : 'Overdue') : 'Unknown';
      rows.push([
        quarter, a.Title, a.Category, 'Preparer',
        a.PreparerSignOffBy || a.Preparer, a.Preparer,
        formatDateET(a.PreparerSignOffDate), signOffWD, dueWD, timeliness, 'No', ''
      ]);
    }
    if (a.ReviewerSignOff) {
      const dueWD      = Number(a.ReviewerWorkday);
      const signOffWD  = getSignOffWorkday(a.ReviewerSignOffDate);
      const timeliness = typeof signOffWD === 'number' ? (signOffWD <= dueWD ? 'On Time' : 'Overdue') : 'Unknown';
      rows.push([
        quarter, a.Title, a.Category, 'Reviewer',
        a.ReviewerSignOffBy || a.Reviewer, a.Reviewer,
        formatDateET(a.ReviewerSignOffDate), signOffWD, dueWD, timeliness, 'No', ''
      ]);
    }
  });
  downloadCSV(rows, `Folio-SignOffLog-${quarter}.csv`);
}

function exportMatrixExcel() {
  const quarter = getReadQuarter();
  const rows = [['Item', 'Section', 'Preparer', 'Reviewer', ...CONFIG.matrixCheckpoints]];
  // Build matrix rows
  STATE.templates
    .filter(t => t.MatrixItem)
    .forEach(t => {
      const row = [t.MatrixItem, t.MatrixSection, '', ''];
      CONFIG.matrixCheckpoints.forEach(cp => {
        const isMatrixOnly = CONFIG.matrixOnlyColumns.includes(cp);
        if (isMatrixOnly) {
          const ms = STATE.matrixStatus.find(m => m.MatrixItem === t.MatrixItem);
          const fm = MATRIX_FIELD_MAP[cp];
          row.push(ms?.[fm.status] || 'Not Started');
        } else {
          const linked = STATE.assignments.find(a => a.MatrixItem === t.MatrixItem && a.MatrixCheckpoint === cp);
          if (!linked) row.push('N/A');
          else {
            const cpFields = getSignOffFields(getCheckpointRole(cp));
            row.push(linked[cpFields.signOff] ? 'Yes' : '');
          }
        }
      });
      rows.push(row);
    });
  downloadCSV(rows, `Folio-Matrix-${quarter}.csv`);
}



function downloadCSV(rows, filename) {
  const csv = rows.map(row => row.map(v => `"${String(v || '').replace(/"/g,'""')}"`).join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function renderAllTasksCards() {
  const wrap = document.getElementById('all-tasks-cards-wrap');
  if (!wrap) return;
  const filtered = getFilteredAssignments();
  wrap.innerHTML = filtered.map(a => renderTaskCard(a, STATE.currentUser?.Email, isTaskOverdue(a))).join('');
  attachCardEvents();
}

// ============================================================
// SETUP SCREEN
// ============================================================
function renderSetupScreen() {
  renderEmojiPicker('emoji-grid', CONFIG.emojiOptions[0], (emoji) => {
    if (!STATE.currentUser) STATE.currentUser = {};
    STATE.currentUser.Emoji = emoji;
    updateSetupPreview();
  });
  renderColorPicker('color-grid', CONFIG.colorOptions[0].hex, (color) => {
    if (!STATE.currentUser) STATE.currentUser = {};
    STATE.currentUser.Color = color;
    updateSetupPreview();
  });

  document.getElementById('setup-name')?.addEventListener('input', (e) => {
    if (!STATE.currentUser) STATE.currentUser = {};
    STATE.currentUser.Title = e.target.value;
    updateSetupPreview();
  });

  document.getElementById('setup-emoji-custom')?.addEventListener('input', (e) => {
    if (e.target.value.trim()) {
      if (!STATE.currentUser) STATE.currentUser = {};
      STATE.currentUser.Emoji = e.target.value.trim();
      updateSetupPreview();
    }
  });

  document.getElementById('btn-save-setup')?.addEventListener('click', completeSetup);
}

function updateSetupPreview() {
  const badge = document.getElementById('preview-badge');
  if (!badge || !STATE.currentUser) return;
  const hex = STATE.currentUser.Color || '#75787B';
  badge.style.background = hex + '22';
  badge.style.color = hex;
  badge.textContent = `${STATE.currentUser.Emoji || '?'} ${STATE.currentUser.Title || 'You'}`;
}

async function completeSetup() {
  const name = document.getElementById('setup-name')?.value?.trim();
  if (!name) { showToast('Please enter your name', 'error'); return; }
  STATE.currentUser.Title = name;

  try {
    await updateListItem(CONFIG.lists.users, STATE.currentUser._id, {
      Title: name,
      Emoji: STATE.currentUser.Emoji,
      Color: STATE.currentUser.Color,
    });
    showApp();
  } catch (err) {
    showToast('Failed to save profile — please try again', 'error');
    logError('Setup save failed:', err);
  }
}

// ============================================================
// SCREEN MANAGEMENT
// ============================================================
function showScreen(screenId) {
  document.querySelectorAll('.screen').forEach(s => s.classList.add('hidden'));
  document.getElementById(screenId)?.classList.remove('hidden');
}

async function showApp() {
  showScreen('screen-app');

  // Show correct nav items based on role
  document.querySelectorAll('.nav-matrix-link').forEach(el => {
    el.classList.toggle('hidden', !STATE.isFinalReviewer);
  });
  document.querySelectorAll('.nav-admin-link').forEach(el => {
    el.classList.toggle('hidden', !STATE.isAdmin);
  });
  // Hide "New Comment" button for non-reviewers — reviewers and admins only
  const newRCBtn = document.getElementById('btn-new-rc');
  if (newRCBtn) newRCBtn.classList.toggle('hidden', !STATE.isFinalReviewer && !STATE.isAdmin);

  updateNavAvatar();

  // Load all data
  showLoading('Loading your tasks...');
  try {
    await loadTemplates();
    if (STATE.activeQuarter) {
      await loadAllData();
    }
  } catch (err) {
    logError('Initial data load failed:', err);
    showStaleBanner(true);
  }
  hideLoading();

  updateWDIndicator();

  // Populate the quarter picker now that we know which quarters exist.
  populateQuarterPicker();

  // Restore persisted filters for this user+quarter, then sync all toolbar
  // controls to match (status buttons, selects, search input).
  restoreFilters();
  syncFilterUI();

  if (!STATE.activeQuarter) {
    // no-quarter is not a routed view — renderMyTasks handles the placeholder display.
    showView('my-tasks');
  } else {
    showView('my-tasks');
  }

  startPolling();
}

// escapeHtml moved to top-of-file utilities section

// ============================================================
// FILTER PERSISTENCE
// ============================================================
// Persists STATUS, CATEGORY, and ASSIGNEE filters per user per quarter in
// localStorage. Search and RC filters are intentionally not persisted —
// they are momentary query states, not recurring preferences.
// Key format: folio:filters:{email}:{quarter}

function filterStorageKey() {
  const email   = STATE.currentUser?.Email || 'unknown';
  const quarter = STATE.activeQuarter || 'none';
  return `folio:filters:${email}:${quarter}`;
}

function saveFilters() {
  if (!STATE.currentUser?.Email || !STATE.activeQuarter) return;
  try {
    const toSave = {
      status:   STATE.filters.status,
      category: STATE.filters.category,
      assignee: STATE.filters.assignee,
      sort:     STATE.filters.sort,
      sortDir:  STATE.filters.sortDir,
    };
    localStorage.setItem(filterStorageKey(), JSON.stringify(toSave));
  } catch (err) {
    // localStorage may be unavailable in some corporate environments — fail silently.
    logError('saveFilters failed:', err);
  }
}

function restoreFilters() {
  if (!STATE.currentUser?.Email || !STATE.activeQuarter) return;
  try {
    const raw = localStorage.getItem(filterStorageKey());
    if (!raw) return;
    const saved = JSON.parse(raw);
    if (saved.status)   STATE.filters.status   = saved.status;
    if (saved.category) STATE.filters.category = saved.category;
    if (saved.assignee) STATE.filters.assignee = saved.assignee;
    if (saved.sort)     STATE.filters.sort      = saved.sort;
    if (saved.sortDir)  STATE.filters.sortDir   = saved.sortDir;
    log('Filters restored for', STATE.activeQuarter, saved);
  } catch (err) {
    logError('restoreFilters failed:', err);
  }
}

// Clears persisted filters for the current user+quarter — called when quarter changes.
function clearSavedFilters() {
  try {
    localStorage.removeItem(filterStorageKey());
  } catch (err) { /* silent */ }
}

// ============================================================
// INITIALIZATION
// ============================================================
async function init() {
  log('Folio v' + CONFIG.version + ' initializing...');

  // Populate version spans from CONFIG so there is a single source of truth.
  document.querySelectorAll('[id^="app-version"]').forEach(el => { el.textContent = CONFIG.version; });

  // Validate config
  if (CONFIG.clientId === 'YOUR_APPLICATION_CLIENT_ID') {
    document.body.innerHTML = `
      <div style="padding:40px;font-family:Arial;max-width:600px;margin:0 auto">
        <h2 style="color:#C8102E">Configuration Required</h2>
        <p>Please fill in your CONFIG values in app.js before deploying:</p>
        <ul>
          <li>clientId — your Azure App Registration Client ID</li>
          <li>tenantId — your Azure Directory (Tenant) ID</li>
          <li>redirectUri — the full URL to this index.html on SharePoint</li>
          <li>siteUrl — your SharePoint site URL</li>
        </ul>
        <p>See Section 4 of the Build Guide for details.</p>
      </div>`;
    return;
  }

  // Initialize MSAL
  msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  // Handle redirect response
  const redirectResult = await msalInstance.handleRedirectPromise();
  if (redirectResult) {
    currentAccount = redirectResult.account;
    msalInstance.setActiveAccount(currentAccount);
  }

  // Check for existing session
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    currentAccount = accounts[0];
    msalInstance.setActiveAccount(currentAccount);
  }

  // Attach global events
  attachGlobalEvents();

  if (currentAccount) {
    // Already signed in
    showScreen('screen-app');
    showLoading('Loading Folio...');
    try {
      await loadActiveQuarter();
      const email = currentAccount.username;
      const isReturning = await loadCurrentUser(email);

      // Update last login
      if (STATE.currentUser?._id) {
        updateListItem(CONFIG.lists.users, STATE.currentUser._id, {
          LastLogin: new Date().toISOString()
        }).catch(() => {});
      }

      if (!isReturning || !STATE.currentUser.Emoji) {
        hideLoading();
        renderSetupScreen();
        showScreen('screen-profile-setup');
      } else {
        hideLoading();
        await showApp();
      }
    } catch (err) {
      hideLoading();
      logError('Init failed:', err);
      showScreen('screen-signin');
    }
  } else {
    showScreen('screen-signin');
    document.getElementById('btn-signin')?.addEventListener('click', () => {
      msalInstance.loginRedirect(loginRequest);
    });
  }
}

// ============================================================
// START
// ============================================================
document.addEventListener('DOMContentLoaded', init);
