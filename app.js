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
    calendarMilestones:   'CalendarMilestones',
    appSettings:          'AppSettings',
    users:                'Users',
    auditLog:             'AuditLog',
    taskSuggestions:      'TaskSuggestions',
    matrixStatus:         'MatrixStatus',
    reviewComments:       'ReviewComments',
    reviewCommentReplies: 'ReviewCommentReplies',
  },

  // App Settings
  // NOTE: When bumping version, also update:
  //   1. The ?v= cache-bust parameter on app.js and style.css in index.html
  //   2. The footer version display in index.html
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
    { hex: '#C0392B', label: 'Crimson' },
    { hex: '#FF7043', label: 'Tangerine' },
    { hex: '#E67300', label: 'Orange' },
    { hex: '#F5A623', label: 'Amber' },
    { hex: '#5A8A00', label: 'Lime' },
    { hex: '#3AB54A', label: 'Green' },
    { hex: '#558B2F', label: 'Olive' },
    { hex: '#00897B', label: 'Teal' },
    { hex: '#00838F', label: 'Petrol' },
    { hex: '#29ABE2', label: 'Sky' },
    { hex: '#2E4DA0', label: 'Navy' },
    { hex: '#5C6BC0', label: 'Indigo' },
    { hex: '#7B61FF', label: 'Purple' },
    { hex: '#8B5CF6', label: 'Lavender' },
    { hex: '#AB47BC', label: 'Violet' },
    { hex: '#E91E8C', label: 'Rose' },
    { hex: '#E86545', label: 'Coral' },
    { hex: '#75787B', label: 'Slate' },
    { hex: '#78909C', label: 'Steel' },
    { hex: '#B5651D', label: 'Chestnut' },
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
// DOMAIN CONSTANTS — single source of truth for magic strings
// ============================================================
// Use these instead of raw strings in logic/comparisons/writes.
// HTML option values and UI labels stay as literals since they
// are presentation layer, not domain logic.

const STATUS = {
  NOT_STARTED: 'Not Started',
  IN_PROGRESS: 'In Progress',
  PREPARED:    'Prepared',
  COMPLETE:    'Complete',
};

const SIGN_OFF_MODE = {
  SEQUENTIAL:    'Sequential',
  PREPARER_ONLY: 'Preparer Only',
};

const ROLE = {
  ADMIN:          'Admin',
  FINAL_REVIEWER: 'FinalReviewer',
  TEAM_MEMBER:    'TeamMember',
  READ_ONLY:      'ReadOnly',
};

// Domain vocabulary — single source of truth for all business string constants.
// Use these instead of inline string literals so spelling/casing changes propagate everywhere.
const PRIORITY = {
  URGENT: 'Urgent',
  NORMAL: 'Normal',
};

const RC_STATUS = {
  OPEN:     'Open',
  RESOLVED: 'Resolved',
};

const FILING = {
  Q:    '10-Q',
  K:    '10-K',
  BOTH: 'Both',
};

const CHECKPOINT = {
  SP_PREPARER:     'SP Preparer',
  SP_1ST_REVIEWER: 'SP 1st Reviewer',
  LOADED_TO_CLARA: 'Loaded to Clara',
  FINAL_REVIEW:    'Final Review',
};

const CATEGORY = {
  TIE_OUT:    'Tie Out',
  FINANCIALS: 'Financials',
  XBRL:       'XBRL',
};

// ============================================================
// STATE MUTATION HELPERS (Feedback B — single instrumentation point)
// Use these instead of direct STATE.assignments mutations so that
// future logging, validation, or optimistic-update guards can be
// added in one place without touching every call site.
// ============================================================
function patchAssignment(id, patch) {
  const a = STATE.assignments.find(x => x._id === id);
  if (a) Object.assign(a, patch);
  return a;
}

function patchMatrixStatus(id, patch) {
  const m = STATE.matrixStatus.find(x => x._id === id);
  if (m) Object.assign(m, patch);
  return m;
}

function patchCalendarRow(id, patch) {
  const c = STATE.calendar.find(x => x._id === id);
  if (c) Object.assign(c, patch);
  return c;
}

function patchUser(id, patch) {
  const u = STATE.users.find(x => x._id === id);
  if (u) Object.assign(u, patch);
  return u;
}


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
  milestones:     [],         // CalendarMilestones for active quarter
  matrixStatus:   [],         // MatrixStatus for active quarter
  reviewComments: [],         // ReviewComments for active quarter
  rcReplies:      [],         // ReviewCommentReplies for active quarter
  currentView:    'my-tasks',
  pollTimer:      null,
  siteId:         null,
  _siteIdPromise: null,   // In-flight getSiteId promise — prevents duplicate concurrent fetches       // SharePoint site ID (auto-populated)
  isAdmin:        false,
  isFinalReviewer: false,
  isReadOnly:      false,
  taskDetailId:   null,       // Currently open task panel assignment ID
  filters: {
    status:      'all',
    category:    'all',
    assignee:    'all',
    search:      '',
    rcStatus:    'all',
    rcPriority:  'all',
    rcQuarter:   'all',
    sort:        'overdue',
    sortDir:     'asc',
    showSkipped: false,
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
  _editUserEmoji:         null,   // emoji selected in edit user modal
  _pendingSkip:           null,   // {id, isSkipping, title} for skip/unskip confirmation
  _pendingBulkAssign:     null,   // {targets, preparer, reviewer, category} for bulk assign confirmation
  _editUserColor:         null,   // color selected in edit user modal
  suggestions:            [],         // TaskSuggestions (loaded when admin panel opens)
  pendingCascade:         null,   // {quarter, fromWD, shiftDays, subsequent}
  pendingRollforward:     null,   // quarter name awaiting rollforward confirm
  pendingWDEdit:          null,   // {assignmentId} awaiting workday edit confirm
  pendingDocLinkEdit:     null,   // assignmentId awaiting doc link edit
  _auditEntries:          [],     // Loaded on-demand when audit log panel opens
  _auditFilter:           { type: 'All', person: '', quarter: '' }, // Audit log filter state
  _matrixUpdateInFlight:  false,       // Guard against double-clicks on matrix cells
  _signOffInFlight:       new Set(),   // Guards against double-click sign-offs per assignment+role
  _allUsers:              [],          // All users including inactive — for badge rendering in audit/SOX
  _calendarLoading:        false,  // Guard against double calendar loads
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
  if (t === 'SVP')           return 'milestone-svp';
  if (t === 'MD')            return 'milestone-md';
  if (t === 'CFO')           return 'milestone-cfo';
  if (t === 'Team Deadline') return 'milestone-team';
  return 'milestone-std';  // Default for Standard or null/empty
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

// Centralizes the common pattern of coercing a date string to a local-noon Date.
// Using T12:00:00 avoids DST ambiguity — midnight local time can shift across
// day boundaries during DST transitions depending on the browser's timezone.
function dateFromString(dateStr) {
  if (!dateStr) return null;
  const clean = String(dateStr).split('T')[0]; // strip time if ISO timestamp
  return new Date(clean + 'T12:00:00');
}

function dateStringFromDate(d) {
  if (!d) return null;
  const et = new Date(d.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
  return `${et.getFullYear()}-${String(et.getMonth()+1).padStart(2,'0')}-${String(et.getDate()).padStart(2,'0')}`;
}

// Normalizes a quarter string to canonical form: "Q2 2026"
// Handles case, extra whitespace, and separator variations.
// Applied at every point of entry — UI inputs, imports, rollforward.
function normalizeQuarter(q) {
  if (!q) return '';
  return q.trim()
    .replace(/\s+/g, ' ')           // collapse multiple spaces
    .replace(/^q/i, 'Q')            // uppercase Q
    .replace(/^(Q\d)\s*[-_]?\s*(\d{4})$/, '$1 $2'); // ensure single space between Q# and year
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
  // Return 'between' for days that fall between two workdays (e.g. non-workday within the close period).
  // Callers should check typeof wd === 'number' before using as a workday number.
  return 'between';
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
  if (wd === 'between') {
    // Find the surrounding workdays for display
    const sorted2 = [...STATE.calendar].filter(c => c.Quarter === quarter)
      .sort((a,b) => Number(a.WorkdayNumber) - Number(b.WorkdayNumber));
    const prev = sorted2.filter(c => c.ActualDate < today).pop();
    const next = sorted2.find(c => c.ActualDate > today);
    return prev && next ? `Between WD${prev.WorkdayNumber} and WD${next.WorkdayNumber}` : quarter;
  }
  const date = resolveWorkday(quarter, wd);
  const dateStr = date ? formatDateShort(date) : '';
  const historyFlag = isViewingHistory() ? ' 🔒' : '';
  return `WD${wd}${dateStr ? ' · ' + dateStr : ''}${historyFlag}`;
}

function isTaskOverdue(assignment) {
  if (assignment.IsSkipped) return false;  // Skipped tasks are never overdue
  if (assignment.Status === STATUS.COMPLETE) return false;
  const wd = getTodaysWorkday(STATE.activeQuarter);
  // Post-close: all incomplete tasks are overdue
  if (wd === 'post-close') return true;
  if (!wd || typeof wd !== 'number') return false;
  const role = assignment.SignOffMode === SIGN_OFF_MODE.PREPARER_ONLY ? 'preparer' :
    !assignment.PreparerSignOff ? 'preparer' : 'reviewer';
  const dueWD = role === 'preparer'
    ? Number(assignment.PreparerWorkday)
    : Number(assignment.ReviewerWorkday);
  return wd > dueWD;
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
        // Allow filtering on non-indexed SharePoint columns.
        // Without this header SharePoint returns 400 on unindexed column filters.
        'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly',
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

// Converts a raw Graph API error into a user-friendly message.
// Called by catch blocks throughout the app to surface actionable guidance.
function classifyGraphError(err) {
  const msg = String(err?.message || '');
  if (msg.includes('403'))             return 'Permission denied — check Azure API permissions are granted';
  if (msg.includes('404'))             return 'Item not found — a list column may have been renamed';
  if (msg.includes('400'))             return 'Bad request — a list column name may not match the schema';
  if (msg.includes('503') ||
      msg.includes('502') ||
      msg.includes('504'))             return 'SharePoint is temporarily unavailable — please try again';
  if (msg.includes('AADSTS700016'))    return 'App not found — check clientId in CONFIG';
  if (msg.includes('AADSTS50011'))     return 'Redirect URI mismatch — check Azure app registration';
  if (msg.includes('AADSTS7000215'))   return 'Client secret expired — regenerate in Azure Portal';
  if (msg.includes('AADSTS'))          return 'Authentication error — try signing out and back in';
  if (msg.includes('getSiteId') ||
      msg.includes('siteId'))          return 'SharePoint site not found — check siteUrl in CONFIG';
  if (msg.includes('NetworkError') ||
      msg.includes('Failed to fetch')) return 'Network error — check your internet connection';
  return 'Unexpected error — check the browser console for details (F12)';
}

// ============================================================
// SHAREPOINT — SITE ID
// ============================================================
async function getSiteId() {
  if (STATE.siteId) return STATE.siteId;
  // Promise-cache: if a fetch is already in-flight, reuse it rather than firing a duplicate request.
  if (STATE._siteIdPromise) return STATE._siteIdPromise;
  const url = CONFIG.siteUrl.replace(/\/$/, '').replace('https://', '').replace('.sharepoint.com', '');
  const parts = url.split('/sites/');
  const hostname = parts[0] + '.sharepoint.com';
  const sitePath = 'sites/' + parts[1];
  STATE._siteIdPromise = graphRequest('GET', `/sites/${hostname}:/${sitePath}`)
    .then(data => {
      STATE.siteId = data.id;
      STATE._siteIdPromise = null;
      log('Site ID:', STATE.siteId);
      return STATE.siteId;
    })
    .catch(err => {
      STATE._siteIdPromise = null;
      throw err;
    });
  return STATE._siteIdPromise;
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
  // Load all AppSettings rows and filter client-side.
  // Avoids server-side filter on Title which requires a SharePoint index.
  // AppSettings is a tiny list (< 10 rows) so loading all is negligible.
  const items = await getListItems(CONFIG.lists.appSettings);
  const match = items.find(i => i.fields.Title === key);
  return match ? match.fields.SettingValue : null;
}

async function getAppSettings(...keys) {
  // Batch version — loads the list once and returns an object with all requested keys.
  // Use instead of multiple sequential getAppSetting() calls to halve round trips.
  const items = await getListItems(CONFIG.lists.appSettings);
  const result = {};
  keys.forEach(key => {
    const match = items.find(i => i.fields.Title === key);
    result[key] = match ? match.fields.SettingValue : null;
  });
  return result;
}

async function setAppSetting(key, value) {
  const items = await getListItems(CONFIG.lists.appSettings);
  const match = items.find(i => i.fields.Title === key);
  if (match) {
    await updateListItem(CONFIG.lists.appSettings, match.id, { SettingValue: value });
  } else {
    await createListItem(CONFIG.lists.appSettings, { Title: key, SettingValue: value });
  }
}

async function setAppSettings(kvPairs) {
  // Batch version — loads the list once and writes all key/value pairs.
  // Use instead of multiple sequential setAppSetting() calls.
  const items = await getListItems(CONFIG.lists.appSettings);
  await Promise.all(Object.entries(kvPairs).map(([key, value]) => {
    const match = items.find(i => i.fields.Title === key);
    return match
      ? updateListItem(CONFIG.lists.appSettings, match.id, { SettingValue: value })
      : createListItem(CONFIG.lists.appSettings, { Title: key, SettingValue: value });
  }));
}

// ============================================================
// AUDIT LOG
// ============================================================
async function writeAuditLog(actionType, details) {
  // Use explicitly provided quarter if given — falls back to activeQuarter.
  // Callers should always pass quarter explicitly to avoid recording against
  // the wrong quarter when viewingQuarter differs from activeQuarter.
  const quarter = details.quarter || STATE.activeQuarter || '';
  try {
    await createListItem(CONFIG.lists.auditLog, {
      Title: `${actionType}: ${details.taskName || details.description || ''}`,
      Quarter:       quarter,
      ActionType:    actionType,
      ActionBy:      STATE.currentUser?.Email || '',
      ActionDate:    new Date().toISOString(),
      WorkdayNumber: (() => { const w = getTodaysWorkday(quarter); return typeof w === 'number' ? w : 0; })(),
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
  const settings = await getAppSettings('ActiveQuarter', 'WorkingQuarter');
  STATE.activeQuarter  = settings['ActiveQuarter'];
  STATE.workingQuarter = settings['WorkingQuarter'];
  // viewingQuarter starts equal to activeQuarter on every login.
  // It diverges only when the user browses a historical quarter.
  STATE.viewingQuarter = STATE.activeQuarter;
  log('Active quarter:', STATE.activeQuarter);
}

async function loadCurrentUser(email) {
  // Load ALL users (active and inactive) so deactivated users still render
  // correctly in sign-off records, badges, and the SOX export.
  const allItems = await getListItems(CONFIG.lists.users);
  // Keep the full user list in a separate cache for badge rendering.
  // STATE.users (filtered to active) is used for assignment dropdowns.
  STATE._allUsers = allItems.map(i => ({ ...i.fields, _id: i.id }));

  const match = allItems.find(i => i.fields.Email?.toLowerCase() === email.toLowerCase());
  if (match) {
    STATE.currentUser = { ...match.fields, _id: match.id };
    // Block deactivated users
    if (STATE.currentUser.IsActive === false || STATE.currentUser.IsActive === 0) {
      throw new Error('ACCESS_DENIED: Your account has been deactivated. Contact your Folio admin.');
    }
  } else {
    // Unknown user — deny access. Admin must pre-add users.
    // This prevents any Moody's tenant user from auto-gaining access.
    throw new Error('ACCESS_DENIED: Your account is not registered in Folio. Contact your admin to be added.');
  }
  STATE.isAdmin        = STATE.currentUser.Role === ROLE.ADMIN;
  STATE.isFinalReviewer = STATE.currentUser.Role === ROLE.FINAL_REVIEWER || STATE.isAdmin;
  STATE.isReadOnly      = STATE.currentUser.Role === ROLE.READ_ONLY;
  return true;
}

async function loadUsers() {
  // If loadCurrentUser already fetched all users, reuse that data to avoid a duplicate request.
  const items = STATE._allUsers.length
    ? STATE._allUsers
    : (await getListItems(CONFIG.lists.users)).map(i => ({ ...i.fields, _id: i.id }));
  // STATE.users contains only active users — used for assignment dropdowns and staging grid.
  // STATE._allUsers contains everyone (active + inactive) — used for badge rendering.
  if (!STATE._allUsers.length) STATE._allUsers = items;
  STATE.users = items.filter(u => u.IsActive !== false && u.IsActive !== 0);
}

async function loadTemplates() {
  // Load all templates and filter client-side — avoids SharePoint boolean filter issues.
  const items = await getListItems(CONFIG.lists.taskTemplates);
  STATE.templates = items
    .map(i => ({ ...i.fields, _id: i.id }))
    .filter(t => t.IsActive !== false && t.IsActive !== 0);
  log('Templates loaded:', STATE.templates.length);
}

async function loadAssignments(quarter) {
  const items = await getListItems(
    CONFIG.lists.quarterlyAssignments,
    `fields/Quarter eq '${quarter}' and fields/IsStaging eq 0`
  );
  STATE.assignments = items.map(i => ({ ...i.fields, _id: i.id }));
  log('Assignments loaded:', STATE.assignments.length);
}

async function loadCalendar(quarter) {
  quarter = normalizeQuarter(quarter);  // Normalize before use as a filter key
  const items = await getListItems(CONFIG.lists.closeCalendar, `fields/Quarter eq '${quarter}'`);
  STATE.calendar = items.map(i => {
    const row = { ...i.fields, _id: i.id };
    // Normalize ActualDate to YYYY-MM-DD — SharePoint Date columns return full ISO timestamps
    if (row.ActualDate && row.ActualDate.includes('T')) {
      row.ActualDate = row.ActualDate.split('T')[0];
    }
    return row;
  }).sort((a,b) => Number(a.WorkdayNumber) - Number(b.WorkdayNumber));
}

async function loadMilestones(quarter) {
  quarter = normalizeQuarter(quarter);
  try {
    const items = await getListItems(CONFIG.lists.calendarMilestones, `fields/Quarter eq '${quarter}'`);
    STATE.milestones = items.map(i => ({ ...i.fields, _id: i.id }))
      .sort((a,b) => Number(a.WorkdayNumber) - Number(b.WorkdayNumber));
  } catch (err) {
    // CalendarMilestones list may not exist yet — fail silently
    STATE.milestones = [];
    log('CalendarMilestones not available:', err.message);
  }
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

async function loadRCReplies() {
  const rcIds = new Set(STATE.reviewComments.map(rc => rc._id));
  if (!rcIds.size) { STATE.rcReplies = []; return; }
  // Use Quarter server-side filter if the column exists (new replies have it).
  // Falls back to fetching all and filtering client-side for older replies without it.
  const quarter = getReadQuarter();
  const items = await getListItems(
    CONFIG.lists.reviewCommentReplies,
    quarter ? `fields/Quarter eq '${quarter}'` : ''
  ).catch(() => getListItems(CONFIG.lists.reviewCommentReplies));
  // Always filter by rcIds as a safety net — catches replies that predate the Quarter column.
  STATE.rcReplies = items
    .map(i => ({ ...i.fields, _id: i.id }))
    .filter(r => rcIds.has(r.ReviewCommentLookupId));
}

async function loadAllData() {
  if (!STATE.activeQuarter) return;
  // Delegate to loadViewingQuarterData so both functions stay in sync.
  // Adding a new data source only requires updating loadViewingQuarterData.
  await loadViewingQuarterData(STATE.activeQuarter);
}

// Returns true when the user is browsing a historical quarter.
function isViewingHistory() {
  return STATE.viewingQuarter && STATE.viewingQuarter !== STATE.activeQuarter;
}

// Permission helpers — single source of truth for role checks.
function canAdmin()  { return STATE.isAdmin; }
function canReview() { return STATE.isFinalReviewer || STATE.isAdmin; }
function canWrite()  { return !isViewingHistory(); }
// Returns true if the current user can act as reviewer on a specific assignment.
function canActAsReviewer(assignment) {
  if (!canWrite() || STATE.isReadOnly) return false;
  const email = STATE.currentUser?.Email;
  // Assigned reviewer can always act regardless of role
  if (assignment.Reviewer === email) return true;
  // FinalReviewer and Admin can act on any task
  return canReview();
}

// Returns true if the current user can post review comments.
// FinalReviewers and Admins can post on any task.
// TeamMembers can post only if assigned as reviewer on the specific task (when taskId provided)
// or on any task in the quarter (when taskId is omitted — used for the New Comment button).
// INTENTIONAL: When taskId is omitted, a TeamMember assigned as reviewer on one task
// can see the New Comment button and access all tasks' comments. This is acceptable
// because the modal asks them to select the task — they cannot comment on tasks
// they're not assigned to without an explicit selection.
function canPostReviewComment(taskId) {
  if (!canWrite() || STATE.isReadOnly) return false;
  if (canReview()) return true;
  // TeamMember assigned as reviewer on this specific task
  if (taskId) {
    return STATE.assignments.some(
      a => a.TaskTemplateLookupId === taskId && a.Reviewer === STATE.currentUser?.Email
    );
  }
  // No taskId — show New Comment button if user is reviewer on ANY task this quarter.
  // The modal itself scopes the comment to the selected task.
  return STATE.assignments.some(a => a.Reviewer === STATE.currentUser?.Email);
}

// Shows a confirmation dialog when an admin is about to write to a past quarter.
// Returns true if the write should proceed, false if the admin cancelled.
// Use this as a guard in any write that could target historical data.
function confirmIfPastQuarter(quarter, action = 'edit this item') {
  // No warning needed for the active quarter or the current staging/working quarter
  if (!quarter) return true;
  if (quarter === STATE.activeQuarter) return true;
  if (quarter === STATE.workingQuarter) return true;
  return window.confirm(
    `⚠️ You are about to ${action} in ${quarter}, which has already closed.\n\n`
    + `This change will be written to SharePoint.\n\n`
    + `Click OK to continue, or Cancel to go back.`
  );
}

// Use these helpers everywhere quarter context matters:
// getReadQuarter()  — the quarter currently being displayed (may be historical)
// getWriteQuarter() — the live quarter all writes must target (never historical)
function getReadQuarter()  { return STATE.viewingQuarter || STATE.activeQuarter; }
function getWriteQuarter() { return STATE.activeQuarter; }

// Loads all data for the viewing quarter (historical or live).
// Does NOT touch STATE.activeQuarter — write operations always use the live quarter.
async function loadViewingQuarterData(quarter) {
  if (!quarter) return;
  await Promise.all([
    loadAssignments(quarter),
    loadCalendar(quarter),
    loadMilestones(quarter),
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
    document.title = `Folio — ${quarter}${quarter !== STATE.activeQuarter ? ' (history)' : ''}`;
    await loadViewingQuarterData(quarter);
    updateHistoryBanner();
    updateWDIndicator(); updateContextRibbon();
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
// Derives quarter list from STATE.assignments (already loaded) plus the working quarter
// so no extra API call is needed.
async function populateQuarterPicker() {
  const sel = document.getElementById('quarter-picker');
  if (!sel) return;
  // Note: this function is sometimes called without await (fire-and-forget).
  // The try/catch inside ensures errors are always logged regardless.

  try {
    // Collect all known quarters from loaded assignments plus active/working quarters.
    // If no assignments loaded yet (first load), fall back to a lightweight API call.
    let quarters;
    if (STATE.assignments.length) {
      const fromAssignments = [...new Set(STATE.assignments.map(a => a.Quarter).filter(Boolean))];
      const extras = [STATE.activeQuarter, STATE.workingQuarter].filter(Boolean);
      quarters = [...new Set([...fromAssignments, ...extras])].sort().reverse();
    } else {
      const items = await getListItems(CONFIG.lists.quarterlyAssignments);
      quarters = [...new Set(items.map(i => i.fields?.Quarter).filter(Boolean))].sort().reverse();
    }

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
      // If the user is viewing a historical quarter, refresh that quarter's data.
      // Using loadAllData() here would replace STATE.assignments with live data
      // while STATE.viewingQuarter still points at history — causing wrong renders.
      if (isViewingHistory()) {
        await loadViewingQuarterData(STATE.viewingQuarter);
      } else {
        await loadAllData();
      }
      refreshCurrentView();
      updateWDIndicator(); updateContextRibbon();
      await populateQuarterPicker();
      showStaleBanner(false);
    } catch (err) {
      logError('Poll failed:', err, '|', classifyGraphError(err));
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
          .then(() => { refreshCurrentView(); updateWDIndicator(); updateContextRibbon(); showStaleBanner(false); })
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
  // Guard against double-clicks — ignore if this assignment+role is already in-flight.
  const guardKey = `${assignmentId}:${role}`;
  if (STATE._signOffInFlight.has(guardKey)) return;
  STATE._signOffInFlight.add(guardKey);

  const assignment = STATE.assignments.find(a => a._id === assignmentId);
  if (!assignment) { STATE._signOffInFlight.delete(guardKey); return; }

  const f = getSignOffFields(role);
  const now = new Date().toISOString();
  const userEmail = STATE.currentUser.Email;

  const fields = {
    [f.signOff]:    true,
    [f.signOffDate]:now,
    [f.signOffBy]:  userEmail,
    Status: role === 'preparer' && assignment.SignOffMode !== SIGN_OFF_MODE.PREPARER_ONLY ? STATUS.PREPARED : STATUS.COMPLETE,
  };

  // Snapshot the fields we are about to overwrite so we can restore them exactly on failure.
  const snapshot = {};
  Object.keys(fields).forEach(k => { snapshot[k] = assignment[k]; });

  // Optimistic update
  Object.assign(assignment, fields);
  refreshCurrentView();

  try {
    // Write to SharePoint first — if this succeeds the sign-off is real.
    await updateListItem(CONFIG.lists.quarterlyAssignments, assignmentId, fields);
    showToast('✓ Signed off', 'success');

    // Audit log write is best-effort — a failure here does NOT revert the sign-off.
    // The data is already committed to SharePoint; only the audit trail entry is missing.
    const assignedEmail = role === 'preparer' ? assignment.Preparer : assignment.Reviewer;
    const onBehalf = assignedEmail && assignedEmail !== userEmail;
    writeAuditLog('SignOff', {
      quarter:      assignment.Quarter || STATE.activeQuarter,
      taskName:     assignment.Title || assignment.TaskTemplateLookupId,
      assignmentId,
      newValue: onBehalf
        ? `${role} signed off by ${userEmail} ON BEHALF OF ${assignedEmail}`
        : `${role} signed off by ${userEmail}`,
    }).catch(auditErr => {
      logError('Audit log write failed for sign-off (sign-off itself succeeded):', auditErr);
    });
  } catch (err) {
    // Restore full snapshot — covers all fields set above, not just a subset.
    Object.assign(assignment, snapshot);
    refreshCurrentView();
    showToast(`Sign-off failed — ${classifyGraphError(err)}`, 'error');
    logError('Sign-off failed:', err);
  } finally {
    STATE._signOffInFlight.delete(guardKey);
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
    fields.Status = STATUS.NOT_STARTED;
  } else {
    prevValue = `Reviewer signed off by ${assignment.ReviewerSignOffBy} on ${assignment.ReviewerSignOffDate}`;
    fields.ReviewerSignOff     = false;
    fields.ReviewerSignOffDate = null;
    fields.ReviewerSignOffBy   = null;
    fields.Status = STATUS.PREPARED;
  }

  // Snapshot before optimistic update so we can restore on failure.
  const snapshot = {};
  Object.keys(fields).forEach(k => { snapshot[k] = assignment[k]; });

  Object.assign(assignment, fields);
  refreshCurrentView();

  try {
    await updateListItem(CONFIG.lists.quarterlyAssignments, assignmentId, fields);
    await writeAuditLog('Reversal', {
      quarter:       assignment.Quarter || STATE.activeQuarter,
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
    showToast(`Reversal failed — ${classifyGraphError(err)}`, 'error');
    logError('Reversal failed:', err);
  }
}

// ============================================================
// MATRIX STATUS UPDATE
// ============================================================
async function performMatrixUpdate(matrixItem, column, newStatus) {
  if (STATE._matrixUpdateInFlight) return;
  STATE._matrixUpdateInFlight = true;

  const existing = STATE.matrixStatus.find(
    m => m.MatrixItem === matrixItem && m.Quarter === STATE.activeQuarter
  );

  const now = new Date().toISOString();
  const userEmail = STATE.currentUser.Email;

  const fm = MATRIX_FIELD_MAP[column];
  if (!fm) { STATE._matrixUpdateInFlight = false; return; }

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
    showToast(`✓ ${column} updated to ${newStatus}`, 'success');
    // Log Final Review sign-offs to the audit trail — it's the last checkpoint before filing
    if (column === CHECKPOINT.FINAL_REVIEW) {
      await writeAuditLog('FinalReview', {
        quarter:       getWriteQuarter(),
        taskName:      `${matrixItem} — Final Review`,
        newValue:      newStatus,
        previousValue: existing?.[fm.status] || STATUS.NOT_STARTED,
      });
    }
  } catch (err) {
    // Revert optimistic update on failure.
    if (existing && Object.keys(snapshot).length) {
      Object.assign(existing, snapshot);
      renderMatrixView();
    }
    showToast('Update failed — please try again', 'error');
    logError('Matrix update failed:', err);
  } finally {
    STATE._matrixUpdateInFlight = false;
  }
}

// ============================================================
// USER HELPERS
// ============================================================
function getUserByEmail(email) {
  if (!email) return null;
  // Search _allUsers first (includes inactive) so deactivated users still render
  // correctly in sign-off records, audit logs, and SOX exports.
  const pool = STATE._allUsers.length ? STATE._allUsers : STATE.users;
  return pool.find(u => u.Email?.toLowerCase() === email.toLowerCase());
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

  // Toggle history-mode body class for visual read-only treatment
  document.body.classList.toggle('history-mode', isViewingHistory());

  // Update nav
  document.querySelectorAll('.nav-link').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.view === viewName);
  });

  // Hide all views — add hidden class so display:none !important applies
  document.querySelectorAll('.view').forEach(v => {
    v.classList.remove('active');
    v.classList.add('hidden');
    v.style.display = '';
  });

  // Show target view — remove hidden first so !important doesn't override display:block
  const target = document.getElementById(`view-${viewName}`);
  if (target) {
    target.classList.remove('hidden');
    target.classList.add('active');
  }

  // Close side panel
  closeTaskPanel();

  // Render the view
  renderCurrentView();
}

function refreshCurrentView() {
  renderCurrentView();
  updateWDIndicator(); updateContextRibbon();
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
// Updates the persistent context ribbon below the nav bar.
// Shows mode, quarter, and the user's effective role — reduces confusion
// about why actions are enabled/disabled in the current context.
function updateContextRibbon() {
  const ribbon = document.getElementById('context-ribbon');
  if (!ribbon) return;

  const quarter = getReadQuarter();
  if (!quarter) { ribbon.classList.add('hidden'); return; }

  const isHistory   = isViewingHistory();
  const isReadOnly  = STATE.isReadOnly;
  const isAdmin     = STATE.isAdmin;
  const isFinal     = STATE.isFinalReviewer;
  const activeQ     = STATE.activeQuarter;
  const workingQ    = STATE.workingQuarter;

  let modeIcon, modeLabel, modeClass, roleLabel;

  if (isHistory) {
    modeIcon  = '🔒';
    modeLabel = 'Historical · Read-only';
    modeClass = 'ribbon-history';
  } else if (!activeQ && workingQ) {
    modeIcon  = '⏳';
    modeLabel = 'Staging · Not yet activated';
    modeClass = 'ribbon-staging';
  } else {
    modeIcon  = '🟢';
    modeLabel = 'Live';
    modeClass = 'ribbon-live';
  }

  if (isAdmin)        roleLabel = 'Admin';
  else if (isReadOnly) roleLabel = 'Read Only';
  else if (isFinal)   roleLabel = 'Final Reviewer';
  else                roleLabel = 'Team Member';

  ribbon.className = `context-ribbon ${modeClass}`;
  ribbon.innerHTML = `
    <span class="ribbon-mode">${modeIcon} ${modeLabel}</span>
    <span class="ribbon-sep">·</span>
    <span class="ribbon-quarter">${escapeHtml(quarter)}</span>
    <span class="ribbon-sep">·</span>
    <span class="ribbon-role">You are: <strong>${roleLabel}</strong></span>
    ${isHistory ? '<span class="ribbon-sep">·</span><span class="ribbon-hint">Sign-offs and edits are disabled</span>' : ''}
  `;
}

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


function openEditDocLinkModal(assignmentId, currentUrl) {
  STATE.pendingDocLinkEdit = assignmentId;
  const urlEl = document.getElementById('doc-link-url');
  if (urlEl) urlEl.value = currentUrl || '';
  showModal('modal-edit-doc-link');
  urlEl?.select();
}

async function confirmEditDocLink() {
  const assignmentId = STATE.pendingDocLinkEdit;
  const url = document.getElementById('doc-link-url')?.value?.trim() || null;
  if (!assignmentId) return;
  hideModal('modal-edit-doc-link');
  try {
    await updateListItem(CONFIG.lists.quarterlyAssignments, assignmentId, {
      DocumentLink:    url,
      HasDocumentLink: !!url,
    });
    const assignment = STATE.assignments.find(a => a._id === assignmentId);
    if (assignment) {
      patchAssignment(assignmentId, { DocumentLink: url, HasDocumentLink: !!url });
    }
    showToast(url ? '✓ Document link updated' : '✓ Document link removed', 'success');
    openTaskPanel(STATE.taskDetailId);
  } catch (err) {
    showToast(`Failed to update link — ${classifyGraphError(err)}`, 'error');
    logError('confirmEditDocLink failed:', err);
  }
  STATE.pendingDocLinkEdit = null;
}


function openAddTaskModal() {
  // Admin can add a one-off task to the active quarter that isn't in templates.
  document.getElementById('add-task-name')?.focus && (document.getElementById('add-task-name').value = '');
  const prepSel = document.getElementById('add-task-preparer');
  const revSel  = document.getElementById('add-task-reviewer');
  const catSel  = document.getElementById('add-task-category');
  if (prepSel) prepSel.innerHTML = '<option value="">— Unassigned —</option>' +
    STATE.users.filter(u => u.IsActive !== false).map(u =>
      `<option value="${escapeHtml(u.Email)}">${escapeHtml(u.Title)}</option>`).join('');
  if (revSel) revSel.innerHTML = '<option value="">— None —</option>' +
    STATE.users.filter(u => u.IsActive !== false).map(u =>
      `<option value="${escapeHtml(u.Email)}">${escapeHtml(u.Title)}</option>`).join('');
  const cats = [...new Set(STATE.templates.map(t => t.Category).filter(Boolean))].sort();
  if (catSel) catSel.innerHTML = '<option value="">— Select category —</option>' +
    cats.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
  showModal('modal-add-task');
}

async function confirmAddTask() {
  const name    = document.getElementById('add-task-name')?.value?.trim();
  const cat     = document.getElementById('add-task-category')?.value;
  const prep    = document.getElementById('add-task-preparer')?.value;
  const rev     = document.getElementById('add-task-reviewer')?.value;
  const prepWD  = Number(document.getElementById('add-task-prepwd')?.value) || null;
  const revWD   = Number(document.getElementById('add-task-revwd')?.value) || null;
  const mode    = document.getElementById('add-task-mode')?.value || SIGN_OFF_MODE.SEQUENTIAL;

  if (!name) { showToast('Task name is required', 'error'); return; }
  if (!cat)  { showToast('Category is required', 'error'); return; }
  if (mode === SIGN_OFF_MODE.SEQUENTIAL && !revWD) {
    showToast('Reviewer Workday is required for Sequential tasks', 'error'); return;
  }

  hideModal('modal-add-task');
  showLoading('Adding task...');
  try {
    const quarter = STATE.activeQuarter;
    const title   = `${quarter} - ${cat} — ${name}`;
    const created = await createListItem(CONFIG.lists.quarterlyAssignments, {
      Title:              title,
      Quarter:            quarter,
      Category:           cat,
      Preparer:           prep || null,
      Reviewer:           rev || null,
      SignOffMode:        mode,
      PreparerWorkday:    prepWD,
      ReviewerWorkday:    revWD || null,
      PreparerSignOff:    false,
      ReviewerSignOff:    false,
      Status:             STATUS.NOT_STARTED,
      IsStaging:          false,
      IsSkipped:          false,
      TaskTemplateLookupId: 'adhoc',
    });
    STATE.assignments.push({ ...created.fields, _id: created.id });
    showToast(`✓ Task added: ${name}`, 'success');
    refreshCurrentView();
  } catch (err) {
    showToast(`Failed to add task — ${classifyGraphError(err)}`, 'error');
    logError('confirmAddTask failed:', err);
  } finally {
    hideLoading();
  }
}


async function signOffAll() {
  // Signs off all tasks where the current user is the preparer/reviewer AND the step is ready.
  const email = STATE.currentUser?.Email;
  const quarter = STATE.activeQuarter;
  if (!email || !quarter || isViewingHistory()) return;

  const ready = STATE.assignments.filter(a => {
    if (a.IsSkipped || a.Status === STATUS.COMPLETE) return false;
    const isPreparer = a.Preparer === email;
    const isReviewer = a.Reviewer === email;
    if (isPreparer && !a.PreparerSignOff) return true;
    if (isReviewer && a.PreparerSignOff && !a.ReviewerSignOff &&
        a.SignOffMode !== SIGN_OFF_MODE.PREPARER_ONLY) return true;
    return false;
  });

  if (!ready.length) { showToast('No tasks ready to sign off', 'info'); return; }

  if (!window.confirm(`Sign off all ${ready.length} ready task(s)? This cannot be undone in bulk.`)) return;

  showLoading(`Signing off ${ready.length} tasks...`);
  let done = 0;
  const errors = [];

  for (const assignment of ready) {
    try {
      const isPreparer = assignment.Preparer === email && !assignment.PreparerSignOff;
      const role = isPreparer ? 'preparer' : 'reviewer';
      await performSignOff(assignment._id, role);
      done++;
    } catch (err) {
      errors.push(assignment.Title || assignment._id);
      logError('Sign-off-all failed for:', assignment.Title, err);
      // Stop on first failure — a network error mid-batch means subsequent
      // sign-offs would also fail, and partial state is clearer to recover from.
      break;
    }
  }

  hideLoading();
  if (errors.length) {
    const failedNames = errors.slice(0, 3).join(', ') + (errors.length > 3 ? ` + ${errors.length - 3} more` : '');
    showToast(`Signed off ${done} tasks. Failed: ${failedNames}`, 'warning');
  } else {
    showToast(`✓ Signed off ${done} tasks`, 'success');
  }
  refreshCurrentView();
}

// ============================================================
// MY TASKS VIEW
// ============================================================
function filterMyAssignments() {
  const email = STATE.currentUser?.Email;
  if (!email) return [];
  return STATE.assignments.filter(a =>
    !a.IsSkipped && (a.Preparer === email || a.Reviewer === email)
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

  // Sign off all button — show when user has more than 1 ready task
  const signOffAllBtn = document.getElementById('btn-signoff-all');
  if (signOffAllBtn && email && quarter && !isViewingHistory()) {
    const readyCount = STATE.assignments.filter(a => {
      if (a.IsSkipped || a.Status === STATUS.COMPLETE) return false;
      if (a.Preparer === email && !a.PreparerSignOff) return true;
      if (a.Reviewer === email && a.PreparerSignOff && !a.ReviewerSignOff &&
          a.SignOffMode !== SIGN_OFF_MODE.PREPARER_ONLY) return true;
      return false;
    }).length;
    signOffAllBtn.style.display = readyCount > 1 ? '' : 'none';
    signOffAllBtn.textContent = `Sign off all (${readyCount})`;
  } else if (signOffAllBtn) {
    signOffAllBtn.style.display = 'none';
  }

  if (!quarter) {
    // Show the no-quarter placeholder directly without going through the view router.
    // We do NOT set STATE.currentView here — no-quarter is not a routed view and the
    // router has no case for it. Keeping STATE.currentView = 'my-tasks' means the next
    // refreshCurrentView() call will re-run renderMyTasks(), which will re-check the
    // quarter and show this placeholder again if still needed.
    // Special case: no-quarter is not a routed view so we manage classes directly.
    document.querySelectorAll('.view').forEach(v => {
      v.classList.remove('active');
      v.classList.add('hidden');
    });
    const noQ = document.getElementById('view-no-quarter');
    if (noQ) { noQ.classList.remove('hidden'); noQ.classList.add('active'); }
    return;
  }

  const tasks = filterMyAssignments();
  const wd          = getTodaysWorkday(quarter);
  const tomorrowWD  = getTomorrowWorkday(quarter);
  const todayWD     = typeof wd === 'number' ? wd : -1;

  const overdue    = tasks.filter(t => isTaskOverdue(t) && t.Status !== STATUS.COMPLETE);
  const waiting    = tasks.filter(t => t.Status !== STATUS.COMPLETE && isLocked(t, email));
  const active     = tasks.filter(t => !isTaskOverdue(t) && t.Status !== STATUS.COMPLETE && !isLocked(t, email));
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
  if (assignment.SignOffMode === SIGN_OFF_MODE.PREPARER_ONLY) return false;
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
  // Fallback: fires for completed tasks or when viewing someone else's task in All Tasks.
  // PreparerWorkday is the primary due date shown in the table for management view.
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
  const canSignPreparer  = !assignment.PreparerSignOff && !STATE.isReadOnly;
  // Any user assigned as reviewer can sign off — role doesn't restrict the assigned reviewer
  const canSignReviewer  = !assignment.ReviewerSignOff && (isReviewer || isAdmin || isFinalReviewer) && !STATE.isReadOnly;
  const role = canSignPreparer ? 'preparer'
    : canSignReviewer ? 'reviewer' : null;
  const locked = isLocked(assignment, currentEmail);
  const dueWD = getDueWD(assignment, currentEmail);
  const dueDate = resolveWorkday(getReadQuarter(), dueWD);

  // Check for urgent review comments
  const hasUrgentRC = STATE.reviewComments.some(
    rc => rc.TaskTemplateLookupId === assignment.TaskTemplateLookupId &&
          rc.Priority === PRIORITY.URGENT && rc.Status === RC_STATUS.OPEN
  );
  const rcCount = STATE.reviewComments.filter(
    rc => rc.TaskTemplateLookupId === assignment.TaskTemplateLookupId
  ).length;

  const prepBadge = renderBadge(assignment.Preparer);
  const revBadge = assignment.Reviewer ? renderBadge(assignment.Reviewer) : '';

  let signoffBtn = '';
  let nudgeBtn = '';
  if (isViewingHistory()) {
    signoffBtn = '';
  } else if (isWaiting || locked) {
    signoffBtn = `<button class="btn-secondary btn-sm" disabled>🔒 Awaiting preparer sign-off</button>`;
    // Nudge button — only for the assigned reviewer on waiting tasks, not ReadOnly
    const isAssignedReviewer = assignment.Reviewer === currentEmail;
    if (isAssignedReviewer && !STATE.isReadOnly) {
      // Rate-limit: don't show if nudged in the last hour
      const lastNudged = assignment.NudgeSent ? new Date(assignment.NudgeSent) : null;
      const hourAgo = new Date(Date.now() - 60 * 60 * 1000);
      const recentlyNudged = lastNudged && lastNudged > hourAgo;
      nudgeBtn = recentlyNudged
        ? `<button class="btn-icon btn-sm" disabled title="Nudge sent recently — wait an hour before sending another">👋 Nudged</button>`
        : `<button class="btn-icon btn-sm" data-action="nudge-preparer" data-id="${assignment._id}" title="Send a reminder to the preparer">👋 Nudge</button>`;
    }
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
        ${nudgeBtn}
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
    (chips.length > 1 ? '<button class="filter-chip-clear-all" id="btn-clear-all-filters" data-action="clear-all-filters">Clear all</button>' : '');

  // Wire chip remove buttons
  bar.querySelectorAll('.filter-chip-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      chips[Number(btn.dataset.chip)].clear();
      syncFilterUI();
      renderAllTasks();
    });
  });

  // Wire clear all
  // btn-clear-all-filters handled by delegation
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
  const activeCount = STATE.assignments.filter(a => !a.IsSkipped).length;
  if (sub) sub.textContent = `${quarter || '—'} · ${activeCount} tasks · ${getCompletionPct()}% complete`;

  populateCategoryFilter();
  populateAssigneeFilter();
  renderSortHeaders();
  renderActiveFilterChips();

  const filtered = getFilteredAssignments();
  const tbody = document.getElementById('all-tasks-tbody');
  if (!tbody) return;

  tbody.innerHTML = '';
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
      <td style="font-weight:500;font-size:12px">${escapeHtml(a.Title || '')}${a.IsSkipped ? ' <span style="font-size:9px;background:#F5F5F5;color:var(--slate);padding:1px 5px;border-radius:4px">SKIPPED</span>' : ''}</td>
      <td><span class="cat-tag">${escapeHtml(a.Category || '')}</span></td>
      <td>${renderBadge(a.Preparer)}</td>
      <td>${a.Reviewer ? renderBadge(a.Reviewer) : '<span style="font-size:10px;color:var(--slate)">Preparer only</span>'}</td>
      <td style="font-size:11px;color:var(--slate)">${a.PreparerWorkday ? 'WD' + a.PreparerWorkday : '—'}</td>
      <td style="font-size:11px;color:var(--slate)">${a.ReviewerWorkday ? 'WD' + a.ReviewerWorkday : '—'}</td>
      <td>${renderStatusBadge(a)}</td>
      <td style="font-size:10px;color:var(--slate)">${getTaskRCCount(a) || '—'}</td>`;
  });

  // Skipped tasks toggle for admins
  const skippedCount = STATE.assignments.filter(a => a.IsSkipped).length;
  const skippedToggleEl = document.getElementById('skipped-tasks-toggle');
  if (skippedToggleEl) {
    if (STATE.isAdmin && skippedCount > 0) {
      skippedToggleEl.style.display = '';
      skippedToggleEl.innerHTML = `<button class="btn-secondary btn-sm" onclick="STATE.filters.showSkipped=!STATE.filters.showSkipped;renderAllTasks()">
        ${STATE.filters.showSkipped ? 'Hide skipped tasks' : `Show ${skippedCount} skipped task${skippedCount !== 1 ? 's' : ''}`}
      </button>`;
    } else {
      skippedToggleEl.style.display = 'none';
    }
  }
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
  const s = assignment.Status || STATUS.NOT_STARTED;

  // Overdue takes priority over all other states.
  if (isTaskOverdue(assignment)) {
    return `<span class="status-badge status-overdue">⚠ Overdue</span>`;
  }

  // Reviewer step locked — preparer has not signed off yet but reviewer is assigned.
  // Only show Locked when there IS a reviewer waiting — not for unassigned sequential tasks.
  if (assignment.SignOffMode !== SIGN_OFF_MODE.PREPARER_ONLY &&
      assignment.Reviewer &&
      !assignment.PreparerSignOff &&
      !assignment.ReviewerSignOff) {
    return `<span class="status-badge status-notstarted">Locked</span>`;
  }

  const map = {
       [STATUS.COMPLETE]:    ['status-complete',   '✓ Complete'],
       [STATUS.PREPARED]:    ['status-prepared',   '→ Ready for review'],
       [STATUS.IN_PROGRESS]: ['status-progress',   'In progress'],
       [STATUS.NOT_STARTED]: ['status-notstarted', 'Not started'],
  };
  const [cls, label] = map[s] || map[STATUS.NOT_STARTED];
  return `<span class="status-badge ${cls}">${label}</span>`;
}

// Status severity order used when sorting by status or overdue-first.
const STATUS_ORDER = { 'Overdue': 0, [STATUS.IN_PROGRESS]: 1, [STATUS.NOT_STARTED]: 2, [STATUS.PREPARED]: 3, [STATUS.COMPLETE]: 4 };

function getEffectiveStatus(a) {
  return isTaskOverdue(a) ? 'Overdue' : (a.Status || STATUS.NOT_STARTED);
}

function getFilteredAssignments() {
  const f = STATE.filters;

  const filtered = STATE.assignments.filter(a => {
    // Show skipped only when explicitly requested
    if (a.IsSkipped && !STATE.filters.showSkipped) return false;
    if (f.status === 'unsigned' && a.Status === STATUS.COMPLETE) return false;
    if (f.status === 'overdue' && !isTaskOverdue(a)) return false;
    if (f.status === 'complete' && a.Status !== STATUS.COMPLETE) return false;
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
  const active = STATE.assignments.filter(a => !a.IsSkipped);
  if (!active.length) return 0;
  const complete = active.filter(a => a.Status === STATUS.COMPLETE).length;
  return Math.round((complete / active.length) * 100);
}

function getTaskRCCount(assignment) {
  return STATE.reviewComments.filter(rc => rc.TaskTemplateLookupId === assignment.TaskTemplateLookupId).length || 0;
}

function populateCategoryFilter() {
  const sel = document.getElementById('filter-category');
  if (!sel) return;
  const current = sel.value;
  const cats = [...new Set(STATE.assignments.filter(a => !a.IsSkipped).map(a => a.Category).filter(Boolean))].sort();
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
  const prepDate = resolveWorkday(getReadQuarter(), assignment.PreparerWorkday);
  const metaText = `${assignment.Category || ''} · Due WD${assignment.PreparerWorkday}${prepDate ? ' · ' + formatDateShort(prepDate) : ''}`;
  const metaEl = document.getElementById('panel-meta');
  if (STATE.isAdmin && !isViewingHistory()) {
    metaEl.innerHTML = `<span>${escapeHtml(metaText)}</span>`
      + ` <button class="btn-icon btn-sm" style="margin-left:6px;font-size:10px" data-action="edit-wd" data-id="${assignment._id}">Edit WD</button>`;
  } else {
    metaEl.textContent = metaText;
  }

  // Assignment section
  const email = STATE.currentUser?.Email;
  const prepBadge = renderBadge(assignment.Preparer);
  const revBadge = assignment.Reviewer ? renderBadge(assignment.Reviewer) : '—';
  // CATEGORY.TIE_OUT is an exact-match comparison — the SharePoint Category value must
  // match 'Tie Out' exactly (case-insensitive). If the category name ever changes,
  // update CATEGORY.TIE_OUT in the constants block above.
  const isTieOut = (assignment.Category || '').toLowerCase() === CATEGORY.TIE_OUT.toLowerCase();
  const docLink = isTieOut && assignment.HasDocumentLink && assignment.DocumentLink
    ? `<a class="panel-doc-link" href="${escapeHtml(assignment.DocumentLink)}" target="_blank">🔗 Open document</a>`
    : '';
  const showDocRow = isTieOut && (docLink || (STATE.isAdmin && !isViewingHistory()));
  const canReassign = STATE.isAdmin && !isViewingHistory() && !STATE.isReadOnly;
  document.getElementById('panel-assignment').innerHTML = `
    <div class="panel-meta-row"><span class="panel-meta-label">Preparer</span>${prepBadge}${canReassign ? `<button class="btn-icon btn-sm" style="margin-left:6px" data-action="reassign" data-id="${assignment._id}" data-role="preparer">Reassign</button>` : ''}</div>
    <div class="panel-meta-row"><span class="panel-meta-label">Reviewer</span>${revBadge}${canReassign && assignment.Reviewer ? `<button class="btn-icon btn-sm" style="margin-left:6px" data-action="reassign" data-id="${assignment._id}" data-role="reviewer">Reassign</button>` : ''}</div>
    <div class="panel-meta-row"><span class="panel-meta-label">Sign-off mode</span><span style="font-size:11px">${assignment.SignOffMode || '—'}</span></div>
    ${showDocRow ? `<div class="panel-meta-row" style="border-bottom:none"><span class="panel-meta-label">Document</span>${docLink || ''}${STATE.isAdmin && !isViewingHistory() ? `<button class="btn-icon btn-sm" style="margin-left:6px" data-action="edit-doc-link" data-id="${assignment._id}" data-url="${escapeHtml(assignment.DocumentLink || '')}" title="Edit document link">✏️</button>` : ''}</div>` : ''}`;

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
        <div class="rc-card ${rc.Priority === PRIORITY.URGENT ? 'urgent' : ''}" style="cursor:default">
          <div class="rc-meta">
            ${renderBadge(rc.CreatedBy)}
            <span class="rc-meta-text">${formatDateET(rc.CreatedDate) || '—'}</span>
            <span class="${rc.Priority === PRIORITY.URGENT ? 'badge-urgent' : 'badge-normal'}">${rc.Priority}</span>
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

  // Notes — preparer free-text field for documenting methodology, caveats, etc.
  const notesEl = document.getElementById('panel-notes');
  const notesInput = document.getElementById('panel-notes-input');
  const notesSave = document.getElementById('btn-panel-notes-save');
  const notesDisplay = document.getElementById('panel-notes-display');

  if (notesEl) {
    const canEditNotes = !isViewingHistory() && !STATE.isReadOnly &&
      (assignment.Preparer === STATE.currentUser?.Email || STATE.isAdmin);
    const currentNotes = assignment.Notes || '';

    // Show read-only display for non-editors; hide it for editors (they see the textarea instead)
    if (notesDisplay) {
      notesDisplay.textContent = currentNotes || 'No notes yet.';
      notesDisplay.style.color = currentNotes ? '' : 'var(--slate)';
      notesDisplay.style.display = canEditNotes ? 'none' : '';
    }
    if (notesInput) {
      notesInput.value = currentNotes;
      notesInput.style.display = canEditNotes ? '' : 'none';
    }
    if (notesSave) {
      notesSave.style.display = canEditNotes ? '' : 'none';
      notesSave.onclick = async () => {
        const newNotes = notesInput?.value?.trim() || '';
        try {
          await updateListItem(CONFIG.lists.quarterlyAssignments, assignment._id, { Notes: newNotes });
          patchAssignment(assignment._id, { Notes: newNotes });
          showToast('✓ Notes saved', 'success');
        } catch (err) {
          showToast(`Failed to save notes — ${classifyGraphError(err)}`, 'error');
        }
      };
    }
  }

  // Show panel
  document.getElementById('task-panel').classList.remove('hidden');
  document.getElementById('panel-overlay').classList.remove('hidden');
}

function renderPanelStatusChain(assignment, email) {
  const chain = document.getElementById('panel-status-chain');
  if (!chain) return;
  const isPrepOnly = assignment.SignOffMode === SIGN_OFF_MODE.PREPARER_ONLY;
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

  const isPrepOnly       = assignment.SignOffMode === SIGN_OFF_MODE.PREPARER_ONLY;
  const prepDone         = assignment.PreparerSignOff;
  const revDone          = assignment.ReviewerSignOff;
  const isPreparer       = assignment.Preparer === email;
  const isReviewer       = assignment.Reviewer === email;
  const isAdmin          = STATE.isAdmin;
  const isFinalReviewer  = STATE.isFinalReviewer;

  // RULE 3: Preparer steps — any team member can sign off (always shown).
  // RULE 4: Reviewer steps — restricted to assigned reviewer, admin, FinalReviewer.
  //         Everyone else sees an "on behalf" override button that logs the actual signer.
  const canSignPreparer  = !STATE.isReadOnly;  // ReadOnly users cannot sign off
  const canSignReviewer  = isReviewer || isAdmin || isFinalReviewer;

  // Reversals stay restricted — only assigned person or admin can reverse.
  const canReversePreparer = (isPreparer || isAdmin) && !STATE.isReadOnly;
  const canReverseReviewer = (isReviewer || isAdmin || isFinalReviewer) && !STATE.isReadOnly;

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

  // Admin skip/unskip button — always shown at bottom for admins on live quarters
  if (STATE.isAdmin && !isViewingHistory() && !STATE.isReadOnly) {
    const isSkipped = assignment.IsSkipped;
    html += `<div style="margin-top:12px;padding-top:10px;border-top:1px solid var(--mid-gray)">
      <button class="btn-${isSkipped ? 'secondary' : 'danger'} btn-sm" 
        data-action="${isSkipped ? 'unskip-task' : 'skip-task'}" data-id="${assignment._id}"
        title="${isSkipped ? 'Restore this task to the quarter' : 'Remove this task from the quarter — it will not appear in any view or count'}">
        ${isSkipped ? '↩ Restore task' : 'Skip this quarter'}
      </button>
      ${isSkipped ? '<span style="font-size:10px;color:var(--slate);margin-left:8px">This task is currently skipped</span>' : ''}
    </div>`;
  }

  actionDiv.innerHTML = html;
  attachCardEvents();
}

function renderMyTasksTable() {
  const email = STATE.currentUser?.Email;
  const tasks = filterMyAssignments();
  const tbody = document.getElementById('my-tasks-tbody');
  if (!tbody) return;

  tbody.innerHTML = '';
  tasks.forEach(a => {
    const isMyPreparer = a.Preparer === email && !a.PreparerSignOff;
    const isMyReviewer = a.Reviewer === email && a.PreparerSignOff && !a.ReviewerSignOff;
    const role = isMyPreparer ? 'Preparer' : isMyReviewer ? 'Reviewer' : 'Observer';
    const dueWD = getDueWD(a, email);
    const dueDate = resolveWorkday(getReadQuarter(), dueWD);
    const overdue = isTaskOverdue(a);

    const row = tbody.insertRow();
    if (overdue) row.classList.add('overdue-row');
    row.style.cursor = 'pointer';
    row.addEventListener('click', () => openTaskPanel(a._id));
    row.innerHTML = `
      <td style="font-weight:500;font-size:12px">${escapeHtml(a.Title || '')}</td>
      <td><span class="cat-tag">${escapeHtml(a.Category || '')}</span></td>
      <td style="font-size:11px">${role}</td>
      <td style="font-size:11px;color:var(--slate)">${a.PreparerWorkday ? 'WD' + a.PreparerWorkday : '—'}</td>
      <td style="font-size:11px;color:var(--slate)">${a.ReviewerWorkday ? 'WD' + a.ReviewerWorkday : '—'}</td>
      <td>${renderStatusBadge(a)}</td>
      <td style="font-size:11px;color:${overdue ? 'var(--red)' : 'var(--slate)'}">${dueDate ? formatDateShort(dueDate) + ' (WD' + dueWD + ')' : '—'}</td>`;
  });
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
    const urgent  = STATE.reviewComments.filter(rc => rc.Priority === PRIORITY.URGENT && rc.Status === RC_STATUS.OPEN).length;
    const open    = STATE.reviewComments.filter(rc => rc.Status === RC_STATUS.OPEN).length;
    const resolved = STATE.reviewComments.filter(rc => rc.Status === RC_STATUS.RESOLVED).length;
    sub.textContent = `${quarter || '—'} · ${urgent} urgent · ${open} open · ${resolved} resolved`;
    document.getElementById('rc-urgent-count').textContent = urgent;
    document.getElementById('rc-open-count').textContent = STATE.reviewComments.filter(rc => rc.Priority === PRIORITY.NORMAL && rc.Status === RC_STATUS.OPEN).length;
    document.getElementById('rc-resolved-count').textContent = resolved;
  }

  const urgentList = document.getElementById('rc-urgent-list');
  const openList   = document.getElementById('rc-open-list');
  const resolvedList = document.getElementById('rc-resolved-list');

  const allRCs  = rcQuarter ? STATE.reviewComments.filter(rc => rc.Quarter === rcQuarter) : STATE.reviewComments;
  const urgent  = allRCs.filter(rc => rc.Priority === PRIORITY.URGENT && rc.Status === RC_STATUS.OPEN);
  const normal  = allRCs.filter(rc => rc.Priority === PRIORITY.NORMAL && rc.Status === RC_STATUS.OPEN);
  const resolved = allRCs.filter(rc => rc.Status === RC_STATUS.RESOLVED);

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
    <div class="rc-card ${rc.Priority === PRIORITY.URGENT ? 'urgent' : ''} ${isResolved ? 'resolved' : ''}">
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
            <span class="${rc.Priority === PRIORITY.URGENT ? 'badge-urgent' : 'badge-normal'}">${rc.Priority}</span>
            <span class="${isResolved ? 'badge-resolved' : 'badge-open'}">${isResolved ? '✓ Resolved' : 'Open'}</span>
          </div>
          ${replyBadge}
        </div>
      </div>
      <div class="rc-comment-text">"${escapeHtml(rc.CommentText || '')}"</div>
      <div class="rc-meta">
        ${renderBadge(rc.CreatedBy)}
        <span class="rc-meta-text">${formatDateET(rc.CreatedDate) || '—'}</span>
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
      <div class="rc-reply-meta">${renderBadge(r.CreatedByEmail)} · ${formatDateET(r.CreatedDate)}${r.TaggedUsers ? ' · Tagged: ' + r.TaggedUsers.split(';').filter(Boolean).map(e => renderBadge(e.trim())).join('') : ''}</div>
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

  // Build sections dynamically from templates — supports any MatrixSection value.
  // For Q4 quarters, 'Form 10-Q' is automatically renamed to 'Form 10-K'.
  const filingType = isQuarterQ4(quarter) ? FILING.K : FILING.Q;
  const sections = {};

  STATE.templates
    .filter(t => t.MatrixItem && t.MatrixSection && (t.FilingType === filingType || t.FilingType === FILING.BOTH))
    .forEach(t => {
      // Rename 'Form 10-Q' → 'Form 10-K' in Q4 quarters
      let section = t.MatrixSection;
      if (isQuarterQ4(quarter) && section === 'Form 10-Q') section = 'Form 10-K';
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
      <th class="left-align" style="min-width:70px">Preparer</th>
      <th class="left-align" style="min-width:70px">1st Reviewer</th>
      ${checkpoints.map(cp => {
        const isMatrixOnly = CONFIG.matrixOnlyColumns.includes(cp);
        const isFinal = cp === CHECKPOINT.FINAL_REVIEW;
        const cls = isFinal ? 'final-col' : isMatrixOnly ? 'matrix-only-col' : '';
        return `<th class="${cls}" style="min-width:52px" title="${escapeHtml(cp)}">${escapeHtml(cp)}</th>`;
      }).join('')}
    </tr></thead>
    <tbody>`;

  Object.entries(sections).forEach(([sectionName, items]) => {
    if (!items.length) return;
    html += `<tr class="section-header"><td colspan="${3 + checkpoints.length}">${escapeHtml(sectionName)}</td></tr>`;

    items.forEach(item => {
      // Get preparer and reviewer from assignments
      const assignments = STATE.assignments.filter(a => a.MatrixItem === item.name);

      // Hide row entirely if every task-linked assignment is skipped this quarter.
      // Matrix-only columns (Final Review etc.) don't count — they're independent.
      const taskLinked = assignments.filter(a =>
        !CONFIG.matrixOnlyColumns.includes(a.MatrixCheckpoint) && a.MatrixCheckpoint
      );
      if (taskLinked.length > 0 && taskLinked.every(a => a.IsSkipped)) return;

      const preparers = [...new Set(
        assignments
          .filter(a => !a.IsSkipped && !a.MatrixCheckpoint?.toLowerCase().toLowerCase() === CATEGORY.XBRL.toLowerCase())
          .map(a => a.Preparer).filter(Boolean)
      )];
      const reviewers = [...new Set(
        assignments
          .filter(a => !a.IsSkipped && !a.MatrixCheckpoint?.toLowerCase().toLowerCase() === CATEGORY.XBRL.toLowerCase())
          .map(a => a.Reviewer).filter(Boolean)
      )];

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
          const status = ms?.[fm.status] || STATUS.NOT_STARTED;
          const isFinalReview = cp === CHECKPOINT.FINAL_REVIEW;
          const canAct = isFinalReview ? (STATE.isFinalReviewer || STATE.isAdmin) : !STATE.isReadOnly;

          const isFinalCell    = cp === CHECKPOINT.FINAL_REVIEW;
          const isMatrixOnlyCell = CONFIG.matrixOnlyColumns.includes(cp);
          const cellClass = isFinalCell ? 'final-td' : isMatrixOnlyCell ? 'matrix-only-td' : '';
          if (status === STATUS.COMPLETE) {
            const tooltip = `Signed off by ${ms?.[fm.by] || '—'} · ${formatDateET(ms?.[fm.date])}`;
            html += `<td class="cell-done ${cellClass}" title="${escapeHtml(tooltip)}">
              <svg width="12" height="12" viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
            </td>`;
          } else if (status === 'N/A') {
            html += `<td class="cell-na ${cellClass}" title="Not applicable">N/A</td>`;
          } else if (canAct && !isViewingHistory()) {
            html += `<td class="cell-actionable ${cellClass}" data-action="matrix-update" data-item="${escapeHtml(item.name)}" data-col="${escapeHtml(cp)}" title="Click to update">
              <svg width="10" height="10" viewBox="0 0 10 10"><circle cx="5" cy="5" r="3.5" fill="none" stroke="#3B72C0" stroke-width="1.5"/></svg>
            </td>`;
          } else {
            html += `<td class="cell-empty ${cellClass}"></td>`;
          }
        } else {
          // Task-linked column — use getCheckpointRole/getSignOffFields for consistent field access
          const linkedAssignment = STATE.assignments.find(
            a => a.MatrixItem === item.name && a.MatrixCheckpoint === cp
          );

          if (!linkedAssignment || linkedAssignment.IsSkipped) {
            html += `<td class="cell-na" title="${linkedAssignment?.IsSkipped ? 'Skipped this quarter' : 'Not applicable'}"></td>`;
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
      const isNA = cell.dataset.col === CHECKPOINT.FINAL_REVIEW && !STATE.isFinalReviewer;
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
  const complete = STATE.assignments.filter(a => a.Status === STATUS.COMPLETE).length;
  const overdue  = STATE.assignments.filter(a => !a.IsSkipped && isTaskOverdue(a)).length;
  const urgentRC = STATE.reviewComments.filter(rc => rc.Priority === PRIORITY.URGENT && rc.Status === RC_STATUS.OPEN).length;
  const pct = total ? Math.round((complete / total) * 100) : 0;

  const metricGrid = document.getElementById('metric-grid');
  if (metricGrid) {
    metricGrid.innerHTML = `
      <div class="metric-card"><div class="metric-label">Overall complete</div><div class="metric-value ${pct > 75 ? 'success' : ''}">${pct}%</div><div class="metric-sub">${complete} of ${total} tasks</div></div>
      <div class="metric-card"><div class="metric-label">Overdue tasks</div><div class="metric-value ${overdue > 0 ? 'danger' : ''}">${overdue}</div><div class="metric-sub">Across all categories</div></div>
      <div class="metric-card"><div class="metric-label">Urgent comments</div><div class="metric-value ${urgentRC > 0 ? 'danger' : ''}">${urgentRC}</div><div class="metric-sub">${STATE.reviewComments.filter(rc => rc.Status === RC_STATUS.OPEN).length} open total</div></div>
      <div class="metric-card"><div class="metric-label">Active quarter</div><div class="metric-value" style="font-size:18px;padding-top:4px">${quarter || '—'}</div><div class="metric-sub">${isQuarterQ4(quarter) ? '10-K · WD1–35' : '10-Q · WD1–20'}</div></div>`;
  }

  // Category progress
  const catBars = document.getElementById('category-bars');
  if (catBars) {
    const cats = [...new Set(STATE.assignments.filter(a => !a.IsSkipped).map(a => a.Category).filter(Boolean))].sort();
    catBars.innerHTML = cats.map(cat => {
      const catTasks = STATE.assignments.filter(a => a.Category === cat && !a.IsSkipped);
      const catComplete = catTasks.filter(a => a.Status === STATUS.COMPLETE).length;
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
      const done = myTasks.filter(a => a.Status === STATUS.COMPLETE).length;
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
    const upcoming = STATE.milestones
      .map(m => ({
        ...m,
        ActualDate: m.MilestoneDate || STATE.calendar.find(c => Number(c.WorkdayNumber) === Number(m.WorkdayNumber))?.ActualDate,
      }))
      .filter(m => m.ActualDate && m.ActualDate >= today)
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
  const overdueTasks = STATE.assignments.filter(a => !a.IsSkipped && isTaskOverdue(a));
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

  const quarter = getReadQuarter() || STATE.workingQuarter;
  const sub = document.getElementById('calendar-sub');
  if (sub) sub.textContent = `${quarter || '—'} · Close calendar`;

  const calBody = document.getElementById('cal-view-body');
  if (!calBody) return;

  if (!STATE.calendar.length) {
    // Try loading the calendar for the working quarter if no active quarter yet
    if (STATE.workingQuarter && !STATE._calendarLoading) {
      STATE._calendarLoading = true;
      calBody.innerHTML = '<p style="font-size:13px;color:var(--slate)">Loading calendar...</p>';
      Promise.all([
        loadCalendar(STATE.workingQuarter),
        loadMilestones(STATE.workingQuarter),
      ]).then(() => {
        STATE._calendarLoading = false;
        renderCalendarView();
      }).catch(() => {
        STATE._calendarLoading = false;
        calBody.innerHTML = '<p style="font-size:13px;color:var(--slate)">No calendar rows set up yet. Go to Admin → Close Calendar → Setup Calendar.</p>';
      });
      return;
    }
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
      <div class="cal-view-legend-item"><span class="milestone-team" style="padding:2px 8px;border-radius:8px">Team</span>&nbsp;Team deadlines</div>
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
  let lastMonth = -1;
  while (cursor <= endSunday) {
    // Determine the month label from the first actual workday in this week,
    // not the Monday — avoids showing "March" when WD1 is April 1st.
    const weekDates = [];
    for (let d = 0; d < 7; d++) {
      const dt = new Date(cursor.getTime() + d * 86400000);
      weekDates.push(toETDateStr(dt));
    }
    const firstWorkdayInWeek = weekDates.find(ds => byDate[ds]);
    const refDateStr = firstWorkdayInWeek || weekDates[0];
    const refDate = new Date(refDateStr + 'T12:00:00');
    const refET = new Date(refDate.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
    const refMonth = refET.getMonth();

    if (refMonth !== lastMonth) {
      lastMonth = refMonth;
      const monthName = refET.toLocaleString('en-US', { month: 'long', year: 'numeric', timeZone: CONFIG.timezone });
      html += `<div class="cal-month-header">${monthName}</div>`;
    }
    html += '<div class="cal-week-row">';
    for (let d = 0; d < 7; d++) {
      const dateStr = toETDateStr(cursor);
      const calRow  = byDate[dateStr];
      const isToday = dateStr === today;
      // Only dim past days during an active close — not post-close (entire quarter would grey out)
      const wd = getTodaysWorkday(STATE.activeQuarter);
      const isActive = wd !== 'post-close' && wd !== 'pre-close' && getReadQuarter() === STATE.activeQuarter;
      const isPast  = isActive && dateStr < today;

      const dayMilestones = STATE.milestones.filter(m => (m.MilestoneDate || '') === dateStr);
      if (!calRow) {
        if (dayMilestones.length) {
          // Non-workday but has milestones — show a minimal cell
          html += `<div class="cal-day" style="opacity:0.85;border-style:dashed">
            <div class="cal-day-top"><span class="cal-day-date">${formatDateShort(dateStr)}</span></div>
            ${dayMilestones.map(m => `<span class="cal-ms ${milestoneClass(m)}">${escapeHtml(m.MilestoneLabel)}</span>`).join('')}
          </div>`;
        } else {
          const etD = new Date(cursor.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
          const isWknd = etD.getDay() === 0 || etD.getDay() === 6;
          html += `<div class="cal-day empty">
            <div class="cal-day-top"><span class="cal-day-date" style="color:${isWknd ? '#bbb' : 'var(--color-text-tertiary)'}">${formatDateShort(dateStr)}</span></div>
          </div>`;
        }
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
            <span class="cal-day-date">${formatDateShort(dateStr)}</span>
          </div>
          ${calRow.IsWeekend ? '<span class="cal-wknd-flag">Weekend</span>' : ''}
          ${STATE.milestones
            .filter(m => (m.MilestoneDate || '') === dateStr)
            .map(m => `<span class="cal-ms ${milestoneClass(m)}">${escapeHtml(m.MilestoneLabel)}</span>`)
            .join('')}
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
    case 'calendar':
      if (!STATE.calendar.length && STATE.workingQuarter) {
        content.innerHTML = '<p style="font-size:12px;color:var(--slate);padding:12px">Loading calendar...</p>';
        Promise.all([
          loadCalendar(STATE.workingQuarter),
          loadMilestones(STATE.workingQuarter),
        ]).then(() => {
          content.innerHTML = renderAdminCalendar();
          attachAdminEvents('calendar');
        }).catch(() => {
          content.innerHTML = renderAdminCalendar();
        });
      } else {
        content.innerHTML = renderAdminCalendar();
      }
      break;
    case 'rollforward': content.innerHTML = renderAdminRollforward(); break;
    case 'templates':
      content.innerHTML = '<p style="font-size:12px;color:var(--slate);padding:12px">Loading templates...</p>';
      loadTemplates().then(() => {
        const activeBtn = document.querySelector('.sidebar-btn.active');
        if (activeBtn?.dataset.panel === 'templates') {
          content.innerHTML = renderAdminTemplates();
          attachAdminEvents('templates');
        }
      }).catch(err => {
        content.innerHTML = `<p style="color:var(--red);padding:12px">Failed to load templates — ${classifyGraphError(err)}</p>`;
      });
      return;
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
    case 'users':
      content.innerHTML = '<p style="font-size:12px;color:var(--slate);padding:12px">Loading users...</p>';
      loadUsers().then(() => {
        const activeBtn = document.querySelector('.sidebar-btn.active');
        if (activeBtn?.dataset.panel === 'users') {
          content.innerHTML = renderAdminUsers();
          attachAdminEvents('users');
        }
      }).catch(err => {
        content.innerHTML = `<p style="color:var(--red);padding:12px">Failed to load users — ${classifyGraphError(err)}</p>`;
      });
      return;
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
        ${STATE.workingQuarter ? `<button class="btn-secondary btn-sm" id="btn-edit-staging" data-action="edit-staging">Edit staging</button>
        <button class="btn-success btn-sm" id="btn-activate-quarter" data-action="activate-quarter">Activate ${STATE.workingQuarter}</button>` : ''}
      </div>
    </div>
    <div class="card" style="margin-bottom:12px">
      <div class="card-title" style="display:flex;align-items:center;justify-content:space-between">
        System diagnostics
        <button class="btn-secondary btn-sm" id="btn-run-diagnostics" data-action="run-diagnostics">Run diagnostics</button>
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
  const items = STATE.milestones.map(m => ({
    ...m,
    ActualDate: m.MilestoneDate || STATE.calendar.find(c => Number(c.WorkdayNumber) === Number(m.WorkdayNumber))?.ActualDate,
  })).slice(0, 8);
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
        </td>
      </tr>`).join('')}
    </tbody></table>`;
}

function renderAdminCalendar() {
  const quarter = STATE.activeQuarter || STATE.workingQuarter;
  const hasRows = STATE.calendar.length > 0;

  // Build a full date-range view — every calendar day from first WD to last WD,
  // including non-workday days so milestones can be added on any date.
  let fullDayRows = '';
  if (hasRows) {
    const firstDate = new Date(STATE.calendar[0].ActualDate + 'T12:00:00');
    const lastDate  = new Date(STATE.calendar[STATE.calendar.length - 1].ActualDate + 'T12:00:00');
    const DAY_NAMES = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const today     = todayET();

    // Build a lookup: dateStr → calendar row
    const calByDate = {};
    STATE.calendar.forEach(c => { calByDate[c.ActualDate] = c; });

    // Build a lookup: dateStr → milestones using MilestoneDate field
    const milestonesByDate = {};
    STATE.milestones.forEach(m => {
      const dateKey = m.MilestoneDate || (() => {
        // Fallback: derive date from WorkdayNumber for older records without MilestoneDate
        const calRow = STATE.calendar.find(c => Number(c.WorkdayNumber) === Number(m.WorkdayNumber));
        return calRow?.ActualDate;
      })();
      if (dateKey) {
        if (!milestonesByDate[dateKey]) milestonesByDate[dateKey] = [];
        milestonesByDate[dateKey].push(m);
      }
    });

    const rows = [];
    let cursor = new Date(firstDate);
    while (cursor <= lastDate) {
      const etDate  = new Date(cursor.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
      const dateStr = `${etDate.getFullYear()}-${String(etDate.getMonth()+1).padStart(2,'0')}-${String(etDate.getDate()).padStart(2,'0')}`;
      const dow     = etDate.getDay();
      const dowName = DAY_NAMES[dow];
      const calRow  = calByDate[dateStr];
      const isToday = dateStr === today;
      const milestones = milestonesByDate[dateStr] || [];
      const isNonWorkday = !calRow;

      // For non-workday rows, milestones attach to the nearest preceding WD
      // We use a virtual WD number for the add-milestone button
      const nearestWD = calRow ? calRow.WorkdayNumber
        : STATE.calendar.filter(c => c.ActualDate < dateStr).pop()?.WorkdayNumber || 0;

      rows.push(`
        <tr style="${isNonWorkday ? 'opacity:0.55;background:var(--light-gray)' : ''}${isToday ? ';background:var(--blue-tint,#EEF3FF)' : ''}">
          <td style="font-weight:${isNonWorkday ? '400' : '600'};color:${isNonWorkday ? 'var(--slate)' : 'inherit'}">
            ${calRow ? `WD${calRow.WorkdayNumber}${calRow.IsWeekend ? ' <span class="weekend-marker">Wknd</span>' : ''}` : '—'}
          </td>
          <td style="color:var(--slate);white-space:nowrap">
            <span style="font-size:10px;color:${dow===0||dow===6?'var(--red)':'var(--slate)'};margin-right:4px">${dowName}</span>
            ${formatDateShort(dateStr)}
          </td>
          <td>
            ${milestones.map(m => `
              <span class="${milestoneClass(m)}" style="display:inline-flex;align-items:center;gap:4px;margin:1px">
                ${escapeHtml(m.MilestoneLabel)}
                <button class="btn-icon" style="font-size:9px;padding:0 3px;line-height:1.4"
                  data-action="delete-milestone" data-id="${m._id}" title="Remove">✕</button>
              </span>`).join('')}
            <button class="btn-secondary btn-sm" style="margin-left:4px;font-size:10px"
              data-action="add-milestone" data-wd="${nearestWD}" data-date="${dateStr}">+ Milestone</button>
          </td>
          <td>
            ${calRow
              ? `<button class="btn-secondary btn-sm" data-action="edit-cal-row" data-id="${calRow._id}">Edit date</button>`
              : '<span style="font-size:11px;color:var(--slate)">Non-workday</span>'}
          </td>
        </tr>`);

      cursor = new Date(cursor.getTime() + 86400000);
    }
    fullDayRows = rows.join('');
  }

  return `
    <div class="admin-section-title">Close Calendar</div>
    <div class="admin-section-sub">${quarter || 'No active quarter'}</div>
    <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center">
      <button class="btn-primary btn-sm" id="btn-setup-calendar" data-action="setup-calendar">Setup Calendar…</button>
      <span style="font-size:11px;color:var(--slate)">${hasRows ? `${STATE.calendar.length} workdays · all days shown` : 'No workdays set up yet — click Setup Calendar to create them'}</span>
    </div>
    <div class="card">
      ${!hasRows
        ? `<p style="font-size:12px;color:var(--slate);padding:8px 0">No calendar rows yet. Click <strong>Setup Calendar</strong> to create workday rows for this quarter.</p>`
        : `<table class="cal-table">
            <thead><tr><th>WD</th><th>Date</th><th>Milestones</th><th>Actions</th></tr></thead>
            <tbody>${fullDayRows}</tbody>
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
        <button class="btn-primary btn-sm" id="btn-start-new-quarter" data-action="start-new-quarter">Start New Quarter</button>
        ${STATE.workingQuarter ? `<button class="btn-secondary btn-sm" id="btn-rollforward" data-action="rollforward">Roll Forward from ${STATE.activeQuarter || 'previous'}</button>` : ''}
        ${STATE.workingQuarter ? `<button class="btn-success btn-sm" id="btn-activate-quarter-rf" data-action="activate-quarter-rf">Activate ${STATE.workingQuarter}</button>` : ''}
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
    if (!STATE._stagingLoading) {
      STATE._stagingLoading = true;
      getListItems(CONFIG.lists.quarterlyAssignments,
        `fields/Quarter eq '${STATE.workingQuarter}' and fields/IsStaging eq 1`
      ).then(items => {
        STATE._stagingItems = items.map(i => ({ ...i.fields, _id: i.id }));
        STATE._stagingLoading = false;
        // Only update the staging grid container — don't re-render the whole panel
        // so the Roll Forward / Activate buttons are never destroyed mid-click.
        const gridContainer = document.getElementById('staging-grid-container');
        if (gridContainer) {
          gridContainer.innerHTML = renderStagingGrid();
          attachAdminEvents('rollforward');
        }
      }).catch(() => { STATE._stagingLoading = false; });
    }
    return `<div class="card" id="staging-grid-container">
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
      <tr class="${item.IsSkipped ? 'staging-skipped' : ''}">
        <td style="font-size:11px;max-width:160px">${escapeHtml(item.Title || '')}${item.IsSkipped ? ' <span style="font-size:9px;color:var(--slate);font-weight:500">SKIPPED</span>' : ''}</td>
        <td><span class="cat-tag">${escapeHtml(item.Category || '')}</span></td>
        <td>
          <input type="number" class="staging-select staging-wd" data-id="${item._id}" data-field="PreparerWorkday"
            value="${item.PreparerWorkday || ''}" min="1" max="35"
            style="font-size:11px;width:52px;text-align:center" title="Preparer workday" />
        </td>
        <td>
          ${item.SignOffMode === SIGN_OFF_MODE.PREPARER_ONLY
            ? '<span style="font-size:10px;color:var(--slate)">—</span>'
            : `<input type="number" class="staging-select staging-wd" data-id="${item._id}" data-field="ReviewerWorkday"
                value="${item.ReviewerWorkday || ''}" min="1" max="35"
                style="font-size:11px;width:52px;text-align:center" title="Reviewer workday" />`}
        </td>
        <td>
          <select class="staging-select staging-mode" data-id="${item._id}" data-field="SignOffMode" style="font-size:11px;width:100%">
            <option value="Sequential"${item.SignOffMode === SIGN_OFF_MODE.SEQUENTIAL ? ' selected' : ''}>Sequential</option>
            <option value="Preparer Only"${item.SignOffMode === SIGN_OFF_MODE.PREPARER_ONLY ? ' selected' : ''}>Prep Only</option>
          </select>
        </td>
        <td>
          <select class="staging-select" data-id="${item._id}" data-field="Preparer" style="font-size:11px;max-width:130px">
            ${blankOpt}${userOpts.replace(`value="${escapeHtml(item.Preparer)}"`, `value="${escapeHtml(item.Preparer)}" selected`)}
          </select>
        </td>
        <td>
          ${item.SignOffMode === SIGN_OFF_MODE.PREPARER_ONLY
            ? '<span style="font-size:10px;color:var(--slate)">—</span>'
            : `<select class="staging-select" data-id="${item._id}" data-field="Reviewer" style="font-size:11px;max-width:130px">
                ${blankOpt}${userOpts.replace(`value="${escapeHtml(item.Reviewer)}"`, `value="${escapeHtml(item.Reviewer)}" selected`)}
              </select>`}
        </td>
        <td style="text-align:center">
          <input type="checkbox" class="staging-select staging-skip" data-id="${item._id}" data-field="IsSkipped"
            ${item.IsSkipped ? 'checked' : ''} title="Skip this task for ${STATE.workingQuarter}" />
        </td>
        <td><span class="staging-save-indicator" style="font-size:10px;white-space:nowrap"></span></td>
      </tr>`).join('');

  return `
    <div class="card" id="staging-grid-container">
      <div class="card-title" style="display:flex;align-items:center;justify-content:space-between">
        Staging grid — ${STATE.workingQuarter}
        <span style="font-size:11px;font-weight:400;color:var(--slate)">${stagingItems.length} assignments · ${stagingItems.filter(i => i.IsSkipped).length} skipped · changes save instantly</span>
      </div>
      <p style="font-size:12px;color:var(--slate);margin-bottom:10px">Review and adjust workday numbers, preparers, and reviewers before activating. Changes here only affect the staging quarter.</p>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:12px;padding:10px 12px;background:var(--light-gray);border-radius:6px">
        <span style="font-size:11px;font-weight:600;color:var(--slate)">Bulk assign:</span>
        <select id="bulk-assign-category" class="field-input" style="font-size:11px;width:160px;padding:4px 8px">
          <option value="all">All categories</option>
          ${[...new Set(stagingItems.map(i => i.Category).filter(Boolean))].sort().map(c =>
            `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('')}
        </select>
        <select id="bulk-assign-preparer" class="field-input" style="font-size:11px;width:160px;padding:4px 8px">
          <option value="">Set Preparer…</option>
          ${STATE.users.filter(u => u.IsActive !== false).map(u =>
            `<option value="${escapeHtml(u.Email)}">${escapeHtml((u.Emoji || '') + ' ' + (u.Title || u.Email.split('@')[0]))}</option>`).join('')}
        </select>
        <select id="bulk-assign-reviewer" class="field-input" style="font-size:11px;width:160px;padding:4px 8px">
          <option value="">Set Reviewer…</option>
          <option value="__clear__">— Clear reviewer —</option>
          ${STATE.users.filter(u => u.IsActive !== false).map(u =>
            `<option value="${escapeHtml(u.Email)}">${escapeHtml((u.Emoji || '') + ' ' + (u.Title || u.Email.split('@')[0]))}</option>`).join('')}
        </select>
        <button class="btn-primary btn-sm" id="btn-bulk-assign-apply" data-action="bulk-assign">Apply</button>
        <span id="bulk-assign-status" style="font-size:11px;color:var(--slate)"></span>
      </div>
      <div class="table-wrap">
        <table class="data-table" style="table-layout:fixed;width:100%">
          <colgroup>
            <col style="width:18%"/><col style="width:10%"/><col style="width:6%"/>
            <col style="width:6%"/><col style="width:10%"/><col style="width:21%"/><col style="width:21%"/><col style="width:8%"/>
          </colgroup>
          <thead><tr>
            <th>Task</th><th>Category</th>
            <th title="Preparer Workday">Prep WD</th>
            <th title="Reviewer Workday">Rev WD</th>
            <th title="Sign-off mode">Mode</th>
            <th>Preparer</th><th>Reviewer</th>
            <th title="Skip this quarter">Skip</th>
            <th style="width:50px"></th>
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
      <button class="btn-primary btn-sm" id="btn-new-template" data-action="new-template">+ New Template</button>
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
              <td style="font-size:11px;white-space:nowrap">
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
      <button class="btn-primary btn-sm" id="btn-add-user" data-action="add-user">+ Add User</button>
    </div>
    <div class="table-wrap">
      <table class="data-table">
        <thead><tr><th>Name</th><th>Email</th><th>Role</th><th>Last Login</th><th>Actions</th></tr></thead>
        <tbody>
          ${STATE.users.map(u => `
            <tr>
              <td>${renderBadge(u.Email)}</td>
              <td style="font-size:11px">${escapeHtml(u.Email || '')}</td>
              <td><span class="cat-tag">${escapeHtml(u.Role || ROLE.TEAM_MEMBER)}</span></td>
              <td style="font-size:11px">${u.LastLogin ? formatDateET(u.LastLogin) : '—'}</td>
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
    FinalReview:           'background:#EEEDFE;color:#3C3489',
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

  const allTypes    = ['All','SignOff','Reversal','Reassignment','FinalReview'];
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
        <div>${formatDateShort(e.ActionDate?.split('T')[0] || e.ActionDate)}</div>
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
      <button class="btn-secondary btn-sm" id="btn-export-audit-excel" data-action="export-audit-excel">Export CSV</button>
      <button class="btn-primary btn-sm" id="btn-export-sox" data-action="export-sox">Audit Log Export…</button>
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
        <button class="btn-secondary btn-sm" id="btn-validate-import" data-action="validate-import">Validate</button>
        <button class="btn-primary btn-sm" id="btn-run-import" data-action="run-import" disabled>Import</button>
      </div>
      <div id="import-status" style="margin-top:12px;font-size:12px;color:var(--slate)"></div>
      <div id="import-progress" style="margin-top:8px"></div>
    </div>`;
}

function attachAdminEvents(panelName) {
  // Note: all modal button listeners are in attachGlobalEvents (run once at startup).
  // All dynamic button actions use data-action delegation on admin-content.
  // Only panel-specific input listeners that need re-attaching per render go here.

  // Template search — use a flag to avoid stacking on repeated renders
  const templateSearch = document.getElementById('template-search');
  if (templateSearch && !templateSearch.dataset.listenerAttached) {
    templateSearch.dataset.listenerAttached = 'true';
    templateSearch.addEventListener('input', e => filterTemplateTable(e.target.value));
  }
  // Staging grid — save preparer/reviewer on dropdown change
  const adminContent2 = document.getElementById('admin-content');
  if (adminContent2 && !adminContent2.dataset.stagingEventsAttached) {
    adminContent2.dataset.stagingEventsAttached = 'true';
    const _stagingDebounce = {};
    adminContent2.addEventListener('change', e => {
      const sel = e.target.closest('.staging-select');
      if (!sel) return;
      const { id, field } = sel.dataset;

      // WD fields are numbers; skip is boolean; SignOffMode and person fields are strings.
      const isWD   = field === 'PreparerWorkday' || field === 'ReviewerWorkday';
      const isSkip = field === 'IsSkipped';
      const raw    = isSkip ? sel.checked : sel.value;
      const value  = isWD   ? (raw ? Number(raw) : null)
                   : isSkip ? Boolean(raw)
                   : (raw || null);

      // Validate WD range immediately — before debounce
      if (isWD && value !== null && (value < 1 || value > 35)) {
        showToast('Workday must be between 1 and 35', 'error');
        return;
      }

      // Show pending indicator
      const row = sel.closest('tr');
      const indicator = row?.querySelector('.staging-save-indicator');
      if (indicator) { indicator.textContent = 'Saving…'; indicator.style.color = 'var(--slate)'; }

      // Debounce: cancel pending save for this field, reschedule
      clearTimeout(_stagingDebounce[id + field]);
      _stagingDebounce[id + field] = setTimeout(async () => {
        try {
          await updateListItem(CONFIG.lists.quarterlyAssignments, id, { [field]: value });
          const item = STATE._stagingItems.find(i => i._id === id);
          if (item) item[field] = value;
          if (indicator) { indicator.textContent = '✓'; indicator.style.color = 'var(--green-mid)'; }
          setTimeout(() => { if (indicator) indicator.textContent = ''; }, 2000);
          if (isSkip || field === 'SignOffMode') {
            const gridContainer = document.getElementById('staging-grid-container');
            if (gridContainer) {
              gridContainer.outerHTML = renderStagingGrid();
              attachAdminEvents('rollforward');
            }
          }
        } catch (err) {
          if (indicator) { indicator.textContent = '✗ Failed'; indicator.style.color = 'var(--red)'; }
          showToast(`Failed to update ${field} — ${classifyGraphError(err)}`, 'error');
          logError('Staging grid update failed:', err);
        }
      }, 350);
    });
  }

  const adminContentEl = document.getElementById('admin-content');
  if (adminContentEl && !adminContentEl.dataset.adminActionsAttached) {
    adminContentEl.dataset.adminActionsAttached = 'true';
    adminContentEl.addEventListener('click', async e => {
      const btn = e.target.closest('[data-action]');
      if (!btn) return;
      const { action, id, email } = btn.dataset;

      if (action === 'start-new-quarter')   startNewQuarter();
      if (action === 'rollforward')           performRollforward();
      if (action === 'activate-quarter-rf')  confirmActivation();
      if (action === 'activate-quarter')      confirmActivation();
      if (action === 'edit-staging')          renderAdminPanel('rollforward');
      if (action === 'run-diagnostics')       runDiagnostics();
      if (action === 'setup-calendar') {
        const calQuarter = STATE.activeQuarter || STATE.workingQuarter || '';
        const quarterEl  = document.getElementById('cal-setup-quarter');
        const maxWDEl    = document.getElementById('cal-setup-maxwd');
        const errEl      = document.getElementById('cal-setup-error');
        if (quarterEl) quarterEl.value = calQuarter;
        if (maxWDEl)   maxWDEl.value   = isQuarterQ4(calQuarter) ? '35' : '20';
        if (errEl)     errEl.classList.add('hidden');
        showModal('modal-cal-setup');
      }
      if (action === 'new-template')          openEditTemplateModal(null);
      if (action === 'add-user')              openAddUserModal();
      if (action === 'export-audit-excel')    exportAuditLog();
      if (action === 'export-sox')            openSOXExportModal();
      if (action === 'validate-import')       validateImport();
      if (action === 'run-import')            runImport();
      if (action === 'clear-all-filters')     clearAllFilters();
      if (action === 'edit-template')         openEditTemplateModal(id);
      if (action === 'retire-template')   await retireTemplate(id);
      if (action === 'edit-cal-row')      openEditCalendarRowModal(id);
      if (action === 'edit-doc-link')      openEditDocLinkModal(id, btn.dataset.url);
      if (action === 'add-milestone')      openAddMilestoneModal(btn.dataset.wd, btn.dataset.date);
      if (action === 'bulk-assign')         confirmBulkAssign();
      if (action === 'delete-milestone')   deleteMilestone(id);
      if (action === 'edit-user')         openEditUserRoleModal(email);
      if (action === 'rc-reply')          openRCReplyInput(id);
      if (action === 'submit-rc-reply')   submitRCReply(id);
      if (action === 'cancel-rc-reply')   e.target.closest('.rc-reply-form')?.remove();
      if (action === 'approve-suggestion') await approveSuggestion(id);
      if (action === 'reject-suggestion') {
        STATE.pendingSuggestionReject = id;
        const noteEl = document.getElementById('reject-suggestion-note');
        if (noteEl) noteEl.value = '';
        showModal('modal-reject-suggestion');
      }
    });
  }

  // Template edit modal confirm/cancel — moved to attachGlobalEvents
  // Calendar edit modal confirm/cancel — moved to attachGlobalEvents
  // User role modal — moved to attachGlobalEvents
  // SOX/audit exports — moved to attachGlobalEvents
  // Import buttons — handled by delegation

  // Audit log filter dropdowns — re-render on change (must stay here, dropdowns are dynamic)
  ['audit-filter-type','audit-filter-person','audit-filter-quarter'].forEach(filterId => {
    const el = document.getElementById(filterId);
    if (el && !el.dataset.listenerAttached) {
      el.dataset.listenerAttached = 'true';
      el.addEventListener('change', e => {
        const field = filterId.replace('audit-filter-', '');
        STATE._auditFilter[field] = e.target.value;
        document.getElementById('admin-content').innerHTML = renderAdminAuditLog();
        attachAdminEvents('auditlog');
      });
    }
  });
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
    // Normalize quarter values in imported rows if present
    rows.forEach(r => { if (r.Quarter) r.Quarter = normalizeQuarter(r.Quarter); });

    const required = ['TaskName', 'Category', 'FilingType', 'SignOffMode', 'PreparerWorkday', 'IsActive'];
    const missing = required.filter(r => !rows[0] || !(r in rows[0]));
    if (missing.length) {
      if (status) status.textContent = `❌ Missing required columns: ${missing.join(', ')}`;
      return;
    }
    // Warn if any Sequential task is missing a ReviewerWorkday — it would never show as overdue
    const missingRevWD = rows.filter(r =>
      r.SignOffMode === SIGN_OFF_MODE.SEQUENTIAL && !r.ReviewerWorkday
    );
    if (missingRevWD.length) {
      if (status) status.textContent =
        `⚠ ${rows.length} tasks ready — ${missingRevWD.length} Sequential task(s) are missing ReviewerWorkday (${missingRevWD.map(r => r.TaskName).join(', ')}). These will never show as overdue for the reviewer.`;
      if (btnImport) { btnImport.disabled = false; btnImport.dataset.rows = JSON.stringify(rows); }
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
          Category:        row.Category || '',
          MatrixItem:      row.MatrixItem || null,
          MatrixCheckpoint:row.MatrixCheckpoint || null,
          MatrixSection:   row.MatrixSection || null,
          FilingType:      row.FilingType || FILING.BOTH,
          SignOffMode:     row.SignOffMode || SIGN_OFF_MODE.SEQUENTIAL,
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

  // Delay focus AND briefly disable all buttons in the modal to prevent
  // carryover clicks from the triggering button firing modal actions.
  const buttons = Array.from(modalEl.querySelectorAll('button'));
  buttons.forEach(b => { b.disabled = true; });
  setTimeout(() => {
    buttons.forEach(b => { b.disabled = false; });
    // Focus the modal title or a non-button element if possible,
    // otherwise focus the cancel button last to avoid accidental confirms.
    const cancelBtn = modalEl.querySelector('.btn-secondary, [id*="cancel"]');
    (cancelBtn || first).focus();
  }, 150);

  if (_modalKeyHandler) document.removeEventListener('keydown', _modalKeyHandler);

  _modalKeyHandler = (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      hideAllModals();
      return;
    }
    if (e.key !== 'Tab') return;
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
  const emoji = u.Emoji || '👤';
  const firstName = (u.Title || u.Email?.split('@')[0] || '').split(' ')[0];
  btn.innerHTML = `<span style="font-size:15px;line-height:1">${emoji}</span><span class="nav-user-name">${escapeHtml(firstName)}</span>`;
  btn.style.background = '';
  btn.style.borderColor = '';
}

// ============================================================
// EVENTS — GLOBAL
// ============================================================
function clearAllFilters() {
  STATE.filters.status   = 'all';
  STATE.filters.category = 'all';
  STATE.filters.assignee = 'all';
  STATE.filters.search   = '';
  saveFilters();
  renderAllTasks();
}

function openAddUserModal() {
  const emailEl    = document.getElementById('add-user-email');
  const nameEl     = document.getElementById('add-user-name');
  const roleEl     = document.getElementById('add-user-role');
  const errEl      = document.getElementById('add-user-error');
  const customEl   = document.getElementById('add-user-emoji-custom');
  const previewWrap = document.getElementById('add-user-preview-wrap');
  if (emailEl)     emailEl.value  = '';
  if (nameEl)      nameEl.value   = '';
  if (roleEl)      roleEl.value   = ROLE.TEAM_MEMBER;
  if (errEl)       errEl.classList.add('hidden');
  if (customEl)    customEl.value = '';
  if (previewWrap) previewWrap.style.display = 'none';
  STATE._addUserEmoji = null;
  STATE._addUserColor = null;
  renderEmojiPicker('add-user-emoji-grid', null, (emoji) => {
    STATE._addUserEmoji = emoji;
    const c = document.getElementById('add-user-emoji-custom');
    if (c) c.value = '';
    updateAddUserPreview();
  });
  renderColorPicker('add-user-color-grid', CONFIG.colorOptions[0].hex, (color) => {
    STATE._addUserColor = color;
    updateAddUserPreview();
  });
  showModal('modal-add-user');
}

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

  // ── Admin modal buttons ──────────────────────────────────────
  // Skip task modal
  document.getElementById('btn-skip-task-confirm')?.addEventListener('click', async () => {
    // Bulk assign takes priority — the modal is reused for both skip and bulk assign
    if (STATE._pendingBulkAssign) {
      const { targets, preparer, reviewer, category } = STATE._pendingBulkAssign;
      STATE._pendingBulkAssign = null;
      hideModal('modal-skip-task');
      await executeBulkAssign(targets, preparer, reviewer);
      return;
    }
    // Skip/restore flow
    const { id, isSkipping } = STATE._pendingSkip || {};
    if (!id) return;
    hideModal('modal-skip-task');
    updateListItem(CONFIG.lists.quarterlyAssignments, id, { IsSkipped: isSkipping })
      .then(() => {
        patchAssignment(id, { IsSkipped: isSkipping });
        showToast(isSkipping ? '✓ Task skipped' : '✓ Task restored', 'success');
        STATE._pendingSkip = null;
        closeTaskPanel();
        refreshCurrentView();
      })
      .catch(err => { showToast(`Failed — ${classifyGraphError(err)}`, 'error'); STATE._pendingSkip = null; });
  });
  document.getElementById('btn-skip-task-cancel')?.addEventListener('click', () => {
    hideModal('modal-skip-task'); STATE._pendingSkip = null;
  });
  // Add milestone modal
  document.getElementById('btn-add-milestone-confirm')?.addEventListener('click', confirmAddMilestone);
  document.getElementById('btn-add-milestone-cancel')?.addEventListener('click', () => hideModal('modal-add-milestone'));
  // Edit document link modal
  document.getElementById('btn-edit-doc-link-confirm')?.addEventListener('click', confirmEditDocLink);
  document.getElementById('btn-edit-doc-link-cancel')?.addEventListener('click', () => hideModal('modal-edit-doc-link'));
  // Add task mid-quarter modal
  document.getElementById('btn-add-task-confirm')?.addEventListener('click', confirmAddTask);
  document.getElementById('btn-add-task-cancel')?.addEventListener('click', () => hideModal('modal-add-task'));
  // Template edit
  document.getElementById('btn-edit-tpl-save')?.addEventListener('click', saveTemplateEdit);
  document.getElementById('btn-edit-tpl-cancel')?.addEventListener('click', () => { hideModal('modal-edit-template'); STATE.pendingTemplateEdit = null; });
  // Calendar edit
  document.getElementById('btn-edit-cal-save')?.addEventListener('click', saveCalendarRowEdit);
  document.getElementById('btn-edit-cal-cancel')?.addEventListener('click', () => { hideModal('modal-edit-calendar'); STATE.pendingCalendarEdit = null; });
  // User role edit
  document.getElementById('btn-edit-user-save')?.addEventListener('click', saveUserRoleEdit);
  document.getElementById('btn-edit-user-cancel')?.addEventListener('click', () => { hideModal('modal-edit-user'); STATE.pendingUserEdit = null; });
  // SOX export
  document.getElementById('btn-sox-confirm')?.addEventListener('click', confirmSOXExport);
  document.getElementById('btn-sox-cancel')?.addEventListener('click', () => hideModal('modal-sox-export'));
  // Edit WD
  document.getElementById('btn-edit-wd-confirm')?.addEventListener('click', confirmEditWD);
  document.getElementById('btn-edit-wd-cancel')?.addEventListener('click', () => { STATE.pendingWDEdit = null; hideModal('modal-edit-wd'); });
  // New quarter
  document.getElementById('btn-new-quarter-confirm')?.addEventListener('click', confirmNewQuarter);
  document.getElementById('btn-new-quarter-cancel')?.addEventListener('click', () => hideModal('modal-new-quarter'));
  document.getElementById('new-quarter-name')?.addEventListener('keydown', e => { if (e.key === 'Enter') confirmNewQuarter(); });
  // Rollforward
  document.getElementById('btn-rollforward-confirm')?.addEventListener('click', confirmRollforward);
  document.getElementById('btn-rollforward-cancel')?.addEventListener('click', () => { hideModal('modal-rollforward-confirm'); STATE.pendingRollforward = null; });
  // Reassign
  document.getElementById('btn-reassign-confirm')?.addEventListener('click', confirmReassign);
  document.getElementById('btn-reassign-cancel')?.addEventListener('click', () => { hideModal('modal-reassign'); STATE.pendingReassign = null; });
  // Calendar setup
  document.getElementById('btn-cal-setup-confirm')?.addEventListener('click', setupCalendarBulk);
  document.getElementById('btn-cal-setup-cancel')?.addEventListener('click', () => hideModal('modal-cal-setup'));
  // Cascade
  document.getElementById('btn-cascade-confirm')?.addEventListener('click', confirmCascade);
  document.getElementById('btn-cascade-no')?.addEventListener('click', () => { hideModal('modal-cascade'); STATE.pendingCascade = null; showToast('✓ Calendar row updated', 'success'); renderAdminPanel('calendar'); });
  // Add user
  document.getElementById('btn-add-user-confirm')?.addEventListener('click', createUser);
  document.getElementById('btn-add-user-cancel')?.addEventListener('click', () => hideModal('modal-add-user'));
  document.getElementById('add-user-email')?.addEventListener('keydown', e => { if (e.key === 'Enter') createUser(); });
  // Retire template
  document.getElementById('btn-retire-template-confirm')?.addEventListener('click', () => { hideModal('modal-retire-template'); confirmRetireTemplate(); });
  document.getElementById('btn-retire-template-cancel')?.addEventListener('click', () => { hideModal('modal-retire-template'); STATE.pendingTemplateRetire = null; });
  // Activate quarter
  document.getElementById('btn-activate-confirm')?.addEventListener('click', async () => { if (!STATE.pendingActivation) return; hideModal('modal-activate'); await activateQuarter(STATE.pendingActivation); STATE.pendingActivation = null; });
  document.getElementById('btn-activate-cancel')?.addEventListener('click', () => { hideModal('modal-activate'); STATE.pendingActivation = null; });

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
    // Reason is optional — a blank reason is stored as empty string, not blocked.
    // Forcing a reason produces garbage data ('n/a') which is worse than nothing.
    const reason = document.getElementById('reversal-reason')?.value?.trim() || '';
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
    if (!canPostReviewComment()) {
      showToast('You must be assigned as a reviewer to post review comments', 'error');
      return;
    }
    const sel = document.getElementById('rc-task-select');
    if (sel) sel.innerHTML = STATE.templates.map(t =>
      `<option value="${escapeHtml(t._id)}">${escapeHtml(t.TaskName || t.Title || '')}</option>`
    ).join('');
    showModal('modal-new-rc');
    renderRCTagPicker();
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

  // Modal listeners moved to attachGlobalEvents — wired once at startup

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

  // Table/card view toggle — All Tasks
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

  // Table/card view toggle — My Tasks
  document.getElementById('btn-my-tasks-cards')?.addEventListener('click', () => {
    document.getElementById('btn-my-tasks-cards').classList.add('active');
    document.getElementById('btn-my-tasks-table').classList.remove('active');
    document.getElementById('my-tasks-card-view')?.classList.remove('hidden');
    document.getElementById('my-tasks-table-view')?.classList.add('hidden');
  });
  document.getElementById('btn-my-tasks-table')?.addEventListener('click', () => {
    document.getElementById('btn-my-tasks-table').classList.add('active');
    document.getElementById('btn-my-tasks-cards').classList.remove('active');
    document.getElementById('my-tasks-card-view')?.classList.add('hidden');
    document.getElementById('my-tasks-table-view')?.classList.remove('hidden');
    renderMyTasksTable();
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
          const confirmBtn = document.getElementById('btn-signoff-confirm');
          if (titleEl) titleEl.textContent = `Sign off as ${role}?`;
          if (bodyEl) bodyEl.innerHTML = `
            <p style="font-size:13px;margin-bottom:8px">${escapeHtml(assignment.Title || '')}</p>
            <p style="font-size:12px;color:var(--slate)">Recorded as ${renderBadge(STATE.currentUser?.Email)} · ${formatDateET(new Date().toISOString())}</p>`;
          // Regular sign-off — button always enabled
          if (confirmBtn) { confirmBtn.disabled = false; confirmBtn.style.opacity = '1'; }
          showModal('modal-signoff');
        }
      }

      if (action === 'reverse') {
        STATE.pendingReversal = { assignmentId: id, role };
        const desc = document.getElementById('reversal-desc');
        if (desc) desc.textContent = `You are reversing the ${role} sign-off. This action will be logged.`;
        const reasonEl = document.getElementById('reversal-reason');
        if (reasonEl) reasonEl.value = '';
        showModal('modal-reversal');
      }

      if (action === 'rc-resolve') {
        resolveReviewComment(id);
      }

      if (action === 'rc-reply') {
        openRCReplyInput(id);
      }

      if (action === 'submit-rc-reply') {
        submitRCReply(id);
      }

      if (action === 'cancel-rc-reply') {
        el.closest('.rc-reply-form')?.remove();
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

      if (action === 'edit-wd') {
        openEditWDModal(id);
      }

      if (action === 'edit-doc-link') {
        openEditDocLinkModal(id, el.dataset.url);
      }

      if (action === 'nudge-preparer') {
        const assignment = STATE.assignments.find(a => a._id === id);
        if (!assignment) return;
        const now = new Date().toISOString();
        // Optimistically update the card so the button disables immediately
        patchAssignment(id, { NudgeSent: now });
        updateListItem(CONFIG.lists.quarterlyAssignments, id, { NudgeSent: now })
          .then(() => {
            showToast(`👋 Nudge sent to ${assignment.Preparer?.split('@')[0] || 'preparer'}`, 'success');
            refreshCurrentView();
          })
          .catch(err => {
            // Revert on failure
            patchAssignment(id, { NudgeSent: assignment.NudgeSent || null });
            showToast(`Failed to send nudge — ${classifyGraphError(err)}`, 'error');
          });
      }

      if (action === 'skip-task' || action === 'unskip-task') {
        const isSkipping = action === 'skip-task';
        const assignment = STATE.assignments.find(a => a._id === id);
        if (!assignment) return;
        STATE._pendingSkip = { id, isSkipping, title: assignment.Title };
        const modalTitle = document.getElementById('modal-skip-title');
        const modalDesc  = document.getElementById('modal-skip-desc');
        if (isSkipping) {
          if (modalTitle) modalTitle.textContent = 'Skip this task?';
          const confirmBtn = document.getElementById('btn-skip-task-confirm');
          if (confirmBtn) { confirmBtn.textContent = 'Skip task'; confirmBtn.className = 'btn-danger'; }
          if (modalDesc) modalDesc.innerHTML = `
            <strong>${escapeHtml(assignment.Title)}</strong> will be removed from:
            <ul style="margin:8px 0 0 18px;font-size:12px">
              <li>My Tasks and All Tasks views</li>
              <li>Overdue counts and Dashboard</li>
              <li>Matrix (shown as N/A)</li>
              <li>Future rollforwards — unless you restore it first</li>
            </ul>
            <p style="margin-top:10px;font-size:12px;color:var(--slate)">You can restore it from the Show Skipped Tasks toggle in All Tasks.</p>`;
        } else {
          if (modalTitle) modalTitle.textContent = 'Restore this task?';
          const confirmBtn = document.getElementById('btn-skip-task-confirm');
          if (confirmBtn) { confirmBtn.textContent = 'Restore task'; confirmBtn.className = 'btn-primary'; }
          if (modalDesc) modalDesc.innerHTML = `<strong>${escapeHtml(assignment.Title)}</strong> will reappear in all views and be included in the next rollforward.`;
        }
        showModal('modal-skip-task');
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
function renderRCTagPicker() {
  // Renders a multi-select tag picker into the rc-tag-users div.
  // Called each time the RC modal opens so the user list is always current.
  const container = document.getElementById('rc-tag-users');
  if (!container) return;
  const currentEmail = STATE.currentUser?.Email;
  const users = STATE.users.filter(u => u.IsActive !== false && u.Email !== currentEmail);
  if (!users.length) {
    container.innerHTML = '<span style="font-size:11px;color:var(--slate)">No other users to tag</span>';
    return;
  }
  container.innerHTML = users.map(u =>
    `<label class="tag-option" style="display:inline-flex;align-items:center;gap:4px;margin:2px 4px 2px 0;cursor:pointer">
      <input type="checkbox" class="rc-tag-checkbox" value="${escapeHtml(u.Email)}" style="margin:0;cursor:pointer">
      ${renderBadge(u.Email)}
    </label>`
  ).join('');
}

function getTaggedUsers() {
  // Returns semicolon-separated emails of checked tag checkboxes.
  return [...document.querySelectorAll('.rc-tag-checkbox:checked')]
    .map(cb => cb.value).join(';');
}

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
      Status:              RC_STATUS.OPEN,
      TaggedUsers:         getTaggedUsers() || null,
    });
    STATE.reviewComments.push({ ...created.fields, _id: created.id });
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
      Status:         RC_STATUS.RESOLVED,
      ResolvedBy:     STATE.currentUser.Email,
      ResolvedDate:   now,
      ResolutionNote: note,
    });
    rc.Status = RC_STATUS.RESOLVED;
    rc.ResolvedBy = STATE.currentUser.Email;
    rc.ResolvedDate = now;
    rc.ResolutionNote = note;
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

    // Generate workday dates — skip weekends by default.
    // Weekends can be added manually via the Edit button on any calendar row
    // if a specific close requires weekend work.
    let current = new Date(startDate + 'T12:00:00');
    const created = [];

    for (let wd = 1; wd <= maxWD; wd++) {
      // Skip weekend days — advance until we land on a weekday
      let currentET = new Date(current.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
      while (currentET.getDay() === 0 || currentET.getDay() === 6) {
        current = new Date(current.getTime() + 86400000);
        currentET = new Date(current.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
      }

      const dateStr   = `${currentET.getFullYear()}-${String(currentET.getMonth()+1).padStart(2,'0')}-${String(currentET.getDate()).padStart(2,'0')}`;

      await createListItem(CONFIG.lists.closeCalendar, {
        Title:         `${quarter}-WD${wd}`,
        Quarter:       quarter,
        WorkdayNumber: wd,
        ActualDate:    dateStr,
        IsWeekend:     false,
      });
      created.push({ WorkdayNumber: wd, ActualDate: dateStr, IsWeekend: false,
                     MilestoneLabel: null, MilestoneType: null, Quarter: quarter });

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

    await 
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
  const weekendEl    = document.getElementById('edit-cal-weekend');
  if (dateEl)      dateEl.value      = row.ActualDate || '';
  if (weekendEl)   weekendEl.checked = !!row.IsWeekend;

  // Auto-update the weekend checkbox when the date changes
  if (dateEl && weekendEl) {
    const updateWeekend = () => {
      if (!dateEl.value) return;
      const d = new Date(dateEl.value + 'T12:00:00');
      weekendEl.checked = d.getDay() === 0 || d.getDay() === 6;
    };
    // Remove any previous listener then attach fresh — modal reuses the same elements
    dateEl.removeEventListener('change', dateEl._weekendHandler);
    dateEl.removeEventListener('input', dateEl._weekendHandler);
    dateEl._weekendHandler = updateWeekend;
    dateEl.addEventListener('change', updateWeekend);
    dateEl.addEventListener('input', updateWeekend);
  }

  const titleEl = document.getElementById('modal-edit-calendar-title');
  if (titleEl) titleEl.textContent = `Edit WD${row.WorkdayNumber}`;
  showModal('modal-edit-calendar');
}

async function saveCalendarRowEdit() {
  const calRowId = STATE.pendingCalendarEdit;
  if (!calRowId) return;
  const row = STATE.calendar.find(c => c._id === calRowId);
  if (!row) return;

  // Warn if editing a calendar row from a past quarter
  if (!confirmIfPastQuarter(row.Quarter, `edit the calendar for ${row.Quarter}`)) {
    hideModal('modal-edit-calendar');
    STATE.pendingCalendarEdit = null;
    return;
  }

  const newDate      = document.getElementById('edit-cal-date')?.value;
  // Auto-detect weekend from the date itself — don't rely solely on the checkbox
  const newDateDay   = newDate ? new Date(newDate + 'T12:00:00').getDay() : -1;
  const newWeekend   = newDateDay === 0 || newDateDay === 6 ||
                       document.getElementById('edit-cal-weekend')?.checked || false;

  if (!newDate) { showToast('Date is required', 'error'); return; }

  const prevDate = row.ActualDate;
  const quarter  = row.Quarter || STATE.activeQuarter || STATE.workingQuarter;

  // Snapshot current values for rollback on failure.
  const snapshot = {
    ActualDate: row.ActualDate,
    IsWeekend:  row.IsWeekend,
  };

  // Optimistic update — apply immediately so the admin panel reflects the change.
  row.ActualDate = newDate;
  row.IsWeekend  = newWeekend;

  try {
    await updateListItem(CONFIG.lists.closeCalendar, calRowId, {
      ActualDate: newDate,
      IsWeekend:  newWeekend,
    });
    await hideModal('modal-edit-calendar');
    STATE.pendingCalendarEdit = null;

    // If the date changed, handle cascade or resequence.
    if (prevDate && newDate !== prevDate) {
      if (newWeekend) {
        // Weekend workday inserted — resequence all subsequent rows from the day after
        // the new date, skipping weekends unless they were intentionally set as weekend workdays.
        const subsequent = STATE.calendar.filter(
          c => c.Quarter === quarter && Number(c.WorkdayNumber) > Number(row.WorkdayNumber)
        );

        if (subsequent.length > 0) {
          showLoading('Resequencing subsequent workdays...');
          let cursor = new Date(newDate + 'T12:00:00');
          cursor = new Date(cursor.getTime() + 86400000); // start from day after new date

          try {
            for (const c of subsequent) {
              // If this row was intentionally a weekend, just advance by 1 day
              // Otherwise skip to the next weekday
              if (!c.IsWeekend) {
                let cursorET = new Date(cursor.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
                while (cursorET.getDay() === 0 || cursorET.getDay() === 6) {
                  cursor = new Date(cursor.getTime() + 86400000);
                  cursorET = new Date(cursor.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
                }
              }
              const cursorET = new Date(cursor.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
              const newDateStr = `${cursorET.getFullYear()}-${String(cursorET.getMonth()+1).padStart(2,'0')}-${String(cursorET.getDate()).padStart(2,'0')}`;

              await updateListItem(CONFIG.lists.closeCalendar, c._id, {
                ActualDate: newDateStr,
                IsWeekend:  c.IsWeekend, // preserve intentional weekend flags
              });
              c.ActualDate = newDateStr;
              cursor = new Date(cursor.getTime() + 86400000);
            }
            showToast(`✓ Resequenced ${subsequent.length} workdays after WD${row.WorkdayNumber}`, 'success');
          } catch (err) {
            showToast('Resequence failed — some rows may be unchanged', 'error');
            logError('Resequence failed:', err);
          } finally {
            hideLoading();
            renderAdminPanel('calendar');
          }
          return;
        }
      } else {
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
      } // end if shiftDays !== 0
      } // end else (regular weekday cascade)
    } // end if date changed

    showToast('✓ Calendar row updated', 'success');
    renderAdminPanel('calendar');
  } catch (err) {
    // Revert optimistic update so the calendar reflects actual SharePoint state.
    patchCalendarRow(calRowId, snapshot);
    showToast('Failed to update calendar row', 'error');
    logError('saveCalendarRowEdit failed:', err);
  }
}

// Applies preparer/reviewer to all staging items in the selected category.
// Shows a confirmation modal before executing — actual write is in executeBulkAssign.
async function confirmBulkAssign() {
  const category = document.getElementById('bulk-assign-category')?.value || 'all';
  const preparer = document.getElementById('bulk-assign-preparer')?.value || '';
  const reviewer  = document.getElementById('bulk-assign-reviewer')?.value  || '';

  if (!preparer && !reviewer) {
    showToast('Select a preparer or reviewer to assign', 'error'); return;
  }

  const targets = STATE._stagingItems.filter(i =>
    !i.IsSkipped && (category === 'all' || i.Category === category)
  );

  if (!targets.length) {
    showToast('No tasks match the selected category', 'error'); return;
  }

  STATE._pendingBulkAssign = { targets, preparer, reviewer, category };
  const modalTitle = document.getElementById('modal-skip-title');
  const modalDesc  = document.getElementById('modal-skip-desc');
  const confirmBtn = document.getElementById('btn-skip-task-confirm');
  if (modalTitle) modalTitle.textContent = 'Confirm bulk assignment';
  if (modalDesc)  modalDesc.innerHTML = `Apply to <strong>${targets.length} task${targets.length !== 1 ? 's' : ''}</strong>${category !== 'all' ? ' in <strong>' + escapeHtml(category) + '</strong>' : ''}?<br><span style="font-size:12px;color:var(--slate)">This will overwrite existing preparer/reviewer assignments.</span>`;
  if (confirmBtn) { confirmBtn.textContent = 'Apply'; confirmBtn.className = 'btn-primary'; }
  showModal('modal-skip-task');
}

// Executes the actual bulk assign writes after confirmation.
async function executeBulkAssign(targets, preparer, reviewer) {
  const statusEl = document.getElementById('bulk-assign-status');
  if (statusEl) statusEl.textContent = `Saving 0 of ${targets.length}…`;

  let done = 0;
  const errors = [];
  for (const item of targets) {
    const fields = {};
    if (preparer) fields.Preparer = preparer;
    if (reviewer === '__clear__') fields.Reviewer = '';
    else if (reviewer) fields.Reviewer = reviewer;

    try {
      await updateListItem(CONFIG.lists.quarterlyAssignments, item._id, fields);
      Object.assign(item, fields);
      done++;
      if (statusEl) statusEl.textContent = `Saved ${done} of ${targets.length}…`;
    } catch (err) {
      errors.push(item.Title);
      logError('Bulk assign failed for:', item.Title, err);
    }
  }

  if (statusEl) statusEl.textContent = '';
  if (errors.length) {
    showToast(`Applied to ${done} tasks. Failed: ${errors.slice(0,3).join(', ')}${errors.length > 3 ? ` +${errors.length-3} more` : ''}`, 'warning');
  } else {
    showToast(`✓ Applied to ${done} tasks`, 'success');
  }

  const gridContainer = document.getElementById('staging-grid-container');
  if (gridContainer) {
    gridContainer.outerHTML = renderStagingGrid();
    attachAdminEvents('rollforward');
  }
}


let _pendingMilestoneWD   = null;
let _pendingMilestoneDate = null;

function openAddMilestoneModal(wd, date) {
  _pendingMilestoneWD   = wd;
  _pendingMilestoneDate = date;
  const labelEl = document.getElementById('milestone-label');
  const typeEl  = document.getElementById('milestone-type');
  if (labelEl) labelEl.value = '';
  if (typeEl)  typeEl.value  = 'Standard';
  const titleEl = document.getElementById('modal-add-milestone-title');
  if (titleEl) titleEl.textContent = `Add Milestone — ${formatDateShort(date)}${wd ? ' (WD' + wd + ')' : ''}`;
  showModal('modal-add-milestone');
}

async function confirmAddMilestone() {
  const label = document.getElementById('milestone-label')?.value?.trim();
  const type  = document.getElementById('milestone-type')?.value || 'Standard';
  if (!label) { showToast('Milestone label is required', 'error'); return; }
  if (!_pendingMilestoneDate) { showToast('No date selected — please close and try again', 'error'); return; }
  const quarter = STATE.activeQuarter || STATE.workingQuarter;
  if (!quarter) { showToast('No active quarter', 'error'); return; }

  hideModal('modal-add-milestone');
  try {
    const milestoneFields = {
      Title:          `${quarter} | WD${_pendingMilestoneWD || '?'} | ${label}`,
      Quarter:        quarter,
      MilestoneDate:  _pendingMilestoneDate,
      MilestoneLabel: label,
      MilestoneType:  type,
    };
    if (_pendingMilestoneWD) milestoneFields.WorkdayNumber = Number(_pendingMilestoneWD);
    const created = await createListItem(CONFIG.lists.calendarMilestones, milestoneFields);
    STATE.milestones.push({ ...created.fields, _id: created.id });
    showToast(`✓ Milestone added`, 'success');
    renderAdminPanel('calendar');
  } catch (err) {
    showToast(`Failed to add milestone — ${classifyGraphError(err)}`, 'error');
    logError('confirmAddMilestone failed:', err);
  }
  _pendingMilestoneWD   = null;
  _pendingMilestoneDate = null;
}

async function deleteMilestone(milestoneId) {
  if (!milestoneId) return;
  if (!window.confirm('Remove this milestone?')) return;
  try {
    await graphRequest('DELETE',
      `/sites/${await getSiteId()}/lists/${CONFIG.lists.calendarMilestones}/items/${milestoneId}`
    );
    STATE.milestones = STATE.milestones.filter(m => m._id !== milestoneId);
    showToast('✓ Milestone removed', 'success');
    renderAdminPanel('calendar');
  } catch (err) {
    showToast(`Failed to remove milestone — ${classifyGraphError(err)}`, 'error');
    logError('deleteMilestone failed:', err);
  }
}

// Applies the pending cascade shift to all subsequent workday rows.
async function confirmCascade() {
  const { quarter, fromWD, shiftDays, subsequent } = STATE.pendingCascade || {};
  if (!subsequent?.length) return;

  // Warn if cascading dates in a past quarter
  if (!confirmIfPastQuarter(quarter, `cascade shift dates in ${quarter}`)) {
    hideModal('modal-cascade');
    STATE.pendingCascade = null;
    return;
  }

  hideModal('modal-cascade');
  STATE.pendingCascade = null;

  showLoading(`Cascading ${Math.abs(shiftDays)}-day shift to ${subsequent.length} workdays...`);
  let updated = 0;
  try {
    for (const c of subsequent) {
      // Shift the date in ET to stay consistent with all other date handling.
      const oldDate  = new Date(c.ActualDate + 'T12:00:00');
      let shifted    = new Date(oldDate.getTime() + shiftDays * 86400000);
      let shiftedET  = new Date(shifted.toLocaleString('en-US', { timeZone: CONFIG.timezone }));

      // If the original row was not a weekend but the shifted date lands on one,
      // advance to the next Monday so weekends stay skipped.
      if (!c.IsWeekend && (shiftedET.getDay() === 0 || shiftedET.getDay() === 6)) {
        while (shiftedET.getDay() === 0 || shiftedET.getDay() === 6) {
          shifted   = new Date(shifted.getTime() + 86400000);
          shiftedET = new Date(shifted.toLocaleString('en-US', { timeZone: CONFIG.timezone }));
        }
      }

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
// EDIT WORKDAY DUE DATE (admin only, live quarter)
// ============================================================
function openEditWDModal(assignmentId) {
  const assignment = STATE.assignments.find(a => a._id === assignmentId);
  if (!assignment) return;

  STATE.pendingWDEdit = { assignmentId };

  const titleEl = document.getElementById('modal-edit-wd-title');
  const bodyEl  = document.getElementById('modal-edit-wd-body');

  if (titleEl) titleEl.textContent = 'Edit Due Dates';
  if (bodyEl) bodyEl.innerHTML = `
    <p style="font-size:13px;font-weight:500;margin-bottom:4px">${escapeHtml(assignment.Title || '')}</p>
    <p style="font-size:11px;color:var(--slate);margin-bottom:12px">${escapeHtml(assignment.Category || '')} · ${escapeHtml(assignment.SignOffMode || '')}</p>
    <div style="display:flex;gap:12px;flex-wrap:wrap">
      <div class="setup-field" style="flex:1">
        <label class="field-label" for="edit-wd-prep">Preparer Workday</label>
        <input type="number" id="edit-wd-prep" class="field-input" min="1" max="35"
          value="${assignment.PreparerWorkday || ''}" placeholder="WD number" />
      </div>
      ${assignment.SignOffMode !== SIGN_OFF_MODE.PREPARER_ONLY ? `
      <div class="setup-field" style="flex:1">
        <label class="field-label" for="edit-wd-rev">Reviewer Workday</label>
        <input type="number" id="edit-wd-rev" class="field-input" min="1" max="35"
          value="${assignment.ReviewerWorkday || ''}" placeholder="WD number" />
      </div>` : ''}
    </div>
    <p style="font-size:11px;color:var(--slate);margin-top:8px">Changes apply to this quarter only. The template workday is unchanged.</p>
    <p id="edit-wd-error" class="modal-desc hidden" style="color:var(--red);margin-top:4px"></p>`;

  showModal('modal-edit-wd');
}

async function confirmEditWD() {
  const { assignmentId } = STATE.pendingWDEdit || {};
  if (!assignmentId) return;

  const assignment = STATE.assignments.find(a => a._id === assignmentId);
  if (!assignment) return;

  const prepWD = Number(document.getElementById('edit-wd-prep')?.value);
  const revWD  = document.getElementById('edit-wd-rev')
    ? Number(document.getElementById('edit-wd-rev').value) : null;
  const errEl  = document.getElementById('edit-wd-error');

  if (!prepWD || prepWD < 1 || prepWD > 35) {
    if (errEl) { errEl.textContent = 'Preparer workday must be between 1 and 35.'; errEl.classList.remove('hidden'); }
    return;
  }
  if (revWD !== null && assignment.SignOffMode !== SIGN_OFF_MODE.PREPARER_ONLY && (revWD < 1 || revWD > 35)) {
    if (errEl) { errEl.textContent = 'Reviewer workday must be between 1 and 35.'; errEl.classList.remove('hidden'); }
    return;
  }

  hideModal('modal-edit-wd');

  const fields = { PreparerWorkday: prepWD };
  if (revWD && assignment.SignOffMode !== SIGN_OFF_MODE.PREPARER_ONLY) fields.ReviewerWorkday = revWD;

  // Snapshot for rollback
  const snapshot = { PreparerWorkday: assignment.PreparerWorkday, ReviewerWorkday: assignment.ReviewerWorkday };

  // Optimistic update
  Object.assign(assignment, fields);
  openTaskPanel(assignmentId);
  refreshCurrentView();

  try {
    await updateListItem(CONFIG.lists.quarterlyAssignments, assignmentId, fields);
    showToast('✓ Due dates updated', 'success');
  } catch (err) {
    Object.assign(assignment, snapshot);
    openTaskPanel(assignmentId);
    refreshCurrentView();
    showToast(`Failed to update due dates — ${classifyGraphError(err)}`, 'error');
    logError('confirmEditWD failed:', err);
  }
  STATE.pendingWDEdit = null;
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
  const confirmBtn = document.getElementById('btn-signoff-confirm');

  const assignedEmail = role === 'preparer' ? assignment.Preparer : assignment.Reviewer;
  const et = formatDateET(new Date().toISOString());

  if (titleEl) titleEl.textContent = `Sign off on behalf`;
  if (bodyEl) bodyEl.innerHTML = `
    <p style="font-size:13px;font-weight:500;margin-bottom:6px">${escapeHtml(assignment.Title || '')}</p>
    <p style="font-size:12px;color:var(--slate);margin-bottom:4px">
      Assigned ${role}: ${renderBadge(assignedEmail)}
    </p>
    <p style="font-size:12px;color:var(--slate);margin-bottom:12px">
      Signing as: ${renderBadge(STATE.currentUser?.Email)} · ${et}
    </p>
    <div style="background:#FFF8E6;border:1px solid #F5C842;border-radius:6px;padding:10px 12px;margin-bottom:12px">
      <p style="font-size:11px;color:#7A5200;font-weight:600;margin-bottom:6px">
        ⚠ This action will be recorded in the audit log as signed on behalf of the assigned ${role}. It cannot be undone without a reversal.
      </p>
      <p style="font-size:11px;color:#7A5200">
        Type <strong>ON BEHALF</strong> below to confirm:
      </p>
    </div>
    <input type="text" id="signoff-behalf-confirm-input" class="field-input"
      placeholder="Type ON BEHALF to confirm" autocomplete="off"
      style="font-size:13px;letter-spacing:1px" />`;

  // Disable confirm button until "ON BEHALF" is typed
  if (confirmBtn) {
    confirmBtn.disabled = true;
    confirmBtn.style.opacity = '0.5';
  }

  // Wire input to gate the button
  setTimeout(() => {
    const input = document.getElementById('signoff-behalf-confirm-input');
    if (input && confirmBtn) {
      input.addEventListener('input', () => {
        const valid = input.value.trim().toUpperCase() === 'ON BEHALF';
        confirmBtn.disabled = !valid;
        confirmBtn.style.opacity = valid ? '1' : '0.5';
      });
      input.focus();
    }
  }, 100);

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
      quarter:       assignment.Quarter || STATE.activeQuarter,
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
    showToast(`Reassignment failed — ${classifyGraphError(err)}`, 'error');
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
  STATE._editUserEmoji = user.Emoji || null;
  STATE._editUserColor = user.Color || CONFIG.colorOptions[0].hex;

  const nameEl   = document.getElementById('edit-user-name');
  const roleEl   = document.getElementById('edit-user-role');
  const activeEl = document.getElementById('edit-user-active');
  const customEl = document.getElementById('edit-user-emoji-custom');
  if (nameEl)   nameEl.textContent  = `${user.Emoji || ''} ${user.Title || email}`;
  if (roleEl)   roleEl.value        = user.Role || ROLE.TEAM_MEMBER;
  if (activeEl) activeEl.checked    = user.IsActive !== false;
  if (customEl) customEl.value      = '';

  renderEmojiPicker('edit-user-emoji-grid', user.Emoji, (emoji) => {
    STATE._editUserEmoji = emoji;
    const c = document.getElementById('edit-user-emoji-custom');
    if (c) c.value = '';
  });
  renderColorPicker('edit-user-color-grid', user.Color || CONFIG.colorOptions[0].hex, (color) => {
    STATE._editUserColor = color;
  });

  // Custom emoji input
  if (customEl) {
    customEl.oninput = () => {
      const val = customEl.value.trim();
      if (val) STATE._editUserEmoji = val;
    };
  }

  showModal('modal-edit-user');
}

async function saveUserRoleEdit() {
  const email = STATE.pendingUserEdit;
  if (!email) return;
  const user = STATE.users.find(u => u.Email === email);
  if (!user) return;

  const newRole   = document.getElementById('edit-user-role')?.value || ROLE.TEAM_MEMBER;
  const newActive = document.getElementById('edit-user-active')?.checked !== false;
  const newEmoji  = STATE._editUserEmoji || user.Emoji || null;
  const newColor  = STATE._editUserColor || user.Color || null;

  try {
    await updateListItem(CONFIG.lists.users, user._id, {
      Role:     newRole,
      IsActive: newActive,
      Emoji:    newEmoji,
      Color:    newColor,
    });
    const prevRole = user.Role;
    patchUser(user._id, { Role: newRole, IsActive: newActive, Emoji: newEmoji, Color: newColor });
    // Update current user's role flags if they edited themselves
    if (email === STATE.currentUser?.Email) {
      STATE.isAdmin        = newRole === ROLE.ADMIN;
      STATE.isFinalReviewer = newRole === ROLE.FINAL_REVIEWER || STATE.isAdmin;
      STATE.isReadOnly      = newRole === ROLE.READ_ONLY;
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
  const role  = roleEl?.value || ROLE.TEAM_MEMBER;
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
  // Derive quarters from assignments (always loaded) + active quarter.
  // Do not rely on STATE._auditEntries which may be empty if the audit log panel hasn't been opened.
  const fromAssignments = [...new Set(STATE.assignments.map(a => a.Quarter).filter(Boolean))];
  const quarters = [...new Set([
    ...fromAssignments,
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
        `fields/Quarter eq '${quarter}' and fields/IsStaging eq 0`);
      assignments = items.map(i => ({ ...i.fields, _id: i.id }));
    }

    // Fetch the calendar for the export quarter if it differs from the active quarter.
    // STATE.calendar only holds the active quarter — historical exports need their own calendar.
    let exportCalendar = STATE.calendar.filter(c => c.Quarter === quarter);
    if (!exportCalendar.length) {
      const calItems = await getListItems(
        CONFIG.lists.closeCalendar, `fields/Quarter eq '${quarter}'`);
      exportCalendar = calItems.map(i => {
        const row = { ...i.fields, _id: i.id };
        if (row.ActualDate?.includes('T')) row.ActualDate = row.ActualDate.split('T')[0];
        return row;
      });
    }

    function getSignOffWD(isoDate) {
      if (!isoDate) return '';
      const dateStr = isoDate.substring(0, 10);
      const match = exportCalendar.find(c => c.ActualDate === dateStr);
      return match ? Number(match.WorkdayNumber) : '';
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
      if (a.ReviewerSignOff && a.SignOffMode !== SIGN_OFF_MODE.PREPARER_ONLY) {
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
    const unsignedRows = [['Quarter','Task Name','Category','Sign-Off Type','Assigned To','Due WD','Status','Skipped']];
    assignments.forEach(a => {
      const skipped = a.IsSkipped ? 'Yes' : 'No';
      if (!a.PreparerSignOff) {
        unsignedRows.push([quarter, a.Title, a.Category, 'Preparer', a.Preparer || 'Unassigned', a.PreparerWorkday || '', a.IsSkipped ? 'Skipped' : (a.Status || ''), skipped]);
      }
      if (!a.ReviewerSignOff && a.SignOffMode !== SIGN_OFF_MODE.PREPARER_ONLY && !a.IsSkipped) {
        unsignedRows.push([quarter, a.Title, a.Category, 'Reviewer', a.Reviewer || 'Unassigned', a.ReviewerWorkday || '', a.Status || '', skipped]);
      }
    });

    // ── 3. Reversals ─────────────────────────────────────────
    // Always fetch audit entries fresh for the export quarter — STATE._auditEntries
    // is only populated when the Audit Log panel is opened and may be empty or stale.
    let auditQ = STATE._auditEntries.filter(e => e.Quarter === quarter);
    if (!auditQ.length) {
      const freshAudit = await getListItems(
        CONFIG.lists.auditLog, `fields/Quarter eq '${quarter}'`);
      auditQ = freshAudit.map(i => ({ ...i.fields, _id: i.id }));
    }
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

    // ── 5b. Final Review sign-offs ───────────────────────────
    const finalReviewRows = [['Quarter','Date ET','Matrix Item','Signed By','Status','Previous Status']];
    auditQ.filter(e => e.ActionType === 'FinalReview').forEach(e => {
      finalReviewRows.push([quarter, formatDateET(e.ActionDate), e.TaskName, e.ActionBy, e.NewValue || '', e.PreviousValue || '']);
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
    const totalRCOpen    = STATE.reviewComments.filter(rc => rc.Quarter === quarter && rc.Status === RC_STATUS.OPEN).length;

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
    addSheet('Final Review',     finalReviewRows);
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
    // Use getReadQuarter so admins viewing a historical quarter export that quarter,
    // not the currently active one.
    const exportQ = getReadQuarter();
    const items = await getListItems(CONFIG.lists.auditLog,
      exportQ ? `fields/Quarter eq '${exportQ}'` : ''
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
    downloadCSV(rows, `Folio-AuditLog-${exportQ || 'all'}.csv`);
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

  const replyTagOpts = STATE.users
    .filter(u => u.IsActive !== false && u.Email !== STATE.currentUser?.Email)
    .map(u => `<label class="tag-option" style="display:inline-flex;align-items:center;gap:3px;margin:2px;cursor:pointer;">`
      + `<input type="checkbox" class="reply-tag-checkbox-${rcId}" value="${u.Email}" style="margin:0;cursor:pointer">`
      + renderBadge(u.Email) + `</label>`).join('');

  actionsDiv.insertAdjacentHTML('beforeend', `
    <div class="rc-reply-form" style="margin-top:8px;width:100%">
      <textarea id="reply-text-${rcId}" class="field-textarea" rows="2"
        placeholder="Type your reply..." style="width:100%;font-size:12px;box-sizing:border-box"></textarea>
      ${replyTagOpts ? `<div style="margin:4px 0 6px;font-size:11px;color:var(--slate)">Tag teammates (optional):</div>
      <div style="display:flex;flex-wrap:wrap;gap:2px;margin-bottom:6px">${replyTagOpts}</div>` : ''}
      <div style="display:flex;gap:6px">
        <button class="btn-primary btn-sm" data-action="submit-rc-reply" data-id="${rcId}">Post</button>
        <button class="btn-secondary btn-sm" data-action="cancel-rc-reply">Cancel</button>
      </div>
    </div>`);
  document.getElementById(`reply-text-${rcId}`)?.focus();
}

async function submitRCReply(rcId) {
  const text = document.getElementById(`reply-text-${rcId}`)?.value?.trim();
  if (!text) { showToast('Please enter a reply', 'error'); return; }

  const now = new Date().toISOString();
  try {
    const replyTagged = [...document.querySelectorAll(`.reply-tag-checkbox-${rcId}:checked`)]
      .map(cb => cb.value).join(';');
    const created = await createListItem(CONFIG.lists.reviewCommentReplies, {
      Title:                 `Reply to RC ${rcId}`,
      ReviewCommentLookupId: rcId,
      ReplyText:             text,
      CreatedByEmail:        STATE.currentUser.Email,
      CreatedDate:           now,
      TaggedUsers:           replyTagged || null,
      Quarter:               STATE.activeQuarter || '',  // Stored for future server-side filtering
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
    'edit-tpl-filingtype':    t?.FilingType || FILING.BOTH,
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

  const signOffMode = document.getElementById('edit-tpl-signoffmode')?.value;
  const revWDVal    = document.getElementById('edit-tpl-revwd')?.value;
  if (signOffMode === SIGN_OFF_MODE.SEQUENTIAL && !revWDVal) {
    showToast('Reviewer Workday is required for Sequential tasks', 'error'); return;
  }

  const prepWD10K = document.getElementById('edit-tpl-prepwd-10k')?.value;
  const revWD10K  = document.getElementById('edit-tpl-revwd-10k')?.value;

  const updates = {
    Title:               name,
    Category:            document.getElementById('edit-tpl-category')?.value || 'Other',
    FilingType:          document.getElementById('edit-tpl-filingtype')?.value || FILING.BOTH,
    SignOffMode:         document.getElementById('edit-tpl-signoffmode')?.value || 'Sequential',
    PreparerWorkday:     Number(document.getElementById('edit-tpl-prepwd')?.value) || 1,
    ReviewerWorkday:     document.getElementById('edit-tpl-revwd')?.value
      ? Number(document.getElementById('edit-tpl-revwd').value) : null,
    PreparerWorkday10K:  prepWD10K ? Number(prepWD10K) : null,
    ReviewerWorkday10K:  revWD10K  ? Number(revWD10K)  : null,
    // IsActive omitted from edit payload — saving a retired template must not un-retire it.
    // IsActive is only set to true on create (new templates start active by default).
  };

  try {
    if (t) {
      // Edit mode — update existing template
      await updateListItem(CONFIG.lists.taskTemplates, templateId, updates);
      Object.assign(t, updates);
      showToast('✓ Template saved', 'success');
    } else {
      // Create mode — new template
      const created = await createListItem(CONFIG.lists.taskTemplates, {
        ...updates,
        TaskName: name, // TaskName mirrors Title for the app's display logic
        IsActive: true, // New templates always start active
      });
      STATE.templates.push({ ...created.fields, _id: created.id });
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
  const quarter = normalizeQuarter(input?.value);

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
  if (!quarter) { showToast('No staging quarter set', 'error'); return; }

  // Check if existing staging items have been customized (preparer populated)
  // and warn the admin they will be overwritten.
  const existingCustomized = STATE._stagingItems.filter(i => i.Preparer);
  const hasCustomizations  = existingCustomized.length > 0;

  STATE.pendingRollforward = quarter;
  const rfDetail = document.getElementById('rollforward-confirm-detail');
  if (rfDetail) {
    const baseMsg = `This will create ~${STATE.templates.length} staging assignments for ${quarter} copied from templates. `;
    const overwriteMsg = hasCustomizations
      ? `⚠️ Warning: ${existingCustomized.length} staging assignment(s) already have customized preparers/reviewers that will be lost. `
      : `Existing staging assignments for ${quarter} will be replaced. `;
    rfDetail.textContent = baseMsg + overwriteMsg + `You can review and adjust before activating.`;
    rfDetail.style.color = hasCustomizations ? 'var(--red-dark)' : '';
  }
  showModal('modal-rollforward-confirm');
}

function confirmActivation() {
  if (!STATE.workingQuarter) { showToast('No staging quarter to activate', 'error'); return; }
  STATE.pendingActivation = STATE.workingQuarter;
  const titleEl = document.getElementById('activate-modal-title');
  const descEl  = document.getElementById('activate-modal-desc');
  if (titleEl) titleEl.textContent = `Activate ${STATE.workingQuarter}?`;
  if (descEl)  descEl.textContent  = `This will immediately make ${STATE.workingQuarter} the live quarter, visible to all ${STATE.users.length} team members. Make sure the staging grid is complete before activating.`;
  showModal('modal-activate');
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
      `fields/Quarter eq '${quarter}' and fields/IsStaging eq 1`
    );
    for (const item of existing) {
      await graphRequest('DELETE',
        `/sites/${await getSiteId()}/lists/${CONFIG.lists.quarterlyAssignments}/items/${item.id}`
      );
    }

    // Determine filing type for this quarter
    const filingType = isQuarterQ4(quarter) ? FILING.K : FILING.Q;
    const eligible = STATE.templates.filter(t =>
      t.IsActive !== false &&
      (t.FilingType === filingType || t.FilingType === FILING.BOTH)
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
      const isQ4 = filingType === FILING.K;
      const prepWD = (isQ4 && template.PreparerWorkday10K)
        ? template.PreparerWorkday10K
        : template.PreparerWorkday || null;
      const revWD  = (isQ4 && template.ReviewerWorkday10K)
        ? template.ReviewerWorkday10K
        : template.ReviewerWorkday || null;

      // Warn if a Sequential template has no reviewer workday — the assignment would
      // never show as overdue for the reviewer. Log it but still create the assignment.
      if (template.SignOffMode === SIGN_OFF_MODE.SEQUENTIAL && !revWD) {
        logError(`Rollforward warning: Sequential template '${template.Title}' has no ReviewerWorkday — reviewer step will never show as overdue.`);
      }

      await createListItem(CONFIG.lists.quarterlyAssignments, {
        Title:        `${quarter} - ${template.TaskName || template.Title || ''}`,
        Quarter:      quarter,
        TaskTemplateLookupId: template._id,
        Preparer:     prev?.Preparer || template.DefaultPreparer || null,
        Reviewer:     prev?.Reviewer || template.DefaultReviewer || null,
        SignOffMode:  template.SignOffMode || SIGN_OFF_MODE.SEQUENTIAL,
        Category:     template.Category || '',
        MatrixItem:   template.MatrixItem || null,
        MatrixCheckpoint: template.MatrixCheckpoint || null,
        PreparerWorkday:  prepWD,
        ReviewerWorkday:  revWD,
        HasDocumentLink:  template.HasDocumentLink || false,
        PreparerSignOff:  false,
        ReviewerSignOff:  false,
        Status:       STATUS.NOT_STARTED,
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

    // Warn if any assignments reference inactive users BEFORE clearing _stagingItems
    const activeEmails = new Set(STATE.users.filter(u => u.IsActive !== false && u.IsActive !== 0).map(u => u.Email));
    const inactiveCount = eligible.filter(t => {
      const prev = prevMap[t._id];
      const preparer = prev?.Preparer || t.DefaultPreparer;
      const reviewer  = prev?.Reviewer  || t.DefaultReviewer;
      return (preparer && !activeEmails.has(preparer)) || (reviewer && !activeEmails.has(reviewer));
    }).length;

    STATE._stagingItems   = [];
    STATE._stagingLoading = false;
    hideModal('modal-rollforward-confirm');
    showToast(`✓ Rolled forward ${created} tasks to ${quarter}`, 'success');

    if (inactiveCount > 0) {
      showToast(`⚠ ${inactiveCount} task(s) may reference inactive users — review the staging grid before activating`, 'warning');
    }

    renderAdminPanel('rollforward');
  } catch (err) {
    if (created > 0) {
      showToast(
        `Rollforward partially completed — ${created} of ${eligible?.length || '?'} tasks created. ` +
        `Review and manually clean up QuarterlyAssignments in SharePoint before retrying.`,
        'error'
      );
    } else {
      showToast(`Rollforward failed — ${classifyGraphError(err)}`, 'error');
    }
    logError('confirmRollforward failed:', err);
  } finally {
    hideLoading();
  }
}

// ============================================================
// QUARTER ACTIVATION
// ============================================================
async function activateQuarter(quarter) {
  // Guard: require a close calendar before activation.
  // Without a calendar, all tasks show no due dates and overdue detection is broken.
  const calCheck = await getListItems(
    CONFIG.lists.closeCalendar, `fields/Quarter eq '${quarter}'`);
  if (!calCheck.length) {
    showToast(
      `Cannot activate — no Close Calendar exists for ${quarter}. ` +
      'Go to Admin → Close Calendar → Setup Calendar first.',
      'error'
    );
    return;
  }

  showLoading(`Activating ${quarter}...`);
  // Declared outside try so catch can safely reference it for the error message.
  let stagingItems = [];
  try {
    stagingItems = await getListItems(
      CONFIG.lists.quarterlyAssignments,
      `fields/Quarter eq '${quarter}' and fields/IsStaging eq 1`
    );
    for (const item of stagingItems) {
      await updateListItem(CONFIG.lists.quarterlyAssignments, item.id, { IsStaging: false });
    }
    await setAppSettings({ ActiveQuarter: quarter, WorkingQuarter: '' });
    STATE.activeQuarter = quarter;
    STATE.workingQuarter = '';
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
          row.push(ms?.[fm.status] || STATUS.NOT_STARTED);
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
  // Export as Excel using SheetJS
  if (typeof XLSX !== 'undefined') {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(rows);
    // Style header row
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c: C });
      if (!ws[addr]) continue;
      ws[addr].s = { font: { bold: true }, fill: { fgColor: { rgb: '0A1264' } } };
    }
    XLSX.utils.book_append_sheet(wb, ws, 'Matrix');
    XLSX.writeFile(wb, `Folio-Matrix-${quarter}.xlsx`);
  } else {
    downloadCSV(rows, `Folio-Matrix-${quarter}.csv`);
  }
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
  document.title = `Folio${STATE.activeQuarter ? ' — ' + STATE.activeQuarter : ''}`;

  // Show correct nav items based on role
  document.querySelectorAll('.nav-matrix-link').forEach(el => {
    el.classList.remove('hidden');
  });
  document.querySelectorAll('.nav-admin-link').forEach(el => {
    el.classList.toggle('hidden', !STATE.isAdmin);
  });
  // Hide "New Comment" button for non-reviewers — reviewers and admins only
  const newRCBtn = document.getElementById('btn-new-rc');
  // Show New Comment button for FinalReviewers, Admins, and TeamMembers (they may be assigned reviewers)
  // ReadOnly users never see it
  if (newRCBtn) newRCBtn.classList.toggle('hidden', STATE.isReadOnly);
  const suggestBtn = document.getElementById('btn-suggest-change');
  if (suggestBtn) suggestBtn.classList.toggle('hidden', STATE.isReadOnly);
  const signoffAllBtn = document.getElementById('btn-signoff-all');
  if (signoffAllBtn && STATE.isReadOnly) signoffAllBtn.style.display = 'none';
  // Add Task button — admins only, only when a quarter is active
  const addTaskBtn = document.getElementById('btn-add-task-midquarter');
  if (addTaskBtn) addTaskBtn.style.display = STATE.isAdmin && STATE.activeQuarter ? '' : 'none';

  updateNavAvatar();

  // Load all data
  showLoading('Loading your tasks...');
  try {
    await loadTemplates();
    await loadUsers();  // Always load users — needed for staging grid and badges regardless of active quarter
    // Load milestones for working quarter if no active quarter yet
    if (!STATE.activeQuarter && STATE.workingQuarter) {
      await loadMilestones(STATE.workingQuarter);
    }
    if (STATE.activeQuarter) {
      await loadAllData();
    }
  } catch (err) {
    logError('Initial data load failed:', err);
    showStaleBanner(true);
  }
  hideLoading();

  updateWDIndicator(); updateContextRibbon();

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
      const errMsg = String(err.message || '');
      if (errMsg.startsWith('ACCESS_DENIED:')) {
        // Show a dedicated access denied screen rather than the sign-in screen
        const accessEl = document.getElementById('access-denied-msg');
        if (accessEl) accessEl.textContent = errMsg.replace('ACCESS_DENIED: ', '');
        showScreen('screen-access-denied');
      } else {
        const classified = classifyGraphError(err);
        const errEl = document.getElementById('signin-error');
        if (errEl) { errEl.textContent = classified; errEl.classList.remove('hidden'); }
        showScreen('screen-signin');
      }
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
