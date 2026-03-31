/* ============================================================
   FINANCIAL REPORTING TRACKER — app.js
   Hosted on:  GitHub Pages
   Database:   SharePoint Lists via Microsoft Graph API
   Auth:       Microsoft SSO (User.Read + Sites.ReadWrite.All)
   Notify:    Writes to FT_Notifications list — PA flow sends emails
   Admin:      IsAdmin = Yes in FT_Users

   Data operations call the Microsoft Graph API directly using
   a Bearer token obtained via MSAL. No Power Automate flows
   needed — the app reads and writes SharePoint lists directly.
   ============================================================ */

// ── CONFIG ────────────────────────────────────────────────────
const CONFIG = {
  // Azure AD — requires User.Read + Sites.ReadWrite.All
  clientId:  "bb00291f-d451-4e74-b8cf-10c334efb0ed",
  tenantId:  "1061a8b8-b1ee-4249-bb84-9a2cd2792fae",

  // SharePoint site URL — e.g. https://moodys.sharepoint.com/sites/FinancialReporting
  siteUrl:   "https://moodys.sharepoint.com/sites/finance_home_finrptg",
};
// ─────────────────────────────────────────────────────────────

// ── MSAL — User.Read + Sites.ReadWrite.All ────────────────────
const msalConfig = {
  auth: {
    clientId:  CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: true,  // avoids Edge tracking prevention
  },
  system: { allowNativeBroker: false }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const GRAPH_SCOPES = ["User.Read", "Sites.ReadWrite.All"];

// ── GRAPH TOKEN ───────────────────────────────────────────────
// Acquires a fresh (or cached) Bearer token for every Graph call.
async function getGraphToken() {
  const account = msalInstance.getAllAccounts()[0];
  if (!account) throw new Error("Not signed in — please sign in with Microsoft first.");
  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
    return result.accessToken;
  } catch(e) {
    // Silent acquisition failed (e.g. consent not yet granted) — fall back to popup
    const result = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    return result.accessToken;
  }
}

// ── GRAPH API HELPER ─────────────────────────────────────────
// Low-level call to any Graph endpoint with automatic token attachment.
// Handles 429 throttling: reads Retry-After header and retries once after
// the indicated delay. Graph's limit is 10k req/10min per app — a 5-person
// team won't hit it under normal use, but during bulk import it's possible.
async function callGraph(method, path, body) {
  if (!CONFIG.siteUrl || CONFIG.siteUrl.startsWith("REPLACE_")) {
    throw new Error("SharePoint site URL not configured — set CONFIG.siteUrl at the top of app.js.");
  }
  const token = await getGraphToken();
  const url   = path.startsWith("https://") ? path : `https://graph.microsoft.com/v1.0${path}`;
  const opts  = {
    method,
    headers: {
      "Authorization": `Bearer ${token}`,
      "Content-Type":  "application/json",
    },
  };
  if (body) opts.body = JSON.stringify(body);

  const res = await fetch(url, opts);

  // 429 Too Many Requests — Graph throttling. Wait Retry-After seconds then retry once.
  if (res.status === 429) {
    const retryAfter = parseInt(res.headers.get('Retry-After') || '10', 10);
    console.warn(`Graph throttled (429) — retrying after ${retryAfter}s`);
    await new Promise(r => setTimeout(r, retryAfter * 1000));
    const retry = await fetch(url, opts);
    if (!retry.ok) {
      const err = await retry.text();
      throw new Error(`Graph ${method} ${path} → ${retry.status} (after retry): ${err.slice(0, 200)}`);
    }
    if (retry.status === 204) return null;
    const retryText = await retry.text();
    return retryText ? JSON.parse(retryText) : null;
  }

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Graph ${method} ${path} → ${res.status}: ${err.slice(0, 200)}`);
  }
  if (res.status === 204) return null; // DELETE / no-content responses
  const text = await res.text();
  if (!text) return null;
  return JSON.parse(text);
}

// ── GRAPH SITE ID (resolved once per session) ────────────────
// Graph list operations require the internal site ID, which we
// resolve from the site URL once and cache for the session.
// Throws on failure — callers (getListItems etc.) propagate to loadAllData/refreshData.
let _graphSiteId = null;
async function getGraphSiteId() {
  if (_graphSiteId) return _graphSiteId;
  // Convert https://tenant.sharepoint.com/sites/SiteName
  //    into  graph.microsoft.com/v1.0/sites/tenant.sharepoint.com:/sites/SiteName
  const url   = new URL(CONFIG.siteUrl);
  const host  = url.hostname;                   // moodys.sharepoint.com
  const path  = url.pathname;                   // /sites/FinancialReporting
  const data  = await callGraph("GET", `/sites/${host}:${path}`);
  _graphSiteId = data.id;
  return _graphSiteId;
}

// ── LIST HELPERS (via Microsoft Graph) ───────────────────────
// Graph returns list items with fields nested under item.fields.
// All normaliser functions already expect a flat fields object,
// so we return item.fields merged with item.id (the SP item ID).
// All four helpers throw on failure — callers are responsible for
// catching errors (typically in try/catch inside the calling function).

async function getListItems(listName) {
  const siteId = await getGraphSiteId();
  const all    = [];
  // Graph paginates at 200 items — follow @odata.nextLink until done
  let url = `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items?expand=fields&$top=999`;
  while (url) {
    const data = await callGraph("GET", url);
    (data?.value || []).forEach(item => {
      all.push({ id: item.id, ID: item.id, ...item.fields });
    });
    url = data?.["@odata.nextLink"] || null;
  }
  return all;
}

async function createListItem(listName, fields) {
  const siteId  = await getGraphSiteId();
  const payload = { fields };
  const data    = await callGraph(
    "POST",
    `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items`,
    payload
  );
  // Graph always returns the created item with a numeric string id at data.id.
  // Throw rather than silently return '' — callers that use the ID for follow-up
  // writes (e.g. applyTemplate creating steps under a new task) would otherwise
  // write orphaned records with TaskId = '' that can never be retrieved.
  const id = data?.id || data?.fields?.id;
  if (!id) throw new Error(`createListItem(${listName}): Graph response missing item ID`);
  return { id };
}

async function updateListItem(listName, itemId, fields) {
  const siteId = await getGraphSiteId();
  await callGraph(
    "PATCH",
    `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items/${itemId}/fields`,
    fields
  );
}

async function deleteListItem(listName, itemId) {
  const siteId = await getGraphSiteId();
  await callGraph(
    "DELETE",
    `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items/${itemId}`
  );
}
// ── LIST NAMES ───────────────────────────────────────────────
const LISTS = {
  tasks:         "FT_Tasks",
  users:         "FT_Users",
  comments:      "FT_Comments",
  templates:     "FT_Templates",
  steps:         "FT_Steps",
  stepTemplates: "FT_StepTemplates",
  signOffs:      "FT_SignOffs",
  locks:         "FT_QuarterLocks",
  attachments:   "FT_Attachments",
  quarterDates:  "FT_QuarterDates",
  notifications: "FT_Notifications",  // watched by PA flow to send emails
};

// ── IN-MEMORY CACHE ───────────────────────────────────────────
let _users         = [];
let _tasks         = [];
let _templates     = [];
let _comments      = [];
let _steps         = [];
let _stepTemplates = [];
let _signOffs      = [];
let _locks         = [];
let _attachments   = [];
let _quarterDates  = [];
let _pollTimer         = null;
let _commentPollTimer  = null;

function getUserById(id)   { if(!id) return null; return _users.find(u => u._spId===id || u.id===id || u.ID===id) || null; }
function getUsers()        { return _users; }
function getTasks()        { return _tasks; }
// Returns tasks excluding fully-locked quarters — used in all views except Admin/Report/Exec
function getActiveTasks()  { return _tasks.filter(t => !isQuarterLocked(t.quarter, t.year)); }
function getTemplates()    { return _templates; }
function getStepsForTask(taskSpId) {
  return _steps.filter(s => s.taskId === taskSpId).sort((a,b) => (a.order||0)-(b.order||0));
}
function isQuarterLocked(quarter, year) {
  return _locks.some(l => l.quarter === quarter && l.year === parseInt(year));
}
function getStepTemplatesForTemplate(templateId) {
  return _stepTemplates.filter(s => s.templateId === templateId).sort((a,b) => (a.order||0)-(b.order||0));
}

// ── DATA NORMALISER ───────────────────────────────────────────
// SharePoint returns fields with capitalised names; normalise to
// the same shape the rest of the app expects.
function normaliseTask(f) {
  return {
    _spId:       f.id || f.ID || '',
    id:          f.TaskId  || f.id || f.ID || '',
    name:        f.Title   || f.TaskName || '',
    type:        f.TaskType || '',
    quarter:     f.Quarter  || '',
    year:        parseInt(f.Year) || new Date().getFullYear(),
    dueDate:     (f.DueDate || '').slice(0, 10),
    status:      f.Status   || 'Not Started',
    ownerId:     f.OwnerId  || '',
    reviewerId:  f.ReviewerId  || '',
    reviewer2Id: f.Reviewer2Id || '',
    description:        f.Description        || '',
    notes:              f.Notes              || '',   // rich notes, can carry forward
    applicability:      f.Applicability      || 'All Quarters',
    workdayNum:         f.WorkdayNum ? parseInt(f.WorkdayNum) : null,
    skipNextRollforward: f.SkipNextRollforward === 'Yes' || f.SkipNextRollforward === true,
    reassignRequested:  f.ReassignRequested === 'Yes' || f.ReassignRequested === true,
    reassignNote:       f.ReassignNote || '',
  };
}
function normaliseUser(f) {
  return {
    _spId:   f.id || f.ID || '',
    id:      f.UserId   || f.id || f.ID || '',
    name:    f.Title    || f.FullName || '',
    role:    f.JobRole  || '',
    email:   (f.Email   || '').toLowerCase().trim(),
    isAdmin: f.IsAdmin === true || f.IsAdmin === 'Yes',
  };
}
function normaliseTemplate(f) {
  return {
    _spId:              f.id || f.ID || '',
    id:                 f.TemplateId || f.id || f.ID,
    name:               f.Title || '',
    type:               f.TaskType || '',
    dueDaysFromQtrEnd:  parseInt(f.DueDaysFromQtrEnd) || 30,
    defaultOwnerId:     f.DefaultOwnerId || '',
  };
}
function normaliseComment(f) {
  return {
    _spId:      f.id || f.ID || '',
    id:         f.CommentId || f.id || f.ID,
    taskId:     f.TaskId    || '',
    stepId:     f.StepId    || '',
    authorId:   f.AuthorId  || '',
    text:       f.CommentText || '',
    time:       f.CommentTime || '',
    ts:         f.Timestamp  || 0,
    tsIso:      f.TimestampISO || '',
    isResolved: f.IsResolved === 'Yes' || f.IsResolved === true,
  };
}
function normaliseSignOff(f) {
  return {
    _spId:      f.id || f.ID || '',
    id:         f.SignOffId   || f.id || f.ID,
    refId:      f.RefId       || '',   // _spId of the task or step
    refType:    f.RefType     || '',   // "task" | "step"
    refName:    f.RefName     || '',
    userId:     f.UserId      || '',
    userName:   f.UserName    || '',
    fromStatus: f.FromStatus  || '',
    toStatus:   f.ToStatus    || '',
    ts:         f.Timestamp   || '',
    tsIso:      f.TimestampISO|| '',
  };
}
function normaliseLock(f) {
  return {
    _spId:   f.id || f.ID || '',
    id:      f.LockId  || f.id || f.ID,
    quarter: f.Quarter || '',
    year:    parseInt(f.Year) || 0,
    lockedBy:   f.LockedBy   || '',
    lockedAt:   f.LockedAt   || '',
  };
}
function normaliseAttachment(f) {
  return {
    _spId:       f.id || f.ID || '',
    id:          f.AttachmentId || f.id || f.ID,
    stepId:      f.StepId       || '',
    taskId:      f.TaskId       || '',
    label:       f.Title        || f.Label || '',
    url:         f.FileUrl      || '',
    linkedBy:    f.LinkedBy     || '',
    linkedAt:    f.LinkedAt     || '',
    versionNote: f.VersionNote  || '',
  };
}

function normaliseQuarterDate(f) {
  // CalOverrides stores { wdNum: 'YYYY-MM-DD' } manual date overrides as JSON
  let calOverrides = {};
  try { calOverrides = JSON.parse(f.CalOverrides || '{}'); } catch { calOverrides = {}; }
  return {
    _spId:          f.id || f.ID || '',
    quarter:        f.Quarter || '',
    year:           parseInt(f.Year) || 0,
    wd1Date:        (f.WD1Date || '').slice(0,10),
    calOverrides,
    secFilingDate:  (f.SECFilingDate  || '').slice(0,10),
    earningsDate:   (f.EarningsCallDate || '').slice(0,10),
    earningsTime:   f.EarningsCallTime  || '',

  };
}
function normaliseStep(f) {
  return {
    _spId:   f.id || f.ID || '',
    id:      f.StepId    || f.id || f.ID,
    taskId:  f.TaskId    || '',
    name:    f.Title     || '',
    order:   parseInt(f.StepOrder) || 0,
    status:  f.Status    || 'Not Started',
    ownerId: f.OwnerId   || '',
    reviewerId:  f.ReviewerId  || '',
    reviewer2Id: f.Reviewer2Id || '',
    dueDate: (f.DueDate  || '').slice(0,10),
    note:          f.Note          || '',   // legacy short note (kept for compat)
    notes:         f.Notes         || '',   // rich multi-line notes
    applicability: f.Applicability || 'All Quarters',
    workdayNum:    f.WorkdayNum ? parseInt(f.WorkdayNum) : null,
    requiresPrev:  f.RequiresPrev === true || f.RequiresPrev === 'Yes',
    reassignRequested: f.ReassignRequested === 'Yes' || f.ReassignRequested === true,
    reassignNote:      f.ReassignNote || '',
  };
}
function normaliseStepTemplate(f) {
  return {
    _spId:      f.id || f.ID || '',
    id:         f.StepTemplateId || f.id || f.ID,
    templateId: f.TemplateId     || '',
    name:       f.Title          || '',
    order:      parseInt(f.StepOrder) || 0,
    defaultOwnerId: f.DefaultOwnerId || '',
    dueDaysFromQtrEnd: parseInt(f.DueDaysFromQtrEnd) || 0,
    workdayNum:    f.WorkdayNum ? parseInt(f.WorkdayNum) : null,
    requiresPrev:  f.RequiresPrev === true || f.RequiresPrev === 'Yes',
  };
}

// ── LOAD ALL DATA ─────────────────────────────────────────────
// _fetchAndStore does the actual fetching and normalisation.
// loadAllData calls it with all 10 lists (initial load).
// refreshData calls it skipping static template lists unless Admin is active.
async function _fetchAndStore({ skipStaticLists = false } = {}) {
  const adminActive = document.querySelector('#view-admin.active') !== null;
  const skipTpls    = skipStaticLists && !adminActive;

  const [rawTasks, rawUsers, rawTemplates, rawComments, rawSteps, rawStepTpls,
         rawSignOffs, rawLocks, rawAttachments, rawQDates] = await Promise.all([
    getListItems(LISTS.tasks),
    getListItems(LISTS.users),
    skipTpls ? Promise.resolve(null) : getListItems(LISTS.templates),
    getListItems(LISTS.comments),
    getListItems(LISTS.steps),
    skipTpls ? Promise.resolve(null) : getListItems(LISTS.stepTemplates),
    getListItems(LISTS.signOffs),
    getListItems(LISTS.locks),
    getListItems(LISTS.attachments),
    getListItems(LISTS.quarterDates),
  ]);

  _tasks        = rawTasks.map(normaliseTask);
  _users        = rawUsers.map(normaliseUser);
  buildOwnerColorMap();
  if (rawTemplates)  _templates     = rawTemplates.map(normaliseTemplate);
  _comments     = rawComments.map(normaliseComment);
  _steps        = rawSteps.map(normaliseStep);
  if (rawStepTpls)   _stepTemplates = rawStepTpls.map(normaliseStepTemplate);
  _signOffs     = rawSignOffs.map(normaliseSignOff);
  _locks        = rawLocks.map(normaliseLock);
  _attachments  = rawAttachments.map(normaliseAttachment);
  _quarterDates = rawQDates.map(normaliseQuarterDate);
  _lastRefreshed = new Date();
}

async function loadAllData() {
  showLoadingOverlay(true);
  try {
    await _fetchAndStore({ skipStaticLists: false });
  } catch(e) {
    console.error("Data load error:", e);
    showError("Could not load data from SharePoint. Check your config and list names. " + e.message);
  } finally {
    showLoadingOverlay(false);
  }
}

async function refreshData() {
  try {
    // Skip static template lists on background polls — see _fetchAndStore
    await _fetchAndStore({ skipStaticLists: true });
    updateRefreshStamp();
    renderCurrentView();
  } catch(e) { console.warn("Refresh error:", e); }
}

async function refreshComments() {
  if (!commentingTaskId) return;
  try {
    const raw  = await getListItems(LISTS.comments);
    _comments  = raw.map(normaliseComment);
    renderCommentList();
  } catch(e) { console.warn("Comment refresh error:", e); }
}

function updateRefreshStamp() {
  const el = document.getElementById('refresh-stamp');
  if (!el || !_lastRefreshed) return;
  const mins = Math.floor((new Date() - _lastRefreshed) / 60000);
  el.textContent = mins === 0 ? 'Updated just now' : `Updated ${mins}m ago`;
}

// Update the stamp every minute so "Xm ago" stays current
const _stampTimer = setInterval(updateRefreshStamp, 60000);

async function manualRefresh() {
  const btn = document.getElementById('refresh-btn');
  if (btn) { btn.disabled = true; btn.textContent = '↺ Refreshing…'; }
  await refreshData();
  if (btn) { btn.disabled = false; btn.textContent = '↺ Refresh'; }
}

function startPolling() {
  if (_pollTimer) clearInterval(_pollTimer);
  _pollTimer = setInterval(async () => {
    await refreshData();
  }, 60000); // every 60 seconds — safe with direct Graph API, no Power Automate run limits
}

function startCommentPolling() {
  if (_commentPollTimer) clearInterval(_commentPollTimer);
  _commentPollTimer = setInterval(refreshComments, 60000); // every 60 seconds
}

function stopCommentPolling() {
  if (_commentPollTimer) { clearInterval(_commentPollTimer); _commentPollTimer = null; }
}

// ── STATE ─────────────────────────────────────────────────────
let _lastRefreshed = null; // Date of last successful data refresh
let _tableSort = { col: null, dir: 1 }; // col: 'due'|'status'|'owner'|'name', dir: 1=asc -1=desc
let currentUser      = null;
let activeFilter     = 'all';
let activeTypeFilter = 'all';
let calYear          = new Date().getFullYear();
let calMonth         = new Date().getMonth();
let editingTaskId    = null;
let commentingTaskId = null;
let _calResolve      = null;

// ── HELPERS ───────────────────────────────────────────────────
function initials(name)  { return (name||'').split(' ').map(p=>p[0]).join('').slice(0,2).toUpperCase(); }
function formatDate(dateStr) {
  if (!dateStr) return '—';
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('en-US', {month:'short', day:'numeric', year:'numeric'});
}
function deadlineStatus(dueDate, status) {
  if (status === 'Not Applicable') return 'done'; // N/A never shown as overdue
  if (status === 'Complete') return 'done';
  const today = new Date(); today.setHours(0,0,0,0);
  const due   = new Date(dueDate + 'T00:00:00');
  const diff  = Math.floor((due - today) / MS_PER_DAY);
  if (diff < 0)  return 'overdue';
  if (diff <= 7) return 'soon';
  return 'ok';
}
function daysLabel(dueDate, status) {
  if (!dueDate || status === 'Complete' || status === 'Not Applicable') return '';
  const today = new Date(); today.setHours(0,0,0,0);
  const diff  = Math.floor((new Date(dueDate+'T00:00:00') - today) / MS_PER_DAY);
  if (diff < 0)  return `<span style="font-size:10px;font-weight:700;color:var(--deadline-overdue)">${Math.abs(diff)}d overdue</span>`;
  if (diff === 0) return `<span style="font-size:10px;font-weight:700;color:var(--deadline-soon)">Today</span>`;
  if (diff <= 7)  return `<span style="font-size:10px;color:var(--deadline-soon)">${diff}d</span>`;
  return '';
}

function isThisWeek(dueDate) {
  if (!dueDate) return false;
  const today = new Date(); today.setHours(0,0,0,0);
  const diff  = Math.floor((new Date(dueDate+'T00:00:00') - today) / MS_PER_DAY);
  return diff >= 0 && diff <= 7;
}
function quarterEndDate(q, yr) {
  return {Q1:`${yr}-03-31`,Q2:`${yr}-06-30`,Q3:`${yr}-09-30`,Q4:`${yr}-12-31`}[q];
}
function addDays(dateStr, days) {
  const d = new Date(dateStr+'T00:00:00'); d.setDate(d.getDate()+days);
  return d.toISOString().slice(0,10);
}

// ══════════════════════════════════════════════════════════════
// ── WORKING DAY ENGINE ───────────────────────────────────────
// Handles US Federal holidays, NYSE extra holidays (Good Friday),
// and custom company holidays managed in Admin.
// Applied only during rollforward & template apply.
// ══════════════════════════════════════════════════════════════

// Custom company holidays stored in localStorage (no SharePoint
// list needed — admin manages them in the Admin panel).
function loadCustomHolidays() {
  try { return JSON.parse(localStorage.getItem('ft_custom_holidays') || '[]'); }
  catch { return []; }
}
function saveCustomHolidays(list) {
  try { localStorage.setItem('ft_custom_holidays', JSON.stringify(list)); }
  catch(e) { console.warn('Could not save custom holidays to localStorage:', e.message); }
}

// ── US Federal Holidays (fixed-date + rule-based) ─────────────
function usFederalHolidays(year) {
  const h = new Set();
  const iso = d => d.toISOString().slice(0,10);

  // Helper: nth weekday of a month  (n=1 first, n=-1 last)
  function nthWeekday(yr, month, weekday, n) {
    if (n > 0) {
      const d = new Date(yr, month, 1);
      while (d.getDay() !== weekday) d.setDate(d.getDate()+1);
      d.setDate(d.getDate() + (n-1)*7);
      return d;
    } else {
      const d = new Date(yr, month+1, 0); // last day of month
      while (d.getDay() !== weekday) d.setDate(d.getDate()-1);
      return d;
    }
  }

  // Observed rule: if holiday falls on Sat → Fri, Sun → Mon
  function observed(d) {
    const dow = d.getDay();
    if (dow === 6) { const f=new Date(d); f.setDate(f.getDate()-1); return f; }
    if (dow === 0) { const m=new Date(d); m.setDate(m.getDate()+1); return m; }
    return d;
  }

  const fixed = [
    new Date(year,  0,  1),   // New Year's Day
    new Date(year,  6,  4),   // Independence Day
    new Date(year, 10, 11),   // Veterans Day
    new Date(year, 11, 25),   // Christmas Day
  ];
  fixed.forEach(d => h.add(iso(observed(d))));

  // Rule-based
  h.add(iso(nthWeekday(year, 0, 1, 3)));   // MLK Day: 3rd Mon Jan
  h.add(iso(nthWeekday(year, 1, 1, 3)));   // Presidents Day: 3rd Mon Feb
  h.add(iso(nthWeekday(year, 4, 1,-1)));   // Memorial Day: last Mon May
  const juneteenth = observed(new Date(year, 5, 19)); h.add(iso(juneteenth));  // Juneteenth: Jun 19, observed
  h.add(iso(nthWeekday(year, 8, 1, 1)));   // Labor Day: 1st Mon Sep
  h.add(iso(nthWeekday(year, 9, 1, 2)));   // Columbus Day: 2nd Mon Oct
  h.add(iso(nthWeekday(year, 10, 4, 4)));  // Thanksgiving: 4th Thu Nov

  return h;
}

// ── NYSE Extra Holidays ───────────────────────────────────────
// Good Friday (Friday before Easter Sunday)
function easterSunday(year) {
  // Anonymous Gregorian algorithm
  const a = year % 19, b = Math.floor(year/100), c = year % 100;
  const d = Math.floor(b/4), e = b % 4, f = Math.floor((b+8)/25);
  const g = Math.floor((b-f+1)/3), h = (19*a+b-d-g+15) % 30;
  const i = Math.floor(c/4), k = c % 4;
  const l = (32+2*e+2*i-h-k) % 7;
  const m = Math.floor((a+11*h+22*l)/451);
  const month = Math.floor((h+l-7*m+114)/31) - 1; // 0-indexed
  const day   = ((h+l-7*m+114) % 31) + 1;
  return new Date(year, month, day);
}

function nyseExtraHolidays(year) {
  const h   = new Set();
  const iso = d => d.toISOString().slice(0,10);
  const easter = easterSunday(year);
  const goodFriday = new Date(easter); goodFriday.setDate(easter.getDate()-2);
  h.add(iso(goodFriday));
  return h;
}

// ── Master holiday checker (memoised by year) ────────────────
const _holidayCache = {};
function isHoliday(dateStr) {
  const year = parseInt(dateStr.slice(0,4));
  if (!_holidayCache[year]) {
    const fed  = usFederalHolidays(year);
    const nyse = nyseExtraHolidays(year);
    loadCustomHolidays().forEach(d => fed.add(d)); // merge custom into fed set
    _holidayCache[year] = fed; // combined set
  }
  return _holidayCache[year].has(dateStr);
}
// Call this whenever custom holidays are changed so the cache is rebuilt
function clearHolidayCache() { Object.keys(_holidayCache).forEach(k => delete _holidayCache[k]); }

function isWeekend(dateStr) {
  const dow = new Date(dateStr+'T00:00:00').getDay();
  return dow === 0 || dow === 6;
}

function isNonWorkingDay(dateStr) {
  return isWeekend(dateStr) || isHoliday(dateStr);
}

// ── Adjust to nearest working day ─────────────────────────────
// Direction: "closest" — if equidistant, prefer the day before.
function nearestWorkingDay(dateStr) {
  if (!isNonWorkingDay(dateStr)) return dateStr; // already a working day

  let before = dateStr, after = dateStr;
  let bDays = 0, aDays = 0;

  // Walk backwards to find prev working day
  for (let i = 1; i <= 14; i++) {
    const d = addDays(dateStr, -i);
    if (!isNonWorkingDay(d)) { before = d; bDays = i; break; }
  }
  // Walk forwards to find next working day
  for (let i = 1; i <= 14; i++) {
    const d = addDays(dateStr, i);
    if (!isNonWorkingDay(d)) { after = d; aDays = i; break; }
  }

  // If no working day found backward, bDays stays 0; fall through to `after`
  // Ties (bDays === aDays > 0) go to the earlier day (more conservative for deadlines)
  if (bDays === 0) return after;
  if (aDays === 0) return before;
  return bDays <= aDays ? before : after;
}

// ── Admin: Custom Holiday Manager ────────────────────────────

// ── CLOSE CALENDARS PANEL (in Admin) ────────────────────────
function renderCloseCalendarsPanel() {
  const el = document.getElementById('admin-calendars-list');
  if (!el) return;
  const cur  = new Date().getFullYear();
  const rows = [];
  // Show last year, current year, next year — enough history without clutter
  for (let yr = cur - 1; yr <= cur + 1; yr++) {
    for (const q of ['Q1','Q2','Q3','Q4']) {
      const cal = loadCloseCalendar(q, yr);
      rows.push({ q, yr, cal });
    }
  }
  el.innerHTML = rows.map(({ q, yr, cal }) => `
    <div class="admin-user-row">
      <div class="admin-user-info">
        <div class="user-name-sm">${q} ${yr}</div>
        <div class="user-role-sm">${cal
          ? `WD1 = ${formatDate(cal.wd1Date)} · ${cal.days.length} working days mapped`
          : '<span style="color:var(--text-faint)">No calendar set</span>'
        }</div>
      </div>
      <button class="btn-secondary" style="font-size:12px;padding:5px 12px"
        onclick="openCloseCalendarAdmin('${q}',${yr})">
        ${cal ? '✏️ Edit' : '+ Set'}
      </button>
    </div>`).join('');
}

function renderCustomHolidays() {
  const el = document.getElementById('admin-holidays-list');
  if (!el) return;

  // Populate built-in holiday preview for current year
  const previewEl = document.getElementById('holiday-preview');
  if (previewEl) {
    const yr   = new Date().getFullYear();
    const summ = holidaySummaryForYear(yr);
    const tag  = (label, dates) => dates.length
      ? `<div style="margin-bottom:6px"><span style="font-weight:600;color:var(--text-muted)">${label}:</span> `
        + dates.map(d => `<span style="display:inline-block;background:var(--bg-secondary);border-radius:4px;padding:1px 6px;margin:1px 2px;font-size:11px">${formatDate(d)}</span>`).join(' ')
        + '</div>'
      : '';
    previewEl.innerHTML =
      tag('Federal', summ.federal) +
      tag('NYSE (Good Friday)', summ.nyse) +
      (summ.custom.length ? tag('Company', summ.custom) : '');
  }
  const holidays = loadCustomHolidays().sort();
  if (!holidays.length) {
    el.innerHTML = '<p class="text-muted" style="font-size:13px;padding:10px 22px">No custom holidays added yet.</p>';
    return;
  }
  el.innerHTML = holidays.map(d => `
    <div class="admin-user-row">
      <div class="admin-user-info">
        <div class="user-name-sm">${formatDate(d)}</div>
        <div class="user-role-sm">${d}</div>
      </div>
      <button class="icon-btn" onclick="removeCustomHoliday('${d}')">🗑</button>
    </div>`).join('');
}

function addCustomHoliday() {
  const input = document.getElementById('holiday-input');
  const val   = input?.value?.trim();
  if (!val || !/^\d{4}-\d{2}-\d{2}$/.test(val)) {
    showToast('Please enter a valid date in YYYY-MM-DD format.', 'warning'); return;
  }
  const list = loadCustomHolidays();
  if (!list.includes(val)) { list.push(val); saveCustomHolidays(list); clearHolidayCache(); }
  if (input) input.value = '';
  renderCustomHolidays();
}

function removeCustomHoliday(dateStr) {
  saveCustomHolidays(loadCustomHolidays().filter(d => d !== dateStr));
  clearHolidayCache();
  renderCustomHolidays();
}

// ── Working day summary for a given year (for display) ────────
function holidaySummaryForYear(year) {
  const fed  = [...usFederalHolidays(year)].sort();
  const nyse = [...nyseExtraHolidays(year)].sort();
  return { federal: fed, nyse, custom: loadCustomHolidays().filter(d=>d.startsWith(String(year))) };
}

// ══════════════════════════════════════════════════════════════
// ── CLOSE CALENDAR ENGINE ────────────────────────────────────
// Each quarter has a "close calendar": a mapping from workday
// numbers (WD1, WD2 …) to real calendar dates, starting from
// a manually-set WD1 date and advancing through working days.
//
// Stored in localStorage keyed by "Q1-2025", "Q2-2025" etc.
// ══════════════════════════════════════════════════════════════

function calendarKey(quarter, year) { return `${quarter}-${year}`; }

function loadCloseCalendar(quarter, year) {
  // FT_QuarterDates is the source of truth — WD1 and overrides are shared across the team.
  // localStorage is only used as a fallback before SharePoint data has loaded.
  const qd = _quarterDates.find(d => d.quarter===quarter && d.year===parseInt(year));
  if (qd?.wd1Date) {
    // Build the calendar using SharePoint WD1 and SharePoint overrides
    return buildCloseCalendar(qd.wd1Date, 40, qd.calOverrides || {});
  }
  // Fallback: localStorage only (setup wizard before SharePoint is connected)
  try { return JSON.parse(localStorage.getItem('ft_cal_' + calendarKey(quarter, year)) || 'null'); }
  catch { return null; }
}

function saveCloseCalendar(quarter, year, calObj) {
  if (typeof calObj === 'string') calObj = buildCloseCalendar(calObj); // accept wd1 string
  if (!calObj) return;

  const wd1       = calObj.wd1Date;
  const overrides = calObj.overrides || {};
  if (!wd1) return;

  // Serialise overrides for SharePoint storage
  const overridesJson = Object.keys(overrides).length
    ? JSON.stringify(overrides)
    : '';

  const existing = _quarterDates.find(d => d.quarter===quarter && d.year===parseInt(year));
  if (existing) {
    // Update WD1 and overrides in SharePoint — both shared with the whole team
    existing.wd1Date     = wd1;
    existing.calOverrides = overrides;
    updateListItem(LISTS.quarterDates, existing._spId, {
      WD1Date:      wd1,
      CalOverrides: overridesJson,
    }).catch(e => console.warn('Could not save calendar to SharePoint:', e.message));
  } else {
    // Create a new FT_QuarterDates record with WD1 and overrides
    const fields = {
      Title:        `${quarter} ${year}`,
      Quarter:      quarter,
      Year:         String(year),
      WD1Date:      wd1,
      CalOverrides: overridesJson,
    };
    createListItem(LISTS.quarterDates, fields).then(created => {
      _quarterDates.push(normaliseQuarterDate({ ...fields, id: created?.id || uid() }));
    }).catch(e => console.warn('Could not create QuarterDates for WD1:', e.message));
  }

  // Keep localStorage in sync as a local cache for faster first render
  try { localStorage.setItem('ft_cal_' + calendarKey(quarter, year), JSON.stringify(calObj)); }
  catch(e) { console.warn('Could not cache calendar in localStorage:', e.message); }
}

// Build a calendar: given WD1 date, generate WD1 … WD(maxDays)
// skipping weekends and holidays. Returns { wd1Date, days: [{num, date}] }
// overrides: optional { wdNum: 'YYYY-MM-DD' } map of manual date edits
function buildCloseCalendar(wd1DateStr, maxDays = 40, overrides = {}) {
  const days = [];
  let current = wd1DateStr;
  // Make sure WD1 itself is a working day (shift forward if not)
  // unless it has been manually overridden
  if (!overrides[1]) {
    while (isNonWorkingDay(current)) current = addDays(current, 1);
  }
  for (let n = 1; n <= maxDays; n++) {
    // Use override if provided for this workday number
    const date = overrides[n] || current;
    days.push({ num: n, date, overridden: !!overrides[n] });
    // Advance to next working day (from the non-overridden position)
    let next = addDays(current, 1);
    while (!overrides[n+1] && isNonWorkingDay(next)) next = addDays(next, 1);
    current = next;
  }
  return { wd1Date: days[0].date, days, overrides };
}

// Convert a workday number to a calendar date for a given quarter
function workdayToDate(workdayNum, quarter, year) {
  if (!workdayNum) return null;
  const cal = loadCloseCalendar(quarter, year);
  if (!cal) return null;
  const entry = cal.days.find(d => d.num === workdayNum);
  return entry ? entry.date : null;
}

// Convert a calendar date back to a workday number for a given quarter
function dateToWorkday(dateStr, quarter, year) {
  if (!dateStr) return null;
  const cal = loadCloseCalendar(quarter, year);
  if (!cal) return null;
  const entry = cal.days.find(d => d.date === dateStr);
  return entry ? entry.num : null;
}

// Format workday + date for display: "WD3 · Apr 3, 2025"
// Accepts optional quarter + year so it can resolve a workday number to a real
// date via the close calendar when no fixed dateStr is present.
function formatWorkdayDate(workdayNum, dateStr, quarter, year) {
  if (!dateStr && !workdayNum) return '—';
  const wdLabel = workdayNum ? `<span class="wd-badge">WD${workdayNum}</span> ` : '';
  // If no fixed date but we have a workday number and quarter/year, resolve via calendar
  const resolvedDate = dateStr || (workdayNum && quarter && year
    ? workdayToDate(workdayNum, quarter, year)
    : null);
  const dateLabel = resolvedDate ? formatDate(resolvedDate) : '';
  return wdLabel + dateLabel;
}

// ── CLOSE CALENDAR UI ─────────────────────────────────────────

// Shared state for the editable calendar grid (used by both prompt + admin)
let _editingCal = { quarter: null, year: null, overrides: {} };

// Renders an editable calendar grid into the given container element.
// Each WD cell has a date input — changing it updates _editingCal.overrides.
function renderEditableCalGrid(containerId, wd1DateStr, maxDays, isPrompt) {
  const el = document.getElementById(containerId);
  if (!el || !wd1DateStr) return;
  // Validate format before embedding in onclick attributes
  if (!/^\d{4}-\d{2}-\d{2}$/.test(wd1DateStr)) return;
  const overrides = _editingCal.overrides;
  const cal       = buildCloseCalendar(wd1DateStr, maxDays, overrides);

  const resetBtn = `<button class="btn-secondary"
    style="font-size:11px;padding:4px 10px"
    onclick="resetCalOverrides('${containerId}','${wd1DateStr}',${maxDays},${isPrompt})">
    ↺ Reset overrides
  </button>`;

  el.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
      <p style="font-size:12px;font-weight:600;color:var(--text-muted);margin:0">
        Working day calendar — click any date to override
      </p>
      ${resetBtn}
    </div>
    <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:6px;max-height:340px;overflow-y:auto;padding-bottom:4px">
      ${cal.days.map(d => {
        const isOverridden = !!overrides[d.num];
        const isHolidayDay = isHoliday(d.date);
        const isWknd       = isWeekend(d.date);
        const bgColor      = isOverridden
          ? 'var(--blue)'
          : isHolidayDay
            ? '#fef3c7'
            : isWknd
              ? '#fee2e2'
              : 'var(--blue-pale)';
        const textColor    = isOverridden ? '#fff' : 'var(--navy)';
        const wdColor      = isOverridden ? 'rgba(255,255,255,.8)' : 'var(--blue)';
        const hint         = isOverridden
          ? 'Manually set'
          : isHolidayDay
            ? '⚠ Holiday'
            : isWknd
              ? '⚠ Weekend'
              : '';
        return `<div style="background:${bgColor};border-radius:6px;padding:5px 6px;text-align:center;position:relative" title="${hint}">
          <div style="font-size:10px;font-weight:700;color:${wdColor}">WD${d.num}</div>
          <input type="date" value="${d.date}"
            style="border:none;background:transparent;font-size:10px;color:${textColor};font-family:inherit;text-align:center;width:100%;cursor:pointer;padding:0;outline:none"
            onchange="overrideCalDay(${d.num},'${containerId}','${wd1DateStr}',${maxDays},${isPrompt},this.value)"
          />
          ${hint ? `<div style="font-size:9px;color:${isOverridden?'rgba(255,255,255,.7)':'#92400e'};margin-top:1px">${hint}</div>` : ''}
        </div>`;
      }).join('')}
    </div>
    <p style="font-size:11px;color:var(--text-faint);margin-top:8px">
      🔵 Overridden &nbsp; 🟡 Holiday &nbsp; 🔴 Weekend &nbsp; All other days are working days
    </p>`;
}

function overrideCalDay(wdNum, containerId, wd1DateStr, maxDays, isPrompt, newDate) {
  if (!newDate) return;
  _editingCal.overrides[wdNum] = newDate;
  renderEditableCalGrid(containerId, wd1DateStr, maxDays, isPrompt);
}

function resetCalOverrides(containerId, wd1DateStr, maxDays, isPrompt) {
  _editingCal.overrides = {};
  renderEditableCalGrid(containerId, wd1DateStr, maxDays, isPrompt);
}

// Called from rollForward — opens the close calendar setup modal
// before proceeding with the actual copy. Returns a Promise that
// resolves with the calendar object once confirmed, or null if cancelled.
function promptCloseCalendar(toQ, toY) {
  return new Promise(resolve => {
    const existing  = loadCloseCalendar(toQ, toY);
    const qEndDate  = quarterEndDate(toQ, toY);
    let suggested   = addDays(qEndDate, 1);
    while (isNonWorkingDay(suggested)) suggested = addDays(suggested, 1);

    // Restore existing overrides if any
    _editingCal = { quarter: toQ, year: toY, overrides: existing?.overrides || {} };

    document.getElementById('modal-title').textContent =
      `Step 1 of 2 — Set Close Calendar for ${toQ} ${toY}`;

    document.getElementById('modal-body').innerHTML = `
      <p style="font-size:13px;color:var(--text-muted);margin-bottom:16px">
        Set <strong>Workday 1</strong> for ${toQ} ${toY}, then adjust any specific
        dates below — for example if your team works on Good Friday, click that
        date and change it to the actual day your team is working.
      </p>
      <div class="form-group">
        <label>Workday 1 date for ${toQ} ${toY}</label>
        <input type="date" id="wd1-input" value="${existing?.wd1Date || suggested}" />
        <p style="font-size:11px;color:var(--text-faint);margin-top:4px">
          First working day of your close — quarter ends ${formatDate(qEndDate)}.
        </p>
      </div>
      <div id="cal-preview" style="margin-top:16px"></div>
      <div class="modal-footer">
        <button class="btn-secondary" onclick="closeAllModals();if(_calResolve){_calResolve(null);_calResolve=null;}">Cancel</button>
        <button class="btn-primary" onclick="confirmCloseCalendar('${toQ}',${toY})">Confirm & Continue →</button>
      </div>`;

    openModal();
    _calResolve = resolve;

    const input = document.getElementById('wd1-input');
    function updatePreview() {
      const val = input.value; if (!val) return;
      renderEditableCalGrid('cal-preview', val, 20, true);
    }
    input.addEventListener('input', updatePreview);
    updatePreview();
  });
}

function confirmCloseCalendar(toQ, toY) {
  const input = document.getElementById('wd1-input');
  const val   = input?.value?.trim();
  if (!val) { showToast('Please set a Workday 1 date.', 'warning'); return; }
  const cal = buildCloseCalendar(val, 40, _editingCal.overrides);
  saveCloseCalendar(toQ, toY, cal);
  closeAllModals();
  if (_calResolve) { _calResolve(cal); _calResolve = null; }
}

// Render/edit the close calendar for a quarter (accessible from Admin)
function openCloseCalendarAdmin(quarter, year) {
  const cal       = loadCloseCalendar(quarter, year);
  const qEndDate  = quarterEndDate(quarter, year);
  let suggested   = addDays(qEndDate, 1);
  while (isNonWorkingDay(suggested)) suggested = addDays(suggested, 1);

  // Load existing overrides so edits are preserved
  _editingCal = { quarter, year, overrides: cal?.overrides || {} };

  document.getElementById('modal-title').textContent = `Close Calendar — ${quarter} ${year}`;
  document.getElementById('modal-body').innerHTML = `
    <p style="font-size:13px;color:var(--text-muted);margin-bottom:12px">
      Set WD1 to anchor the calendar, then click any date cell to override it —
      for example, change Good Friday to the actual day your team is working.
    </p>
    <div class="form-group">
      <label>Workday 1 date</label>
      <input type="date" id="wd1-admin-input" value="${cal?.wd1Date || suggested}" />
    </div>
    <div id="admin-cal-preview" style="margin-top:12px"></div>
    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeAllModals()">Cancel</button>
      <button class="btn-primary" onclick="saveCalendarAdmin('${quarter}',${year})">Save Calendar</button>
    </div>`;

  openModal();

  const input = document.getElementById('wd1-admin-input');
  function updatePreview() {
    const val = input?.value; if (!val) return;
    renderEditableCalGrid('admin-cal-preview', val, 30, false);
  }
  input.addEventListener('input', updatePreview);
  updatePreview();
}

function saveCalendarAdmin(quarter, year) {
  const input = document.getElementById('wd1-admin-input');
  const val   = input?.value?.trim();
  if (!val) { showToast('Please set a Workday 1 date.', 'warning'); return; }
  const cal = buildCloseCalendar(val, 40, _editingCal.overrides);
  saveCloseCalendar(quarter, year, cal);
  closeAllModals();
  renderAdmin();
  const overrideCount = Object.keys(_editingCal.overrides).length;
  showToast(`Close calendar saved — WD1 ${formatDate(cal.wd1Date)}${overrideCount ? `, ${overrideCount} override(s)` : ''}`, 'success');
}

// Live WD# → date resolver called from modal inputs
function resolveWdToDate(wdInputId, dateInputId, quarter, year) {
  const wdEl   = document.getElementById(wdInputId);
  const dateEl = document.getElementById(dateInputId);
  if (!wdEl || !dateEl || !quarter || !year) return;
  const wdNum  = parseInt(wdEl.value);
  if (!wdNum) return;
  const resolved = workdayToDate(wdNum, quarter, parseInt(year));
  if (resolved) dateEl.value = resolved;
}
function uid() { return 'ft-' + Date.now() + '-' + Math.random().toString(36).slice(2,6); }
function typeBadgeClass(type) {
  return {
    'Close':          'badge-type-close',
    'Financial Report':'badge-type-financial',
    'Master SS':      'badge-type-master',
    'Ops Book':       'badge-type-ops',
    'Other':          'badge-type-other',
    'Press Release':  'badge-type-press',
    'Post-Filing':    'badge-type-post',
    'Pre-Filing':     'badge-type-pre',
  }[type] || 'badge-type-other';
}
function statusBadgeClass(s) {
  return {'Not Started':'status-not-started','In Progress':'status-in-progress',
          'Ready for Review 1':'status-review','Ready for Review':'status-review',
          'Ready for Review 2':'status-review2',
          'Complete':'status-complete','Not Applicable':'status-na'}[s]||'status-not-started';
}
function dotClass(ds) {
  return {overdue:'dot-overdue',soon:'dot-soon',ok:'dot-ok',done:'dot-done'}[ds]||'dot-ok';
}
// Returns the background colour for a task-type calendar chip.
function calChipBg(type) {
  return {
    'Close':           '#e0f2fe',
    'Financial Report':'#dbeafe',
    'Master SS':       '#d1fae5',
    'Ops Book':        '#fce7f3',
    'Other':           '#f1f5f9',
    'Press Release':   '#ede9fe',
    'Post-Filing':     '#fef3c7',
    'Pre-Filing':      '#ecfdf5',
  }[type] || '#f1f5f9';
}
// Returns the foreground colour for a task-type calendar chip.
function calChipFg(type) {
  return {
    'Close':           '#0369a1',
    'Financial Report':'#1d4ed8',
    'Master SS':       '#065f46',
    'Ops Book':        '#9d174d',
    'Other':           '#475569',
    'Press Release':   '#5b21b6',
    'Post-Filing':     '#92400e',
    'Pre-Filing':      '#065f46',
  }[type] || '#334155';
}
function escHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── UI HELPERS ────────────────────────────────────────────────

// ── IN-APP TOAST / CONFIRM SYSTEM ───────────────────────────
// Replaces native alert() / confirm() throughout the app

function showToast(msg, type='info', duration=4000) {
  // type: 'info' | 'success' | 'error' | 'warning'
  let container = document.getElementById('toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    document.body.appendChild(container);
  }
  const id   = 'toast-' + Date.now();
  const icons = { info:'ℹ', success:'✅', error:'❌', warning:'⚠️' };
  const toast = document.createElement('div');
  toast.id        = id;
  toast.className = `app-toast app-toast-${type}`;
  toast.innerHTML = `<span class="toast-icon">${icons[type]}</span>
    <span class="toast-msg">${escHtml(msg)}</span>
    <button class="toast-close" onclick="this.closest('.app-toast').remove()">✕</button>`;
  container.appendChild(toast);
  // Animate in
  requestAnimationFrame(() => toast.classList.add('visible'));
  if (duration > 0) setTimeout(() => {
    toast.classList.remove('visible');
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

function showConfirm(msg, onConfirm, onCancel, confirmLabel='Confirm', danger=false) {
  // In-app confirm replaces confirm()
  let overlay = document.getElementById('confirm-overlay');
  if (overlay) overlay.remove();
  overlay = document.createElement('div');
  overlay.id        = 'confirm-overlay';
  overlay.className = 'confirm-overlay';
  overlay.innerHTML = `
    <div class="confirm-box">
      <div class="confirm-msg">${escHtml(msg)}</div>
      <div class="confirm-actions">
        <button class="btn-secondary" id="confirm-cancel">${escHtml('Cancel')}</button>
        <button class="${danger?'btn-danger':'btn-primary'}" id="confirm-ok">${escHtml(confirmLabel)}</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  document.getElementById('confirm-cancel').onclick = () => {
    overlay.remove();
    if (onCancel) onCancel();
  };
  document.getElementById('confirm-ok').onclick = () => {
    overlay.remove();
    onConfirm();
  };
  requestAnimationFrame(() => overlay.classList.add('visible'));
}

function showInlineConfirm(anchorEl, msg, onConfirm, confirmLabel='Delete', danger=true) {
  // Small inline confirm that appears next to the element
  const existing = document.getElementById('inline-confirm-pop');
  if (existing) existing.remove();
  const pop = document.createElement('div');
  pop.id        = 'inline-confirm-pop';
  pop.className = 'inline-confirm-pop';
  pop.innerHTML = `<span>${escHtml(msg)}</span>
    <button class="${danger?'btn-danger':'btn-primary'} small" id="inline-confirm-ok">${escHtml(confirmLabel)}</button>
    <button class="btn-secondary small" onclick="document.getElementById('inline-confirm-pop').remove()">Cancel</button>`;
  document.body.appendChild(pop);
  // Position near anchor
  const rect = anchorEl.getBoundingClientRect();
  pop.style.top  = (rect.bottom + window.scrollY + 6) + 'px';
  pop.style.left = Math.max(8, rect.left + window.scrollX - 60) + 'px';
  // Close on outside click — store handler reference so confirm-ok can also remove
  // it, preventing a stale listener if the pop is removed programmatically.
  function _outsideClickHandler(e) {
    if (!pop.contains(e.target) && e.target !== anchorEl) {
      pop.remove();
      document.removeEventListener('click', _outsideClickHandler);
    }
  }
  document.getElementById('inline-confirm-ok').onclick = () => {
    pop.remove();
    document.removeEventListener('click', _outsideClickHandler);
    onConfirm();
  };
  setTimeout(() => document.addEventListener('click', _outsideClickHandler), 0);
}

function showLoadingOverlay(show, msg='') {
  let el = document.getElementById('loading-overlay');
  if (!el) {
    el = document.createElement('div');
    el.id = 'loading-overlay';
    document.body.appendChild(el);
  }
  if (show) {
    el.className = 'loading-overlay';
    el.innerHTML = `<div class="loading-spinner"></div><span>${escHtml(msg||'Loading…')}</span>`;
    el.style.display = 'flex';
  } else {
    el.style.display = 'none';
  }
}

function showError(msg) {
  const el = document.getElementById('sp-error');
  if (el) { el.textContent = msg; el.classList.remove('hidden'); }
}

// ── LOGIN (Microsoft SSO) ────────────────────────────────────
async function loginWithMicrosoft() {
  // Clear any previous error before retrying
  const errEl = document.getElementById('sp-error');
  if (errEl) { errEl.textContent = ''; errEl.classList.add('hidden'); }
  try {
    await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
    await afterMicrosoftLogin();
  } catch(e) {
    showError("Microsoft sign-in failed: " + e.message);
  }
}

async function afterMicrosoftLogin() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) return;
  const msAccount = accounts[0];

  showLoadingOverlay(true);
  try {
    await loadAllData();

    const email = msAccount.username?.toLowerCase() || '';
    const user  = _users.find(u => (u.email||'').toLowerCase() === email);

    if (!user) {
      showError(`Your Microsoft account (${msAccount.username}) was not found in FT_Users. Ask your admin to add your email address to the FT_Users list.`);
      showLoadingOverlay(false);
      return;
    }

    // Log straight in — no PIN needed
    currentUser = user;
    sessionStorage.setItem('ft_session', user._spId || user.id);
    launchApp();
  } catch(e) {
    showError('Could not load data: ' + e.message);
  } finally {
    showLoadingOverlay(false);
  }
}

// PIN login removed — Microsoft login handles identity directly

// Returns the canonical SharePoint item ID for the current user.
// Always use this instead of repeating currentUser._spId || currentUser.id throughout the code.
function currentUserId() { return currentUser?._spId || currentUser?.id || ''; }

function logout() {
  currentUser = null;
  _graphSiteId = null; // reset so next login resolves fresh
  sessionStorage.removeItem('ft_session');
  stopCommentPolling();
  if (_pollTimer) clearInterval(_pollTimer);
  if (typeof _stampTimer !== 'undefined') clearInterval(_stampTimer);
  msalInstance.logoutPopup().catch(()=>{});
  document.getElementById('app-screen').classList.remove('active');
  document.getElementById('login-screen').classList.add('active');
}

function initAdminYearSelects() {
  const cur = new Date().getFullYear();
  ['audit-year-select'].forEach(id => {
    const el = document.getElementById(id); if (!el) return;
    el.innerHTML = [-3,-2,-1,0,1].map(d=>`<option value="${cur+d}">${cur+d}</option>`).join('');
  });
}

function launchApp() {
  document.getElementById('login-screen').classList.remove('active');
  document.getElementById('app-screen').classList.add('active');
  document.getElementById('sidebar-user-info').innerHTML = `
    <div class="user-avatar">${initials(currentUser.name)}</div>
    <div class="user-name">${escHtml(currentUser.name)}</div>
    <div class="user-role">${escHtml(currentUser.role)}</div>`;
  document.querySelectorAll('.admin-only').forEach(el =>
    el.classList.toggle('hidden', !currentUser.isAdmin));
  populateYearSelects();
  setCurrentQuarter();
  initCompactMode();
  startPolling();
  initAdminYearSelects();
  renderSavedFilters();
  renderDashboard();
  renderPriorityCard();
  setTimeout(updateRefreshStamp, 100); // show stamp after first data load
}

function setCurrentQuarter() {
  const now  = new Date(); const m = now.getMonth();
  const defQ = m<3?'Q1':m<6?'Q2':m<9?'Q3':'Q4';
  const pref = loadQuarterPref();
  const q    = pref.q  || defQ;
  const yr   = pref.yr ? parseInt(pref.yr) : now.getFullYear();
  const qf   = document.getElementById('quarter-filter'); if(qf) qf.value = q;
  const yf   = document.getElementById('year-filter');    if(yf) yf.value = yr;
  const lbl  = document.getElementById('dashboard-quarter-label');
  if(lbl) lbl.textContent = `${q} ${yr} · Financial Reporting`;
}
function populateYearSelects() {
  const cur = new Date().getFullYear();
  // Show 4 years back and 1 forward so historical quarters remain accessible
  // as the team builds up multiple years of data. cur-3 through cur+1 = 5 options.
  const years = [cur-3, cur-2, cur-1, cur, cur+1];
  ['year-filter','template-year-select','rf-from-year','rf-to-year','kanban-year','report-year','team-year','exec-year','mytasks-year','all-year-filter'].forEach(id => {
    const el = document.getElementById(id); if(!el) return;
    el.innerHTML = years.map(y=>`<option value="${y}">${y}</option>`).join('');
    el.value = cur;
  });
}

// ── MOBILE SIDEBAR ────────────────────────────────────────────
function toggleSidebar() {
  const sb  = document.querySelector('.sidebar');
  const bd  = document.getElementById('sidebar-backdrop');
  const open = sb?.classList.toggle('open');
  if (bd) bd.classList.toggle('visible', open);
}
// Closes the mobile sidebar and removes the backdrop.
function closeSidebar() {
  document.querySelector('.sidebar')?.classList.remove('open');
  document.getElementById('sidebar-backdrop')?.classList.remove('visible');
}

// ── VIEWS ─────────────────────────────────────────────────────
function switchView(view, el) {
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  if (el) el.classList.add('active');
  document.getElementById('view-'+view).classList.add('active');
  closeSidebar(); // close on mobile nav
  // Dismiss any active undo toast when navigating — stale undo actions
  // from a previous view would be confusing if triggered from a different context.
  // Exception: if executeUndo itself is navigating back, _undoSourceView will
  // already be null (cleared by removeUndoToast) so this is a no-op.
  if (_undoTimer) removeUndoToast();
  if (view==='dashboard') renderDashboard();
  if (view==='tasks')     { populateOwnerFilter(); renderAllTasks(); }
  if (view==='calendar')  renderCalendar();
  if (view==='team') {
    const now=new Date(); const m=now.getMonth();
    const q=m<3?'Q1':m<6?'Q2':m<9?'Q3':'Q4';
    const qEl=document.getElementById('team-quarter'); if(qEl) qEl.value=q;
    const yEl=document.getElementById('team-year');    if(yEl) yEl.value=now.getFullYear();
    toggleTeamYearVisibility();
    renderTeam();
  }
  if (view==='admin')     renderAdmin();
  if (view==='mytasks') {
    const now=new Date(); const m=now.getMonth();
    const q=m<3?'Q1':m<6?'Q2':m<9?'Q3':'Q4';
    const qEl=document.getElementById('mytasks-quarter'); if(qEl&&qEl.value==='all') qEl.value=q;
    const yEl=document.getElementById('mytasks-year');    if(yEl&&!yEl.value) yEl.value=now.getFullYear();
    toggleMyTasksYearVisibility();
    renderMyTasks();
  }
  if (view==='kanban')    { initKanbanSelects(); renderKanban(); }
  if (view==='report')    { initReportSelects(); renderReport(); }
  if (view==='exec') {
    const now=new Date(); const m=now.getMonth();
    const q=m<3?'Q1':m<6?'Q2':m<9?'Q3':'Q4';
    const qEl=document.getElementById('exec-quarter'); if(qEl&&!qEl.value) qEl.value=q;
    renderExecView();
  }
}
function sortTaskTable(col) {
  if (_tableSort.col === col) {
    _tableSort.dir *= -1; // toggle direction
  } else {
    _tableSort.col = col;
    _tableSort.dir = 1;
  }
  // Update sort indicators — check both dashboard ('sort-ind-X')
  // and All Tasks ('all-sort-ind-X') spans so arrows update in both views
  ['name','owner','due','status'].forEach(c => {
    ['sort-ind-' + c, 'all-sort-ind-' + c].forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      if (c === col) el.textContent = _tableSort.dir === 1 ? '↑' : '↓';
      else el.textContent = '⇅';
    });
  });
  renderCurrentView();
}

function setQuickFilter(filter, btn) {
  // 'mine' is no longer a dashboard filter — redirect to My Tasks view
  if (filter === 'mine') {
    switchView('mytasks', document.querySelector('[data-view="mytasks"]'));
    return;
  }
  activeFilter = filter;
  document.querySelectorAll('.qfilter').forEach(b => b.classList.remove('active'));
  if (btn) btn.classList.add('active');
  renderDashboard();
}
function setTypeFilter(val) { activeTypeFilter = val; renderDashboard(); }

// ── DASHBOARD ─────────────────────────────────────────────────
function renderDashboard() {
  const q  = document.getElementById('quarter-filter')?.value || 'Q1';
  const yr = parseInt(document.getElementById('year-filter')?.value || new Date().getFullYear());
  saveQuarterPref(q, yr);
  let tasks = getActiveTasks().filter(t => t.quarter===q && t.year===yr);
  const hr  = new Date().getHours();
  const greet = document.getElementById('dashboard-greeting');
  if(greet) greet.textContent = `${hr<12?'Good morning':hr<17?'Good afternoon':'Good evening'}, ${currentUser.name.split(' ')[0]}`;
  // Update quarter label whenever filter changes
  const lbl = document.getElementById('dashboard-quarter-label');
  if (lbl) lbl.textContent = `${q} ${yr} · Financial Reporting`;
  const activeTasks = tasks.filter(t=>t.status!=='Not Applicable');
  const total    = activeTasks.length;
  const naCount  = tasks.filter(t=>t.status==='Not Applicable').length;
  const complete = activeTasks.filter(t=>t.status==='Complete').length,
        overdue  = activeTasks.filter(t=>deadlineStatus(t.dueDate,t.status)==='overdue').length,
        inprog   = activeTasks.filter(t=>t.status==='In Progress').length;
  const sg = document.getElementById('stat-grid');
  if(sg) sg.innerHTML = `
    <div class="stat-card"><div class="stat-label">Total Deliverables</div>
      <div class="stat-value">${total}</div><div class="stat-sub">${q} ${yr}</div></div>
    <div class="stat-card complete"><div class="stat-label">Complete</div>
      <div class="stat-value">${complete}</div>
      <div class="stat-sub">${total?Math.round(complete/total*100):0}% of quarter</div></div>
    <div class="stat-card progress"><div class="stat-label">In Progress</div>
      <div class="stat-value">${inprog}</div><div class="stat-sub">tasks active</div></div>
    <div class="stat-card overdue"><div class="stat-label">Overdue</div>
      <div class="stat-value">${overdue}</div>
      <div class="stat-sub">${overdue>0?'needs attention':'all on track'}</div></div>
    ${naCount>0?`<div class="stat-card na"><div class="stat-label">Not Applicable</div>
      <div class="stat-value">${naCount}</div><div class="stat-sub">excluded from stats</div></div>`:''}`;
  // Show lock banner if this quarter is locked
  const lockBanner = document.getElementById('lock-banner');
  if (lockBanner) {
    const locked = isQuarterLocked(q, yr);
    lockBanner.classList.toggle('hidden', !locked);
    if (locked) {
      const lock = _locks.find(l => l.quarter===q && l.year===yr);
      lockBanner.textContent = `🔒 ${q} ${yr} is locked (by ${lock?.lockedBy||'Admin'} on ${lock?.lockedAt||'—'}). Tasks are read-only.`;
    }
  }

  if (activeFilter==='overdue')   tasks=tasks.filter(t=>deadlineStatus(t.dueDate,t.status)==='overdue');
  if (activeFilter==='this-week') tasks=tasks.filter(t=>isThisWeek(t.dueDate));
  if (activeTypeFilter!=='all')   tasks=tasks.filter(t=>t.type===activeTypeFilter);
  renderDeadlineStrip(q, yr);
  renderPriorityCard();

  // If there are no tasks at all for any quarter (first-time setup),
  // show an onboarding prompt instead of a generic empty table.
  if (!getTasks().length && currentUser?.isAdmin) {
    const tbody = document.getElementById('dashboard-task-body');
    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="7"><div class="empty-state" style="padding:40px 0">
        <div class="empty-icon">📋</div>
        <p style="font-size:15px;font-weight:500;margin-bottom:8px">No tasks yet — let's get set up</p>
        <p style="font-size:13px;color:var(--text-faint);margin-bottom:20px">Import your task list from Excel, or add tasks one at a time.</p>
        <div style="display:flex;gap:10px;justify-content:center;flex-wrap:wrap">
          <button class="btn-primary" onclick="switchView('admin',document.querySelector('[data-view=\'admin\']'))">
            📥 Go to Import →
          </button>
          <button class="btn-secondary" onclick="openAddTask()">+ Add First Task</button>
        </div>
      </div></td></tr>`;
      return;
    }
  }

  renderTaskTable(tasks, 'dashboard-task-body', true);
}

// ── FEATURE 2 — EXTERNAL DEADLINE STRIP ──────────────────────
function renderDeadlineStrip(q, yr) {
  const el = document.getElementById('deadline-strip');
  if (!el) return;
  const qd = _quarterDates.find(d => d.quarter===q && d.year===yr);
  if (!qd) { el.innerHTML = ''; el.style.display='none'; return; }
  const today = new Date(); today.setHours(0,0,0,0);
  function daysUntil(dateStr) {
    if (!dateStr) return null;
    const d = new Date(dateStr+'T00:00:00');
    return Math.ceil((d-today)/ MS_PER_DAY);
  }
  function chip(label, dateStr, time) {
    const days = daysUntil(dateStr);
    if (days === null) return '';
    const cls = days < 0 ? 'dl-past' : days <= 3 ? 'dl-urgent' : days <= 7 ? 'dl-soon' : 'dl-ok';
    const txt = days < 0 ? `${Math.abs(days)}d ago` : days === 0 ? 'Today' : `${days}d`;
    const timeStr = time ? ` · ${time}` : '';
    return `<div class="dl-chip ${cls}"><span class="dl-label">${label}</span><span class="dl-days">${txt}${timeStr}</span></div>`;
  }
  const chips = [
    chip('SEC Filing',   qd.secFilingDate, ''),
    chip('Earnings Call',qd.earningsDate,  qd.earningsTime),
  ].filter(Boolean).join('');
  if (!chips) { el.style.display='none'; return; }
  el.style.display='flex';
  el.innerHTML = `<span style="font-size:11px;font-weight:600;color:var(--navy);margin-right:8px;white-space:nowrap">Key Dates</span>${chips}`;
}

// ── ALL TASKS ─────────────────────────────────────────────────
function populateOwnerFilter() {
  const sel = document.getElementById('all-owner-filter'); if(!sel) return;
  sel.innerHTML = `<option value="all">All Owners</option>`+
    getUsers().map(u=>`<option value="${u._spId||u.id}">${escHtml(u.name)}</option>`).join('');
}
function renderAllTasks() {
  const search  = (document.getElementById('task-search')?.value||'').toLowerCase();
  const status  = document.getElementById('all-status-filter')?.value||'all';
  const type    = document.getElementById('all-type-filter')?.value||'all';
  const owner   = document.getElementById('all-owner-filter')?.value||'all';
  const quarter = document.getElementById('all-quarter-filter')?.value||'all';
  const year    = parseInt(document.getElementById('all-year-filter')?.value||0);
  let tasks=getActiveTasks();
  if(search)         tasks=tasks.filter(t=>t.name.toLowerCase().includes(search)||(t.notes||t.description||'').toLowerCase().includes(search));
  if(status!=='all') tasks=tasks.filter(t=>t.status===status);
  if(type!=='all')   tasks=tasks.filter(t=>t.type===type);
  if(owner!=='all')  tasks=tasks.filter(t=>t.ownerId===owner);
  if(quarter!=='all') tasks=tasks.filter(t=>t.quarter===quarter && (!year||t.year===year));
  renderTaskTable(tasks,'all-task-body',false);
}

// ── TASK TABLE ────────────────────────────────────────────────
function renderTaskTable(tasks, tbodyId, hiddenQuarter) {
  const tbody = document.getElementById(tbodyId); if(!tbody) return;
  if (!tasks.length) {
    tbody.innerHTML=`<tr><td colspan="${hiddenQuarter?'7':'8'}"><div class="empty-state">
      <div class="empty-icon">📋</div>
      <p>No tasks found. <a href="#" onclick="openAddTask();return false;">Add a new task</a>.</p>
    </div></td></tr>`; return;
  }
  // Apply sort if set
  if (_tableSort.col) {
    const col = _tableSort.col, d = _tableSort.dir;
    tasks = [...tasks].sort((a,b) => {
      let av, bv;
      if (col==='due')    { av=a.dueDate||''; bv=b.dueDate||''; }
      else if(col==='status') { const o={'Not Started':0,'In Progress':1,'Ready for Review 1':2,'Ready for Review 2':3,'Complete':4,'Not Applicable':5}; av=o[a.status]??0; bv=o[b.status]??0; }
      else if(col==='owner')  { av=(getUserById(a.ownerId)?.name||'').toLowerCase(); bv=(getUserById(b.ownerId)?.name||'').toLowerCase(); }
      else if(col==='name')   { av=a.name.toLowerCase(); bv=b.name.toLowerCase(); }
      else return 0;
      return av<bv?-d:av>bv?d:0;
    });
  }
  tbody.innerHTML = tasks.map(task => {
    const owner   = getUserById(task.ownerId);
    const ds      = deadlineStatus(task.dueDate, task.status);
    const locked   = isQuarterLocked(task.quarter, task.year);
    const canEdit  = !locked && (currentUser.isAdmin || task.ownerId===currentUserId());
    const taskComments   = _comments.filter(c=>c.taskId===task.id||c.taskId===task._spId);
    const commentCount   = taskComments.length;
    const unresolvedCount = taskComments.filter(c=>!c.isResolved).length;
    const taskSteps    = getStepsForTask(task._spId);
    const doneSteps    = taskSteps.filter(s=>s.status==='Complete').length;
    const stepPct      = taskSteps.length ? Math.round(doneSteps/taskSteps.length*100) : null;
    const qCol = hiddenQuarter ? '' :
      `<td><span class="badge ${typeBadgeClass(task.type)}">${task.quarter} ${task.year}</span></td>`;
    const stepBar = taskSteps.length ? `
      <div class="step-mini-bar" onclick="openSteps('${task._spId}','${escHtml(task.name)}')" title="Steps: ${doneSteps}/${taskSteps.length} complete">
        <div class="step-mini-bar-track"><div class="step-mini-fill" style="width:${stepPct}%"></div></div>
        <span class="step-mini-label">${doneSteps}/${taskSteps.length}</span>
      </div>` : `<span class="step-mini-add" onclick="openSteps('${task._spId}','${escHtml(task.name)}')">+ add steps</span>`;
    const appBadge = task.applicability && task.applicability !== 'All Quarters'
      ? `<span class="app-badge ${task.applicability.startsWith('10-K')?'app-badge-10k':'app-badge-10q'}">${task.applicability.startsWith('10-K')?'10-K':'10-Q'}</span>`
      : '';
    const taskTrail = renderSignOffTrail(task._spId, false);
    const isNA = task.status === 'Not Applicable';
    return `<tr style="${isNA?'opacity:0.5;':''}${unresolvedCount>0?'border-left:3px solid var(--amber);':''}">
      <td style="width:32px;padding:8px 6px"><input type="checkbox" class="bulk-check" data-spid="${task._spId}" onchange="updateBulkBar()" style="cursor:pointer;width:15px;height:15px" /></td>
      <td>
        <div class="task-name"
          onclick="openSteps('${task._spId}','${escHtml(task.name)}')"
          style="cursor:pointer;" title="Click to open steps">
          ${escHtml(task.name)}${appBadge}${task.skipNextRollforward
            ? '<span class="skip-badge" title="Skipped on next rollforward">⊘ skip</span>'
            : ''}
        </div>
        ${(task.notes||task.description)?`<div class="task-desc">${escHtml(task.notes||task.description)}</div>`:''}
        ${stepBar}
        ${taskTrail}
      </td>
      <td><span class="badge ${typeBadgeClass(task.type)}">${escHtml(task.type)}</span></td>
      ${qCol}
      <td><div class="owner-chip">
        ${coloredAvatar(task.ownerId, owner?.name||'?')}
        ${owner?escHtml(owner.name):'—'}
      </div></td>
      <td><div class="deadline-cell">
        <span class="deadline-dot ${dotClass(ds)}"></span>
        ${formatWorkdayDate(task.workdayNum, task.dueDate, task.quarter, task.year)}
        ${daysLabel(task.dueDate, task.status)}
      </div></td>
      <td>
        ${locked ? '<span style="font-size:11px;color:var(--text-faint)">🔒</span> ' : ''}
        ${(() => {
          const uid = currentUserId();
          const canInteract = !isQuarterLocked(task.quarter,task.year) && (
            currentUser.isAdmin ||
            task.ownerId===uid || task.reviewerId===uid || task.reviewer2Id===uid
          );
          return `<span class="status-badge ${statusBadgeClass(task.status)}"
            onclick="${canInteract?`cycleStatus('${task._spId}')`:''}"
            oncontextmenu="event.preventDefault();${canInteract?`openQuickStatus('${task._spId}','task')`:''}"
            style="${canInteract?'cursor:pointer':''}"
            title="${!locked?'Click to advance · Right-click for all options incl. Not Applicable':''}"
          >${escHtml(task.status)}</span>`;
        })()}</td>
      <td><div class="action-row">
        <button class="icon-btn" title="Steps (${taskSteps.length})" onclick="openSteps('${task._spId}','${escHtml(task.name)}')">📋</button>
        <button class="icon-btn ${unresolvedCount>0?'comment-btn-unresolved':''}"
          title="Comments (${commentCount}${unresolvedCount>0?' — '+unresolvedCount+' unresolved':''})"
          onclick="openComments('${task._spId}','${escHtml(task.name)}')">
          💬${unresolvedCount>0
            ? `<span class="unresolved-badge">${unresolvedCount}</span>`
            : (commentCount>0 ? ` <sup style="font-size:9px">${commentCount}</sup>` : '')}
        </button>
        <button class="icon-btn" title="Quick note" onclick="openQuickNote('${task._spId}')">📝</button>
        ${canEdit?`<button class="icon-btn" title="Edit" onclick="openEditTask('${task._spId}')">✏️</button>`:''}
        ${currentUser.isAdmin?`<button class="icon-btn" title="Delete" onclick="deleteTask('${task._spId}',this)">🗑</button>`:''}
        ${task.reassignRequested && currentUser.isAdmin
          ? `<button class="icon-btn reassign-flag-btn" title="Reassignment requested${task.reassignNote?' — '+task.reassignNote:''}" onclick="clearReassignFlag('${task._spId}','task')">🔄</button>`
          : (!currentUser.isAdmin && (task.ownerId===currentUser.id||task.ownerId===currentUser._spId||task.reviewerId===currentUser.id||task.reviewer2Id===currentUser.id))
            ? `<button class="icon-btn" title="Request reassignment" onclick="openReassignRequest('${task._spId}','task','${escHtml(task.name)}')">🔄</button>`
            : ''}
      </div></td>
    </tr>`;
  }).join('');
}

// ── APPLICABILITY HELPERS ────────────────────────────────────
function appliesToQuarter(applicability, quarter) {
  if (!applicability || applicability === 'All Quarters') return true;
  if (applicability === '10-K only (Q4)')     return quarter === 'Q4';
  if (applicability === '10-Q only (Q1, Q2, Q3)') return quarter !== 'Q4';
  return true;
}

// ── SHARED ROLLFORWARD CORE ─────────────────────────────────
// Used by both rollForward() and runQuarterWizard()
// carryNotes = { [spId]: bool } — if present, wizard carry-forward behaviour is active
async function _copyTasksToQuarter(srcTasks, fromQ, fromY, toQ, toY, carryNotes) {
  const qEnd = quarterEndDate(toQ, toY);
  const srcEnd = quarterEndDate(fromQ, fromY);
  const srcEndDate = new Date(srcEnd + 'T00:00:00');
  let copied = 0;
  const errors = [];                  // collect per-task errors rather than aborting early
  const createdTaskIds = [];          // track every SP item ID created so we can roll back

  for (const task of srcTasks) {
    if (!appliesToQuarter(task.applicability, toQ)) continue;
    if (task.skipNextRollforward) {
      // Clear the skip flag on the source task regardless — it was consumed
      updateListItem(LISTS.tasks, task._spId, { SkipNextRollforward: 'No' })
        .then(() => { task.skipNextRollforward = false; })
        .catch(e => console.warn('Could not clear SkipNextRollforward:', e.message));
      continue;
    }
    // Case-insensitive duplicate check to catch minor name edits between quarters
    const nameLower = task.name.trim().toLowerCase();
    const exists = getTasks().some(t =>
      t.name.trim().toLowerCase() === nameLower && t.quarter === toQ && t.year === toY
    );
    if (exists) continue;

    try {
      let taskWdNum = task.workdayNum || null;
      if (!taskWdNum && task.dueDate) taskWdNum = dateToWorkday(task.dueDate, fromQ, fromY);
      let newDueDate = taskWdNum ? workdayToDate(taskWdNum, toQ, toY) : null;
      if (!newDueDate && task.dueDate) {
        const taskDue = new Date((task.dueDate || srcEnd) + 'T00:00:00');
        const offset  = Math.round((taskDue - srcEndDate) / MS_PER_DAY);
        newDueDate    = nearestWorkingDay(addDays(qEnd, offset));
        taskWdNum     = taskWdNum || dateToWorkday(newDueDate, toQ, toY);
      }

      // Carry-forward note handling (wizard only)
      let taskNoteValue = '';
      if (carryNotes) {
        const rawNote = carryNotes[task._spId] ? (task.notes || task.description || '') : '';
        taskNoteValue = rawNote ? `[↩ from ${fromQ} ${fromY}] ${rawNote}` : '';
      }

      const newTaskId = uid();
      const created = await createListItem(LISTS.tasks, {
        Title: task.name, TaskId: newTaskId,
        TaskType: task.type, Quarter: toQ, Year: String(toY),
        DueDate: newDueDate || '', Status: 'Not Started',
        OwnerId: task.ownerId,
        ReviewerId:  task.reviewerId  || '',
        Reviewer2Id: task.reviewer2Id || '',
        Description: taskNoteValue || task.description || '',
        Notes:       taskNoteValue || task.notes || '',
        Applicability: task.applicability || 'All Quarters',
        WorkdayNum: taskWdNum ? String(taskWdNum) : '',
        SkipNextRollforward: 'No',
      });
      const newTaskSpId = created.id; // createListItem now throws if missing — safe to access directly
      createdTaskIds.push({ list: LISTS.tasks, id: newTaskSpId });

      const srcSteps = getStepsForTask(task._spId);
      for (const step of srcSteps) {
        if (!appliesToQuarter(step.applicability, toQ)) continue;
        let stepWdNum = step.workdayNum || null;
        if (!stepWdNum && step.dueDate) stepWdNum = dateToWorkday(step.dueDate, fromQ, fromY);
        let stepDue = stepWdNum ? workdayToDate(stepWdNum, toQ, toY) : null;
        if (!stepDue && step.dueDate) {
          const sd  = new Date((step.dueDate || srcEnd) + 'T00:00:00');
          const off = Math.round((sd - srcEndDate) / MS_PER_DAY);
          stepDue   = nearestWorkingDay(addDays(qEnd, off));
          stepWdNum = stepWdNum || dateToWorkday(stepDue, toQ, toY);
        }
        let stepNoteValue = '';
        if (carryNotes) {
          const rawNote = carryNotes[step._spId] ? (step.notes || step.note || '') : '';
          stepNoteValue = rawNote ? `[↩ from ${fromQ} ${fromY}] ${rawNote}` : '';
        }
        const createdStep = await createListItem(LISTS.steps, {
          Title: step.name, StepId: uid(),
          TaskId: String(newTaskSpId),
          StepOrder: String(step.order),
          Status: 'Not Started',
          OwnerId: step.ownerId,
          ReviewerId:  step.reviewerId  || '',
          Reviewer2Id: step.reviewer2Id || '',
          DueDate: stepDue || null,
          Note:  stepNoteValue || step.note  || '',
          Notes: stepNoteValue || step.notes || '',
          Applicability: step.applicability || 'All Quarters',
          WorkdayNum: stepWdNum ? String(stepWdNum) : '',
          RequiresPrev: step.requiresPrev ? 'Yes' : 'No',
        });
        createdTaskIds.push({ list: LISTS.steps, id: createdStep.id });
      }
      copied++;
    } catch(e) {
      errors.push(`"${task.name}": ${e.message}`);
    }
  }

  // If any tasks failed to copy, roll back everything created so far and surface a
  // clear error. A partial quarter is worse than no quarter — the user can retry cleanly.
  if (errors.length) {
    console.warn(`Rollforward: ${errors.length} task(s) failed. Rolling back ${createdTaskIds.length} created records.`);
    await Promise.allSettled(
      createdTaskIds.map(({ list, id }) => deleteListItem(list, id).catch(() => {}))
    );
    throw new Error(
      `${errors.length} task(s) could not be copied and the rollforward was cancelled. ` +
      `No records were changed. Errors: ` +
      errors.slice(0, 5).join(' | ') +
      (errors.length > 5 ? ` …and ${errors.length - 5} more.` : '')
    );
  }

  return copied;
}
// ─────────────────────────────────────────────────────────────
// ── ROLLFORWARD ENGINE ────────────────────────────────────────
async function rollForward() {
  const fromQ   = document.getElementById('rf-from-quarter').value;
  const fromY   = parseInt(document.getElementById('rf-from-year').value);
  const toQ     = document.getElementById('rf-to-quarter').value;
  const toY     = parseInt(document.getElementById('rf-to-year').value);

  if (fromQ === toQ && fromY === toY) {
    showToast('Source and target quarter cannot be the same.', 'warning'); return;
  }
  if (isQuarterLocked(toQ, toY)) {
    showToast(`${toQ} ${toY} is already locked. Unlock it in Admin first.`, 'warning'); return;
  }

  const srcTasks = getTasks().filter(t => t.quarter === fromQ && t.year === fromY);
  if (!srcTasks.length) {
    showToast(`No tasks found in ${fromQ} ${fromY}.`, 'warning'); return;
  }

  // ── Step 1: Set close calendar for target quarter ────────────
  const cal = await promptCloseCalendar(toQ, toY);
  if (!cal) return; // user cancelled

  showLoadingOverlay(true, `Rolling forward to ${toQ} ${toY}…`);
  try {
    // Lock only after a clean copy — _copyTasksToQuarter throws and rolls back on failure
    const copied = await _copyTasksToQuarter(srcTasks, fromQ, fromY, toQ, toY, null);
    await lockQuarter(fromQ, fromY);
    await refreshData();
    showToast(`Rolled forward: ${copied} task(s) copied to ${toQ} ${toY}. ${fromQ} ${fromY} locked.`, 'success', TOAST_DURATION_LONG);
  } catch(e) {
    showToast('Rollforward failed: ' + e.message, 'error');
  } finally {
    showLoadingOverlay(false);
  }
}

async function lockQuarter(quarter, year) {
  const already = isQuarterLocked(quarter, year);
  if (already) return;
  const entry = {
    Title:    `${quarter} ${year}`,
    LockId:   uid(),
    Quarter:  quarter,
    Year:     String(year),
    LockedBy: currentUser.name,
    LockedAt: nowLabel(),
  };
  const created = await createListItem(LISTS.locks, entry);
  _locks.push(normaliseLock({ ...entry, id: created?.id || entry.LockId }));
}

async function unlockQuarter(lockSpId) {
  showConfirm('Unlock this quarter? Team members will be able to edit tasks again.', async () => {
    showLoadingOverlay(true, 'Unlocking quarter…');
    try {
      await deleteListItem(LISTS.locks, lockSpId);
      _locks = _locks.filter(l => l._spId !== lockSpId);
      await refreshData();
      renderAdmin();
      showToast('Quarter unlocked — team members can now edit tasks.', 'success');
    } catch(e) { showToast('Unlock failed: ' + e.message, 'error'); }
    finally { showLoadingOverlay(false); }
  }, null, 'Unlock', false);
}

// ══════════════════════════════════════════════════════════════
// ── FILE LINKS (SharePoint link attachments) ─────────────────
// Team members paste SharePoint URLs to tie-outs or any other
// file. Links are stored in FT_Attachments with a label,
// url, who linked it, and when. No file upload needed.
// ══════════════════════════════════════════════════════════════

function getAttachmentsForStep(stepSpId) {
  return _attachments
    .filter(a => a.stepId === String(stepSpId))
    .sort((a, b) => {
      // Sort newest first — uid() embeds Date.now() so id comparison is chronological
      const ai = (a.id||a._spId||'').toString();
      const bi = (b.id||b._spId||'').toString();
      return bi.localeCompare(ai);
    });
}

// Infer a tidy display label from a SharePoint URL if none given
function labelFromUrl(url) {
  try {
    const decoded = decodeURIComponent(url);
    const parts   = decoded.split('/');
    // Walk back to find first meaningful segment (skip empty, 'Forms', etc.)
    for (let i = parts.length - 1; i >= 0; i--) {
      const p = parts[i];
      if (p && p !== 'Forms' && !p.startsWith('?') && p !== 'AllItems.aspx') return p;
    }
  } catch {}
  return url;
}

function fileIcon(label) {
  const ext = (label || '').split('.').pop().toLowerCase();
  return { xlsx:'📊', xls:'📊', csv:'📊', pdf:'📄',
           docx:'📝', doc:'📝', pptx:'📑', ppt:'📑',
           zip:'🗜', msg:'📧' }[ext] || '🔗';
}

// Open the "Add link" modal for a step
function openAddLink(stepSpId) {
  document.getElementById('modal-title').textContent = 'Link SharePoint File';
  document.getElementById('modal-body').innerHTML = `
    <p style="font-size:13px;color:var(--text-muted);margin-bottom:16px">
      Paste the SharePoint link to any file — tie-out, workbook, PDF, or folder.
      To copy a link in SharePoint, right-click the file → <strong>Copy link</strong>.
    </p>
    <div class="form-group">
      <label>SharePoint URL</label>
      <input type="url" id="link-url" placeholder="https://moodys.sharepoint.com/sites/finance_home_finrptg/…"
        style="width:100%" oninput="previewLinkLabel()" />
    </div>
    <div class="form-group">
      <label>Display label <span style="font-weight:400;color:var(--text-faint)">(optional — auto-filled from URL)</span></label>
      <input type="text" id="link-label" placeholder="e.g. Q1 2025 Footnote 1 Tie Out" />
    </div>
    <div class="form-group">
      <label>Version / note <span style="font-weight:400;color:var(--text-faint)">(optional — e.g. Draft v2, Final as of Apr 3)</span></label>
      <input type="text" id="link-version" placeholder="e.g. Draft v2 as of Apr 3" />
    </div>
    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeAllModals()">Cancel</button>
      <button class="btn-primary" onclick="saveLink('${stepSpId}')">Save Link</button>
    </div>`;
  openModal();
  document.getElementById('link-url').focus();
}

function previewLinkLabel() {
  const url   = document.getElementById('link-url')?.value?.trim();
  const label = document.getElementById('link-label');
  if (label && url && !label.value) {
    label.placeholder = labelFromUrl(url);
  }
}

async function saveLink(stepSpId) {
  const url     = document.getElementById('link-url')?.value?.trim();
  const label   = document.getElementById('link-label')?.value?.trim() || labelFromUrl(url);
  const version = document.getElementById('link-version')?.value?.trim() || '';
  if (!url) { showToast('Please paste a SharePoint URL.', 'warning'); return; }
  if (!url.startsWith('http')) { showToast('Please enter a valid URL starting with https://', 'warning'); return; }

  closeAllModals();
  showLoadingOverlay(true);
  try {
    const attId = uid();
    const now   = new Date().toLocaleString('en-US', {
      month:'short', day:'numeric', year:'numeric',
      hour:'numeric', minute:'2-digit', hour12:true,
    });
    const fields = {
      Title:         label,
      AttachmentId:  attId,
      StepId:        String(stepSpId),
      TaskId:        String(stepsTaskSpId),
      Label:         label,
      FileUrl:       url,
      LinkedBy:      currentUser.name,
      LinkedAt:      now,
      VersionNote:   version,
    };
    const created = await createListItem(LISTS.attachments, fields);
    _attachments.push(normaliseAttachment({ ...fields, id: created?.id || attId }));
    renderStepsPanel();
  } catch(e) { showToast('Could not save link: ' + e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}

async function deleteAttachment(attSpId, btnEl) {
  showInlineConfirm(btnEl || document.body, 'Remove this link?', async () => {
    try {
      _attachments = _attachments.filter(a => a._spId !== attSpId); // optimistic
      renderStepsPanel();
      await deleteListItem(LISTS.attachments, attSpId);
    } catch(e) {
      showToast('Could not remove link: ' + e.message, 'error');
      await refreshData();
    }
  }, 'Remove', true);
}

function renderAttachmentPanel(stepSpId) {
  const atts    = getAttachmentsForStep(stepSpId);
  const step    = _steps.find(s => s._spId === stepSpId);
  const isOwner = currentUser.isAdmin
    || step?.ownerId === currentUser.id
    || step?.ownerId === currentUser._spId;

  const attRows = atts.map(a => {
    const display = a.label || labelFromUrl(a.url);
    return `<div class="att-row">
      <span class="att-icon">${fileIcon(display)}</span>
      <div class="att-info">
        <a href="${escHtml(a.url)}" target="_blank" rel="noopener" class="att-name"
           title="${escHtml(a.url)}">${escHtml(display)}</a>
        ${a.versionNote ? `<div class="att-version">${escHtml(a.versionNote)}</div>` : ''}
        <div class="att-meta">Linked by ${escHtml(a.linkedBy)} · ${escHtml(a.linkedAt)}</div>
      </div>
      ${isOwner || currentUser.isAdmin
        ? `<button class="icon-btn" style="width:24px;height:24px;font-size:11px"
             title="Remove link" onclick="deleteAttachment('${a._spId}',this)">✕</button>`
        : ''}
    </div>`;
  }).join('');

  return `
    <div class="att-panel">
      <div class="att-panel-header">
        <span class="att-panel-title">🔗 Links${atts.length ? ` (${atts.length})` : ''}</span>
        ${isOwner
          ? `<button class="att-upload-btn" onclick="openAddLink('${stepSpId}')">+ Add link</button>`
          : ''}
      </div>
      ${atts.length
        ? `<div class="att-list">${attRows}</div>`
        : `<div class="att-empty">No links yet${isOwner ? ' — click "+ Add link" to attach a SharePoint file.' : '.'}</div>`
      }
    </div>`;
}

// ── SIGN-OFF LOG ─────────────────────────────────────────────

function nowISO()  { return new Date().toISOString(); }
function nowLabel() {
  return new Date().toLocaleString('en-US', {
    month: 'short', day: 'numeric', year: 'numeric',
    hour: 'numeric', minute: '2-digit', hour12: true,
  });
}

async function writeSignOff(refId, refType, refName, fromStatus, toStatus) {
  const entry = {
    SignOffId:     uid(),
    RefId:         String(refId),
    RefType:       refType,
    RefName:       refName,
    UserId:        currentUserId(),
    UserName:      currentUser.name,
    FromStatus:    fromStatus,
    ToStatus:      toStatus,
    Timestamp:     nowLabel(),
    TimestampISO:  nowISO(),
    Title:         `${currentUser.name} → ${toStatus}`,
  };
  // Add to local cache immediately so it renders inline without waiting
  _signOffs.push(normaliseSignOff({ ...entry, id: entry.SignOffId }));
  // Persist to SharePoint in background (non-blocking)
  createListItem(LISTS.signOffs, entry).catch(e =>
    console.warn('Sign-off write failed:', e)
  );
}

function getSignOffsFor(refId) {
  return _signOffs
    .filter(s => s.refId === String(refId))
    .sort((a, b) => { const at=a.tsIso||a.ts||''; const bt=b.tsIso||b.ts||''; return bt>at?1:bt<at?-1:0; });
}

// Renders the inline sign-off trail (most recent first, max 5 shown)
function renderSignOffTrail(refId, showAll) {
  const entries = getSignOffsFor(refId);
  if (!entries.length) return '';
  const visible  = showAll ? entries : entries.slice(0, 3);
  const overflow = entries.length - visible.length;
  const rows = visible.map(e => {
    const icon = {
      'Complete':           '✅',
      'Ready for Review 1': '🔍',
      'Ready for Review 2': '🔎',
      'In Progress':        '▶️',
      'Not Started':        '⏸',
      'Not Applicable':     '⊘',
    }[e.toStatus] || '•';
    return `<div class="signoff-entry">
      <span class="signoff-icon">${icon}</span>
      <span class="signoff-status">${escHtml(e.toStatus)}</span>
      <span class="signoff-by">${escHtml(e.userName)}</span>
      <span class="signoff-time">${escHtml(e.ts)}</span>
    </div>`;
  }).join('');
  const more = overflow > 0
    ? `<span class="signoff-more" onclick="toggleSignOffHistory('${refId}', this)">+ ${overflow} earlier entries</span>`
    : '';
  return `<div class="signoff-trail" id="trail-${refId}">${rows}${more}</div>`;
}

function toggleSignOffHistory(refId, el) {
  const trail = document.getElementById('trail-' + refId);
  if (!trail) return;
  trail.outerHTML = renderSignOffTrail(refId, true);
}

// ── STATUS CYCLE ──────────────────────────────────────────────
const STATUS_ORDER = ['Not Started','In Progress','Ready for Review 1','Ready for Review 2','Complete','Not Applicable'];
// N/A is a manual override — never in the cycle
// Cycle is dynamic based on reviewer assignments — see getStatusCycle()
// STATUS_CYCLE_FULL was redundant with STATUS_ORDER.filter(s=>s!=='Not Applicable') — removed

// Returns the applicable status cycle for a task or step object
// Skips Review 1 / Review 2 if no reviewer is assigned
function getStatusCycle(item) {
  const cycle = ['Not Started','In Progress'];
  if (item.reviewerId)  cycle.push('Ready for Review 1');
  if (item.reviewer2Id) cycle.push('Ready for Review 2');
  cycle.push('Complete');
  return cycle;
}

// ── FEATURE 4 — EMAIL NOTIFICATION (via FT_Notifications list) ───
// Writes a record to FT_Notifications. A Power Automate flow watches
// this list ("When an item is created") and sends the email via Outlook.
// No Mail.Send permission required — PA handles delivery.
async function notifyReviewer(item, itemType, newStatus) {
  try {
    let reviewerEmail = '';
    if (newStatus === 'Ready for Review 1' && item.reviewerId) {
      const reviewer = _users.find(u => (u._spId||u.id) === item.reviewerId);
      reviewerEmail = reviewer?.email || '';
    } else if (newStatus === 'Ready for Review 2' && item.reviewer2Id) {
      const reviewer2 = _users.find(u => (u._spId||u.id) === item.reviewer2Id);
      reviewerEmail = reviewer2?.email || '';
    }
    if (!reviewerEmail) return;

    const taskName = itemType === 'task'
      ? item.name
      : (getTasks().find(t => t._spId === item.taskId)?.name || '');
    const stepName = itemType === 'step' ? item.name : '';

    await createListItem(LISTS.notifications, {
      Title:         `[FRT] ${stepName ? 'Step' : 'Task'} ready for review: ${stepName || taskName}`,
      ToEmail:       reviewerEmail,
      TaskName:      taskName,
      StepName:      stepName,
      StatusChanged: newStatus,
      ChangedBy:     currentUser?.name || '',
      AppUrl:        window.location.href,
      NotifyType:    'review',
    });
  } catch(e) { console.warn('Notification write failed (non-critical):', e.message); }
}

// Tracks task spIds currently being written to prevent concurrent undo/redo races.
const _statusWriteInFlight = new Set();

async function cycleStatus(spId) {
  const task = _tasks.find(t => t._spId === spId); if(!task) return;
  if (isQuarterLocked(task.quarter, task.year)) {
    showToast(`${task.quarter} ${task.year} is locked — unlock it in Admin to make changes.`, 'warning'); return;
  }
  // Debounce: ignore rapid double-clicks while a write is in flight for this task
  if (_statusWriteInFlight.has(spId)) return;

  const prev = task.status;
  const cycle = getStatusCycle(task);
  const next = cycle[(cycle.indexOf(prev)+1) % cycle.length];
  task.status = next;
  renderCurrentView();
  showUndoToast(`Status changed to "${next}"`, async () => {
    if (_statusWriteInFlight.has(spId)) return; // forward write still in flight — skip
    task.status = prev;
    renderCurrentView();
    _statusWriteInFlight.add(spId);
    try {
      await updateListItem(LISTS.tasks, spId, { Status: prev });
      await writeSignOff(spId, 'task', task.name, next, prev);
    } finally { _statusWriteInFlight.delete(spId); }
  });
  _statusWriteInFlight.add(spId);
  try {
    await updateListItem(LISTS.tasks, spId, { Status: next });
    await writeSignOff(spId, 'task', task.name, prev, next);
    if (next === 'Ready for Review 1' || next === 'Ready for Review 2') await notifyReviewer(task, 'task', next);
  } catch(e) {
    console.error("Status update failed:", e);
    showToast('Status update failed — changes were not saved: ' + e.message, 'error');
    await refreshData();
  } finally { _statusWriteInFlight.delete(spId); }
}

// ── STEPS PANEL ──────────────────────────────────────────────
let stepsTaskSpId      = null;
let editingStepId      = null;
let _stepsPanelDirty   = false; // track if any changes made

function openSteps(taskSpId, taskName) {
  stepsTaskSpId = taskSpId;
  document.getElementById('steps-task-title').textContent = taskName;
  renderStepsPanel();
  document.getElementById('steps-overlay').classList.remove('hidden');
}

// Closes the steps panel. Triggers a full data refresh if any steps were modified.
function closeStepsPanel() {
  document.getElementById('steps-overlay').classList.add('hidden');
  stepsTaskSpId = null;
  if (_stepsPanelDirty) { _stepsPanelDirty = false; refreshData(); }
  else { renderCurrentView(); } // just re-render without full reload
}

function renderStepsPanel() {
  const steps = getStepsForTask(stepsTaskSpId);
  const list  = document.getElementById('steps-list');
  if (!list) return;
  // Show/hide Add Step button based on lock state
  const parentTask   = _tasks.find(t => t._spId === stepsTaskSpId);
  const isLocked     = parentTask ? isQuarterLocked(parentTask.quarter, parentTask.year) : false;
  const footerEl     = document.getElementById('steps-panel-footer');
  if (footerEl) {
    footerEl.innerHTML = (!isLocked || currentUser.isAdmin)
      ? '<button class="btn-primary small" onclick="openAddStep()">+ Add Step</button>'
      : '<span style="font-size:11px;color:var(--text-faint)">🔒 Quarter locked — steps are read-only</span>';
  }

  if (!steps.length) {
    list.innerHTML = `<div class="empty-state" style="padding:24px 0">
      <div class="empty-icon">📝</div>
      <p>No steps yet. Add the first step below.</p>
    </div>`;
    return;
  }

  const activeSteps = steps.filter(s => s.status !== 'Not Applicable');
  const total = activeSteps.length;
  const done  = activeSteps.filter(s => s.status === 'Complete').length;
  const pct   = total ? Math.round(done/total*100) : 0;

  // ── Progress header + individual step rows ──────────────────
  list.innerHTML = `
    <div class="steps-progress-header">
      <span class="steps-progress-label">${done} of ${total} complete</span>
      <span class="steps-progress-pct">${pct}%</span>
    </div>
    <div class="progress-bar-wrap" style="margin-bottom:16px">
      <div class="progress-bar-fill" style="width:${pct}%;background:var(--blue)"></div>
    </div>
    ${steps.map((step, idx) => {
      const owner    = getUserById(step.ownerId);
      const reviewer  = step.reviewerId  ? getUserById(step.reviewerId)  : null;
      const reviewer2 = step.reviewer2Id ? getUserById(step.reviewer2Id) : null;
      const ds      = step.dueDate ? deadlineStatus(step.dueDate, step.status) : 'ok';
      const isOwner = currentUser.isAdmin || step.ownerId === currentUserId();
      const isDone  = step.status === 'Complete';
      const stepTrail    = renderSignOffTrail(step._spId, false);
      const stepAtts     = renderAttachmentPanel(step._spId);
      const stepComments = _comments.filter(c => c.stepId === step._spId);
      const stepCommentCount = stepComments.length;
      return `<div class="step-row ${isDone ? 'step-done' : ''}"
        draggable="true"
        ondragstart="stepDragStart(event,'${step._spId}')"
        ondragend="stepDragEnd(event)"
        ondragover="stepDragOver(event,'${step._spId}')"
        ondrop="stepDrop(event,'${step._spId}')">
        <div class="step-number step-drag-handle" title="Drag to reorder">⠿ ${idx+1}</div>
        <div class="step-body">
          <div class="step-name ${isDone ? 'step-name-done' : ''}">${escHtml(step.name)}</div>
          <div class="step-meta">
            <span class="role-label">Preparer:</span>
            ${owner ? `<span class="owner-chip">${coloredAvatar(step.ownerId,owner.name,22)}${escHtml(owner.name)}</span>` : '<span class="text-muted">Unassigned</span>'}
            ${reviewer
              ? `<span class="role-label" style="margin-left:8px">Reviewer 1:</span>
                 <span class="owner-chip reviewer-chip">${coloredAvatar(step.reviewerId,reviewer.name,22)}${escHtml(reviewer.name)}</span>`
              : ''}
            ${reviewer2
              ? `<span class="role-label" style="margin-left:8px">Reviewer 2:</span>
                 <span class="owner-chip reviewer2-chip">${coloredAvatar(step.reviewer2Id,reviewer2.name,22)}${escHtml(reviewer2.name)}</span>`
              : ''}
            ${step.dueDate || step.workdayNum
              ? `<span class="deadline-cell">
                   <span class="deadline-dot ${dotClass(ds)}"></span>
                   ${formatWorkdayDate(step.workdayNum, step.dueDate, parentTask?.quarter, parentTask?.year)}
                 </span>`
              : ''}
            ${(() => {
              const noteText = step.notes || step.note;
              if (!noteText) return '';
              if (noteText.startsWith('[↩')) {
                const label   = noteText.split(']')[0] + ']';
                const content = noteText.split('] ').slice(1).join('] ');
                return `<div class="step-notes-display">
                  <span class="note-carryforward-label">${escHtml(label)}</span>
                  ${escHtml(content)}
                </div>`;
              }
              return `<div class="step-notes-display">${escHtml(noteText)}</div>`;
            })()}
          </div>
          ${stepTrail}
          ${stepAtts}
          <div class="step-comments-section">
            <button class="step-comment-toggle" onclick="toggleStepComments('${step._spId}')">
              💬 ${stepCommentCount > 0 ? `${stepCommentCount} comment${stepCommentCount>1?'s':''}` : 'Add comment'}
            </button>
            <div id="step-comments-${step._spId}" class="step-comments-panel hidden">
              ${stepComments.map(c => {
                const a = getUserById(c.authorId);
                return `<div class="comment-item">
                  <span class="comment-author">${escHtml(a ? a.name : 'Former team member')}</span>
                  <span class="comment-time">${escHtml(c.time)}</span>
                  <div class="comment-text">${escHtml(c.text)}</div>
                </div>`;
              }).join('')}
              <div class="step-comment-input">
                <input type="text"
                  id="step-comment-input-${step._spId}"
                  placeholder="Add a comment…"
                  style="width:100%;padding:6px 8px;border:1px solid var(--border);border-radius:6px;font-size:12px"
                  onkeydown="if(event.key==='Enter'&&!event.shiftKey)addStepComment('${step._spId}')" />
                <button class="btn-primary" style="margin-top:6px;font-size:12px;padding:5px 12px" onclick="addStepComment('${step._spId}')">Post</button>
              </div>
            </div>
          </div>
        </div>
        <div class="step-actions">
          ${(() => {
            const blocker    = getBlockingPredecessor(step);
            const isGated    = !!blocker && !isDone;
            const gateLabel  = isGated ? `🔒 Needs: ${escHtml(blocker.name)}` : '';
            const badgeTitle = isGated ? `Gated — "${blocker.name}" must be Complete first` : '';
            if (!isOwner) {
              return `<span class="status-badge ${statusBadgeClass(step.status)}" style="font-size:11px">${escHtml(step.status)}</span>`;
            }
            return `
              <div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px">
                <span class="status-badge ${statusBadgeClass(step.status)} ${isGated?'step-gated':''}"
                  style="cursor:pointer;font-size:11px" title="${badgeTitle}"
                  onclick="cycleStepStatus('${step._spId}')">
                  ${isGated ? '🔒 ' : ''}${escHtml(step.status)}
                </span>
                ${isGated && currentUser.isAdmin ? `<button class="icon-btn step-force-btn" title="Admin: force unlock" onclick="forceUnlockStep('${step._spId}')">⚡ Force</button>` : ''}
                ${isGated ? `<span class="gate-label">${gateLabel}</span>` : ''}
              </div>
              <button class="icon-btn" onclick="openEditStep('${step._spId}')">✏️</button>`;
          })()}
          ${step.reassignRequested && currentUser.isAdmin
            ? `<button class="icon-btn reassign-flag-btn" title="Reassignment requested${step.reassignNote?' — '+step.reassignNote:''}" onclick="clearReassignFlag('${step._spId}','step')">🔄</button>`
            : (!currentUser.isAdmin && (step.ownerId===currentUser.id||step.ownerId===currentUser._spId||step.reviewerId===currentUser.id||step.reviewer2Id===currentUser.id))
              ? `<button class="icon-btn" title="Request reassignment" onclick="openReassignRequest('${step._spId}','step','${escHtml(step.name)}')">🔄</button>`
              : ''}
          ${currentUser.isAdmin ? `<button class="icon-btn" onclick="deleteStep('${step._spId}',this)">🗑</button>` : ''}
        </div>
      </div>`;
    }).join('')}`;
}

// ── STEP GATE HELPER ─────────────────────────────────────────
// Returns the predecessor step (same task, order = this.order - 1) if
// this step has requiresPrev = true and that predecessor is not Complete.
// Returns null if the gate is clear or doesn't apply.
function getBlockingPredecessor(step) {
  if (!step.requiresPrev) return null;
  const siblings = getStepsForTask(step.taskId)
    .sort((a,b) => (a.order||0) - (b.order||0));
  const idx = siblings.findIndex(s => s._spId === step._spId);
  if (idx <= 0) return null;                        // first step — no predecessor
  const prev = siblings[idx - 1];
  if (prev.status === 'Complete') return null;       // gate is clear
  return prev;                                       // gate is blocked
}

async function cycleStepStatus(stepSpId) {
  const step = _steps.find(s => s._spId === stepSpId); if(!step) return;
  const prev = step.status;
  const cycle = getStatusCycle(step);
  const next = cycle[(cycle.indexOf(prev)+1) % cycle.length];

  try {
    // Only check gate when advancing forward (not cycling back to Not Started)
    const movingForward = STATUS_ORDER.indexOf(next) > STATUS_ORDER.indexOf(prev);
    if (movingForward) {
      const blocker = getBlockingPredecessor(step);
      if (blocker) {
        await new Promise((resolve, reject) => {
          showConfirm(
            `"${blocker.name}" (step ${blocker.order}) is not yet complete. This step requires the previous step to be finished first. Proceed anyway?`,
            resolve, reject, 'Proceed', true
          );
        }).catch(() => { throw new Error('cancelled'); });
      }
    }

    step.status = next;
    _stepsPanelDirty = true;
    renderStepsPanel();
    showUndoToast(`Step status changed to "${next}"`, async () => {
      step.status = prev;
      renderStepsPanel();
      await updateListItem(LISTS.steps, stepSpId, { Status: prev });
      await writeSignOff(stepSpId, 'step', step.name, next, prev);
    });
    await updateListItem(LISTS.steps, stepSpId, { Status: next });
    await writeSignOff(stepSpId, 'step', step.name, prev, next);
    if (next === 'Ready for Review 1' || next === 'Ready for Review 2') await notifyReviewer(step, 'step', next);
  } catch(e) {
    if(e.message==='cancelled'){step.status=prev;_stepsPanelDirty=false;renderStepsPanel();return;}
    console.error("Step status update failed:", e); await refreshData();
  }
}

// Admin force-unlock: mark the blocking predecessor Complete, then advance this step
async function forceUnlockStep(stepSpId) {
  const step    = _steps.find(s => s._spId === stepSpId); if(!step) return;
  const blocker = getBlockingPredecessor(step);
  if (!blocker) { cycleStepStatus(stepSpId); return; }
  showConfirm(
    `Force unlock? This will mark "${blocker.name}" as Complete (bypassing the gate) and then advance "${step.name}". A sign-off entry will be created for both changes.`,
    async () => {
      const blockerPrev = blocker.status;
      blocker.status = 'Complete';
      await writeSignOff(blocker._spId, 'step', blocker.name, blockerPrev, 'Complete');
      await updateListItem(LISTS.steps, blocker._spId, { Status: 'Complete' });
      await cycleStepStatus(stepSpId);
    },
    null, 'Force Unlock', true
  );
}

function openAddStep() {
  editingStepId = null;
  showStepModal(null);
}

function openEditStep(stepSpId) {
  editingStepId = stepSpId;
  showStepModal(_steps.find(s => s._spId === stepSpId));
}

function showStepModal(step) {
  const steps     = getStepsForTask(stepsTaskSpId);
  const _uid = u => u._spId||u.id;
  const ownerOpts = getUsers().map(u =>
    `<option value="${_uid(u)}" ${step?.ownerId===_uid(u)?'selected':''}>${escHtml(u.name)}</option>`).join('');
  const srOpts  = `<option value="">— None —</option>` + getUsers().map(u=>`<option value="${_uid(u)}" ${step?.reviewerId===_uid(u)?'selected':''}>${escHtml(u.name)}</option>`).join('');
  const sr2Opts = `<option value="">— None —</option>` + getUsers().map(u=>`<option value="${_uid(u)}" ${step?.reviewer2Id===_uid(u)?'selected':''}>${escHtml(u.name)}</option>`).join('');
  document.getElementById('modal-title').textContent = step ? 'Edit Step' : 'Add Step';
  document.getElementById('modal-body').innerHTML = `
    <div class="form-group"><label>Step Name</label>
      <input type="text" id="sf-name" value="${escHtml(step?.name||'')}" placeholder="e.g. Rollforward" /></div>
    <div class="form-group"><label>Order</label>
      <input type="number" id="sf-order" value="${step?.order ?? steps.length+1}" min="1" /></div>
    <div class="form-group"><label>Preparer (Owner)</label>
      <select id="sf-owner"><option value="">— Unassigned —</option>${ownerOpts}</select></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="form-group"><label>Reviewer 1 <span style="font-weight:400;color:var(--text-faint)">(optional)</span></label>
        <select id="sf-reviewer">${srOpts}</select></div>
      <div class="form-group"><label>Reviewer 2 <span style="font-weight:400;color:var(--text-faint)">(optional)</span></label>
        <select id="sf-reviewer2">${sr2Opts}</select></div>
    </div>
    <div class="form-group"><label>Due Date</label>
      <div style="display:grid;grid-template-columns:90px 1fr;gap:10px;align-items:end">
        <div>
          <label style="font-size:11px;color:var(--text-faint);display:block;margin-bottom:4px">Workday #</label>
          <input type="number" id="sf-workday"
            value="${step?.workdayNum||''}" min="1" max="60" placeholder="WD #"
            style="width:100%;padding:10px 8px;border:1.5px solid var(--border);border-radius:var(--radius);font-family:inherit;font-size:14px"
            oninput="(()=>{const pt=_tasks.find(t=>t._spId===stepsTaskSpId);if(pt)resolveWdToDate('sf-workday','sf-due',pt.quarter,pt.year)})()" />
        </div>
        <input type="date" id="sf-due" value="${step?.dueDate||''}" />
      </div>
    </div>
    <div class="form-group"><label>Status</label>
      <select id="sf-status">${STATUS_ORDER.map(s=>`<option ${step?.status===s?'selected':''}>${s}</option>`).join('')}</select></div>
    <div class="form-group"><label>Notes <span style="font-weight:400;color:var(--text-faint)">(optional — carries forward to next quarter if selected during rollforward)</span></label>
      <textarea id="sf-notes" rows="3"
        placeholder="e.g. Use tab 'FN1 RF' in the Master SS. Check rate changed to 4.8% per Q3 auditor note."
        style="width:100%;resize:vertical">${escHtml(step?.notes||step?.note||'')}</textarea></div>
    <div class="form-group"><label>Applicability</label>
      <select id="sf-applicability">
        <option value="All Quarters" ${(!step?.applicability||step?.applicability==='All Quarters')?'selected':''}>All Quarters</option>
        <option value="10-K only (Q4)" ${step?.applicability==='10-K only (Q4)'?'selected':''}>10-K only (Q4)</option>
        <option value="10-Q only (Q1, Q2, Q3)" ${step?.applicability==='10-Q only (Q1, Q2, Q3)'?'selected':''}>10-Q only (Q1, Q2, Q3)</option>
      </select></div>
    <div class="form-group">
      <label style="display:flex;align-items:center;gap:10px;cursor:pointer;user-select:none">
        <input type="checkbox" id="sf-requires-prev" ${step?.requiresPrev?'checked':''} style="width:16px;height:16px;cursor:pointer" />
        <span>Requires previous step to be Complete before this step can advance</span>
      </label>
      <p style="font-size:11px;color:var(--text-faint);margin-top:4px;margin-left:26px">
        When checked, a warning appears if someone tries to advance this step before
        the one above it is finished. Admins can force-override.
      </p>
    </div>
    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeAllModals()">Cancel</button>
      <button class="btn-primary" onclick="saveStep()">Save Step</button>
    </div>`;
  openModal();
}

async function saveStep() {
  const name          = document.getElementById('sf-name').value.trim();
  const order         = parseInt(document.getElementById('sf-order').value) || 1;
  const ownerId       = document.getElementById('sf-owner').value;
  const dueDate       = document.getElementById('sf-due').value;
  const status        = document.getElementById('sf-status').value;
  const notes         = document.getElementById('sf-notes')?.value.trim() || '';
  const applicability = document.getElementById('sf-applicability').value;
  const workdayNum    = parseInt(document.getElementById('sf-workday')?.value) || null;
  const requiresPrev  = document.getElementById('sf-requires-prev')?.checked || false;
  const reviewerId   = document.getElementById('sf-reviewer')?.value  || '';
  const reviewer2Id  = document.getElementById('sf-reviewer2')?.value || '';
  if (!name) { showToast('Please enter a step name.', 'warning'); return; }
  // If workday set and a calendar exists for this task's quarter, resolve date
  let resolvedDue = dueDate;
  if (workdayNum && stepsTaskSpId) {
    const parentTask = _tasks.find(t => t._spId === stepsTaskSpId);
    if (parentTask) {
      const wd = workdayToDate(workdayNum, parentTask.quarter, parentTask.year);
      if (wd) resolvedDue = wd;
    }
  }
  closeAllModals();
  showLoadingOverlay(true);
  try {
    const fields = {
      Title: name, StepOrder: String(order), OwnerId: ownerId,
      ReviewerId: reviewerId, Reviewer2Id: reviewer2Id,
      DueDate: resolvedDue || null, Status: status,
      Note: notes, Notes: notes,
      TaskId: stepsTaskSpId, Applicability: applicability,
      WorkdayNum: workdayNum ? String(workdayNum) : '',
      RequiresPrev: requiresPrev ? 'Yes' : 'No',
    };
    if (editingStepId) {
      const oldStep = _steps.find(s=>s._spId===editingStepId);
      if (oldStep && oldStep.dueDate !== (resolvedDue||null)) {
        await recordDueDateChange(editingStepId, 'step', name, oldStep.dueDate, resolvedDue||null);
      }
      await updateListItem(LISTS.steps, editingStepId, fields);
      const idx = _steps.findIndex(s => s._spId === editingStepId);
      if (idx >= 0) Object.assign(_steps[idx], { name, order, ownerId, reviewerId, reviewer2Id, dueDate: resolvedDue, status, note: notes, notes, applicability, workdayNum, requiresPrev });
    } else {
      fields.StepId = uid();
      const created = await createListItem(LISTS.steps, fields);
      _steps.push(normaliseStep({ ...fields, id: created?.id || fields.StepId }));
    }
    renderStepsPanel();
  } catch(e) { showToast('Could not save step: '+e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}

async function deleteStep(stepSpId, btnEl) {
  showInlineConfirm(btnEl || document.body, 'Delete this step?', async () => {
    _steps = _steps.filter(s => s._spId !== stepSpId); // optimistic
    renderStepsPanel();
    try {
      await deleteListItem(LISTS.steps, stepSpId);
      _stepsPanelDirty = true;
    } catch(e) {
      showToast('Could not delete step: '+e.message, 'error');
      await refreshData(); renderStepsPanel();
    }
  }, 'Delete', true);
}

// ── STEP TEMPLATES — manage default steps per task template ───
let editingStepTemplateId  = null;
let stepsTemplateId        = null;

function openStepTemplates(templateSpId, templateName) {
  stepsTemplateId   = templateSpId;
  document.getElementById('modal-title').textContent = `Steps for "${templateName}"`;
  renderStepTemplatesModal();
  openModal();
}

function renderStepTemplatesModal() {
  const tpSteps = getStepTemplatesForTemplate(stepsTemplateId);
  const ownerOpts = () => getUsers().map(u =>
    `<option value="${u._spId||u.id}">${escHtml(u.name)}</option>`).join('');

  document.getElementById('modal-body').innerHTML = `
    <div id="step-tpl-list">
      ${tpSteps.length === 0
        ? '<p class="text-muted" style="font-size:13px;margin-bottom:12px">No default steps yet.</p>'
        : tpSteps.map((st,i) => {
            const owner = getUserById(st.defaultOwnerId);
            return `<div class="step-tpl-row">
              <span class="step-number">${i+1}</span>
              <div class="step-body">
                <div class="step-name">
                  ${escHtml(st.name)}
                  ${st.requiresPrev
                    ? '<span title="Requires previous step" style="font-size:10px;background:#fef3c7;color:#92400e;border-radius:4px;padding:1px 5px">🔒 gated</span>'
                    : ''}
                </div>
                <div class="step-meta">
                  ${owner?`<span class="owner-chip"><span class="mini-avatar">${initials(owner.name)}</span>${escHtml(owner.name)}</span>`:''}
                  <span class="text-muted">${st.dueDaysFromQtrEnd} days after qtr end</span>
                </div>
              </div>
              <div class="step-actions">
                <button class="icon-btn" onclick="openEditStepTemplate('${st._spId}')">✏️</button>
                <button class="icon-btn" onclick="deleteStepTemplate('${st._spId}',this)">🗑</button>
              </div>
            </div>`;
          }).join('')
      }
    </div>
    <div style="border-top:1px solid var(--bg-secondary);margin-top:12px;padding-top:14px">
      <p style="font-size:12px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px">Add default step</p>
      <div class="form-group"><label>Step Name</label>
        <input type="text" id="stf-name" placeholder="e.g. Rollforward" /></div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div class="form-group"><label>Order</label>
          <input type="number" id="stf-order" value="${tpSteps.length+1}" min="1" /></div>
        <div class="form-group"><label>Days After Qtr End</label>
          <input type="number" id="stf-days" value="5" min="0" max="120" /></div>
      </div>
      <div class="form-group"><label>Default Owner</label>
        <select id="stf-owner"><option value="">— Unassigned —</option>${ownerOpts()}</select></div>
      <div class="form-group">
        <label style="display:flex;align-items:center;gap:10px;cursor:pointer;user-select:none">
          <input type="checkbox" id="stf-requires-prev" style="width:16px;height:16px;cursor:pointer" />
          <span>Requires previous step to be Complete before advancing</span>
        </label>
      </div>
      <div class="modal-footer">
        <button class="btn-secondary" onclick="closeAllModals()">Done</button>
        <button class="btn-primary" onclick="saveStepTemplate()">Add Step</button>
      </div>
    </div>`;
}

function openEditStepTemplate(spId) {
  const st = _stepTemplates.find(s => s._spId === spId); if (!st) return;
  editingStepTemplateId = spId;
  const nameEl  = document.getElementById('stf-name');
  const orderEl = document.getElementById('stf-order');
  const daysEl  = document.getElementById('stf-days');
  const ownerEl = document.getElementById('stf-owner');
  const reqEl   = document.getElementById('stf-requires-prev');
  if (nameEl)  nameEl.value  = st.name;
  if (orderEl) orderEl.value = st.order;
  if (daysEl)  daysEl.value  = st.dueDaysFromQtrEnd;
  if (ownerEl) ownerEl.value = st.defaultOwnerId || '';
  if (reqEl)   reqEl.checked = st.requiresPrev || false;
  const btn = document.querySelector('#modal-body .btn-primary[onclick="saveStepTemplate()"]');
  if (btn) btn.textContent = 'Save Changes';
  nameEl?.focus();
}

async function saveStepTemplate() {
  const name       = document.getElementById('stf-name').value.trim();
  const order      = parseInt(document.getElementById('stf-order').value) || 1;
  const days       = parseInt(document.getElementById('stf-days').value) || 0;
  const ownerId    = document.getElementById('stf-owner').value;
  const reqPrev    = document.getElementById('stf-requires-prev')?.checked || false;
  if (!name) { showToast('Please enter a step name.', 'warning'); return; }
  const fields = {
    Title: name, StepOrder: String(order),
    DueDaysFromQtrEnd: String(days),
    DefaultOwnerId: ownerId,
    RequiresPrev: reqPrev ? 'Yes' : 'No',
  };
  showLoadingOverlay(true);
  try {
    if (editingStepTemplateId) {
      await updateListItem(LISTS.stepTemplates, editingStepTemplateId, fields);
      const idx = _stepTemplates.findIndex(s => s._spId === editingStepTemplateId);
      if (idx >= 0) Object.assign(_stepTemplates[idx], { name, order, dueDaysFromQtrEnd: days, defaultOwnerId: ownerId, requiresPrev: reqPrev });
      editingStepTemplateId = null;
    } else {
      const id = uid();
      const created = await createListItem(LISTS.stepTemplates, { ...fields, StepTemplateId: id, TemplateId: stepsTemplateId });
      _stepTemplates.push(normaliseStepTemplate({ ...fields, StepTemplateId: id, TemplateId: stepsTemplateId, id: created?.id || id }));
    }
    renderStepTemplatesModal();
  } catch(e) { showToast('Could not save step template: '+e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}

async function deleteStepTemplate(spId, btnEl) {
  showInlineConfirm(btnEl || document.body, 'Delete this default step?', async () => {
    showLoadingOverlay(true, 'Deleting…');
    try {
      await deleteListItem(LISTS.stepTemplates, spId);
      _stepTemplates = _stepTemplates.filter(s => s._spId !== spId);
      renderStepTemplatesModal();
      showToast('Default step deleted.', 'success');
    } catch(e) { showToast('Could not delete step template: '+e.message, 'error'); }
    finally { showLoadingOverlay(false); }
  }, 'Delete', true);
}

// ── CALENDAR ─────────────────────────────────────────────────
function changeCalMonth(dir) {
  calMonth+=dir;
  if(calMonth>11){calMonth=0;calYear++;}
  if(calMonth<0) {calMonth=11;calYear--;}
  renderCalendar();
}
function jumpCalToToday() {
  const now = new Date();
  calYear  = now.getFullYear();
  calMonth = now.getMonth();
  renderCalendar();
}
function renderCalendar() {
  const grid=document.getElementById('calendar-grid'); if(!grid) return;
  const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
  const lbl=document.getElementById('cal-month-label'); if(lbl) lbl.textContent=`${MONTHS[calMonth]} ${calYear}`;
  const firstDay=new Date(calYear,calMonth,1).getDay();
  const daysInMonth=new Date(calYear,calMonth+1,0).getDate();
  const today=new Date(); today.setHours(0,0,0,0);
  let html=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'].map(d=>`<div class="cal-header">${d}</div>`).join('');
  for(let i=0;i<firstDay;i++){
    const pd=new Date(calYear,calMonth,-firstDay+i+1);
    html+=`<div class="cal-day other-month"><div class="cal-day-num">${pd.getDate()}</div></div>`;
  }
  for(let d=1;d<=daysInMonth;d++){
    const thisDate=new Date(calYear,calMonth,d); thisDate.setHours(0,0,0,0);
    const dateStr=thisDate.toISOString().slice(0,10);
    const isToday=thisDate.getTime()===today.getTime();
    const dayTasks=getActiveTasks().filter(t=>t.dueDate===dateStr);
    let chips=dayTasks.slice(0,3).map(t=>
      `<span class="cal-task-chip" style="background:${calChipBg(t.type)};color:${calChipFg(t.type)}"
        onclick="openSteps('${t._spId}','${escHtml(t.name)}')" title="Click to view steps — ${escHtml(t.name)}">${escHtml(t.name)}</span>`).join('');
    if(dayTasks.length>3) chips+=`<span style="font-size:10px;color:var(--text-faint)">+${dayTasks.length-3} more</span>`;
    html+=`<div class="cal-day${isToday?' today':''}"><div class="cal-day-num">${d}</div>${chips}</div>`;
  }
  grid.innerHTML=html;
}

// ── TEAM ──────────────────────────────────────────────────────

// ── FEATURE 3 — STEP COMMENTS ────────────────────────────────
function toggleStepComments(stepSpId) {
  const panel = document.getElementById(`step-comments-${stepSpId}`);
  if (panel) panel.classList.toggle('hidden');
}

async function addStepComment(stepSpId) {
  const input = document.getElementById(`step-comment-input-${stepSpId}`);
  if (!input) return;
  const text = input.value.trim();
  if (!text) return;
  if (text.length > COMMENT_MAX_LENGTH) {
    showToast('Comment is too long — please keep it under 2,000 characters.', 'warning'); return;
  }
  const step  = _steps.find(s => s._spId === stepSpId); if (!step) return;
  const task  = _tasks.find(t => t._spId === step.taskId);
  const now   = new Date();
  const label = now.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}) + ', ' +
                now.toLocaleTimeString('en-US',{hour:'numeric',minute:'2-digit'});
  try {
    const fields = {
      Title:        `${currentUser.name} → step comment`,
      CommentId:    uid(),
      TaskId:       task?._spId || '',
      StepId:       stepSpId,
      AuthorId:     currentUserId(),
      CommentText:  text,
      CommentTime:  label,
      Timestamp:    String(Date.now()),
      TimestampISO: new Date().toISOString(),
      IsResolved:   'No',
    };
    const created = await createListItem(LISTS.comments, fields);
    _comments.push(normaliseComment({ ...fields, id: created?.id || fields.CommentId }));
    input.value = '';
    renderStepsPanel();
  } catch(e) { showToast('Could not post comment: ' + e.message, 'error'); }
}

// ═══════════════════════════════════════════════════════════════
// ── REASSIGN REQUEST ─────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function openReassignRequest(spId, type, itemName) {
  document.getElementById('modal-title').textContent = '🔄 Request Reassignment';
  document.getElementById('modal-body').innerHTML = `
    <p style="font-size:13px;color:var(--text-muted);margin-bottom:16px">
      This will notify all admins that you are requesting reassignment of:<br>
      <b style="color:var(--navy)">${escHtml(itemName)}</b>
    </p>
    <div class="form-group">
      <label>Reason <span style="font-weight:400;color:var(--text-faint)">(optional — helps the admin understand the situation)</span></label>
      <textarea id="reassign-note-input" rows="3" placeholder="e.g. Out sick this week, travelling for client meetings, overloaded with footnotes…" style="width:100%;resize:vertical"></textarea>
    </div>
    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeAllModals()">Cancel</button>
      <button class="btn-primary" onclick="submitReassignRequest('${spId}','${type}','${escHtml(itemName)}')">Send Request to Admin</button>
    </div>`;
  openModal();
}

async function submitReassignRequest(spId, type, itemName) {
  const note    = document.getElementById('reassign-note-input')?.value.trim() || '';
  const list    = type==='task' ? LISTS.tasks : LISTS.steps;
  const arr     = type==='task' ? _tasks : _steps;
  const item    = arr.find(i=>i._spId===spId); if(!item) return;
  closeAllModals();
  showLoadingOverlay(true);
  try {
    // Set flag on the item
    await updateListItem(list, spId, { ReassignRequested: 'Yes', ReassignNote: note });
    item.reassignRequested = true; item.reassignNote = note;

    // Notify all admins via FT_Notifications list (PA flow sends the email)
    const admins = _users.filter(u => u.isAdmin && u.email);
    const rTaskName = type==='task' ? itemName : (getTasks().find(t=>t._spId===item.taskId)?.name||'');
    const rStepName = type==='step' ? itemName : '';
    for (const admin of admins) {
      createListItem(LISTS.notifications, {
        Title:         `[FRT] Reassignment requested: ${rTaskName}${rStepName ? ' — '+rStepName : ''}`,
        ToEmail:       admin.email,
        TaskName:      rTaskName,
        StepName:      rStepName,
        StatusChanged: 'Reassignment Requested',
        ChangedBy:     currentUser.name + (note ? ` — "${note}"` : ''),
        AppUrl:        window.location.href,
        NotifyType:    'reassign',
      }).catch(e => console.warn('Reassign notify write failed:', e.message));
    }
    renderCurrentView();
    if (document.getElementById('steps-overlay') && !document.getElementById('steps-overlay').classList.contains('hidden')) renderStepsPanel();
    showToast('Reassignment request sent to all admins.', 'success');
  } catch(e) { showToast('Could not send reassignment request: '+e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}

async function clearReassignFlag(spId, type) {
  const list = type==='task' ? LISTS.tasks : LISTS.steps;
  const arr  = type==='task' ? _tasks : _steps;
  const item = arr.find(i=>i._spId===spId); if(!item) return;
  try {
    await updateListItem(list, spId, { ReassignRequested: 'No', ReassignNote: '' });
    item.reassignRequested = false; item.reassignNote = '';
    renderCurrentView();
    if (document.getElementById('steps-overlay') && !document.getElementById('steps-overlay').classList.contains('hidden')) renderStepsPanel();
  } catch(e) { showToast('Could not clear flag: '+e.message, 'error'); }
}

function toggleTeamYearVisibility() {
  const q   = document.getElementById('team-quarter')?.value;
  const yEl = document.getElementById('team-year');
  if (yEl) yEl.style.display = q ? '' : 'none';
}
function toggleMyTasksYearVisibility() {
  const q   = document.getElementById('mytasks-quarter')?.value;
  const yEl = document.getElementById('mytasks-year');
  if (yEl) yEl.style.display = (q && q !== 'all') ? '' : 'none';
}
function renderTeam() {
  const q   = document.getElementById('team-quarter')?.value;
  const yr  = parseInt(document.getElementById('team-year')?.value || 0);
  const grid=document.getElementById('team-grid'); if(!grid) return;
  grid.innerHTML=getUsers().map(user=>{
    const userId=user._spId||user.id;
    let allTasks = getActiveTasks().filter(t=>
      t.ownerId===userId || t.reviewerId===userId || t.reviewer2Id===userId
    );
    if (q) allTasks = allTasks.filter(t => t.quarter===q && (!yr || t.year===yr));
    const activeMyTasks=allTasks.filter(t=>t.status!=='Not Applicable');
    const complete=activeMyTasks.filter(t=>t.status==='Complete').length;
    const overdue=activeMyTasks.filter(t=>deadlineStatus(t.dueDate,t.status)==='overdue').length;
    const inReview=activeMyTasks.filter(t=>t.status==='Ready for Review 1'||t.status==='Ready for Review 2').length;
    const pct=activeMyTasks.length?Math.round(complete/activeMyTasks.length*100):0;
    const taskList=allTasks.slice(0,4).map(t=>{
      const ds=deadlineStatus(t.dueDate,t.status);
      return `<div class="team-task-item" onclick="openSteps('${t._spId}','${escHtml(t.name)}')" style="cursor:pointer" title="View steps — ${escHtml(t.name)}">
        <span class="team-task-name" style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escHtml(t.name)}</span>
        <span class="deadline-dot ${dotClass(ds)}" style="width:8px;height:8px;border-radius:50%;display:inline-block;flex-shrink:0"></span>
      </div>`;
    }).join('')||'<div class="text-muted" style="font-size:12px;padding:6px 0">No tasks assigned</div>';
    return `<div class="team-card">
      <div class="team-card-header">
        <div class="team-avatar">${initials(user.name)}</div>
        <div><div class="team-name">${escHtml(user.name)}</div><div class="team-role">${escHtml(user.role)}</div></div>
      </div>
      <div class="team-stats">
        <div class="team-stat"><div class="team-stat-val">${activeMyTasks.length}</div><div class="team-stat-label">Total</div></div>
        <div class="team-stat"><div class="team-stat-val">${complete}</div><div class="team-stat-label">Done</div></div>
        <div class="team-stat"><div class="team-stat-val" style="color:var(--purple-600)">${inReview}</div><div class="team-stat-label">For Review</div></div>
        <div class="team-stat"><div class="team-stat-val" style="color:var(--deadline-overdue)">${overdue}</div><div class="team-stat-label">Overdue</div></div>
      </div>
      <div class="progress-bar-wrap" style="margin-bottom:14px">
        <div class="progress-bar-fill" style="width:${pct}%"></div>
      </div>
      <div class="team-task-list">${taskList}</div>
    </div>`;
  }).join('');
}

// ── FEATURE 2 — QUARTER DATES ADMIN ──────────────────────────
function renderQuarterDatesPanel() {
  const el = document.getElementById('admin-quarter-dates-list');
  if (!el) return;
  if (!_quarterDates.length) {
    el.innerHTML = `<p class="text-muted" style="font-size:13px;padding:12px 22px">
      No key dates set yet. Click + Add Quarter to add SEC filing dates,
      earnings call times, and other key deadlines.
    </p>`;
    return;
  }
  el.innerHTML = _quarterDates
    .sort((a,b) => { const ak=`${b.year}${b.quarter}`,bk=`${a.year}${a.quarter}`; return ak>bk?1:ak<bk?-1:0; })
    .map(qd => `<div class="admin-user-row">
      <div class="admin-user-info">
        <div class="user-name-sm">${escHtml(qd.quarter)} ${qd.year}</div>
        <div class="user-role-sm">
          ${qd.secFilingDate    ? `SEC: ${formatDate(qd.secFilingDate)} · ` : ''}
          ${qd.earningsDate     ? `Earnings: ${formatDate(qd.earningsDate)}${qd.earningsTime?' '+qd.earningsTime:''} · ` : ''}

        </div>
      </div>
      <div style="display:flex;gap:8px">
        <button class="icon-btn" onclick="openEditQuarterDate('${qd._spId}')">✏️</button>
        <button class="icon-btn" onclick="deleteQuarterDate('${qd._spId}',this)">🗑</button>
      </div>
    </div>`).join('');
}

function openAddQuarterDate()        { showQuarterDateModal(null); }
function openEditQuarterDate(spId)   { showQuarterDateModal(_quarterDates.find(q=>q._spId===spId)); }

function showQuarterDateModal(qd) {
  const cur = new Date().getFullYear();
  const yrOpts = [cur-3,cur-2,cur-1,cur,cur+1].map(y=>`<option value="${y}" ${qd?.year===y?'selected':''}>${y}</option>`).join('');
  document.getElementById('modal-title').textContent = qd ? 'Edit Quarter Dates' : 'Add Quarter Dates';
  document.getElementById('modal-body').innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="form-group"><label>Quarter</label>
        <select id="qd-quarter">${['Q1','Q2','Q3','Q4'].map(q=>`<option ${qd?.quarter===q?'selected':''}>${q}</option>`).join('')}</select></div>
      <div class="form-group"><label>Year</label>
        <select id="qd-year">${yrOpts}</select></div>
    </div>
    <div class="form-group"><label>SEC Filing Date</label>
      <input type="date" id="qd-sec" value="${qd?.secFilingDate||''}" /></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="form-group"><label>Earnings Call Date</label>
        <input type="date" id="qd-earn-date" value="${qd?.earningsDate||''}" /></div>
      <div class="form-group"><label>Earnings Call Time <span style="font-weight:400;color:var(--text-faint)">(optional)</span></label>
        <input type="text" id="qd-earn-time" value="${escHtml(qd?.earningsTime||'')}" placeholder="e.g. 8:30 AM ET" /></div>
    </div>

    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeAllModals()">Cancel</button>
      <button class="btn-primary" onclick="saveQuarterDate('${qd?._spId||''}')">Save</button>
    </div>`;
  openModal();
}

async function saveQuarterDate(spId) {
  const fields = {
    Title:            document.getElementById('qd-quarter').value + ' ' + document.getElementById('qd-year').value,
    Quarter:          document.getElementById('qd-quarter').value,
    Year:             document.getElementById('qd-year').value,
    SECFilingDate:    document.getElementById('qd-sec').value       || null,
    EarningsCallDate: document.getElementById('qd-earn-date').value || null,
    EarningsCallTime: document.getElementById('qd-earn-time').value || '',

  };
  closeAllModals(); showLoadingOverlay(true);
  try {
    if (spId) {
      await updateListItem(LISTS.quarterDates, spId, fields);
      const idx = _quarterDates.findIndex(q=>q._spId===spId);
      if (idx>=0) _quarterDates[idx] = normaliseQuarterDate({...fields, id: spId});
    } else {
      const created = await createListItem(LISTS.quarterDates, fields);
      _quarterDates.push(normaliseQuarterDate({...fields, id: created?.id||uid()}));
    }
    renderQuarterDatesPanel();
    renderCurrentView();
  } catch(e) { showToast('Could not save quarter dates: '+e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}

async function deleteQuarterDate(spId, btnEl) {
  showInlineConfirm(btnEl || document.body, 'Delete these key dates?', async () => {
    try {
      await deleteListItem(LISTS.quarterDates, spId);
      _quarterDates = _quarterDates.filter(q=>q._spId!==spId);
      renderQuarterDatesPanel();
      renderCurrentView();
      showToast('Key dates deleted.', 'success');
    } catch(e) { showToast('Could not delete: '+e.message, 'error'); }
  }, 'Delete', true);
}

// ── INLINE QUICK NOTE ────────────────────────────────────────
function openQuickNote(spId) {
  const _t = _tasks.find(t=>t._spId===spId);
  if (_t && isQuarterLocked(_t.quarter, _t.year) && !currentUser.isAdmin) {
    showToast('This quarter is locked — notes are read-only.', 'warning'); return;
  }
  const currentNote = _t ? (_t.notes||_t.description||'') : '';
  const existing = document.getElementById('quick-note-pop');
  if (existing) existing.remove();
  const pop = document.createElement('div');
  pop.id        = 'quick-note-pop';
  pop.className = 'quick-note-pop';
  pop.innerHTML = `
    <div class="quick-note-header">Quick Note</div>
    <textarea id="quick-note-ta" class="quick-note-ta" placeholder="Add a note — carries forward during rollforward…" rows="3">${escHtml(currentNote||'')}</textarea>
    <div class="quick-note-footer">
      <button class="btn-secondary small" onclick="document.getElementById('quick-note-pop').remove()">Cancel</button>
      <button class="btn-primary small" onclick="saveQuickNote('${spId}')">Save</button>
    </div>`;
  document.body.appendChild(pop);
  const ta = document.getElementById('quick-note-ta');
  if (ta) ta.focus();
  // Single named handler so Cancel button can also remove it
  function _qnOutsideClick(e) {
    const qn = document.getElementById('quick-note-pop');
    if (!qn) { document.removeEventListener('click', _qnOutsideClick); return; }
    if (!qn.contains(e.target)) { qn.remove(); document.removeEventListener('click', _qnOutsideClick); }
  }
  // Cancel button explicitly removes the handler
  pop.querySelector('.btn-secondary').addEventListener('click', () => {
    document.removeEventListener('click', _qnOutsideClick);
  });
  setTimeout(() => document.addEventListener('click', _qnOutsideClick), 0);
}

async function saveQuickNote(spId) {
  const ta   = document.getElementById('quick-note-ta'); if(!ta) return;
  const note = ta.value.trim();
  const task = _tasks.find(t=>t._spId===spId); if(!task) return;
  document.getElementById('quick-note-pop')?.remove();
  task.notes       = note;
  task.description = note;
  renderCurrentView();
  try {
    await updateListItem(LISTS.tasks, spId, { Notes: note, Description: note });
    showToast('Note saved.', 'success', 2000);
  } catch(e) { showToast('Could not save note: '+e.message, 'error'); }
}

// ── ADMIN ─────────────────────────────────────────────────────
function renderAdmin() {
  if(!currentUser?.isAdmin) return;
  // ── Reassignment request banner ──────────────────────────
  const pendingTasks = _tasks.filter(t=>t.reassignRequested);
  const pendingSteps = _steps.filter(s=>s.reassignRequested);
  const pendingTotal = pendingTasks.length + pendingSteps.length;
  const reassignBanner = document.getElementById('admin-reassign-banner');
  if (reassignBanner) {
    if (pendingTotal > 0) {
      reassignBanner.style.display = 'block';
      reassignBanner.innerHTML = `
      <div style="background:#FEF3C7;border:1px solid #FDE68A;border-radius:8px;
                  padding:12px 16px;margin-bottom:16px;display:flex;
                  align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px">
        <span style="font-size:13px;color:#92400E">
          🔄 <b>${pendingTotal} reassignment request${pendingTotal!==1?'s':''} pending</b>
          — ${[...pendingTasks.map(t => t.name), ...pendingSteps.map(s => s.name)]
              .slice(0, 3).map(n => '<i>' + escHtml(n) + '</i>').join(', ')}
          ${pendingTotal > 3 ? ' and ' + (pendingTotal - 3) + ' more' : ''}
        </span>
        <button class="btn-secondary" style="font-size:12px;padding:5px 12px" onclick="switchView('tasks',document.querySelector('[data-view=\'tasks\']'))">View in All Tasks →</button>
      </div>`;
    } else {
      reassignBanner.style.display = 'none';
      reassignBanner.innerHTML = '';
    }
  }
  renderCustomHolidays();
  renderCloseCalendarsPanel();
  renderQuarterDatesPanel();

  // Locked quarters panel
  const lkEl = document.getElementById('admin-locks-list');
  if (lkEl) {
    if (!_locks.length) {
      lkEl.innerHTML = '<p class="text-muted" style="font-size:13px;padding:12px 22px">No quarters locked yet.</p>';
    } else {
      lkEl.innerHTML = _locks
        .sort((a,b) => { const ak=`${b.year}${b.quarter}`,bk=`${a.year}${a.quarter}`; return ak>bk?1:ak<bk?-1:0; })
        .map(l => `<div class="admin-user-row">
          <div class="admin-user-info">
            <div class="user-name-sm">🔒 ${escHtml(l.quarter)} ${l.year}</div>
            <div class="user-role-sm">Locked by ${escHtml(l.lockedBy)} · ${escHtml(l.lockedAt)}</div>
          </div>
          <button class="btn-secondary" style="font-size:12px;padding:5px 12px" onclick="unlockQuarter('${l._spId}')">Unlock</button>
        </div>`).join('');
    }
  }

  const uEl=document.getElementById('admin-user-list');
  if(uEl) uEl.innerHTML=getUsers().map(u=>`
    <div class="admin-user-row">
      <div class="admin-user-info">
        <div class="user-name-sm">${escHtml(u.name)}</div>
        <div class="user-role-sm">${escHtml(u.role)}</div>
      </div>
      <div style="display:flex;align-items:center;gap:8px">
        <span class="${u.isAdmin?'admin-badge-admin':'admin-badge-member'}">${u.isAdmin?'Admin':'Member'}</span>
        <button class="icon-btn" onclick="openEditUser('${u._spId}')">✏️</button>
        ${u._spId!==currentUser._spId?`<button class="icon-btn" onclick="deleteUser('${u._spId}',this)">🗑</button>`:''}
      </div>
    </div>`).join('');
  const tEl=document.getElementById('admin-template-list');
  if(tEl) tEl.innerHTML=getTemplates().map(tp=>{
    const stpCount = getStepTemplatesForTemplate(tp._spId||tp.id).length;
    return `<div class="template-row">
      <div>
        <div class="template-name">${escHtml(tp.name)}</div>
        <div class="template-type">${escHtml(tp.type)} · Due ${tp.dueDaysFromQtrEnd} days after quarter end · ${stpCount} default step${stpCount!==1?'s':''}</div>
      </div>
      <div style="display:flex;gap:8px;align-items:center">
        <button class="btn-secondary" style="font-size:11px;padding:5px 10px" onclick="openStepTemplates('${tp._spId||tp.id}','${escHtml(tp.name)}')">📋 Steps</button>
        <button class="icon-btn" onclick="openEditTemplate('${tp._spId}')">✏️</button>
        <button class="icon-btn" onclick="deleteTemplate('${tp._spId}',this)">🗑</button>
      </div>
    </div>`;
  }).join('');
}

// Creates tasks for the selected quarter based on saved quarterly templates.
async function applyTemplate() {
  const quarter = document.getElementById('template-quarter-select').value;
  const year    = parseInt(document.getElementById('template-year-select').value);
  if (isQuarterLocked(quarter, year)) {
    showToast(`${quarter} ${year} is locked — unlock it in Admin before applying templates.`, 'warning'); return;
  }
  if (!getTemplates().length) {
    showToast('No templates defined yet. Add templates in the Quarterly Templates section above.', 'warning'); return;
  }
  const qEnd    = quarterEndDate(quarter, year);
  let added = 0;
  showLoadingOverlay(true, `Applying templates to ${quarter} ${year}…`);
  try {
  for (const tp of getTemplates()) {
    const exists = getTasks().some(t => t.name===tp.name && t.quarter===quarter && t.year===year);
    if (!exists) {
      const taskId   = uid();
      const created  = await createListItem(LISTS.tasks, {
        Title: tp.name, TaskId: taskId, TaskType: tp.type,
        Quarter: quarter, Year: String(year),
        DueDate: addDays(qEnd, tp.dueDaysFromQtrEnd),
        Status: 'Not Started', OwnerId: tp.defaultOwnerId, Notes: '', Description: '',
        Applicability: 'All Quarters', WorkdayNum: '', SkipNextRollforward: 'No',
      });
      const taskSpId = created?.id || taskId;
      // Auto-create steps from step templates
      const stpls = getStepTemplatesForTemplate(tp._spId || tp.id);
      for (const stpl of stpls) {
        await createListItem(LISTS.steps, {
          Title: stpl.name, StepId: uid(),
          TaskId: String(taskSpId),
          StepOrder: String(stpl.order),
          Status: 'Not Started',
          OwnerId: stpl.defaultOwnerId || '',
          DueDate: stpl.dueDaysFromQtrEnd ? nearestWorkingDay(addDays(qEnd, stpl.dueDaysFromQtrEnd)) : '',
          RequiresPrev: stpl.requiresPrev ? 'Yes' : 'No',
          Applicability: 'All Quarters',
          WorkdayNum: stpl.workdayNum ? String(stpl.workdayNum) : '',
          Note: '', Notes: '',
        });
      }
      added++;
    }
  }
    await refreshData();
    showToast(`Applied template: ${added} task${added!==1?'s':''} added to ${quarter} ${year}.`, 'success');
  } catch(e) { showToast('Template apply failed: ' + e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}

// ── TASK MODAL ────────────────────────────────────────────────
// Opens the Add Task modal pre-filled with the currently viewed quarter.
function openAddTask() {
  editingTaskId = null;
  // Pre-select the currently viewed quarter/year when adding a new task
  const q  = document.getElementById('quarter-filter')?.value;
  const yr = parseInt(document.getElementById('year-filter')?.value || 0);
  showTaskModal(null, q, yr);
}
function openEditTask(spId)   { editingTaskId=spId;   showTaskModal(_tasks.find(t=>t._spId===spId)); }
function showTaskModal(task, defaultQ, defaultY) {
  document.getElementById('modal-title').textContent=task?'Edit Task':'Add New Task';
  const _tuid = u => u._spId||u.id;
  const ownerOpts=getUsers().map(u=>`<option value="${_tuid(u)}" ${task?.ownerId===_tuid(u)?'selected':''}>${escHtml(u.name)}</option>`).join('');
  const reviewerOpts = `<option value="">— None —</option>` + getUsers().map(u=>`<option value="${_tuid(u)}" ${task?.reviewerId===_tuid(u)?'selected':''}>${escHtml(u.name)}</option>`).join('');
  const reviewer2Opts = `<option value="">— None —</option>` + getUsers().map(u=>`<option value="${_tuid(u)}" ${task?.reviewer2Id===_tuid(u)?'selected':''}>${escHtml(u.name)}</option>`).join('');
  const cur=new Date().getFullYear();
  const effQ = task?.quarter || defaultQ;
  const effY = task?.year    || defaultY || cur;
  const yearOpts=[cur-3,cur-2,cur-1,cur,cur+1].map(y=>`<option value="${y}" ${effY===y?'selected':''}>${y}</option>`).join('');
  document.getElementById('modal-body').innerHTML=`
    <div class="form-group"><label>Deliverable Name</label>
      <input type="text" id="tf-name" value="${escHtml(task?.name||'')}" placeholder="e.g. 10-Q Filing" /></div>
    <div class="form-group"><label>Type</label>
      <select id="tf-type">
        ${['Close','Financial Report','Master SS','Ops Book','Other','Press Release','Post-Filing','Pre-Filing']
          .map(t => `<option ${task?.type===t?'selected':''}>${t}</option>`).join('')}
      </select></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="form-group"><label>Quarter</label>
        <select id="tf-quarter">${['Q1','Q2','Q3','Q4'].map(q=>`<option ${effQ===q?'selected':''}>${q}</option>`).join('')}</select></div>
      <div class="form-group"><label>Year</label>
        <select id="tf-year">${yearOpts}</select></div>
    </div>
    <div class="form-group"><label>Due Date</label>
      <div style="display:grid;grid-template-columns:90px 1fr;gap:10px;align-items:end">
        <div>
          <label style="font-size:11px;color:var(--text-faint);display:block;margin-bottom:4px">Workday #</label>
          <input type="number" id="tf-workday"
            value="${task?.workdayNum||''}" min="1" max="60" placeholder="WD #"
            style="width:100%;padding:10px 8px;border:1.5px solid var(--border);border-radius:var(--radius);font-family:inherit;font-size:14px"
            oninput="resolveWdToDate('tf-workday','tf-due',document.getElementById('tf-quarter')?.value,document.getElementById('tf-year')?.value)" />
        </div>
        <input type="date" id="tf-due" value="${task?.dueDate||''}" />
      </div>
      <p style="font-size:11px;color:var(--text-faint);margin-top:4px">Enter a workday number to auto-fill the date, or set the date directly.</p>
    </div>
    <div class="form-group"><label>Preparer (Owner)</label>
      <select id="tf-owner">${ownerOpts}</select></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="form-group"><label>Reviewer 1 <span style="font-weight:400;color:var(--text-faint)">(optional)</span></label>
        <select id="tf-reviewer">${reviewerOpts}</select></div>
      <div class="form-group"><label>Reviewer 2 <span style="font-weight:400;color:var(--text-faint)">(optional)</span></label>
        <select id="tf-reviewer2">${reviewer2Opts}</select></div>
    </div>
    <div class="form-group"><label>Status</label>
      <select id="tf-status">${STATUS_ORDER.map(s=>`<option ${task?.status===s?'selected':''}>${s}</option>`).join('')}</select></div>
    <div class="form-group" style="display:flex;align-items:center;gap:10px">
      <input type="checkbox" id="tf-skip" ${task?.skipNextRollforward?'checked':''} style="width:16px;height:16px;cursor:pointer" />
      <label for="tf-skip" style="margin:0;cursor:pointer">Skip this task on the next rollforward <span style="font-weight:400;color:var(--text-faint)">(auto-clears after one quarter)</span></label>
    </div>
    <div class="form-group"><label>Applicability</label>
      <select id="tf-applicability">
        <option value="All Quarters" ${(!task?.applicability||task?.applicability==='All Quarters')?'selected':''}>All Quarters</option>
        <option value="10-K only (Q4)" ${task?.applicability==='10-K only (Q4)'?'selected':''}>10-K only (Q4)</option>
        <option value="10-Q only (Q1, Q2, Q3)" ${task?.applicability==='10-Q only (Q1, Q2, Q3)'?'selected':''}>10-Q only (Q1, Q2, Q3)</option>
      </select></div>
    <div class="form-group"><label>Notes <span style="font-weight:400;color:var(--text-faint)">(optional — can carry forward during rollforward)</span></label>
      <textarea id="tf-desc" rows="3"
        placeholder="e.g. Key instructions, prior quarter context, things to watch out for…"
        style="resize:vertical">${escHtml(task?.notes||task?.description||'')}</textarea></div>
    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeAllModals()">Cancel</button>
      <button class="btn-primary" onclick="saveTask()">Save Task</button>
    </div>`;
  openModal();
}
async function saveTask() {
  const name=document.getElementById('tf-name').value.trim();
  const quarter = document.getElementById('tf-quarter').value;
  const year    = parseInt(document.getElementById('tf-year').value);
  // Prevent saving to a locked quarter (non-admins)
  if (!currentUser.isAdmin) {
    if (editingTaskId) {
      const existing = _tasks.find(t=>t._spId===editingTaskId);
      if (existing && isQuarterLocked(existing.quarter, existing.year)) {
        showToast(`${existing.quarter} ${existing.year} is locked — unlock it in Admin to make changes.`, 'warning');
        closeAllModals(); return;
      }
    } else {
      // Creating a new task — check the selected quarter/year
      const q  = document.getElementById('tf-quarter')?.value;
      const yr = parseInt(document.getElementById('tf-year')?.value || 0);
      if (q && yr && isQuarterLocked(q, yr)) {
        showToast(`${q} ${yr} is locked — unlock it in Admin before adding tasks.`, 'warning');
        closeAllModals(); return;
      }
    }
  }
  const wdNum   = parseInt(document.getElementById('tf-workday')?.value) || null;
  let   dueDate = document.getElementById('tf-due').value;
  // If workday set and calendar exists, resolve to real date
  if (wdNum) {
    const resolved = workdayToDate(wdNum, quarter, year);
    if (resolved) dueDate = resolved;
  }
  if(!name){showToast('Please fill in the task name.','warning');return;}
  if(!dueDate){
    if(wdNum) showToast(`WD${wdNum} could not resolve to a date — set the close calendar for ${quarter} ${year} in Admin first, then retry.`,'warning');
    else showToast('Please enter a due date or workday number.','warning');
    return;
  }
  const fields={
    Title:       name,
    TaskType:    document.getElementById('tf-type').value,
    Quarter:     quarter,
    Year:        String(year),
    DueDate:     dueDate,
    OwnerId:     document.getElementById('tf-owner').value,
    ReviewerId:  document.getElementById('tf-reviewer')?.value  || '',
    Reviewer2Id: document.getElementById('tf-reviewer2')?.value || '',
    Status:      document.getElementById('tf-status').value,
    Notes:       document.getElementById('tf-desc').value.trim(),
    Description: document.getElementById('tf-desc').value.trim(), // kept in sync with Notes for SP compat
    Applicability: document.getElementById('tf-applicability')?.value || 'All Quarters',
    WorkdayNum:  wdNum ? String(wdNum) : '',
    SkipNextRollforward: document.getElementById('tf-skip')?.checked ? 'Yes' : 'No',
  };
  closeAllModals();
  showLoadingOverlay(true);
  try {
    if(editingTaskId) {
      const old = _tasks.find(t=>t._spId===editingTaskId);
      if (old && old.dueDate !== dueDate) {
        await recordDueDateChange(editingTaskId, 'task', name, old.dueDate, dueDate);
      }
      await updateListItem(LISTS.tasks, editingTaskId, fields);
    } else {
      fields.TaskId = uid();
      await createListItem(LISTS.tasks, fields);
    }
    await refreshData();
  } catch(e) { showToast('Could not save task: '+e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}
// Confirms then deletes a task and all its child records (steps, comments, sign-offs, attachments).
async function deleteTask(spId, btnEl) {
  const childSteps = _steps.filter(s => s.taskId === spId);
  const childCount = childSteps.length;
  const msg = childCount
    ? `Delete this task and its ${childCount} step${childCount!==1?'s':''}? This cannot be undone.`
    : 'Delete this task? This cannot be undone.';
  showInlineConfirm(btnEl || document.body, msg, async () => {
    const task = _tasks.find(t=>t._spId===spId);
    // Optimistic removal from all caches
    _tasks       = _tasks.filter(t=>t._spId!==spId);
    _steps       = _steps.filter(s=>s.taskId!==spId);
    _comments    = _comments.filter(c=>c.taskId!==spId&&c.taskId!==task?.id);
    _signOffs    = _signOffs.filter(s=>s.refId!==spId);
    _attachments = _attachments.filter(a=>a.taskId!==spId);
    renderCurrentView();
    try {
      // Delete task row first, then cascade all child records in background.
      // Sign-offs are deleted for both the task itself AND each child step —
      // previously only step sign-offs were cleaned from SharePoint.
      await deleteListItem(LISTS.tasks, spId);

      // Task-level sign-offs (refType='task' or 'due_date_change' on this task)
      const taskSignOffs = _signOffs.filter(so => so.refId === spId);
      taskSignOffs.forEach(so => deleteListItem(LISTS.signOffs, so._spId).catch(()=>{}));

      // Task-level comments
      const taskComments = _comments.filter(c => c.taskId === spId && !c.stepId);
      taskComments.forEach(c => deleteListItem(LISTS.comments, c._spId).catch(()=>{}));

      // Child steps and all their children
      for (const s of childSteps) {
        deleteListItem(LISTS.steps, s._spId).catch(()=>{});
        _comments.filter(c=>c.stepId===s._spId).forEach(c =>
          deleteListItem(LISTS.comments, c._spId).catch(()=>{})
        );
        _signOffs.filter(so=>so.refId===s._spId).forEach(so =>
          deleteListItem(LISTS.signOffs, so._spId).catch(()=>{})
        );
        _attachments.filter(a=>a.stepId===s._spId).forEach(a =>
          deleteListItem(LISTS.attachments, a._spId).catch(()=>{})
        );
      }
    } catch(e) {
      showToast('Could not delete task: '+e.message, 'error');
      if(task) _tasks.push(task); renderCurrentView();
    }
  }, 'Delete', true);
}

// ── COMMENTS ─────────────────────────────────────────────────
// Opens the comment modal for a task and starts comment polling.
function openComments(taskSpId, taskName) {
  commentingTaskId = taskSpId;
  document.getElementById('comment-task-title').textContent = taskName || 'Comments';
  document.getElementById('comment-input').value = '';
  renderCommentList();
  startCommentPolling();
  document.getElementById('comment-overlay').classList.remove('hidden');
}
function renderCommentList() {
  const list=document.getElementById('comment-list'); if(!list) return;
  const taskComments=_comments
    .filter(c=>c.taskId===commentingTaskId)
    .sort((a,b)=>{ const at=a.tsIso||String(a.ts||""); const bt=b.tsIso||String(b.ts||""); return at<bt?-1:at>bt?1:0; });
  if(!taskComments.length){
    list.innerHTML='<div class="text-muted" style="font-size:13px;padding:8px 0">No comments yet.</div>';
    return;
  }
  const unresolved = taskComments.filter(c=>!c.isResolved);
  const resolved   = taskComments.filter(c=>c.isResolved);
  const renderComment = (c, dimmed) => {
    const authorName = c.authorId ? (getUserById(c.authorId)?.name || 'Former team member') : 'Unknown';
    return `<div class="comment-item ${dimmed?'comment-resolved':''}">
      <div class="comment-header">
        <span class="comment-author">${escHtml(authorName)}</span>
        <span class="comment-time">${escHtml(c.time)}</span>
        ${c.isResolved
          ? '<span class="comment-resolved-label">✓ Resolved</span>'
          : `<button class="comment-resolve-btn" onclick="resolveComment('${c._spId}')">✓ Resolve</button>`}
      </div>
      <div class="comment-text" style="${dimmed?'opacity:0.5':''}"> ${escHtml(c.text)}</div>
    </div>`;
  };
  list.innerHTML =
    unresolved.map(c=>renderComment(c,false)).join('') +
    (resolved.length ? `
      <details style="margin-top:8px">
        <summary style="font-size:11px;color:var(--text-faint);cursor:pointer;padding:4px 0">
          ${resolved.length} resolved comment${resolved.length!==1?'s':''}
        </summary>
        ${resolved.map(c=>renderComment(c,true)).join('')}
      </details>` : '');
  list.scrollTop=list.scrollHeight;
}

async function resolveComment(commentSpId) {
  const comment = _comments.find(c=>c._spId===commentSpId); if(!comment) return;
  try {
    await updateListItem(LISTS.comments, commentSpId, { IsResolved: 'Yes' });
    comment.isResolved = true;
    renderCommentList();
    renderCurrentView(); // refresh unresolved badges on task rows
  } catch(e) { showToast('Could not resolve comment: '+e.message, 'error'); }
}
// Posts a new comment on the currently open task. Called from the comment modal.
async function addComment() {
  const input=document.getElementById('comment-input');
  const text=input.value.trim(); if(!text||!commentingTaskId) return;
  if (text.length > COMMENT_MAX_LENGTH) {
    showToast('Comment is too long — please keep it under 2,000 characters.', 'warning'); return;
  }
  input.value='';
  const now=new Date();
  const commentId = uid();
  const commentFields = {
    Title:       text.slice(0,50),
    CommentId:   commentId,
    TaskId:      commentingTaskId,
    AuthorId:    currentUserId(),
    CommentText: text,
    CommentTime: now.toLocaleString('en-US',{month:'short',day:'numeric',year:'numeric',hour:'numeric',minute:'2-digit'}),
    Timestamp:   String(now.getTime()),
    TimestampISO: now.toISOString(),
    IsResolved:  'No',
  };
  try {
    const created = await createListItem(LISTS.comments, commentFields);
    _comments.push(normaliseComment({...commentFields, id: created?.id||commentId}));
    renderCommentList();
    renderCurrentView(); // update unresolved badges
  } catch(e) { showToast('Could not save comment: '+e.message, 'error'); }
}
function closeCommentModal() {
  stopCommentPolling();
  document.getElementById('comment-overlay').classList.add('hidden');
  commentingTaskId=null;
  renderCurrentView(); // badges already up to date — no need to re-fetch all 10 lists
}

// ── USER MODAL ────────────────────────────────────────────────
function openAddUser()        { showUserModal(null); }
function openEditUser(spId)   { showUserModal(_users.find(u=>u._spId===spId)); }
function showUserModal(user) {
  document.getElementById('modal-title').textContent=user?'Edit Team Member':'Add Team Member';
  document.getElementById('modal-body').innerHTML=`
    <div class="form-group"><label>Full Name</label>
      <input type="text" id="uf-name" value="${escHtml(user?.name||'')}" /></div>
    <div class="form-group"><label>Role / Title</label>
      <input type="text" id="uf-role" value="${escHtml(user?.role||'')}" /></div>
    <div class="form-group"><label>Work Email (for auto-login matching)</label>
      <input type="email" id="uf-email" value="${escHtml(user?.email||'')}" placeholder="name@company.com" /></div>

    <div class="form-group"><label>Access Level</label>
      <select id="uf-admin">
        <option value="false" ${!user?.isAdmin?'selected':''}>Member</option>
        <option value="true"  ${user?.isAdmin ?'selected':''}>Admin</option>
      </select></div>
    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeAllModals()">Cancel</button>
      <button class="btn-primary" onclick="saveUser('${user?._spId||''}')">Save</button>
    </div>`;
  openModal();
}
async function saveUser(existingSpId) {
  const name=document.getElementById('uf-name').value.trim();
  const role=document.getElementById('uf-role').value.trim();
  const email=document.getElementById('uf-email').value.trim();
  const isAdmin=document.getElementById('uf-admin').value==='true';
  if(!name){showToast('Please fill in the name field.','warning');return;}
  closeAllModals(); showLoadingOverlay(true);
  try {
    const fields={ Title:name, FullName:name, JobRole:role, Email:email, IsAdmin:isAdmin?'Yes':'No' };
    if(existingSpId) { await updateListItem(LISTS.users, existingSpId, fields); }
    else             { fields.UserId=uid(); await createListItem(LISTS.users, fields); }
    await refreshData();
  } catch(e) { showToast('Could not save user: '+e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}
async function deleteUser(spId, btnEl) {
  showInlineConfirm(btnEl || document.body, 'Remove this team member?', async () => {
    showLoadingOverlay(true, 'Removing…');
    try { await deleteListItem(LISTS.users, spId); await refreshData(); showToast('Team member removed.', 'success'); }
    catch(e) { showToast('Could not remove user: '+e.message, 'error'); }
    finally { showLoadingOverlay(false); }
  }, 'Remove', true);
}

// ── TEMPLATE MODAL ────────────────────────────────────────────
function openAddTemplate()       { showTemplateModal(null); }
function openEditTemplate(spId)  { showTemplateModal(_templates.find(t=>t._spId===spId)); }
function showTemplateModal(tp) {
  const ownerOpts=getUsers().map(u=>`<option value="${u._spId||u.id}" ${tp?.defaultOwnerId===(u._spId||u.id)?'selected':''}>${escHtml(u.name)}</option>`).join('');
  document.getElementById('modal-title').textContent=tp?'Edit Template':'Add Template Task';
  document.getElementById('modal-body').innerHTML=`
    <div class="form-group"><label>Task Name</label>
      <input type="text" id="tpl-name" value="${escHtml(tp?.name||'')}" /></div>
    <div class="form-group"><label>Type</label>
      <select id="tpl-type">
        ${['Close','Financial Report','Master SS','Ops Book','Other','Press Release','Post-Filing','Pre-Filing']
          .map(t => `<option ${tp?.type===t ? 'selected' : ''}>${t}</option>`).join('')}
      </select></div>
    <div class="form-group"><label>Days After Quarter End</label>
      <input type="number" id="tpl-days" value="${tp?.dueDaysFromQtrEnd||30}" min="1" max="120" /></div>
    <div class="form-group"><label>Default Owner</label>
      <select id="tpl-owner">${ownerOpts}</select></div>
    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeAllModals()">Cancel</button>
      <button class="btn-primary" onclick="saveTemplate('${tp?._spId||''}')">Save Template</button>
    </div>`;
  openModal();
}
async function saveTemplate(existingSpId) {
  const name=document.getElementById('tpl-name').value.trim();
  const type=document.getElementById('tpl-type').value;
  const days=parseInt(document.getElementById('tpl-days').value);
  const ownerId=document.getElementById('tpl-owner').value;
  if(!name||isNaN(days)){showToast('Please fill in all fields.', 'warning'); return;}
  closeAllModals(); showLoadingOverlay(true);
  try {
    const fields={ Title:name, TaskType:type, DueDaysFromQtrEnd:String(days), DefaultOwnerId:ownerId };
    if(existingSpId) { await updateListItem(LISTS.templates, existingSpId, fields); }
    else             { fields.TemplateId=uid(); await createListItem(LISTS.templates, fields); }
    await refreshData();
  } catch(e) { showToast('Could not save template: '+e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}
async function deleteTemplate(spId, btnEl) {
  showInlineConfirm(btnEl || document.body, 'Delete this template?', async () => {
    showLoadingOverlay(true, 'Deleting…');
    try { await deleteListItem(LISTS.templates, spId); await refreshData(); showToast('Template deleted.', 'success'); }
    catch(e) { showToast('Could not delete template: '+e.message, 'error'); }
    finally { showLoadingOverlay(false); }
  }, 'Delete', true);
}

// ── MODAL HELPERS ─────────────────────────────────────────────
function openModal()      { document.getElementById('modal-overlay').classList.remove('hidden'); }
// Closes all modal overlays. Called from modal close buttons and keyboard Escape.
function closeAllModals() { document.getElementById('modal-overlay').classList.add('hidden'); editingTaskId=null; editingStepId=null; editingStepTemplateId=null; }
function closeModal(e) {
  if(e.target===document.getElementById('modal-overlay'))   closeAllModals();
  if(e.target===document.getElementById('comment-overlay')) closeCommentModal();
}
function renderCurrentView() {
  const active=document.querySelector('.view.active'); if(!active) return;
  const id=active.id.replace('view-','');
  if(id==='dashboard') renderDashboard();
  if(id==='tasks')     renderAllTasks();
  if(id==='calendar')  renderCalendar();
  if(id==='team')      renderTeam();
  if(id==='admin')     renderAdmin();
  if(id==='mytasks')   renderMyTasks();
  if(id==='kanban')    renderKanban();
  if(id==='report')    renderReport();
  if(id==='exec')      renderExecView();
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 11 — VERSION ────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
const APP_VERSION = 'v2.0';

// ── NAMED CONSTANTS ───────────────────────────────────────────
const MS_PER_DAY          = 86400000; // milliseconds in one day
const TOAST_DURATION      = 4000;     // default toast display time (ms)
const TOAST_DURATION_LONG = 6000;     // longer toast for important confirmations (ms)
const UNDO_DURATION       = 10000;    // undo toast window (ms)
const COMMENT_MAX_LENGTH  = 2000;     // maximum comment character count

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 1 — PRIORITY CARD ───────────────────────────────
// ═══════════════════════════════════════════════════════════════
function renderPriorityCard() {
  const card = document.getElementById('priority-card');
  const body = document.getElementById('priority-card-body');
  if (!card || !body || !currentUser) return;

  const uid = currentUserId();
  const today   = new Date(); today.setHours(0,0,0,0);
  const in3days = new Date(today); in3days.setDate(in3days.getDate()+3);

  // Overdue tasks or tasks due within 3 days assigned to current user
  const urgentTasks = getActiveTasks().filter(t => {
    if (t.status === 'Complete' || t.status === 'Not Applicable') return false;
    const isAssigned = t.ownerId===uid || t.reviewerId===uid || t.reviewer2Id===uid;
    if (!isAssigned) return false;
    if (!t.dueDate) return false;
    const due = new Date(t.dueDate+'T00:00:00');
    return due <= in3days;
  }).sort((a,b) => (a.dueDate||'').localeCompare(b.dueDate||''));

  // Steps where user is next in chain and they're unblocked
  const urgentSteps = _steps.filter(s => {
    if (s.status === 'Complete' || s.status === 'Not Applicable') return false;
    const isAssigned = s.ownerId===currentUserId || s.reviewerId===currentUserId || s.reviewer2Id===currentUserId;
    if (!isAssigned) return false;
    const parentTask = _tasks.find(t => t._spId === s.taskId);
    if (parentTask && isQuarterLocked(parentTask.quarter, parentTask.year)) return false;
    if (getBlockingPredecessor(s)) return false; // still gated
    if (!s.dueDate) return false;
    const due = new Date(s.dueDate+'T00:00:00');
    return due <= in3days;
  }).sort((a,b) => (a.dueDate||'').localeCompare(b.dueDate||''));

  const total = urgentTasks.length + urgentSteps.length;
  if (!total) { card.classList.add('hidden'); return; }
  card.classList.remove('hidden');

  const taskRows = urgentTasks.map(t => {
    const ds  = deadlineStatus(t.dueDate, t.status);
    const isOverdue = ds === 'overdue';
    return `<div class="priority-item ${isOverdue?'priority-overdue':'priority-soon'}">
      <div class="priority-item-info">
        <span class="priority-item-name">${escHtml(t.name)}</span>
        <span class="priority-item-meta">${escHtml(t.type)} · ${isOverdue?'<span style="color:#ef4444;font-weight:600">OVERDUE</span>':'Due '+formatDate(t.dueDate)}</span>
      </div>
      <span class="status-badge ${statusBadgeClass(t.status)}" style="cursor:pointer;font-size:11px"
        onclick="cycleStatus('${t._spId}')">${escHtml(t.status)}</span>
    </div>`;
  }).join('');

  const stepRows = urgentSteps.map(s => {
    const task = getTasks().find(t=>t._spId===s.taskId);
    const ds   = deadlineStatus(s.dueDate, s.status);
    const isOverdue = ds === 'overdue';
    return `<div class="priority-item ${isOverdue?'priority-overdue':'priority-soon'}">
      <div class="priority-item-info">
        <span class="priority-item-name">${escHtml(s.name)}</span>
        <span class="priority-item-meta">Step · ${escHtml(task?.name||'')} · ${isOverdue?'<span style="color:#ef4444;font-weight:600">OVERDUE</span>':'Due '+formatDate(s.dueDate)}</span>
      </div>
      <span class="status-badge ${statusBadgeClass(s.status)}" style="cursor:pointer;font-size:11px"
        onclick="cycleStepStatus('${s._spId}')">${escHtml(s.status)}</span>
    </div>`;
  }).join('');

  body.innerHTML = taskRows + stepRows;
}

function togglePriorityCard() {
  const body = document.getElementById('priority-card-body');
  const btn  = document.getElementById('priority-toggle-btn');
  if (!body) return;
  const hidden = body.style.display === 'none';
  body.style.display = hidden ? '' : 'none';
  if (btn) btn.textContent = hidden ? '▾ Hide' : '▸ Show';
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 2 — QUICK STATUS ACTION ────────────────────────
// ═══════════════════════════════════════════════════════════════
function openQuickStatus(spId, type) {
  const item = type==='task' ? _tasks.find(t=>t._spId===spId) : _steps.find(s=>s._spId===spId);
  if (!item) return;
  const cycle = getStatusCycle(item);
  const existing = document.getElementById('quick-status-menu');
  if (existing) existing.remove();

  const menu = document.createElement('div');
  menu.id = 'quick-status-menu';
  menu.className = 'quick-status-menu';
  menu.innerHTML = cycle.concat(['Not Applicable']).map(s =>
    `<div class="quick-status-option ${s===item.status?'active':''}" onclick="quickSetStatus('${spId}','${type}','${s}');document.getElementById('quick-status-menu').remove()">
      <span class="status-badge ${statusBadgeClass(s)}" style="font-size:11px;pointer-events:none">${escHtml(s)}</span>
    </div>`
  ).join('');

  // Position near the badge that was clicked
  document.body.appendChild(menu);
  const closeMenu = (e) => { if (!menu.contains(e.target)) { menu.remove(); document.removeEventListener('click', closeMenu); } };
  setTimeout(() => document.addEventListener('click', closeMenu), 0);
}

async function quickSetStatus(spId, type, newStatus) {
  const list = type==='task' ? LISTS.tasks : LISTS.steps;
  const arr  = type==='task' ? _tasks : _steps;
  const item = arr.find(i=>i._spId===spId); if(!item) return;
  const prev = item.status;
  if (prev===newStatus) return;
  item.status = newStatus;
  try {
    await updateListItem(list, spId, { Status: newStatus });
    await writeSignOff(spId, type, item.name, prev, newStatus);
    if (newStatus==='Ready for Review 1'||newStatus==='Ready for Review 2') await notifyReviewer(item, type, newStatus);
    renderCurrentView();
    if (document.getElementById('steps-overlay') && !document.getElementById('steps-overlay').classList.contains('hidden')) renderStepsPanel();
  } catch(e) { item.status=prev; showToast('Could not update status: '+e.message, 'error'); }
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 5 — CLOSE CHECKLIST MODE ───────────────────────
// ═══════════════════════════════════════════════════════════════
// Switches to full-screen checklist mode showing all incomplete tasks for the current quarter.
function enterChecklistMode() {
  const q  = document.getElementById('quarter-filter')?.value || 'Q1';
  const yr = parseInt(document.getElementById('year-filter')?.value || new Date().getFullYear());
  const tasks = getActiveTasks()
    .filter(t => t.quarter===q && t.year===yr && t.status!=='Complete' && t.status!=='Not Applicable')
    .sort((a,b) => (a.dueDate||'').localeCompare(b.dueDate||''));

  const sub = document.getElementById('checklist-subtitle');
  if (sub) sub.textContent = `${q} ${yr} · ${tasks.length} items remaining`;

  const body = document.getElementById('checklist-body');
  if (!body) return;
  if (!tasks.length) {
    body.innerHTML = '<div style="text-align:center;color:rgba(255,255,255,.6);padding:60px 0;font-size:16px">🎉 All tasks complete for this quarter!</div>';
  } else {
    body.innerHTML = tasks.map(t => {
      const ds    = deadlineStatus(t.dueDate, t.status);
      const cycle = getStatusCycle(t);
      const next  = cycle[(cycle.indexOf(t.status)+1)%cycle.length];
      const locked = isQuarterLocked(t.quarter, t.year);
      const canEdit = !locked && (currentUser.isAdmin || t.ownerId===currentUser.id || t.ownerId===currentUser._spId);
      return `<div class="checklist-item">
        <div class="checklist-item-left">
          <span class="status-badge ${statusBadgeClass(t.status)}" style="font-size:12px">${escHtml(t.status)}</span>
          <div>
            <div style="color:#fff;font-size:14px;font-weight:500">${escHtml(t.name)}</div>
            <div style="color:rgba(255,255,255,.4);font-size:12px;margin-top:2px">
              ${escHtml(t.type)} · ${ds==='overdue'?'<span style="color:#f87171">OVERDUE</span>':formatDate(t.dueDate)}
            </div>
          </div>
        </div>
        ${canEdit ? `<button class="checklist-advance-btn" onclick="checklistAdvance('${t._spId}')">→ ${escHtml(next)}</button>` : ''}
      </div>`;
    }).join('');
  }
  document.getElementById('checklist-overlay').classList.remove('hidden');
}

async function checklistAdvance(spId) {
  const task = _tasks.find(t=>t._spId===spId); if(!task) return;
  const cycle = getStatusCycle(task);
  const next  = cycle[(cycle.indexOf(task.status)+1)%cycle.length];
  const prev  = task.status;
  task.status = next;
  try {
    await updateListItem(LISTS.tasks, spId, { Status: next });
    await writeSignOff(spId, 'task', task.name, prev, next);
    if (next==='Ready for Review 1'||next==='Ready for Review 2') await notifyReviewer(task,'task',next);
    enterChecklistMode(); // re-render
  } catch(e) { task.status=prev; showToast('Could not update: '+e.message, 'error'); }
}

function exitChecklistMode() {
  document.getElementById('checklist-overlay').classList.add('hidden');
  renderCurrentView();
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 6 — SETUP WIZARD ────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function checkSetupRequired() {
  const c = CONFIG;
  const missing = [];
  if (!c.clientId || c.clientId.startsWith('REPLACE')) missing.push('clientId');
  if (!c.tenantId || c.tenantId.startsWith('REPLACE')) missing.push('tenantId');
  if (!c.siteUrl  || c.siteUrl.startsWith('REPLACE'))  missing.push('siteUrl');
  return missing;
}

function showSetupWizardIfNeeded() {
  const missing = checkSetupRequired();
  if (!missing.length) return;
  const wizardEl = document.getElementById('setup-wizard');
  if (!wizardEl) return;

  const fields = [
    { key:'clientId', label:'Azure AD Client ID',       hint:'From portal.azure.com → App registrations → Overview',                 val: CONFIG.clientId },
    { key:'tenantId', label:'Azure AD Tenant ID',        hint:'From portal.azure.com → App registrations → Overview',                 val: CONFIG.tenantId },
    { key:'siteUrl',  label:'SharePoint Site URL',       hint:'e.g. https://moodys.sharepoint.com/sites/FinancialReporting',          val: CONFIG.siteUrl  },
  ];

  document.getElementById('wizard-steps').innerHTML = fields.map(f => {
    const isMissing = missing.some(m => m === f.key);
    return `<div class="wizard-field ${isMissing?'wizard-field-missing':''}">
      <label class="wizard-label">${escHtml(f.label)} ${isMissing?'<span style="color:#ef4444">*</span>':''}</label>
      <div class="wizard-hint">${escHtml(f.hint)}</div>
      <input type="text" class="wizard-input" data-key="${f.key}"
        value="${escHtml((!f.val||f.val.startsWith('REPLACE'))?'':f.val)}"
        placeholder="${f.key==='siteUrl'?'https://moodys.sharepoint.com/sites/...':'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'}" />
    </div>`;
  }).join('');

  wizardEl.classList.remove('hidden');
}

function saveWizardConfig() {
  const inputs = document.querySelectorAll('.wizard-input');
  inputs.forEach(inp => {
    const key = inp.dataset.key;
    const val = inp.value.trim();
    if (!val) return;
    if (key === 'clientId') CONFIG.clientId = val;
    else if (key === 'tenantId') CONFIG.tenantId = val;
    else if (key === 'siteUrl')  CONFIG.siteUrl  = val;
  });
  // Reset cached site ID so it is re-resolved with the new URL
  _graphSiteId = null;
  const stillMissing = checkSetupRequired();
  if (stillMissing.length) {
    showToast('Please fill in all required fields: ' + stillMissing.join(', '), 'warning');
    return;
  }
  document.getElementById('setup-wizard').classList.add('hidden');
  showToast('Configuration saved for this session. To make it permanent, update the CONFIG block in app.js.', 'success', TOAST_DURATION_LONG + 1000);
  loadAllData().then(() => renderCurrentView());
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 7 — HEALTH CHECK ────────────────────────────────
// ═══════════════════════════════════════════════════════════════
// Tests Graph token, SharePoint site connection, and each list. Renders results in Admin panel.
async function runHealthCheck() {
  const el = document.getElementById('health-check-results');
  if (!el) return;
  el.innerHTML = '<p style="font-size:13px;color:var(--text-faint)">Running checks…</p>';

  const results = [];

  // 1. Config present
  const cfgOk = CONFIG.clientId && !CONFIG.clientId.startsWith('REPLACE')
             && CONFIG.tenantId && !CONFIG.tenantId.startsWith('REPLACE')
             && CONFIG.siteUrl  && !CONFIG.siteUrl.startsWith('REPLACE');
  results.push({ name: 'CONFIG values', status: cfgOk ? 'ok' : 'error',
    msg: cfgOk ? 'clientId, tenantId, siteUrl all set' : 'One or more CONFIG values are still placeholders' });

  // 2. Graph token
  let tokenOk = false;
  try {
    await getGraphToken();
    tokenOk = true;
    results.push({ name: 'Graph token (User.Read + Sites.ReadWrite.All)', status: 'ok', msg: 'Token acquired successfully' });
  } catch(e) {
    results.push({ name: 'Graph token', status: 'error', msg: e.message.slice(0, 100) });
  }

  // 3. SharePoint site reachable
  let siteOk = false;
  if (tokenOk && cfgOk) {
    try {
      _graphSiteId = null; // force fresh resolution
      await getGraphSiteId();
      siteOk = true;
      results.push({ name: 'SharePoint site', status: 'ok', msg: `Resolved: ${CONFIG.siteUrl}` });
    } catch(e) {
      results.push({ name: 'SharePoint site', status: 'error', msg: e.message.slice(0, 100) });
    }
  }

  // 4. FT_Notifications list accessible (PA flow watches this for email delivery)
  if (siteOk) {
    try {
      await getListItems(LISTS.notifications);
      results.push({ name: 'FT_Notifications list', status: 'ok',
        msg: 'Accessible — Power Automate flow will pick up notification records' });
    } catch(e) {
      results.push({ name: 'FT_Notifications list', status: 'error',
        msg: 'Not found — create this list in SharePoint and set up the PA notify flow' });
    }
  }

  // 5. Each SharePoint list accessible
  if (siteOk) {
    const listChecks = [
      LISTS.tasks, LISTS.users, LISTS.steps, LISTS.comments,
      LISTS.signOffs, LISTS.locks, LISTS.attachments, LISTS.quarterDates,
    ];
    for (const listName of listChecks) {
      try {
        await getListItems(listName);
        results.push({ name: `List: ${listName}`, status: 'ok', msg: 'Accessible' });
      } catch(e) {
        results.push({ name: `List: ${listName}`, status: 'error', msg: 'Not found or inaccessible — check list name in SharePoint' });
      }
    }
  }

  el.innerHTML = results.map(r => `
    <div style="display:flex;align-items:center;gap:10px;padding:6px 0;border-bottom:0.5px solid var(--border)">
      <span style="font-size:14px">${r.status==='ok'?'✅':r.status==='skip'?'⚪':'❌'}</span>
      <div>
        <div style="font-size:12px;font-weight:500;color:var(--text)">${escHtml(r.name)}</div>
        <div style="font-size:11px;color:${r.status==='ok'?'var(--green-600)':r.status==='skip'?'var(--text-faint)':'var(--red-500)'}">${escHtml(r.msg)}</div>
      </div>
    </div>`).join('');
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 9 — EXCEL DATA IMPORT ───────────────────────────
// ═══════════════════════════════════════════════════════════════
async function handleImportFile(input) {
  const file = input.files[0]; if (!file) return;
  const preview = document.getElementById('import-preview');
  if (!preview) return;
  preview.innerHTML = '<p style="font-size:13px;color:var(--text-faint)">Reading file…</p>';

  try {
    const XLSX = await import('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/xlsx.mjs');
    const buf  = await file.arrayBuffer();
    const wb   = XLSX.read(buf, { type: 'array' });

    const readSheet = (name) => {
      const ws = wb.Sheets[name];
      if (!ws) return [];
      // Template rows: 1=title banner, 2=subtitle, 3=column headers, 4=tooltip row, 5+=data
      // range:2 tells the parser to treat row 3 (0-indexed=2) as the header row
      const all = XLSX.utils.sheet_to_json(ws, { defval: '', range: 2 });
      // Row 4 (tooltip) becomes the first parsed entry — drop it by checking for known tip text
      // Also drop completely empty rows (blank filler rows at the bottom of the template)
      return all.filter((r, i) => {
        if (i === 0) {
          const t = String(r.Title || r.FullName || '');
          // Drop tooltip row: empty title OR matches known tooltip phrases
          // Deliberately NOT using t.length > 50 — that would drop long but valid task names
          const isTooltip = t === ''
            || /e\.g\.|e\.g |^full name$|^full name,|^full name displayed|^deliverable name|^step name|^team member$|^team member,|^team member who/i.test(t);
          return !isTooltip;
        }
        return Object.values(r).some(v => String(v).trim() !== '');
      });
    };

    const users  = readSheet('FT_Users');
    const tasks  = readSheet('FT_Tasks');
    const steps  = readSheet('FT_Steps');

    // Build preview HTML
    preview.innerHTML = `
      <div style="border:1px solid var(--border);border-radius:8px;overflow:hidden;margin-bottom:14px">
        <div style="background:var(--bg-secondary);padding:10px 14px;font-size:12px;font-weight:600;color:var(--text-muted)">
          Import Preview
        </div>
        <div style="padding:12px 14px;font-size:13px">
          <div style="display:flex;gap:20px;margin-bottom:10px">
            <span>👤 <b>${users.length}</b> users</span>
            <span>📋 <b>${tasks.length}</b> tasks</span>
            <span>📝 <b>${steps.length}</b> steps</span>
          </div>
          ${users.length ? `<p style="color:var(--gray-500);font-size:12px">First user: ${escHtml(String(users[0].Title||users[0].FullName||''))}</p>` : ''}
          ${tasks.length ? `<p style="color:var(--gray-500);font-size:12px">First task: ${escHtml(String(tasks[0].Title||''))}</p>` : ''}
        </div>
      </div>
      <p style="font-size:12px;color:var(--red-500);margin-bottom:10px">⚠ This will ADD new rows to your SharePoint lists. Existing rows will not be changed.</p>
      <button class="btn-primary" id="import-run-btn">Import ${users.length+tasks.length+steps.length} rows →</button>
    `;
    // Bind parsed data via closure rather than injecting JSON into onclick attribute
    const btn = document.getElementById('import-run-btn');
    if (btn) btn.addEventListener('click', () => runImport(users, tasks, steps));
  } catch(e) {
    preview.innerHTML = `<p style="color:var(--red-500);font-size:13px">Could not read file: ${escHtml(e.message)}. Make sure you are uploading the Financial Reporting Tracker Import Template.</p>`;
  }
}

async function runImport(users, tasks, steps) {
  const preview = document.getElementById('import-preview');
  if (!preview) return;
  showLoadingOverlay(true);
  let imported = 0, errors = 0;

  // Writes in batches of 8 concurrent Graph POSTs — reduces import time from
  // ~2 minutes (sequential) to ~15 seconds for a typical 400-row dataset.
  // Batch size of 8 keeps well under Graph's 10k req/10min throttle limit.
  async function writeBatch(listName, fieldsList) {
    const BATCH = 8;
    for (let i = 0; i < fieldsList.length; i += BATCH) {
      const chunk = fieldsList.slice(i, i + BATCH);
      const results = await Promise.allSettled(
        chunk.map(fields => createListItem(listName, fields))
      );
      results.forEach(r => {
        if (r.status === 'fulfilled') imported++;
        else errors++;
      });
    }
  }

  try {
    const userFields = users.map(u => ({
      Title: String(u.Title||u.FullName||''), FullName: String(u.FullName||u.Title||''),
      JobRole: String(u.JobRole||''), Email: String(u.Email||''),
      IsAdmin: String(u.IsAdmin||'No'), UserId: String(u.UserId||uid()),
    }));
    const taskFields = tasks.map(t => ({
      Title: String(t.Title||''), TaskId: String(t.TaskId||uid()),
      TaskType: String(t.TaskType||'Other'), Quarter: String(t.Quarter||'Q1'),
      Year: String(t.Year||new Date().getFullYear()), Status: String(t.Status||'Not Started'),
      OwnerId: String(t.OwnerId||''), DueDate: t.DueDate||null,
      Applicability: String(t.Applicability||'All Quarters'),
      WorkdayNum: t.WorkdayNum ? String(t.WorkdayNum) : '',
      Description: String(t.Description||''), ReviewerId: String(t.ReviewerId||''),
      Reviewer2Id: String(t.Reviewer2Id||''), SkipNextRollforward: 'No',
    }));
    const stepFields = steps.map(s => ({
      Title: String(s.Title||''), StepId: String(s.StepId||uid()),
      TaskId: String(s.TaskId||''), StepOrder: String(s.StepOrder||1),
      Status: String(s.Status||'Not Started'), OwnerId: String(s.OwnerId||''),
      DueDate: s.DueDate||null,
      Applicability: String(s.Applicability||'All Quarters'),
      WorkdayNum: s.WorkdayNum ? String(s.WorkdayNum) : '',
      Note: String(s.Note||''), RequiresPrev: String(s.RequiresPrev||'No'),
      ReviewerId: String(s.ReviewerId||''), Reviewer2Id: String(s.Reviewer2Id||''),
    }));

    // Users first (tasks reference them), tasks second (steps reference them), steps last
    await writeBatch(LISTS.users, userFields);
    await writeBatch(LISTS.tasks, taskFields);
    await writeBatch(LISTS.steps, stepFields);

    await refreshData();
    preview.innerHTML = `
      <div style="background:#D1FAE5;border:1px solid #6EE7B7;border-radius:8px;padding:12px 16px;font-size:13px;color:#065F46">
        ✅ Import complete — ${imported} rows added${errors
          ? `, <span style="color:#B45309">${errors} errors skipped (check browser console for details)</span>`
          : ''}.
      </div>`;
  } catch(e) {
    preview.innerHTML = `<p style="color:var(--red-500);font-size:13px">Import failed: ${escHtml(e.message)}</p>`;
  } finally { showLoadingOverlay(false); }
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 10 — NEW QUARTER WIZARD ─────────────────────────
// ═══════════════════════════════════════════════════════════════
let _qwizState = {};

function openQuarterWizard() {
  _qwizState = { step: 1 };
  const cur = new Date();
  const curQ = ['Q1','Q2','Q3','Q4'][Math.floor(cur.getMonth()/3)];
  const curY = cur.getFullYear();
  // Default: roll from current quarter to next
  const nextIdx = ['Q1','Q2','Q3','Q4'].indexOf(curQ);
  const nextQ   = ['Q2','Q3','Q4','Q1'][nextIdx];
  const nextY   = nextQ==='Q1' ? curY+1 : curY;
  _qwizState.fromQ = curQ; _qwizState.fromY = curY;
  _qwizState.toQ   = nextQ; _qwizState.toY   = nextY;
  renderQuarterWizardStep();
  document.getElementById('quarter-wizard').classList.remove('hidden');
}

function renderQuarterWizardStep() {
  const body  = document.getElementById('qwiz-body');
  const title = document.getElementById('qwiz-title');
  if (!body) return;
  const { step } = _qwizState;

  if (step === 1) {
    title.textContent = '🔄 Step 1 of 4 — Choose Quarters';
    const curY = new Date().getFullYear();
    body.innerHTML = `
      <p style="font-size:13px;color:var(--text-muted);margin-bottom:20px">Select the quarter you are rolling <b>from</b> and the new quarter you are rolling <b>into</b>.</p>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:24px">
        <div>
          <h4 style="font-size:13px;font-weight:600;margin-bottom:10px;color:var(--navy)">From (source quarter)</h4>
          <div class="form-group"><label>Quarter</label>
            <select id="qwiz-from-q" class="filter-select" style="width:100%">
              ${['Q1','Q2','Q3','Q4'].map(q=>`<option value="${q}" ${_qwizState.fromQ===q?'selected':''}>${q}</option>`).join('')}
            </select></div>
          <div class="form-group"><label>Year</label>
            <select id="qwiz-from-y" class="filter-select" style="width:100%">
              ${[-1,0,1].map(d=>`<option value="${curY+d}" ${_qwizState.fromY===curY+d?'selected':''}>${curY+d}</option>`).join('')}
            </select></div>
        </div>
        <div>
          <h4 style="font-size:13px;font-weight:600;margin-bottom:10px;color:var(--navy)">To (new quarter)</h4>
          <div class="form-group"><label>Quarter</label>
            <select id="qwiz-to-q" class="filter-select" style="width:100%">
              ${['Q1','Q2','Q3','Q4'].map(q=>`<option value="${q}" ${_qwizState.toQ===q?'selected':''}>${q}</option>`).join('')}
            </select></div>
          <div class="form-group"><label>Year</label>
            <select id="qwiz-to-y" class="filter-select" style="width:100%">
              ${[-1,0,1,2].map(d=>`<option value="${curY+d}" ${_qwizState.toY===curY+d?'selected':''}>${curY+d}</option>`).join('')}
            </select></div>
        </div>
      </div>
      <div style="display:flex;justify-content:flex-end">
        <button class="btn-primary" onclick="qwizNext1()">Next → Review Tasks</button>
      </div>`;

  } else if (step === 2) {
    title.textContent = `🔄 Step 2 of 4 — Review Tasks & Notes`;
    const tasks = getTasks().filter(t =>
      t.quarter===_qwizState.fromQ && t.year===_qwizState.fromY &&
      !t.skipNextRollforward &&
      appliesToQuarter(t.applicability, _qwizState.toQ)
    );
    _qwizState.selectedTaskIds    = tasks.map(t=>t._spId);
    _qwizState.carryForwardNotes  = {}; // { taskSpId: true/false, stepSpId: true/false }

    // Pre-populate carry-forward: checked = true for any task/step that has notes
    tasks.forEach(t => {
      if (t.notes||t.description) _qwizState.carryForwardNotes[t._spId] = true;
      getStepsForTask(t._spId).forEach(s => {
        if (s.notes||s.note) _qwizState.carryForwardNotes[s._spId] = true;
      });
    });

    body.innerHTML = `
      <p style="font-size:13px;color:var(--text-muted);margin-bottom:4px">
        <b>${tasks.length}</b> tasks will be copied to ${_qwizState.toQ} ${_qwizState.toY}.
        Deselect tasks to skip them. For tasks and steps with notes, choose whether to carry the note forward.
      </p>
      <p style="font-size:11px;color:var(--text-faint);margin-bottom:14px">Notes carry forward with a label showing which quarter they came from.</p>
      <div style="max-height:400px;overflow-y:auto;border:0.5px solid var(--border);border-radius:8px;margin-bottom:16px">
        ${tasks.length ? tasks.map(t => {
          const taskNote   = t.notes || t.description || '';
          const taskSteps  = getStepsForTask(t._spId).filter(s => appliesToQuarter(s.applicability, _qwizState.toQ));
          const stepsWithNotes = taskSteps.filter(s => s.notes || s.note);
          return `
          <div style="border-bottom:0.5px solid var(--border)">
            <div style="display:flex;align-items:flex-start;gap:10px;padding:10px 14px;background:var(--bg-secondary)">
              <input type="checkbox" value="${t._spId}" checked
                onchange="qwizToggleTask('${t._spId}',this.checked)"
                style="width:15px;height:15px;margin-top:2px;flex-shrink:0" />
              <div style="flex:1;min-width:0">
                <div style="font-size:13px;font-weight:600">${escHtml(t.name)}</div>
                <div style="font-size:11px;color:var(--text-faint)">${escHtml(t.type)} · ${escHtml(t.status)}</div>
                ${taskNote ? `
                <div style="margin-top:6px;background:#fff;border:1px solid var(--border);border-radius:6px;padding:7px 10px;font-size:12px;color:var(--text-muted)">
                  <span style="font-weight:600;color:var(--navy);font-size:10px;text-transform:uppercase;letter-spacing:.04em">Task note: </span>${escHtml(taskNote)}
                  <label style="display:flex;align-items:center;gap:6px;margin-top:6px;cursor:pointer;font-size:11px;color:var(--text-muted)">
                    <input type="checkbox" ${_qwizState.carryForwardNotes[t._spId]?'checked':''}
                      onchange="_qwizState.carryForwardNotes['${t._spId}']=this.checked"
                      style="width:13px;height:13px" />
                    Carry this note forward to ${_qwizState.toQ} ${_qwizState.toY}
                  </label>
                </div>` : ''}
              </div>
            </div>
            ${taskSteps.map((s,si) => `
            <div style="display:flex;align-items:flex-start;gap:10px;padding:8px 14px 8px 36px;border-top:0.5px solid #F0F2F7">
              <div style="font-size:11px;color:var(--text-faint);min-width:16px;margin-top:2px">${s.order}.</div>
              <div style="flex:1;min-width:0">
                <div style="font-size:12px;font-weight:500">${escHtml(s.name)}</div>
                ${(s.notes||s.note) ? `
                <div style="margin-top:5px;background:#F9FAFB;border:1px solid var(--border);border-radius:5px;padding:6px 9px;font-size:11px;color:var(--text-muted)">
                  ${escHtml(s.notes||s.note)}
                  <label style="display:flex;align-items:center;gap:6px;margin-top:5px;cursor:pointer;font-size:11px;color:var(--text-muted)">
                    <input type="checkbox" ${_qwizState.carryForwardNotes[s._spId]?'checked':''}
                      onchange="_qwizState.carryForwardNotes['${s._spId}']=this.checked"
                      style="width:13px;height:13px" />
                    Carry this note forward
                  </label>
                </div>` : ''}
              </div>
            </div>`).join('')}
          </div>`;
        }).join('') : '<p style="padding:20px;text-align:center;color:var(--text-faint);font-size:13px">No tasks found in source quarter.</p>'}
      </div>
      <div style="display:flex;justify-content:space-between">
        <button class="btn-secondary" onclick="_qwizState.step=1;renderQuarterWizardStep()">← Back</button>
        <button class="btn-primary" onclick="qwizValidateAndNextStep()">Next → Set Close Calendar</button>
      </div>`;

  } else if (step === 3) {
    title.textContent = `🔄 Step 3 of 4 — Set Close Calendar`;
    const existing = loadCloseCalendar(_qwizState.toQ, _qwizState.toY);
    body.innerHTML = `
      <p style="font-size:13px;color:var(--text-muted);margin-bottom:16px">
        Set Workday 1 for ${_qwizState.toQ} ${_qwizState.toY} — the first working day of your close process. This drives all workday-based due dates.
      </p>
      <div class="form-group" style="max-width:280px">
        <label>Workday 1 date</label>
        <input type="date" id="qwiz-wd1" value="${existing?.wd1Date||''}" />
      </div>
      <div id="qwiz-cal-preview" style="margin-top:16px"></div>
      <div style="display:flex;justify-content:space-between;margin-top:16px">
        <button class="btn-secondary" onclick="_qwizState.step=2;renderQuarterWizardStep()">← Back</button>
        <button class="btn-primary" onclick="qwizNext3()">Next → Set Key Dates</button>
      </div>`;

    // Attach input listener now that innerHTML is set (inline <script> tags don't execute)
    const wd1Input = document.getElementById('qwiz-wd1');
    if (wd1Input) {
      wd1Input.addEventListener('input', function() {
        renderEditableCalGrid('qwiz-cal-preview', this.value, 20, true);
      });
      if (wd1Input.value) renderEditableCalGrid('qwiz-cal-preview', wd1Input.value, 20, true);
    }

  } else if (step === 4) {
    title.textContent = `🔄 Step 4 of 4 — Key Dates & Confirm`;
    const existing = _quarterDates.find(d=>d.quarter===_qwizState.toQ&&d.year===_qwizState.toY);
    body.innerHTML = `
      <p style="font-size:13px;color:var(--text-muted);margin-bottom:16px">
        Optional — set external key dates for ${_qwizState.toQ} ${_qwizState.toY}.
        These show as a countdown strip on the dashboard.
      </p>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px">
        <div class="form-group"><label>SEC Filing Date</label><input type="date" id="qwiz-sec" value="${existing?.secFilingDate||''}" /></div>
        <div class="form-group"><label>Earnings Call Date</label><input type="date" id="qwiz-earn" value="${existing?.earningsDate||''}" /></div>
        <div class="form-group"><label>Earnings Call Time</label><input type="text" id="qwiz-earntime" value="${escHtml(existing?.earningsTime||'')}" placeholder="e.g. 8:30 AM ET" /></div>

      </div>
      <div style="background:var(--bg-secondary);border-radius:8px;padding:14px;margin-bottom:16px;font-size:13px">
        <b>Summary:</b> Copying ${_qwizState.selectedTaskIds?.length||0} tasks from ${_qwizState.fromQ} ${_qwizState.fromY} → ${_qwizState.toQ} ${_qwizState.toY}
      </div>
      <div style="background:#FEF3C7;border:1px solid #FDE68A;border-radius:8px;padding:12px 14px;margin-bottom:16px;font-size:12px;color:#92400E">
        🔒 <b>Note:</b> ${_qwizState.fromQ} ${_qwizState.fromY} will be <b>locked</b> after rollforward — team members will no longer be able to edit those tasks. You can unlock it in Admin if needed.
      </div>
      <div style="display:flex;justify-content:space-between">
        <button class="btn-secondary" onclick="_qwizState.step=3;renderQuarterWizardStep()">← Back</button>
        <button class="btn-primary" onclick="runQuarterWizard()">🚀 Run Rollforward</button>
      </div>`;
  }
}

function qwizNext1() {
  _qwizState.fromQ = document.getElementById('qwiz-from-q').value;
  _qwizState.fromY = parseInt(document.getElementById('qwiz-from-y').value);
  _qwizState.toQ   = document.getElementById('qwiz-to-q').value;
  _qwizState.toY   = parseInt(document.getElementById('qwiz-to-y').value);
  if (_qwizState.fromQ===_qwizState.toQ && _qwizState.fromY===_qwizState.toY) {
    showToast('From and To quarters cannot be the same.', 'warning'); return;
  }
  if (isQuarterLocked(_qwizState.toQ, _qwizState.toY)) {
    showToast(`${_qwizState.toQ} ${_qwizState.toY} is already locked. Unlock it in Admin before rolling forward into it.`, 'warning'); return;
  }
  const srcTasks = getTasks().filter(t => t.quarter===_qwizState.fromQ && t.year===_qwizState.fromY);
  if (!srcTasks.length) {
    showToast(`No tasks found in ${_qwizState.fromQ} ${_qwizState.fromY}. Nothing to roll forward.`, 'warning'); return;
  }
  _qwizState.step = 2;
  renderQuarterWizardStep();
}

function qwizToggleTask(spId, checked) {
  if (checked) {
    if (!_qwizState.selectedTaskIds.includes(spId)) _qwizState.selectedTaskIds.push(spId);
  } else {
    _qwizState.selectedTaskIds = _qwizState.selectedTaskIds.filter(id=>id!==spId);
  }
}

// Validates step 2 of the quarter wizard before advancing to step 3.
// Extracted from inline onclick to keep the template readable.
function qwizValidateAndNextStep() {
  if (!_qwizState.selectedTaskIds?.length) {
    showToast('No tasks selected — select at least one task to roll forward.', 'warning');
    return;
  }
  _qwizState.step = 3;
  renderQuarterWizardStep();
}

function qwizNext3() {
  const wd1 = document.getElementById('qwiz-wd1')?.value;
  if (!wd1) { showToast('Please set Workday 1 before continuing.', 'warning'); return; }
  // Store WD1 in wizard state but do NOT save to localStorage yet.
  // The calendar is saved inside runQuarterWizard() so that hitting
  // Back from step 4 or closing the wizard does not leave a saved
  // calendar for a quarter whose rollforward was never confirmed.
  _qwizState.wd1 = wd1;
  _qwizState.step = 4;
  renderQuarterWizardStep();
}

async function runQuarterWizard() {
  const { fromQ, fromY, toQ, toY, selectedTaskIds } = _qwizState;
  const qEnd = quarterEndDate(toQ, toY);

  // Save key dates if entered
  const secDate   = document.getElementById('qwiz-sec')?.value   || null;
  const earnDate  = document.getElementById('qwiz-earn')?.value  || null;
  const earnTime  = document.getElementById('qwiz-earntime')?.value || '';
  if (secDate||earnDate) {
    // Save quarter dates inline (bypasses modal)
    const qdFields = {
      Title: `${toQ} ${toY}`,
      Quarter: toQ, Year: String(toY),
      SECFilingDate: secDate, EarningsCallDate: earnDate,
      EarningsCallTime: earnTime,
    };
    const existQd = _quarterDates.find(d=>d.quarter===toQ&&d.year===toY);
    try {
      if (existQd) { await updateListItem(LISTS.quarterDates, existQd._spId, qdFields); Object.assign(existQd, normaliseQuarterDate({...qdFields,id:existQd._spId})); }
      else { const c = await createListItem(LISTS.quarterDates, qdFields); _quarterDates.push(normaliseQuarterDate({...qdFields,id:c?.id||uid()})); }
    } catch(e) { console.warn('Could not save key dates:', e.message); }
  }

  document.getElementById('quarter-wizard').classList.add('hidden');
  showLoadingOverlay(true);

  try {
    // Save close calendar now that rollforward is confirmed — deferred from qwizNext3
    if (_qwizState.wd1) saveCloseCalendar(_qwizState.toQ, _qwizState.toY, _qwizState.wd1);

    const selectedTasks = getTasks().filter(t => selectedTaskIds.includes(t._spId));
    if (!selectedTasks.length) {
      showLoadingOverlay(false);
      showToast('No tasks selected — rollforward cancelled. The source quarter was not locked.', 'warning');
      return;
    }
    // _copyTasksToQuarter throws and rolls back if any task fails — lockQuarter
    // is only reached when the full copy succeeded. This ordering is intentional:
    // never lock the source quarter if the copy did not complete cleanly.
    const copied = await _copyTasksToQuarter(selectedTasks, fromQ, fromY, toQ, toY, _qwizState.carryForwardNotes || {});
    await lockQuarter(fromQ, fromY);
    await refreshData();
    showToast(`${copied} task${copied!==1?'s':''} copied to ${toQ} ${toY}. ${fromQ} ${fromY} locked.`, 'success', TOAST_DURATION_LONG);
  } catch(e) { showToast('Quarter wizard failed: '+e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 12 — FULL AUDIT LOG EXPORT ──────────────────────
// ═══════════════════════════════════════════════════════════════
function exportFullAuditLog(scope) {
  let tasks = getTasks();
  let label = 'AllQuarters';
  if (scope === 'quarter') {
    const q  = document.getElementById('audit-quarter-select')?.value || 'Q1';
    const yr = parseInt(document.getElementById('audit-year-select')?.value || new Date().getFullYear());
    tasks = tasks.filter(t => t.quarter===q && t.year===yr);
    label = `${q}_${yr}`;
  }

  // Sign-offs
  const soRows = [];
  tasks.forEach(t => {
    // include refType as element [2] so we can split status vs date changes below
    getSignOffsFor(t._spId).forEach(s => soRows.push([t.quarter, t.year, s.refType||'task', t.name, s.userName, s.fromStatus, s.toStatus, s.ts, s.tsIso]));
    getStepsForTask(t._spId).forEach(step => {
      getSignOffsFor(step._spId).forEach(s => soRows.push([t.quarter, t.year, s.refType||'step', step.name, s.userName, s.fromStatus, s.toStatus, s.ts, s.tsIso]));
    });
  });

  // Comments
  const allComments = _comments.filter(c => tasks.some(t => t._spId===c.taskId || t.id===c.taskId));
  const commentRows = allComments.map(c => {
    const task = tasks.find(t=>t._spId===c.taskId||t.id===c.taskId);
    const cAuthor = getUserById(c.authorId)?.name || (c.authorId ? 'Former team member' : 'Unknown');
    return [task?.quarter||'', task?.year||'', task?.name||'', c.stepId?'Step':'Task', cAuthor, c.text, c.time];
  });

  // Tasks summary
  const taskRows = tasks.map(t => [t.quarter, t.year, t.name, t.type, t.status,
    getUserById(t.ownerId)?.name||t.ownerId,
    getUserById(t.reviewerId)?.name||'',
    getUserById(t.reviewer2Id)?.name||'',
    t.dueDate, t.applicability]);

  // Separate due_date_change sign-offs into their own section for clarity
  const statusChanges = soRows.filter(r => r[2] !== 'due_date_change');
  const dateChanges   = soRows.filter(r => r[2] === 'due_date_change').map(r =>
    [r[0], r[1], r[3], r[4], r[5], r[6], r[7], r[8]]  // Quarter, Year, Item, Changed By, Old Date, New Date, Display, ISO
  );

  const shHtml = (header, rows) => `<table>
    <thead><tr>${header.map(h=>`<th style="background:#0f2140;color:#fff;font-weight:bold;padding:6px 10px;font-size:12px">${h}</th>`).join('')}</tr></thead>
    <tbody>
      ${rows.map((r,i) => `
        <tr style="background:${i%2 ? '#f7f9fc' : '#fff'}">
          ${r.map(c => `<td style="padding:6px 10px;font-size:12px;border:1px solid #dde3ed">${escHtml(String(c||''))}</td>`).join('')}
        </tr>`).join('')}
    </tbody>
  </table>`;

  const wb = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>
  <x:ExcelWorksheet><x:Name>Tasks</x:Name><x:WorksheetOptions><x:Selected/></x:WorksheetOptions></x:ExcelWorksheet>
  <x:ExcelWorksheet><x:Name>Sign-Off Log</x:Name><x:WorksheetOptions/></x:ExcelWorksheet>
  <x:ExcelWorksheet><x:Name>Date Changes</x:Name><x:WorksheetOptions/></x:ExcelWorksheet>
  <x:ExcelWorksheet><x:Name>Comments</x:Name><x:WorksheetOptions/></x:ExcelWorksheet>
</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
<style>table{border-collapse:collapse}</style></head><body>
${shHtml(['Quarter','Year','Task','Type','Status','Preparer','Reviewer 1','Reviewer 2','Due Date','Applicability'], taskRows)}
${shHtml(['Quarter','Year','Item','Changed By','From Status','To Status','Display Time','ISO Time'], statusChanges)}
${shHtml(['Quarter','Year','Item','Changed By','Old Date','New Date','Display Time','ISO Time'], dateChanges)}
${shHtml(['Quarter','Year','Task','Level','Author','Comment','Time'], commentRows)}
</body></html>`;

  const blob = new Blob([wb], {type:'application/vnd.ms-excel;charset=utf-8'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href=url; a.download=`AuditLog_${label}_FRT.xls`;
  a.click(); URL.revokeObjectURL(url);
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 9 — STICKY QUARTER ──────────────────────────────
// ═══════════════════════════════════════════════════════════════
function saveQuarterPref(q, yr) {
  try { sessionStorage.setItem('ft_quarter', q); sessionStorage.setItem('ft_year', String(yr)); } catch(e) {}
}
function loadQuarterPref() {
  try { return { q: sessionStorage.getItem('ft_quarter'), yr: sessionStorage.getItem('ft_year') }; } catch(e) { return {}; }
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 5 — COLOUR-CODED OWNERS ─────────────────────────
// ═══════════════════════════════════════════════════════════════
const OWNER_COLOURS = [
  {bg:'#EDE9FE',text:'#5B21B6'}, {bg:'#CCFBF1',text:'#0F766E'},
  {bg:'#FEE2E2',text:'var(--red-dark)'}, {bg:'#FEF9C3',text:'#713F12'},
  {bg:'#DBEAFE',text:'#1D4ED8'}, {bg:'#FCE7F3',text:'#9D174D'},
  {bg:'#D1FAE5',text:'#065F46'}, {bg:'#FED7AA',text:'#C2410C'},
  {bg:'#E0E7FF',text:'#3730A3'}, {bg:'#F0FDF4',text:'#166534'},
];
const _ownerColorMap = {};
function getOwnerColour(userId) {
  if (!userId) return OWNER_COLOURS[0];
  if (!_ownerColorMap[userId]) {
    const idx = Object.keys(_ownerColorMap).length % OWNER_COLOURS.length;
    _ownerColorMap[userId] = OWNER_COLOURS[idx];
  }
  return _ownerColorMap[userId];
}
function buildOwnerColorMap() {
  // Clear existing keys without reassigning the const reference
  Object.keys(_ownerColorMap).forEach(k => delete _ownerColorMap[k]);
  _users.forEach((u, i) => {
    _ownerColorMap[u._spId || u.id] = OWNER_COLOURS[i % OWNER_COLOURS.length];
  });
}
function coloredAvatar(userId, name, size=28) {
  const col = getOwnerColour(userId);
  return `<div class="mini-avatar" style="width:${size}px;height:${size}px;font-size:${Math.round(size*0.38)}px;background:${col.bg};color:${col.text}">${initials(name||'?')}</div>`;
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 2 — UNDO TOAST ──────────────────────────────────
// ═══════════════════════════════════════════════════════════════
let _undoTimer = null;
let _undoAction = null;
let _undoSourceView = null; // the view that was active when the undoable action was taken

function showUndoToast(message, undoFn) {
  if (_undoTimer) { clearTimeout(_undoTimer); removeUndoToast(); }
  _undoAction = undoFn;
  // Remember which view triggered this so we can navigate back on undo
  const activeView = document.querySelector('.view.active');
  _undoSourceView = activeView ? activeView.id.replace('view-', '') : null;
  let toast = document.getElementById('undo-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'undo-toast';
    toast.className = 'undo-toast';
    document.body.appendChild(toast);
  }
  toast.innerHTML = `<span class="undo-msg">${escHtml(message)}</span>
    <button class="undo-btn" onclick="executeUndo()">↩ Undo</button>
    <button class="undo-close" onclick="removeUndoToast()">✕</button>`;
  toast.classList.add('visible');
  _undoTimer = setTimeout(removeUndoToast, UNDO_DURATION);
}
function removeUndoToast() {
  const t = document.getElementById('undo-toast');
  if (t) { t.classList.remove('visible'); }
  _undoAction = null;
  if (_undoTimer) { clearTimeout(_undoTimer); _undoTimer = null; }
}
async function executeUndo() {
  if (_undoAction) {
    const sourceView = _undoSourceView;
    removeUndoToast();
    await _undoAction();
    // Navigate back to the view where the action originated if user has moved away
    if (sourceView && !document.getElementById('view-' + sourceView)?.classList.contains('active')) {
      const navEl = document.querySelector(`[data-view="${sourceView}"]`);
      switchView(sourceView, navEl);
    }
  }
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 8 — DUE DATE CHANGE HISTORY ─────────────────────
// ═══════════════════════════════════════════════════════════════
// Stored in sign-off log with refType='due_date_change'
async function recordDueDateChange(spId, itemType, itemName, oldDate, newDate) {
  if (oldDate === newDate) return;
  const signOffId = uid(); // single ID shared between SharePoint record and local cache
  const entry = {
    Title:        `Due date changed: ${itemName}`,
    SignOffId:    signOffId,
    RefId:        String(spId),
    RefType:      'due_date_change',
    RefName:      itemName,
    UserId:       currentUserId(),
    UserName:     currentUser.name,
    FromStatus:   oldDate || '(none)',
    ToStatus:     newDate || '(none)',
    Timestamp:    nowLabel(),
    TimestampISO: nowISO(),
  };
  // Optimistically push to local cache first
  _signOffs.unshift(normaliseSignOff({ ...entry, id: signOffId }));
  try {
    const created = await createListItem(LISTS.signOffs, entry);
    // Patch the local cache entry with the real SharePoint ID once available
    const cached = _signOffs.find(s => s.id === signOffId);
    if (cached && created?.id) cached._spId = created.id;
  } catch(e) { console.warn('Could not record due date change:', e.message); }
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 3 — GLOBAL SEARCH ───────────────────────────────
// ═══════════════════════════════════════════════════════════════
function openGlobalSearch() {
  let overlay = document.getElementById('global-search-overlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'global-search-overlay';
    overlay.className = 'gs-overlay';
    overlay.innerHTML = `
      <div class="gs-box" onclick="event.stopPropagation()">
        <div class="gs-input-wrap">
          <span class="gs-icon">🔍</span>
          <input type="text" id="gs-input" class="gs-input" placeholder="Search tasks, steps, comments, users…" oninput="runGlobalSearch(this.value)" autocomplete="off" />
          <button class="gs-close" onclick="closeGlobalSearch()">✕</button>
        </div>
        <div id="gs-results" class="gs-results"></div>
      </div>`;
    overlay.addEventListener('click', closeGlobalSearch);
    document.body.appendChild(overlay);
  }
  overlay.classList.add('visible');
  setTimeout(() => document.getElementById('gs-input')?.focus(), 50);
}

function closeGlobalSearch() {
  const overlay = document.getElementById('global-search-overlay');
  if (overlay) overlay.classList.remove('visible');
}

function runGlobalSearch(q) {
  const el = document.getElementById('gs-results'); if(!el) return;
  q = (q||'').toLowerCase().trim();
  if (!q || q.length < 2) { el.innerHTML = '<div class="gs-hint">Type at least 2 characters to search…</div>'; return; }

  const results = [];

  // Tasks
  getActiveTasks().filter(t => t.name.toLowerCase().includes(q) || (t.description||'').toLowerCase().includes(q) || (t.notes||'').toLowerCase().includes(q))
    .slice(0,5).forEach(t => results.push({
      type:'Task', icon:'📋', label:t.name, sub:`${t.type} · ${t.quarter} ${t.year} · ${statusBadgeClass(t.status)?t.status:''}`,
      action: `closeGlobalSearch();
               switchView('dashboard',document.querySelector('[data-view=\'dashboard\']'));
               document.getElementById('quarter-filter').value='${t.quarter}';
               document.getElementById('year-filter').value='${t.year}';
               renderDashboard();`
    }));

  // Steps
  _steps.filter(s => s.name.toLowerCase().includes(q) || (s.notes||s.note||'').toLowerCase().includes(q))
    .slice(0,5).forEach(s => {
      const task = _tasks.find(t=>t._spId===s.taskId);
      results.push({ type:'Step', icon:'📝', label:s.name, sub:`Step in: ${task?.name||'Unknown task'}`,
        action: `openSteps('${s.taskId}','${escHtml(task?.name||'')}');closeGlobalSearch();` });
    });

  // Comments
  _comments.filter(c => c.text.toLowerCase().includes(q))
    .slice(0,3).forEach(c => {
      const task = _tasks.find(t=>t._spId===c.taskId||t.id===c.taskId);
      results.push({ type:'Comment', icon:'💬', label:c.text.slice(0,60)+(c.text.length>60?'…':''),
        sub:`By ${getUserById(c.authorId)?.name||'Former team member'} on ${task?.name||'Unknown task'}`,
        action: `openComments('${c.taskId}','${escHtml(task?.name||'')}');closeGlobalSearch();` });
    });

  // Users
  _users.filter(u => u.name.toLowerCase().includes(q) || u.role.toLowerCase().includes(q) || (u.email||'').toLowerCase().includes(q))
    .slice(0,3).forEach(u => results.push({
      type:'User', icon:'👤', label:u.name, sub:`${u.role}${u.isAdmin?' · Admin':''}`,
      action: `closeGlobalSearch();`
    }));

  if (!results.length) { el.innerHTML = `<div class="gs-hint">No results for "<b>${escHtml(q)}</b>"</div>`; return; }

  el.innerHTML = results.map(r => `
    <div class="gs-result" onclick="${r.action}">
      <span class="gs-result-icon">${r.icon}</span>
      <div class="gs-result-body">
        <div class="gs-result-label">${escHtml(r.label)}</div>
        <div class="gs-result-sub">${escHtml(r.sub)}</div>
      </div>
      <span class="gs-result-type">${r.type}</span>
    </div>`).join('');
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 4 — SAVED FILTERS ───────────────────────────────
// ═══════════════════════════════════════════════════════════════
const MAX_SAVED_FILTERS = 3;

function getSavedFilters() {
  try { return JSON.parse(localStorage.getItem('ft_saved_filters')||'[]'); } catch(e) { return []; }
}
function setSavedFilters(filters) {
  try { localStorage.setItem('ft_saved_filters', JSON.stringify(filters)); } catch(e) {}
}

function saveCurrentFilter() {
  const filters = getSavedFilters();
  if (filters.length >= MAX_SAVED_FILTERS) {
    showToast(`You can save up to ${MAX_SAVED_FILTERS} filters. Delete one first.`, 'warning'); return;
  }
  const q    = document.getElementById('quarter-filter')?.value || 'Q1';
  const yr   = document.getElementById('year-filter')?.value || new Date().getFullYear();
  const type = activeTypeFilter;
  const filter = activeFilter;
  document.getElementById('modal-title').textContent = 'Save Filter';
  document.getElementById('modal-body').innerHTML = `
    <p style="font-size:13px;color:var(--text-muted);margin-bottom:14px">
      Saves the current quarter (${q} ${yr}), quick filter, and type filter as a named shortcut.
    </p>
    <div class="form-group"><label>Filter name</label>
      <input type="text" id="save-filter-name" placeholder='e.g. My Q1 overdue' style="width:100%" />
    </div>
    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeAllModals()">Cancel</button>
      <button class="btn-primary" onclick="
        const n=document.getElementById('save-filter-name')?.value.trim();
        if(!n){showToast('Please enter a name.','warning');return;}
        const f=getSavedFilters();
        f.push({name:n,q:'${q}',yr:'${yr}',type:'${type}',filter:'${filter}'});
        setSavedFilters(f);renderSavedFilters();closeAllModals();
        showToast('Filter saved: '+n,'success',2500);
      ">Save</button>
    </div>`;
  openModal();
  setTimeout(()=>document.getElementById('save-filter-name')?.focus(),50);
}

function applySavedFilter(idx) {
  const filters = getSavedFilters();
  const f = filters[idx]; if(!f) return;
  // Legacy filters saved with filter='mine' → redirect to My Tasks
  if (f.filter === 'mine') {
    switchView('mytasks', document.querySelector('[data-view="mytasks"]'));
    return;
  }
  const qEl = document.getElementById('quarter-filter');
  const yEl = document.getElementById('year-filter');
  if (qEl) qEl.value = f.q;
  if (yEl) yEl.value = f.yr;
  activeFilter     = f.filter || 'all';
  activeTypeFilter = f.type   || 'all';
  const typeEl = document.getElementById('type-filter');
  if (typeEl) typeEl.value = activeTypeFilter;
  document.querySelectorAll('.qfilter').forEach(b => {
    b.classList.toggle('active', b.dataset.filter === activeFilter);
  });
  saveQuarterPref(f.q, f.yr);
  renderDashboard();
  switchView('dashboard', document.querySelector('[data-view="dashboard"]'));
}

function deleteSavedFilter(idx) {
  const filters = getSavedFilters();
  filters.splice(idx, 1);
  setSavedFilters(filters);
  renderSavedFilters();
}

function renderSavedFilters() {
  const bar     = document.getElementById('saved-filters-bar');
  const details = document.getElementById('saved-filters-details');
  const lbl     = document.getElementById('saved-filters-summary-label');
  if (!bar) return;
  const filters = getSavedFilters();
  // Update summary label with count
  if (lbl) lbl.textContent = filters.length ? `☆ Saved filters (${filters.length})` : '☆ Saved filters';
  bar.innerHTML = filters.map((f,i) => `
    <div class="saved-filter-chip">
      <span onclick="applySavedFilter(${i})">${escHtml(f.name)}</span>
      <button onclick="deleteSavedFilter(${i})" title="Delete filter">✕</button>
    </div>`).join('');
}

// ═══════════════════════════════════════════════════════════════
// ── COMPACT MODE ─────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
// Toggles compact row display. Persists preference in localStorage.
function toggleCompactMode() {
  const isCompact = document.body.classList.toggle('compact');
  try { localStorage.setItem('ft_compact', isCompact ? '1' : '0'); } catch(e) {}
  const btn = document.getElementById('compact-toggle');
  if (btn) btn.textContent = isCompact ? '⊞ Normal' : '⊟ Compact';
}
function initCompactMode() {
  try {
    if (localStorage.getItem('ft_compact') === '1') {
      document.body.classList.add('compact');
      const btn = document.getElementById('compact-toggle');
      if (btn) btn.textContent = '⊞ Normal';
    }
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════
// ── FEATURE 11 — DRAG TO REORDER STEPS ──────────────────────
// ═══════════════════════════════════════════════════════════════
let _dragStepSpId = null;

function stepDragStart(e, stepSpId) {
  _dragStepSpId = stepSpId;
  e.dataTransfer.effectAllowed = 'move';
  e.currentTarget.classList.add('step-dragging');
}
function stepDragEnd(e) {
  e.currentTarget.classList.remove('step-dragging');
  document.querySelectorAll('.step-drag-over').forEach(el => el.classList.remove('step-drag-over'));
}
function stepDragOver(e, stepSpId) {
  e.preventDefault();
  if (_dragStepSpId === stepSpId) return;
  e.dataTransfer.dropEffect = 'move';
  document.querySelectorAll('.step-drag-over').forEach(el => el.classList.remove('step-drag-over'));
  e.currentTarget.closest('.step-row')?.classList.add('step-drag-over');
}
async function stepDrop(e, targetSpId) {
  e.preventDefault();
  document.querySelectorAll('.step-drag-over').forEach(el => el.classList.remove('step-drag-over'));
  if (!_dragStepSpId || _dragStepSpId === targetSpId) return;
  const steps = getStepsForTask(stepsTaskSpId).sort((a,b)=>(a.order||0)-(b.order||0));
  const dragIdx   = steps.findIndex(s=>s._spId===_dragStepSpId);
  const targetIdx = steps.findIndex(s=>s._spId===targetSpId);
  if (dragIdx < 0 || targetIdx < 0) return;
  // Reorder in memory
  const [moved] = steps.splice(dragIdx, 1);
  steps.splice(targetIdx, 0, moved);
  // Assign new order numbers and save
  showLoadingOverlay(true);
  try {
    for (let i=0; i<steps.length; i++) {
      steps[i].order = i+1;
      await updateListItem(LISTS.steps, steps[i]._spId, { StepOrder: String(i+1) });
    }
    renderStepsPanel();
  } catch(e) { showToast('Could not reorder steps: '+e.message, 'error'); }
  finally { showLoadingOverlay(false); _dragStepSpId = null; }
}
// ── INIT ──────────────────────────────────────────────────────
(function init() {
  // Set version in sidebar
  const vEl = document.getElementById('app-version');
  if (vEl) vEl.textContent = APP_VERSION;

  // Show current quarter on login screen
  const qCtx = document.getElementById('login-quarter-context');
  if (qCtx) {
    const now = new Date();
    const m   = now.getMonth();
    const q   = m<3?'Q1':m<6?'Q2':m<9?'Q3':'Q4';
    const yr  = now.getFullYear();
    qCtx.textContent = `${q} ${yr} Close`;
  }

  msalInstance.handleRedirectPromise().catch(console.error);

  // Check if setup is needed (after redirect promise so auth state is settled)
  showSetupWizardIfNeeded();

  // Keyboard shortcuts
  document.addEventListener('keydown', e => {
    const tag = document.activeElement?.tagName?.toLowerCase();
    const inInput = tag === 'input' || tag === 'textarea' || tag === 'select';
    // "/" opens global search (when not typing)
    if (e.key === '/' && !e.ctrlKey && !e.metaKey && !e.altKey && !inInput) {
      e.preventDefault();
      openGlobalSearch();
      return;
    }
    // Escape closes any open overlay/modal
    if (e.key === 'Escape') {
      const search = document.getElementById('global-search-overlay');
      if (search?.classList.contains('visible')) { closeGlobalSearch(); return; }
      const steps = document.getElementById('steps-overlay');
      if (steps && !steps.classList.contains('hidden')) { closeStepsPanel(); return; }
      const comments = document.getElementById('comment-overlay');
      if (comments && !comments.classList.contains('hidden')) { closeCommentModal(); return; }
      const modal = document.getElementById('modal-overlay');
      if (modal && !modal.classList.contains('hidden')) { closeAllModals(); return; }
      const checklist = document.getElementById('checklist-overlay');
      if (checklist && !checklist.classList.contains('hidden')) { exitChecklistMode(); return; }
    }
  });

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    // Already signed in with Microsoft — restore session or re-identify
    const savedId = sessionStorage.getItem('ft_session');
    if (savedId) {
      loadAllData().then(() => {
        const user = _users.find(u => (u._spId||u.id) === savedId);
        if (user) { currentUser = user; launchApp(); }
        else { afterMicrosoftLogin(); } // session expired
      }).catch(e => {
        showError('Could not restore session: ' + e.message);
      });
    } else {
      afterMicrosoftLogin(); // signed in but no session — identify and launch
    }
  }
})();

// ═══════════════════════════════════════════════════════════════
// ── MY TASKS VIEW ────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function renderMyTasks() {
  const userId = currentUserId(); // canonical ID for the current user
  const titleEl = document.getElementById('mytasks-title');
  if (titleEl) titleEl.textContent = `My Tasks — ${currentUser.name.split(' ')[0]}`;
  const container = document.getElementById('mytasks-content');
  if (!container) return;

  // Quarter filter
  const mqEl = document.getElementById('mytasks-quarter');
  const myEl = document.getElementById('mytasks-year');
  const mq   = mqEl?.value || 'all';
  const my   = myEl?.value ? parseInt(myEl.value) : 0;

  // All tasks/steps where I am preparer, reviewer 1, or reviewer 2
  let allTasks = getActiveTasks().filter(t =>
    t.status !== 'Not Applicable' &&
    (t.ownerId===userId || t.reviewerId===userId || t.reviewer2Id===userId)
  );
  if (mq !== 'all') allTasks = allTasks.filter(t => t.quarter===mq && (!my || t.year===my));
  let allSteps = _steps.filter(s => {
    if (s.status === 'Not Applicable') return false;
    if (!(s.ownerId===userId || s.reviewerId===userId || s.reviewer2Id===userId)) return false;
    const pt = _tasks.find(t => t._spId === s.taskId);
    if (!pt || isQuarterLocked(pt.quarter, pt.year)) return false; // exclude locked quarters
    return true;
  });
  if (mq !== 'all') {
    allSteps = allSteps.filter(s => {
      const pt = _tasks.find(t=>t._spId===s.taskId);
      return pt && pt.quarter===mq && (!my || pt.year===my);
    });
  }

  if (!allTasks.length && !allSteps.length) {
    container.innerHTML = `<div class="empty-state"><div class="empty-icon">🎉</div>
      <p>Nothing assigned to you right now.</p></div>`;
    return;
  }

  // ── Grouping: Ready for Me / Waiting / Complete ───────────
  function isReadyForMe(item) {
    const isOwner     = item.ownerId    === userId;
    const isReviewer1 = item.reviewerId === userId;
    const isReviewer2 = item.reviewer2Id === userId;
    if (item.status === 'Not Started' || item.status === 'In Progress') {
      if (!isOwner) return false;
      // For steps: also check whether a predecessor is blocking this step
      if (item.taskId !== undefined && getBlockingPredecessor(item)) return false;
      return true;
    }
    if (item.status === 'Ready for Review 1') return isReviewer1;
    if (item.status === 'Ready for Review 2') return isReviewer2;
    return false;
  }

  function isWaiting(item) {
    if (item.status === 'Complete') return false;
    return !isReadyForMe(item);
  }

  const taskGroups = {
    readyForMe: allTasks.filter(t => isReadyForMe(t)),
    waiting:    allTasks.filter(t => isWaiting(t)),
    complete:   allTasks.filter(t => t.status === 'Complete'),
  };
  const stepGroups = {
    readyForMe: allSteps.filter(s => isReadyForMe(s)),
    waiting:    allSteps.filter(s => isWaiting(s)),
    complete:   allSteps.filter(s => s.status === 'Complete'),
  };

  function taskRow(task) {
    const locked  = isQuarterLocked(task.quarter, task.year);
    const canEdit = !locked;
    const ds      = deadlineStatus(task.dueDate, task.status);
    const myRole  = task.ownerId===userId ? 'Preparer' : task.reviewerId===userId ? 'Reviewer 1' : 'Reviewer 2';
    const unresolvedComments = _comments.filter(c=>(c.taskId===task._spId||c.taskId===task.id)&&!c.isResolved).length;
    return `<div class="mytask-row">
      <div class="mytask-main">
        <div class="mytask-name"
          onclick="openSteps('${task._spId}','${escHtml(task.name)}')"
          style="cursor:pointer" title="Click to open steps">
          ${escHtml(task.name)}
          <span class="mytask-role-chip">${myRole}</span>
          ${task.skipNextRollforward?'<span class="skip-badge">⊘ skip</span>':''}
        </div>
        <div class="mytask-meta">
          <span class="badge ${typeBadgeClass(task.type)}" style="font-size:10px">${escHtml(task.type)}</span>
          <span class="deadline-cell">
            <span class="deadline-dot ${dotClass(ds)}"></span>
            ${formatWorkdayDate(task.workdayNum, task.dueDate, task.quarter, task.year)}
            ${ds === 'overdue' ? '<span class="overdue-label"> OVERDUE</span>' : ''}
          </span>
          <span class="text-muted" style="font-size:11px">${task.quarter} ${task.year}</span>
        </div>
      </div>
      <div class="mytask-actions">
        <span class="status-badge ${statusBadgeClass(task.status)}" ${canEdit?`onclick="cycleStatus('${task._spId}')" style="cursor:pointer"`:''}>${escHtml(task.status)}</span>
        <button class="icon-btn" title="Steps" onclick="openSteps('${task._spId}','${escHtml(task.name)}')">📋</button>
        <button class="icon-btn ${unresolvedComments>0 ? 'comment-btn-unresolved' : ''}"
          title="Comments"
          onclick="openComments('${task._spId}','${escHtml(task.name)}')">
          💬${unresolvedComments > 0 ? `<span class="unresolved-badge">${unresolvedComments}</span>` : ''}
        </button>
      </div>
    </div>`;
  }

  function stepRow(step) {
    const parentTask = _tasks.find(t=>t._spId===step.taskId);
    const locked     = parentTask ? isQuarterLocked(parentTask.quarter, parentTask.year) : false;
    const ds         = step.dueDate ? deadlineStatus(step.dueDate, step.status) : 'ok';
    const myRole     = step.ownerId===userId ? 'Preparer' : step.reviewerId===userId ? 'Reviewer 1' : 'Reviewer 2';
    const blocker    = getBlockingPredecessor(step);
    return `<div class="mytask-row step">
      <div class="mytask-main">
        <div class="mytask-name">${escHtml(step.name)}
          <span class="mytask-role-chip">${myRole}</span>
          ${blocker?`<span class="mytask-blocked-chip">🔒 needs: ${escHtml(blocker.name)}</span>`:''}
        </div>
        <div class="mytask-meta">
          <span class="text-muted" style="font-size:11px">↳ ${escHtml(parentTask?.name||'')}</span>
          ${step.dueDate||step.workdayNum?`<span class="deadline-cell"><span class="deadline-dot ${dotClass(ds)}"></span>${formatWorkdayDate(step.workdayNum,step.dueDate,parentTask?.quarter,parentTask?.year)}</span>`:''}
        </div>
      </div>
      <div class="mytask-actions">
        ${!locked
          ? `<span class="status-badge ${statusBadgeClass(step.status)}"
               style="cursor:pointer;font-size:11px"
               onclick="cycleStepStatus('${step._spId}')">${escHtml(step.status)}</span>`
          : `<span class="status-badge ${statusBadgeClass(step.status)}"
               style="font-size:11px">${escHtml(step.status)}</span>`}
      </div>
    </div>`;
  }

  function section(label, emoji, tasks, steps, emptyMsg) {
    const total = tasks.length + steps.length;
    if (!total && emptyMsg==='hide') return '';
    return `<div class="mytask-section">
      <div class="mytask-section-header">
        <span class="mytask-section-emoji">${emoji}</span>
        <span class="mytask-section-label">${label}</span>
        <span class="mytask-section-count">${total}</span>
      </div>
      ${total ? `<div class="mytask-section-body">
        ${tasks.map(taskRow).join('')}
        ${steps.map(stepRow).join('')}
      </div>` : `<div class="mytask-empty">${emptyMsg}</div>`}
    </div>`;
  }

  container.innerHTML =
    section('Ready for Me',      '⚡', taskGroups.readyForMe, stepGroups.readyForMe, 'Nothing needs your attention right now.') +
    section('Waiting on Others', '⏳', taskGroups.waiting,    stepGroups.waiting,    'hide') +
    section('Complete',          '✅', taskGroups.complete,   stepGroups.complete,   'hide');
}

// ═══════════════════════════════════════════════════════════════
// ── KANBAN BOARD ─────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function initKanbanSelects() {
  const cur = new Date();
  const m   = cur.getMonth();
  const q   = m<3?'Q1':m<6?'Q2':m<9?'Q3':'Q4';
  // Only set defaults if no value is already selected — preserve user's last choice
  const qEl = document.getElementById('kanban-quarter');
  if (qEl && !qEl.value) qEl.value = q;
  const yEl = document.getElementById('kanban-year');
  if (yEl && !yEl.value) yEl.value = cur.getFullYear();
}

function renderKanban() {
  const mode    = document.getElementById('kanban-mode')?.value    || 'personal';
  const quarter = document.getElementById('kanban-quarter')?.value || 'Q1';
  const year    = parseInt(document.getElementById('kanban-year')?.value || new Date().getFullYear());
  const locked  = isQuarterLocked(quarter, year);
  const userId = currentUserId();
  const subEl   = document.getElementById('kanban-sub');
  if (subEl) subEl.textContent = mode==='personal'
    ? `${currentUser.name.split(' ')[0]}'s board · ${quarter} ${year}${locked?' 🔒':''}`
    : `Full team · ${quarter} ${year}${locked?' 🔒':''}`;

  let tasks = getActiveTasks().filter(t => t.quarter===quarter && t.year===year);
  if (mode === 'personal') tasks = tasks.filter(t => t.ownerId===userId);

  const board = document.getElementById('kanban-board');
  if (!board) return;

  const cols = STATUS_ORDER.filter(s => s !== 'Not Applicable').map(status => {
    const colTasks = tasks.filter(t => t.status === status);
    const cards    = colTasks.map(task => {
      const owner    = getUserById(task.ownerId);
      const ds       = deadlineStatus(task.dueDate, task.status);
      const canEdit  = !locked && (currentUser.isAdmin || task.ownerId===uid);
      const steps    = getStepsForTask(task._spId);
      const doneS    = steps.filter(s=>s.status==='Complete').length;
      const appBadge = task.applicability && task.applicability !== 'All Quarters'
        ? `<span class="app-badge ${task.applicability.startsWith('10-K')?'app-badge-10k':'app-badge-10q'} " style="margin-left:0">${task.applicability.startsWith('10-K')?'10-K':'10-Q'}</span>` : '';
      return `<div class="kanban-card" draggable="${canEdit}" ondragstart="kanbanDragStart(event,'${task._spId}')" ondragover="event.preventDefault()" ondrop="kanbanDrop(event,'${status}')">
        <div class="kanban-card-header">
          <span class="badge ${typeBadgeClass(task.type)}" style="font-size:10px">${escHtml(task.type)}</span>
          ${appBadge}
        </div>
        <div class="kanban-card-name">${escHtml(task.name)}</div>
        ${steps.length ? `
          <div class="kanban-step-bar">
            <div class="kanban-step-fill" style="width:${Math.round(doneS/steps.length*100)}%"></div>
          </div>
          <div class="kanban-step-label">${doneS}/${steps.length} steps</div>` : ''}
        <div class="kanban-card-footer">
          <div class="owner-chip">
            <div class="mini-avatar" style="width:20px;height:20px;font-size:9px">${owner?initials(owner.name):'?'}</div>
            ${mode==='team'&&owner?`<span style="font-size:11px">${escHtml(owner.name.split(' ')[0])}</span>`:''}
          </div>
          <div class="deadline-cell" style="font-size:11px">
            <span class="deadline-dot ${dotClass(ds)}"></span>${formatWorkdayDate(task.workdayNum, task.dueDate, task.quarter, task.year)}
          </div>
        </div>
        <div class="kanban-card-actions">
          <button class="icon-btn" style="width:24px;height:24px;font-size:11px" onclick="openSteps('${task._spId}','${escHtml(task.name)}')">📋</button>
          <button class="icon-btn" style="width:24px;height:24px;font-size:11px" onclick="openComments('${task._spId}','${escHtml(task.name)}')">💬</button>
          ${canEdit?`<button class="icon-btn" style="width:24px;height:24px;font-size:11px" onclick="openEditTask('${task._spId}')">✏️</button>`:''}
        </div>
      </div>`;
    }).join('');

    return `<div class="kanban-col" ondragover="event.preventDefault()" ondrop="kanbanDrop(event,'${status}')">
      <div class="kanban-col-header">
        <span class="kanban-col-title">${status}</span>
        <span class="kanban-col-count">${colTasks.length}</span>
      </div>
      <div class="kanban-col-body" id="kanban-col-${status.replace(/\s/g,'-')}">
        ${cards || `<div class="kanban-empty">No tasks</div>`}
      </div>
    </div>`;
  });

  board.innerHTML = cols.join('');
}

let _dragTaskSpId = null;
function kanbanDragStart(e, spId) {
  _dragTaskSpId = spId;
  e.dataTransfer.effectAllowed = 'move';
}
async function kanbanDrop(e, newStatus) {
  e.preventDefault();
  if (!_dragTaskSpId) return;
  const task = _tasks.find(t => t._spId === _dragTaskSpId);
  if (!task || task.status === newStatus) { _dragTaskSpId = null; return; }
  const prev = task.status;
  task.status = newStatus;
  renderKanban();
  try {
    await updateListItem(LISTS.tasks, _dragTaskSpId, { Status: newStatus });
    await writeSignOff(_dragTaskSpId, 'task', task.name, prev, newStatus);
  } catch(e) {
    task.status = prev; // revert optimistic update
    renderKanban();
    console.error('Kanban drop failed:', e);
    showToast('Could not move task: ' + e.message, 'error');
  }
  _dragTaskSpId = null;
}

// ═══════════════════════════════════════════════════════════════
// ── SUMMARY REPORT VIEW ──────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function initReportSelects() {
  switchReportTab('quarter');

  const cur = new Date();
  const m   = cur.getMonth();
  const q   = m<3?'Q1':m<6?'Q2':m<9?'Q3':'Q4';
  const qEl = document.getElementById('report-quarter'); if(qEl) qEl.value = q;
  const yEl = document.getElementById('report-year');    if(yEl) yEl.value = cur.getFullYear();
}

function renderReport() {
  const quarter = document.getElementById('report-quarter')?.value || 'Q1';
  const year    = parseInt(document.getElementById('report-year')?.value || new Date().getFullYear());
  const tasks   = getTasks().filter(t => t.quarter===quarter && t.year===year);
  const locked  = isQuarterLocked(quarter, year);
  const el      = document.getElementById('report-content');
  if (!el) return;

  const activeTasks2 = tasks.filter(t=>t.status!=='Not Applicable');
  const naCount2     = tasks.filter(t=>t.status==='Not Applicable').length;
  const total        = activeTasks2.length;
  const complete     = activeTasks2.filter(t=>t.status==='Complete').length;
  const review       = activeTasks2.filter(t=>t.status==='Ready for Review 1'||t.status==='Ready for Review 2').length;
  const inprog       = activeTasks2.filter(t=>t.status==='In Progress').length;
  const notstart     = activeTasks2.filter(t=>t.status==='Not Started').length;
  const overdue      = activeTasks2.filter(t=>deadlineStatus(t.dueDate,t.status)==='overdue').length;
  const pct          = total ? Math.round(complete/total*100) : 0;

  const byType   = {};
  tasks.forEach(t => { byType[t.type] = (byType[t.type]||0)+1; });

  const taskRows = tasks.map(task => {
    const owner     = getUserById(task.ownerId);
    const reviewer  = task.reviewerId  ? getUserById(task.reviewerId)  : null;
    const reviewer2 = task.reviewer2Id ? getUserById(task.reviewer2Id) : null;
    const steps     = getStepsForTask(task._spId);
    const doneS     = steps.filter(s=>s.status==='Complete').length;
    const signoffs  = getSignOffsFor(task._spId);
    const lastSO    = signoffs[0];
    const ds        = deadlineStatus(task.dueDate, task.status);
    // Find who signed off at each key stage
    const preparedBy  = signoffs.slice().reverse().find(s=>s.toStatus==='In Progress');
    const reviewed1By = signoffs.slice().reverse().find(s=>s.toStatus==='Ready for Review 1'||s.toStatus==='Ready for Review');
    const reviewed2By = signoffs.slice().reverse().find(s=>s.toStatus==='Ready for Review 2');
    const completedBy = signoffs.slice().reverse().find(s=>s.toStatus==='Complete');
    return `<tr>
      <td style="font-weight:600">${escHtml(task.name)}</td>
      <td><span class="badge ${typeBadgeClass(task.type)}" style="font-size:10px">${escHtml(task.type)}</span></td>
      <td><div style="font-size:12px">${owner?escHtml(owner.name):'—'}
        ${reviewer?`<div style="font-size:10px;color:var(--text-faint)">Rev: ${escHtml(reviewer.name)}</div>`:''}
        ${reviewer2?`<div style="font-size:10px;color:var(--text-faint)">VP: ${escHtml(reviewer2.name)}</div>`:''}
      </div></td>
      <td><div class="deadline-cell" style="font-size:12px"><span class="deadline-dot ${dotClass(ds)}"></span>${formatWorkdayDate(task.workdayNum, task.dueDate, task.quarter, task.year)}</div></td>
      <td><span class="status-badge ${statusBadgeClass(task.status)}" style="font-size:11px">${escHtml(task.status)}</span></td>
      <td style="font-size:11px;color:var(--text-muted)">${steps.length?`${doneS}/${steps.length}`:'—'}</td>
      <td style="font-size:11px;color:var(--text-muted)">
        ${preparedBy  ? `<div>▶ ${escHtml(preparedBy.userName)} <span style="color:var(--text-faint)">${escHtml(preparedBy.ts)}</span></div>`:''}
        ${reviewed1By ? `<div>🔍 ${escHtml(reviewed1By.userName)} <span style="color:var(--text-faint)">${escHtml(reviewed1By.ts)}</span></div>`:''}
        ${reviewed2By ? `<div>🔎 ${escHtml(reviewed2By.userName)} <span style="color:var(--text-faint)">${escHtml(reviewed2By.ts)}</span></div>`:''}
        ${completedBy ? `<div>✅ ${escHtml(completedBy.userName)} <span style="color:var(--text-faint)">${escHtml(completedBy.ts)}</span></div>`:''}
        ${!preparedBy&&!reviewed1By&&!completedBy?'—':''}
      </td>
    </tr>`;
  }).join('');

  el.innerHTML = `
    <div class="report-header">
      <div>
        <div class="report-title">${quarter} ${year} Financial Reporting</div>
        <div class="report-subtitle">Quarter Summary Report · Generated ${nowLabel()}${locked?' · 🔒 Locked':''}</div>
      </div>
      <div class="report-logo">FRT · Financial Reporting Tracker</div>
    </div>

    <div class="report-stat-row">
      <div class="report-stat"><div class="report-stat-val">${total}</div><div class="report-stat-lbl">Active Deliverables</div></div>
      <div class="report-stat complete"><div class="report-stat-val">${complete}</div><div class="report-stat-lbl">Complete</div></div>
      <div class="report-stat review"><div class="report-stat-val">${review}</div><div class="report-stat-lbl">Ready for Review</div></div>
      <div class="report-stat progress"><div class="report-stat-val">${inprog}</div><div class="report-stat-lbl">In Progress</div></div>
      <div class="report-stat overdue"><div class="report-stat-val">${overdue}</div><div class="report-stat-lbl">Overdue</div></div>
      <div class="report-stat"><div class="report-stat-val">${pct}%</div><div class="report-stat-lbl">Completion Rate</div></div>
    </div>

    <div class="report-progress-wrap">
      <div class="report-progress-fill" style="width:${pct}%"></div>
    </div>

    <div style="display:flex;gap:16px;margin:16px 0;flex-wrap:wrap">
      ${Object.entries(byType).map(([type,count])=>`
        <div style="display:flex;align-items:center;gap:6px;font-size:12px">
          <span class="badge ${typeBadgeClass(type)}" style="font-size:10px">${escHtml(type)}</span>
          <span style="color:var(--text-muted)">${count} task${count!==1?'s':''}</span>
        </div>`).join('')}
    </div>

    <table class="task-table report-table">
      <thead><tr>
        <th>Deliverable</th><th>Type</th><th>Owner / Reviewers</th>
        <th>Due Date</th><th>Status</th><th>Steps</th><th>Sign-Off Trail</th>
      </tr></thead>
      <tbody>${taskRows}</tbody>
    </table>

    <div class="report-footer">
      Financial Reporting Tracker · ${quarter} ${year} · Confidential · ${nowLabel()}
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════════
// ── EXECUTIVE STATUS PAGE ────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function renderExecView() {
  const q   = document.getElementById('exec-quarter')?.value || document.getElementById('quarter-filter')?.value || 'Q1';
  const yr  = parseInt(document.getElementById('exec-year')?.value || document.getElementById('year-filter')?.value || new Date().getFullYear());
  const el  = document.getElementById('exec-content');
  if (!el) return;

  const allTasks    = getTasks().filter(t => t.quarter===q && t.year===yr);
  const tasks       = allTasks.filter(t => t.status !== 'Not Applicable');
  const naCount     = allTasks.filter(t => t.status === 'Not Applicable').length;
  const total       = tasks.length;
  const complete    = tasks.filter(t => t.status==='Complete').length;
  const inprog      = tasks.filter(t => t.status==='In Progress').length;
  const rfr         = tasks.filter(t => t.status==='Ready for Review 1'||t.status==='Ready for Review 2').length;
  const notstarted  = tasks.filter(t => t.status==='Not Started').length;
  const overdue     = tasks.filter(t => deadlineStatus(t.dueDate,t.status)==='overdue').length;
  const pct         = total ? Math.round(complete/total*100) : 0;
  const locked      = isQuarterLocked(q, yr);

  // Traffic light helper
  function tl(task) {
    const ds = deadlineStatus(task.dueDate, task.status);
    if (task.status==='Complete')                          return {dot:'tl-green',  label:'Complete'};
    if (ds==='overdue')                                    return {dot:'tl-red',    label:'Overdue'};
    if (task.status==='Ready for Review 1'||task.status==='Ready for Review 2') return {dot:'tl-blue', label:'For Review'};
    if (ds==='soon')                                       return {dot:'tl-amber',  label:'Due Soon'};
    return {dot:'tl-gray', label:task.status};
  }

  const taskRows = tasks
    .sort((a,b) => {
      const order = {'Complete':4,'Not Applicable':5,'Ready for Review 2':1,'Ready for Review 1':1,'In Progress':2,'Not Started':3};
      const oa = order[a.status]||3, ob = order[b.status]||3;
      if (oa!==ob) return oa-ob;
      return (a.dueDate||'').localeCompare(b.dueDate||'');
    })
    .map(task => {
      const owner    = getUserById(task.ownerId);
      const reviewer = task.reviewerId ? getUserById(task.reviewerId) : null;
      const signoffs = getSignOffsFor(task._spId);
      const lastSO   = signoffs[0];
      const completedBy = signoffs.slice().reverse().find(s=>s.toStatus==='Complete');
      const ds = deadlineStatus(task.dueDate, task.status);
      const traf = tl(task);
      const unresolvedComments = _comments.filter(c=>(c.taskId===task._spId||c.taskId===task.id)&&!c.isResolved).length;
      return `<tr>
        <td style="width:24px;text-align:center"><span class="tl-dot ${traf.dot}"></span></td>
        <td style="font-weight:500;font-size:13px">${escHtml(task.name)}</td>
        <td><span class="badge ${typeBadgeClass(task.type)}" style="font-size:10px">${escHtml(task.type)}</span></td>
        <td style="font-size:12px">${owner?escHtml(owner.name):'—'}${reviewer?`<span style="font-size:10px;color:var(--text-muted)"> / ${escHtml(reviewer.name)}</span>`:''}</td>
        <td style="font-size:12px">
          ${formatWorkdayDate(task.workdayNum, task.dueDate, task.quarter, task.year)}
          ${ds === 'overdue' ? '<span style="color:var(--red);font-size:10px;font-weight:700;margin-left:4px">OVERDUE</span>' : ''}
        </td>
        <td><span class="status-badge ${statusBadgeClass(task.status)}" style="font-size:11px;pointer-events:none">${escHtml(task.status)}</span></td>
        <td style="font-size:11px;color:var(--text-muted)">
          ${completedBy ? `✅ ${escHtml(completedBy.userName)}` : lastSO ? `${escHtml(lastSO.toStatus)} — ${escHtml(lastSO.userName)}` : '—'}
          ${unresolvedComments>0?`<span style="color:var(--amber);font-weight:600;margin-left:4px">💬 ${unresolvedComments}</span>`:''}
        </td>
      </tr>`;
    }).join('');

  el.innerHTML = `
    <div class="exec-header">
      <div>
        <div class="exec-title">${q} ${yr} — Financial Reporting Status</div>
        <div class="exec-subtitle">As of ${nowLabel()}${locked?' · Quarter Locked':''}</div>
      </div>
      <div class="exec-logo">FRT · Financial Reporting Tracker</div>
    </div>

    <div class="exec-stats">
      <div class="exec-stat">
        <div class="exec-stat-val">${pct}%</div>
        <div class="exec-stat-lbl">Complete</div>
        <div class="exec-progress"><div class="exec-progress-fill ${pct>=100?'exec-progress-done':pct>=50?'exec-progress-mid':''}" style="width:${pct}%"></div></div>
      </div>
      <div class="exec-stat-grid">
        <div class="exec-mini-stat tl-green-bg"><span>${complete}</span><label>Complete</label></div>
        <div class="exec-mini-stat tl-blue-bg"><span>${rfr}</span><label>For Review</label></div>
        <div class="exec-mini-stat tl-amber-bg"><span>${inprog}</span><label>In Progress</label></div>
        <div class="exec-mini-stat tl-gray-bg"><span>${notstarted}</span><label>Not Started</label></div>
        <div class="exec-mini-stat tl-red-bg"><span>${overdue}</span><label>Overdue</label></div>
        <div class="exec-mini-stat" style="background:#F1F5F9"><span>${total}</span><label>Total</label></div>
      </div>
    </div>

    <table class="exec-table">
      <thead><tr>
        <th style="width:24px"></th>
        <th>Deliverable</th><th>Type</th><th>Owner / Reviewer</th>
        <th>Due Date</th><th>Status</th><th>Last Action</th>
      </tr></thead>
      <tbody>${taskRows}</tbody>
    </table>

    <div class="exec-footer">
      Financial Reporting Tracker · ${q} ${yr} · Confidential · Printed ${nowLabel()}
    </div>`;
}
// ═══════════════════════════════════════════════════════════════
// ── EXCEL EXPORT ─────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
// Exports tasks, steps, and sign-offs for the selected quarter as a multi-sheet .xls file.
function exportExcel() {
  const quarter = document.getElementById('report-quarter')?.value || 'Q1';
  const year    = parseInt(document.getElementById('report-year')?.value || new Date().getFullYear());
  const tasks   = getTasks().filter(t => t.quarter===quarter && t.year===year);

  // Build CSV content for three tabs (we use multi-sheet CSV trick via XLSX-style tab encoding)
  // Since we have no XLSX library, generate a proper HTML table that Excel can open natively

  // --- Sheet 1: Tasks ---
  const taskHeader = ['Task Name','Type','Quarter','Year','Owner','Due Date','Status','Applicability','Description'];
  const taskRows   = tasks.map(t => {
    const owner = getUserById(t.ownerId);
    return [t.name, t.type, t.quarter, t.year, owner?owner.name:'', t.dueDate, t.status, t.applicability||'All Quarters', t.description||''];
  });

  // --- Sheet 2: Steps ---
  const stepHeader = ['Task Name','Step Name','Step Order','Owner','Due Date','Status','Applicability','Note'];
  const stepRows   = [];
  tasks.forEach(t => {
    getStepsForTask(t._spId).forEach(s => {
      const owner = getUserById(s.ownerId);
      stepRows.push([t.name, s.name, s.order, owner?owner.name:'', s.dueDate, s.status, s.applicability||'All Quarters', s.note||'']);
    });
  });

  // --- Sheet 3: Sign-Off Log (status changes only, not due-date edits) ---
  const soHeader = ['Type','Item','Changed By','From Status','To Status','Date & Time'];
  const soRows   = [];
  tasks.forEach(t => {
    getSignOffsFor(t._spId).filter(s=>s.refType!=='due_date_change').forEach(s => soRows.push(['Task', t.name, s.userName, s.fromStatus, s.toStatus, s.ts]));
    getStepsForTask(t._spId).forEach(step => {
      getSignOffsFor(step._spId).filter(s=>s.refType!=='due_date_change').forEach(s => soRows.push(['Step', step.name, s.userName, s.fromStatus, s.toStatus, s.ts]));
    });
  });

  // Build multi-sheet HTML workbook (Excel opens this natively)
  const sheetHtml = (name, header, rows) => `
    <table>
      <thead><tr>${header.map(h=>`<th style="background:#0f2140;color:#fff;font-weight:bold">${h}</th>`).join('')}</tr></thead>
      <tbody>${rows.map((r,i)=>`<tr style="background:${i%2?'#f7f9fc':'#fff'}">${r.map(c=>`<td>${escHtml(String(c||''))}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>`;

  const wb = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<head><meta charset="UTF-8">
<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>
  <x:ExcelWorksheet><x:Name>Tasks</x:Name><x:WorksheetOptions><x:Selected/></x:WorksheetOptions></x:ExcelWorksheet>
  <x:ExcelWorksheet><x:Name>Steps</x:Name><x:WorksheetOptions/></x:ExcelWorksheet>
  <x:ExcelWorksheet><x:Name>Sign-Off Log</x:Name><x:WorksheetOptions/></x:ExcelWorksheet>
</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
<style>td,th{border:1px solid #dde3ed;padding:6px 10px;font-size:12px;font-family:Calibri,sans-serif}table{border-collapse:collapse}</style>
</head><body>
${sheetHtml('Tasks',     taskHeader, taskRows)}
${sheetHtml('Steps',     stepHeader, stepRows)}
${sheetHtml('Sign-Offs', soHeader,   soRows)}
</body></html>`;

  const blob = new Blob([wb], { type: 'application/vnd.ms-excel;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `FinancialReportingTracker_${quarter}_${year}.xls`;
  a.click();
  URL.revokeObjectURL(url);
}

// ── FEATURE 7 — TEAM SUMMARY REPORT ──────────────────────────

function switchReportTab(tab) {
  const qContent = document.getElementById('report-content');
  const tContent = document.getElementById('team-summary-content');
  const qTab = document.getElementById('report-tab-quarter');
  const tTab = document.getElementById('report-tab-team');
  if (tab === 'quarter') {
    if (qContent) qContent.style.display = '';
    if (tContent) tContent.style.display = 'none';
    if (qTab) qTab.classList.add('report-tab-active');
    if (tTab) tTab.classList.remove('report-tab-active');
  } else {
    if (qContent) qContent.style.display = 'none';
    if (tContent) tContent.style.display = '';
    if (qTab) qTab.classList.remove('report-tab-active');
    if (tTab) tTab.classList.add('report-tab-active');
    renderTeamSummary();
  }
}
function renderTeamSummary() {
  const quarter = document.getElementById('report-quarter')?.value || 'Q1';
  const year    = parseInt(document.getElementById('report-year')?.value || new Date().getFullYear());
  const tasks   = getTasks().filter(t => t.quarter===quarter && t.year===year && t.status!=='Not Applicable');
  const el      = document.getElementById('team-summary-content');
  if (!el) return;

  const users = getUsers().filter(u => {
    const uid = u._spId || u.id;
    return tasks.some(t => t.ownerId===uid || t.reviewerId===uid || t.reviewer2Id===uid);
  });

  if (!users.length) { el.innerHTML = '<p class="text-muted" style="padding:16px">No team members assigned to tasks this quarter.</p>'; return; }

  const rows = users.map(u => {
    const uid  = u._spId || u.id;
    const prep = tasks.filter(t => t.ownerId===uid);
    const rev1 = tasks.filter(t => t.reviewerId===uid && t.ownerId!==uid);
    const rev2 = tasks.filter(t => t.reviewer2Id===uid && t.ownerId!==uid && t.reviewerId!==uid);
    const all  = [...new Set([...prep, ...rev1, ...rev2])];
    const complete = all.filter(t => t.status==='Complete').length;
    const overdue  = all.filter(t => deadlineStatus(t.dueDate,t.status)==='overdue').length;
    const rfr      = all.filter(t => t.status==='Ready for Review 1'||t.status==='Ready for Review 2').length;
    const inprog   = all.filter(t => t.status==='In Progress').length;
    const pct      = all.length ? Math.round(complete/all.length*100) : 0;
    return `<tr>
      <td><div class="owner-chip"><div class="mini-avatar">${initials(u.name)}</div>${escHtml(u.name)}</div></td>
      <td style="text-align:center">${prep.length}</td>
      <td style="text-align:center">${rev1.length + rev2.length}</td>
      <td style="text-align:center">${all.length}</td>
      <td style="text-align:center;color:var(--green-600)">${complete}</td>
      <td style="text-align:center;color:var(--blue)">${inprog}</td>
      <td style="text-align:center;color:var(--purple)">${rfr}</td>
      <td style="text-align:center;color:${overdue>0?'var(--red-500)':'inherit'};font-weight:${overdue>0?'600':'400'}">${overdue||'—'}</td>
      <td><div style="display:flex;align-items:center;gap:8px">
        <div style="flex:1;background:#EEF1F6;border-radius:4px;height:6px;overflow:hidden">
          <div style="width:${pct}%;height:100%;background:var(--blue);border-radius:4px"></div>
        </div>
        <span style="font-size:11px;color:var(--gray-500);min-width:30px">${pct}%</span>
      </div></td>
    </tr>`;
  }).join('');

  el.innerHTML = `
    <table class="report-team-table">
      <thead><tr>
        <th>Team Member</th>
        <th>Prep</th>
        <th>Review</th>
        <th>Total</th>
        <th>Complete</th>
        <th>In Progress</th>
        <th>For Review</th>
        <th>Overdue</th>
        <th>Progress</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
}

// Exports a per-team-member workload summary for the selected quarter as a .xls file.
function exportTeamSummary() {
  const quarter = document.getElementById('report-quarter')?.value || 'Q1';
  const year    = parseInt(document.getElementById('report-year')?.value || new Date().getFullYear());
  const tasks   = getTasks().filter(t => t.quarter===quarter && t.year===year && t.status!=='Not Applicable');
  const users   = getUsers().filter(u => {
    const uid = u._spId || u.id;
    return tasks.some(t => t.ownerId===uid || t.reviewerId===uid || t.reviewer2Id===uid);
  });

  const header = ['Name','Role','Prep Tasks','Review Tasks','Total','Complete','In Progress','For Review','Overdue','% Complete'];
  const rows   = users.map(u => {
    const userId = u._spId || u.id;
    const prep = tasks.filter(t => t.ownerId===userId);
    const rev  = tasks.filter(t => (t.reviewerId===userId||t.reviewer2Id===userId) && t.ownerId!==userId);
    const all  = [...new Set([...prep, ...rev])];
    const complete = all.filter(t => t.status==='Complete').length;
    const overdue  = all.filter(t => deadlineStatus(t.dueDate,t.status)==='overdue').length;
    const rfr      = all.filter(t => t.status==='Ready for Review 1'||t.status==='Ready for Review 2').length;
    const inprog   = all.filter(t => t.status==='In Progress').length;
    const pct      = all.length ? Math.round(complete/all.length*100) : 0;
    return [u.name, u.role, prep.length, rev.length, all.length, complete, inprog, rfr, overdue, pct+'%'];
  });

  const sheetHtml = (header, rows) => `<table>
    <thead><tr>${header.map(h=>`<th style="background:#0f2140;color:#fff;font-weight:bold">${h}</th>`).join('')}</tr></thead>
    <tbody>${rows.map((r,i)=>`<tr style="background:${i%2?'#f7f9fc':'#fff'}">${r.map(c=>`<td>${escHtml(String(c||''))}</td>`).join('')}</tr>`).join('')}</tbody>
  </table>`;

  const wb = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<head><meta charset="UTF-8"><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>
  <x:ExcelWorksheet><x:Name>Team Summary</x:Name><x:WorksheetOptions><x:Selected/></x:WorksheetOptions></x:ExcelWorksheet>
</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
<style>td,th{border:1px solid #dde3ed;padding:6px 10px;font-size:12px;font-family:Calibri,sans-serif}table{border-collapse:collapse}</style>
</head><body>${sheetHtml(header, rows)}</body></html>`;

  const blob = new Blob([wb], { type: 'application/vnd.ms-excel;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = `TeamSummary_${quarter}_${year}.xls`;
  a.click(); URL.revokeObjectURL(url);
}

// ═══════════════════════════════════════════════════════════════
// ── BULK EDIT ────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function toggleBulkSelectAll(checkbox, tbodyId) {
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  tbody.querySelectorAll('.bulk-check').forEach(cb => cb.checked = checkbox.checked);
  updateBulkBar();
}

function updateBulkBar() {
  const checked = document.querySelectorAll('.bulk-check:checked').length;
  const bar     = document.getElementById('bulk-bar');
  if (!bar) return;
  bar.classList.toggle('hidden', checked === 0);
  const lbl = document.getElementById('bulk-count-label');
  if (lbl) lbl.textContent = `${checked} task${checked!==1?'s':''} selected`;
}

function getSelectedTaskSpIds() {
  return [...document.querySelectorAll('.bulk-check:checked')].map(cb => cb.dataset.spid);
}

function openBulkEdit() {
  const spIds = getSelectedTaskSpIds();
  if (!spIds.length) return;
  const ownerOpts = getUsers().map(u =>
    `<option value="${u._spId||u.id}">${escHtml(u.name)}</option>`).join('');
  document.getElementById('modal-title').textContent = `Bulk Edit — ${spIds.length} tasks`;
  document.getElementById('modal-body').innerHTML = `
    <p style="font-size:13px;color:var(--text-muted);margin-bottom:16px">
      Leave a field blank to keep existing values unchanged.
    </p>
    <div class="form-group"><label>Reassign Preparer (Owner)</label>
      <select id="bulk-owner"><option value="">— Keep existing —</option>${ownerOpts}</select></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="form-group"><label>Reassign Reviewer 1</label>
        <select id="bulk-reviewer"><option value="">— Keep existing —</option><option value="__clear__">Clear reviewer</option>${ownerOpts}</select></div>
      <div class="form-group"><label>Reassign Reviewer 2</label>
        <select id="bulk-reviewer2"><option value="">— Keep existing —</option><option value="__clear__">Clear reviewer</option>${ownerOpts}</select></div>
    </div>
    <div class="form-group"><label>Shift Due Dates by (days)</label>
      <input type="number" id="bulk-shift" placeholder="e.g. 7 or -3" /></div>
    <div class="form-group"><label>Set Status</label>
      <select id="bulk-status">
        <option value="">— Keep existing —</option>
        ${STATUS_ORDER.map(s=>`<option>${s}</option>`).join('')}
      </select></div>
    <div class="modal-footer">
      <button class="btn-secondary" onclick="closeAllModals()">Cancel</button>
      <button class="btn-primary" onclick="applyBulkEdit()">Apply to ${spIds.length} tasks</button>
    </div>`;
  openModal();
}

async function applyBulkEdit() {
  const spIds     = getSelectedTaskSpIds();
  const owner     = document.getElementById('bulk-owner').value;
  const reviewer  = document.getElementById('bulk-reviewer')?.value;
  const reviewer2 = document.getElementById('bulk-reviewer2')?.value;
  const shift     = parseInt(document.getElementById('bulk-shift').value);
  const status    = document.getElementById('bulk-status').value;
  if (!owner && !reviewer && !reviewer2 && isNaN(shift) && !status) {
    showToast('Please set at least one field to change.', 'warning'); return;
  }
  closeAllModals();
  showLoadingOverlay(true, `Updating ${spIds.length} tasks…`);
  try {
    for (const spId of spIds) {
      const task   = _tasks.find(t => t._spId === spId); if(!task) continue;
      const fields = {};
      if (owner)    { fields.OwnerId    = owner;    task.ownerId    = owner; }
      if (reviewer) {
        const rv = reviewer==='__clear__' ? '' : reviewer;
        fields.ReviewerId  = rv; task.reviewerId  = rv;
      }
      if (reviewer2) {
        const rv2 = reviewer2==='__clear__' ? '' : reviewer2;
        fields.Reviewer2Id = rv2; task.reviewer2Id = rv2;
      }
      if (!isNaN(shift) && shift !== 0 && task.dueDate) {
        const newDate  = addDays(task.dueDate, shift);
        fields.DueDate = newDate; task.dueDate = newDate;
      }
      if (status) {
        const prev   = task.status;
        fields.Status = status; task.status = status;
        await writeSignOff(spId, 'task', task.name, prev, status);
      }
      if (Object.keys(fields).length) await updateListItem(LISTS.tasks, spId, fields);
    }
    await refreshData();
  } catch(e) { showToast('Bulk edit failed: ' + e.message, 'error'); }
  finally { showLoadingOverlay(false); }
}
