/* ============================================================
   FINANCIAL REPORTING TRACKER — app.js
   Hosted on:  GitHub Pages
   Database:   SharePoint Lists via Power Automate Flows
   Auth:       Microsoft SSO (User.Read only — no admin consent) + PIN

   Data operations route through Power Automate HTTP flows
   running inside the Moodys M365 tenant. No Sites.ReadWrite.All
   permission needed. User.Read handles identity only.
   ============================================================ */

// ── CONFIG ────────────────────────────────────────────────────
const CONFIG = {
  // Azure AD — User.Read only, no admin consent needed
  clientId:  "bb00291f-d451-4e74-b8cf-10c334efb0ed",
  tenantId:  "1061a8b8-b1ee-4249-bb84-9a2cd2792fae",

  // Power Automate flow URLs — paste each after creating in flow.microsoft.com
  flows: {
    getItems:   "https://default1061a8b8b1ee4249bb849a2cd2792f.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7e9e722c97e94f4dbe06ddffc20af421/triggers/manual/paths/invoke?api-version=1",
    createItem: "https://default1061a8b8b1ee4249bb849a2cd2792f.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/c4ccc22635ba4da191a19042a278ab74/triggers/manual/paths/invoke?api-version=1",
    updateItem: "https://default1061a8b8b1ee4249bb849a2cd2792f.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/4e29b81a714a4bd684bd660c04bf8fc8/triggers/manual/paths/invoke?api-version=1",
    deleteItem: "https://default1061a8b8b1ee4249bb849a2cd2792f.ae.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/1559f122bf474ecab813533ab3d27545/triggers/manual/paths/invoke?api-version=1",
  }
};
// ─────────────────────────────────────────────────────────────

// ── MSAL — User.Read only ─────────────────────────────────────
const msalConfig = {
  auth: {
    clientId:  CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true,
  },
  system: { allowNativeBroker: false }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const _msalReady   = Promise.resolve();

const GRAPH_SCOPES = ["User.Read"];

// ── USER IDENTITY ─────────────────────────────────────────────
async function getUserToken() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) throw new Error("Not authenticated");
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: GRAPH_SCOPES, account: accounts[0]
    });
    return result.accessToken;
  } catch {
    const result = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
    return result.accessToken;
  }
}

// ── POWER AUTOMATE FLOW HELPERS ───────────────────────────────
async function callFlow(url, body) {
  if (!url || url.startsWith("REPLACE_")) {
    throw new Error("Power Automate flow URL not configured — paste your flow URLs into the CONFIG block at the top of app.js.");
  }
  const res = await fetch(url, {
    method:  "POST",
    headers: { "Content-Type": "application/json" },
    body:    JSON.stringify(body),
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Flow error ${res.status}: ${err.slice(0, 200)}`);
  }
  if (res.status === 202) return null;
  const text = await res.text();
  if (!text) return null;
  return JSON.parse(text);
}

// ── LIST HELPERS (via Power Automate) ─────────────────────────
async function getListItems(listName) {
  const data = await callFlow(CONFIG.flows.getItems, { listName });
  return data?.items || [];
}

async function createListItem(listName, fields) {
  const data = await callFlow(CONFIG.flows.createItem, { listName, fields });
  return data || {};
}

async function updateListItem(listName, itemId, fields) {
  await callFlow(CONFIG.flows.updateItem, { listName, itemId: String(itemId), fields });
}

async function deleteListItem(listName, itemId) {
  await callFlow(CONFIG.flows.deleteItem, { listName, itemId: String(itemId) });
}
// ── LIST NAMES ───────────────────────────────────────────────
const LISTS = {
  tasks:         "FT_Tasks",
  users:         "FT_Users",
  comments:      "FT_Comments",
  templates:     "FT_Templates",
  steps:         "FT_Steps",
  stepTemplates: "FT_StepTemplates",
  signOffs:      "FT_SignOffs",   // audit trail for every status change
  locks:         "FT_QuarterLocks", // locked quarters (read-only)
  attachments:   "FT_Attachments",  // SharePoint links attached to steps
};

// ── IN-MEMORY CACHE ───────────────────────────────────────────
let _users         = [];
let _tasks         = [];
let _templates     = [];
let _comments      = [];
let _steps         = [];
let _stepTemplates = [];
let _signOffs      = [];   // all sign-off log entries
let _locks         = [];   // locked quarter records
let _attachments   = [];   // file attachment metadata
let _pollTimer         = null;
let _commentPollTimer  = null;

function getUserById(id)   { return _users.find(u => u.id === id || u.ID === id) || null; }
function getTaskById(id)   { return _tasks.find(t => t.id === id || t.ID === id) || null; }
function getUsers()        { return _users; }
function getTasks()        { return _tasks; }
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
    _spId:       f.id || f.ID || '',   // SharePoint item ID for updates
    id:          f.TaskId  || f.id || f.ID || '',
    name:        f.Title   || f.TaskName || '',
    type:        f.TaskType || '',
    quarter:     f.Quarter  || '',
    year:        parseInt(f.Year) || new Date().getFullYear(),
    dueDate:     (f.DueDate || '').slice(0, 10),
    status:      f.Status   || 'Not Started',
    ownerId:     f.OwnerId  || '',
    description:   f.Description   || '',
    applicability: f.Applicability || 'All Quarters',
    workdayNum:    f.WorkdayNum ? parseInt(f.WorkdayNum) : null,
  };
}
function normaliseUser(f) {
  return {
    _spId:   f.id || f.ID || '',
    id:      f.UserId   || f.id || f.ID || '',
    name:    f.Title    || f.FullName || '',
    role:    f.JobRole  || '',
    pin:     f.PIN      || '0000',
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
    _spId:    f.id || f.ID || '',
    id:       f.CommentId || f.id || f.ID,
    taskId:   f.TaskId    || '',
    authorId: f.AuthorId  || '',
    text:     f.CommentText || '',
    time:     f.CommentTime || '',
    ts:       f.Timestamp  || 0,
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
    _spId:      f.id || f.ID || '',
    id:         f.AttachmentId || f.id || f.ID,
    stepId:     f.StepId       || '',
    taskId:     f.TaskId       || '',
    label:      f.Title        || f.Label || '',
    url:        f.FileUrl      || '',
    linkedBy:   f.LinkedBy     || '',
    linkedAt:   f.LinkedAt     || '',
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
    dueDate: (f.DueDate  || '').slice(0,10),
    note:          f.Note          || '',
    applicability: f.Applicability || 'All Quarters',
    workdayNum:    f.WorkdayNum ? parseInt(f.WorkdayNum) : null,
    requiresPrev:  f.RequiresPrev === true || f.RequiresPrev === 'Yes',
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
async function loadAllData() {
  showLoadingOverlay(true);
  try {
    const [rawTasks, rawUsers, rawTemplates, rawComments, rawSteps, rawStepTpls, rawSignOffs] = await Promise.all([
      getListItems(LISTS.tasks),
      getListItems(LISTS.users),
      getListItems(LISTS.templates),
      getListItems(LISTS.comments),
      getListItems(LISTS.steps),
      getListItems(LISTS.stepTemplates),
      getListItems(LISTS.signOffs),
      getListItems(LISTS.locks),
      getListItems(LISTS.attachments),
    ]);
    _tasks         = rawTasks.map(normaliseTask);
    _users         = rawUsers.map(normaliseUser);
    _templates     = rawTemplates.map(normaliseTemplate);
    _comments      = rawComments.map(normaliseComment);
    _steps         = rawSteps.map(normaliseStep);
    _stepTemplates = rawStepTpls.map(normaliseStepTemplate);
    _signOffs      = rawSignOffs.map(normaliseSignOff);
    _locks         = (await getListItems(LISTS.locks)).map(normaliseLock);
    _attachments   = (await getListItems(LISTS.attachments)).map(normaliseAttachment);
  } catch(e) {
    console.error("Data load error:", e);
    showError("Could not load data from SharePoint. Check your config and list names. " + e.message);
  } finally {
    showLoadingOverlay(false);
  }
}

async function refreshData() {
  try {
    const [rawTasks, rawUsers, rawTemplates, rawSteps, rawStepTpls, rawSignOffs] = await Promise.all([
      getListItems(LISTS.tasks),
      getListItems(LISTS.users),
      getListItems(LISTS.templates),
      getListItems(LISTS.steps),
      getListItems(LISTS.stepTemplates),
      getListItems(LISTS.signOffs),
    ]);
    _tasks         = rawTasks.map(normaliseTask);
    _users         = rawUsers.map(normaliseUser);
    _templates     = rawTemplates.map(normaliseTemplate);
    _steps         = rawSteps.map(normaliseStep);
    _stepTemplates = rawStepTpls.map(normaliseStepTemplate);
    _signOffs      = rawSignOffs.map(normaliseSignOff);
    _locks         = (await getListItems(LISTS.locks)).map(normaliseLock);
    _attachments   = (await getListItems(LISTS.attachments)).map(normaliseAttachment);
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

function startPolling() {
  if (_pollTimer) clearInterval(_pollTimer);
  _pollTimer = setInterval(refreshData, 30000); // every 30s
}

function startCommentPolling() {
  if (_commentPollTimer) clearInterval(_commentPollTimer);
  _commentPollTimer = setInterval(refreshComments, 10000); // every 10s
}

function stopCommentPolling() {
  if (_commentPollTimer) { clearInterval(_commentPollTimer); _commentPollTimer = null; }
}

// ── STATE ─────────────────────────────────────────────────────
let currentUser      = null;
let activeFilter     = 'all';
let activeTypeFilter = 'all';
let calYear          = new Date().getFullYear();
let calMonth         = new Date().getMonth();
let editingTaskId    = null;
let commentingTaskId = null;

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
  const diff  = Math.floor((due - today) / 86400000);
  if (diff < 0)  return 'overdue';
  if (diff <= 7) return 'soon';
  return 'ok';
}
function isThisWeek(dueDate) {
  const today = new Date(); today.setHours(0,0,0,0);
  const diff  = Math.floor((new Date(dueDate+'T00:00:00') - today) / 86400000);
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
  localStorage.setItem('ft_custom_holidays', JSON.stringify(list));
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
  h.add(iso(nthWeekday(year, 5, 1, 1)));   // Juneteenth observed (nearest Mon if needed) — fixed Jun 19 observed
  const juneteenth = observed(new Date(year, 5, 19)); h.add(iso(juneteenth));
  h.add(iso(nthWeekday(year, 8, 1, 1)));   // Labor Day: 1st Mon Sep
  h.add(iso(nthWeekday(year, 9, 4, 4)));   // Columbus Day: 2nd Mon Oct  (4th Thu = Thanksgiving below)
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

// ── Master holiday checker ────────────────────────────────────
function isHoliday(dateStr) {
  const year  = parseInt(dateStr.slice(0,4));
  const fed   = usFederalHolidays(year);
  const nyse  = nyseExtraHolidays(year);
  const custom = new Set(loadCustomHolidays());
  return fed.has(dateStr) || nyse.has(dateStr) || custom.has(dateStr);
}

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

  // Pick closer; ties go to the day before (earlier = more conservative for deadlines)
  return bDays <= aDays ? before : after;
}

// ── Admin: Custom Holiday Manager ────────────────────────────

// ── CLOSE CALENDARS PANEL (in Admin) ────────────────────────
function renderCloseCalendarsPanel() {
  const el = document.getElementById('admin-calendars-list');
  if (!el) return;
  const cur  = new Date().getFullYear();
  const rows = [];
  // Show last year, current year, next year × 4 quarters
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
          : '<span style="color:var(--gray-400)">No calendar set</span>'
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
      ? `<div style="margin-bottom:6px"><span style="font-weight:600;color:var(--gray-600)">${label}:</span> `
        + dates.map(d => `<span style="display:inline-block;background:var(--gray-100);border-radius:4px;padding:1px 6px;margin:1px 2px;font-size:11px">${formatDate(d)}</span>`).join(' ')
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
    alert('Please enter a valid date in YYYY-MM-DD format.'); return;
  }
  const list = loadCustomHolidays();
  if (!list.includes(val)) { list.push(val); saveCustomHolidays(list); }
  if (input) input.value = '';
  renderCustomHolidays();
}

function removeCustomHoliday(dateStr) {
  saveCustomHolidays(loadCustomHolidays().filter(d => d !== dateStr));
  renderCustomHolidays();
}

// ── Working day summary for a given year (for display) ────────
function holidaySummaryForYear(year) {
  const fed  = [...usFederalHolidays(year)].sort();
  const nyse = [...nyseExtraHolidays(year)].sort();
  return { federal: fed, nyse, custom: loadCustomHolidays().filter(d=>d.startsWith(year)) };
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
  try {
    return JSON.parse(localStorage.getItem('ft_cal_' + calendarKey(quarter, year)) || 'null');
  } catch { return null; }
}

function saveCloseCalendar(quarter, year, calObj) {
  localStorage.setItem('ft_cal_' + calendarKey(quarter, year), JSON.stringify(calObj));
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
function formatWorkdayDate(workdayNum, dateStr) {
  if (!dateStr && !workdayNum) return '—';
  const wdLabel = workdayNum ? `<span class="wd-badge">WD${workdayNum}</span> ` : '';
  const dateLabel = dateStr ? formatDate(dateStr) : '';
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
  const overrides = _editingCal.overrides;
  const cal       = buildCloseCalendar(wd1DateStr, maxDays, overrides);

  const resetBtn = isPrompt
    ? `<button class="btn-secondary" style="font-size:11px;padding:4px 10px" onclick="resetCalOverrides('${containerId}','${wd1DateStr}',${maxDays},${isPrompt})">↺ Reset overrides</button>`
    : `<button class="btn-secondary" style="font-size:11px;padding:4px 10px" onclick="resetCalOverrides('${containerId}','${wd1DateStr}',${maxDays},${isPrompt})">↺ Reset overrides</button>`;

  el.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
      <p style="font-size:12px;font-weight:600;color:var(--gray-600);margin:0">
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
    <p style="font-size:11px;color:var(--gray-400);margin-top:8px">
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
      <p style="font-size:13px;color:var(--gray-600);margin-bottom:16px">
        Set <strong>Workday 1</strong> for ${toQ} ${toY}, then adjust any specific
        dates below — for example if your team works on Good Friday, click that
        date and change it to the actual day your team is working.
      </p>
      <div class="form-group">
        <label>Workday 1 date for ${toQ} ${toY}</label>
        <input type="date" id="wd1-input" value="${existing?.wd1Date || suggested}" />
        <p style="font-size:11px;color:var(--gray-400);margin-top:4px">
          First working day of your close — quarter ends ${formatDate(qEndDate)}.
        </p>
      </div>
      <div id="cal-preview" style="margin-top:16px"></div>
      <div class="modal-footer">
        <button class="btn-secondary" onclick="closeAllModals();window._calResolve(null)">Cancel</button>
        <button class="btn-primary" onclick="confirmCloseCalendar('${toQ}',${toY})">Confirm & Continue →</button>
      </div>`;

    openModal();
    window._calResolve = resolve;

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
  if (!val) { alert('Please set a Workday 1 date.'); return; }
  const cal = buildCloseCalendar(val, 40, _editingCal.overrides);
  saveCloseCalendar(toQ, toY, cal);
  closeAllModals();
  if (window._calResolve) { window._calResolve(cal); window._calResolve = null; }
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
    <p style="font-size:13px;color:var(--gray-600);margin-bottom:12px">
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
  if (!val) { alert('Please set a Workday 1 date.'); return; }
  const cal = buildCloseCalendar(val, 40, _editingCal.overrides);
  saveCloseCalendar(quarter, year, cal);
  closeAllModals();
  renderAdmin();
  const overrideCount = Object.keys(_editingCal.overrides).length;
  alert(`Close calendar saved for ${quarter} ${year}.\nWD1 = ${formatDate(cal.wd1Date)}${overrideCount ? `\n${overrideCount} date override(s) applied.` : ''}`);
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
          'Ready for Review':'status-review','Complete':'status-complete',
          'Not Applicable':'status-na'}[s]||'status-not-started';
}
function dotClass(ds) {
  return {overdue:'dot-overdue',soon:'dot-soon',ok:'dot-ok',done:'dot-done'}[ds]||'dot-ok';
}
function calChipBg(type)  { return {'Close':'#e0f2fe','Financial Report':'#dbeafe','Master SS':'#d1fae5','Ops Book':'#fce7f3','Other':'#f1f5f9','Press Release':'#ede9fe','Post-Filing':'#fef3c7','Pre-Filing':'#ecfdf5'}[type]||'#f1f5f9'; }
function calChipFg(type)  { return {'Close':'#0369a1','Financial Report':'#1d4ed8','Master SS':'#065f46','Ops Book':'#9d174d','Other':'#475569','Press Release':'#5b21b6','Post-Filing':'#92400e','Pre-Filing':'#065f46'}[type]||'#334155'; }
function escHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ── UI HELPERS ────────────────────────────────────────────────
function showLoadingOverlay(show) {
  let el = document.getElementById('loading-overlay');
  if (!el) {
    el = document.createElement('div');
    el.id = 'loading-overlay';
    el.style.cssText = `position:fixed;inset:0;background:rgba(15,33,64,.55);
      display:flex;align-items:center;justify-content:center;z-index:999;
      color:#fff;font-family:'DM Sans',sans-serif;font-size:15px;gap:12px;`;
    el.innerHTML = `<span style="font-size:22px">⏳</span> Connecting to SharePoint…`;
    document.body.appendChild(el);
  }
  el.style.display = show ? 'flex' : 'none';
}
function showError(msg) {
  const el = document.getElementById('sp-error');
  if (el) { el.textContent = msg; el.classList.remove('hidden'); }
}

// ── LOGIN (Microsoft SSO) ────────────────────────────────────
async function loginWithMicrosoft() {
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

  await loadAllData();

  const email = msAccount.username?.toLowerCase() || '';
  let user = _users.find(u => (u.email||'').toLowerCase() === email);

  populateUserSelect();

  if (user) {
    document.getElementById('login-user-select').value = user._spId || user.id;
  }

  document.getElementById('ms-login-row').classList.add('hidden');
  document.getElementById('pin-login-row').classList.remove('hidden');
}

function populateUserSelect() {
  const sel = document.getElementById('login-user-select');
  sel.innerHTML = '<option value="">— Choose team member —</option>';
  _users.sort((a,b) => a.name.localeCompare(b.name)).forEach(u => {
    const opt = document.createElement('option');
    opt.value = u._spId || u.id;
    opt.textContent = u.name;
    sel.appendChild(opt);
  });
}

function login() {
  const selVal = document.getElementById('login-user-select').value;
  const pin    = document.getElementById('login-pin').value;
  const errEl  = document.getElementById('login-error');
  const user   = _users.find(u => (u._spId||u.id) === selVal && u.pin === pin);
  if (!user) { errEl.classList.remove('hidden'); return; }
  errEl.classList.add('hidden');
  currentUser = user;
  sessionStorage.setItem('ft_session', user._spId || user.id);
  launchApp();
}

function logout() {
  currentUser = null;
  sessionStorage.removeItem('ft_session');
  sessionStorage.removeItem('ft_pending_user');
  stopCommentPolling();
  if (_pollTimer) clearInterval(_pollTimer);
  msalInstance.logoutPopup().catch(()=>{});
  document.getElementById('app-screen').classList.remove('active');
  document.getElementById('login-screen').classList.add('active');
  document.getElementById('ms-login-row').classList.remove('hidden');
  document.getElementById('pin-login-row').classList.add('hidden');
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
  startPolling();
  renderDashboard();
}

function setCurrentQuarter() {
  const now = new Date(); const m = now.getMonth();
  const q   = m<3?'Q1':m<6?'Q2':m<9?'Q3':'Q4';
  const qf  = document.getElementById('quarter-filter'); if(qf) qf.value = q;
  const yf  = document.getElementById('year-filter');    if(yf) yf.value = now.getFullYear();
  const lbl = document.getElementById('dashboard-quarter-label');
  if(lbl) lbl.textContent = `${q} ${now.getFullYear()} · Financial Reporting`;
}
function populateYearSelects() {
  const cur = new Date().getFullYear();
  ['year-filter','template-year-select','rf-from-year','rf-to-year','kanban-year','report-year'].forEach(id => {
    const el = document.getElementById(id); if(!el) return;
    el.innerHTML = [cur-1,cur,cur+1].map(y=>`<option value="${y}">${y}</option>`).join('');
    el.value = cur;
  });
}

// ── VIEWS ─────────────────────────────────────────────────────
function switchView(view, el) {
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
  if (el) el.classList.add('active');
  document.getElementById('view-'+view).classList.add('active');
  if (view==='dashboard') renderDashboard();
  if (view==='tasks')     { populateOwnerFilter(); renderAllTasks(); }
  if (view==='calendar')  renderCalendar();
  if (view==='team')      renderTeam();
  if (view==='admin')     renderAdmin();
  if (view==='mytasks')   renderMyTasks();
  if (view==='kanban')    { initKanbanSelects(); renderKanban(); }
  if (view==='report')    { initReportSelects(); renderReport(); }
}
function setQuickFilter(filter, btn) {
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
  let tasks = getTasks().filter(t => t.quarter===q && t.year===yr);
  const hr  = new Date().getHours();
  const greet = document.getElementById('dashboard-greeting');
  if(greet) greet.textContent = `${hr<12?'Good morning':hr<17?'Good afternoon':'Good evening'}, ${currentUser.name.split(' ')[0]}`;
  const lbl = document.getElementById('dashboard-quarter-label');
  if(lbl) lbl.textContent = `${q} ${yr} · Financial Reporting`;
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
  if (activeFilter==='mine')      tasks=tasks.filter(t=>t.ownerId===currentUser.id||t.ownerId===currentUser._spId);
  if (activeTypeFilter!=='all')   tasks=tasks.filter(t=>t.type===activeTypeFilter);
  renderTaskTable(tasks, 'dashboard-task-body', true);
}

// ── ALL TASKS ─────────────────────────────────────────────────
function populateOwnerFilter() {
  const sel = document.getElementById('all-owner-filter'); if(!sel) return;
  sel.innerHTML = `<option value="all">All Owners</option>`+
    getUsers().map(u=>`<option value="${u.id}">${escHtml(u.name)}</option>`).join('');
}
function renderAllTasks() {
  const search=(document.getElementById('task-search')?.value||'').toLowerCase();
  const status=document.getElementById('all-status-filter')?.value||'all';
  const type  =document.getElementById('all-type-filter')?.value||'all';
  const owner =document.getElementById('all-owner-filter')?.value||'all';
  let tasks=getTasks();
  if(search) tasks=tasks.filter(t=>t.name.toLowerCase().includes(search)||t.description?.toLowerCase().includes(search));
  if(status!=='all') tasks=tasks.filter(t=>t.status===status);
  if(type!=='all')   tasks=tasks.filter(t=>t.type===type);
  if(owner!=='all')  tasks=tasks.filter(t=>t.ownerId===owner);
  renderTaskTable(tasks,'all-task-body',false);
}

// ── TASK TABLE ────────────────────────────────────────────────
function renderTaskTable(tasks, tbodyId, hiddenQuarter) {
  const tbody = document.getElementById(tbodyId); if(!tbody) return;
  if (!tasks.length) {
    tbody.innerHTML=`<tr><td colspan="8"><div class="empty-state">
      <div class="empty-icon">📋</div>
      <p>No tasks found. <a href="#" onclick="openAddTask();return false;">Add a new task</a>.</p>
    </div></td></tr>`; return;
  }
  tbody.innerHTML = tasks.map(task => {
    const owner   = getUserById(task.ownerId);
    const ds      = deadlineStatus(task.dueDate, task.status);
    const locked   = isQuarterLocked(task.quarter, task.year);
    const canEdit  = !locked && (currentUser.isAdmin || task.ownerId===currentUser.id || task.ownerId===currentUser._spId);
    const commentCount = _comments.filter(c=>c.taskId===task.id||c.taskId===task._spId).length;
    const taskSteps    = getStepsForTask(task._spId);
    const doneSteps    = taskSteps.filter(s=>s.status==='Complete').length;
    const stepPct      = taskSteps.length ? Math.round(doneSteps/taskSteps.length*100) : null;
    const qCol = hiddenQuarter ? '' :
      `<td><span class="badge ${typeBadgeClass(task.type)}">${task.quarter} ${task.year}</span></td>`;
    const stepBar = taskSteps.length ? `
      <div class="step-mini-bar" onclick="openSteps('${task._spId}','${escHtml(task.name)}')" title="Steps: ${doneSteps}/${taskSteps.length} complete">
        <div class="step-mini-fill" style="width:${stepPct}%"></div>
        <span class="step-mini-label">${doneSteps}/${taskSteps.length} steps</span>
      </div>` : `<span class="step-mini-add" onclick="openSteps('${task._spId}','${escHtml(task.name)}')">+ add steps</span>`;
    const appBadge = task.applicability && task.applicability !== 'All Quarters'
      ? `<span class="app-badge ${task.applicability.startsWith('10-K')?'app-badge-10k':'app-badge-10q'}">${task.applicability.startsWith('10-K')?'10-K':'10-Q'}</span>`
      : '';
    const taskTrail = renderSignOffTrail(task._spId, false);
    return `<tr style="${isNA?'opacity:0.5;':''}">
      <td style="width:32px;padding:8px 6px"><input type="checkbox" class="bulk-check" data-spid="${task._spId}" onchange="updateBulkBar()" style="cursor:pointer;width:15px;height:15px" /></td>
      <td>
        <div class="task-name">${escHtml(task.name)}${appBadge}</div>
        ${task.description?`<div class="task-desc">${escHtml(task.description)}</div>`:''}
        ${stepBar}
        ${taskTrail}
      </td>
      <td><span class="badge ${typeBadgeClass(task.type)}">${escHtml(task.type)}</span></td>
      ${qCol}
      <td><div class="owner-chip">
        <div class="mini-avatar">${owner?initials(owner.name):'?'}</div>
        ${owner?escHtml(owner.name):'—'}
      </div></td>
      <td><div class="deadline-cell">
        <span class="deadline-dot ${dotClass(ds)}"></span>
        ${formatWorkdayDate(task.workdayNum, task.dueDate)}
        ${ds==='overdue'?'<span class="text-danger fw-600" style="font-size:11px"> OVERDUE</span>':''}
      </div></td>
      <td>
        ${locked ? '<span style="font-size:11px;color:var(--gray-400)">🔒</span> ' : ''}
        <span class="status-badge ${statusBadgeClass(task.status)}"
        onclick="${canEdit?`cycleStatus('${task._spId}')`:''}"        style="${canEdit?'cursor:pointer':''}">${escHtml(task.status)}</span></td>
      <td><div class="action-row">
        <button class="icon-btn" title="Steps (${taskSteps.length})" onclick="openSteps('${task._spId}','${escHtml(task.name)}')">📋</button>
        <button class="icon-btn" title="Comments (${commentCount})" onclick="openComments('${task._spId}','${escHtml(task.name)}')">💬${commentCount>0?` <sup style="font-size:9px">${commentCount}</sup>`:''}</button>
        ${canEdit?`<button class="icon-btn" title="Edit" onclick="openEditTask('${task._spId}')">✏️</button>`:''}
        ${currentUser.isAdmin?`<button class="icon-btn" title="Delete" onclick="deleteTask('${task._spId}')">🗑</button>`:''}
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

// ── ROLLFORWARD ENGINE ────────────────────────────────────────
async function rollForward() {
  const fromQ   = document.getElementById('rf-from-quarter').value;
  const fromY   = parseInt(document.getElementById('rf-from-year').value);
  const toQ     = document.getElementById('rf-to-quarter').value;
  const toY     = parseInt(document.getElementById('rf-to-year').value);

  if (fromQ === toQ && fromY === toY) {
    alert('Source and target quarter cannot be the same.'); return;
  }
  if (isQuarterLocked(toQ, toY)) {
    alert(`${toQ} ${toY} is already locked. Unlock it first if you want to roll into it.`); return;
  }

  const srcTasks = getTasks().filter(t => t.quarter === fromQ && t.year === fromY);
  if (!srcTasks.length) {
    alert(`No tasks found in ${fromQ} ${fromY}.`); return;
  }

  // ── Step 1: Set close calendar for target quarter ────────────
  const cal = await promptCloseCalendar(toQ, toY);
  if (!cal) return; // user cancelled

  const qEnd = quarterEndDate(toQ, toY);
  const alreadyExists = getTasks().some(t => t.quarter === toQ && t.year === toY);
  if (alreadyExists) {
    if (!confirm(`${toQ} ${toY} already has tasks. Roll forward will add any missing ones. Continue?`)) return;
  }

  showLoadingOverlay(true);
  let copied = 0;
  try {
    for (const task of srcTasks) {
      // Skip tasks that don't apply to the target quarter
      if (!appliesToQuarter(task.applicability, toQ)) continue;

      const exists = getTasks().some(t => t.name === task.name && t.quarter === toQ && t.year === toY);
      if (exists) continue;

      // Compute new due date via workday number (preferred) or date offset fallback
      const srcEnd       = quarterEndDate(fromQ, fromY);
      const srcEndDate   = new Date(srcEnd + 'T00:00:00');
      // Determine workday number: use stored workdayNum, or derive from old calendar
      let taskWdNum = task.workdayNum || null;
      if (!taskWdNum && task.dueDate) {
        taskWdNum = dateToWorkday(task.dueDate, fromQ, fromY);
      }
      if (!taskWdNum && task.dueDate) {
        // Last resort: preserve offset from quarter end and find nearest WD in new cal
        const taskDue  = new Date((task.dueDate||srcEnd) + 'T00:00:00');
        const offset   = Math.round((taskDue - srcEndDate) / 86400000);
        taskWdNum = null; // will fall back to offset below
      }
      // Resolve to actual date in new calendar
      let newDueDate = taskWdNum ? workdayToDate(taskWdNum, toQ, toY) : null;
      if (!newDueDate && task.dueDate) {
        const taskDue  = new Date((task.dueDate||srcEnd) + 'T00:00:00');
        const offset   = Math.round((taskDue - srcEndDate) / 86400000);
        newDueDate     = nearestWorkingDay(addDays(qEnd, offset));
        // Attempt to map to a workday number in the new calendar
        taskWdNum = taskWdNum || dateToWorkday(newDueDate, toQ, toY);
      }

      const newTaskId  = uid();
      const created    = await createListItem(LISTS.tasks, {
        Title: task.name, TaskId: newTaskId,
        TaskType: task.type, Quarter: toQ, Year: String(toY),
        DueDate: newDueDate || '', Status: 'Not Started',
        OwnerId: task.ownerId, Description: task.description,
        Applicability: task.applicability || 'All Quarters',
        WorkdayNum: taskWdNum ? String(taskWdNum) : '',
      });
      const newTaskSpId = created?.id || newTaskId;

      // Copy steps, reset status + sign-offs
      const srcSteps = getStepsForTask(task._spId);
      for (const step of srcSteps) {
        if (!appliesToQuarter(step.applicability, toQ)) continue;
        // Resolve step due date via workday number
        let stepWdNum = step.workdayNum || null;
        if (!stepWdNum && step.dueDate) stepWdNum = dateToWorkday(step.dueDate, fromQ, fromY);
        let stepDue = stepWdNum ? workdayToDate(stepWdNum, toQ, toY) : null;
        if (!stepDue && step.dueDate) {
          const sd  = new Date((step.dueDate||srcEnd)+'T00:00:00');
          const off = Math.round((sd - srcEndDate)/86400000);
          stepDue   = nearestWorkingDay(addDays(qEnd, off));
          stepWdNum = stepWdNum || dateToWorkday(stepDue, toQ, toY);
        }
        await createListItem(LISTS.steps, {
          Title: step.name, StepId: uid(),
          TaskId: String(newTaskSpId),
          StepOrder: String(step.order),
          Status: 'Not Started',
          OwnerId: step.ownerId,
          DueDate: stepDue || null,
          Note: step.note || '',
          Applicability: step.applicability || 'All Quarters',
          WorkdayNum: stepWdNum ? String(stepWdNum) : '',
          RequiresPrev: step.requiresPrev ? 'Yes' : 'No',
        });
      }
      copied++;
    }

    // Lock the source quarter automatically
    await lockQuarter(fromQ, fromY);

    await refreshData();
    alert(`✅ Rolled forward: ${copied} task(s) copied to ${toQ} ${toY}.\n🔒 ${fromQ} ${fromY} has been locked.`);
  } catch(e) {
    showError('Rollforward failed: ' + e.message);
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
  if (!confirm('Unlock this quarter? Team members will be able to edit tasks again.')) return;
  showLoadingOverlay(true);
  try {
    await deleteListItem(LISTS.locks, lockSpId);
    _locks = _locks.filter(l => l._spId !== lockSpId);
    renderAdmin();
  } catch(e) { showError('Unlock failed: ' + e.message); }
  finally { showLoadingOverlay(false); }
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
    .sort((a, b) => b.linkedAt.localeCompare(a.linkedAt));
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
    <p style="font-size:13px;color:var(--gray-600);margin-bottom:16px">
      Paste the SharePoint link to any file — tie-out, workbook, PDF, or folder.
      To copy a link in SharePoint, right-click the file → <strong>Copy link</strong>.
    </p>
    <div class="form-group">
      <label>SharePoint URL</label>
      <input type="url" id="link-url" placeholder="https://moodys.sharepoint.com/sites/finance_home_finrptg/…"
        style="width:100%" oninput="previewLinkLabel()" />
    </div>
    <div class="form-group">
      <label>Display label <span style="font-weight:400;color:var(--gray-400)">(optional — auto-filled from URL)</span></label>
      <input type="text" id="link-label" placeholder="e.g. Q1 2025 Footnote 1 Tie Out" />
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
  const url   = document.getElementById('link-url')?.value?.trim();
  const label = document.getElementById('link-label')?.value?.trim() || labelFromUrl(url);
  if (!url) { alert('Please paste a SharePoint URL.'); return; }
  if (!url.startsWith('http')) { alert('Please enter a valid URL starting with https://'); return; }

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
    };
    const created = await createListItem(LISTS.attachments, fields);
    _attachments.push(normaliseAttachment({ ...fields, id: created?.id || attId }));
    renderStepsPanel();
  } catch(e) { showError('Could not save link: ' + e.message); }
  finally { showLoadingOverlay(false); }
}

async function deleteAttachment(attSpId) {
  if (!confirm('Remove this link?')) return;
  try {
    await deleteListItem(LISTS.attachments, attSpId);
    _attachments = _attachments.filter(a => a._spId !== attSpId);
    renderStepsPanel();
  } catch(e) { showError('Could not remove link: ' + e.message); }
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
        <div class="att-meta">Linked by ${escHtml(a.linkedBy)} · ${escHtml(a.linkedAt)}</div>
      </div>
      ${isOwner || currentUser.isAdmin
        ? `<button class="icon-btn" style="width:24px;height:24px;font-size:11px"
             title="Remove link" onclick="deleteAttachment('${a._spId}')">✕</button>`
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
    UserId:        currentUser._spId || currentUser.id,
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
    .sort((a, b) => (b.tsIso || b.ts) > (a.tsIso || a.ts) ? 1 : -1);
}

// Renders the inline sign-off trail (most recent first, max 5 shown)
function renderSignOffTrail(refId, showAll) {
  const entries = getSignOffsFor(refId);
  if (!entries.length) return '';
  const visible  = showAll ? entries : entries.slice(0, 3);
  const overflow = entries.length - visible.length;
  const rows = visible.map(e => {
    const icon = {
      'Complete':        '✅',
      'Ready for Review': '🔍',
      'In Progress':     '▶️',
      'Not Started':     '⏸',
      'Not Applicable':  '⊘',
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
const STATUS_ORDER = ['Not Started','In Progress','Ready for Review','Complete','Not Applicable'];
// Cycle only moves through the first four — N/A is set manually
const STATUS_CYCLE  = ['Not Started','In Progress','Ready for Review','Complete'];
async function cycleStatus(spId) {
  const task = _tasks.find(t => t._spId === spId); if(!task) return;
  if (isQuarterLocked(task.quarter, task.year)) {
    alert(`Q${task.quarter} ${task.year} is locked and cannot be edited.`); return;
  }
  const prev = task.status;
  const next = STATUS_CYCLE[(STATUS_CYCLE.indexOf(prev)+1) % STATUS_CYCLE.length];
  task.status = next;
  await writeSignOff(spId, 'task', task.name, prev, next);
  renderCurrentView();
  try {
    await updateListItem(LISTS.tasks, spId, { Status: next });
  } catch(e) {
    console.error("Status update failed:", e);
    await refreshData();
  }
}

// ── STEPS PANEL ──────────────────────────────────────────────
let stepsTaskSpId   = null;
let stepsTaskName   = null;
let editingStepId   = null;

function openSteps(taskSpId, taskName) {
  stepsTaskSpId = taskSpId;
  stepsTaskName = taskName;
  document.getElementById('steps-task-title').textContent = taskName;
  renderStepsPanel();
  document.getElementById('steps-overlay').classList.remove('hidden');
}

function closeStepsPanel() {
  document.getElementById('steps-overlay').classList.add('hidden');
  stepsTaskSpId = null; stepsTaskName = null;
  refreshData();
}

function renderStepsPanel() {
  const steps = getStepsForTask(stepsTaskSpId);
  const list  = document.getElementById('steps-list');
  if (!list) return;

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

  list.innerHTML = `
    <div class="steps-progress-header">
      <span class="steps-progress-label">${done} of ${total} complete</span>
      <span class="steps-progress-pct">${pct}%</span>
    </div>
    <div class="progress-bar-wrap" style="margin-bottom:16px">
      <div class="progress-bar-fill" style="width:${pct}%;background:var(--blue)"></div>
    </div>
    ${steps.map((step, idx) => {
      const owner   = getUserById(step.ownerId);
      const ds      = step.dueDate ? deadlineStatus(step.dueDate, step.status) : 'ok';
      const isOwner = currentUser.isAdmin || step.ownerId === currentUser.id || step.ownerId === currentUser._spId;
      const isDone  = step.status === 'Complete';
      const stepTrail = renderSignOffTrail(step._spId, false);
      const stepAtts  = renderAttachmentPanel(step._spId);
      return `<div class="step-row ${isDone ? 'step-done' : ''}">
        <div class="step-number">${idx+1}</div>
        <div class="step-body">
          <div class="step-name ${isDone ? 'step-name-done' : ''}">${escHtml(step.name)}</div>
          <div class="step-meta">
            ${owner ? `<span class="owner-chip"><span class="mini-avatar">${initials(owner.name)}</span>${escHtml(owner.name)}</span>` : '<span class="text-muted">Unassigned</span>'}
            ${step.dueDate || step.workdayNum ? `<span class="deadline-cell"><span class="deadline-dot ${dotClass(ds)}"></span>${formatWorkdayDate(step.workdayNum, step.dueDate)}</span>` : ''}
            ${step.note ? `<span class="step-note">${escHtml(step.note)}</span>` : ''}
          </div>
          ${stepTrail}
          ${stepAtts}
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
          ${currentUser.isAdmin ? `<button class="icon-btn" onclick="deleteStep('${step._spId}')">🗑</button>` : ''}
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
  const next = STATUS_CYCLE[(STATUS_CYCLE.indexOf(prev)+1) % STATUS_CYCLE.length];

  // Only check gate when advancing forward (not cycling back to Not Started)
  const movingForward = STATUS_ORDER.indexOf(next) > STATUS_ORDER.indexOf(prev);
  if (movingForward) {
    const blocker = getBlockingPredecessor(step);
    if (blocker) {
      const go = confirm(
        `⚠ Gate warning

` +
        `"${blocker.name}" (step ${blocker.order}) is not yet Complete.

` +
        `This step is configured to require the previous step to be finished first.

` +
        `Proceed anyway?`
      );
      if (!go) return;
    }
  }

  step.status = next;
  await writeSignOff(stepSpId, 'step', step.name, prev, next);
  renderStepsPanel();
  try {
    await updateListItem(LISTS.steps, stepSpId, { Status: next });
  } catch(e) { console.error("Step status update failed:", e); await refreshData(); }
}

// Admin force-unlock: mark the blocking predecessor Complete, then advance this step
async function forceUnlockStep(stepSpId) {
  const step    = _steps.find(s => s._spId === stepSpId); if(!step) return;
  const blocker = getBlockingPredecessor(step);
  if (!blocker) { cycleStepStatus(stepSpId); return; }
  if (!confirm(
    `Force unlock?

` +
    `This will mark "${blocker.name}" as Complete (bypassing the gate) ` +
    `and then advance "${step.name}".

` +
    `A sign-off log entry will be created for both changes.`
  )) return;
  // Mark blocker Complete
  blocker.status = 'Complete';
  await writeSignOff(blocker._spId, 'step', blocker.name, blocker.status, 'Complete');
  await updateListItem(LISTS.steps, blocker._spId, { Status: 'Complete' });
  // Now advance this step
  await cycleStepStatus(stepSpId);
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
  const ownerOpts = getUsers().map(u =>
    `<option value="${u.id||u._spId}" ${step?.ownerId===(u.id||u._spId)?'selected':''}>${escHtml(u.name)}</option>`).join('');
  document.getElementById('modal-title').textContent = step ? 'Edit Step' : 'Add Step';
  document.getElementById('modal-body').innerHTML = `
    <div class="form-group"><label>Step Name</label>
      <input type="text" id="sf-name" value="${escHtml(step?.name||'')}" placeholder="e.g. Rollforward" /></div>
    <div class="form-group"><label>Order</label>
      <input type="number" id="sf-order" value="${step?.order ?? steps.length+1}" min="1" /></div>
    <div class="form-group"><label>Assigned To</label>
      <select id="sf-owner"><option value="">— Unassigned —</option>${ownerOpts}</select></div>
    <div class="form-group"><label>Due Date</label>
      <div style="display:grid;grid-template-columns:90px 1fr;gap:10px;align-items:end">
        <div>
          <label style="font-size:11px;color:var(--gray-400);display:block;margin-bottom:4px">Workday #</label>
          <input type="number" id="sf-workday" value="${step?.workdayNum||''}" min="1" max="60" placeholder="WD #" style="width:100%;padding:10px 8px;border:1.5px solid var(--gray-200);border-radius:var(--radius);font-family:inherit;font-size:14px"
            oninput="(()=>{const pt=_tasks.find(t=>t._spId===stepsTaskSpId);if(pt)resolveWdToDate('sf-workday','sf-due',pt.quarter,pt.year)})()" />
        </div>
        <input type="date" id="sf-due" value="${step?.dueDate||''}" />
      </div>
    </div>
    <div class="form-group"><label>Status</label>
      <select id="sf-status">${STATUS_ORDER.map(s=>`<option ${step?.status===s?'selected':''}>${s}</option>`).join('')}</select></div>
    <div class="form-group"><label>Note (optional)</label>
      <input type="text" id="sf-note" value="${escHtml(step?.note||'')}" placeholder="Any brief note…" /></div>
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
      <p style="font-size:11px;color:var(--gray-400);margin-top:4px;margin-left:26px">
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
  const note          = document.getElementById('sf-note').value.trim();
  const applicability = document.getElementById('sf-applicability').value;
  const workdayNum    = parseInt(document.getElementById('sf-workday')?.value) || null;
  const requiresPrev  = document.getElementById('sf-requires-prev')?.checked || false;
  if (!name) { alert('Please enter a step name.'); return; }
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
      DueDate: resolvedDue || null, Status: status, Note: note,
      TaskId: stepsTaskSpId, Applicability: applicability,
      WorkdayNum: workdayNum ? String(workdayNum) : '',
      RequiresPrev: requiresPrev ? 'Yes' : 'No',
    };
    if (editingStepId) {
      await updateListItem(LISTS.steps, editingStepId, fields);
      const idx = _steps.findIndex(s => s._spId === editingStepId);
      if (idx >= 0) Object.assign(_steps[idx], { name, order, ownerId, dueDate: resolvedDue, status, note, applicability, workdayNum, requiresPrev });
    } else {
      fields.StepId = uid();
      const created = await createListItem(LISTS.steps, fields);
      _steps.push(normaliseStep({ ...fields, id: created?.id || fields.StepId }));
    }
    renderStepsPanel();
  } catch(e) { showError("Could not save step: "+e.message); }
  finally { showLoadingOverlay(false); }
}

async function deleteStep(stepSpId) {
  if (!confirm('Delete this step?')) return;
  showLoadingOverlay(true);
  try {
    await deleteListItem(LISTS.steps, stepSpId);
    _steps = _steps.filter(s => s._spId !== stepSpId);
    renderStepsPanel();
  } catch(e) { showError("Could not delete step: "+e.message); }
  finally { showLoadingOverlay(false); }
}

// ── STEP TEMPLATES — manage default steps per task template ───
let editingStepTemplateId  = null;
let stepsTemplateId        = null;
let stepsTemplateName      = null;

function openStepTemplates(templateSpId, templateName) {
  stepsTemplateId   = templateSpId;
  stepsTemplateName = templateName;
  document.getElementById('modal-title').textContent = `Steps for "${templateName}"`;
  renderStepTemplatesModal();
  openModal();
}

function renderStepTemplatesModal() {
  const tpSteps = getStepTemplatesForTemplate(stepsTemplateId);
  const ownerOpts = () => getUsers().map(u =>
    `<option value="${u.id||u._spId}">${escHtml(u.name)}</option>`).join('');

  document.getElementById('modal-body').innerHTML = `
    <div id="step-tpl-list">
      ${tpSteps.length === 0
        ? '<p class="text-muted" style="font-size:13px;margin-bottom:12px">No default steps yet.</p>'
        : tpSteps.map((st,i) => {
            const owner = getUserById(st.defaultOwnerId);
            return `<div class="step-tpl-row">
              <span class="step-number">${i+1}</span>
              <div class="step-body">
                <div class="step-name">${escHtml(st.name)}${st.requiresPrev?' <span title="Requires previous step" style="font-size:10px;background:#fef3c7;color:#92400e;border-radius:4px;padding:1px 5px">🔒 gated</span>':''}</div>
                <div class="step-meta">
                  ${owner?`<span class="owner-chip"><span class="mini-avatar">${initials(owner.name)}</span>${escHtml(owner.name)}</span>`:''}
                  <span class="text-muted">${st.dueDaysFromQtrEnd} days after qtr end</span>
                </div>
              </div>
              <div class="step-actions">
                <button class="icon-btn" onclick="openEditStepTemplate('${st._spId}')">✏️</button>
                <button class="icon-btn" onclick="deleteStepTemplate('${st._spId}')">🗑</button>
              </div>
            </div>`;
          }).join('')
      }
    </div>
    <div style="border-top:1px solid var(--gray-100);margin-top:12px;padding-top:14px">
      <p style="font-size:12px;font-weight:600;color:var(--gray-600);text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px">Add default step</p>
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

async function saveStepTemplate() {
  const name       = document.getElementById('stf-name').value.trim();
  const order      = parseInt(document.getElementById('stf-order').value) || 1;
  const days       = parseInt(document.getElementById('stf-days').value) || 0;
  const ownerId    = document.getElementById('stf-owner').value;
  const reqPrev    = document.getElementById('stf-requires-prev')?.checked || false;
  if (!name) { alert('Please enter a step name.'); return; }
  showLoadingOverlay(true);
  try {
    const id = uid();
    const created = await createListItem(LISTS.stepTemplates, {
      Title: name, StepTemplateId: id,
      TemplateId: stepsTemplateId,
      StepOrder: String(order),
      DueDaysFromQtrEnd: String(days),
      DefaultOwnerId: ownerId,
      RequiresPrev: reqPrev ? 'Yes' : 'No',
    });
    _stepTemplates.push(normaliseStepTemplate({
      Title: name, StepTemplateId: id,
      TemplateId: stepsTemplateId,
      StepOrder: String(order),
      DueDaysFromQtrEnd: String(days),
      DefaultOwnerId: ownerId,
      RequiresPrev: reqPrev ? 'Yes' : 'No',
      id: created?.id || id,
    }));
    renderStepTemplatesModal();
  } catch(e) { showError("Could not save step template: "+e.message); }
  finally { showLoadingOverlay(false); }
}

async function deleteStepTemplate(spId) {
  if (!confirm('Remove this default step?')) return;
  showLoadingOverlay(true);
  try {
    await deleteListItem(LISTS.stepTemplates, spId);
    _stepTemplates = _stepTemplates.filter(s => s._spId !== spId);
    renderStepTemplatesModal();
  } catch(e) { showError("Could not delete step template: "+e.message); }
  finally { showLoadingOverlay(false); }
}

// ── CALENDAR ─────────────────────────────────────────────────
function changeCalMonth(dir) {
  calMonth+=dir;
  if(calMonth>11){calMonth=0;calYear++;}
  if(calMonth<0) {calMonth=11;calYear--;}
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
    const dayTasks=getTasks().filter(t=>t.dueDate===dateStr);
    let chips=dayTasks.slice(0,3).map(t=>
      `<span class="cal-task-chip" style="background:${calChipBg(t.type)};color:${calChipFg(t.type)}"
        onclick="openEditTask('${t._spId}')" title="${escHtml(t.name)}">${escHtml(t.name)}</span>`).join('');
    if(dayTasks.length>3) chips+=`<span style="font-size:10px;color:var(--gray-400)">+${dayTasks.length-3} more</span>`;
    html+=`<div class="cal-day${isToday?' today':''}"><div class="cal-day-num">${d}</div>${chips}</div>`;
  }
  grid.innerHTML=html;
}

// ── TEAM ──────────────────────────────────────────────────────
function renderTeam() {
  const grid=document.getElementById('team-grid'); if(!grid) return;
  grid.innerHTML=getUsers().map(user=>{
    const uid=user.id||user._spId;
    const myTasks=getTasks().filter(t=>t.ownerId===uid);
    const activeMyTasks=myTasks.filter(t=>t.status!=='Not Applicable');
    const complete=activeMyTasks.filter(t=>t.status==='Complete').length;
    const overdue=activeMyTasks.filter(t=>deadlineStatus(t.dueDate,t.status)==='overdue').length;
    const pct=activeMyTasks.length?Math.round(complete/activeMyTasks.length*100):0;
    const taskList=myTasks.slice(0,4).map(t=>{
      const ds=deadlineStatus(t.dueDate,t.status);
      return `<div class="team-task-item">
        <span class="team-task-name">${escHtml(t.name)}</span>
        <span class="deadline-dot ${dotClass(ds)}" style="width:8px;height:8px;border-radius:50%;display:inline-block;flex-shrink:0"></span>
      </div>`;
    }).join('')||'<div class="text-muted" style="font-size:12px;padding:6px 0">No tasks assigned</div>';
    return `<div class="team-card">
      <div class="team-card-header">
        <div class="team-avatar">${initials(user.name)}</div>
        <div><div class="team-name">${escHtml(user.name)}</div><div class="team-role">${escHtml(user.role)}</div></div>
      </div>
      <div class="team-stats">
        <div class="team-stat"><div class="team-stat-val">${myTasks.length}</div><div class="team-stat-label">Total</div></div>
        <div class="team-stat"><div class="team-stat-val">${complete}</div><div class="team-stat-label">Done</div></div>
        <div class="team-stat"><div class="team-stat-val" style="color:var(--deadline-overdue)">${overdue}</div><div class="team-stat-label">Overdue</div></div>
      </div>
      <div class="progress-bar-wrap" style="margin-bottom:14px">
        <div class="progress-bar-fill" style="width:${pct}%"></div>
      </div>
      <div class="team-task-list">${taskList}</div>
    </div>`;
  }).join('');
}

// ── ADMIN ─────────────────────────────────────────────────────
function renderAdmin() {
  if(!currentUser?.isAdmin) return;
  renderCustomHolidays();
  renderCloseCalendarsPanel();

  // Locked quarters panel
  const lkEl = document.getElementById('admin-locks-list');
  if (lkEl) {
    if (!_locks.length) {
      lkEl.innerHTML = '<p class="text-muted" style="font-size:13px;padding:12px 22px">No quarters locked yet.</p>';
    } else {
      lkEl.innerHTML = _locks
        .sort((a,b) => `${b.year}${b.quarter}` > `${a.year}${a.quarter}` ? 1 : -1)
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
        ${u._spId!==currentUser._spId?`<button class="icon-btn" onclick="deleteUser('${u._spId}')">🗑</button>`:''}
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
        <button class="icon-btn" onclick="deleteTemplate('${tp._spId}')">🗑</button>
      </div>
    </div>`;
  }).join('');
}

async function applyTemplate() {
  const quarter = document.getElementById('template-quarter-select').value;
  const year    = parseInt(document.getElementById('template-year-select').value);
  const qEnd    = quarterEndDate(quarter, year);
  let added = 0;
  for (const tp of getTemplates()) {
    const exists = getTasks().some(t => t.name===tp.name && t.quarter===quarter && t.year===year);
    if (!exists) {
      const taskId   = uid();
      const created  = await createListItem(LISTS.tasks, {
        Title: tp.name, TaskId: taskId, TaskType: tp.type,
        Quarter: quarter, Year: String(year),
        DueDate: addDays(qEnd, tp.dueDaysFromQtrEnd),
        Status: 'Not Started', OwnerId: tp.defaultOwnerId, Description: '',
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
          Note: '',
        });
      }
      added++;
    }
  }
  await refreshData();
  alert(`Applied template: ${added} task(s) added to ${quarter} ${year}.`);
}

// ── TASK MODAL ────────────────────────────────────────────────
function openAddTask()        { editingTaskId=null;   showTaskModal(null); }
function openEditTask(spId)   { editingTaskId=spId;   showTaskModal(_tasks.find(t=>t._spId===spId)); }
function showTaskModal(task) {
  document.getElementById('modal-title').textContent=task?'Edit Task':'Add New Task';
  const ownerOpts=getUsers().map(u=>`<option value="${u.id||u._spId}" ${task?.ownerId===(u.id||u._spId)?'selected':''}>${escHtml(u.name)}</option>`).join('');
  const cur=new Date().getFullYear();
  const yearOpts=[cur-1,cur,cur+1].map(y=>`<option value="${y}" ${task?.year===y?'selected':''}>${y}</option>`).join('');
  document.getElementById('modal-body').innerHTML=`
    <div class="form-group"><label>Deliverable Name</label>
      <input type="text" id="tf-name" value="${escHtml(task?.name||'')}" placeholder="e.g. 10-Q Filing" /></div>
    <div class="form-group"><label>Type</label>
      <select id="tf-type">${['Close', 'Financial Report', 'Master SS', 'Ops Book', 'Other', 'Press Release', 'Post-Filing', 'Pre-Filing'].map(t=>`<option ${task?.type===t?'selected':''}>${t}</option>`).join('')}</select></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div class="form-group"><label>Quarter</label>
        <select id="tf-quarter">${['Q1','Q2','Q3','Q4'].map(q=>`<option ${task?.quarter===q?'selected':''}>${q}</option>`).join('')}</select></div>
      <div class="form-group"><label>Year</label>
        <select id="tf-year">${yearOpts}</select></div>
    </div>
    <div class="form-group"><label>Due Date</label>
      <div style="display:grid;grid-template-columns:90px 1fr;gap:10px;align-items:end">
        <div>
          <label style="font-size:11px;color:var(--gray-400);display:block;margin-bottom:4px">Workday #</label>
          <input type="number" id="tf-workday" value="${task?.workdayNum||''}" min="1" max="60" placeholder="WD #" style="width:100%;padding:10px 8px;border:1.5px solid var(--gray-200);border-radius:var(--radius);font-family:inherit;font-size:14px" oninput="resolveWdToDate('tf-workday','tf-due',document.getElementById('tf-quarter')?.value,document.getElementById('tf-year')?.value)" />
        </div>
        <input type="date" id="tf-due" value="${task?.dueDate||''}" />
      </div>
      <p style="font-size:11px;color:var(--gray-400);margin-top:4px">Enter a workday number to auto-fill the date, or set the date directly.</p>
    </div>
    <div class="form-group"><label>Assigned To</label>
      <select id="tf-owner">${ownerOpts}</select></div>
    <div class="form-group"><label>Status</label>
      <select id="tf-status">${STATUS_ORDER.map(s=>`<option ${task?.status===s?'selected':''}>${s}</option>`).join('')}</select></div>
    <div class="form-group"><label>Applicability</label>
      <select id="tf-applicability">
        <option value="All Quarters" ${(!task?.applicability||task?.applicability==='All Quarters')?'selected':''}>All Quarters</option>
        <option value="10-K only (Q4)" ${task?.applicability==='10-K only (Q4)'?'selected':''}>10-K only (Q4)</option>
        <option value="10-Q only (Q1, Q2, Q3)" ${task?.applicability==='10-Q only (Q1, Q2, Q3)'?'selected':''}>10-Q only (Q1, Q2, Q3)</option>
      </select></div>
    <div class="form-group"><label>Description (optional)</label>
      <textarea id="tf-desc" rows="2" placeholder="Brief notes…">${escHtml(task?.description||'')}</textarea></div>
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
  const wdNum   = parseInt(document.getElementById('tf-workday')?.value) || null;
  let   dueDate = document.getElementById('tf-due').value;
  // If workday set and calendar exists, resolve to real date
  if (wdNum) {
    const resolved = workdayToDate(wdNum, quarter, year);
    if (resolved) dueDate = resolved;
  }
  if(!name||!dueDate){alert('Please fill in the task name and due date.');return;}
  const fields={
    Title:       name,
    TaskType:    document.getElementById('tf-type').value,
    Quarter:     quarter,
    Year:        String(year),
    DueDate:     dueDate,
    OwnerId:     document.getElementById('tf-owner').value,
    Status:      document.getElementById('tf-status').value,
    Description: document.getElementById('tf-desc').value.trim(),
    Applicability: document.getElementById('tf-applicability')?.value || 'All Quarters',
    WorkdayNum:  wdNum ? String(wdNum) : '',
  };
  closeAllModals();
  showLoadingOverlay(true);
  try {
    if(editingTaskId) {
      await updateListItem(LISTS.tasks, editingTaskId, fields);
    } else {
      fields.TaskId = uid();
      await createListItem(LISTS.tasks, fields);
    }
    await refreshData();
  } catch(e) { showError("Could not save task: "+e.message); }
  finally { showLoadingOverlay(false); }
}
async function deleteTask(spId) {
  if(!confirm('Delete this task? This cannot be undone.')) return;
  showLoadingOverlay(true);
  try { await deleteListItem(LISTS.tasks, spId); await refreshData(); }
  catch(e) { showError("Could not delete task: "+e.message); }
  finally { showLoadingOverlay(false); }
}

// ── COMMENTS ─────────────────────────────────────────────────
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
    .sort((a,b)=>(a.ts||0)-(b.ts||0));
  if(!taskComments.length){
    list.innerHTML='<div class="text-muted" style="font-size:13px;padding:8px 0">No comments yet.</div>';
    return;
  }
  list.innerHTML=taskComments.map(c=>{
    const author=getUserById(c.authorId);
    return `<div class="comment-item">
      <span class="comment-author">${author?escHtml(author.name):'Unknown'}</span>
      <span class="comment-time">${escHtml(c.time)}</span>
      <div class="comment-text">${escHtml(c.text)}</div>
    </div>`;
  }).join('');
  list.scrollTop=list.scrollHeight;
}
async function addComment() {
  const input=document.getElementById('comment-input');
  const text=input.value.trim(); if(!text||!commentingTaskId) return;
  input.value='';
  const now=new Date();
  await createListItem(LISTS.comments, {
    Title:       text.slice(0,50),
    CommentId:   uid(),
    TaskId:      commentingTaskId,
    AuthorId:    currentUser._spId||currentUser.id,
    CommentText: text,
    CommentTime: now.toLocaleString('en-US',{month:'short',day:'numeric',hour:'numeric',minute:'2-digit'}),
    Timestamp:   String(now.getTime()),
  });
  await refreshComments();
}
function closeCommentModal() {
  stopCommentPolling();
  document.getElementById('comment-overlay').classList.add('hidden');
  commentingTaskId=null;
  refreshData();
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
    <div class="form-group"><label>PIN (4 digits)</label>
      <input type="password" id="uf-pin" maxlength="4" value="${user?.pin||''}" placeholder="····" /></div>
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
  const pin=document.getElementById('uf-pin').value.trim();
  const isAdmin=document.getElementById('uf-admin').value==='true';
  if(!name||pin.length!==4){alert('Please fill in all fields. PIN must be 4 digits.');return;}
  closeAllModals(); showLoadingOverlay(true);
  try {
    const fields={ Title:name, FullName:name, JobRole:role, Email:email, PIN:pin, IsAdmin:isAdmin?'Yes':'No' };
    if(existingSpId) { await updateListItem(LISTS.users, existingSpId, fields); }
    else             { fields.UserId=uid(); await createListItem(LISTS.users, fields); }
    await refreshData();
  } catch(e) { showError("Could not save user: "+e.message); }
  finally { showLoadingOverlay(false); }
}
async function deleteUser(spId) {
  if(!confirm('Remove this team member?')) return;
  showLoadingOverlay(true);
  try { await deleteListItem(LISTS.users, spId); await refreshData(); }
  catch(e) { showError("Could not delete user: "+e.message); }
  finally { showLoadingOverlay(false); }
}

// ── TEMPLATE MODAL ────────────────────────────────────────────
function openAddTemplate()       { showTemplateModal(null); }
function openEditTemplate(spId)  { showTemplateModal(_templates.find(t=>t._spId===spId)); }
function showTemplateModal(tp) {
  const ownerOpts=getUsers().map(u=>`<option value="${u.id||u._spId}" ${tp?.defaultOwnerId===(u.id||u._spId)?'selected':''}>${escHtml(u.name)}</option>`).join('');
  document.getElementById('modal-title').textContent=tp?'Edit Template':'Add Template Task';
  document.getElementById('modal-body').innerHTML=`
    <div class="form-group"><label>Task Name</label>
      <input type="text" id="tpl-name" value="${escHtml(tp?.name||'')}" /></div>
    <div class="form-group"><label>Type</label>
      <select id="tpl-type">${['Close', 'Financial Report', 'Master SS', 'Ops Book', 'Other', 'Press Release', 'Post-Filing', 'Pre-Filing'].map(t=>`<option ${tp?.type===t?'selected':''}>${t}</option>`).join('')}</select></div>
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
  if(!name||isNaN(days)){alert('Please fill in all fields.');return;}
  closeAllModals(); showLoadingOverlay(true);
  try {
    const fields={ Title:name, TaskType:type, DueDaysFromQtrEnd:String(days), DefaultOwnerId:ownerId };
    if(existingSpId) { await updateListItem(LISTS.templates, existingSpId, fields); }
    else             { fields.TemplateId=uid(); await createListItem(LISTS.templates, fields); }
    await refreshData();
  } catch(e) { showError("Could not save template: "+e.message); }
  finally { showLoadingOverlay(false); }
}
async function deleteTemplate(spId) {
  if(!confirm('Delete this template?')) return;
  showLoadingOverlay(true);
  try { await deleteListItem(LISTS.templates, spId); await refreshData(); }
  catch(e) { showError("Could not delete template: "+e.message); }
  finally { showLoadingOverlay(false); }
}

// ── MODAL HELPERS ─────────────────────────────────────────────
function openModal()      { document.getElementById('modal-overlay').classList.remove('hidden'); }
function closeAllModals() { document.getElementById('modal-overlay').classList.add('hidden'); editingTaskId=null; }
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
}

// ── INIT ──────────────────────────────────────────────────────
(function init() {
  msalInstance.handleRedirectPromise().catch(console.error);

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    loadAllData().then(() => {
      populateUserSelect();
      const savedId = sessionStorage.getItem('ft_session');
      if (savedId) {
        const user = _users.find(u => (u._spId||u.id) === savedId);
        if (user) { currentUser = user; launchApp(); return; }
      }
      document.getElementById('ms-login-row').classList.add('hidden');
      document.getElementById('pin-login-row').classList.remove('hidden');
    });
  }

  document.getElementById('login-pin').addEventListener('keydown', e => {
    if(e.key === 'Enter') login();
  });
})();

// ═══════════════════════════════════════════════════════════════
// ── MY TASKS VIEW ────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function renderMyTasks() {
  const uid      = currentUser._spId || currentUser.id;
  const filter   = document.getElementById('mytasks-status-filter')?.value || 'open';
  const titleEl  = document.getElementById('mytasks-title');
  if (titleEl) titleEl.textContent = `My Tasks — ${currentUser.name.split(' ')[0]}`;

  let myTasks = getTasks().filter(t => t.ownerId === uid);
  if (filter === 'open')     myTasks = myTasks.filter(t => t.status !== 'Complete' && t.status !== 'Not Applicable');
  if (filter === 'Complete') myTasks = myTasks.filter(t => t.status === 'Complete');

  // Also include steps assigned to me across all tasks
  let mySteps = _steps.filter(s => s.ownerId === uid);
  if (filter === 'open')     mySteps = mySteps.filter(s => s.status !== 'Complete' && s.status !== 'Not Applicable');
  if (filter === 'Complete') mySteps = mySteps.filter(s => s.status === 'Complete');

  const container = document.getElementById('mytasks-content');
  if (!container) return;

  if (!myTasks.length && !mySteps.length) {
    container.innerHTML = `<div class="empty-state"><div class="empty-icon">🎉</div>
      <p>No ${filter === 'open' ? 'open' : ''} items assigned to you.</p></div>`;
    return;
  }

  // Group tasks by quarter
  const quarters = [...new Set(myTasks.map(t => `${t.quarter} ${t.year}`))].sort().reverse();

  let html = '';

  // --- Tasks section ---
  if (myTasks.length) {
    html += `<div class="mytasks-section-title">Tasks assigned to me</div>`;
    quarters.forEach(qLabel => {
      const qTasks = myTasks.filter(t => `${t.quarter} ${t.year}` === qLabel);
      const locked  = isQuarterLocked(qTasks[0].quarter, qTasks[0].year);
      html += `<div class="card mt-24">
        <div class="card-header">
          <h3 class="card-title">${qLabel} ${locked ? '🔒' : ''}</h3>
          <span class="text-muted" style="font-size:12px">${qTasks.length} task${qTasks.length!==1?'s':''}</span>
        </div>
        <table class="task-table">
          <thead><tr>
            <th>Deliverable</th><th>Type</th><th>Due Date</th><th>Status</th><th>Actions</th>
          </tr></thead>
          <tbody>
            ${qTasks.map(task => {
              const ds       = deadlineStatus(task.dueDate, task.status);
              const canEdit  = !locked;
              const trail    = renderSignOffTrail(task._spId, false);
              const appBadge = task.applicability && task.applicability !== 'All Quarters'
                ? `<span class="app-badge ${task.applicability.startsWith('10-K')?'app-badge-10k':'app-badge-10q'}">${task.applicability.startsWith('10-K')?'10-K':'10-Q'}</span>` : '';
              return `<tr>
                <td><div class="task-name">${escHtml(task.name)}${appBadge}</div>${trail}</td>
                <td><span class="badge ${typeBadgeClass(task.type)}">${escHtml(task.type)}</span></td>
                <td><div class="deadline-cell"><span class="deadline-dot ${dotClass(ds)}"></span>${formatWorkdayDate(task.workdayNum, task.dueDate)}${ds==='overdue'?'<span class="text-danger fw-600" style="font-size:11px"> OVERDUE</span>':''}</div></td>
                <td><span class="status-badge ${statusBadgeClass(task.status)}" onclick="${canEdit?`cycleStatus('${task._spId}')`:''}" style="${canEdit?'cursor:pointer':''}">${escHtml(task.status)}</span></td>
                <td><div class="action-row">
                  <button class="icon-btn" onclick="openSteps('${task._spId}','${escHtml(task.name)}')">📋</button>
                  <button class="icon-btn" onclick="openComments('${task._spId}','${escHtml(task.name)}')">💬</button>
                  ${canEdit?`<button class="icon-btn" onclick="openEditTask('${task._spId}')">✏️</button>`:''}
                </div></td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>`;
    });
  }

  // --- Steps section ---
  if (mySteps.length) {
    html += `<div class="mytasks-section-title" style="margin-top:28px">Steps assigned to me</div>`;
    html += `<div class="card mt-24">
      <div class="card-header"><h3 class="card-title">Step Checklist</h3>
        <span class="text-muted" style="font-size:12px">${mySteps.filter(s=>s.status==='Complete').length}/${mySteps.length} complete</span>
      </div>
      <div style="padding:4px 0">
        ${mySteps.map(step => {
          const parentTask = _tasks.find(t => t._spId === step.taskId);
          const ds         = step.dueDate ? deadlineStatus(step.dueDate, step.status) : 'ok';
          const locked     = parentTask ? isQuarterLocked(parentTask.quarter, parentTask.year) : false;
          const isDone     = step.status === 'Complete';
          const trail      = renderSignOffTrail(step._spId, false);
          return `<div class="step-row ${isDone?'step-done':''}">
            <div class="step-body" style="flex:1">
              <div class="step-name ${isDone?'step-name-done':''}">${escHtml(step.name)}</div>
              <div class="step-meta">
                ${parentTask?`<span class="text-muted" style="font-size:11px">↳ ${escHtml(parentTask.name)} · ${parentTask.quarter} ${parentTask.year}</span>`:''}
                ${step.dueDate||step.workdayNum?`<span class="deadline-cell"><span class="deadline-dot ${dotClass(ds)}"></span>${formatWorkdayDate(step.workdayNum, step.dueDate)}</span>`:''}
              </div>
              ${trail}
            </div>
            <div class="step-actions">
              ${!locked?`<span class="status-badge ${statusBadgeClass(step.status)}" style="cursor:pointer;font-size:11px" onclick="cycleStepStatus('${step._spId}')">${escHtml(step.status)}</span>`:`<span class="status-badge ${statusBadgeClass(step.status)}" style="font-size:11px">${escHtml(step.status)}</span>`}
            </div>
          </div>`;
        }).join('')}
      </div>
    </div>`;
  }

  container.innerHTML = html;
}

// ═══════════════════════════════════════════════════════════════
// ── KANBAN BOARD ─────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function initKanbanSelects() {
  const cur = new Date();
  const m   = cur.getMonth();
  const q   = m<3?'Q1':m<6?'Q2':m<9?'Q3':'Q4';
  const qEl = document.getElementById('kanban-quarter'); if(qEl) qEl.value = q;
  const yEl = document.getElementById('kanban-year');    if(yEl) yEl.value = cur.getFullYear();
}

function renderKanban() {
  const mode    = document.getElementById('kanban-mode')?.value    || 'personal';
  const quarter = document.getElementById('kanban-quarter')?.value || 'Q1';
  const year    = parseInt(document.getElementById('kanban-year')?.value || new Date().getFullYear());
  const locked  = isQuarterLocked(quarter, year);
  const uid     = currentUser._spId || currentUser.id;
  const subEl   = document.getElementById('kanban-sub');
  if (subEl) subEl.textContent = mode==='personal'
    ? `${currentUser.name.split(' ')[0]}'s board · ${quarter} ${year}${locked?' 🔒':''}`
    : `Full team · ${quarter} ${year}${locked?' 🔒':''}`;

  let tasks = getTasks().filter(t => t.quarter===quarter && t.year===year);
  if (mode === 'personal') tasks = tasks.filter(t => t.ownerId===uid);

  const board = document.getElementById('kanban-board');
  if (!board) return;

  const cols = STATUS_CYCLE.map(status => {
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
        ${steps.length?`<div class="kanban-step-bar"><div class="kanban-step-fill" style="width:${steps.length?Math.round(doneS/steps.length*100):0}%"></div></div><div class="kanban-step-label">${doneS}/${steps.length} steps</div>`:''}
        <div class="kanban-card-footer">
          <div class="owner-chip">
            <div class="mini-avatar" style="width:20px;height:20px;font-size:9px">${owner?initials(owner.name):'?'}</div>
            ${mode==='team'&&owner?`<span style="font-size:11px">${escHtml(owner.name.split(' ')[0])}</span>`:''}
          </div>
          <div class="deadline-cell" style="font-size:11px">
            <span class="deadline-dot ${dotClass(ds)}"></span>${formatWorkdayDate(task.workdayNum, task.dueDate)}
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
  await writeSignOff(_dragTaskSpId, 'task', task.name, prev, newStatus);
  renderKanban();
  try {
    await updateListItem(LISTS.tasks, _dragTaskSpId, { Status: newStatus });
  } catch(e) { console.error('Kanban drop failed:', e); await refreshData(); }
  _dragTaskSpId = null;
}

// ═══════════════════════════════════════════════════════════════
// ── SUMMARY REPORT VIEW ──────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function initReportSelects() {
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
  const review       = activeTasks2.filter(t=>t.status==='Ready for Review').length;
  const inprog       = activeTasks2.filter(t=>t.status==='In Progress').length;
  const notstart     = activeTasks2.filter(t=>t.status==='Not Started').length;
  const overdue      = activeTasks2.filter(t=>deadlineStatus(t.dueDate,t.status)==='overdue').length;
  const pct          = total ? Math.round(complete/total*100) : 0;

  const byType   = {};
  tasks.forEach(t => { byType[t.type] = (byType[t.type]||0)+1; });

  const taskRows = tasks.map(task => {
    const owner    = getUserById(task.ownerId);
    const steps    = getStepsForTask(task._spId);
    const doneS    = steps.filter(s=>s.status==='Complete').length;
    const signoffs = getSignOffsFor(task._spId);
    const lastSO   = signoffs[0];
    const ds       = deadlineStatus(task.dueDate, task.status);
    return `<tr>
      <td style="font-weight:600">${escHtml(task.name)}</td>
      <td><span class="badge ${typeBadgeClass(task.type)}" style="font-size:10px">${escHtml(task.type)}</span></td>
      <td>${owner?escHtml(owner.name):'—'}</td>
      <td><div class="deadline-cell" style="font-size:12px"><span class="deadline-dot ${dotClass(ds)}"></span>${formatWorkdayDate(task.workdayNum, task.dueDate)}</div></td>
      <td><span class="status-badge ${statusBadgeClass(task.status)}" style="font-size:11px">${escHtml(task.status)}</span></td>
      <td style="font-size:11px;color:var(--gray-600)">${steps.length?`${doneS}/${steps.length}`:'—'}</td>
      <td style="font-size:11px;color:var(--gray-600)">${lastSO?`${escHtml(lastSO.userName)} · ${escHtml(lastSO.ts)}`:'—'}</td>
    </tr>`;
  }).join('');

  el.innerHTML = `
    <div class="report-header">
      <div>
        <div class="report-title">${quarter} ${year} Financial Reporting</div>
        <div class="report-subtitle">Quarter Summary Report · Generated ${nowLabel()}${locked?' · 🔒 Locked':''}</div>
      </div>
      <div class="report-logo">◈ Financial Reporting Tracker</div>
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
          <span style="color:var(--gray-600)">${count} task${count!==1?'s':''}</span>
        </div>`).join('')}
    </div>

    <table class="task-table report-table">
      <thead><tr>
        <th>Deliverable</th><th>Type</th><th>Owner</th>
        <th>Due Date</th><th>Status</th><th>Steps</th><th>Last Sign-Off</th>
      </tr></thead>
      <tbody>${taskRows}</tbody>
    </table>

    <div class="report-footer">
      Financial Reporting Tracker · ${quarter} ${year} · Confidential · ${nowLabel()}
    </div>
  `;
}

// ═══════════════════════════════════════════════════════════════
// ── EXCEL EXPORT ─────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════
function exportExcel() {
  const quarter = document.getElementById('report-quarter')?.value || 'Q1';
  const year    = parseInt(document.getElementById('report-year')?.value || new Date().getFullYear());
  const tasks   = getTasks().filter(t => t.quarter===quarter && t.year===year);

  // Build CSV content for three tabs (we use multi-sheet CSV trick via XLSX-style tab encoding)
  // Since we have no XLSX library, generate a proper HTML table that Excel can open natively
  const escape = v => `"${String(v||'').replace(/"/g,'""')}"`;

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

  // --- Sheet 3: Sign-Off Log ---
  const soHeader = ['Ref Type','Ref Name','Changed By','From Status','To Status','Date & Time'];
  const soRows   = [];
  tasks.forEach(t => {
    getSignOffsFor(t._spId).forEach(s => soRows.push(['Task', t.name, s.userName, s.fromStatus, s.toStatus, s.ts]));
    getStepsForTask(t._spId).forEach(step => {
      getSignOffsFor(step._spId).forEach(s => soRows.push(['Step', step.name, s.userName, s.fromStatus, s.toStatus, s.ts]));
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
    `<option value="${u.id||u._spId}">${escHtml(u.name)}</option>`).join('');
  document.getElementById('modal-title').textContent = `Bulk Edit — ${spIds.length} tasks`;
  document.getElementById('modal-body').innerHTML = `
    <p style="font-size:13px;color:var(--gray-600);margin-bottom:16px">
      Leave a field blank to keep existing values unchanged.
    </p>
    <div class="form-group"><label>Reassign Owner</label>
      <select id="bulk-owner"><option value="">— Keep existing —</option>${ownerOpts}</select></div>
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
  const spIds   = getSelectedTaskSpIds();
  const owner   = document.getElementById('bulk-owner').value;
  const shift   = parseInt(document.getElementById('bulk-shift').value);
  const status  = document.getElementById('bulk-status').value;
  if (!owner && isNaN(shift) && !status) { alert('Please set at least one field to change.'); return; }
  closeAllModals();
  showLoadingOverlay(true);
  try {
    for (const spId of spIds) {
      const task   = _tasks.find(t => t._spId === spId); if(!task) continue;
      const fields = {};
      if (owner)         { fields.OwnerId = owner;            task.ownerId = owner; }
      if (!isNaN(shift) && shift !== 0) {
        const newDate    = addDays(task.dueDate, shift);
        fields.DueDate   = newDate; task.dueDate = newDate;
      }
      if (status) {
        const prev       = task.status;
        fields.Status    = status; task.status = status;
        await writeSignOff(spId, 'task', task.name, prev, status);
      }
      if (Object.keys(fields).length) await updateListItem(LISTS.tasks, spId, fields);
    }
    await refreshData();
  } catch(e) { showError('Bulk edit failed: ' + e.message); }
  finally { showLoadingOverlay(false); }
}
