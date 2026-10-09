/**
 * ============================================================================
 * OFFICE LEAVE CALENDAR — JSON API BACKEND
 * ----------------------------------------------------------------------------
 * This Apps Script exposes a JSON API over HTTP.
 * The frontend (index.html) lives separately and connects to this API
 * using the Web App URL you paste into the tool once.
 *
 * Sheet Tabs (auto-created):
 *   Staff       — Name | Role | Joining Date | Active
 *   Leaves      — ID | Date | Staff Name | Type | Note | Created At
 *   Holidays    — Date | Name
 *   Settings    — Key | Value
 *
 * Deployment:
 *   Deploy -> New Deployment -> Web app
 *     • Execute as: Me
 *     • Who has access: Anyone   ← IMPORTANT for cross-origin fetch
 *   Copy the Web App URL. Paste it into index.html on first launch.
 * ============================================================================
 */

const SHEET_STAFF        = 'Staff';
const SHEET_LEAVES       = 'Leaves';
const SHEET_HOLIDAYS     = 'Holidays';
const SHEET_SETTINGS     = 'Settings';
const SHEET_TEMPLATES    = 'Templates';
const SHEET_COMPENSATION = 'Compensation';

const DEFAULT_SETTINGS = {
  CYCLE_START_DAY:      '26',
  OFFICE_NAME:          'My Office',
  WEEKLY_OFF:           'Sunday',
  ANNUAL_LEAVE_QUOTA:   '12',
  QUOTA_YEAR_START:     '01-10',  // MM-DD — quota renews on this day each year
  FREEZE_OVERRIDE:      'false',  // 'true' = allow editing past (paid) cycles
  OPENING_BALANCE_ASOF: ''        // ISO date the starting balances were counted through
};

/* ────────────────────────────────────────────────────────────────────────── */
/* HTTP ENTRY POINTS                                                          */
/* ────────────────────────────────────────────────────────────────────────── */
function doGet(e)  { return handle_(e && e.parameter ? e.parameter : {}); }
function doPost(e) {
  const params = (e && e.parameter) ? Object.assign({}, e.parameter) : {};
  if (e && e.postData && e.postData.contents) {
    try { Object.assign(params, JSON.parse(e.postData.contents)); } catch (err) {}
  }
  return handle_(params);
}

function handle_(p) {
  try { ensureSheetsExist(); } catch (err) { return json_({ ok:false, error:'Sheet init failed: ' + err.message }); }

  const action = String(p.action || 'ping').toLowerCase();

  // ── PIN GATE ───────────────────────────────────────────────────────────────
  // If an ACCESS_PIN is configured, every action except these must present the
  // matching pin hash. This protects the raw data, not just the UI.
  const OPEN_ACTIONS = { ping:1, pinstatus:1, verifypin:1 };
  const storedPin = String(getSettings().ACCESS_PIN || '');
  if (storedPin && !OPEN_ACTIONS[action]) {
    if (String(p.pin || '') !== storedPin) {
      return json_({ ok:false, error:'Locked — PIN required', code:'PIN_REQUIRED' });
    }
  }

  let result;
  try {
    switch (action) {
      case 'ping':          result = pingResponse_(); break;
      case 'pinstatus':     result = { ok:true, hasPin: !!storedPin }; break;
      case 'verifypin':     result = verifyPin_(p.pin); break;
      case 'setpin':        result = setPin_(p.newPin); break;
      case 'init':          result = getInitData(); break;
      case 'boot':          result = getBootData_(); break;
      case 'month':         result = getMonthData(+p.year, +p.month); break;
      case 'year':          result = getYearData(+p.year); break;
      case 'staffdetail':   result = getStaffDetail(p.name, +p.year); break;
      case 'addleave':      result = addLeave(p.date, p.staff, p.type, p.note, p.category, p.compId); break;
      case 'addleaves':     result = addLeaves(p.leaves); break;
      case 'deleteleave':   result = deleteLeave(p.id); break;
      case 'deleteleaves':  result = deleteLeaves(p.ids); break;
      case 'updateleave':   result = updateLeave(p.id, p.type, p.note); break;
      case 'compensation':  result = getCompensation_(); break;
      case 'addcomp':       result = addComp_(p.staff, p.type, p.date, p.note); break;
      case 'deletecomp':    result = deleteComp_(p.id); break;
      case 'staff':         result = getStaffList(); break;
      case 'addstaff':      result = addStaff(p.name, p.role); break;
      case 'updatestaff':   result = updateStaff(p.name, p.newName, p.role); break;
      case 'deactivatestaff': result = deactivateStaff(p.name); break;
      case 'reactivatestaff': result = reactivateStaff(p.name); break;
      case 'deletestaff':   result = deleteStaff(p.name, p.deleteLeaves); break;
      case 'setopening':    result = setOpeningBalance_(p.name, p.value); break;
      case 'setopenings':   result = setOpeningBalances_(p.balances); break;
      case 'holidays':      result = getAllHolidays(); break;
      case 'addholiday':    result = addHoliday(p.date, p.name, p.officeClosed); break;
      case 'addholidays':   result = addHolidays(p.holidays); break;
      case 'deleteholiday': result = deleteHoliday(p.date); break;
      case 'setholidayclosed': result = setHolidayClosed_(p.date, p.closed); break;
      case 'smartparse':    result = smartParse_(p.text); break;
      case 'setapikey':     result = setApiKey_(p.key); break;
      case 'aistatus':      result = aiStatus_(); break;
      case 'templates':     result = getTemplates(); break;
      case 'savetemplate':  result = saveTemplate(p.id, p.name, p.staff, p.days, p.type, p.note); break;
      case 'deletetemplate':result = deleteTemplate(p.id); break;
      case 'settings':      result = getPublicSettings_(); break;
      case 'updatesetting': result = updateSetting(p.key, p.value); break;
      case 'report':        result = getReport(+p.year, +p.month); break;
      case 'sheeturl':      result = { url: SpreadsheetApp.getActive().getUrl() }; break;
      default:              result = { ok:false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { ok:false, error: err.message || String(err) };
  }
  return json_(result);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function pingResponse_() {
  const settings = getSettings();
  return {
    ok: true,
    pong: new Date().toISOString(),
    office: settings.OFFICE_NAME || 'Office',
    hasPin: !!String(settings.ACCESS_PIN || ''),   // lets the client skip a separate pinstatus call
    version: '1.1'
  };
}

/* One-shot launch payload: staff, settings, today's month, compensation, freeze cutoff.
   Collapses 3 sequential calls (init + month + compensation) into one round-trip. */
function getBootData_() {
  ensureSheetsExist();
  const now = new Date();
  const y = now.getFullYear(), m = now.getMonth() + 1;
  return {
    ok: true,
    staff: getStaffList(),
    settings: getPublicSettings_(),
    todayISO: fmtDate_(now),
    freezeCutoffISO: getFreezeCutoffISO_(),
    currentMonth: { year: y, month: m, data: getMonthData(y, m) },
    compensation: getCompensation_(),
    holidaysAll: getAllHolidays()   // full list w/ officeClosed → powers the office-closed exclusion set
  };
}

/* ────────────────────────────────────────────────────────────────────────── */
/* SHEET INITIALIZATION                                                       */
/* ────────────────────────────────────────────────────────────────────────── */
function ensureSheetsExist() {
  const ss = SpreadsheetApp.getActive();
  const requiredTabs = [
    { name: SHEET_STAFF,        headers: ['Name', 'Role', 'Joining Date', 'Active', 'Opening Balance'] },
    { name: SHEET_LEAVES,       headers: ['ID', 'Date', 'Staff Name', 'Type', 'Note', 'Created At', 'Category', 'Comp ID'] },
    { name: SHEET_HOLIDAYS,     headers: ['Date', 'Name', 'Office Closed'] },
    { name: SHEET_SETTINGS,     headers: ['Key', 'Value'] },
    { name: SHEET_TEMPLATES,    headers: ['ID', 'Name', 'Staff', 'Days', 'Type', 'Note', 'Created At'] },
    { name: SHEET_COMPENSATION, headers: ['ID', 'Staff Name', 'Type', 'Earned Date', 'Reason', 'Status', 'Used Leave ID', 'Used Date', 'Created At'] }
  ];
  requiredTabs.forEach(t => {
    let sh = ss.getSheetByName(t.name);
    if (!sh) {
      sh = ss.insertSheet(t.name);
      sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers]).setFontWeight('bold').setBackground('#f0f0f0');
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, t.headers.length);
    }
  });
  // MIGRATION helper: append a missing header column to an existing sheet
  const addColumnIfMissing_ = (sheetName, header) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastColumn() < 1) return;
    const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (hdr.indexOf(header) < 0) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(header).setFontWeight('bold').setBackground('#f0f0f0');
    }
  };
  addColumnIfMissing_(SHEET_STAFF,    'Opening Balance');
  addColumnIfMissing_(SHEET_LEAVES,   'Category');   // 'Regular' | 'OT'
  addColumnIfMissing_(SHEET_LEAVES,   'Comp ID');    // links an OT leave to the comp credit it consumed
  addColumnIfMissing_(SHEET_HOLIDAYS, 'Office Closed'); // TRUE = office shut → leaves that day don't count
  const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  const data = settingsSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Object.entries(DEFAULT_SETTINGS).forEach(([k, v]) => settingsSheet.appendRow([k, v]));
  }
  return { ok: true };
}

function initializeSheets() { return ensureSheetsExist(); }  // Manual runner

/* ────────────────────────────────────────────────────────────────────────── */
/* HELPERS                                                                    */
/* ────────────────────────────────────────────────────────────────────────── */
function sh_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }

function readRows_(sheetName) {
  const s = sh_(sheetName);
  const data = s.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function fmtDate_(d) {
  if (!d) return '';
  const date = (d instanceof Date) ? d : new Date(d);
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function parseDate_(str) {
  const [y, m, d] = String(str).split('-').map(Number);
  return new Date(y, m - 1, d);
}

function newId_() {
  return 'L' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
}

/* ────────────────────────────────────────────────────────────────────────── */
/* API IMPLEMENTATIONS                                                        */
/* ────────────────────────────────────────────────────────────────────────── */
function getInitData() {
  ensureSheetsExist();
  return {
    ok: true,
    staff: getStaffList(),
    settings: getPublicSettings_(),
    todayISO: fmtDate_(new Date())
  };
}

function getSettings() {
  const rows = readRows_(SHEET_SETTINGS);
  const out = Object.assign({}, DEFAULT_SETTINGS);
  rows.forEach(r => { if (r.Key) out[r.Key] = String(r.Value); });
  return out;
}

/* Settings safe to send to the client — never expose the PIN hash. */
function getPublicSettings_() {
  const s = getSettings();
  delete s.ACCESS_PIN;
  return s;
}

/* Verify a provided PIN hash against the stored one. Open action (no gate).
   On success it also returns the full boot payload, so unlocking loads the app in ONE round-trip. */
function verifyPin_(pinHash) {
  const stored = String(getSettings().ACCESS_PIN || '');
  if (!stored) { const b = getBootData_(); b.valid = true; b.hasPin = false; return b; }
  const valid = (String(pinHash || '') === stored);
  if (!valid) return { ok:true, valid:false, hasPin:true };
  const b = getBootData_(); b.valid = true; b.hasPin = true; return b;
}

/* Set, change, or remove the access PIN.
   newPin is the SHA-256 hash from the client, or '' to remove protection.
   When a PIN already exists this action is gated (caller must present the current PIN). */
function setPin_(newPin) {
  updateSetting('ACCESS_PIN', String(newPin || ''));
  return { ok:true, hasPin: !!String(newPin || '') };
}

function updateSetting(key, value) {
  if (!key) return { ok:false, error:'Key required' };
  const s = sh_(SHEET_SETTINGS);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      s.getRange(i + 1, 2).setValue(value);
      return { ok: true };
    }
  }
  s.appendRow([key, value]);
  return { ok: true };
}

function getStaffList() {
  return readRows_(SHEET_STAFF)
    .filter(r => r.Name)
    .map(r => ({
      name: String(r.Name).trim(),
      role: r.Role || '',
      joiningDate: r['Joining Date'] ? fmtDate_(r['Joining Date']) : '',
      active: (r.Active === '' || r.Active === true || String(r.Active).toLowerCase() === 'true' || String(r.Active).toLowerCase() === 'yes'),
      openingBalance: parseFloat(r['Opening Balance']) || 0
    }));
}

function addStaff(name, role) {
  if (!name || !String(name).trim()) return { ok:false, error:'Name required' };
  const clean = String(name).trim();
  const existing = getStaffList().find(s => s.name.toLowerCase() === clean.toLowerCase());
  if (existing) return { ok:true, existed:true, name: existing.name };
  sh_(SHEET_STAFF).appendRow([clean, role || '', new Date(), true, 0]);
  return { ok:true, name:clean };
}

/* Set a staff's opening balance (leaves already taken this quota year before tool tracking).
   Stored in the Staff sheet's "Opening Balance" column. */
function setOpeningBalance_(name, value) {
  if (!name || !String(name).trim()) return { ok:false, error:'Name required' };
  const num = parseFloat(value);
  const bal = isNaN(num) ? 0 : num;
  const s = sh_(SHEET_STAFF);
  const data = s.getDataRange().getValues();
  const hdr = data[0].map(String);
  let col = hdr.indexOf('Opening Balance');
  if (col < 0) { col = hdr.length; s.getRange(1, col + 1).setValue('Opening Balance'); }
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(name).trim().toLowerCase()) {
      s.getRange(i + 1, col + 1).setValue(bal);
      return { ok:true, name:String(name).trim(), openingBalance:bal };
    }
  }
  return { ok:false, error:'Staff not found' };
}

/* Bulk set opening balances: balances = [{name, value}, ...] */
function setOpeningBalances_(balances) {
  if (typeof balances === 'string') { try { balances = JSON.parse(balances); } catch (e) { return { ok:false, error:'Invalid balances JSON' }; } }
  if (!Array.isArray(balances)) return { ok:false, error:'balances must be an array' };
  let updated = 0;
  balances.forEach(b => { const r = setOpeningBalance_(b.name, b.value); if (r.ok) updated++; });
  return { ok:true, updated };
}

function updateStaff(name, newName, role) {
  if (!name || !String(name).trim()) return { ok:false, error:'Name required' };
  const s = sh_(SHEET_STAFF);
  const data = s.getDataRange().getValues();
  const cleanNew = newName ? String(newName).trim() : String(name).trim();
  // If renaming, ensure no duplicate
  if (cleanNew.toLowerCase() !== String(name).trim().toLowerCase()) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === cleanNew.toLowerCase()) {
        return { ok:false, error:'Another staff already has this name' };
      }
    }
  }
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(name).trim().toLowerCase()) {
      s.getRange(i + 1, 1).setValue(cleanNew);
      s.getRange(i + 1, 2).setValue(role || '');
      // Cascade rename in Leaves sheet + Compensation sheet
      if (cleanNew !== String(name).trim()) {
        const oldLower = String(name).trim().toLowerCase();
        const lsh = sh_(SHEET_LEAVES);
        const ldata = lsh.getDataRange().getValues();
        for (let j = 1; j < ldata.length; j++) {
          if (String(ldata[j][2]).trim().toLowerCase() === oldLower) lsh.getRange(j + 1, 3).setValue(cleanNew);
        }
        const csh = sh_(SHEET_COMPENSATION);
        if (csh) {
          const cdata = csh.getDataRange().getValues();
          const cName = cdata[0].map(String).indexOf('Staff Name');
          if (cName >= 0) {
            for (let j = 1; j < cdata.length; j++) {
              if (String(cdata[j][cName]).trim().toLowerCase() === oldLower) csh.getRange(j + 1, cName + 1).setValue(cleanNew);
            }
          }
        }
      }
      return { ok:true, name:cleanNew };
    }
  }
  return { ok:false, error:'Staff not found' };
}

function deactivateStaff(name) {
  const s = sh_(SHEET_STAFF);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(name).trim().toLowerCase()) {
      s.getRange(i + 1, 4).setValue(false);
      return { ok:true };
    }
  }
  return { ok:false, error:'Staff not found' };
}

function reactivateStaff(name) {
  const s = sh_(SHEET_STAFF);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(name).trim().toLowerCase()) {
      s.getRange(i + 1, 4).setValue(true);
      return { ok:true };
    }
  }
  return { ok:false, error:'Staff not found' };
}

/**
 * Permanently delete a staff member. Optionally cascade-delete their leaves.
 * @param {string} name - Staff name to delete
 * @param {boolean} deleteLeaves - If true, also delete all leaves for this staff
 */
function deleteStaff(name, deleteLeaves) {
  if (!name || !String(name).trim()) return { ok:false, error:'Name required' };
  const cleanName = String(name).trim().toLowerCase();
  const s = sh_(SHEET_STAFF);
  const data = s.getDataRange().getValues();
  let removed = false;
  // Reverse iterate to safely delete rows
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim().toLowerCase() === cleanName) {
      s.deleteRow(i + 1);
      removed = true;
      break;
    }
  }
  if (!removed) return { ok:false, error:'Staff not found' };

  let leavesDeleted = 0;
  if (deleteLeaves === true || deleteLeaves === 'true') {
    const lsh = sh_(SHEET_LEAVES);
    const ldata = lsh.getDataRange().getValues();
    for (let j = ldata.length - 1; j >= 1; j--) {
      if (String(ldata[j][2]).trim().toLowerCase() === cleanName) {
        lsh.deleteRow(j + 1);
        leavesDeleted++;
      }
    }
  }
  return { ok:true, leavesDeleted };
}

/* Shared holiday mapper — includes the Office Closed flag. */
function mapHolidayRow_(r) {
  const oc = r['Office Closed'];
  return {
    date: fmtDate_(r.Date),
    name: String(r.Name || ''),
    officeClosed: (oc === true || String(oc).toLowerCase() === 'true' || String(oc).toLowerCase() === 'yes')
  };
}

/* Shared mapper so every endpoint returns leaves with the same shape (incl. OT category). */
function mapLeaveRow_(r) {
  return {
    id: r.ID,
    date: fmtDate_(r.Date),
    staff: String(r['Staff Name'] || '').trim(),
    type: r.Type || 'Full',
    note: r.Note || '',
    category: String(r.Category || 'Regular') || 'Regular',
    compId: r['Comp ID'] || ''
  };
}

function getMonthData(year, month) {
  if (!year || !month) return { ok:false, error:'year and month required' };
  const leaves = readRows_(SHEET_LEAVES).map(mapLeaveRow_).filter(l => {
    if (!l.date) return false;
    const [y, m] = l.date.split('-').map(Number);
    return y === year && m === month;
  });

  const holidays = readRows_(SHEET_HOLIDAYS).map(mapHolidayRow_).filter(h => {
    if (!h.date) return false;
    const [y, m] = h.date.split('-').map(Number);
    return y === year && m === month;
  });

  return { ok:true, leaves, holidays };
}

/**
 * Return all leaves & holidays for an entire year in one call.
 * Powers the year-heatmap view without needing 12 separate month requests.
 */
function getYearData(year) {
  if (!year) return { ok:false, error:'year required' };
  const startISO = year + '-01-01';
  const endISO = year + '-12-31';
  const leaves = readRows_(SHEET_LEAVES).map(mapLeaveRow_).filter(l => l.date >= startISO && l.date <= endISO);
  const holidays = readRows_(SHEET_HOLIDAYS).map(mapHolidayRow_).filter(h => h.date >= startISO && h.date <= endISO);
  return { ok:true, year:year, leaves:leaves, holidays:holidays };
}

/**
 * Complete history + patterns for a single staff member across a year window.
 */
function getStaffDetail(name, year) {
  if (!name) return { ok:false, error:'name required' };
  const y = year || new Date().getFullYear();
  const startISO = y + '-01-01';
  const endISO = y + '-12-31';
  const clean = String(name).trim().toLowerCase();
  const leaves = readRows_(SHEET_LEAVES).map(r => ({
    id: r.ID,
    date: fmtDate_(r.Date),
    staff: String(r['Staff Name'] || '').trim(),
    type: r.Type || 'Full',
    note: r.Note || ''
  })).filter(l => l.staff.toLowerCase() === clean && l.date >= startISO && l.date <= endISO)
    .sort((a,b) => a.date.localeCompare(b.date));

  // Aggregations
  const monthly = {}; // '2026-09' -> {full, half, total}
  const dowCount = [0,0,0,0,0,0,0]; // day-of-week counter
  let fullDays = 0, halfDays = 0;
  leaves.forEach(l => {
    const mk = l.date.substring(0,7);
    monthly[mk] = monthly[mk] || { full:0, half:0, total:0 };
    if (l.type === 'Half') { monthly[mk].half++; halfDays++; monthly[mk].total += 0.5; }
    else                   { monthly[mk].full++; fullDays++; monthly[mk].total += 1; }
    const dt = new Date(l.date + 'T00:00:00');
    dowCount[dt.getDay()]++;
  });
  const totalDays = fullDays + halfDays * 0.5;

  // Most-common day-of-week
  let topDow = -1, topDowCount = 0;
  dowCount.forEach((c, i) => { if (c > topDowCount) { topDowCount = c; topDow = i; } });

  return {
    ok: true,
    name: String(name).trim(),
    year: y,
    leaves: leaves,
    fullDays: fullDays,
    halfDays: halfDays,
    totalDays: totalDays,
    monthly: monthly,
    dowCount: dowCount,
    topDow: topDow,
    topDowCount: topDowCount
  };
}

/* ────────────────────────────────────────────────────────────────────────── */
/* FREEZE (lock past / already-paid cycles)                                   */
/* ────────────────────────────────────────────────────────────────────────── */
/* Cutoff = day before the current salary cycle's start. Everything on/before it
   belongs to a cycle whose salary is already paid, so it's locked.
   Returns '' when FREEZE_OVERRIDE is on (nothing locked). */
function getFreezeCutoffISO_() {
  const settings = getSettings();
  if (String(settings.FREEZE_OVERRIDE || 'false').toLowerCase() === 'true') return '';
  const startDay = parseInt(settings.CYCLE_START_DAY, 10) || 1;
  const today = new Date();
  const y = today.getFullYear(), m = today.getMonth(), d = today.getDate();
  let curStart;
  if (startDay === 1) {
    curStart = new Date(y, m, 1);
  } else {
    curStart = (d >= startDay) ? new Date(y, m, startDay) : new Date(y, m - 1, startDay);
  }
  const cutoff = new Date(curStart.getFullYear(), curStart.getMonth(), curStart.getDate() - 1);
  return fmtDate_(cutoff);
}
function isFrozen_(dateISO) {
  const cutoff = getFreezeCutoffISO_();
  return !!cutoff && String(dateISO) <= cutoff;
}

function addLeave(dateISO, staffName, type, note, category, compId) {
  if (!dateISO || !staffName) return { ok:false, error:'Date and staff required' };
  if (isFrozen_(dateISO)) return { ok:false, error:'That date is in a locked (already-paid) cycle', code:'FROZEN' };
  const clean = String(staffName).trim();
  const t = (String(type) === 'Half') ? 'Half' : 'Full';
  const cat = (String(category) === 'OT') ? 'OT' : 'Regular';
  addStaff(clean, '');
  const existing = readRows_(SHEET_LEAVES).find(r =>
    fmtDate_(r.Date) === dateISO &&
    String(r['Staff Name']).trim().toLowerCase() === clean.toLowerCase() &&
    r.Type === t
  );
  if (existing) return { ok:false, error:'Duplicate leave already exists' };

  // If this is an OT (comp) leave, consume a matching Available comp credit
  let usedCompId = '';
  if (cat === 'OT') {
    const credit = consumeCompCredit_(clean, t, compId, dateISO);
    if (!credit.ok) return { ok:false, error: credit.error || 'No matching compensation credit available' };
    usedCompId = credit.id;
  }

  const id = newId_();
  sh_(SHEET_LEAVES).appendRow([id, parseDate_(dateISO), clean, t, note || '', new Date(), cat, usedCompId]);
  // Link the consumed credit back to this leave id
  if (usedCompId) linkCompToLeave_(usedCompId, id, dateISO);
  return { ok:true, id, category:cat, compId:usedCompId };
}

/**
 * BULK: add multiple leaves in a single sheet write.
 * Solves the parallel-request race condition where individual addLeave calls
 * can be silently dropped by Apps Script's sheet lock contention.
 * Accepts: leaves = [{date, staff, type, note}, ...]
 * Returns: {ok, added:[{id,date,staff,type,note}], skipped:[{date,staff,type,error}]}
 */
function addLeaves(leaves) {
  // leaves may come in as an array (from POST body JSON merge) or a JSON string (from GET query param)
  if (typeof leaves === 'string') {
    try { leaves = JSON.parse(leaves); } catch (e) { return { ok:false, error:'Invalid leaves JSON' }; }
  }
  if (!Array.isArray(leaves) || !leaves.length) return { ok:false, error:'leaves must be a non-empty array' };

  const s = sh_(SHEET_LEAVES);
  // Snapshot existing keys once — O(n) instead of O(n²)
  const existingSet = {};
  readRows_(SHEET_LEAVES).forEach(r => {
    const key = fmtDate_(r.Date) + '|' + String(r['Staff Name']).trim().toLowerCase() + '|' + (r.Type || 'Full');
    existingSet[key] = true;
  });

  const added = [];
  const skipped = [];
  const rowsToAppend = [];
  const seenInBatch = {};
  const uniqueStaff = {};

  const freezeCutoff = getFreezeCutoffISO_();
  leaves.forEach(l => {
    if (!l.date || !l.staff) { skipped.push({ date:l.date, staff:l.staff, type:l.type, error:'Missing date or staff' }); return; }
    if (freezeCutoff && String(l.date) <= freezeCutoff) { skipped.push({ date:l.date, staff:l.staff, type:l.type, error:'Locked cycle' }); return; }
    const clean = String(l.staff).trim();
    const t = (String(l.type) === 'Half') ? 'Half' : 'Full';
    const key = String(l.date) + '|' + clean.toLowerCase() + '|' + t;
    if (existingSet[key] || seenInBatch[key]) {
      skipped.push({ date:l.date, staff:clean, type:t, error:'Duplicate' });
      return;
    }
    seenInBatch[key] = true;
    uniqueStaff[clean] = true;
    const id = newId_();
    // Bulk path is Regular leaves only (OT comp leaves go through addLeave individually)
    rowsToAppend.push([id, parseDate_(l.date), clean, t, l.note || '', new Date(), 'Regular', '']);
    added.push({ id:id, date:l.date, staff:clean, type:t, note:l.note || '', category:'Regular' });
  });

  // Auto-register any new staff (before appending leaves)
  Object.keys(uniqueStaff).forEach(n => addStaff(n, ''));

  // Single atomic sheet write — no race condition possible
  if (rowsToAppend.length) {
    const startRow = s.getLastRow() + 1;
    s.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  return { ok:true, added:added, skipped:skipped };
}

function deleteLeave(leaveId) {
  if (!leaveId) return { ok:false, error:'ID required' };
  const s = sh_(SHEET_LEAVES);
  const data = s.getDataRange().getValues();
  const hdr = data[0].map(String);
  const iDate = hdr.indexOf('Date'), iCat = hdr.indexOf('Category'), iComp = hdr.indexOf('Comp ID');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(leaveId)) {
      if (isFrozen_(fmtDate_(data[i][iDate]))) return { ok:false, error:'That date is in a locked (already-paid) cycle', code:'FROZEN' };
      // If this was an OT leave, release its comp credit back to Available
      if (iCat >= 0 && String(data[i][iCat]) === 'OT' && iComp >= 0 && data[i][iComp]) {
        releaseCompCredit_(String(data[i][iComp]));
      }
      s.deleteRow(i + 1);
      return { ok:true };
    }
  }
  return { ok:false, error:'Leave not found' };
}

/**
 * Bulk delete multiple leaves in a single sheet operation.
 * Skips frozen dates; releases comp credits for any OT leaves removed.
 */
function deleteLeaves(ids) {
  if (typeof ids === 'string') { try { ids = JSON.parse(ids); } catch (e) { return { ok:false, error:'Invalid ids JSON' }; } }
  if (!Array.isArray(ids) || !ids.length) return { ok:false, error:'ids must be a non-empty array' };
  const idSet = {};
  ids.forEach(id => { idSet[String(id)] = true; });
  const s = sh_(SHEET_LEAVES);
  const data = s.getDataRange().getValues();
  const hdr = data[0].map(String);
  const iDate = hdr.indexOf('Date'), iCat = hdr.indexOf('Category'), iComp = hdr.indexOf('Comp ID');
  const deleted = [];
  const skippedFrozen = [];
  const compsToRelease = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const rowId = String(data[i][0]);
    if (!idSet[rowId]) continue;
    if (isFrozen_(fmtDate_(data[i][iDate]))) { skippedFrozen.push(rowId); continue; }
    if (iCat >= 0 && String(data[i][iCat]) === 'OT' && iComp >= 0 && data[i][iComp]) compsToRelease.push(String(data[i][iComp]));
    s.deleteRow(i + 1);
    deleted.push(rowId);
  }
  compsToRelease.forEach(releaseCompCredit_);
  return { ok:true, deleted:deleted, count:deleted.length, skippedFrozen:skippedFrozen };
}

function updateLeave(leaveId, type, note) {
  const s = sh_(SHEET_LEAVES);
  const data = s.getDataRange().getValues();
  const headers = data[0].map(String);
  const idxType = headers.indexOf('Type');
  const idxNote = headers.indexOf('Note');
  const idxDate = headers.indexOf('Date');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(leaveId)) {
      if (isFrozen_(fmtDate_(data[i][idxDate]))) return { ok:false, error:'That date is in a locked (already-paid) cycle', code:'FROZEN' };
      if (type) s.getRange(i + 1, idxType + 1).setValue(type === 'Half' ? 'Half' : 'Full');
      if (note !== undefined) s.getRange(i + 1, idxNote + 1).setValue(note);
      return { ok:true };
    }
  }
  return { ok:false, error:'Leave not found' };
}

function addHoliday(dateISO, name, officeClosed) {
  if (!dateISO || !name) return { ok:false, error:'Date and name required' };
  const closed = (officeClosed === true || String(officeClosed).toLowerCase() === 'true');
  sh_(SHEET_HOLIDAYS).appendRow([parseDate_(dateISO), String(name).trim(), closed]);
  return { ok:true, officeClosed:closed };
}

function deleteHoliday(dateISO) {
  const s = sh_(SHEET_HOLIDAYS);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (fmtDate_(data[i][0]) === dateISO) {
      s.deleteRow(i + 1);
      return { ok:true };
    }
  }
  return { ok:false, error:'Holiday not found' };
}

/* ────────────────────────────────────────────────────────────────────────── */
/* COMPENSATION (overtime comp-off credits)                                   */
/* Each row = one credit the authority granted for overtime. When a staff takes */
/* an OT leave, a matching Available credit is consumed and linked to that leave. */
/* ────────────────────────────────────────────────────────────────────────── */
function getCompensation_() {
  return {
    ok: true,
    credits: readRows_(SHEET_COMPENSATION).filter(r => r.ID).map(r => ({
      id: r.ID,
      staff: String(r['Staff Name'] || '').trim(),
      type: (String(r.Type) === 'Half') ? 'Half' : 'Full',
      earnedDate: r['Earned Date'] ? fmtDate_(r['Earned Date']) : '',
      reason: r.Reason || '',
      status: (String(r.Status || 'Available') === 'Used') ? 'Used' : 'Available',
      usedLeaveId: r['Used Leave ID'] || '',
      usedDate: r['Used Date'] ? fmtDate_(r['Used Date']) : ''
    }))
  };
}

function addComp_(staff, type, dateISO, reason) {
  if (!staff || !String(staff).trim()) return { ok:false, error:'Staff required' };
  const clean = String(staff).trim();
  const t = (String(type) === 'Half') ? 'Half' : 'Full';
  addStaff(clean, '');
  const id = 'C' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
  sh_(SHEET_COMPENSATION).appendRow([id, clean, t, dateISO ? parseDate_(dateISO) : new Date(), reason || '', 'Available', '', '', new Date()]);
  return { ok:true, id, staff:clean, type:t };
}

function deleteComp_(id) {
  if (!id) return { ok:false, error:'ID required' };
  const s = sh_(SHEET_COMPENSATION);
  const data = s.getDataRange().getValues();
  const iStatus = data[0].map(String).indexOf('Status');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      if (String(data[i][iStatus]) === 'Used') return { ok:false, error:'This credit is already used — delete its (OT) leave first to free it' };
      s.deleteRow(i + 1);
      return { ok:true };
    }
  }
  return { ok:false, error:'Not found' };
}

/* Reserve a matching Available credit (marks it Used). Prefers an explicit id, else oldest. */
function consumeCompCredit_(staff, type, preferredId, dateISO) {
  const s = sh_(SHEET_COMPENSATION);
  const data = s.getDataRange().getValues();
  const hdr = data[0].map(String);
  const iStaff = hdr.indexOf('Staff Name'), iType = hdr.indexOf('Type'), iStatus = hdr.indexOf('Status'), iEarned = hdr.indexOf('Earned Date');
  const lower = String(staff).trim().toLowerCase();
  let target = -1;
  if (preferredId) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(preferredId) && String(data[i][iStatus] || 'Available') !== 'Used' &&
          String(data[i][iStaff]).trim().toLowerCase() === lower && String(data[i][iType]) === type) { target = i; break; }
    }
  }
  if (target < 0) {
    let oldest = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][iStaff]).trim().toLowerCase() === lower && String(data[i][iType]) === type && String(data[i][iStatus] || 'Available') !== 'Used') {
        const ed = data[i][iEarned] ? new Date(data[i][iEarned]).getTime() : 0;
        if (oldest === null || ed < oldest.ed) oldest = { row:i, ed };
      }
    }
    if (oldest) target = oldest.row;
  }
  if (target < 0) return { ok:false, error:'No available ' + type + ' comp credit for ' + staff };
  s.getRange(target + 1, iStatus + 1).setValue('Used');
  return { ok:true, id:String(data[target][0]) };
}

function linkCompToLeave_(compId, leaveId, dateISO) {
  const s = sh_(SHEET_COMPENSATION);
  const data = s.getDataRange().getValues();
  const hdr = data[0].map(String);
  const iUsedLeave = hdr.indexOf('Used Leave ID'), iUsedDate = hdr.indexOf('Used Date');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(compId)) {
      if (iUsedLeave >= 0) s.getRange(i + 1, iUsedLeave + 1).setValue(leaveId);
      if (iUsedDate >= 0) s.getRange(i + 1, iUsedDate + 1).setValue(dateISO ? parseDate_(dateISO) : new Date());
      return { ok:true };
    }
  }
  return { ok:false };
}

function releaseCompCredit_(compId) {
  const s = sh_(SHEET_COMPENSATION);
  const data = s.getDataRange().getValues();
  const hdr = data[0].map(String);
  const iStatus = hdr.indexOf('Status'), iUsedLeave = hdr.indexOf('Used Leave ID'), iUsedDate = hdr.indexOf('Used Date');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(compId)) {
      if (iStatus >= 0) s.getRange(i + 1, iStatus + 1).setValue('Available');
      if (iUsedLeave >= 0) s.getRange(i + 1, iUsedLeave + 1).setValue('');
      if (iUsedDate >= 0) s.getRange(i + 1, iUsedDate + 1).setValue('');
      return { ok:true };
    }
  }
  return { ok:false };
}

/**
 * Bulk add holidays (used by "Import Indian Holidays" button).
 * Skips duplicates by (date + name).
 */
function addHolidays(list) {
  if (typeof list === 'string') { try { list = JSON.parse(list); } catch (e) { return { ok:false, error:'Invalid list JSON' }; } }
  if (!Array.isArray(list) || !list.length) return { ok:false, error:'list must be a non-empty array' };
  const s = sh_(SHEET_HOLIDAYS);
  const existing = {};
  readRows_(SHEET_HOLIDAYS).forEach(r => {
    existing[fmtDate_(r.Date) + '|' + String(r.Name || '').toLowerCase()] = true;
  });
  const added = [];
  const skipped = [];
  const rowsToAppend = [];
  list.forEach(h => {
    if (!h.date || !h.name) { skipped.push({ ...h, error:'Missing date/name' }); return; }
    const key = h.date + '|' + String(h.name).toLowerCase();
    if (existing[key]) { skipped.push({ ...h, error:'Already exists' }); return; }
    existing[key] = true;
    const closed = (h.officeClosed === true || String(h.officeClosed).toLowerCase() === 'true');
    rowsToAppend.push([parseDate_(h.date), String(h.name).trim(), closed]);
    added.push({ date:h.date, name:h.name, officeClosed:closed });
  });
  if (rowsToAppend.length) {
    const startRow = s.getLastRow() + 1;
    s.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }
  return { ok:true, added:added, skipped:skipped };
}

function getAllHolidays() {
  return readRows_(SHEET_HOLIDAYS).map(mapHolidayRow_).filter(h => h.date);
}

/* Set a holiday's Office Closed flag (by date). */
function setHolidayClosed_(dateISO, closed) {
  if (!dateISO) return { ok:false, error:'Date required' };
  const s = sh_(SHEET_HOLIDAYS);
  const data = s.getDataRange().getValues();
  let col = data[0].map(String).indexOf('Office Closed');
  if (col < 0) { col = data[0].length; s.getRange(1, col + 1).setValue('Office Closed'); }
  const want = (closed === true || String(closed).toLowerCase() === 'true');
  for (let i = 1; i < data.length; i++) {
    if (fmtDate_(data[i][0]) === dateISO) { s.getRange(i + 1, col + 1).setValue(want); return { ok:true, date:dateISO, officeClosed:want }; }
  }
  return { ok:false, error:'Holiday not found' };
}

/**
 * Salary-cycle report.
 * CYCLE_START_DAY = 1  → calendar month
 * CYCLE_START_DAY = 26 → 26th of previous month through 25th of current month
 */
function getReport(year, month) {
  if (!year || !month) return { ok:false, error:'year and month required' };
  const settings = getSettings();
  const startDay = parseInt(settings.CYCLE_START_DAY, 10) || 1;
  const weeklyOffs = String(settings.WEEKLY_OFF || 'Sunday').split(',').map(function(s){ return s.trim().toLowerCase(); });

  let startDate, endDate;
  if (startDay === 1) {
    startDate = new Date(year, month - 1, 1);
    endDate   = new Date(year, month, 0);
  } else {
    startDate = new Date(year, month - 2, startDay);
    endDate   = new Date(year, month - 1, startDay - 1);
  }
  const startISO = fmtDate_(startDate);
  const endISO   = fmtDate_(endDate);

  const DAY_NAMES = ['sunday','monday','tuesday','wednesday','thursday','friday','saturday'];
  function isWeekOffISO_(iso) {
    const d = new Date(iso + 'T00:00:00');
    return weeklyOffs.indexOf(DAY_NAMES[d.getDay()]) >= 0;
  }
  // Office-closed holiday dates — leaves on these don't count (office was shut)
  const officeClosedSet = {};
  getAllHolidays().forEach(function(h){ if (h.officeClosed) officeClosedSet[h.date] = true; });
  function isNonWorking_(iso) { return isWeekOffISO_(iso) || !!officeClosedSet[iso]; }

  const rawLeaves = readRows_(SHEET_LEAVES).map(mapLeaveRow_)
    .filter(function(l){ return l.date >= startISO && l.date <= endISO; })
    .filter(function(l){ return l.category !== 'OT'; });   // OT (comp) leaves never count in salary

  // Skip non-working-day leaves (weekly off + office-closed holidays) — office was already shut
  const excluded = rawLeaves.filter(function(l){ return isNonWorking_(l.date); });
  const leaves = rawLeaves.filter(function(l){ return !isNonWorking_(l.date); });

  const totals = {};
  leaves.forEach(function(l){
    if (!totals[l.staff]) totals[l.staff] = { name: l.staff, fullDays: 0, halfDays: 0, totalDays: 0, leaves: [] };
    if (l.type === 'Half') { totals[l.staff].halfDays += 1; totals[l.staff].totalDays += 0.5; }
    else                   { totals[l.staff].fullDays += 1; totals[l.staff].totalDays += 1; }
    totals[l.staff].leaves.push(l);
  });
  getStaffList().filter(function(s){ return s.active; }).forEach(function(s){
    if (!totals[s.name]) totals[s.name] = { name: s.name, fullDays: 0, halfDays: 0, totalDays: 0, leaves: [] };
  });

  const rows = Object.values(totals).sort(function(a, b){ return b.totalDays - a.totalDays || a.name.localeCompare(b.name); });

  return {
    ok: true,
    cycleLabel: monthLabel_(year, month),
    startDate: startISO,
    endDate:   endISO,
    startDay:  startDay,
    rows: rows,
    totalLeaves: leaves.length,
    weekOffLeavesExcluded: excluded.length
  };
}

/* ────────────────────────────────────────────────────────────────────────── */
/* TEMPLATES (weekly recurrence)                                              */
/* ────────────────────────────────────────────────────────────────────────── */
/**
 * A template = a per-staff, per-weekday leave rule
 * Days is stored as comma-separated day names ("Mon,Wed,Fri") or single ("Sun")
 * On apply, the frontend generates the actual leave records via addLeaves.
 */
function getTemplates() {
  return readRows_(SHEET_TEMPLATES).map(r => ({
    id: r.ID,
    name: String(r.Name || '').trim(),
    staff: String(r.Staff || '').trim(),
    days: String(r.Days || '').split(',').map(d => d.trim()).filter(Boolean),
    type: r.Type || 'Full',
    note: r.Note || ''
  })).filter(t => t.id);
}

function saveTemplate(id, name, staff, days, type, note) {
  if (!staff) return { ok:false, error:'Staff required' };
  if (!days) return { ok:false, error:'At least one day required' };
  const cleanDays = String(days).split(',').map(d => d.trim()).filter(Boolean).join(',');
  const t = (String(type) === 'Half') ? 'Half' : 'Full';
  const s = sh_(SHEET_TEMPLATES);
  const rows = s.getDataRange().getValues();

  if (id) {
    // Update existing
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        s.getRange(i+1, 1, 1, 7).setValues([[id, name || '', staff, cleanDays, t, note || '', new Date()]]);
        return { ok:true, id:id };
      }
    }
  }
  // Create new
  const newId = 'T' + Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
  s.appendRow([newId, name || (staff + ' - ' + cleanDays), staff, cleanDays, t, note || '', new Date()]);
  return { ok:true, id:newId };
}

function deleteTemplate(id) {
  if (!id) return { ok:false, error:'ID required' };
  const s = sh_(SHEET_TEMPLATES);
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) { s.deleteRow(i+1); return { ok:true }; }
  }
  return { ok:false, error:'Not found' };
}

/* ────────────────────────────────────────────────────────────────────────── */
/* AI SMART VOICE — parse a spoken command into structured JSON via Gemini      */
/* Key is stored in Script Properties (GEMINI_API_KEY), never returned to client */
/* ────────────────────────────────────────────────────────────────────────── */
var GEMINI_MODEL = 'gemini-2.0-flash';   // change here if you prefer another free model

function setApiKey_(key) {
  const clean = String(key || '').trim();
  const props = PropertiesService.getScriptProperties();
  if (clean) props.setProperty('GEMINI_API_KEY', clean);
  else props.deleteProperty('GEMINI_API_KEY');
  return { ok:true, hasKey: !!clean };
}
function aiStatus_() {
  const has = !!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return { ok:true, hasKey: has, model: GEMINI_MODEL };
}

function smartParse_(text) {
  if (!text || !String(text).trim()) return { ok:false, error:'No text' };
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) return { ok:false, error:'AI key not set', code:'NO_AI_KEY' };

  const staff = getStaffList().filter(function(s){ return s.active; }).map(function(s){ return s.name; });
  const settings = getSettings();
  const today = new Date();
  const todayISO = fmtDate_(today);
  const weekday = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][today.getDay()];

  const prompt = [
    'You convert a spoken office leave-management command into STRICT JSON. Output JSON only, no prose.',
    'Today is ' + todayISO + ' (' + weekday + '). Salary cycle start day: ' + (settings.CYCLE_START_DAY || '1') + '.',
    'Active staff — match spoken names to the EXACT name from this list: ' + JSON.stringify(staff) + '.',
    'Return this shape: {"intent":"add|remove|comp_add|comp_remove|ot_leave|office_closed|office_open","staff":[exact names],"startDate":"YYYY-MM-DD","endDate":"YYYY-MM-DD or null","type":"Full|Half","reason":"","holidayName":""}',
    'Meanings: add=add leave; remove=delete a leave; comp_add=grant an overtime comp-off credit; comp_remove=remove a comp credit; ot_leave=mark a leave that uses a comp credit (free); office_closed / office_open=set a day office closed/open.',
    'Resolve every date to an absolute YYYY-MM-DD from today. Handle "today, tomorrow, next monday, 5th, 5 to 8, 28 oct to 5 nov, till 10th". Use the current or nearest sensible occurrence; default the year to the current year. If a single day, endDate=null.',
    'Match misheard/mis-spelled names to the closest active staff. type defaults to "Full" unless half/aadha/aadhi is said. reason=overtime or holiday reason if any. holidayName only for office_closed.',
    'Command: "' + String(text).replace(/"/g, "'") + '"'
  ].join('\n');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(key);
  const payload = { contents:[{ parts:[{ text: prompt }] }], generationConfig:{ temperature:0, responseMimeType:'application/json' } };
  try {
    const res = UrlFetchApp.fetch(url, { method:'post', contentType:'application/json', payload: JSON.stringify(payload), muteHttpExceptions:true });
    const code = res.getResponseCode();
    const bodyTxt = res.getContentText();
    if (code !== 200) return { ok:false, error:'AI HTTP ' + code + ': ' + String(bodyTxt).slice(0, 200) };
    const data = JSON.parse(bodyTxt);
    const txt = data && data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0].text;
    if (!txt) return { ok:false, error:'AI returned no content' };
    var parsed;
    try { parsed = JSON.parse(txt); }
    catch (e) { const mm = String(txt).match(/\{[\s\S]*\}/); if (mm) parsed = JSON.parse(mm[0]); else return { ok:false, error:'AI non-JSON' }; }
    return { ok:true, parsed: parsed };
  } catch (err) { return { ok:false, error:'AI call failed: ' + (err.message || err) }; }
}

function monthLabel_(year, month) {
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return months[month - 1] + ' ' + year;
}
