/**
 * PPR Calendar Sync + Deadline Notifier  (v4)
 * Toronto Metropolitan University — YSGPS
 *
 * The flat "Stages" tab now includes a "Phase" column (derived from
 * the master's PHASE I/II/III header rows). Phase is used by AppSheet for grouping
 * but is intentionally NOT shown on the Google Calendar.
 *
 * FLOW
 *   buildStagesTable()  — master matrix -> flat "Stages" tab (one row per program × stage)
 *   syncPPRtoCalendar() — "Stages" tab -> "PPR Status Tracker" calendar
 *   sendPPRDigest()     — emails overdue / due-soon / starting-soon
 *   rebuildAndSync()    — both build + sync (use the "PPR Tracker" menu)
 *
 * Statuses (case-insensitive): Complete, In progress, Incomplete, N/A.
 *   N/A and blank -> not tracked (no event, no digest).
 */

// ===================== CONFIG =====================
const CONFIG = {
  MASTER_SHEET: 'PPR Tracking (Master)',
  SHEET_NAME: 'Stages',
  CALENDAR_NAME: 'PPR Status Tracker',
  SYNC_SHEET: '_PPR_SYNC',

  DUE_SOON_DAYS: 7,
  START_SOON_DAYS: 3,
  DIGEST_RECIPIENTS: '',
  DIGEST_HOUR: 7,
  SYNC_HOUR: 1,

  EVENT_REMINDER_MINUTES: [7 * 24 * 60, 24 * 60],
  RUN_BUDGET_MS: 5 * 60 * 1000,

  HEADERS: {
    'Program': 'program', 'Faculty': 'faculty', 'Degree': 'degree', 'Cycle': 'cycle',
    'Phase': 'phase', 'Stage No': 'stageNum', 'Stage': 'stageName',
    'Status': 'status', 'Start': 'start', 'Complete': 'complete', 'Key': 'key'
  }
};

const STAGE_HEADERS = ['Program', 'Faculty', 'Degree', 'Cycle', 'Phase', 'Stage No',
                       'Stage', 'Status', 'Start', 'Complete', 'Duration (days)', 'Key'];

const PHASE_MAP = { 'I': 'Phase 1: Self-Study', 'II': 'Phase 2: Site Visit', 'III': 'Phase 3: PRT Report' };

function normStatus_(raw) {
  const s = String(raw || '').trim().toLowerCase();
  if (s === 'complete') return 'Complete';
  if (s === 'in progress' || s === 'in-progress' || s === 'inprogress') return 'In progress';
  if (s === 'incomplete') return 'Incomplete';
  return null;
}

function statusStyle_(status) {
  switch (status) {
    case 'Complete':    return { emoji: '✅', color: CalendarApp.EventColor.GRAY };
    case 'In progress': return { emoji: '🔄', color: CalendarApp.EventColor.BLUE };
    case 'Incomplete':  return { emoji: '⬜', color: null };
    default:            return null;
  }
}

// ===================== MENU =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('PPR Tracker')
    .addItem('Rebuild + sync now', 'rebuildAndSync')
    .addSeparator()
    .addItem('Rebuild Stages from master', 'buildStagesTable')
    .addItem('Sync to calendar only', 'syncPPRtoCalendar')
    .addItem('Send deadline digest now', 'sendPPRDigest')
    .addSeparator()
    .addItem('Install daily triggers', 'installTriggers')
    .addToUi();
}

function rebuildAndSync() { buildStagesTable(); syncPPRtoCalendar(); }

// ===================== FLATTEN MASTER -> STAGES =====================
function buildStagesTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.MASTER_SHEET);
  if (!master) throw new Error('Master sheet not found: ' + CONFIG.MASTER_SHEET);

  const V = master.getDataRange().getValues();
  const nRows = V.length, nCols = V[0].length;
  const scan = Math.min(nRows, 15);

  let rProgram = -1, rFac = -1, rDeg = -1, rCyc = -1;
  for (let r = 0; r < scan; r++) {
    const a = String(V[r][0] || '').trim().toLowerCase();
    if (a === 'program') rProgram = r;
    else if (a === 'faculty') rFac = r;
    else if (a === 'degree') rDeg = r;
    else if (a === 'cycle') rCyc = r;
  }
  let rSub = -1, best = 0;
  for (let r = 0; r < scan; r++) {
    let cnt = 0;
    for (let c = 0; c < nCols; c++) if (String(V[r][c]).trim() === 'Status') cnt++;
    if (cnt > best) { best = cnt; rSub = r; }
  }
  if (rProgram < 0 || rSub < 0)
    throw new Error('Could not find the PROGRAM row or the Status sub-header row in the master.');

  const programs = [];
  for (let c = 0; c < nCols; c++) {
    if (String(V[rSub][c]).trim() !== 'Status') continue;
    const name = String(V[rProgram][c] || '').replace(/\s+/g, ' ').trim();
    if (!name) continue;
    programs.push({
      name: name,
      faculty: rFac >= 0 ? String(V[rFac][c] || '').trim() : '',
      degree: rDeg >= 0 ? String(V[rDeg][c] || '').replace(/\s+/g, ' ').trim() : '',
      cycle: rCyc >= 0 ? String(V[rCyc][c] || '').trim() : '',
      s: c, st: c + 1, co: c + 2
    });
  }

  // Single ascending pass: PHASE headers set the current phase; STAGE rows inherit it.
  const stages = [];
  let curPhase = '';
  for (let r = 0; r < nRows; r++) {
    const a = String(V[r][0] || '').trim();
    const pm = a.match(/^PHASE\s+([IVXLC]+)\s*:\s*(.+)$/i);
    if (pm) {
      const rn = pm[1].toUpperCase();
      curPhase = PHASE_MAP[rn] || ('Phase ' + rn + ': ' + titleish_(pm[2]));
      continue;
    }
    const m = a.match(/^STAGE\s+(\d+)\s*:\s*(.+)$/i);
    if (m) stages.push({ row: r, num: parseInt(m[1], 10), name: m[2].trim(), dur: V[r][1], phase: curPhase });
  }

  const out = [STAGE_HEADERS.slice()];
  programs.forEach(function (p) {
    stages.forEach(function (s) {
      out.push([
        p.name, p.faculty, p.degree, p.cycle, s.phase, s.num, s.name,
        V[s.row][p.s] || '',
        cleanDate_(V[s.row][p.st]),
        cleanDate_(V[s.row][p.co]),
        (typeof s.dur === 'number' ? s.dur : ''),
        p.name + ' ||S' + s.num
      ]);
    });
  });

  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, out.length, STAGE_HEADERS.length).setValues(out);

  if (out.length > 1) {
    const n = out.length - 1;
    sh.getRange(2, 9, n, 2).setNumberFormat('yyyy-mm-dd');     // Start (9), Complete (10)
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Complete', 'Incomplete', 'In progress', 'N/A'], true).build();
    sh.getRange(2, 8, n, 1).setDataValidation(rule);           // Status (8)
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, STAGE_HEADERS.length).setFontWeight('bold');

  try {
    ss.toast('Stages rebuilt: ' + (out.length - 1) + ' rows (with Phase) from "' +
             CONFIG.MASTER_SHEET + '".', 'PPR Tracker', 6);
  } catch (e) {}
}

function titleish_(s) {
  return String(s).toLowerCase().replace(/\b([a-z])/g, function (m) { return m.toUpperCase(); });
}
function cleanDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  return '';
}

// ===================== READ FLAT TABLE =====================
function getItems_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error('Sheet not found: ' + CONFIG.SHEET_NAME + ' — run "Rebuild Stages from master" first.');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const idx = {};
  values[0].forEach(function (h, c) {
    const field = CONFIG.HEADERS[String(h).trim()];
    if (field) idx[field] = c;
  });

  const items = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const status = normStatus_(row[idx.status]);
    if (!status) continue;
    const start = toDate_(row[idx.start]);
    const complete = toDate_(row[idx.complete]);
    if (!start && !complete) continue;

    const program = String(row[idx.program] || '').trim();
    const stageNum = row[idx.stageNum];
    items.push({
      key: (idx.key !== undefined && row[idx.key]) ? String(row[idx.key]).trim() : program + ' ||S' + stageNum,
      program: program,
      faculty: idx.faculty !== undefined ? String(row[idx.faculty] || '').trim() : '',
      degree: idx.degree !== undefined ? String(row[idx.degree] || '').trim() : '',
      cycle: idx.cycle !== undefined ? String(row[idx.cycle] || '').trim() : '',
      // phase is read but deliberately not used in the calendar event
      phase: idx.phase !== undefined ? String(row[idx.phase] || '').trim() : '',
      stageNum: stageNum, stageName: String(row[idx.stageName] || '').trim(),
      status: status, start: start, complete: complete
    });
  }
  return items;
}

function toDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  return null;
}

// ===================== CALENDAR SYNC =====================
function syncPPRtoCalendar() {
  const startMs = Date.now();
  const cal = getCalendar_();
  const map = readSyncMap_();
  const seen = {};
  const items = getItems_();

  let partial = false;
  for (let i = 0; i < items.length; i++) {
    if (Date.now() - startMs > CONFIG.RUN_BUDGET_MS) { partial = true; break; }
    const it = items[i];
    seen[it.key] = true;

    const style = statusStyle_(it.status);
    const title = style.emoji + ' ' + it.program + ' · S' + it.stageNum + ': ' + it.stageName;
    const desc = buildDesc_(it);

    let s = it.start || it.complete;
    let e = it.complete || it.start;
    if (s.getTime() > e.getTime()) { const t = s; s = e; e = t; }
    const singleDay = s.getTime() === e.getTime();
    const endExclusive = new Date(e.getFullYear(), e.getMonth(), e.getDate() + 1);

    let ev = null;
    if (map[it.key]) { try { ev = cal.getEventById(map[it.key]); } catch (err) { ev = null; } }

    if (ev) {
      ev.setTitle(title); ev.setDescription(desc);
      if (singleDay) ev.setAllDayDate(s); else ev.setAllDayDates(s, endExclusive);
    } else {
      ev = singleDay ? cal.createAllDayEvent(title, s) : cal.createAllDayEvent(title, s, endExclusive);
      ev.setDescription(desc);
      map[it.key] = ev.getId();
    }

    if (style.color) ev.setColor(style.color);
    ev.removeAllReminders();
    if (it.status !== 'Complete') CONFIG.EVENT_REMINDER_MINUTES.forEach(function (m) { ev.addPopupReminder(m); });
  }

  if (!partial) {
    Object.keys(map).forEach(function (k) {
      if (seen[k]) return;
      try { const ev = cal.getEventById(map[k]); if (ev) ev.deleteEvent(); } catch (e) {}
      delete map[k];
    });
  }

  writeSyncMap_(map);
  try {
    SpreadsheetApp.getActive().toast(
      partial ? 'Partial sync — run again to finish.'
              : 'Synced ' + items.length + ' tracked stages to "' + CONFIG.CALENDAR_NAME + '".',
      'PPR Tracker', 6);
  } catch (e) {}
}

function getCalendar_() {
  const cals = CalendarApp.getCalendarsByName(CONFIG.CALENDAR_NAME);
  if (!cals || !cals.length)
    throw new Error('Calendar "' + CONFIG.CALENDAR_NAME + '" not found. Run listMyCalendars to see exact titles.');
  return cals[0];
}

function buildDesc_(it) {
  const tz = Session.getScriptTimeZone();
  const fmt = function (d) { return d ? Utilities.formatDate(d, tz, 'EEE, MMM d, yyyy') : '—'; };
  // Note: Phase is intentionally omitted from the calendar event for now.
  return [
    'Program: ' + it.program,
    'Faculty: ' + (it.faculty || '—') + '    Degree: ' + (it.degree || '—'),
    (it.cycle ? 'Cycle: ' + it.cycle : ''),
    'Stage ' + it.stageNum + ': ' + it.stageName,
    'Status: ' + it.status,
    'Start: ' + fmt(it.start),
    'Target complete: ' + fmt(it.complete),
    '',
    'Synced from the Stages table.'
  ].filter(String).join('\n');
}

// ===================== BOOKKEEPING TAB =====================
function syncSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SYNC_SHEET);
  if (!sh) { sh = ss.insertSheet(CONFIG.SYNC_SHEET); sh.appendRow(['key', 'eventId']); sh.hideSheet(); }
  return sh;
}
function readSyncMap_() {
  const vals = syncSheet_().getDataRange().getValues();
  const map = {};
  for (let i = 1; i < vals.length; i++) if (vals[i][0]) map[vals[i][0]] = vals[i][1];
  return map;
}
function writeSyncMap_(map) {
  const sh = syncSheet_();
  sh.clearContents();
  const rows = [['key', 'eventId']];
  Object.keys(map).forEach(function (k) { rows.push([k, map[k]]); });
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
}

// ===================== DAILY DIGEST =====================
function sendPPRDigest() {
  const items = getItems_();
  const now = new Date();
  const t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const DAY = 86400000;

  const overdue = [], dueSoon = [], startingSoon = [];
  items.forEach(function (it) {
    if (it.status === 'Complete') return;
    if (it.complete) {
      const d = Math.round((midnight_(it.complete) - t0) / DAY);
      if (d < 0) overdue.push([it, d]); else if (d <= CONFIG.DUE_SOON_DAYS) dueSoon.push([it, d]);
    }
    if (it.start) {
      const ds = Math.round((midnight_(it.start) - t0) / DAY);
      if (ds >= 0 && ds <= CONFIG.START_SOON_DAYS) startingSoon.push([it, ds]);
    }
  });
  if (!overdue.length && !dueSoon.length && !startingSoon.length) return;

  const byDays = function (a, b) { return a[1] - b[1]; };
  overdue.sort(byDays); dueSoon.sort(byDays); startingSoon.sort(byDays);

  const tz = Session.getScriptTimeZone();
  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222">' +
    '<h2 style="margin:0 0 4px">PPR deadlines</h2>' +
    '<p style="color:#666;margin:0 0 16px">' + Utilities.formatDate(t0, tz, 'EEEE, MMMM d, yyyy') + '</p>' +
    digestBlock_('⚠️ Overdue', '#b00020', overdue, 'due', tz) +
    digestBlock_('⏰ Due soon', '#c47f00', dueSoon, 'due', tz) +
    digestBlock_('▶️ Starting soon', '#1a73e8', startingSoon, 'start', tz) +
    '<p style="color:#999;font-size:12px;margin-top:20px">From the Stages table. Complete and N/A stages are excluded.</p></div>';

  MailApp.sendEmail({
    to: CONFIG.DIGEST_RECIPIENTS || Session.getActiveUser().getEmail(),
    subject: 'PPR deadlines — ' + Utilities.formatDate(t0, tz, 'MMM d, yyyy'),
    htmlBody: html
  });
}

function digestBlock_(heading, color, rows, mode, tz) {
  if (!rows.length) return '';
  let html = '<h3 style="color:' + color + ';margin:14px 0 6px">' + heading + ' (' + rows.length + ')</h3>' +
             '<table cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">';
  rows.forEach(function (pair) {
    const it = pair[0], d = pair[1];
    const date = mode === 'due' ? it.complete : it.start;
    const when = d === 0 ? 'today' : d < 0 ? (-d + ' day(s) ago') : ('in ' + d + ' day(s)');
    html += '<tr style="border-bottom:1px solid #eee">' +
      '<td style="white-space:nowrap;color:#555">' + Utilities.formatDate(date, tz, 'MMM d') + '</td>' +
      '<td><b>' + escapeHtml_(it.program) + '</b><br><span style="color:#666">S' + it.stageNum +
        ': ' + escapeHtml_(it.stageName) + ' — ' + escapeHtml_(it.status) + '</span></td>' +
      '<td style="white-space:nowrap;text-align:right;color:' + color + '">' + when + '</td></tr>';
  });
  return html + '</table>';
}

function midnight_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime(); }
function escapeHtml_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

// ===================== TRIGGERS / HELPERS =====================
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'sendPPRDigest' || fn === 'rebuildAndSync' || fn === 'syncPPRtoCalendar') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rebuildAndSync').timeBased().everyDays(1).atHour(CONFIG.SYNC_HOUR).create();
  ScriptApp.newTrigger('sendPPRDigest').timeBased().everyDays(1).atHour(CONFIG.DIGEST_HOUR).create();
  try { SpreadsheetApp.getActive().toast('Daily triggers installed.', 'PPR Tracker', 5); } catch (e) {}
}

function listMyCalendars() {
  CalendarApp.getAllOwnedCalendars().forEach(function (c) {
    Logger.log('"' + c.getName() + '"  id: ' + c.getId());
  });
}