/**
 * PPR Calendar Sync + Deadline Notifier  (v7)
 * Toronto Metropolitan University — YSGPS
 *
 * Changes:
 *   - Adds Step name support.
 *   - Google Calendar event title shows Step name instead of Stage name.
 *   - Calendar events no longer stretch from Start to Complete.
 *   - If Complete date exists, event displays for the final full week:
 *       Complete date - 6 days through Complete date.
 *   - No emoji icons in event titles.
 *   - No status-based event colors.
 *   - Event color is based on Phase.
 *   - Rebuild + sync now first deletes all existing events from the three
 *     PPR phase calendars, clears sync map, rebuilds Stages, then syncs fresh.
 *
 * Expected master layout:
 *   Column A: PHASE / STAGE labels
 *   Column B: Step name
 *   Column C: Duration
 *   Program blocks begin at their Status / Start / Complete columns.
 *
 * FLOW
 *   buildStagesTable()  — master matrix -> flat "Stages" tab
 *   syncPPRtoCalendar() — "Stages" tab -> three phase-specific calendars
 *   sendPPRDigest()     — emails overdue / due-soon / starting-soon
 *   rebuildAndSync()    — delete old events + rebuild + sync
 *
 * Statuses:
 *   Complete, In progress, Incomplete, N/A.
 *   N/A and blank -> not tracked.
 */

// ===================== CONFIG =====================
const CONFIG = {
  MASTER_SHEET: 'PPR Tracking (Master)',
  SHEET_NAME: 'Stages',

  CALENDAR_NAMES: {
    'Phase 1: Self-Study': 'PPR Tracker Phase 1',
    'Phase 2: Site Visit': 'PPR Tracker Phase 2',
    'Phase 3: PRT Report': 'PPR Tracker Phase 3'
  },

  PHASE_COLORS: {
    'Phase 1: Self-Study': CalendarApp.EventColor.BLUE,
    'Phase 2: Site Visit': CalendarApp.EventColor.GREEN,
    'Phase 3: PRT Report': CalendarApp.EventColor.ORANGE
  },

  SYNC_SHEET: '_PPR_SYNC',

  DUE_SOON_DAYS: 7,
  START_SOON_DAYS: 3,
  DIGEST_RECIPIENTS: '',
  DIGEST_HOUR: 7,
  SYNC_HOUR: 1,

  EVENT_REMINDER_MINUTES: [7 * 24 * 60, 24 * 60],
  RUN_BUDGET_MS: 5 * 60 * 1000,

  HEADERS: {
    'Program': 'program',
    'Faculty': 'faculty',
    'Degree': 'degree',
    'Cycle': 'cycle',
    'Phase': 'phase',
    'Stage No': 'stageNum',
    'Stage': 'stageName',
    'Step': 'stepName',
    'Status': 'status',
    'Start': 'start',
    'Complete': 'complete',
    'Duration (days)': 'duration',
    'Key': 'key'
  }
};

const STAGE_HEADERS = [
  'Program',
  'Faculty',
  'Degree',
  'Cycle',
  'Phase',
  'Stage No',
  'Stage',
  'Step',
  'Status',
  'Start',
  'Complete',
  'Duration (days)',
  'Key'
];

const PHASE_MAP = {
  'I': 'Phase 1: Self-Study',
  'II': 'Phase 2: Site Visit',
  'III': 'Phase 3: PRT Report'
};

// ===================== NORMALIZATION =====================
function normStatus_(raw) {
  const s = String(raw || '').trim().toLowerCase();

  if (s === 'complete') return 'Complete';
  if (s === 'in progress' || s === 'in-progress' || s === 'inprogress') return 'In progress';
  if (s === 'incomplete' || s === 'in complete') return 'Incomplete';

  return null;
}

function phaseColor_(phase) {
  return CONFIG.PHASE_COLORS[phase] || CalendarApp.EventColor.BLUE;
}

// ===================== MENU =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('PPR Tracker')
    .addItem('Delete old events + rebuild + sync now', 'rebuildAndSync')
    .addSeparator()
    .addItem('Rebuild Stages from master', 'buildStagesTable')
    .addItem('Sync to calendar only', 'syncPPRtoCalendar')
    .addItem('Send deadline digest now', 'sendPPRDigest')
    .addSeparator()
    .addItem('Install daily triggers', 'installTriggers')
    .addToUi();
}

function rebuildAndSync() {
  deleteAllPPRCalendarEvents_();
  clearSyncMap_();
  buildStagesTable();
  syncPPRtoCalendar();
}

// ===================== FLATTEN MASTER -> STAGES =====================
function buildStagesTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName(CONFIG.MASTER_SHEET);

  if (!master) {
    throw new Error('Master sheet not found: ' + CONFIG.MASTER_SHEET);
  }

  const V = master.getDataRange().getValues();
  const nRows = V.length;
  const nCols = V[0].length;
  const scan = Math.min(nRows, 15);

  let rProgram = -1;
  let rFac = -1;
  let rDeg = -1;
  let rCyc = -1;

  for (let r = 0; r < scan; r++) {
    const a = String(V[r][0] || '').trim().toLowerCase();

    if (a === 'program') rProgram = r;
    else if (a === 'faculty') rFac = r;
    else if (a === 'degree') rDeg = r;
    else if (a === 'cycle') rCyc = r;
  }

  let rSub = -1;
  let best = 0;

  for (let r = 0; r < scan; r++) {
    let cnt = 0;

    for (let c = 0; c < nCols; c++) {
      if (String(V[r][c]).trim() === 'Status') cnt++;
    }

    if (cnt > best) {
      best = cnt;
      rSub = r;
    }
  }

  if (rProgram < 0 || rSub < 0) {
    throw new Error('Could not find the PROGRAM row or the Status sub-header row in the master.');
  }

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
      s: c,
      st: c + 1,
      co: c + 2
    });
  }

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

    if (m) {
      stages.push({
        row: r,
        num: parseInt(m[1], 10),
        name: m[2].trim(),
        stepName: String(V[r][1] || '').trim(),
        dur: V[r][2],
        phase: curPhase
      });
    }
  }

  const out = [STAGE_HEADERS.slice()];

  programs.forEach(function (p) {
    stages.forEach(function (s) {
      out.push([
        p.name,
        p.faculty,
        p.degree,
        p.cycle,
        s.phase,
        s.num,
        s.name,
        s.stepName,
        V[s.row][p.s] || '',
        cleanDate_(V[s.row][p.st]),
        cleanDate_(V[s.row][p.co]),
        typeof s.dur === 'number' ? s.dur : '',
        p.name + ' ||S' + s.num
      ]);
    });
  });

  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  sh.clear();
  sh.getRange(1, 1, out.length, STAGE_HEADERS.length).setValues(out);

  if (out.length > 1) {
    const n = out.length - 1;

    // Start and Complete are columns 10 and 11.
    sh.getRange(2, 10, n, 2).setNumberFormat('yyyy-mm-dd');

    // Status is column 9.
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Complete', 'Incomplete', 'In progress', 'N/A'], true)
      .build();

    sh.getRange(2, 9, n, 1).setDataValidation(rule);
  }

  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, STAGE_HEADERS.length).setFontWeight('bold');

  try {
    ss.toast(
      'Stages rebuilt: ' + (out.length - 1) + ' rows with Phase and Step from "' + CONFIG.MASTER_SHEET + '".',
      'PPR Tracker',
      6
    );
  } catch (e) {}
}

function titleish_(s) {
  return String(s)
    .toLowerCase()
    .replace(/\b([a-z])/g, function (m) {
      return m.toUpperCase();
    });
}

function cleanDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  return '';
}

// ===================== READ FLAT TABLE =====================
function getItems_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);

  if (!sh) {
    throw new Error('Sheet not found: ' + CONFIG.SHEET_NAME + ' — run "Rebuild Stages from master" first.');
  }

  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const idx = {};

  values[0].forEach(function (h, c) {
    const field = CONFIG.HEADERS[String(h).trim()];

    if (field) {
      idx[field] = c;
    }
  });

  const required = ['program', 'phase', 'stageNum', 'stageName', 'status', 'start', 'complete'];

  required.forEach(function (field) {
    if (idx[field] === undefined) {
      throw new Error('Missing required column in Stages sheet: ' + field);
    }
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
      key: idx.key !== undefined && row[idx.key]
        ? String(row[idx.key]).trim()
        : program + ' ||S' + stageNum,

      program: program,
      faculty: idx.faculty !== undefined ? String(row[idx.faculty] || '').trim() : '',
      degree: idx.degree !== undefined ? String(row[idx.degree] || '').trim() : '',
      cycle: idx.cycle !== undefined ? String(row[idx.cycle] || '').trim() : '',
      phase: idx.phase !== undefined ? String(row[idx.phase] || '').trim() : '',
      stageNum: stageNum,
      stageName: String(row[idx.stageName] || '').trim(),
      stepName: idx.stepName !== undefined ? String(row[idx.stepName] || '').trim() : '',
      status: status,
      start: start,
      complete: complete
    });
  }

  return items;
}

function toDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  return null;
}

// ===================== CALENDAR SYNC =====================
function syncPPRtoCalendar() {
  const startMs = Date.now();
  const calendarsByPhase = getCalendarsByPhase_();
  const map = readSyncMap_();
  const seen = {};
  const items = getItems_();

  let partial = false;

  for (let i = 0; i < items.length; i++) {
    if (Date.now() - startMs > CONFIG.RUN_BUDGET_MS) {
      partial = true;
      break;
    }

    const it = items[i];
    const cal = getCalendarForItem_(it, calendarsByPhase);

    seen[it.key] = true;

    const displayName = it.stepName || it.stageName;
    const title = it.program + ' · ' + displayName;
    const desc = buildDesc_(it);

    const eventWindow = getEventWindow_(it);
    const s = eventWindow.start;
    const endExclusive = eventWindow.endExclusive;

    let ev = null;
    const existing = map[it.key];

    if (existing && existing.phase && existing.phase !== it.phase) {
      deleteEventFromAnyPhase_(existing.eventId, calendarsByPhase);
      ev = null;
    } else if (existing && existing.eventId) {
      try {
        ev = cal.getEventById(existing.eventId);
      } catch (err) {
        ev = null;
      }
    }

    if (ev) {
      ev.setTitle(title);
      ev.setDescription(desc);
      ev.setAllDayDates(s, endExclusive);
    } else {
      ev = cal.createAllDayEvent(title, s, endExclusive);
      ev.setDescription(desc);
    }

    ev.setColor(phaseColor_(it.phase));

    ev.removeAllReminders();

    if (it.status !== 'Complete') {
      CONFIG.EVENT_REMINDER_MINUTES.forEach(function (m) {
        ev.addPopupReminder(m);
      });
    }

    map[it.key] = {
      eventId: ev.getId(),
      phase: it.phase
    };
  }

  if (!partial) {
    Object.keys(map).forEach(function (k) {
      if (seen[k]) return;

      deleteEventFromAnyPhase_(map[k].eventId, calendarsByPhase);
      delete map[k];
    });
  }

  writeSyncMap_(map);

  try {
    SpreadsheetApp.getActive().toast(
      partial
        ? 'Partial sync — run again to finish.'
        : 'Synced ' + items.length + ' tracked stages to the three PPR phase calendars.',
      'PPR Tracker',
      6
    );
  } catch (e) {}
}

function getEventWindow_(it) {
  const anchor = it.complete || it.start;

  if (!anchor) {
    throw new Error('Missing both Start and Complete date for: ' + it.program + ' S' + it.stageNum);
  }

  // Preferred rule:
  // If Complete exists, show the final full week ending on Complete.
  // Example: Complete June 29 -> event runs June 23 to June 29.
  // Google Calendar all-day end date is exclusive, so endExclusive is Complete + 1 day.
  if (it.complete) {
    return {
      start: new Date(
        it.complete.getFullYear(),
        it.complete.getMonth(),
        it.complete.getDate() - 6
      ),
      endExclusive: new Date(
        it.complete.getFullYear(),
        it.complete.getMonth(),
        it.complete.getDate() + 1
      )
    };
  }

  // Fallback if Complete is blank:
  // create a one-day event on Start.
  return {
    start: new Date(
      anchor.getFullYear(),
      anchor.getMonth(),
      anchor.getDate()
    ),
    endExclusive: new Date(
      anchor.getFullYear(),
      anchor.getMonth(),
      anchor.getDate() + 1
    )
  };
}

function getCalendarsByPhase_() {
  const out = {};

  Object.keys(CONFIG.CALENDAR_NAMES).forEach(function (phase) {
    const calName = CONFIG.CALENDAR_NAMES[phase];
    const cals = CalendarApp.getCalendarsByName(calName);

    if (!cals || !cals.length) {
      throw new Error(
        'Calendar "' + calName + '" not found. Please create it first or run listMyCalendars to check the exact title.'
      );
    }

    out[phase] = cals[0];
  });

  return out;
}

function getCalendarForItem_(it, calendarsByPhase) {
  if (!it.phase) {
    throw new Error('Missing phase for: ' + it.program + ' S' + it.stageNum);
  }

  const cal = calendarsByPhase[it.phase];

  if (!cal) {
    throw new Error(
      'No calendar configured for phase "' + it.phase + '". ' +
      'Check CONFIG.CALENDAR_NAMES and the Phase values in the Stages sheet.'
    );
  }

  return cal;
}

function deleteEventFromAnyPhase_(eventId, calendarsByPhase) {
  if (!eventId) return false;

  return Object.keys(calendarsByPhase).some(function (phase) {
    try {
      const ev = calendarsByPhase[phase].getEventById(eventId);

      if (ev) {
        ev.deleteEvent();
        return true;
      }
    } catch (e) {}

    return false;
  });
}

// ===================== FULL CALENDAR CLEANUP =====================
function deleteAllPPRCalendarEvents_() {
  const calendarsByPhase = getCalendarsByPhase_();

  // Wide date window to catch old and future PPR events.
  const from = new Date(2000, 0, 1);
  const to = new Date(2100, 0, 1);

  Object.keys(calendarsByPhase).forEach(function (phase) {
    const cal = calendarsByPhase[phase];
    const events = cal.getEvents(from, to);

    events.forEach(function (ev) {
      try {
        ev.deleteEvent();
      } catch (e) {}
    });
  });

  try {
    SpreadsheetApp.getActive().toast(
      'Deleted existing events from all PPR phase calendars.',
      'PPR Tracker',
      5
    );
  } catch (e) {}
}

function clearSyncMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SYNC_SHEET);

  if (!sh) return;

  sh.clearContents();
  sh.getRange(1, 1, 1, 3).setValues([['key', 'eventId', 'phase']]);

  try {
    sh.hideSheet();
  } catch (e) {}
}

function buildDesc_(it) {
  const tz = Session.getScriptTimeZone();

  const fmt = function (d) {
    return d ? Utilities.formatDate(d, tz, 'EEE, MMM d, yyyy') : '—';
  };

  return [
    'Program: ' + it.program,
    'Faculty: ' + (it.faculty || '—') + '    Degree: ' + (it.degree || '—'),
    it.cycle ? 'Cycle: ' + it.cycle : '',
    'Step: ' + (it.stepName || it.stageName),
    'Stage ' + it.stageNum + ': ' + it.stageName,
    'Status: ' + it.status,
    'Start: ' + fmt(it.start),
    'Target complete: ' + fmt(it.complete),
    '',
    'Calendar display window: final week ending on the target complete date.',
    'Synced from the Stages table.'
  ].filter(String).join('\n');
}

// ===================== BOOKKEEPING TAB =====================
function syncSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SYNC_SHEET);

  if (!sh) {
    sh = ss.insertSheet(CONFIG.SYNC_SHEET);
    sh.appendRow(['key', 'eventId', 'phase']);
    sh.hideSheet();
  }

  return sh;
}

function readSyncMap_() {
  const vals = syncSheet_().getDataRange().getValues();
  const map = {};

  for (let i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;

    map[vals[i][0]] = {
      eventId: vals[i][1],
      phase: vals[i][2] || ''
    };
  }

  return map;
}

function writeSyncMap_(map) {
  const sh = syncSheet_();

  sh.clearContents();

  const rows = [['key', 'eventId', 'phase']];

  Object.keys(map).forEach(function (k) {
    rows.push([
      k,
      map[k].eventId || '',
      map[k].phase || ''
    ]);
  });

  sh.getRange(1, 1, rows.length, 3).setValues(rows);
}

// ===================== DAILY DIGEST =====================
function sendPPRDigest() {
  const items = getItems_();
  const now = new Date();
  const t0 = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const DAY = 86400000;

  const overdue = [];
  const dueSoon = [];
  const startingSoon = [];

  items.forEach(function (it) {
    if (it.status === 'Complete') return;

    if (it.complete) {
      const d = Math.round((midnight_(it.complete) - t0) / DAY);

      if (d < 0) {
        overdue.push([it, d]);
      } else if (d <= CONFIG.DUE_SOON_DAYS) {
        dueSoon.push([it, d]);
      }
    }

    if (it.start) {
      const ds = Math.round((midnight_(it.start) - t0) / DAY);

      if (ds >= 0 && ds <= CONFIG.START_SOON_DAYS) {
        startingSoon.push([it, ds]);
      }
    }
  });

  if (!overdue.length && !dueSoon.length && !startingSoon.length) {
    return;
  }

  const byDays = function (a, b) {
    return a[1] - b[1];
  };

  overdue.sort(byDays);
  dueSoon.sort(byDays);
  startingSoon.sort(byDays);

  const tz = Session.getScriptTimeZone();

  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222">' +
    '<h2 style="margin:0 0 4px">PPR deadlines</h2>' +
    '<p style="color:#666;margin:0 0 16px">' +
    Utilities.formatDate(t0, tz, 'EEEE, MMMM d, yyyy') +
    '</p>' +
    digestBlock_('Overdue', '#b00020', overdue, 'due', tz) +
    digestBlock_('Due soon', '#c47f00', dueSoon, 'due', tz) +
    digestBlock_('Starting soon', '#1a73e8', startingSoon, 'start', tz) +
    '<p style="color:#999;font-size:12px;margin-top:20px">From the Stages table. Complete and N/A stages are excluded.</p>' +
    '</div>';

  MailApp.sendEmail({
    to: CONFIG.DIGEST_RECIPIENTS || Session.getActiveUser().getEmail(),
    subject: 'PPR deadlines — ' + Utilities.formatDate(t0, tz, 'MMM d, yyyy'),
    htmlBody: html
  });
}

function digestBlock_(heading, color, rows, mode, tz) {
  if (!rows.length) {
    return '';
  }

  let html =
    '<h3 style="color:' + color + ';margin:14px 0 6px">' +
    heading +
    ' (' + rows.length + ')' +
    '</h3>' +
    '<table cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">';

  rows.forEach(function (pair) {
    const it = pair[0];
    const d = pair[1];
    const date = mode === 'due' ? it.complete : it.start;
    const when = d === 0
      ? 'today'
      : d < 0
        ? -d + ' day(s) ago'
        : 'in ' + d + ' day(s)';

    html +=
      '<tr style="border-bottom:1px solid #eee">' +
      '<td style="white-space:nowrap;color:#555">' +
      Utilities.formatDate(date, tz, 'MMM d') +
      '</td>' +
      '<td><b>' +
      escapeHtml_(it.program) +
      '</b><br><span style="color:#666">' +
      escapeHtml_(it.stepName || it.stageName) +
      ' — ' +
      escapeHtml_(it.status) +
      '</span></td>' +
      '<td style="white-space:nowrap;text-align:right;color:' +
      color +
      '">' +
      when +
      '</td>' +
      '</tr>';
  });

  return html + '</table>';
}

function midnight_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// ===================== TRIGGERS / HELPERS =====================
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();

    if (fn === 'sendPPRDigest' || fn === 'rebuildAndSync' || fn === 'syncPPRtoCalendar') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('rebuildAndSync')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.SYNC_HOUR)
    .create();

  ScriptApp.newTrigger('sendPPRDigest')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.DIGEST_HOUR)
    .create();

  try {
    SpreadsheetApp.getActive().toast('Daily triggers installed.', 'PPR Tracker', 5);
  } catch (e) {}
}

function listMyCalendars() {
  CalendarApp.getAllOwnedCalendars().forEach(function (c) {
    Logger.log('"' + c.getName() + '"  id: ' + c.getId());
  });
}