// ═══════════════════════════════════════════════════════════════
// YearPlannerService.gs — CDOA Annual Calendar / Year Planner.
// Recurring events (bills, meetings, contract renewals) reproduced
// from the association's own color-coded template, computed for any
// year, synced to a dedicated Google Calendar, and reminded to EC
// Members only — by email automatically (1 month / 1 week / 1 day
// before) and by a one-click WhatsApp-ready message (WhatsApp cannot
// be sent programmatically without a paid Business API, so this is
// the honest, free equivalent: pre-filled, one tap to send).
// ═══════════════════════════════════════════════════════════════
var YearPlannerService = (function() {
  var RULES_SHEET = 'PlannerRules';
  var RULES_HEADERS = ['rule_id','title','color','rule_type','params','active','notes','created_at','updated_at'];
  var RC = { ID:0, TITLE:1, COLOR:2, TYPE:3, PARAMS:4, ACTIVE:5, NOTES:6, CREATED:7, UPDATED:8 };

  var LOG_SHEET = 'PlannerReminderLog';
  var LOG_HEADERS = ['log_id','rule_id','occurrence_date','stage','channel','sent_at'];

  var MAP_SHEET = 'PlannerCalendarEvents';
  var MAP_HEADERS = ['rule_id','occurrence_date','calendar_event_id'];

  function ensureSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    [[RULES_SHEET, RULES_HEADERS], [LOG_SHEET, LOG_HEADERS], [MAP_SHEET, MAP_HEADERS]].forEach(function(spec) {
      var sh = ss.getSheetByName(spec[0]);
      if (!sh) {
        sh = ss.insertSheet(spec[0]);
        sh.getRange(1, 1, 1, spec[1].length).setValues([spec[1]]).setFontWeight('bold');
        sh.setFrozenRows(1);
      }
    });
  }

  // ── Seed the 11 events straight from the association's template ──
  function seedInitial() {
    ensureSheets();
    if (getAllRules().length) throw new Error('Planner rules already exist — seeding only runs once, so it can never overwrite live edits.');
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RULES_SHEET);
    var now = new Date().toISOString();
    var rows = [
      ['Send Treasurer\'s Report to all Owners', '#34A853', 'specific_months_day', { day: 2, months: [1,2,3,4,5,6,7,8,9,10,11,12] }, 'Monthly'],
      ['CDOA EC Meeting', '#00FFFF', 'monthly_last_weekday', { weekday: 0 }, 'Last Sunday of every month'],
      ['Salary Payments', '#4A86E8', 'monthly_last_day', {}, 'Last day of the month'],
      ['KWA Bill Due Date', '#FF00FF', 'specific_months_day', { day: 10, months: [2,4,6,8,10,12] }, 'Every other month'],
      ['KSEB Bill Due Date', '#FF0000', 'specific_months_day', { day: 11, months: [1,2,3,4,5,6,7,8,9,10,11,12] }, '11th of every month'],
      ['AQUAPURE (STP/WTP) Contract Renewal', '#FFFF00', 'annual_date', { month: 1, day: 1 }, 'Annual'],
      ['Land Tax Payment Due', '#9900FF', 'annual_date', { month: 4, day: 1 }, 'Annual'],
      ['Fire NOC Renewal', '#0000FF', 'annual_date', { month: 7, day: 25 }, 'Annual'],
      ['CDOA General Body Meeting (GBM)', '#FBBC04', 'annual_date', { month: 11, day: 10 }, 'Annual'],
      ['WTP/STP Contract Expiration', '#8B4513', 'annual_date', { month: 12, day: 31 }, 'Annual'],
      ['Lift Contract Expiration — #1', '#000000', 'one_off_date', { year: 2024, month: 6, day: 21 }, 'One-off; add a new rule for future renewals'],
      ['Lift Contract Expiration — #2', '#000000', 'one_off_date', { year: 2024, month: 7, day: 4 }, 'One-off; add a new rule for future renewals'],
      ['Lift Contract Expiration — #3', '#000000', 'one_off_date', { year: 2024, month: 8, day: 16 }, 'One-off; add a new rule for future renewals']
    ];
    var out = rows.map(function(r) {
      return [Database.generateId('PLR'), r[0], r[1], r[2], JSON.stringify(r[3]), 'Yes', r[4], now, now];
    });
    sh.getRange(sh.getLastRow() + 1, 1, out.length, RULES_HEADERS.length).setValues(out);
    return { success: true, rules: out.length };
  }

  function _toObj(r) {
    var params = {};
    try { params = JSON.parse(r[RC.PARAMS] || '{}'); } catch (e) {}
    return { rule_id: String(r[RC.ID] || ''), title: String(r[RC.TITLE] || ''),
             color: String(r[RC.COLOR] || '#64748b'), rule_type: String(r[RC.TYPE] || ''),
             params: params, active: String(r[RC.ACTIVE]) !== 'No', notes: String(r[RC.NOTES] || '') };
  }

  function getAllRules() {
    ensureSheets();
    return Database.getAll(RULES_SHEET).map(_toObj).filter(function(r) { return r.rule_id; });
  }

  function addRule(d) {
    if (!d || !d.title || !d.rule_type) throw new Error('title and rule_type are required.');
    ensureSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RULES_SHEET);
    var now = new Date().toISOString();
    var id = Database.generateId('PLR');
    sh.appendRow([id, String(d.title), String(d.color || '#64748b'), String(d.rule_type),
                  JSON.stringify(d.params || {}), d.active === false ? 'No' : 'Yes', String(d.notes || ''), now, now]);
    return { success: true, rule_id: id };
  }

  function updateRule(id, d) {
    var found = Database.findByColumn(RULES_SHEET, RC.ID, id);
    if (!found) throw new Error('Rule not found: ' + id);
    var r = found.data;
    if (d.title !== undefined)     r[RC.TITLE]  = String(d.title);
    if (d.color !== undefined)     r[RC.COLOR]  = String(d.color);
    if (d.rule_type !== undefined) r[RC.TYPE]   = String(d.rule_type);
    if (d.params !== undefined)    r[RC.PARAMS] = JSON.stringify(d.params);
    if (d.active !== undefined)    r[RC.ACTIVE] = d.active ? 'Yes' : 'No';
    if (d.notes !== undefined)     r[RC.NOTES]  = String(d.notes);
    r[RC.UPDATED] = new Date().toISOString();
    Database.updateRow(RULES_SHEET, found.rowIndex, r);
    return { success: true };
  }

  function deleteRule(id) {
    var found = Database.findByColumn(RULES_SHEET, RC.ID, id);
    if (!found) throw new Error('Rule not found: ' + id);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RULES_SHEET).deleteRow(found.rowIndex);
    return { success: true };
  }

  // ── Occurrence calculation — pure date math, no side effects ──
  function _pad2(n) { return n < 10 ? '0' + n : String(n); }
  function _iso(y, m, d) { return y + '-' + _pad2(m) + '-' + _pad2(d); }
  function _daysInMonth(y, m) { return new Date(y, m, 0).getDate(); } // m is 1-indexed here

  function _occurrencesForRule(rule, year) {
    var out = [];
    var p = rule.params || {};
    if (rule.rule_type === 'annual_date') {
      out.push(_iso(year, p.month, p.day));
    } else if (rule.rule_type === 'one_off_date') {
      if (Number(p.year) === year) out.push(_iso(year, p.month, p.day));
    } else if (rule.rule_type === 'specific_months_day') {
      (p.months || []).forEach(function(m) {
        var dim = _daysInMonth(year, m);
        out.push(_iso(year, m, Math.min(p.day, dim)));
      });
    } else if (rule.rule_type === 'monthly_last_day') {
      // Last calendar day of the month — shifted to the preceding
      // Saturday if that day is a Sunday (banks are typically closed
      // Sundays; this matched 11 of the 12 months in the association's
      // own template exactly).
      for (var m1 = 1; m1 <= 12; m1++) {
        var dim1 = _daysInMonth(year, m1);
        var wd1 = new Date(year, m1 - 1, dim1).getDay();
        out.push(_iso(year, m1, wd1 === 0 ? dim1 - 1 : dim1));
      }
    } else if (rule.rule_type === 'monthly_last_weekday') {
      var wd = Number(p.weekday) || 0; // 0 = Sunday
      for (var m2 = 1; m2 <= 12; m2++) {
        var dim2 = _daysInMonth(year, m2);
        for (var d2 = dim2; d2 >= dim2 - 6; d2--) {
          if (new Date(year, m2 - 1, d2).getDay() === wd) { out.push(_iso(year, m2, d2)); break; }
        }
      }
    }
    return out;
  }

  // Every occurrence for a year, across all active rules, sorted.
  function getOccurrencesForYear(year) {
    year = Number(year);
    var rules = getAllRules().filter(function(r) { return r.active; });
    var out = [];
    rules.forEach(function(rule) {
      _occurrencesForRule(rule, year).forEach(function(dateStr) {
        out.push({ date: dateStr, rule_id: rule.rule_id, title: rule.title, color: rule.color, notes: rule.notes });
      });
    });
    out.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    return { year: year, occurrences: out };
  }

  // ── EC Member reminders — email now, WhatsApp one-click ──
  var STAGES = [
    { key: '1mo', days: 30, label: '1 month before' },
    { key: '1wk', days: 7,  label: '1 week before' },
    { key: '1day', days: 1, label: '1 day before' }
  ];

  function _addDays(dateStr, n) {
    var parts = dateStr.split('-').map(Number);
    var d = new Date(parts[0], parts[1] - 1, parts[2]);
    d.setDate(d.getDate() - n);
    return _iso(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }

  function _todayIso() {
    var tz = Session.getScriptTimeZone();
    return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  }

  function _ecEmails() {
    var out = [];
    try {
      CommitteeService.getAll().current.forEach(function(m) {
        if (m.email && m.email.indexOf('@') > -1) out.push({ name: m.name, email: m.email });
      });
    } catch (e) {}
    return out;
  }

  function _alreadyLogged(ruleId, date, stage) {
    var rows = Database.getAll(LOG_SHEET);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][1]) === ruleId && String(rows[i][2]) === date && String(rows[i][3]) === stage) return true;
    }
    return false;
  }

  function _logSent(ruleId, date, stage, channel) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
    sh.appendRow([Database.generateId('PRL'), ruleId, date, stage, channel, new Date().toISOString()]);
  }

  // What's due TODAY across the surrounding years — used both by the
  // page (to show/send) and the daily trigger (to email automatically).
  function getDueReminders() {
    var today = _todayIso();
    var y = Number(today.slice(0, 4));
    var occ = getOccurrencesForYear(y - 1).occurrences
      .concat(getOccurrencesForYear(y).occurrences)
      .concat(getOccurrencesForYear(y + 1).occurrences);

    var due = [];
    occ.forEach(function(o) {
      STAGES.forEach(function(stage) {
        var reminderDate = _addDays(o.date, stage.days);
        if (reminderDate !== today) return;
        due.push({ rule_id: o.rule_id, title: o.title, color: o.color, event_date: o.date,
                   stage: stage.key, stage_label: stage.label,
                   emailSent: _alreadyLogged(o.rule_id, o.date, stage.key + ':email'),
                   whatsappSent: _alreadyLogged(o.rule_id, o.date, stage.key + ':whatsapp') });
      });
    });
    return { today: today, due: due, ecEmails: _ecEmails().map(function(e) { return e.name; }) };
  }

  // Called by a daily time-driven trigger (set up once in the Apps
  // Script editor — see Year Planner page for instructions). Sends
  // exactly the reminders due today, exactly once each (logged so a
  // second trigger firing the same day never double-sends).
  function sendDuePlannerReminders() {
    var d = getDueReminders();
    var ec = _ecEmails();
    if (!ec.length) return { sent: 0, reason: 'No EC Member has an email on file.' };
    var sent = 0;
    d.due.forEach(function(item) {
      if (item.emailSent) return;
      var subject = '[CDOA Reminder] ' + item.title + ' — ' + item.event_date + ' (' + item.stage_label + ')';
      var body = 'This is an automatic reminder from OXYGEN.\n\n' +
        item.title + '\nDate: ' + item.event_date + '\nReminder: ' + item.stage_label + '\n\n' +
        (item.notes || '') + '\n\n— Confident Daffodils Owners Association';
      try {
        MailApp.sendEmail({ to: ec.map(function(e) { return e.email; }).join(','), subject: subject, body: body });
        _logSent(item.rule_id, item.event_date, item.stage + ':email', 'email');
        sent++;
      } catch (e) { /* one failure shouldn't block the others */ }
    });
    return { sent: sent, total: d.due.length };
  }

  // Manual "send now" from the page — same logic, on demand.
  function sendReminderNow(ruleId, eventDate, stage) {
    var ec = _ecEmails();
    if (!ec.length) throw new Error('No current EC Member has an email on file (add one in EC Committee).');
    var rules = getAllRules();
    var rule = null;
    rules.forEach(function(r) { if (r.rule_id === ruleId) rule = r; });
    if (!rule) throw new Error('Rule not found.');
    var stageInfo = STAGES.filter(function(s) { return s.key === stage; })[0];
    var subject = '[CDOA Reminder] ' + rule.title + ' — ' + eventDate + (stageInfo ? ' (' + stageInfo.label + ')' : '');
    var body = 'This is a reminder from OXYGEN.\n\n' + rule.title + '\nDate: ' + eventDate + '\n\n' +
      (rule.notes || '') + '\n\n— Confident Daffodils Owners Association';
    MailApp.sendEmail({ to: ec.map(function(e) { return e.email; }).join(','), subject: subject, body: body });
    _logSent(ruleId, eventDate, stage + ':email', 'email');
    return { success: true, recipients: ec.length };
  }

  // Marks a WhatsApp reminder as sent (called after the page opens the
  // wa.me link — the actual send is the human's own tap, this just
  // stops the page nagging about it again today).
  function markWhatsappSent(ruleId, eventDate, stage) {
    _logSent(ruleId, eventDate, stage + ':whatsapp', 'whatsapp');
    return { success: true };
  }

  function whatsappMessageFor(ruleId, eventDate, stage) {
    var rule = null;
    getAllRules().forEach(function(r) { if (r.rule_id === ruleId) rule = r; });
    if (!rule) throw new Error('Rule not found.');
    var stageInfo = STAGES.filter(function(s) { return s.key === stage; })[0];
    var text = '📅 *CDOA Reminder*\n' + rule.title + '\nDate: ' + eventDate +
      (stageInfo ? '\n(' + stageInfo.label + ')' : '') + (rule.notes ? '\n' + rule.notes : '');
    return { text: text, url: 'https://wa.me/?text=' + encodeURIComponent(text) };
  }

  // ── Google Calendar sync ──
  function _getOrCreateCalendar() {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('PLANNER_CALENDAR_ID');
    if (id) {
      try { return CalendarApp.getCalendarById(id); } catch (e) { /* fall through and recreate */ }
    }
    var cal = CalendarApp.createCalendar('CDOA Year Planner');
    props.setProperty('PLANNER_CALENDAR_ID', cal.getId());
    return cal;
  }

  function _nearestCalendarColor(hex) {
    // Google Calendar only accepts a fixed palette (CalendarApp.EventColor).
    // Map our arbitrary hex to the closest of those by simple RGB distance.
    var palette = [
      ['#A4BDFC', CalendarApp.EventColor.PALE_BLUE], ['#7AE7BF', CalendarApp.EventColor.PALE_GREEN],
      ['#DBADFF', CalendarApp.EventColor.MAUVE], ['#FF887C', CalendarApp.EventColor.PALE_RED],
      ['#FBD75B', CalendarApp.EventColor.YELLOW], ['#FFB878', CalendarApp.EventColor.ORANGE],
      ['#46D6DB', CalendarApp.EventColor.CYAN], ['#E1E1E1', CalendarApp.EventColor.GRAY],
      ['#5484ED', CalendarApp.EventColor.BLUE], ['#51B749', CalendarApp.EventColor.GREEN],
      ['#DC2127', CalendarApp.EventColor.RED]
    ];
    function toRgb(h) { h = h.replace('#', ''); return [parseInt(h.slice(0,2),16), parseInt(h.slice(2,4),16), parseInt(h.slice(4,6),16)]; }
    var target = toRgb(hex);
    var best = palette[0][1], bestDist = Infinity;
    palette.forEach(function(p) {
      var c = toRgb(p[0]);
      var dist = Math.pow(c[0]-target[0],2) + Math.pow(c[1]-target[1],2) + Math.pow(c[2]-target[2],2);
      if (dist < bestDist) { bestDist = dist; best = p[1]; }
    });
    return best;
  }

  // Idempotent: creates events that don't exist yet, updates ones that
  // do (tracked in PlannerCalendarEvents), never duplicates on re-sync.
  function syncYearToCalendar(year) {
    ensureSheets();
    var cal = _getOrCreateCalendar();
    var occ = getOccurrencesForYear(year).occurrences;
    var mapSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAP_SHEET);
    var mapRows = Database.getAll(MAP_SHEET);
    var mapIndex = {};
    for (var i = 0; i < mapRows.length; i++) {
      mapIndex[String(mapRows[i][0]) + '|' + String(mapRows[i][1])] = { rowIndex: i + 2, eventId: String(mapRows[i][2]) };
    }
    var created = 0, updated = 0, failed = 0;
    occ.forEach(function(o) {
      var key = o.rule_id + '|' + o.date;
      var parts = o.date.split('-').map(Number);
      var d = new Date(parts[0], parts[1] - 1, parts[2]);
      try {
        var existing = mapIndex[key];
        if (existing) {
          var ev = cal.getEventById(existing.eventId);
          if (ev) { ev.setTitle(o.title); ev.setColor(_nearestCalendarColor(o.color)); updated++; return; }
        }
        var newEv = cal.createAllDayEvent(o.title, d, { description: o.notes || '' });
        newEv.setColor(_nearestCalendarColor(o.color));
        if (existing) {
          mapSheet.getRange(existing.rowIndex, 3).setValue(newEv.getId());
        } else {
          mapSheet.appendRow([o.rule_id, o.date, newEv.getId()]);
        }
        created++;
      } catch (e) { failed++; }
    });
    return { success: true, created: created, updated: updated, failed: failed, calendarId: cal.getId(), calendarName: cal.getName() };
  }

  function getCalendarInfo() {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('PLANNER_CALENDAR_ID');
    if (!id) return { exists: false };
    try {
      var cal = CalendarApp.getCalendarById(id);
      return { exists: true, id: id, name: cal.getName() };
    } catch (e) {
      return { exists: false };
    }
  }

  return {
    ensureSheets: ensureSheets,
    seedInitial: seedInitial,
    getAllRules: getAllRules,
    addRule: addRule,
    updateRule: updateRule,
    deleteRule: deleteRule,
    getOccurrencesForYear: getOccurrencesForYear,
    getDueReminders: getDueReminders,
    sendDuePlannerReminders: sendDuePlannerReminders,
    sendReminderNow: sendReminderNow,
    markWhatsappSent: markWhatsappSent,
    whatsappMessageFor: whatsappMessageFor,
    syncYearToCalendar: syncYearToCalendar,
    getCalendarInfo: getCalendarInfo
  };
})();

// Entry point for the daily time-driven trigger (Apps Script editor →
// Triggers → Add Trigger → sendDuePlannerReminders → Time-driven →
// Day timer). Kept as a bare global function so it's selectable there.
function sendDuePlannerReminders() {
  return YearPlannerService.sendDuePlannerReminders();
}

// Bare global wrapper purely so this is selectable in the Apps Script
// editor's Run dropdown — running THIS (not the email one above) is
// what actually triggers Google's Calendar permission prompt, since
// it's the only planner function that touches CalendarApp.
function testCalendarAuthorization() {
  var result = YearPlannerService.syncYearToCalendar(new Date().getFullYear());
  Logger.log(result);
  return result;
}
