// ============================================================
// SettingsService.gs — App Settings Management
// ============================================================

var SettingsService = (function() {
  var SHEET = 'Settings';

  var SCHEMA = [
    // ── Community Forms ───────────────────────────────────────
    // Each Google Form's responses land in a linked spreadsheet (Form
    // editor → Responses → "View in Sheets"). Paste that spreadsheet's
    // URL here and OXYGEN reads it live — new responses appear on the
    // corresponding page as soon as it's refreshed.
    { key:'form_movein_responses_url',    group:'Community Forms', label:'Move-In/Out — responses spreadsheet URL',       type:'text', default:'', hint:'Form editor → Responses tab → View in Sheets → copy that spreadsheet\'s URL' },
    { key:'form_partyhall_responses_url', group:'Community Forms', label:'Party Hall Rental — responses spreadsheet URL', type:'text', default:'', hint:'Form editor → Responses tab → View in Sheets → copy that spreadsheet\'s URL' },
    { key:'form_vehicle_responses_url',   group:'Community Forms', label:'Vehicle Tracking — responses spreadsheet URL',  type:'text', default:'', hint:'Form editor → Responses tab → View in Sheets → copy that spreadsheet\'s URL' },

    // ── Documents ─────────────────────────────────────────────
    { key:'documents_folder_url', group:'Documents', label:'CDOA Documents — Google Drive folder URL', type:'text', default:'https://drive.google.com/drive/u/1/folders/1nPaTaNr5Yvyd3rxFYDACzvbvBZJO0KnH', hint:'The shared Drive folder holding the association\'s documents — listed live on the Documents page' },

    // ── Payment Screenshots ───────────────────────────────────
    { key:'screenshots_folder_url', group:'Payment Screenshots', label:'Payment confirmation screenshots — Drive folder URL', type:'text', default:'', hint:'The Drive folder where WhatsApp payment-confirmation screenshots are uploaded. Name each file with the flat (e.g. A101_15Jan.jpg) so the unit is captured.' },

    // ── Payment Fees ──────────────────────────────────────────
    { key:'fee_maintenance',      group:'Payment Fees',        label:'Current Maintenance Fee (₹)',     type:'number',   default:'2000',   hint:'The maintenance charge in effect now — used first when detecting payments' },
    { key:'fee_maintenance_history', group:'Payment Fees',     label:'Other Maintenance Fee amounts (₹, comma-separated)', type:'text', default:'1500', hint:'Older/alternate MF amounts, e.g. 1500. These and their multiples are also recognised in bank statements' },
    { key:'fee_waste',            group:'Payment Fees',        label:'Current Waste Management Fee (₹)', type:'number',   default:'170',    hint:'The waste management charge in effect now' },
    { key:'fee_waste_history',    group:'Payment Fees',        label:'Other Waste Mgmt Fee amounts (₹, comma-separated)', type:'text', default:'', hint:'Older/alternate WMF amounts. These and their multiples are also recognised' },
    { key:'lpg_conv_factor',   group:'Payment Fees', label:'LPG Conversion Factor (Kg per NM³) — fallback', type:'number', default:'2.6',   hint:'Used only for months with no rate history entry below' },
    { key:'lpg_price_per_kg',  group:'Payment Fees', label:'LPG Price per Kg (₹) — fallback',               type:'number', default:'78.31', hint:'Used only for months with no rate history entry below' },
    { key:'lpg_min_cylinders', group:'Payment Fees', label:'LPG Minimum Cylinder Stock (reorder threshold)', type:'number', default:'5', hint:'Net stock at or below this triggers "Order LPG Cylinders"' },
    { key:'fee_corpus',           group:'Payment Fees',        label:'Corpus Fund Amount (₹)',           type:'number',   default:'',       hint:'One-time or periodic corpus contribution' },
    { key:'fee_caution_deposit',  group:'Payment Fees',        label:'Caution Deposit Amount (₹)',      type:'number',   default:'1600',   hint:'Standard caution deposit for tenants' },
    { key:'fee_caution_deposit_history', group:'Payment Fees',  label:'Other Caution Deposit amounts (₹, comma-separated)', type:'text', default:'', hint:'Older/alternate caution deposit amounts — recognised in bank statements and offered when recording' },
    // ── Association Bank (reference only — for statement matching) ──
    // ── Association Details ───────────────────────────────────
    // ── Executive Committee ───────────────────────────────────
    // ── Building Details ──────────────────────────────────────
    // ── Notifications ─────────────────────────────────────────
  ];

  function ensureSheet() {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET);
      sheet.appendRow(['key','value','updated_at']);
      sheet.getRange(1,1,1,3)
        .setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    // Force the "value" column to plain text, permanently. Several
    // settings store date-like text ("2024-01" for the Caretaker ledger's
    // starting month, etc.) — without this, Sheets silently auto-detects
    // that as a real date and converts the cell, so reading it back later
    // produces a full JS Date object instead of the clean string that was
    // actually stored. Same root cause we've already fixed for Bank
    // Statement and LPG dates; this is the same fix applied here.
    sheet.getRange(2, 2, 500, 1).setNumberFormat('@');
    return sheet;
  }

  function getAll() {
    ensureSheet();
    var rows   = Database.getAll(SHEET);
    var stored = {};
    for (var i = 0; i < rows.length; i++) {
      stored[String(rows[i][0])] = String(rows[i][1] !== null && rows[i][1] !== undefined ? rows[i][1] : '');
    }
    var result = {};
    for (var i = 0; i < SCHEMA.length; i++) {
      var s = SCHEMA[i];
      result[s.key] = stored.hasOwnProperty(s.key) ? stored[s.key] : s.default;
    }
    return result;
  }

  function get(key) {
    var all = getAll();
    return all.hasOwnProperty(key) ? all[key] : '';
  }

  // The Universal Active Year — the single, whole-app "which year are we
  // working in" setting shown on the Dashboard. Everyone reads it; only
  // an Administrator can change it (enforced in ApiRouter, not here).
  // Defaults to the current calendar year the first time it's read.
  //
  // Deliberately reads the sheet directly rather than going through
  // get()/getAll() — those only return keys listed in SCHEMA (the fixed
  // set of fields the main Settings page renders), and 'active_year' is
  // not one of them. Going through get() silently returned '' every
  // time, no matter what was actually stored, so the app always fell
  // back to the current calendar year regardless of what was set.
  function getActiveYear() {
    ensureSheet();
    var rows = Database.getAll(SHEET);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === 'active_year') {
        var v = Number(rows[i][1]);
        return v || new Date().getFullYear();
      }
    }
    return new Date().getFullYear();
  }

  function setActiveYear(year) {
    var y = Number(year);
    if (!y || y < 2000 || y > 2100) throw new Error('Enter a valid year.');
    ensureSheet();
    var result = Database.findByColumn(SHEET, 0, 'active_year');
    var now = new Date().toISOString();
    if (result) {
      Database.updateRow(SHEET, result.rowIndex, ['active_year', String(y), now]);
    } else {
      Database.insert(SHEET, ['active_year', String(y), now]);
    }
    return { success: true, active_year: y };
  }

  // ── Effective-dated Fee Schedule ──────────────────────────────
  // Some fees change over time (Maintenance was ₹1,500 up to Feb 2025, then
  // ₹2,000 from Mar 2025). This is the same idea as LPG Rate History: each
  // entry says "from this month onward the fee is ₹X", and feeForMonth()
  // returns the amount that was in force for any given month — so past months
  // keep their old amount and never get retro-changed by a new rate.
  var FEE_SHEET = 'FeeSchedule';
  // Columns: schedule_id | fee_type | year | month | amount | set_by | updated_at
  var FS = { ID: 0, TYPE: 1, YEAR: 2, MONTH: 3, AMOUNT: 4, SET_BY: 5, UPDATED: 6 };

  function ensureFeeScheduleSheet() {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(FEE_SHEET);
    if (sheet) return sheet;
    sheet = ss.insertSheet(FEE_SHEET);
    sheet.appendRow(['schedule_id','fee_type','year','month','amount','set_by','updated_at']);
    sheet.getRange(1, 1, 1, 7)
      .setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    // Seed the known Maintenance history the first time the sheet is created:
    // ₹1,500 from the start, ₹2,000 from Mar 2025 onward.
    var now = new Date().toISOString();
    sheet.appendRow([Database.generateId('FSC'), 'Maintenance', 2020, 1, 1500, 'System', now]);
    sheet.appendRow([Database.generateId('FSC'), 'Maintenance', 2025, 3, 2000, 'System', now]);
    return sheet;
  }

  function listFeeSchedule(feeType) {
    ensureFeeScheduleSheet();
    var rows = Database.getAll(FEE_SHEET);
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      if (feeType && String(rows[i][FS.TYPE]) !== feeType) continue;
      out.push({
        schedule_id: String(rows[i][FS.ID]),
        fee_type:    String(rows[i][FS.TYPE]),
        year:        Number(rows[i][FS.YEAR]),
        month:       Number(rows[i][FS.MONTH]),
        amount:      Number(rows[i][FS.AMOUNT]),
        set_by:      String(rows[i][FS.SET_BY] || '')
      });
    }
    out.sort(function(a, b) { return (a.year * 12 + a.month) - (b.year * 12 + b.month); });
    return out;
  }

  // Upsert one "from this month, fee = amount" entry for a fee type.
  function setFeeSchedule(data) {
    var type = String(data.fee_type || '').trim();
    var year = Number(data.year), month = Number(data.month), amount = Number(data.amount);
    if (!type) throw new Error('Fee type is required.');
    if (!year || !month || month < 1 || month > 12) throw new Error('Valid year and month are required.');
    if (isNaN(amount) || amount < 0) throw new Error('Valid amount is required.');
    var sheet = ensureFeeScheduleSheet();
    var rows = Database.getAll(FEE_SHEET);
    var now = new Date().toISOString();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][FS.TYPE]) === type && Number(rows[i][FS.YEAR]) === year && Number(rows[i][FS.MONTH]) === month) {
        sheet.getRange(i + 2, FS.AMOUNT + 1).setValue(amount);
        sheet.getRange(i + 2, FS.SET_BY + 1).setValue(data.set_by || 'Manager');
        sheet.getRange(i + 2, FS.UPDATED + 1).setValue(now);
        return { success: true, updated: true };
      }
    }
    sheet.appendRow([Database.generateId('FSC'), type, year, month, amount, data.set_by || 'Manager', now]);
    return { success: true, created: true };
  }

  function deleteFeeSchedule(scheduleId) {
    var sheet = ensureFeeScheduleSheet();
    var rows = Database.getAll(FEE_SHEET);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][FS.ID]) === String(scheduleId)) { sheet.deleteRow(i + 2); return { success: true }; }
    }
    throw new Error('Schedule entry not found.');
  }

  // The fee amount in force for a fee type in a given 'YYYY-MM'. Picks the
  // entry with the latest effective-from that is still on or before the
  // target month. Falls back to the current fee_<type> setting, then 0.
  function feeForMonth(feeType, monthKey) {
    var m = String(monthKey || '').match(/^(\d{4})-(\d{1,2})/);
    var targetY = m ? Number(m[1]) : (new Date()).getFullYear();
    var targetM = m ? Number(m[2]) : ((new Date()).getMonth() + 1);
    var target = targetY * 12 + targetM;
    var list = listFeeSchedule(feeType);
    var best = null, bestKey = -1;
    for (var i = 0; i < list.length; i++) {
      var key = list[i].year * 12 + list[i].month;
      if (key <= target && key > bestKey) { bestKey = key; best = list[i].amount; }
    }
    if (best !== null) return best;
    // No schedule entry on/before the target month — fall back to the plain
    // current-fee setting for that type, if there is one.
    var fallbackKey = feeType === 'Maintenance' ? 'fee_maintenance'
                    : feeType === 'Waste Management' ? 'fee_waste'
                    : feeType === 'Caution Deposit' ? 'fee_caution_deposit' : '';
    var fv = fallbackKey ? Number(get(fallbackKey)) : 0;
    return isNaN(fv) ? 0 : fv;
  }

  function set(key, value) {
    ensureSheet();
    var result = Database.findByColumn(SHEET, 0, key);
    var now    = new Date().toISOString();
    if (result) {
      Database.updateRow(SHEET, result.rowIndex, [key, value, now]);
    } else {
      Database.insert(SHEET, [key, value, now]);
    }
    return true;
  }

  function saveAll(data) {
    var keys = Object.keys(data);
    for (var i = 0; i < keys.length; i++) {
      set(keys[i], data[keys[i]]);
    }
    return { success: true, saved: keys.length };
  }

  function getAllWithSchema() {
    var values = getAll();
    var result = [];
    for (var i = 0; i < SCHEMA.length; i++) {
      var s = SCHEMA[i];
      result.push({
        key:     s.key,
        group:   s.group,
        label:   s.label,
        type:    s.type,
        hint:    s.hint,
        value:   values.hasOwnProperty(s.key) ? values[s.key] : s.default,
        default: s.default
      });
    }
    return result;
  }

  return {
    ensureSheet:     ensureSheet,
    getAll:          getAll,
    get:             get,
    set:             set,
    getActiveYear:   getActiveYear,
    setActiveYear:   setActiveYear,
    saveAll:         saveAll,
    getAllWithSchema: getAllWithSchema,
    listFeeSchedule:  listFeeSchedule,
    setFeeSchedule:   setFeeSchedule,
    deleteFeeSchedule: deleteFeeSchedule,
    feeForMonth:      feeForMonth
  };
})();
