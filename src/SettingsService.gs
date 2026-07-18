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
    getAllWithSchema: getAllWithSchema
  };
})();
