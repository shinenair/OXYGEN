// ============================================================
// CategoriesService.gs — Editable payment categories
//
// Two kinds of category, stored in one 'Categories' sheet:
//   • kind = 'credit'  → Fees RECEIVED by the association
//                        (Maintenance, Waste Management, LPG,
//                         Caution Deposit, Party Hall Rental, …)
//   • kind = 'debit'   → Payments MADE by the association
//                        (Care Taker, Sewage, Insurance, KWA, …)
//
// The user can add / edit / delete categories from the Settings page.
// The Bank Statement page uses them to classify each transaction, and
// the Fees Received page uses the credit categories for its tabs.
// ============================================================

var CategoriesService = (function() {
  var SHEET   = 'Categories';
  var HEADERS = ['category_id', 'kind', 'name', 'sort_order', 'active', 'created_at'];

  // Seed lists — written on first run so the app is usable immediately.
  var SEED_CREDIT = [
    'Maintenance', 'Waste Management', 'LPG',
    'Caution Deposit', 'Party Hall Rental', 'Miscellaneous'
  ];
  var SEED_DEBIT = [
    'Waste Mgmt', 'Misc. Purchase', 'Care Taker', 'Sewage', 'LPG Reading',
    'Insurance', 'House Keeping (HK)', 'Electrical Inspectorate', 'KWA',
    'Incinerator AMC', 'DG/Pump', 'Lift AMC', 'Cess/Property Tax',
    'Fire Renewal', 'Bye-Law', 'Tenants', 'Maintenance', 'LPG',
    'LPG Inventory', 'Reserve money to Caretaker', 'We Care (Salary)',
    'IOB Bank Charges', 'Facility Manager'
  ];

  function ensureSheet() {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      _seed(sheet);
    } else {
      var r1c1 = String(sheet.getRange(1, 1).getValue() || '');
      if (r1c1 === '') {
        sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
        sheet.getRange(1, 1, 1, HEADERS.length)
          .setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      }
      // Seed only if completely empty (header row exists but no data)
      if (sheet.getLastRow() < 2) _seed(sheet);
    }
    return sheet;
  }

  function _seed(sheet) {
    var now  = new Date().toISOString();
    var rows = [];
    var n    = 0;
    for (var i = 0; i < SEED_CREDIT.length; i++) {
      rows.push(['CAT_C' + (i + 1), 'credit', SEED_CREDIT[i], (i + 1) * 10, 'Yes', now]);
      n++;
    }
    for (var j = 0; j < SEED_DEBIT.length; j++) {
      rows.push(['CAT_D' + (j + 1), 'debit', SEED_DEBIT[j], (j + 1) * 10, 'Yes', now]);
      n++;
    }
    if (rows.length) sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
    SpreadsheetApp.flush();
    return n;
  }

  function _rows() {
    ensureSheet();
    return Database.getAll(SHEET);
  }

  function _toObj(r) {
    return {
      category_id: String(r[0] || ''),
      kind:        String(r[1] || ''),
      name:        String(r[2] || ''),
      sort_order:  Number(r[3] || 0),
      active:      String(r[4] || 'Yes'),
      created_at:  String(r[5] || '')
    };
  }

  // getAll(kind) — kind optional ('credit' | 'debit'); returns active first, sorted
  function getAll(kind) {
    var rows = _rows();
    var out  = [];
    for (var i = 0; i < rows.length; i++) {
      var o = _toObj(rows[i]);
      if (!o.category_id) continue;
      if (kind && o.kind !== kind) continue;
      out.push(o);
    }
    out.sort(function(a, b) { return a.sort_order - b.sort_order; });
    return out;
  }

  // Convenience: array of active names for a kind
  function names(kind) {
    var all = getAll(kind);
    var out = [];
    for (var i = 0; i < all.length; i++) {
      if (all[i].active !== 'No') out.push(all[i].name);
    }
    return out;
  }

  function add(kind, name) {
    if (!name || !String(name).trim()) throw new Error('Name is required');
    if (kind !== 'credit' && kind !== 'debit') throw new Error('Invalid category kind');
    var sheet = ensureSheet();
    var rows  = Database.getAll(SHEET);

    // Prevent duplicate name within the same kind
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][1]) === kind &&
          String(rows[i][2]).toLowerCase() === String(name).trim().toLowerCase()) {
        throw new Error('That category already exists');
      }
    }
    var maxOrder = 0;
    for (var j = 0; j < rows.length; j++) {
      if (String(rows[j][1]) === kind && Number(rows[j][3]) > maxOrder) maxOrder = Number(rows[j][3]);
    }
    var id = 'CAT_' + (kind === 'credit' ? 'C' : 'D') + new Date().getTime();
    sheet.appendRow([id, kind, String(name).trim(), maxOrder + 10, 'Yes', new Date().toISOString()]);
    SpreadsheetApp.flush();
    return { category_id: id };
  }

  function update(categoryId, data) {
    var sheet = ensureSheet();
    var rows  = Database.getAll(SHEET);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(categoryId)) {
        var r = rows[i].slice();
        while (r.length < HEADERS.length) r.push('');
        if (data.name   !== undefined && String(data.name).trim() !== '') r[2] = String(data.name).trim();
        if (data.active !== undefined) r[4] = data.active ? 'Yes' : 'No';
        if (data.sort_order !== undefined && !isNaN(Number(data.sort_order))) r[3] = Number(data.sort_order);
        sheet.getRange(i + 2, 1, 1, HEADERS.length).setValues([r]);
        SpreadsheetApp.flush();
        return { updated: true };
      }
    }
    throw new Error('Category not found');
  }

  function remove(categoryId) {
    var sheet = ensureSheet();
    var rows  = Database.getAll(SHEET);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(categoryId)) {
        sheet.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { deleted: true };
      }
    }
    throw new Error('Category not found');
  }

  function resetToDefaults() {
    var sheet = ensureSheet();
    var last  = sheet.getLastRow();
    if (last > 1) sheet.getRange(2, 1, last - 1, HEADERS.length).clearContent();
    var n = _seed(sheet);
    return { seeded: n };
  }

  return {
    ensureSheet:     ensureSheet,
    getAll:          getAll,
    names:           names,
    add:             add,
    update:          update,
    remove:          remove,
    resetToDefaults: resetToDefaults
  };
})();

