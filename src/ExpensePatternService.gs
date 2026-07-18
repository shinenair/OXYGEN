// ═══════════════════════════════════════════════════════════════
// ExpensePatternService.gs — recurring-payment patterns for
// auto-categorizing Transaction Ledger debits by narration text,
// so repetitive monthly expenses (security, housekeeping, etc.)
// get their Expense Category assigned automatically on import.
// ═══════════════════════════════════════════════════════════════
var ExpensePatternService = (function() {

  var SHEET = 'ExpensePatterns';
  var HEADERS = ['pattern_id', 'pattern_text', 'amount', 'expense_category', 'recorded_at'];
  var C = { ID: 0, PATTERN: 1, AMOUNT: 2, CATEGORY: 3, RECORDED_AT: 4 };

  function ensureSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET);
    if (!sh) {
      sh = ss.insertSheet(SHEET);
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      return;
    }

    // Migrate an existing sheet from the OLD 4-column layout
    // (pattern_id, pattern_text, expense_category, recorded_at) to the
    // current 5-column layout that has 'amount' inserted before
    // expense_category. Without this, an existing sheet keeps its old
    // column positions while this code reads/writes using the NEW
    // indices — every existing row appears garbled (category read as
    // amount, recorded_at read as category), and every new row appended
    // with 5 values lands misaligned against a 4-column header.
    var headerRow = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    var hasAmountCol = false;
    for (var i = 0; i < headerRow.length; i++) {
      if (String(headerRow[i]).trim().toLowerCase() === 'amount') { hasAmountCol = true; break; }
    }
    if (!hasAmountCol) {
      // Insert a new blank column at position 3 (C) — Sheets automatically
      // shifts expense_category/recorded_at (and every row's data in them)
      // one column to the right, preserving each row's existing values.
      sh.insertColumnAfter(2);
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
    }
  }

  function getAll() {
    return Database.getAll(SHEET).map(function(r) {
      return { pattern_id: String(r[C.ID]), pattern_text: String(r[C.PATTERN] || ''),
               amount: r[C.AMOUNT] === '' || r[C.AMOUNT] === undefined || r[C.AMOUNT] === null ? null : Number(r[C.AMOUNT]),
               expense_category: String(r[C.CATEGORY] || ''), recorded_at: String(r[C.RECORDED_AT] || '') };
    }).sort(function(a, b) { return a.pattern_text.localeCompare(b.pattern_text); });
  }

  // A pattern can specify narration text, an exact amount, or both. Both
  // conditions present on one pattern must BOTH be true to match (more
  // precise — useful when the same amount could come from more than one
  // vendor). Either alone is also valid: an amount-only pattern catches a
  // fixed-cost recurring payment (e.g. a cylinder delivery that's always
  // exactly ₹9,120) even when the narration's reference number is random
  // every time; a text-only pattern catches a vendor whose amount varies
  // month to month (KSEB electricity, KWA water) but whose narration
  // reliably contains the same recognizable text.
  function addPattern(patternText, amount, category) {
    patternText = String(patternText || '').trim();
    var amt = (amount === '' || amount === undefined || amount === null) ? null : Number(amount);
    category = String(category || '').trim();
    if (!patternText && amt === null) throw new Error('Enter narration text, an amount, or both.');
    if (!category) throw new Error('Expense category is required.');
    var sheet = Database.getSheet(SHEET);
    sheet.appendRow([Utilities.getUuid(), patternText, amt === null ? '' : amt, category, new Date().toISOString()]);
    return { success: true };
  }

  function deletePattern(id) {
    var sheet = Database.getSheet(SHEET);
    var rows = Database.getAll(SHEET);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.ID]) === String(id)) { sheet.deleteRow(i + 2); return { success: true }; }
    }
    throw new Error('Pattern not found.');
  }

  // Checks a narration + amount against every pattern, in stored order,
  // and returns the FIRST matching category, or '' if none match.
  function matchCategory(narration, amount, patterns) {
    var n = String(narration || '').toUpperCase();
    var amt = (amount === '' || amount === undefined || amount === null) ? null : Number(amount);
    var list = patterns || getAll();
    for (var i = 0; i < list.length; i++) {
      var p = list[i];
      var textOk = true, amountOk = true;
      var pText = String(p.pattern_text || '').toUpperCase().trim();
      if (pText) textOk = n.indexOf(pText) > -1;
      if (p.amount !== null && p.amount !== undefined && p.amount !== '') {
        amountOk = amt !== null && Math.abs(amt - Number(p.amount)) < 0.005;
      }
      if (!pText && (p.amount === null || p.amount === undefined || p.amount === '')) continue; // empty rule, skip
      if (textOk && amountOk) return p.expense_category;
    }
    return '';
  }

  return {
    ensureSheet:   ensureSheet,
    getAll:        getAll,
    addPattern:    addPattern,
    deletePattern: deletePattern,
    matchCategory: matchCategory
  };
})();
