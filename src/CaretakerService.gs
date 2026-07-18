// ═══════════════════════════════════════════════════════════════
// CaretakerService.gs — the Caretaker's petty-cash (imprest account)
// ledger: small cash purchases (bulbs, stationery, minor repairs)
// that never touch the bank. "Reserve Money" transfers INTO this
// fund are pulled automatically from the Bank Statement (the
// existing "Reserve money to Caretaker" expense category); the
// Caretaker's own spending is logged here manually, since it has
// no bank record at all.
// ═══════════════════════════════════════════════════════════════
var CaretakerService = (function() {

  var SHEET = 'CaretakerExpenses';
  var HEADERS = ['expense_id', 'date', 'voucher_no', 'description', 'amount', 'recorded_by', 'recorded_at'];
  var C = { ID: 0, DATE: 1, VOUCHER: 2, DESC: 3, AMOUNT: 4, RECORDED_BY: 5, RECORDED_AT: 6 };

  // Where the imprest chain starts — a one-time seed for the very first
  // month OXYGEN tracks (e.g. "₹9,516 as of 31-Dec-2023"), stored as a
  // Settings key so it's a single, clearly-labelled source of truth
  // rather than a synthetic transaction. Every month after this one
  // computes its own opening balance from the previous month's closing.
  var SEED_KEY = 'caretaker_opening_balance';
  var SEED_MONTH_KEY = 'caretaker_opening_month'; // yyyy-MM the seed applies to (the month BEFORE tracking starts)

  function ensureSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET);
    if (!sh) {
      sh = ss.insertSheet(SHEET);
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    // Dates stay plain text permanently — same reasoning as every other
    // date column in this project (Sheets silently auto-converting a
    // date-looking string has broken reconciliation-style math before).
    sh.getRange(2, C.DATE + 1, 3000, 1).setNumberFormat('@');
  }

  function getAll() {
    return Database.getAll(SHEET).map(function(r) {
      return { expense_id: String(r[C.ID]), date: String(r[C.DATE] || ''),
               voucher_no: String(r[C.VOUCHER] || ''), description: String(r[C.DESC] || ''),
               amount: Number(r[C.AMOUNT]) || 0, recorded_by: String(r[C.RECORDED_BY] || '') };
    }).sort(function(a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
  }

  function getByMonth(year, month) {
    var key = year + '-' + (month < 10 ? '0' + month : month);
    return getAll().filter(function(e) { return e.date.slice(0, 7) === key; });
  }

  function addExpense(data) {
    if (!data.date) throw new Error('Date is required.');
    var amt = Number(data.amount);
    if (isNaN(amt) || amt <= 0) throw new Error('Amount must be a positive number.');
    if (!String(data.description || '').trim()) throw new Error('Description is required.');
    // Bulk imports (an admin-only action already gated at the router
    // level) skip the interactive per-row confirmation — asking to
    // confirm 11+ separate rows one at a time would be absurd. A single
    // manual add via the Reports page still goes through the normal check.
    if (!data.skipYearCheck) {
      var recordYear = Number(data.date.slice(0, 4));
      checkYearEditable(recordYear, data.confirmed_historical);
    }
    ensureSheet();
    var sheet = Database.getSheet(SHEET);
    var id = Utilities.getUuid();
    sheet.appendRow([id, data.date, data.voucher_no || '', data.description.trim(), amt,
                      data.recorded_by || 'Manager', new Date().toISOString()]);
    return { success: true, expense_id: id };
  }

  function updateExpense(id, data) {
    var sheet = Database.getSheet(SHEET);
    var rows = Database.getAll(SHEET);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.ID]) !== String(id)) continue;
      var oldYear = Number(String(rows[i][C.DATE]).slice(0, 4));
      var newYear = Number(String(data.date || rows[i][C.DATE]).slice(0, 4));
      // A genuine year change (moving the entry into or out of a
      // historical year) needs the confirm — checking both the old and
      // new year covers editing either direction.
      checkYearEditable(oldYear, data.confirmed_historical);
      checkYearEditable(newYear, data.confirmed_historical);
      var amt = Number(data.amount);
      if (!data.date) throw new Error('Date is required.');
      if (isNaN(amt) || amt <= 0) throw new Error('Amount must be a positive number.');
      if (!String(data.description || '').trim()) throw new Error('Description is required.');
      var row = rows[i];
      row[C.DATE] = data.date; row[C.VOUCHER] = data.voucher_no || '';
      row[C.DESC] = data.description.trim(); row[C.AMOUNT] = amt;
      sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
      return { success: true };
    }
    throw new Error('Expense not found.');
  }

  function deleteExpense(id, confirmedHistorical) {
    var sheet = Database.getSheet(SHEET);
    var rows = Database.getAll(SHEET);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.ID]) === String(id)) {
        var recordYear = Number(String(rows[i][C.DATE]).slice(0, 4));
        checkYearEditable(recordYear, confirmedHistorical);
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    throw new Error('Expense not found.');
  }

  function getMonthSummary() {
    var counts = {};
    getAll().forEach(function(e) {
      var k = e.date.slice(0, 7);
      if (!k) return;
      counts[k] = (counts[k] || 0) + 1;
    });
    var out = []; for (var k in counts) out.push({ month: k, count: counts[k] });
    out.sort(function(a, b) { return a.month < b.month ? -1 : 1; });
    return out;
  }

  function deleteByMonth(monthKey) {
    var sheet = Database.getSheet(SHEET);
    var rows = Database.getAll(SHEET);
    var deleted = 0;
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][C.DATE]).slice(0, 7) === monthKey) { sheet.deleteRow(i + 2); deleted++; }
    }
    return { success: true, deleted: deleted };
  }

  function setSeed(openingBalance, asOfMonthKey) {
    var bal = Number(openingBalance);
    if (isNaN(bal) || bal < 0) throw new Error('Enter a valid opening balance.');
    if (!/^\d{4}-\d{2}$/.test(asOfMonthKey)) throw new Error('Enter a valid starting month.');
    SettingsService.ensureSheet(); // makes sure the plain-text column format is applied first
    SettingsService.set(SEED_KEY, String(bal));
    // Leading apostrophe forces plain text even if the column format
    // somehow hasn't taken effect yet — the same proven technique used
    // for Bank Statement date corrections elsewhere in this project.
    SettingsService.set(SEED_MONTH_KEY, "'" + asOfMonthKey);
    return { success: true };
  }

  function getSeed() {
    // getAll() only returns SCHEMA-listed keys (see the Active Year fix) —
    // these two are not in SCHEMA either, so read the sheet directly.
    var bal = null, monthKey = '';
    var rows = Database.getAll('Settings');
    rows.forEach(function(r) {
      if (String(r[0]) === SEED_KEY) bal = Number(r[1]);
      if (String(r[0]) === SEED_MONTH_KEY) {
        var v = r[1];
        if (v instanceof Date) {
          // Already-corrupted from before the plain-text fix — Sheets
          // silently turned the stored "2024-01" into a real Date object.
          // Recover the intended yyyy-MM from its own LOCAL calendar
          // fields (never String(dateObject), which produces the long
          // "Mon Jan 01 2024 00:00:00 GMT+0530..." form instead).
          monthKey = v.getFullYear() + '-' + _pad2(v.getMonth() + 1);
        } else {
          monthKey = String(v || '');
        }
      }
    });
    return { openingBalance: bal, asOfMonth: monthKey };
  }

  // Reserve Money transfers FROM the association TO the caretaker, for one
  // month — pulled directly from the Bank Statement's existing "Reserve
  // money to Caretaker" expense category, so this never needs manual
  // re-entry; it's the same transaction the Treasurer's report already
  // counts as an association expense.
  // Accepts an OPTIONAL pre-fetched array (avoids a re-read on every call
  // — critical when this runs inside getMonthReport()'s month-chain loop).
  function getReserveMoney(year, month, preloadedRows) {
    var key = year + '-' + (month < 10 ? '0' + month : month);
    var rows = preloadedRows || BankService.getDebitsByCategory('Reserve money to Caretaker');
    var out = [];
    rows.forEach(function(r) {
      if (String(r.date || '').indexOf(key) > -1) {
        // Prefer the human-written Note (set from the Bank Statement page)
        // over the raw bank narration, which is usually a cryptic
        // reference string — falls back to narration only if no note
        // has been added for that transaction.
        out.push({ date: r.date, description: (r.note && r.note.trim()) ? r.note.trim() : r.narration, amount: r.debit });
      }
    });
    return out;
  }

  // Full month picture: opening balance (chained from the previous month,
  // or the one-time seed if this IS the seed month's successor), reserve
  // money received, itemized expenses, and the resulting closing balance.
  function getMonthReport(year, month) {
    year = Number(year); month = Number(month);
    if (!year || !month || month < 1 || month > 12) {
      throw new Error('getMonthReport: invalid year/month (got year=' + year + ', month=' + month + ').');
    }

    // Fetch ONCE — the chain below may need to look back across many
    // months (from the seed month up to whichever month was requested).
    // Deliberately uses the lightweight, UNSORTED reader here (not
    // getAllTransactions), since this ledger only ever groups by month —
    // it never needs chronological order, so there's no reason to pay for
    // the O(n log n) sort getAllTransactions() always performs.
    var reserveRows = BankService.getDebitsByCategory('Reserve money to Caretaker');
    var allExpenses = getAll();
    function reserveFor(y, m) { return getReserveMoney(y, m, reserveRows); }
    function expensesFor(y, m) {
      var key = y + '-' + (m < 10 ? '0' + m : m);
      return allExpenses.filter(function(e) { return e.date.slice(0, 7) === key; });
    }

    var seed = getSeed();
    var opening = 0;
    if (seed.asOfMonth) {
      var seedParts = seed.asOfMonth.split('-');
      var seedY = Number(seedParts[0]), seedM = Number(seedParts[1]);
      if (!seedY || !seedM || seedM < 1 || seedM > 12) {
        throw new Error('getMonthReport: the stored starting month ("' + seed.asOfMonth + '") is not a valid yyyy-MM value — re-set it from the Reports page.');
      }
      // Walk forward from the seed month up to (but not including) the
      // requested month, accumulating each in-between month's
      // (reserve money − expenses). The requested month's own figures
      // are applied separately below, exactly once.
      //
      // Hard safety cap: this loop MUST terminate in, at most, a few
      // hundred steps for any sane seed/target combination (decades
      // apart at most). If it ever needs more than that, something is
      // wrong with the stored data (e.g. a corrupted seed year like
      // "0243" instead of "2024") rather than a real gap to walk — in
      // that case, fail fast with a clear, specific error instead of
      // spinning indefinitely.
      opening = seed.openingBalance || 0;
      var cursorY = seedY, cursorM = seedM;
      var guard = 0, GUARD_MAX = 1200; // 100 years — no legitimate case needs more
      var chainSteps = [];
      while (true) {
        guard++;
        if (guard > GUARD_MAX) {
          throw new Error('getMonthReport: the month chain did not terminate after ' + GUARD_MAX +
            ' steps (seed=' + seed.asOfMonth + ', target=' + year + '-' + month +
            ', currently at ' + cursorY + '-' + cursorM + '). This usually means the stored starting ' +
            'month or year is corrupted — re-set the starting balance from the Reports page.');
        }
        var nextM = cursorM + 1, nextY = cursorY;
        if (nextM > 12) { nextM = 1; nextY++; }
        if (nextY > year || (nextY === year && nextM >= month)) break;
        cursorM = nextM; cursorY = nextY;
        var stepReserveRows = reserveFor(cursorY, cursorM);
        var stepExpenseRows = expensesFor(cursorY, cursorM);
        var rm = stepReserveRows.reduce(function(s, r) { return s + r.amount; }, 0);
        var exp = stepExpenseRows.reduce(function(s, e) { return s + e.amount; }, 0);
        opening = opening + rm - exp;
        chainSteps.push({
          year: cursorY, month: cursorM,
          reserveCount: stepReserveRows.length, reserveTotal: Math.round(rm * 100) / 100,
          expenseCount: stepExpenseRows.length, expenseTotal: Math.round(exp * 100) / 100,
          runningOpening: Math.round(opening * 100) / 100
        });
      }
    }
    var reserveMoney = reserveFor(year, month);
    var expenses = expensesFor(year, month);
    var reserveTotal = reserveMoney.reduce(function(s, r) { return s + r.amount; }, 0);
    var expenseTotal = expenses.reduce(function(s, e) { return s + e.amount; }, 0);
    var closing = opening + reserveTotal - expenseTotal;
    return {
      year: year, month: month, opening: Math.round(opening * 100) / 100,
      reserveMoney: reserveMoney, reserveTotal: Math.round(reserveTotal * 100) / 100,
      expenses: expenses, expenseTotal: Math.round(expenseTotal * 100) / 100,
      closing: Math.round(closing * 100) / 100,
      seedUsed: seed, // for verification — the exact starting balance/month this calculation chained from
      chainSteps: chainSteps || [] // every month walked through to reach this opening balance, with counts
    };
  }

  // Imports the "MANAGER'S EXPENSES" / "CARETAKER'S EXPENSES" table from
  // the association's own Caretaker's Monthly Report template. Finds the
  // section by scanning for its header row (containing "Voucher") rather
  // than trusting a fixed row number, since different months' files may
  // have the Reserve Money section above it at a different length.
  //
  // The file's own date cells cannot be trusted for year/month — a real
  // example had "2026" baked into a report titled "January 2024" — so
  // only the DAY is read from each row; the year and month always come
  // from what the user explicitly selects for this import.
  function importExpenses(rows, year, month) {
    year = Number(year); month = Number(month);
    if (!year || !month || month < 1 || month > 12) throw new Error('A valid year and month are required.');

    // Find the header row for the EXPENSES table specifically. Both the
    // "Reserve Money" table and the "Manager's Expenses" table share the
    // exact same column headers ("#, Date, Voucher No., Description,
    // Amount"), so matching on that text alone would latch onto whichever
    // one comes first (Reserve Money) — wrong table entirely. Instead,
    // require the row to be preceded within a few rows by a section title
    // containing "expense", which only the correct table has above it.
    var headerRow = -1;
    for (var i = 0; i < rows.length; i++) {
      var joined = rows[i].join(' ').toLowerCase();
      if (joined.indexOf('voucher') === -1 || joined.indexOf('description') === -1) continue;
      var precededByExpenseTitle = false;
      for (var back = 1; back <= 3 && i - back >= 0; back++) {
        if (rows[i - back].join(' ').toLowerCase().indexOf('expense') > -1) { precededByExpenseTitle = true; break; }
      }
      if (precededByExpenseTitle) { headerRow = i; break; }
    }
    if (headerRow === -1) {
      throw new Error('Could not find the Expenses table (a "Voucher No. / Description" header preceded by an "Expenses" section title) — is this the Caretaker\'s Monthly Report file?');
    }

    var imported = 0, errors = [], warnings = [];
    for (var r = headerRow + 1; r < rows.length; r++) {
      var row = rows[r];
      if (!row || !row.length) continue;
      var seqRaw = String(row[0] || '').trim();
      if (!seqRaw) continue; // genuinely blank row — normal, not worth mentioning
      if (!/^\d+$/.test(seqRaw)) {
        // Has SOME value in the "#" column, but it's not a clean integer
        // (a "TOTAL EXPENSES" label row is expected here — but if this
        // was actually meant to be a numbered data row, it would
        // otherwise vanish with no trace, so flag it just in case).
        if (seqRaw.toUpperCase().indexOf('TOTAL') === -1) {
          warnings.push('Row skipped — "#" column contained ' + JSON.stringify(seqRaw) + ', not a plain number.');
        }
        continue;
      }

      var dateRaw = row[1];
      var description = String(row[6] || '').trim();
      // Amount is the last non-empty cell in the row (trailing columns
      // are blank spacer columns in this template).
      var amountRaw = '';
      for (var c = row.length - 1; c >= 0; c--) {
        if (row[c] !== '' && row[c] !== null && row[c] !== undefined) { amountRaw = row[c]; break; }
      }
      var amount = Number(String(amountRaw).replace(/[₹,\s]/g, ''));

      if (!description || isNaN(amount) || amount <= 0) {
        errors.push('Row #' + seqRaw + ': skipped — description was ' + JSON.stringify(description) +
          ', amount raw value was ' + JSON.stringify(amountRaw) + ' (parsed as ' + amount + ').');
        continue;
      }

      var parsedDate = parseFlexibleDate(dateRaw);
      var day = parsedDate ? Number(parsedDate.split('-')[2]) : NaN;
      if (!day) {
        // Fallback for a cell displayed as just "D-Mon" (no year at all,
        // e.g. "1-Feb") — this template's date column is formatted that
        // way, and since only the day is ever used here anyway (year and
        // month always come from what's selected for this import), this
        // is a completely reliable source, not a guess.
        var dm = String(dateRaw || '').trim().match(/^(\d{1,2})[-\/ ]([A-Za-z]{3})[A-Za-z]*\.?$/);
        if (dm) day = parseInt(dm[1], 10);
      }
      if (!day || day < 1 || day > 31) {
        errors.push('Row #' + seqRaw + ' ("' + description + '"): could not read a day-of-month from its date — raw value was: ' + JSON.stringify(dateRaw) + ' (type: ' + (typeof dateRaw) + ')');
        continue;
      }
      // Cap at the actual number of days in the target month (handles a
      // stray "31" landing in a 30-day month, etc.) rather than failing.
      var daysInMonth = new Date(year, month, 0).getDate();
      if (day > daysInMonth) { warnings.push('Row #' + seqRaw + ': day ' + day + ' does not exist in the selected month — used day ' + daysInMonth + ' instead.'); day = daysInMonth; }
      var date = year + '-' + _pad2(month) + '-' + _pad2(day);

      var voucherNo = String(row[3] || '').trim();
      try {
        addExpense({ date: date, voucher_no: voucherNo, description: description, amount: amount,
                     recorded_by: 'Import', skipYearCheck: true });
        imported++;
      } catch (e) { errors.push('Row #' + seqRaw + ' ("' + description + '"): ' + e.message); }
    }

    if (imported === 0 && !errors.length) {
      throw new Error('No expense rows found under the header row — check that this file matches the Caretaker\'s Monthly Report template.');
    }
    return { success: true, imported: imported, errors: errors, warnings: warnings };
  }

  return {
    ensureSheet:      ensureSheet,
    getAll:           getAll,
    getByMonth:        getByMonth,
    addExpense:       addExpense,
    importExpenses:   importExpenses,
    updateExpense:    updateExpense,
    deleteExpense:    deleteExpense,
    getMonthSummary:  getMonthSummary,
    deleteByMonth:    deleteByMonth,
    setSeed:          setSeed,
    getSeed:          getSeed,
    getReserveMoney:  getReserveMoney,
    getMonthReport:   getMonthReport
  };
})();
