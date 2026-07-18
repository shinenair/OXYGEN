// ═══════════════════════════════════════════════════════════════
// ReportService.gs — assembles the exact data each Association
// report needs, from data OXYGEN already has (Bank Statement,
// Payments ledger, Caretaker ledger). Pure data assembly only —
// PDF rendering/styling lives in ReportPdfService.gs.
// ═══════════════════════════════════════════════════════════════
var ReportService = (function() {

  var STANDARD_INCOME_TYPES = ['Maintenance', 'Waste Management', 'LPG', 'Caution Deposit', 'Party Hall Rental'];
  var INCOME_LABELS = {
    'Maintenance': 'Maintenance Fee (MF)',
    'Waste Management': 'CREDAI',
    'LPG': 'LPG',
    'Caution Deposit': 'Caution Deposit',
    'Party Hall Rental': 'Party Hall Rental'
  };

  function _monthKey(year, month) { return year + '-' + _pad2(month); }
  function _monthEndDisplay(year, month) {
    var mn = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var lastDay = new Date(year, month, 0).getDate();
    return (lastDay < 10 ? '0' + lastDay : String(lastDay)) + ' ' + mn[month] + ' ' + year;
  }

  // Treasurer's Monthly Report — a full month's cash flow: opening bank
  // balance, income grouped by fee type (plus any one-off miscellaneous
  // credits listed individually), every individual expense debit, and
  // the resulting closing balance.
  function getTreasurerReportData(year, month) {
    year = Number(year); month = Number(month);
    var key = _monthKey(year, month);
    var prevY = month === 1 ? year - 1 : year, prevM = month === 1 ? 12 : month - 1;

    // Every transaction for this month, fetched once — used for
    // opening/closing balance derivation, Income, and Expenses below.
    var txns = BankService.getAllTransactions({ month: key });
    var txns2 = [];
    try { txns2 = Bank2Service.getAllTransactions({ month: key }); } catch (e0) {}

    // Opening/closing balance — each account's own balance is only
    // meaningful within ITS OWN continuous ledger (their stated balances
    // have nothing to do with each other), so each is derived separately
    // exactly as before, and the two are then ADDED to produce the
    // combined figure the Treasurer actually wants to see.
    var firstBalTxn = null, lastBalTxn = null;
    txns.forEach(function(t) {
      var bal = (t.balance === '' || t.balance === null || t.balance === undefined) ? NaN : Number(String(t.balance).replace(/[₹,\s]/g, ''));
      if (isNaN(bal)) return;
      if (!firstBalTxn) firstBalTxn = t;
      lastBalTxn = t; // txns is already chronologically sorted by getAllTransactions()
    });
    var openingMain = 0, closingMain = 0;
    if (firstBalTxn) {
      var fBal = Number(String(firstBalTxn.balance).replace(/[₹,\s]/g, ''));
      var fCr = Number(firstBalTxn.credit || 0), fDb = Number(firstBalTxn.debit || 0);
      openingMain = fBal - fCr + fDb;
    }
    if (lastBalTxn) {
      closingMain = Number(String(lastBalTxn.balance).replace(/[₹,\s]/g, ''));
    }

    var account2 = null;
    try { account2 = _account2Summary(year, month); } catch (e2) { account2 = null; }
    var openingAcct2 = (account2 && account2.available) ? account2.openingBalance : 0;
    var closingAcct2 = (account2 && account2.available) ? account2.closingBalance : 0;
    var opening = openingMain + openingAcct2;
    var closing = closingMain + closingAcct2;

    // Income — grouped DIRECTLY from Bank Statement credits by each
    // transaction's own Type (payment_type), not from the Payments
    // ledger. This matters: the Payments ledger only has a row when a
    // credit was successfully auto-matched AND posted, which depends on
    // owner/tenant name-matching succeeding — a transaction's Type on
    // the Bank Statement itself is set independently of that, so this
    // is the more complete, reliable source, consistent with how the
    // rest of this report already reads opening/closing balance and
    // expenses straight from the bank data too.
    //
    // EXCEPTION: a transaction that's been Split (combined MF+WMF+LPG in
    // one payment, a skipped month caught up with the next one, etc.)
    // no longer has one single Type on the bank row — it's tagged
    // "Split (N)" instead. For those, the actual per-type amounts live
    // in the Payments ledger (created by the split itself), so they're
    // looked up there instead of read from the bank row directly. Built
    // as a single map up front — fetching Payments once for the whole
    // month, not once per split transaction — since splits could be
    // common enough that repeating this per-row would add up.
    var splitPaymentsByTxn = {};
    var allPayments = PaymentsService.getAllPayments(null); // NOT filtered by month — a split line
    // (e.g. a skipped month caught up with the next one) can legitimately
    // be tagged with a DIFFERENT month than the bank transaction's own
    // date, so filtering by this report's month here would silently
    // miss exactly the case this feature exists for.
    allPayments.forEach(function(p) {
      var m = String(p.notes || '').match(/^BANK:(\S+)\s+\(split/);
      if (!m) return;
      var tid = m[1];
      if (!splitPaymentsByTxn[tid]) splitPaymentsByTxn[tid] = [];
      splitPaymentsByTxn[tid].push(p);
    });

    // Income and Expenses are now accumulated across BOTH accounts into
    // the SAME groups — one combined report, not two stacked ones. Every
    // AGGREGATED line (the five standard fee types) shows one combined
    // total, since that's the whole point of merging. Every INDIVIDUAL
    // line (Bank Interest, Miscellaneous, uncategorised items — never
    // aggregated to begin with) keeps a small account tag, so a mistaken
    // payment or interest credit from the IOB LPG account is still
    // traceable back to it even inside the merged view.
    var incomeGroups = {}; // type -> total
    var miscRows = [];
    function collectIncome(list, acctLabel) {
      list.forEach(function(t) {
        var credit = Number(t.credit || 0);
        if (!(credit > 0)) return;
        var isSplit = !!(splitPaymentsByTxn[t.txn_id] && splitPaymentsByTxn[t.txn_id].length);
        if (isSplit) {
          splitPaymentsByTxn[t.txn_id].forEach(function(p) {
            var amt = Number(p.amount || 0);
            if (STANDARD_INCOME_TYPES.indexOf(p.payment_type) > -1) {
              incomeGroups[p.payment_type] = (incomeGroups[p.payment_type] || 0) + amt;
            } else {
              miscRows.push({ date: _shortDate(t.date), description: p.payment_type || 'Miscellaneous', amount: Math.round(amt * 100) / 100, account: acctLabel });
            }
          });
          return;
        }
        if (STANDARD_INCOME_TYPES.indexOf(t.payment_type) > -1) {
          incomeGroups[t.payment_type] = (incomeGroups[t.payment_type] || 0) + credit;
        } else {
          var label = (t.payment_type === 'Bank Interest') ? 'Bank Interest' : (t.note && t.note.trim() ? t.note.trim() : 'Miscellaneous');
          miscRows.push({ date: _shortDate(t.date), description: label, amount: Math.round(credit * 100) / 100, account: acctLabel });
        }
      });
    }
    collectIncome(txns, '');
    collectIncome(txns2, 'IOB LPG');

    var incomeRows = [];
    var incomeSeq = 1;
    STANDARD_INCOME_TYPES.forEach(function(type) {
      var total = incomeGroups[type] || 0;
      if (total > 0) {
        incomeRows.push({ seq: incomeSeq++, date: _monthEndDisplay(year, month), description: INCOME_LABELS[type], amount: Math.round(total * 100) / 100, account: '' });
      }
    });
    miscRows.forEach(function(row) { row.seq = incomeSeq++; incomeRows.push(row); });
    var incomeTotal = incomeRows.reduce(function(s, r) { return s + r.amount; }, 0);

    // Expenses — grouped into ONE line per description (Expense Category,
    // or the Note/narration when there's no category) rather than one
    // row per individual transaction, so a recurring charge like "IOB
    // Bank Charges" or "LPG Inventory" that happens more than once in a
    // month shows as a single combined total instead of repeating.
    // Description prefers the Note (a human-written explanation) over
    // the raw bank narration, matching the same preference used
    // elsewhere in this ledger for readability. Merged across both
    // accounts the same way as Income above.
    var expenseGroups = {}; // description -> { amount, firstDate, accounts: {label: true} }
    var expenseOrder = [];
    function collectExpense(list, acctLabel) {
      list.forEach(function(t) {
        var debit = Number(t.debit || 0);
        if (!(debit > 0)) return;
        var desc = (t.expense_category && t.expense_category.trim()) ? t.expense_category.trim()
                  : (t.note && t.note.trim()) ? t.note.trim()
                  : t.narration;
        if (!expenseGroups[desc]) { expenseGroups[desc] = { amount: 0, firstDate: t.date, accounts: {} }; expenseOrder.push(desc); }
        expenseGroups[desc].amount += debit;
        if (acctLabel) expenseGroups[desc].accounts[acctLabel] = true;
      });
    }
    collectExpense(txns, '');
    collectExpense(txns2, 'IOB LPG');

    var expenseRows = [];
    var expSeq = 1;
    expenseOrder.forEach(function(desc) {
      var g = expenseGroups[desc];
      // Only tag a row when EVERY contributing transaction came from the
      // second account — a description mixing both accounts is genuinely
      // combined and shouldn't wear a single-account label.
      var acctKeys = Object.keys(g.accounts);
      var tag = (acctKeys.length === 1 && acctKeys[0] === 'IOB LPG') ? 'IOB LPG' : '';
      expenseRows.push({ seq: expSeq++, date: _shortDate(g.firstDate), description: desc, amount: Math.round(g.amount * 100) / 100, account: tag });
    });
    var expenseTotal = expenseRows.reduce(function(s, r) { return s + r.amount; }, 0);

    // Diagnostic — every credit transaction this month with its raw
    // Type, so an income discrepancy can be checked directly against
    // real data instead of guessed at. Both accounts included, tagged.
    var creditDiagnostic = [];
    function collectDiagnostic(list, acctLabel) {
      list.forEach(function(t) {
        var credit = Number(t.credit || 0);
        if (credit > 0) creditDiagnostic.push({ date: t.date, amount: credit, paymentType: t.payment_type || '(blank)', matchUnit: t.match_unit || '(unmatched)', narration: t.narration, account: acctLabel });
      });
    }
    collectDiagnostic(txns, '');
    collectDiagnostic(txns2, 'IOB LPG');

    // Reconciliation check for the warning banner — both accounts must
    // reconcile for this month (a month the second account has no data
    // for at all simply isn't checked against it).
    var reconciled = true;
    try {
      var recon = BankService.reconcile();
      var thisMonthRecon = null;
      recon.forEach(function(r) { if (r.month === key) thisMonthRecon = r; });
      reconciled = !thisMonthRecon || thisMonthRecon.ok;
    } catch (e) {}
    try {
      var recon2 = Bank2Service.reconcile();
      var thisMonthRecon2 = null;
      recon2.forEach(function(r) { if (r.month === key) thisMonthRecon2 = r; });
      if (thisMonthRecon2 && !thisMonthRecon2.ok) reconciled = false;
    } catch (e3) {}

    var mnFull = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
    var mn3 = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    return {
      year: year, month: month, monthName: mnFull[month],
      reportDate: _lastDayDisplay(year, month),
      prevMonthLabel: mn3[prevM] + (prevM === 12 ? ' ' + prevY : ''),
      openingBalance: Math.round(opening * 100) / 100,
      openingBreakdown: { main: Math.round(openingMain * 100) / 100, iobLpg: Math.round(openingAcct2 * 100) / 100 },
      openingSource: firstBalTxn ? { date: firstBalTxn.date, statedBalance: Number(String(firstBalTxn.balance).replace(/[₹,\s]/g,'')), credit: Number(firstBalTxn.credit||0), debit: Number(firstBalTxn.debit||0) } : null,
      incomeRows: incomeRows, incomeTotal: Math.round(incomeTotal * 100) / 100,
      totalWithIncome: Math.round((opening + incomeTotal) * 100) / 100,
      expenseRows: expenseRows, expenseTotal: Math.round(expenseTotal * 100) / 100,
      closingBalance: Math.round(closing * 100) / 100,
      closingBreakdown: { main: Math.round(closingMain * 100) / 100, iobLpg: Math.round(closingAcct2 * 100) / 100 },
      closingSource: lastBalTxn ? { date: lastBalTxn.date, statedBalance: Number(String(lastBalTxn.balance).replace(/[₹,\s]/g,'')) } : null,
      creditDiagnostic: creditDiagnostic,
      reconciled: reconciled
    };
  }

  // Monthly position of the second (IOB LPG) account. Dormant-safe: a
  // month can have zero transactions, in which case the balance simply
  // carries forward from the last transaction BEFORE the month — so
  // opening/closing are derived from the latest stated balance at each
  // boundary, not from assuming in-month activity exists.
  function _account2Summary(year, month) {
    var all = Bank2Service.getAllTransactions(null); // chronological
    if (!all.length) return null;
    var startKey = year * 100 + month;               // yyyymm of this month
    var credits = 0, debits = 0, interest = 0, txnCount = 0;
    var lastBalBefore = null, lastBalThrough = null;

    all.forEach(function(t) {
      // parseFlexibleDate returns a "yyyy-MM-dd" STRING (or '' when
      // unparseable), NOT a Date object — the year/month key is sliced
      // from the string directly.
      var iso = parseFlexibleDate(t.date);
      if (!iso) return;
      var k = Number(iso.slice(0, 4)) * 100 + Number(iso.slice(5, 7));
      var bal = (t.balance === '' || t.balance === null || t.balance === undefined) ? NaN : Number(String(t.balance).replace(/[₹,\s]/g, ''));
      if (k < startKey && !isNaN(bal)) lastBalBefore = bal;
      if (k <= startKey && !isNaN(bal)) lastBalThrough = bal;
      if (k === startKey) {
        txnCount++;
        var cr = Number(t.credit || 0), db = Number(t.debit || 0);
        credits += cr; debits += db;
        var isInterest = t.payment_type === 'Bank Interest' || /\bint(?:erest)?\b/i.test(String(t.narration || ''));
        if (isInterest && cr > 0) interest += cr;
      }
    });

    var opening = lastBalBefore;
    var closing = lastBalThrough;
    // If the month itself had activity but the running figures above
    // couldn't establish one of the boundaries (e.g. first-ever month),
    // derive what's derivable rather than reporting nothing.
    if (opening === null && closing !== null) opening = closing - credits + debits;
    if (opening === null && closing === null) return { available: false };

    return {
      available: true,
      label: 'IOB LPG Account',
      openingBalance: Math.round((opening === null ? 0 : opening) * 100) / 100,
      closingBalance: Math.round((closing === null ? opening : closing) * 100) / 100,
      credits: Math.round(credits * 100) / 100,
      debits: Math.round(debits * 100) / 100,
      interest: Math.round(interest * 100) / 100,
      txnCount: txnCount,
      dormant: txnCount === 0
    };
  }


  // Shared: map of txn_id -> its portions, from the BankPortions table.
  // A portioned transaction's income must be attributed by its portions'
  // types and amounts, never the bank row's single Type. (Replaces the
  // old split-notes convention, removed with the split feature.)
  function _splitMap() {
    var map = {};
    try {
      var all = PortionsService.getPortionsMap(null); // both accounts
      for (var tid in all) {
        map[tid] = all[tid].map(function(p) {
          return { unit_id: p.unit_id, payment_type: p.payment_type, amount: p.amount };
        });
      }
    } catch (e) {}
    return map;
  }

  // Per-transaction income detail for one month — every credit, with the
  // payer name parsed from the narration, the matched unit, and the fee
  // type (split transactions contribute one row per split portion, since
  // one credit can genuinely carry several types/units). Cash basis: a
  // credit belongs to the month it landed in the bank.
  function getMonthlyIncomeData(year, month) {
    year = Number(year); month = Number(month);
    var key = _monthKey(year, month);
    var splitMap = _splitMap();

    function rowsFor(svc, accountLabel) {
      var out = [];
      svc.getAllTransactions({ month: key }).forEach(function(t) {
        var credit = Number(t.credit || 0);
        if (!(credit > 0)) return;
        var sortKey = parseFlexibleDate(t.date) || ''; // "yyyy-MM-dd" — for chronological merge only
        var splits = splitMap[t.txn_id];
        if (splits && splits.length) {
          splits.forEach(function(p) {
            out.push({ date: _shortDate(t.date), sortKey: sortKey, payer: t.parsed_name || '', unit: p.unit_id || t.match_unit || '', type: p.payment_type || 'Unknown', amount: Math.round(Number(p.amount || 0) * 100) / 100, account: accountLabel, split: true });
          });
        } else {
          out.push({ date: _shortDate(t.date), sortKey: sortKey, payer: t.parsed_name || '', unit: t.match_unit || '', type: t.payment_type || 'Unknown', amount: Math.round(credit * 100) / 100, account: accountLabel, split: false });
        }
      });
      return out;
    }

    var rows = rowsFor(BankService, '');
    try { rows = rows.concat(rowsFor(Bank2Service, 'IOB LPG')); } catch (e) {}
    rows.sort(function(a, b) { return a.sortKey < b.sortKey ? -1 : a.sortKey > b.sortKey ? 1 : 0; });
    rows.forEach(function(r) { delete r.sortKey; });

    var typeTotals = {};
    var grand = 0;
    rows.forEach(function(r) {
      typeTotals[r.type] = (typeTotals[r.type] || 0) + r.amount;
      grand += r.amount;
    });
    var totalsList = [];
    for (var ty in typeTotals) totalsList.push({ type: ty, label: INCOME_LABELS[ty] || ty, total: Math.round(typeTotals[ty] * 100) / 100 });
    totalsList.sort(function(a, b) { return b.total - a.total; });

    var mnFull = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
    return { year: year, month: month, monthName: mnFull[month], rows: rows,
             typeTotals: totalsList, grandTotal: Math.round(grand * 100) / 100 };
  }

  // Per-transaction expense detail for one month — every debit with its
  // expense category and description. Cash basis by bank date.
  function getMonthlyExpenseData(year, month) {
    year = Number(year); month = Number(month);
    var key = _monthKey(year, month);

    function rowsFor(svc, accountLabel) {
      var out = [];
      svc.getAllTransactions({ month: key }).forEach(function(t) {
        var debit = Number(t.debit || 0);
        if (!(debit > 0)) return;
        var desc = (t.note && String(t.note).trim()) ? String(t.note).trim() : String(t.narration || '');
        var sortKey = parseFlexibleDate(t.date) || '';
        out.push({ date: _shortDate(t.date), sortKey: sortKey, category: t.expense_category || 'Uncategorised', description: desc, amount: Math.round(debit * 100) / 100, account: accountLabel });
      });
      return out;
    }

    var rows = rowsFor(BankService, '');
    try { rows = rows.concat(rowsFor(Bank2Service, 'IOB LPG')); } catch (e) {}
    rows.sort(function(a, b) { return a.sortKey < b.sortKey ? -1 : a.sortKey > b.sortKey ? 1 : 0; });
    rows.forEach(function(r) { delete r.sortKey; });

    var catTotals = {};
    var grand = 0;
    rows.forEach(function(r) {
      catTotals[r.category] = (catTotals[r.category] || 0) + r.amount;
      grand += r.amount;
    });
    var totalsList = [];
    for (var c in catTotals) totalsList.push({ category: c, total: Math.round(catTotals[c] * 100) / 100 });
    totalsList.sort(function(a, b) { return b.total - a.total; });

    var mnFull = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
    return { year: year, month: month, monthName: mnFull[month], rows: rows,
             categoryTotals: totalsList, grandTotal: Math.round(grand * 100) / 100 };
  }

  // Annual Income matrix — one row per income type, one column per month
  // (Jan..Dec), each cell the type's total credits that month. Main
  // account, cash basis, split-aware. Bank Interest and anything outside
  // the five standard fee types roll up into their own rows.
  function getAnnualIncomeData(year) {
    year = Number(year);
    var splitMap = _splitMap();
    var txns = BankService.getAllTransactions(null);
    var txns2 = [];
    try { txns2 = Bank2Service.getAllTransactions(null); } catch (e) {}
    var matrix = {}; // type -> [12 month totals]

    function add(type, mIdx, amt) {
      if (!matrix[type]) { matrix[type] = [0,0,0,0,0,0,0,0,0,0,0,0]; }
      matrix[type][mIdx] += amt;
    }

    function collect(list) {
      list.forEach(function(t) {
        var credit = Number(t.credit || 0);
        if (!(credit > 0)) return;
        var iso = parseFlexibleDate(t.date); // "yyyy-MM-dd" string, or '' if unparseable
        if (!iso || Number(iso.slice(0, 4)) !== year) return;
        var mIdx = Number(iso.slice(5, 7)) - 1;
        var splits = splitMap[t.txn_id];
        if (splits && splits.length) {
          splits.forEach(function(p) { add(p.payment_type || 'Unknown', mIdx, Number(p.amount || 0)); });
        } else {
          add(t.payment_type || 'Unknown', mIdx, credit);
        }
      });
    }
    collect(txns);
    collect(txns2);

    // Fixed, meaningful row order: the five standard types first, then
    // whatever else showed up (Bank Interest, Miscellaneous, Unknown...)
    var order = STANDARD_INCOME_TYPES.slice();
    for (var ty in matrix) if (order.indexOf(ty) === -1) order.push(ty);

    var rows = [];
    var colTotals = [0,0,0,0,0,0,0,0,0,0,0,0];
    var grand = 0;
    order.forEach(function(ty) {
      if (!matrix[ty]) return;
      var months = matrix[ty].map(function(v) { return Math.round(v * 100) / 100; });
      var rowTotal = 0;
      months.forEach(function(v, i) { rowTotal += v; colTotals[i] += v; });
      grand += rowTotal;
      rows.push({ type: ty, label: INCOME_LABELS[ty] || ty, months: months, total: Math.round(rowTotal * 100) / 100 });
    });

    return { year: year, rows: rows,
             colTotals: colTotals.map(function(v) { return Math.round(v * 100) / 100; }),
             grandTotal: Math.round(grand * 100) / 100 };
  }

  // Annual Expense matrix — one row per expense category, one column per
  // month. Main account debits, cash basis: an expense belongs to the
  // month the money LEFT the bank, even if it pays for a previous
  // month's bill (e.g. December's electricity paid in January counts in
  // January — the description carries the arrears context, the money
  // movement defines the column).
  function getAnnualExpenseData(year) {
    year = Number(year);
    var txns = BankService.getAllTransactions(null);
    var txns2 = [];
    try { txns2 = Bank2Service.getAllTransactions(null); } catch (e) {}
    var matrix = {}; // category -> [12 totals]

    function collect(list) {
      list.forEach(function(t) {
        var debit = Number(t.debit || 0);
        if (!(debit > 0)) return;
        var iso = parseFlexibleDate(t.date); // "yyyy-MM-dd" string, or '' if unparseable
        if (!iso || Number(iso.slice(0, 4)) !== year) return;
        var cat = t.expense_category || 'Uncategorised';
        if (!matrix[cat]) matrix[cat] = [0,0,0,0,0,0,0,0,0,0,0,0];
        matrix[cat][Number(iso.slice(5, 7)) - 1] += debit;
      });
    }
    collect(txns);
    collect(txns2);

    var cats = [];
    for (var c in matrix) cats.push(c);
    cats.sort();

    var rows = [];
    var colTotals = [0,0,0,0,0,0,0,0,0,0,0,0];
    var grand = 0;
    cats.forEach(function(cat) {
      var months = matrix[cat].map(function(v) { return Math.round(v * 100) / 100; });
      var rowTotal = 0;
      months.forEach(function(v, i) { rowTotal += v; colTotals[i] += v; });
      grand += rowTotal;
      rows.push({ category: cat, months: months, total: Math.round(rowTotal * 100) / 100 });
    });

    return { year: year, rows: rows,
             colTotals: colTotals.map(function(v) { return Math.round(v * 100) / 100; }),
             grandTotal: Math.round(grand * 100) / 100 };
  }

  function _shortDate(dateStr) {
    // BankService dates are "DD-Mon-YYYY" text. Standardized to the
    // app-wide "01 Jan 2024" format: zero-padded day, full year.
    var m = String(dateStr || '').match(/^(\d{1,2})-([A-Za-z]{3})[A-Za-z]*-(\d{4})/);
    if (!m) return String(dateStr || '');
    var dd = Number(m[1]);
    return (dd < 10 ? '0' + dd : String(dd)) + ' ' + m[2] + ' ' + m[3];
  }

  function _lastDayDisplay(year, month) {
    var lastDay = new Date(year, month, 0);
    var dd = _pad2(lastDay.getDate());
    var mn = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return dd + ' ' + mn[month] + ' ' + year;
  }

  return {
    getTreasurerReportData: getTreasurerReportData,
    getMonthlyIncomeData:   getMonthlyIncomeData,
    getMonthlyExpenseData:  getMonthlyExpenseData,
    getAnnualIncomeData:    getAnnualIncomeData,
    getAnnualExpenseData:   getAnnualExpenseData
  };
})();
