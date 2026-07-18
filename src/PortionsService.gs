// ═══════════════════════════════════════════════════════════════
// PortionsService.gs — the PARKED FUNDS model for edge-case payments.
//
// PRINCIPLES (agreed after the split/redistribution failure):
//   1. The bank row is the TRUE SOURCE — never modified by this module.
//   2. Simple payments (one unit, one type) keep the proven automatic
//      posting pipeline. This module handles ONLY the edge cases:
//      multi-unit payments, combined MF+WMF+LPG payments, and bulk
//      amounts to be spread across months.
//   3. "Assign Portions" on a bank transaction records WHO paid WHAT
//      (unit + fee type + amount, summing to the transaction total).
//      Portions create PARKED MONEY — never Fees Received records.
//   4. Month allocation is 100% MANUAL. The system never guesses which
//      months parked money covers; only the user knows. Nothing is
//      written until the user presses Save.
//   5. NOTHING IS EVER DELETED OR MODIFIED SILENTLY. Every destructive
//      action lists exactly what will be removed and requires explicit
//      confirmation. Once the user has approved data, no tool touches
//      it without the user's consent.
//
// Data: sheet 'BankPortions' (one row per portion), and manual
// allocations in the normal Payments ledger tagged 'PARKED-ALLOC' in
// notes. parked(unit,type) = Σ portions − Σ PARKED-ALLOC payments.
// ═══════════════════════════════════════════════════════════════
var PortionsService = (function() {
  var SHEET = 'BankPortions';
  var HEADERS = ['portion_id','txn_id','account','unit_id','payment_type','amount','created_at','updated_at'];
  var C = { ID:0, TXN:1, ACCT:2, UNIT:3, TYPE:4, AMOUNT:5, CREATED:6, UPDATED:7 };
  var ALLOC_TAG = 'PARKED-ALLOC';

  function ensureSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET);
    if (!sh) {
      sh = ss.insertSheet(SHEET);
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  }

  function _toObj(r) {
    return { portion_id: String(r[C.ID] || ''), txn_id: String(r[C.TXN] || ''),
             account: String(r[C.ACCT] || '1'), unit_id: String(r[C.UNIT] || '').toUpperCase(),
             payment_type: String(r[C.TYPE] || ''), amount: Number(r[C.AMOUNT]) || 0 };
  }

  function _allPortions() {
    ensureSheet();
    return Database.getAll(SHEET).map(_toObj).filter(function(p) { return p.portion_id; });
  }

  // ── Assign / replace the portions of one bank transaction ──
  // Refuses (with an explicit list) when auto-posted Fees records exist
  // for this transaction, unless the user has confirmed their removal.
  function assignPortions(d) {
    if (!d || !d.txn_id) throw new Error('txn_id is required.');
    var account = String(d.account) === '2' ? '2' : '1';
    var svc = account === '2' ? Bank2Service : BankService;
    var portions = d.portions || [];
    if (!portions.length) throw new Error('At least one portion is required.');

    // The bank row is the reference: portions must sum to its credit.
    var txn = null;
    svc.getAllTransactions(null).forEach(function(t) { if (t.txn_id === d.txn_id) txn = t; });
    if (!txn) throw new Error('Transaction not found: ' + d.txn_id);
    var credit = Number(txn.credit || 0);
    if (!(credit > 0)) throw new Error('Portions apply to credit (incoming) transactions only.');

    var sum = 0;
    portions.forEach(function(p) {
      if (!p.unit_id || !String(p.unit_id).trim()) throw new Error('Every portion needs a unit.');
      if (!p.payment_type) throw new Error('Every portion needs a payment type.');
      var a = Number(p.amount);
      if (!(a > 0)) throw new Error('Every portion needs an amount greater than zero.');
      sum += a;
    });
    if (Math.abs(sum - credit) > 0.01) {
      throw new Error('Portions must add up to the transaction total: entered ' + sum + ', transaction is ' + credit + '.');
    }

    // Guard 1 — auto-posted records from the simple pipeline: never
    // removed silently. First call returns the list; the client shows
    // it and re-sends with confirm_remove_autoposted=true.
    var autoPosted = _paymentsWithBankRef(d.txn_id);
    if (autoPosted.length && d.confirm_remove_autoposted !== true) {
      return { needs_confirm: true,
               reason: 'auto_posted',
               records: autoPosted.map(function(p) {
                 return { payment_id: p.payment_id, unit_id: p.unit_id, payment_type: p.payment_type, month: p.month, amount: p.amount };
               }) };
    }

    // Guard 2 — existing manual allocations that drew on this txn's
    // OLD portions: replacing portions could strand them. Block until
    // the user unallocates first (explicit, reviewable, no surprises).
    var existing = getPortionsForTxn(d.txn_id);
    if (existing.length) {
      var overAlloc = [];
      var seenPairs = {};
      existing.forEach(function(p) {
        var k = p.unit_id + '|' + p.payment_type;
        if (seenPairs[k]) return;
        seenPairs[k] = true;
        var summ = _pairSummary(p.unit_id, p.payment_type);
        // If removing this txn's parked money would leave the pair
        // over-allocated, the user must unallocate first.
        var newPortionHere = 0;
        portions.forEach(function(np) { if (String(np.unit_id).toUpperCase() === p.unit_id && np.payment_type === p.payment_type) newPortionHere += Number(np.amount); });
        var oldPortionHere = 0;
        existing.forEach(function(op) { if (op.unit_id === p.unit_id && op.payment_type === p.payment_type) oldPortionHere += op.amount; });
        var futureParked = summ.parked - oldPortionHere + newPortionHere;
        if (summ.allocated - futureParked > 0.01) {
          overAlloc.push(p.unit_id + ' ' + p.payment_type + ': allocated ' + summ.allocated + ' would exceed parked ' + futureParked);
        }
      });
      if (overAlloc.length) {
        throw new Error('These units already have manual month allocations drawing on this transaction\'s parked money — remove those allocations first (Unit Profile → Parked Funds → Remove allocations): ' + overAlloc.join(' · '));
      }
    }

    // All guards passed. Apply, in order: remove confirmed auto-posted
    // records (listed to and approved by the user), replace portions.
    if (autoPosted.length) {
      var pSheet = Database.getSheet('Payments');
      // Delete bottom-up by row index.
      autoPosted.sort(function(a, b) { return b.rowIndex - a.rowIndex; });
      autoPosted.forEach(function(p) { pSheet.deleteRow(p.rowIndex); });
    }

    var sh = ensureSheet();
    var rows = Database.getAll(SHEET);
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][C.TXN]) === d.txn_id) sh.deleteRow(i + 2);
    }
    var now = new Date().toISOString();
    var out = portions.map(function(p) {
      return [Database.generateId('POR'), d.txn_id, account, String(p.unit_id).toUpperCase(),
              String(p.payment_type), Number(p.amount), now, now];
    });
    sh.getRange(sh.getLastRow() + 1, 1, out.length, HEADERS.length).setValues(out);

    return { success: true, portions: out.length, autoPostedRemoved: autoPosted.length };
  }

  function _paymentsWithBankRef(txnId) {
    var out = [];
    var all = Database.getAll('Payments');
    var re = new RegExp('^BANK:' + txnId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b');
    for (var i = 0; i < all.length; i++) {
      if (re.test(String(all[i][8] || ''))) { // 8 = NOTES
        out.push({ rowIndex: i + 2, payment_id: String(all[i][0]), unit_id: String(all[i][1]),
                   payment_type: String(all[i][3]), amount: Number(all[i][4]) || 0,
                   month: String(all[i][5] || '').replace(/^'/, '') });
      }
    }
    return out;
  }

  function getPortionsForTxn(txnId) {
    return _allPortions().filter(function(p) { return p.txn_id === txnId; });
  }

  // txn_id -> portion count, for badges/addenda; account-filterable.
  function getPortionTxnIds(account) {
    var map = {};
    _allPortions().forEach(function(p) {
      if (account && p.account !== String(account)) return;
      map[p.txn_id] = (map[p.txn_id] || 0) + 1;
    });
    return map;
  }

  // txn_id -> [portions], one call for the whole Bank Statement page:
  // powers both the 🧩 badges and the addendum sub-rows.
  function getPortionsMap(account) {
    var map = {};
    _allPortions().forEach(function(p) {
      if (account !== null && account !== undefined && account !== '' && p.account !== String(account)) return;
      if (!map[p.txn_id]) map[p.txn_id] = [];
      map[p.txn_id].push({ unit_id: p.unit_id, payment_type: p.payment_type, amount: p.amount });
    });
    return map;
  }

  function getPortionTxnIdSet() {
    var map = {};
    _allPortions().forEach(function(p) { map[p.txn_id] = true; });
    return map;
  }

  // All manual allocations for unit+type (the PARKED-ALLOC records).
  function _allocationsFor(unitId, type) {
    var unit = String(unitId).toUpperCase();
    var out = [];
    var all = Database.getAll('Payments');
    for (var i = 0; i < all.length; i++) {
      if (String(all[i][1]).toUpperCase() !== unit) continue;
      if (String(all[i][3]) !== type) continue;
      if (String(all[i][8] || '').indexOf(ALLOC_TAG) !== 0) continue;
      if (String(all[i][7]) === 'Rejected') continue;
      out.push({ rowIndex: i + 2, payment_id: String(all[i][0]),
                 month: String(all[i][5] || '').replace(/^'/, ''), amount: Number(all[i][4]) || 0 });
    }
    return out;
  }

  function _pairSummary(unitId, type) {
    var unit = String(unitId).toUpperCase();
    var parked = 0;
    var portionList = [];
    _allPortions().forEach(function(p) {
      if (p.unit_id === unit && p.payment_type === type) { parked += p.amount; portionList.push(p); }
    });
    var allocated = 0;
    var allocs = _allocationsFor(unit, type);
    allocs.forEach(function(a) { allocated += a.amount; });
    return { parked: Math.round(parked * 100) / 100, allocated: Math.round(allocated * 100) / 100,
             remaining: Math.round((parked - allocated) * 100) / 100,
             portions: portionList, allocations: allocs };
  }

  // Everything the Unit Profile's Parked Funds card needs.
  function getParkedSummary(unitId) {
    var unit = String(unitId).toUpperCase();
    var types = {};
    _allPortions().forEach(function(p) { if (p.unit_id === unit) types[p.payment_type] = true; });
    // Also surface pairs that have allocations but lost their portions
    // (e.g. the bank month was wiped) — shown as over-allocated, never hidden.
    var all = Database.getAll('Payments');
    for (var i = 0; i < all.length; i++) {
      if (String(all[i][1]).toUpperCase() !== unit) continue;
      if (String(all[i][8] || '').indexOf(ALLOC_TAG) !== 0) continue;
      types[String(all[i][3])] = true;
    }
    var out = [];
    for (var ty in types) {
      var s = _pairSummary(unit, ty);
      out.push({ payment_type: ty, parked: s.parked, allocated: s.allocated, remaining: s.remaining,
                 portions: s.portions.map(function(p) { return { txn_id: p.txn_id, account: p.account, amount: p.amount }; }),
                 allocations: s.allocations.map(function(a) { return { payment_id: a.payment_id, month: a.month, amount: a.amount }; }) });
    }
    out.sort(function(a, b) { return a.payment_type.localeCompare(b.payment_type); });
    return { unit_id: unit, pairs: out };
  }

  // ── 100% manual month allocation. Validates against the parked
  //    balance, writes only what the user typed, one Verified payment
  //    record per month, tagged PARKED-ALLOC. ──
  function allocate(d) {
    if (!d || !d.unit_id || !d.payment_type) throw new Error('unit_id and payment_type are required.');
    var unit = String(d.unit_id).toUpperCase();
    var months = d.months || {}; // { '2024-01': 1500, ... }
    var total = 0;
    var keys = [];
    for (var k in months) {
      var a = Number(months[k]);
      if (!(a > 0)) continue;
      if (!/^\d{4}-\d{2}$/.test(k)) throw new Error('Bad month key: ' + k);
      total += a;
      keys.push(k);
    }
    if (!keys.length) throw new Error('Enter at least one month amount.');

    var s = _pairSummary(unit, d.payment_type);
    if (total - s.remaining > 0.01) {
      throw new Error('You entered ' + total + ' but only ' + s.remaining + ' is parked and unallocated for ' + unit + ' ' + d.payment_type + '. Nothing was saved.');
    }

    // Historical-year rule applies exactly as elsewhere.
    keys.sort();
    var created = [];
    keys.forEach(function(mk) {
      var res = PaymentsService.submitPayment({
        unit_id: unit, payment_type: d.payment_type, amount: Number(months[mk]),
        month: mk, status: 'Verified', reviewed_by: 'Manual allocation',
        notes: ALLOC_TAG + ' (manual month allocation from parked funds)',
        confirmed_historical: d.confirmed_historical === true
      });
      created.push({ month: mk, amount: Number(months[mk]), payment_id: res.payment_id });
    });
    var after = _pairSummary(unit, d.payment_type);
    return { success: true, created: created, remaining: after.remaining };
  }

  // ── Two-phase removal of manual allocations: phase 1 lists exactly
  //    what would be deleted; phase 2 (confirm=true) deletes precisely
  //    that list. Only PARKED-ALLOC records are ever touched. ──
  function removeAllocations(d) {
    if (!d || !d.unit_id || !d.payment_type) throw new Error('unit_id and payment_type are required.');
    var unit = String(d.unit_id).toUpperCase();
    var allocs = _allocationsFor(unit, d.payment_type);
    if (d.year) allocs = allocs.filter(function(a) { return String(a.month).slice(0, 4) === String(d.year); });
    if (!allocs.length) return { success: true, deleted: 0, records: [] };

    if (d.confirm !== true) {
      return { needs_confirm: true,
               records: allocs.map(function(a) { return { payment_id: a.payment_id, month: a.month, amount: a.amount }; }) };
    }

    // Year-edit rule: deleting a past year's records needs the same
    // confirmation as any other historical edit.
    var years = {};
    allocs.forEach(function(a) { years[String(a.month).slice(0, 4)] = true; });
    for (var y in years) checkYearEditable(Number(y), d.confirmed_historical === true);

    var sheet = Database.getSheet('Payments');
    allocs.sort(function(a, b) { return b.rowIndex - a.rowIndex; });
    allocs.forEach(function(a) { sheet.deleteRow(a.rowIndex); });
    return { success: true, deleted: allocs.length };
  }

  // Bank-month deletion cascade: when a bank month is wiped for
  // re-import, its transactions' portions go with it (they describe
  // rows that no longer exist). Manual allocations are NOT touched —
  // they are user-approved data; if their parked backing disappears,
  // the Parked Funds card shows the pair as over-allocated instead.
  function deleteForTxns(txnIds) {
    if (!txnIds || !txnIds.length) return { deleted: 0 };
    var set = {};
    txnIds.forEach(function(t) { set[String(t)] = true; });
    var sh = ensureSheet();
    var rows = Database.getAll(SHEET);
    var deleted = 0;
    for (var i = rows.length - 1; i >= 0; i--) {
      if (set[String(rows[i][C.TXN])]) { sh.deleteRow(i + 2); deleted++; }
    }
    return { deleted: deleted };
  }

  return {
    ensureSheet:       ensureSheet,
    assignPortions:    assignPortions,
    getPortionsForTxn: getPortionsForTxn,
    getPortionTxnIds:  getPortionTxnIds,
    getPortionsMap:    getPortionsMap,
    getPortionTxnIdSet: getPortionTxnIdSet,
    getParkedSummary:  getParkedSummary,
    allocate:          allocate,
    removeAllocations: removeAllocations,
    deleteForTxns:     deleteForTxns
  };
})();
