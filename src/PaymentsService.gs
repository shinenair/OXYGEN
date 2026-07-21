// ============================================================
// PaymentsService.gs — Payment Management
// ============================================================

var PaymentsService = (function() {
  var SHEET = 'Payments';

  var C = {
    PAYMENT_ID:    0,
    UNIT_ID:       1,
    TENANT_ID:     2,
    PAYMENT_TYPE:  3,
    AMOUNT:        4,
    MONTH:         5,
    SCREENSHOT_URL:6,
    STATUS:        7,
    NOTES:         8,
    SUBMITTED_AT:  9,
    REVIEWED_AT:   10,
    REVIEWED_BY:   11
  };

  function submitPayment(data) {
    _validate(data);
    var now       = new Date().toISOString();
    var paymentId = Database.generateId('PAY');
    var month     = _normMonth(data.month || _currentMonth());
    var recordYear = Number(String(month).split('-')[0]);
    checkYearEditable(recordYear, data.confirmed_historical);
    var status    = data.status || 'Pending';

    Database.insert(SHEET, [
      paymentId,
      data.unit_id,
      data.tenant_id || '',
      data.payment_type,
      Number(data.amount),
      "'" + month,
      data.screenshot_url || '',
      status,
      data.notes          || '',
      now,
      status === 'Verified' ? now : '',
      status === 'Verified' ? (data.reviewed_by || 'Manager') : ''
    ]);

    return { success: true, payment_id: paymentId };
  }

  function verifyPayment(paymentId, reviewedBy, notes) {
    return _updateStatus(paymentId, 'Verified', reviewedBy, notes);
  }

  function rejectPayment(paymentId, reviewedBy, notes) {
    return _updateStatus(paymentId, 'Rejected', reviewedBy, notes);
  }

  function _updateStatus(paymentId, status, reviewedBy, notes) {
    var result = Database.findByColumn(SHEET, C.PAYMENT_ID, paymentId);
    if (!result) throw new Error('Payment not found: ' + paymentId);

    var now = new Date().toISOString();
    var row = result.data;
    row[C.STATUS]      = status;
    row[C.REVIEWED_AT] = now;
    row[C.REVIEWED_BY] = reviewedBy || 'Manager';
    if (notes !== undefined) row[C.NOTES] = notes;

    Database.updateRow(SHEET, result.rowIndex, row);
    return { success: true };
  }

  function getPaymentById(paymentId) {
    var result = Database.findByColumn(SHEET, C.PAYMENT_ID, paymentId);
    return result ? _toObj(result.data) : null;
  }

  function getPaymentsByUnit(unitId, limit) {
    var results = Database.findAllByColumn(SHEET, C.UNIT_ID, unitId);
    var payments = results.map(function(r) { return _toObj(r.data); })
      .sort(function(a, b) { return b.submitted_at.localeCompare(a.submitted_at); });
    return limit ? payments.slice(0, limit) : payments;
  }

  function getPaymentsByTenantAndMonth(tenantId, month) {
    var rows = Database.getAll(SHEET);
    return rows.map(function(r) { return _toObj(r); })
      .filter(function(p) {
        return String(p.tenant_id) === String(tenantId) &&
               String(p.month)     === String(month);
      });
  }

  function getPendingPayments() {
    var rows = Database.getAll(SHEET);
    return rows.map(function(r) { return _toObj(r); })
      .filter(function(p) { return p.status === 'Pending'; })
      .sort(function(a, b) { return a.submitted_at.localeCompare(b.submitted_at); });
  }

  // Every distinct month present in the ledger, with a row count — used
  // by the Danger Zone's "delete a specific month" table.
  function getMonthSummary() {
    var rows = Database.getAll(SHEET);
    var counts = {};
    for (var i = 0; i < rows.length; i++) {
      // Normalize the same way _toObj / the Fees Received page does: the
      // month is stored apostrophe-prefixed and Sheets often returns it
      // as a Date or unpadded string. Reading it raw produced keys like
      // "Sun Jan 01 2024…" that never matched the bank's YYYY-MM keys, so
      // payment counts vanished from the combined month table.
      var m = _normMonth(rows[i][C.MONTH]);
      if (!m) continue;
      counts[m] = (counts[m] || 0) + 1;
    }
    var out = [];
    for (var k in counts) out.push({ month: k, count: counts[k] });
    out.sort(function(a, b) { return a.month < b.month ? -1 : 1; });
    return out;
  }

  // Deletes every payment row for one month key (e.g. '2024-01'). Used for
  // a clean re-import of that month's data without touching any other month.
  // Delete ONE payment record by id — the surgical broom for residual
  // or wrong records, reachable from the fees cell inspector. Respects
  // the historical-year rule, and returns an echo of what was deleted.
  function deleteRecord(paymentId, confirmedHistorical) {
    var found = Database.findByColumn(SHEET, C.PAYMENT_ID, paymentId);
    if (!found) throw new Error('Payment record not found: ' + paymentId);
    var r = found.data;
    var month = String(r[C.MONTH] || '').replace(/^'/, '');
    var yr = Number(month.split('-')[0]);
    if (yr) checkYearEditable(yr, confirmedHistorical === true);
    var echo = { payment_id: String(r[C.PAYMENT_ID]), unit_id: String(r[C.UNIT_ID]),
                 payment_type: String(r[C.PAYMENT_TYPE]), month: month, amount: Number(r[C.AMOUNT]) || 0 };
    Database.getSheet(SHEET).deleteRow(found.rowIndex);
    return { success: true, deleted: echo };
  }

  function deleteByMonth(monthKey) {
    var sheet = Database.getSheet(SHEET);
    var rows = Database.getAll(SHEET);
    var deleted = 0;
    // Normalize the stored month (Date / unpadded / apostrophe-prefixed)
    // before comparing — a raw string compare missed rows Sheets returns
    // as Dates, so a "delete this month" reported success while leaving
    // those fee records behind (also affected the bank-delete cascade).
    for (var i = rows.length - 1; i >= 0; i--) {
      if (_normMonth(rows[i][C.MONTH]) === monthKey) { sheet.deleteRow(i + 2); deleted++; }
    }
    return { success: true, deleted: deleted };
  }

  // Shift every real payment (Verified/Pending) for one unit + type by a
  // whole number of months — the one-click fix for the LPG timing offset,
  // where a unit's payments were all recorded one month late. UO markers and
  // Rejected rows are left in place (they're tied to the actual month, not a
  // payment). Amounts, bank references and dates are untouched — only the
  // month a payment is recorded FOR moves. Returns a from→to echo per row.
  function shiftUnitMonths(unitId, paymentType, offset, year) {
    var unit = String(unitId || '').toUpperCase();
    var off  = Number(offset) || 0;
    var yr   = Number(year) || 0;   // 0 = no year filter (legacy callers)
    if (!unit || !paymentType) throw new Error('Unit and payment type are required.');
    if (!off) return { success: true, changed: 0, skippedBoundary: 0, details: [] };
    var sheet = Database.getSheet(SHEET);
    var rows  = Database.getAll(SHEET);
    var changed = 0, skippedBoundary = 0, details = [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.UNIT_ID]).toUpperCase() !== unit) continue;
      if (String(rows[i][C.PAYMENT_TYPE]) !== paymentType) continue;
      var st = String(rows[i][C.STATUS]);
      if (st !== 'Verified' && st !== 'Pending') continue; // leave UO / Rejected
      var mk = _normMonth(rows[i][C.MONTH]);
      if (!mk) continue;
      var parts = mk.split('-');
      var recY = Number(parts[0]), recM = Number(parts[1]);
      // Year scope: only touch records in the active/viewed year, so shifting
      // 2024 never affects 2023 or 2025.
      if (yr && recY !== yr) continue;
      var newM = recM + off;
      // Stay inside the SAME year — a shift that would cross the year boundary
      // (Jan → prev Dec, or Dec → next Jan) is intentionally left as-is so the
      // neighbouring year is untouched; those are fixed via the Edit window.
      if (newM < 1 || newM > 12) { skippedBoundary++; continue; }
      var nk = _mk(recY, newM);
      sheet.getRange(i + 2, C.MONTH + 1).setValue("'" + nk);
      changed++;
      details.push({ from: mk, to: nk, amount: Number(rows[i][C.AMOUNT]) || 0 });
    }
    if (changed) SpreadsheetApp.flush();
    return { success: true, changed: changed, skippedBoundary: skippedBoundary, details: details };
  }

  function getAllPayments(filters) {
    var rows = Database.getAll(SHEET);
    var payments = rows.map(function(r) { return _toObj(r); });

    if (filters) {
      if (filters.status)       payments = payments.filter(function(p) { return p.status       === filters.status; });
      if (filters.unit_id)      payments = payments.filter(function(p) { return p.unit_id      === filters.unit_id; });
      if (filters.tenant_id)    payments = payments.filter(function(p) { return p.tenant_id    === filters.tenant_id; });
      if (filters.payment_type) payments = payments.filter(function(p) { return p.payment_type === filters.payment_type; });
      if (filters.month)        payments = payments.filter(function(p) { return p.month        === filters.month; });
    }

    return payments.sort(function(a, b) { return b.submitted_at.localeCompare(a.submitted_at); });
  }

  function getMonthlyPaymentHistory(unitId) {
    var payments = getPaymentsByUnit(unitId);
    var grouped  = {};

    payments.forEach(function(p) {
      if (!grouped[p.month]) grouped[p.month] = { month: p.month, payments: [], totalVerified: 0, totalPending: 0 };
      grouped[p.month].payments.push(p);
      if (p.status === 'Verified') grouped[p.month].totalVerified += Number(p.amount);
      if (p.status === 'Pending')  grouped[p.month].totalPending  += Number(p.amount);
    });

    return Object.keys(grouped).sort().reverse().map(function(k) { return grouped[k]; });
  }

  function getPaymentStats() {
    var all        = Database.getAll(SHEET).map(function(r) { return _toObj(r); });
    var month      = _currentMonth();
    var thisMonth  = all.filter(function(p) { return p.month === month; });

    return {
      totalPending:          all.filter(function(p) { return p.status === 'Pending';  }).length,
      totalVerified:         all.filter(function(p) { return p.status === 'Verified'; }).length,
      totalRejected:         all.filter(function(p) { return p.status === 'Rejected'; }).length,
      currentMonthCollected: thisMonth.filter(function(p) { return p.status === 'Verified'; })
                               .reduce(function(s, p) { return s + Number(p.amount); }, 0),
      currentMonthPending:   thisMonth.filter(function(p) { return p.status === 'Pending'; })
                               .reduce(function(s, p) { return s + Number(p.amount); }, 0)
    };
  }

  /**
   * updatePayment — manual correction of any payment record.
   * data may contain: payment_type, month (YYYY-MM), amount, unit_id, status, notes
   */
  function updatePayment(paymentId, data) {
    var result = Database.findByColumn(SHEET, C.PAYMENT_ID, paymentId);
    if (!result) throw new Error('Payment not found: ' + paymentId);
    var row = result.data;
    if (data.payment_type !== undefined && data.payment_type !== '') row[C.PAYMENT_TYPE] = data.payment_type;
    if (data.month        !== undefined && data.month        !== '') row[C.MONTH]        = "'" + _normMonth(data.month);
    if (data.amount       !== undefined && data.amount       !== '') row[C.AMOUNT]       = Number(data.amount);
    if (data.unit_id      !== undefined && data.unit_id      !== '') row[C.UNIT_ID]      = String(data.unit_id).toUpperCase();
    if (data.status       !== undefined && data.status       !== '') row[C.STATUS]       = data.status;
    if (data.notes        !== undefined)                             row[C.NOTES]        = data.notes;
    row[C.REVIEWED_AT] = new Date().toISOString();
    if (data.reviewed_by) row[C.REVIEWED_BY] = data.reviewed_by;
    Database.updateRow(SHEET, result.rowIndex, row);
    return { success: true };
  }

  /**
   * allocateFromBankBatch — the automatic bank -> payments pipeline.
   *
   * txns: array of { txn_id, unit_id, tenant_id, payment_type, amount, dateStr }
   * fees: { 'Maintenance': 2000, 'Waste Management': 170 } from Settings
   *
   * Rules:
   *  - The txn date decides the STARTING month (e.g. 05-May-2025 -> 2025-05).
   *  - If that month already has a Verified payment of the same type for the
   *    unit, the allocation moves FORWARD to the next unpaid month.
   *  - Exact multiples of the fee cover several months (₹4000 -> 2 months).
   *  - Idempotent: a bank txn that already posted payments is skipped
   *    (payments carry the reference "BANK:<txn_id>" in notes).
   *
   * Fully batched: one payments read, one write. Returns
   * { created, skipped, allocations: [{txn_id, months:[..]}], ownerSync: {unit:{monKey:amount}} }
   */
  function allocateFromBankBatch(txns, fees) {
    var sheet   = Database.getSheet(SHEET);
    var all     = Database.getAll(SHEET);
    var now     = new Date().toISOString();

    // HYBRID MODEL GUARD: transactions that have PORTIONS assigned
    // (multi-unit / multi-type / bulk edge cases) are handled purely by
    // manual month allocation — the automatic pipeline must never post
    // for them, including on re-match runs.
    var portioned = {};
    try { portioned = PortionsService.getPortionTxnIdSet(); } catch (ePor) {}

    // Existing bank refs (dedup) + paid months per unit|type
    var bankRefs = {};
    var paidSet  = {};
    for (var i = 0; i < all.length; i++) {
      var notes = String(all[i][C.NOTES] || '');
      var m = notes.match(/BANK:([\w]+)/);
      if (m) bankRefs[m[1]] = true;
      if (String(all[i][C.STATUS]) === 'Verified') {
        paidSet[String(all[i][C.UNIT_ID]).toUpperCase() + '|' + String(all[i][C.PAYMENT_TYPE]) + '|' + _normMonth(all[i][C.MONTH])] = true;
      }
    }

    var newRows     = [];
    var allocations = [];
    var ownerSync   = {};
    var tenantSync  = {};
    var created     = 0;
    var skipped     = 0;

    for (var t = 0; t < txns.length; t++) {
      var txn = txns[t];
      var amount = Number(txn.amount || 0);
      if (!txn.unit_id || !(amount > 0)) { skipped++; continue; }
      if (bankRefs[txn.txn_id])          { skipped++; continue; }
      if (portioned[txn.txn_id])         { skipped++; continue; } // portions own this txn

      var type = txn.payment_type && txn.payment_type !== 'Unknown' ? txn.payment_type : 'LPG';

      // One-time / occasional income is never split across months.
      var ONE_TIME = { 'Caution Deposit': true, 'Party Hall Rental': true, 'Bank Interest': true, 'Miscellaneous': true };

      // Fee may be a single number or a LIST [current, older1, older2…].
      // The first amount that divides the payment evenly decides the split,
      // so the current fee takes priority over historical ones.
      var feeList = fees && fees[type] ? fees[type] : [];
      if (Object.prototype.toString.call(feeList) !== '[object Array]') feeList = [Number(feeList)];

      var nMonths   = 1;
      var perAmount = amount;
      if (!ONE_TIME[type]) {
        for (var f = 0; f < feeList.length; f++) {
          var fee = Number(feeList[f]);
          if (fee > 0 && amount % fee === 0 && amount / fee >= 1) {
            nMonths   = Math.min(24, amount / fee);
            perAmount = fee;
            break;
          }
        }
      }

      var ym = _monthFromDateStr(txn.dateStr);
      if (!ym) { skipped++; continue; }  // date unreadable: do NOT post to a guessed month
      var y  = ym.y, mo = ym.m;
      // LPG convention: a payment lands the month AFTER the gas it pays for
      // (December's reading is settled in January). Book it to the CONSUMPTION
      // (previous) month so "Payment For Month" lines up with that month's
      // reading — the bank date stays the real-world truth, this is only the
      // month it is FOR. The treasurer can still override it in the edit window.
      if (type === 'LPG') { mo--; if (mo < 1) { mo = 12; y--; } }
      var unit = String(txn.unit_id).toUpperCase();
      var months = [];

      for (var k = 0; k < nMonths; k++) {
        // Walk forward past months that are already paid (max 4 years)
        var guard = 0;
        while (paidSet[unit + '|' + type + '|' + _mk(y, mo)] && guard < 48) {
          mo++; if (mo > 12) { mo = 1; y++; }
          guard++;
        }
        var monKey = _mk(y, mo);
        var payId  = Database.generateId('PAY');
        newRows.push([
          payId, unit, txn.tenant_id || '', type, perAmount, "'" + monKey,
          '', 'Verified',
          'BANK:' + txn.txn_id + (nMonths > 1 ? ' (' + (k + 1) + '/' + nMonths + ')' : ''),
          now, now, 'Bank Import'
        ]);
        paidSet[unit + '|' + type + '|' + monKey] = true;
        months.push(monKey);
        created++;

        // Sync the owner-sheet Jan–Dec columns for the current year
        if (type === 'Maintenance' && y === (new Date()).getFullYear()) {
          if (!ownerSync[unit]) ownerSync[unit] = {};
          ownerSync[unit][monKey] = perAmount;
        }
        // Sync the tenant-sheet Jan–Dec waste columns for the current year
        if (type === 'Waste Management' && y === (new Date()).getFullYear()) {
          if (!tenantSync[unit]) tenantSync[unit] = {};
          tenantSync[unit][monKey] = perAmount;
        }

        mo++; if (mo > 12) { mo = 1; y++; }
      }
      allocations.push({ txn_id: txn.txn_id, months: months });
    }

    // One batched append
    if (newRows.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, newRows.length, 12).setValues(newRows);
      SpreadsheetApp.flush();
    }

    return { created: created, skipped: skipped, allocations: allocations,
             ownerSync: ownerSync, tenantSync: tenantSync };
  }

  // '01-May-2025' / '01/05/2025' / ISO -> {y, m}; fallback: today
  function _monthFromDateStr(s) {
    var months = { JAN:1,FEB:2,MAR:3,APR:4,MAY:5,JUN:6,JUL:7,AUG:8,SEP:9,OCT:10,NOV:11,DEC:12 };
    var str = String(s || '').trim();
    var m = str.match(/^(\d{1,2})[-\s]([A-Za-z]{3})[a-z]*[-\s](\d{4})/);
    if (m && months[m[2].toUpperCase()]) return { y: Number(m[3]), m: months[m[2].toUpperCase()] };
    m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) return { y: Number(m[3]), m: Number(m[2]) };
    m = str.match(/^(\d{4})-(\d{2})/);
    if (m) return { y: Number(m[1]), m: Number(m[2]) };
    return null; // unparsable — caller must skip, never guess
  }

  function _mk(y, m) { return y + '-' + (m < 10 ? '0' + m : String(m)); }

  function _currentMonth() {
    var now = new Date();
    return now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  }

  function _validate(data) {
    if (!data.unit_id)      throw new Error('unit_id is required.');
    if (!data.payment_type) throw new Error('payment_type is required.');
    var amt = Number(data.amount);
    if (data.amount === undefined || data.amount === null || data.amount === '' || isNaN(amt) || amt < 0) throw new Error('Valid amount is required.');
    if (amt === 0 && data.status !== 'UO') throw new Error('Amount must be greater than zero (0 is only for UO markers).');
    // tenant_id is optional: maintenance and most fees are recorded against
    // the unit, not a specific tenant.
  }

  // Google Sheets auto-converts '2026-01' text into a real Date. Normalize
  // anything the sheet gives back into the canonical 'YYYY-MM' key.
  // NOTE: uses local year/month (not toISOString) to avoid timezone shifts.
  function _normMonth(v) {
    if (v === undefined || v === null || v === '') return '';
    if (v instanceof Date) {
      return v.getFullYear() + '-' + (v.getMonth() + 1 < 10 ? '0' + (v.getMonth() + 1) : String(v.getMonth() + 1));
    }
    var s = String(v).trim();
    var m = s.match(/^(\d{4})-(\d{1,2})/);
    if (m) return m[1] + '-' + (m[2].length < 2 ? '0' + m[2] : m[2]);
    var ym = _monthFromDateStr(s);
    return ym ? _mk(ym.y, ym.m) : '';
  }

  function _toObj(row) {
    function safe(i) {
      var v = row[i];
      if (v === undefined || v === null) return '';
      if (v instanceof Date) return v.toISOString();
      return String(v);
    }
    return {
      payment_id:    safe(C.PAYMENT_ID),
      unit_id:       safe(C.UNIT_ID),
      tenant_id:     safe(C.TENANT_ID),
      payment_type:  safe(C.PAYMENT_TYPE),
      amount:        safe(C.AMOUNT),
      month:         _normMonth(row[C.MONTH]),
      screenshot_url:safe(C.SCREENSHOT_URL),
      status:        safe(C.STATUS),
      notes:         safe(C.NOTES),
      submitted_at:  safe(C.SUBMITTED_AT),
      reviewed_at:   safe(C.REVIEWED_AT),
      reviewed_by:   safe(C.REVIEWED_BY)
    };
  }






  return {
    submitPayment:              submitPayment,
    verifyPayment:              verifyPayment,
    rejectPayment:              rejectPayment,
    getPaymentById:             getPaymentById,
    getPaymentsByUnit:          getPaymentsByUnit,
    getPaymentsByTenantAndMonth:getPaymentsByTenantAndMonth,
    getPendingPayments:         getPendingPayments,
    getAllPayments:              getAllPayments,
    getMonthlyPaymentHistory:   getMonthlyPaymentHistory,
    getPaymentStats:            getPaymentStats,
    updatePayment:              updatePayment,
    deleteRecord:               deleteRecord,
    getMonthSummary:            getMonthSummary,
    deleteByMonth:              deleteByMonth,
    shiftUnitMonths:            shiftUnitMonths,
    allocateFromBankBatch:      allocateFromBankBatch
  };
})();
