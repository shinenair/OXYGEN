// ============================================================
// BankPosting.gs — Posts matched bank transactions as payments
//
// For every matched credit (owner / tenant / unit_hint) with a
// known payment type, this module:
//   1. Reads the transaction DATE -> year + month
//   2. Allocates the payment to the first UNPAID month starting
//      from the transaction month (e.g. Jan already paid ->
//      posts to Feb), checking BOTH the Payments ledger and the
//      owner's Jan–Dec grid
//   3. Writes a Verified record into the Payments sheet — the
//      permanent historic ledger (month stored as YYYY-MM)
//   4. Updates the owner's pay_<month> column so the Owners page
//      and Unit Profile reflect it immediately
//
// Idempotent: each posting is tagged BANK:<txn_id> in the notes,
// so re-imports and re-matching never double-post.
// All reads/writes are batched (safe for hundreds of txns).
// ============================================================

var BankPosting = (function() {

  var MONTHS      = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  var MONTH_INDEX = { JAN:1, FEB:2, MAR:3, APR:4, MAY:5, JUN:6, JUL:7, AUG:8, SEP:9, OCT:10, NOV:11, DEC:12 };

  // Owners sheet column layout (must match OwnersService)
  var OWN_UNIT_ID  = 1;
  var OWN_PAY_JAN  = 31;    // pay_jan ... pay_dec = 31..42
  var OWN_TOTAL    = 45;

  // Payments sheet column layout (must match PaymentsService)
  var PAY_COLS = 12;

  /**
   * txns: array of { txn_id, date, credit, payment_type, match_type, match_id, match_unit }
   * Returns { posted, skipped, allocations:[{txn_id,unit,type,month,amount}], errors }
   */
  function postBatch(txns) {
    var result = { posted: 0, skipped: 0, allocations: [], errors: [] };
    if (!txns || !txns.length) return result;

    // ── 1. Read the Payments ledger once ──────────────────────
    var payRows = Database.getAll('Payments');
    var paidKey = {};   // 'UNIT|TYPE|YYYY-MM' -> true (Verified only)
    var bankRef = {};   // txn ids already posted
    for (var i = 0; i < payRows.length; i++) {
      var pr = payRows[i];
      var status = String(pr[7] || '');
      if (status === 'Verified') {
        paidKey[String(pr[1]) + '|' + String(pr[3]) + '|' + String(pr[5])] = true;
      }
      var notes = String(pr[8] || '');
      var bm = notes.match(/BANK:(\S+)/);
      if (bm) bankRef[bm[1]] = true;
    }

    // ── 2. Read Owners once ────────────────────────────────────
    var ownerRows = Database.getAll('Owners');
    var ownerByUnit = {};
    for (var i = 0; i < ownerRows.length; i++) {
      while (ownerRows[i].length < OWN_TOTAL) ownerRows[i].push('');
      var u = String(ownerRows[i][OWN_UNIT_ID]).trim().toUpperCase();
      if (u && !ownerByUnit.hasOwnProperty(u)) ownerByUnit[u] = i;
    }

    var now            = new Date().toISOString();
    var newPayRows     = [];
    var ownersChanged  = false;

    // ── 3. Process each transaction ────────────────────────────
    for (var t = 0; t < txns.length; t++) {
      var txn = txns[t];

      var amount = Number(txn.credit || 0);
      var unit   = String(txn.match_unit || '').trim().toUpperCase();
      var type   = String(txn.payment_type || '');

      // Only post credits with a known unit and a real payment type
      if (!txn.txn_id || amount <= 0 || !unit ||
          (type !== 'Maintenance' && type !== 'Waste Management' && type !== 'LPG')) {
        result.skipped++;
        continue;
      }
      if (bankRef[txn.txn_id]) { result.skipped++; continue; }  // already posted

      var ym = _parseDate(txn.date);
      if (!ym) {
        result.errors.push(txn.txn_id + ': could not read date "' + txn.date + '"');
        result.skipped++;
        continue;
      }

      // ── Month allocation ─────────────────────────────────────
      var ownerIdx = ownerByUnit.hasOwnProperty(unit) ? ownerByUnit[unit] : -1;
      var allocY = ym.y, allocM = ym.m;

      if (type === 'Maintenance' || type === 'Waste Management') {
        var found = false;
        for (var k = 0; k < 12; k++) {
          var y2 = ym.y + Math.floor((ym.m - 1 + k) / 12);
          var m2 = ((ym.m - 1 + k) % 12) + 1;
          var key = unit + '|' + type + '|' + _ym(y2, m2);

          var ledgerPaid = paidKey[key] === true;
          // Owner Jan–Dec grid counts as "paid" for Maintenance,
          // within the transaction's own year (the grid is yearless)
          var gridPaid = false;
          if (type === 'Maintenance' && ownerIdx >= 0 && y2 === ym.y) {
            gridPaid = String(ownerRows[ownerIdx][OWN_PAY_JAN + m2 - 1] || '') !== '';
          }

          if (!ledgerPaid && !gridPaid) { allocY = y2; allocM = m2; found = true; break; }
        }
        if (!found) { allocY = ym.y; allocM = ym.m; }  // everything paid — book on txn month
      }
      // LPG: always the transaction month (variable amounts, no rolling)

      var monthStr = _ym(allocY, allocM);

      // ── Build ledger record ──────────────────────────────────
      var payRow = [];
      for (var c = 0; c < PAY_COLS; c++) payRow.push('');
      payRow[0]  = Database.generateId('PAY');
      payRow[1]  = unit;
      payRow[2]  = (txn.match_type === 'tenant') ? String(txn.match_id || '') : '';
      payRow[3]  = type;
      payRow[4]  = amount;
      payRow[5]  = monthStr;
      payRow[7]  = 'Verified';
      payRow[8]  = 'BANK:' + txn.txn_id + ' · txn date ' + String(txn.date);
      payRow[9]  = now;
      payRow[10] = now;
      payRow[11] = 'Bank Auto-Post';
      newPayRows.push(payRow);

      paidKey[unit + '|' + type + '|' + monthStr] = true;
      bankRef[txn.txn_id] = true;

      // ── Update the owner's Jan–Dec grid (Maintenance, same year) ──
      if (type === 'Maintenance' && ownerIdx >= 0 && allocY === ym.y) {
        ownerRows[ownerIdx][OWN_PAY_JAN + allocM - 1] = amount;
        ownerRows[ownerIdx][OWN_TOTAL - 1] = now;  // updated_at
        ownersChanged = true;
      }

      result.posted++;
      result.allocations.push({
        txn_id: txn.txn_id, unit: unit, type: type,
        month: monthStr, amount: amount
      });
    }

    // ── 4. Batched writes ──────────────────────────────────────
    if (newPayRows.length > 0) {
      var pSheet = Database.getSheet('Payments');
      if (pSheet.getMaxColumns() < PAY_COLS) {
        pSheet.insertColumnsAfter(pSheet.getMaxColumns(), PAY_COLS - pSheet.getMaxColumns());
      }
      pSheet.getRange(pSheet.getLastRow() + 1, 1, newPayRows.length, PAY_COLS)
        .setValues(newPayRows);
    }
    if (ownersChanged && ownerRows.length > 0) {
      var oSheet = Database.getSheet('Owners');
      if (oSheet.getMaxColumns() < OWN_TOTAL) {
        oSheet.insertColumnsAfter(oSheet.getMaxColumns(), OWN_TOTAL - oSheet.getMaxColumns());
      }
      oSheet.getRange(2, 1, ownerRows.length, OWN_TOTAL).setValues(ownerRows);
    }
    if (newPayRows.length > 0 || ownersChanged) SpreadsheetApp.flush();

    return result;
  }

  // "01-May-2025", "2025-05-01", or a Date -> { y, m }
  function _parseDate(d) {
    if (!d) return null;
    if (d instanceof Date && !isNaN(d.getTime())) {
      return { y: d.getFullYear(), m: d.getMonth() + 1 };
    }
    var s = String(d).trim();
    var m1 = s.match(/^(\d{1,2})-([A-Za-z]{3})[A-Za-z]*-(\d{4})/);   // 01-May-2025
    if (m1) {
      var mi = MONTH_INDEX[m1[2].toUpperCase().slice(0, 3)];
      if (mi) return { y: parseInt(m1[3], 10), m: mi };
    }
    var m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);                // 2025-05-01
    if (m2) return { y: parseInt(m2[1], 10), m: parseInt(m2[2], 10) };
    var m3 = s.match(/^(\d{1,2})[\/](\d{1,2})[\/](\d{4})/);          // 01/05/2025 (DD/MM/YYYY)
    if (m3) return { y: parseInt(m3[3], 10), m: parseInt(m3[2], 10) };
    var dt = new Date(s);
    if (!isNaN(dt.getTime())) return { y: dt.getFullYear(), m: dt.getMonth() + 1 };
    return null;
  }

  function _ym(y, m) {
    return y + '-' + (m < 10 ? '0' + m : String(m));
  }

  return { postBatch: postBatch };
})();
