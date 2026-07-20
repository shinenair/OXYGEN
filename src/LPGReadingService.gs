// ═══════════════════════════════════════════════════════════════
// LPGReadingService — meter-based LPG billing with per-month rate
// history, integrated payment tracking, and batch save.
// Amount = (Current − Previous) × Conversion Factor (Kg/NM³) × Price/Kg
// "Previous" for a month auto-carries from the prior month's Current
// reading, so the user only ever types the new meter value.
// The rate used is whatever was in effect for THAT month (LPGRates),
// so past calculations never shift when today's rate changes.
// ═══════════════════════════════════════════════════════════════
var LPGReadingService = (function() {

  var SHEET = 'LPGReadings';
  var C = {
    READING_ID: 0, UNIT_ID: 1, YEAR: 2, MONTH: 3,
    PREVIOUS: 4, CURRENT: 5, READING_DATE: 6, CONSUMED: 7,
    CONV_FACTOR: 8, PRICE_PER_KG: 9, AMOUNT: 10,
    PAID_AMOUNT: 11, PAYMENT_DATE: 12, PAID_STATUS: 13,
    NOTES: 14, RECORDED_BY: 15, RECORDED_AT: 16,
    FLAGGED: 17, FLAG_NOTE: 18, FLAGGED_BY: 19, FLAGGED_AT: 20
  };
  var TOTAL_COLS = 21;
  var HEADERS = ['reading_id','unit_id','year','month','previous_reading','current_reading',
                 'reading_date','consumed','conv_factor','price_per_kg','amount',
                 'paid_amount','payment_date','paid_status','notes','recorded_by','recorded_at',
                 'flagged','flag_note','flagged_by','flagged_at'];

  function ensureColumns() {
    var sheet = Database.getSheet(SHEET);
    if (sheet.getMaxColumns() < TOTAL_COLS) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), TOTAL_COLS - sheet.getMaxColumns());
    }
    var head = sheet.getRange(1, 1, 1, TOTAL_COLS).getValues()[0];
    if (String(head[0]) !== HEADERS[0] || String(head[TOTAL_COLS - 1]) !== HEADERS[TOTAL_COLS - 1]) {
      sheet.getRange(1, 1, 1, TOTAL_COLS).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, TOTAL_COLS).setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    // Force the date columns to stay PLAIN TEXT forever. Without this,
    // Sheets auto-detects a written "2024-02-03"-looking string and
    // silently converts the cell to its own internal Date type, anchored
    // to the spreadsheet's timezone — reading it back can then drift by a
    // day depending on timezone math. Plain text sidesteps that entirely:
    // the string we wrote is exactly the string we read back, always.
    sheet.getRange(2, C.READING_DATE + 1, 5000, 1).setNumberFormat('@');
    sheet.getRange(2, C.PAYMENT_DATE + 1, 5000, 1).setNumberFormat('@');
  }

  function _rows() { return Database.getAll(SHEET); }
  function _fmtCellDate(v) {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    return String(v);
  }

  function _toObj(row) {
    function n(i) { var v = row[i]; return (v === '' || v === undefined || v === null) ? null : Number(v); }
    function s(i) { var v = row[i]; return (v === undefined || v === null) ? '' : String(v); }
    return {
      reading_id: s(C.READING_ID), unit_id: s(C.UNIT_ID),
      year: n(C.YEAR), month: n(C.MONTH),
      previous_reading: n(C.PREVIOUS), current_reading: n(C.CURRENT),
      reading_date: _fmtCellDate(row[C.READING_DATE]), consumed: n(C.CONSUMED),
      conv_factor: n(C.CONV_FACTOR), price_per_kg: n(C.PRICE_PER_KG), amount: n(C.AMOUNT),
      paid_amount: n(C.PAID_AMOUNT), payment_date: _fmtCellDate(row[C.PAYMENT_DATE]),
      paid_status: s(C.PAID_STATUS) || 'Pending',
      notes: s(C.NOTES), recorded_by: s(C.RECORDED_BY), recorded_at: s(C.RECORDED_AT),
      flagged: s(C.FLAGGED) === 'Yes', flag_note: s(C.FLAG_NOTE),
      flagged_by: s(C.FLAGGED_BY), flagged_at: s(C.FLAGGED_AT)
    };
  }

  // ── Rate history ──────────────────────────────────────────────
  var RATE_SHEET = 'LPGRates';
  var RC = { RATE_ID:0, YEAR:1, MONTH:2, CONV:3, PRICE:4, SET_BY:5, SET_AT:6 };

  function _rateRows() { return Database.getAll(RATE_SHEET); }

  // The rate in effect for {year, month}: the most recent LPGRates entry
  // with (year, month) <= requested; falls back to Settings' current
  // values if no rate history exists yet (first-time use).
  function rateFor(year, month) {
    var rows = _rateRows(), best = null, bestKey = -1;
    var reqKey = Number(year) * 100 + Number(month);
    for (var i = 0; i < rows.length; i++) {
      var key = Number(rows[i][RC.YEAR]) * 100 + Number(rows[i][RC.MONTH]);
      if (key <= reqKey && key > bestKey) { bestKey = key; best = rows[i]; }
    }
    if (best) return { conv_factor: Number(best[RC.CONV]), price_per_kg: Number(best[RC.PRICE]) };
    var settings = SettingsService.getAll();
    return {
      conv_factor: Number(settings.lpg_conv_factor) || 2.6,
      price_per_kg: Number(settings.lpg_price_per_kg) || 78.31
    };
  }

  // All rates on record, oldest first.
  function listRates() {
    var rows = _rateRows();
    var out = rows.map(function(r) {
      return { rate_id: String(r[RC.RATE_ID]), year: Number(r[RC.YEAR]), month: Number(r[RC.MONTH]),
               conv_factor: Number(r[RC.CONV]), price_per_kg: Number(r[RC.PRICE]),
               set_by: String(r[RC.SET_BY] || ''), set_at: String(r[RC.SET_AT] || '') };
    });
    out.sort(function(a, b) { return (a.year * 100 + a.month) - (b.year * 100 + b.month); });
    return out;
  }

  // Set (or replace) the rate effective from {year, month} onward.
  function setRate(year, month, convFactor, pricePerKg, setBy) {
    year = Number(year); month = Number(month);
    if (!year || !month || month < 1 || month > 12) throw new Error('Valid year and month are required.');
    var cf = Number(convFactor), pk = Number(pricePerKg);
    if (isNaN(cf) || cf <= 0) throw new Error('Conversion factor must be a positive number.');
    if (isNaN(pk) || pk <= 0) throw new Error('Price per Kg must be a positive number.');

    var sheet = Database.getSheet(RATE_SHEET);
    var rows = _rateRows();
    for (var i = 0; i < rows.length; i++) {
      if (Number(rows[i][RC.YEAR]) === year && Number(rows[i][RC.MONTH]) === month) {
        sheet.getRange(i + 2, RC.CONV + 1, 1, 2).setValues([[cf, pk]]);
        return { success: true };
      }
    }
    sheet.appendRow([Utilities.getUuid(), year, month, cf, pk, setBy || 'Manager', new Date().toISOString()]);
    return { success: true };
  }

  function deleteRate(rateId) {
    var sheet = Database.getSheet(RATE_SHEET);
    var rows = _rateRows();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][RC.RATE_ID]) === String(rateId)) { sheet.deleteRow(i + 2); return { success: true }; }
    }
    throw new Error('Rate entry not found.');
  }

  // ── Readings ──────────────────────────────────────────────────

  // All readings for one calendar month, keyed by UNIT_ID.
  // Every distinct year+month present, with a row count — for the Danger
  // Zone's "delete a specific month" table.
  function getMonthSummary() {
    var rows = _rows();
    var counts = {};
    for (var i = 0; i < rows.length; i++) {
      var y = Number(rows[i][C.YEAR]), m = Number(rows[i][C.MONTH]);
      if (!y || !m) continue;
      var key = y + '-' + (m < 10 ? '0' + m : m);
      counts[key] = (counts[key] || 0) + 1;
    }
    var out = [];
    for (var k in counts) out.push({ month: k, count: counts[k] });
    out.sort(function(a, b) { return a.month < b.month ? -1 : 1; });
    return out;
  }

  // Deletes every reading row for one year+month. Note: this does NOT
  // touch any OTHER month's "previous reading" auto-carry chain — deleting
  // Jan's row means the next saved reading after Jan will simply look
  // further back for its previous value, same as if Jan had never been read.
  function deleteReadingsByMonth(year, month) {
    var sheet = Database.getSheet(SHEET);
    var rows = _rows();
    var deleted = 0;
    for (var i = rows.length - 1; i >= 0; i--) {
      if (Number(rows[i][C.YEAR]) === Number(year) && Number(rows[i][C.MONTH]) === Number(month)) {
        sheet.deleteRow(i + 2); deleted++;
      }
    }
    return { success: true, deleted: deleted };
  }

  function getMonth(year, month) {
    var rows = _rows(), out = {};
    for (var i = 0; i < rows.length; i++) {
      if (Number(rows[i][C.YEAR]) === Number(year) && Number(rows[i][C.MONTH]) === Number(month)) {
        var o = _toObj(rows[i]);
        out[o.unit_id.toUpperCase()] = o;
      }
    }
    return out;
  }

  // For every unit that has NO reading row yet for {year, month}, the most
  // recent prior current-reading (the value that WOULD carry forward).
  // Lets the page show "last: 92.957" on a still-pending row.
  function getPreviousHints(year, month) {
    var rows = _rows(), cache = {};
    for (var i = 0; i < rows.length; i++) {
      var o = _toObj(rows[i]);
      var k = o.unit_id.toUpperCase();
      if (!cache[k]) cache[k] = [];
      cache[k].push(o);
    }
    var reqKey = Number(year) * 100 + Number(month), out = {};
    for (var uk in cache) {
      var hasThisMonth = false, best = null, bestKey = -1;
      var list = cache[uk];
      for (var j = 0; j < list.length; j++) {
        if (list[j].year === Number(year) && list[j].month === Number(month)) hasThisMonth = true;
        var key = list[j].year * 100 + list[j].month;
        if (key < reqKey && key > bestKey && list[j].current_reading !== null) { bestKey = key; best = list[j].current_reading; }
      }
      if (!hasThisMonth && best !== null) out[uk] = best;
    }
    return out;
  }

  // The comparison this powers: a reading taken in month M is what the
  // resident is BILLED for that month's gas — but by the association's
  // own real-world billing rhythm, the matching PAYMENT always lands in
  // month M+1 (December's reading is paid in January, etc). So "Paid"
  // here is never the stale value once written to this row — it's a
  // LIVE lookup, every time, against the actual LPG payments recorded
  // in Fees Received for the following month. That's what keeps this
  // honest as new payments come in, instead of freezing at whatever an
  // old import happened to say.
  function getUnitHistory(unitId) {
    var unit = String(unitId).toUpperCase();
    var rows = _rows(), out = [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.UNIT_ID]).toUpperCase() === unit) out.push(_toObj(rows[i]));
    }
    out.sort(function(a, b) { return (b.year * 100 + b.month) - (a.year * 100 + a.month); });

    // The real payment date is the Bank Statement transaction's date, not
    // when the payment RECORD was created (which is "today" for a just-run
    // bank import). Map every bank txn id (both accounts) to its date once.
    var txnDate = {};
    try { BankService.getAllTransactions(null).forEach(function(t) { if (t.txn_id) txnDate[String(t.txn_id)] = t.date; }); } catch (eB1) {}
    try { Bank2Service.getAllTransactions(null).forEach(function(t) { if (t.txn_id) txnDate[String(t.txn_id)] = t.date; }); } catch (eB2) {}

    // One call for every LPG payment this unit has ever made, then
    // grouped by month — far cheaper than a lookup per reading row.
    var payMap = {};
    try {
      PaymentsService.getAllPayments({ unit_id: unit, payment_type: 'LPG' }).forEach(function(p) {
        if (p.status === 'Rejected') return;
        var k = p.month;
        if (!payMap[k]) payMap[k] = { amount: 0, dates: [] };
        payMap[k].amount += Number(p.amount) || 0;
        // Prefer the bank transaction's own date; fall back to the record's
        // creation time only for payments not posted from a bank line.
        var ref = String(p.notes || '').match(/BANK:(\S+?)(?:[\s(]|$)/);
        var dstr = (ref && txnDate[ref[1]]) ? txnDate[ref[1]] : p.submitted_at;
        var ymd = _lpgYmd(dstr);
        if (ymd) payMap[k].dates.push(ymd);
      });
    } catch (e) {}

    var today = new Date();
    var todayKey = today.getFullYear() * 100 + (today.getMonth() + 1);

    out.forEach(function(r) {
      // New convention: an LPG payment is stamped with the CONSUMPTION month
      // it pays for — so a reading for month M pairs with a payment whose
      // "For Month" is M itself. Legacy records (posted before this change)
      // are still stamped with the following month (M+1), so we accept EITHER
      // and prefer the exact-month match. This means no risky bulk re-label:
      // old data keeps working, new/corrected data lines up cleanly.
      var sameKey = r.year + '-' + (r.month < 10 ? '0' + r.month : r.month);
      var nextY = r.month === 12 ? r.year + 1 : r.year;
      var nextM = r.month === 12 ? 1 : r.month + 1;
      var nextKey = nextY + '-' + (nextM < 10 ? '0' + nextM : nextM);
      var pay = payMap[sameKey] || payMap[nextKey];
      // The actual month recorded against the payment on the Fees Received
      // page — surfaced verbatim so the history shows the treasurer's truth.
      r.payment_for_month = payMap[sameKey] ? sameKey : (payMap[nextKey] ? nextKey : '');

      if (pay) {
        var amt = Number(pay.amount);
        r.live_paid_amount = isFinite(amt) ? Math.round(amt * 100) / 100 : null;
        // Dates are already normalized to YYYY-MM-DD, so a lexical sort is
        // chronological; the latest is the representative payment date.
        r.live_payment_date = pay.dates.length ? pay.dates.sort().slice(-1)[0] : '';
        r.live_status = (r.live_paid_amount !== null && Math.abs(r.live_paid_amount - (r.amount || 0)) > 0.5) ? 'Mismatch' : 'Paid';
      } else {
        r.live_paid_amount = null;
        r.live_payment_date = '';
        // Only truly "pending" if next month hasn't happened yet or has
        // only just started — otherwise it's an unpaid month that's
        // actually overdue, which should still surface, just not be
        // confused with "not due yet".
        var nextKeyNum = nextY * 100 + nextM;
        r.live_status = nextKeyNum >= todayKey ? 'Pending' : 'Unpaid';
      }
    });

    return out;
  }

  // Admin-only: flags a specific reading's payment comparison as
  // incorrect, with a note — purely informational, never adjusts any
  // amount or payment automatically. A human decision recorded, not an
  // automatic fix.
  function setFlag(readingId, flagged, note) {
    UsersService.requireAdmin();
    var found = Database.findByColumn(SHEET, C.READING_ID, readingId);
    if (!found) throw new Error('Reading not found: ' + readingId);
    var row = found.data;
    row[C.FLAGGED] = flagged ? 'Yes' : 'No';
    row[C.FLAG_NOTE] = flagged ? String(note || '') : '';
    row[C.FLAGGED_BY] = flagged ? (Session.getActiveUser().getEmail() || 'Admin') : '';
    row[C.FLAGGED_AT] = flagged ? new Date().toISOString() : '';
    Database.updateRow(SHEET, found.rowIndex, row);
    return { success: true };
  }

  // Normalize any date the bank or a payment record gives back — ISO
  // ("2025-05-01T…"), DD-Mon-YYYY ("01-May-2025"), or DD/MM/YYYY — to a
  // canonical YYYY-MM-DD, so dates sort and display consistently. Unknown
  // shapes pass through unchanged (formatDate handles them client-side).
  function _lpgYmd(v) {
    if (v === undefined || v === null || v === '') return '';
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var s = String(v).trim();
    var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
    var mon = { jan:'01', feb:'02', mar:'03', apr:'04', may:'05', jun:'06', jul:'07', aug:'08', sep:'09', oct:'10', nov:'11', dec:'12' };
    var m = s.match(/^(\d{1,2})[-\/\s]([A-Za-z]{3})[a-z]*[-\/\s](\d{4})/);
    if (m && mon[m[2].toLowerCase()]) return m[3] + '-' + mon[m[2].toLowerCase()] + '-' + (m[1].length < 2 ? '0' + m[1] : m[1]);
    var n = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if (n) return n[3] + '-' + (n[2].length < 2 ? '0' + n[2] : n[2]) + '-' + (n[1].length < 2 ? '0' + n[1] : n[1]);
    return s;
  }

  // The most recent CURRENT reading strictly before {year, month}.
  function _priorCurrent(unitId, year, month, cache) {
    var reqKey = year * 100 + month, best = null, bestKey = -1;
    var list = cache[unitId] || [];
    for (var i = 0; i < list.length; i++) {
      var key = list[i].year * 100 + list[i].month;
      if (key < reqKey && key > bestKey && list[i].current_reading !== null) { bestKey = key; best = list[i].current_reading; }
    }
    return best;
  }

  function _buildUnitCache() {
    var rows = _rows(), cache = {};
    for (var i = 0; i < rows.length; i++) {
      var o = _toObj(rows[i]);
      var k = o.unit_id.toUpperCase();
      if (!cache[k]) cache[k] = [];
      cache[k].push(o);
    }
    return cache;
  }

  // Save (insert or update) one unit's reading for one month.
  // rateOverride lets the importer stamp the historical rate straight
  // from the source spreadsheet instead of looking one up.
  function _saveOne(data, cache) {
    var unitId = String(data.unit_id || '').toUpperCase();
    if (!unitId) throw new Error('unit_id is required.');
    var year = Number(data.year), month = Number(data.month);
    if (!year || !month || month < 1 || month > 12) throw new Error('Valid year and month are required for ' + unitId + '.');
    if (data.current_reading === '' || data.current_reading === null || data.current_reading === undefined) {
      throw new Error('Current reading is required for ' + unitId + ' — use Delete to remove an entry instead of leaving it blank.');
    }
    var current = Number(data.current_reading);
    if (isNaN(current)) throw new Error('Current reading must be a number for ' + unitId + '.');

    var previous = (data.previous_reading !== undefined && data.previous_reading !== null && data.previous_reading !== '')
      ? Number(data.previous_reading)
      : _priorCurrent(unitId, year, month, cache);
    if (previous === null) previous = current; // first-ever reading: zero consumption

    var consumed = Math.round((current - previous) * 1000) / 1000;
    if (consumed < -0.001) throw new Error(unitId + ': current reading (' + current + ') is lower than previous (' + previous + ').');
    if (consumed < 0) consumed = 0;

    var rate = (data.conv_factor && data.price_per_kg)
      ? { conv_factor: Number(data.conv_factor), price_per_kg: Number(data.price_per_kg) }
      : rateFor(year, month);
    var amount = Math.round(consumed * rate.conv_factor * rate.price_per_kg * 100) / 100;

    return {
      unit_id: unitId, year: year, month: month, previous: previous, current: current,
      reading_date: data.reading_date || '', consumed: consumed,
      conv_factor: rate.conv_factor, price_per_kg: rate.price_per_kg, amount: amount,
      paid_amount: (data.paid_amount !== undefined && data.paid_amount !== null && data.paid_amount !== '') ? Number(data.paid_amount) : '',
      payment_date: data.payment_date || '',
      // Derived, not manually chosen: a populated payment date means paid.
      paid_status: data.payment_date ? 'Verified' : 'Pending',
      notes: data.notes || '', recorded_by: data.recorded_by || 'Manager'
    };
  }

  function saveReading(data) {
    checkYearEditable(data.year, data.confirmed_historical);
    var cache = _buildUnitCache();

    // Overriding the previous reading (to a DIFFERENT value than what is
    // already stored, or than auto-carry would compute) changes Consumed
    // and Amount for that month — only an Administrator may do that.
    // Editing payment/remarks alone re-sends the EXISTING previous_reading
    // unchanged, which must NOT trigger this — only a genuine change should.
    if (data.previous_reading !== undefined && data.previous_reading !== null && data.previous_reading !== '') {
      var unitKey = String(data.unit_id || '').toUpperCase();
      var existing = null;
      (cache[unitKey] || []).forEach(function(r) {
        if (r.year === Number(data.year) && r.month === Number(data.month)) existing = r;
      });
      var incoming = Number(data.previous_reading);
      var isRealChange = !existing || Number(existing.previous_reading) !== incoming;
      if (isRealChange) {
        var callerEmail = Session.getActiveUser().getEmail();
        if (typeof UsersService !== 'undefined' && UsersService.getRole(callerEmail) !== 'admin') {
          throw new Error('Only an Administrator can override the previous reading.');
        }
      }
    }

    var built = _saveOne(data, cache);
    _writeRows([built]);
    return { success: true, previous_reading: built.previous, consumed: built.consumed, amount: built.amount };
  }

  // Save many readings in ONE pass — used by the "Save All Changes" button
  // and the Excel importer, avoiding a round trip per row.
  function saveReadingsBatch(list) {
    if (!list || !list.length) return { success: true, saved: 0, errors: [] };
    var cache = _buildUnitCache();
    var built = [], errors = [];
    for (var i = 0; i < list.length; i++) {
      try { built.push(_saveOne(list[i], cache)); }
      catch (e) { errors.push((list[i].unit_id || '?') + ': ' + e.message); }
    }
    if (built.length) _writeRows(built);
    return { success: true, saved: built.length, errors: errors };
  }

  function _writeRows(built) {
    var sheet = Database.getSheet(SHEET);
    var existingRows = _rows();
    var indexByKey = {}; // 'UNIT|year|month' -> sheet row number
    for (var i = 0; i < existingRows.length; i++) {
      var r = existingRows[i];
      indexByKey[String(r[C.UNIT_ID]).toUpperCase() + '|' + r[C.YEAR] + '|' + r[C.MONTH]] = i + 2;
    }
    var now = new Date().toISOString();
    for (var b = 0; b < built.length; b++) {
      var x = built[b];
      var key = x.unit_id + '|' + x.year + '|' + x.month;
      var readingId = indexByKey[key] ? sheet.getRange(indexByKey[key], 1).getValue() : Utilities.getUuid();
      var row = [readingId, x.unit_id, x.year, x.month, x.previous, x.current,
                 x.reading_date, x.consumed, x.conv_factor, x.price_per_kg, x.amount,
                 x.paid_amount, x.payment_date, x.paid_status, x.notes, x.recorded_by, now];
      if (indexByKey[key]) sheet.getRange(indexByKey[key], 1, 1, row.length).setValues([row]);
      else { sheet.appendRow(row); indexByKey[key] = sheet.getLastRow(); }
    }
  }

  function deleteReading(readingId) {
    var sheet = Database.getSheet(SHEET);
    var rows = _rows();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.READING_ID]) === String(readingId)) { sheet.deleteRow(i + 2); return { success: true }; }
    }
    throw new Error('Reading not found.');
  }

  return {
    ensureColumns:      ensureColumns,
    getMonth:           getMonth,
    getMonthSummary:    getMonthSummary,
    deleteReadingsByMonth: deleteReadingsByMonth,
    getUnitHistory:     getUnitHistory,
    getPreviousHints:   getPreviousHints,
    saveReading:        saveReading,
    saveReadingsBatch:  saveReadingsBatch,
    deleteReading:      deleteReading,
    rateFor:            rateFor,
    listRates:          listRates,
    setRate:            setRate,
    deleteRate:         deleteRate,
    setFlag:            setFlag
  };
})();
