// ── Shared, timezone-agnostic date parser ──────────────────────────
// Global (not inside any IIFE) so every importer in the project can use
// it: LPG readings, Stock Inward/Outward, Inventory-vs-Bank.
//
// This NEVER routes through JavaScript's `new Date(string)` for anything
// but a final ISO check, because that constructor's interpretation of a
// string depends on its exact punctuation (dashes = UTC, slashes = local
// time) — a landmine that previously shifted dates by a day under IST.
// Every format is instead parsed by hand, digit by digit, so the result
// is always exactly the calendar day the sheet intended, regardless of
// server/spreadsheet timezone. Malformed input (e.g. a typo'd year) is
// rejected outright rather than silently imported as a wrong date.
function _pad2(n) { return (n < 10 ? '0' : '') + n; }
function parseFlexibleDate(v) {
  if (v === '' || v === null || v === undefined) return '';
  var s = String(v).trim();

  // Pure number, no separators -> Excel/Sheets serial (epoch 1899-12-30).
  // A serial has no timezone concept at all, so UTC getters are correct.
  if (/^\d+(\.\d+)?$/.test(s)) {
    var n = Number(s);
    var ms = Math.round((n - 25569) * 86400 * 1000);
    var d = new Date(ms);
    if (isNaN(d.getTime())) return '';
    var y = d.getUTCFullYear();
    if (y < 1990 || y > 2100) return ''; // sanity check — rejects garbage serials
    return y + '-' + _pad2(d.getUTCMonth() + 1) + '-' + _pad2(d.getUTCDate());
  }

  // "31-Oct-2024" / "03 Feb 24" style.
  var mn = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
  var m = s.toLowerCase().match(/^(\d{1,2})[-\/ ]([a-z]{3})[a-z]*[-\/ ](\d{2,4})$/);
  if (m && mn[m[2]]) {
    var day = parseInt(m[1], 10), mo = mn[m[2]], yr = parseInt(m[3], 10);
    if (yr < 100) yr += 2000;
    if (yr < 1990 || yr > 2100 || day < 1 || day > 31) return ''; // catches typo'd years like "0243"
    return yr + '-' + _pad2(mo) + '-' + _pad2(day);
  }

  // "24/12/2024" or "24-12-2024" — numeric, day-first (Indian convention),
  // with a same-day-swap fallback if the first number can't be a month.
  m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if (m) {
    var a = parseInt(m[1], 10), b = parseInt(m[2], 10), yr2 = parseInt(m[3], 10);
    if (yr2 < 100) yr2 += 2000;
    var day2 = a, mo2 = b;
    if (b > 12 && a <= 12) { day2 = b; mo2 = a; }
    if (yr2 < 1990 || yr2 > 2100 || mo2 < 1 || mo2 > 12 || day2 < 1 || day2 > 31) return '';
    return yr2 + '-' + _pad2(mo2) + '-' + _pad2(day2);
  }

  // Already ISO "yyyy-MM-dd", optionally with a trailing time component
  // ("yyyy-MM-dd HH:MM:SS" or "...THH:MM:SS") — a real Google Sheets CSV
  // export of a date-type cell with no time set still writes it as
  // "2026-02-01 00:00:00", not the bare date. The time is irrelevant
  // here and simply discarded once matched.
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$/);
  if (m) return m[1] + '-' + _pad2(parseInt(m[2], 10)) + '-' + _pad2(parseInt(m[3], 10));

  return ''; // unparseable — caller should treat as missing, not guess
}

// ═══════════════════════════════════════════════════════════════
// LPGImport — bulk-import historical LPG readings from the
// association's original Excel format (FLAT NO / Name / Previous /
// Current / Reading Date / Volume / Factor / Price / Amount /
// Payment Date / Paid / Remarks).
// ═══════════════════════════════════════════════════════════════
var LPGImport = (function() {

  function _serialToDate(v) { return parseFlexibleDate(v); }

  // rows: array of arrays exactly as pasted/parsed from the sheet, one
  // row per unit, columns matching the original template:
  // [#, FLAT NO, Name, _, _, _, Previous, Current, ReadingDateSerial,
  //  Volume, Factor, Price, Amount, PaymentDateSerial, Paid, Remarks]
  // OR the "UNOCCUPIED" short form: [#, FLAT NO, 'UNOCCUPIED']
  function importFromRows(rows, year, month) {
    year = Number(year); month = Number(month);
    if (!year || !month || month < 1 || month > 12) throw new Error('Valid year and month are required.');
    if (!rows || !rows.length) throw new Error('No rows to import.');

    var toSave = [];
    var skippedUO = 0, skippedNoReading = 0, errors = [];

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r || !r.length) continue;
      var unitId = String(r[1] || '').trim().toUpperCase();
      if (!unitId || !/^[AB]\d{3}$/.test(unitId)) continue; // skip header/blank rows

      var name = String(r[2] || '').trim();
      if (name.toUpperCase() === 'UNOCCUPIED') { skippedUO++; continue; }

      var previous = r[6], current = r[7];
      if (current === '' || current === undefined || current === null) { skippedNoReading++; continue; }

      var readingDate = _serialToDate(r[8]);
      var priceCell   = r[11]; // e.g. "₹78.31" or 78.31
      var paymentDate = _serialToDate(r[13]);
      var paidFlag    = String(r[14] || '').trim().toLowerCase();
      var remarks     = String(r[15] || '').trim();

      var price = Number(String(priceCell || '').replace(/[₹,\s]/g, ''));
      var factor = Number(r[10]) || 2.6;

      var paidStatus = 'Pending';
      var paidAmount = '';
      if (paidFlag === 'yes') {
        paidStatus = 'Verified';
        // Default: the sheet's own computed Amount column.
        var amtCell = Number(String(r[12] || '').replace(/[₹,\s]/g, ''));
        paidAmount = isNaN(amtCell) ? '' : amtCell;
        // But if the Remarks explicitly state a different figure actually
        // paid (e.g. "Paid 446 on 1Feb2024" against a computed ₹466), that
        // real paid amount takes precedence — this is exactly the kind of
        // calculated-vs-paid mismatch the reading history is meant to flag.
        var m = remarks.match(/paid\s*[:₹]?\s*([\d][\d,]*)/i);
        if (m) {
          var stated = Number(m[1].replace(/,/g, ''));
          if (!isNaN(stated) && stated > 0) paidAmount = stated;
        }
      }

      var entry = {
        unit_id: unitId, year: year, month: month,
        previous_reading: (previous === '' || previous === undefined) ? undefined : Number(previous),
        current_reading: Number(current),
        reading_date: readingDate,
        paid_amount: paidAmount, payment_date: paymentDate, paid_status: paidStatus,
        notes: remarks, recorded_by: 'Import'
      };
      // Stamp the historical rate straight from the source file so the
      // amount matches the original spreadsheet exactly, rather than
      // whatever today's rate happens to be.
      if (!isNaN(price) && price > 0) { entry.conv_factor = factor; entry.price_per_kg = price; }

      toSave.push(entry);
    }

    if (!toSave.length) {
      return { success: true, imported: 0, skippedUO: skippedUO, skippedNoReading: skippedNoReading, errors: errors };
    }
    var result = LPGReadingService.saveReadingsBatch(toSave);
    return {
      success: true, imported: result.saved,
      skippedUO: skippedUO, skippedNoReading: skippedNoReading,
      errors: result.errors || []
    };
  }

  return { importFromRows: importFromRows };
})();
