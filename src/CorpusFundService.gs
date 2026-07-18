// ═══════════════════════════════════════════════════════════════
// CorpusFundService.gs — Corpus Fund tracking (Admin-only).
//
// SECURITY: every single function here checks UsersService.requireAdmin()
// itself, in addition to the router's own admin gate — this data must
// never be reachable by a non-admin, whether through the UI or by
// calling the underlying function directly. Belt and suspenders,
// deliberately, given how sensitive this information is.
//
// MODEL: CDOA-wide owed/received summary (simple, admin-maintained
// figures — no per-unit tracking here, that stays exactly as it is on
// Unit Profile). Then any number of Fixed Deposits, each an entirely
// separate deposit with IOB (a new registration payment becomes a
// NEW FD, never added to an existing one), each with its own
// "Statement of Account" — a small ledger of Deposit/Interest/
// Withdrawal lines. Every FD's current balance is computed live by
// summing its own lines — never stored, never goes stale.
// ═══════════════════════════════════════════════════════════════
var CorpusFundService = (function() {
  var OWED_SHEET = 'CorpusOwed';
  var OWED_HEADERS = ['flats_owed', 'amount_per_flat', 'updated_at'];

  var RECEIVED_SHEET = 'CorpusReceived';
  var RECEIVED_HEADERS = ['row_id', 'year', 'flats_paid', 'amount_per_flat', 'created_at', 'updated_at'];
  var RC = { ID: 0, YEAR: 1, FLATS: 2, RATE: 3, CREATED: 4, UPDATED: 5 };

  var FD_SHEET = 'CorpusFixedDeposits';
  var FD_HEADERS = ['fd_id', 'label', 'bank_name', 'amount', 'deposit_date', 'maturity_date', 'interest_rate', 'status', 'created_at', 'updated_at'];
  var FC = { ID: 0, LABEL: 1, BANK: 2, AMOUNT: 3, DEPOSIT: 4, MATURITY: 5, RATE: 6, STATUS: 7, CREATED: 8, UPDATED: 9 };

  var LINE_SHEET = 'CorpusFDLines';
  var LINE_HEADERS = ['line_id', 'fd_id', 'date', 'description', 'amount', 'created_at'];
  var LC = { ID: 0, FD: 1, DATE: 2, DESC: 3, AMOUNT: 4, CREATED: 5 };

  // Deposit and Interest add to the balance; Withdrawal subtracts —
  // this is what lets the running balance compute itself with no
  // manual bookkeeping.
  var DESC_SIGN = { 'Deposit': 1, 'Interest': 1, 'Withdrawal': -1 };

  function ensureSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    function ensure(name, headers) {
      var sh = ss.getSheetByName(name);
      if (!sh) {
        sh = ss.insertSheet(name);
        sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
        sh.setFrozenRows(1);
      }
      return sh;
    }
    ensure(OWED_SHEET, OWED_HEADERS);
    ensure(RECEIVED_SHEET, RECEIVED_HEADERS);
    ensure(FD_SHEET, FD_HEADERS);
    ensure(LINE_SHEET, LINE_HEADERS);
  }

  // ── CDOA-wide "Owed" — one editable row, no history needed ──
  function getOwed() {
    UsersService.requireAdmin();
    ensureSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OWED_SHEET);
    if (sh.getLastRow() < 2) return { flats_owed: 136, amount_per_flat: 10000 };
    var row = sh.getRange(2, 1, 1, 3).getValues()[0];
    return { flats_owed: Number(row[0]) || 0, amount_per_flat: Number(row[1]) || 0 };
  }

  function setOwed(flatsOwed, amountPerFlat) {
    UsersService.requireAdmin();
    ensureSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OWED_SHEET);
    var now = new Date().toISOString();
    if (sh.getLastRow() < 2) sh.appendRow([Number(flatsOwed) || 0, Number(amountPerFlat) || 0, now]);
    else sh.getRange(2, 1, 1, 3).setValues([[Number(flatsOwed) || 0, Number(amountPerFlat) || 0, now]]);
    return { success: true };
  }

  // ── CDOA-wide "Received" — one row per year, admin-maintained ──
  function getReceived() {
    UsersService.requireAdmin();
    ensureSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECEIVED_SHEET);
    if (sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues()
      .map(function(r) { return { row_id: String(r[RC.ID]), year: String(r[RC.YEAR]), flats_paid: Number(r[RC.FLATS]) || 0, amount_per_flat: Number(r[RC.RATE]) || 0 }; })
      .filter(function(r) { return r.row_id; })
      .sort(function(a, b) { return a.year < b.year ? -1 : 1; });
  }

  function addReceivedRow(year, flatsPaid, amountPerFlat) {
    UsersService.requireAdmin();
    if (!year) throw new Error('Year is required.');
    ensureSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECEIVED_SHEET);
    var now = new Date().toISOString();
    var id = Database.generateId('CFR');
    sh.appendRow([id, String(year), Number(flatsPaid) || 0, Number(amountPerFlat) || 0, now, now]);
    return { success: true, row_id: id };
  }

  function updateReceivedRow(rowId, year, flatsPaid, amountPerFlat) {
    UsersService.requireAdmin();
    var found = Database.findByColumn(RECEIVED_SHEET, RC.ID, rowId);
    if (!found) throw new Error('Row not found: ' + rowId);
    var row = found.data;
    if (year !== undefined) row[RC.YEAR] = String(year);
    if (flatsPaid !== undefined) row[RC.FLATS] = Number(flatsPaid) || 0;
    if (amountPerFlat !== undefined) row[RC.RATE] = Number(amountPerFlat) || 0;
    row[RC.UPDATED] = new Date().toISOString();
    Database.updateRow(RECEIVED_SHEET, found.rowIndex, row);
    return { success: true };
  }

  function deleteReceivedRow(rowId) {
    UsersService.requireAdmin();
    var found = Database.findByColumn(RECEIVED_SHEET, RC.ID, rowId);
    if (!found) throw new Error('Row not found: ' + rowId);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECEIVED_SHEET).deleteRow(found.rowIndex);
    return { success: true };
  }

  // ── Fixed Deposits ──
  function _fdRows() {
    ensureSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FD_SHEET);
    if (sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, FD_HEADERS.length).getValues();
  }

  function _fdToObj(r) {
    return { fd_id: String(r[FC.ID]), label: String(r[FC.LABEL] || ''), bank_name: String(r[FC.BANK] || ''),
             amount: Number(r[FC.AMOUNT]) || 0,
             deposit_date: r[FC.DEPOSIT] instanceof Date ? Utilities.formatDate(r[FC.DEPOSIT], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[FC.DEPOSIT] || ''),
             maturity_date: r[FC.MATURITY] instanceof Date ? Utilities.formatDate(r[FC.MATURITY], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[FC.MATURITY] || ''),
             interest_rate: Number(r[FC.RATE]) || 0, status: String(r[FC.STATUS] || 'Active') };
  }

  function _linesForFd(fdId) {
    ensureSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LINE_SHEET);
    if (sh.getLastRow() < 2) return [];
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, LINE_HEADERS.length).getValues();
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][LC.FD]) !== fdId) continue;
      var d = rows[i][LC.DATE];
      out.push({ line_id: String(rows[i][LC.ID]), fd_id: fdId,
                 date: d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(d || ''),
                 description: String(rows[i][LC.DESC] || ''), amount: Number(rows[i][LC.AMOUNT]) || 0 });
    }
    out.sort(function(a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    return out;
  }

  // Every FD with its full statement AND its live, self-computed
  // balance — this is the one call the page needs for everything
  // below the CDOA-wide summary.
  function getAllFDs() {
    UsersService.requireAdmin();
    var fds = _fdRows().map(_fdToObj).filter(function(f) { return f.fd_id; });
    var grandTotal = 0, grandAsOf = '';
    fds.forEach(function(fd) {
      var lines = _linesForFd(fd.fd_id);
      var balance = 0, asOf = '';
      lines.forEach(function(l) {
        balance += (DESC_SIGN[l.description] || 1) * l.amount;
        if (l.date > asOf) asOf = l.date;
      });
      fd.lines = lines;
      fd.balance = Math.round(balance * 100) / 100;
      fd.as_of = asOf;
      grandTotal += fd.balance;
      if (asOf > grandAsOf) grandAsOf = asOf;
    });
    return { fds: fds, grand_total: Math.round(grandTotal * 100) / 100, grand_as_of: grandAsOf };
  }

  function addFD(d) {
    UsersService.requireAdmin();
    if (!d || !d.label) throw new Error('A label is required (e.g. "CDOA Fixed Deposit-03").');
    ensureSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FD_SHEET);
    var now = new Date().toISOString();
    var id = Database.generateId('CFD');
    sh.appendRow([id, String(d.label), String(d.bank_name || ''), Number(d.amount) || 0,
                  String(d.deposit_date || ''), String(d.maturity_date || ''), Number(d.interest_rate) || 0,
                  String(d.status || 'Active'), now, now]);
    return { success: true, fd_id: id };
  }

  function updateFD(fdId, d) {
    UsersService.requireAdmin();
    var found = Database.findByColumn(FD_SHEET, FC.ID, fdId);
    if (!found) throw new Error('Fixed Deposit not found: ' + fdId);
    var row = found.data;
    if (d.label !== undefined) row[FC.LABEL] = String(d.label);
    if (d.bank_name !== undefined) row[FC.BANK] = String(d.bank_name);
    if (d.amount !== undefined) row[FC.AMOUNT] = Number(d.amount) || 0;
    if (d.deposit_date !== undefined) row[FC.DEPOSIT] = String(d.deposit_date);
    if (d.maturity_date !== undefined) row[FC.MATURITY] = String(d.maturity_date);
    if (d.interest_rate !== undefined) row[FC.RATE] = Number(d.interest_rate) || 0;
    if (d.status !== undefined) row[FC.STATUS] = String(d.status);
    row[FC.UPDATED] = new Date().toISOString();
    Database.updateRow(FD_SHEET, found.rowIndex, row);
    return { success: true };
  }

  function deleteFD(fdId) {
    UsersService.requireAdmin();
    var found = Database.findByColumn(FD_SHEET, FC.ID, fdId);
    if (!found) throw new Error('Fixed Deposit not found: ' + fdId);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FD_SHEET).deleteRow(found.rowIndex);
    // Its statement lines describe a deposit that no longer exists — remove them too.
    var lineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LINE_SHEET);
    var rows = Database.getAll(LINE_SHEET);
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][LC.FD]) === fdId) lineSheet.deleteRow(i + 2);
    }
    return { success: true };
  }

  // ── Statement-of-account lines, per FD ──
  function addLine(fdId, date, description, amount) {
    UsersService.requireAdmin();
    if (!fdId) throw new Error('fdId is required.');
    if (DESC_SIGN[description] === undefined) throw new Error('description must be Deposit, Interest, or Withdrawal.');
    var amt = Number(amount);
    if (!(amt > 0)) throw new Error('Amount must be greater than zero.');
    ensureSheets();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LINE_SHEET);
    var id = Database.generateId('CFL');
    sh.appendRow([id, fdId, String(date), String(description), amt, new Date().toISOString()]);
    return { success: true, line_id: id };
  }

  function deleteLine(lineId) {
    UsersService.requireAdmin();
    var found = Database.findByColumn(LINE_SHEET, LC.ID, lineId);
    if (!found) throw new Error('Line not found: ' + lineId);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LINE_SHEET).deleteRow(found.rowIndex);
    return { success: true };
  }

  return {
    ensureSheets: ensureSheets,
    getOwed: getOwed, setOwed: setOwed,
    getReceived: getReceived, addReceivedRow: addReceivedRow, updateReceivedRow: updateReceivedRow, deleteReceivedRow: deleteReceivedRow,
    getAllFDs: getAllFDs, addFD: addFD, updateFD: updateFD, deleteFD: deleteFD,
    addLine: addLine, deleteLine: deleteLine
  };
})();
