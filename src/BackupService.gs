// ═══════════════════════════════════════════════════════════════
// BackupService — full-database backup (single downloadable JSON
// file), restore from that file, and a "wipe for a new facility"
// reset. All three are Administrator-only and irreversible except
// via a fresh backup.
// ═══════════════════════════════════════════════════════════════
var BackupService = (function() {

  // Every sheet OXYGEN manages. A backup captures ALL of them exactly as
  // they are (including headers), so restoring rebuilds the database to
  // an identical, fully working state from this one file.
  var ALL_SHEETS = [
    'Units', 'Owners', 'HISTORY_Owners', 'Tenants', 'HISTORY_Residents',
    'Payments', 'BankStatements', 'LPGReadings', 'LPGRates',
    'LPGStockInward', 'LPGStockOutward', 'Settings', 'Categories', 'Users'
  ];

  // Sheets cleared by "wipe for a new facility" — the association's own
  // data. Settings, Categories and Users are deliberately LEFT ALONE, so
  // the app stays configured and nobody gets locked out; a brand-new
  // facility's units/owners/tenants/financials start from zero instead.
  var FACILITY_DATA_SHEETS = [
    'Units', 'Owners', 'HISTORY_Owners', 'Tenants', 'HISTORY_Residents',
    'Payments', 'BankStatements', 'LPGReadings', 'LPGRates',
    'LPGStockInward', 'LPGStockOutward'
  ];

  function _requireAdmin() {
    var email = Session.getActiveUser().getEmail();
    if (UsersService.getRole(email) !== 'admin') throw new Error('Administrator access required.');
  }

  // Returns { version, exported_at, sheets: { name: [[row],[row],...] } }
  // as a JSON STRING — the client turns this into a downloadable file.
  function exportBackup() {
    _requireAdmin();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var out = { version: 1, exported_at: new Date().toISOString(), spreadsheet_name: ss.getName(), sheets: {} };
    ALL_SHEETS.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      out.sheets[name] = sh ? sh.getDataRange().getValues() : [];
    });
    return JSON.stringify(out);
  }

  // Rebuilds every sheet in the backup EXACTLY as captured — clears each
  // sheet first, then writes the backup's rows back verbatim. Any sheet
  // not present in the backup (e.g. an older backup taken before a new
  // feature existed) is left untouched rather than guessed at.
  function restoreBackup(jsonText) {
    _requireAdmin();
    var data;
    try { data = JSON.parse(jsonText); }
    catch (e) { throw new Error('This file is not a valid OXYGEN backup (could not parse JSON).'); }
    if (!data || !data.sheets) throw new Error('This file is not a valid OXYGEN backup.');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var restored = [];
    for (var name in data.sheets) {
      var rows = data.sheets[name];
      if (!rows || !rows.length) continue;
      var sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      sh.clearContents();
      var maxCols = 0;
      for (var i = 0; i < rows.length; i++) maxCols = Math.max(maxCols, rows[i].length);
      // Pad ragged rows so setValues() gets a rectangular array.
      var padded = rows.map(function(r) {
        var row = r.slice();
        while (row.length < maxCols) row.push('');
        return row;
      });
      sh.getRange(1, 1, padded.length, maxCols).setValues(padded);
      if (padded.length) {
        sh.getRange(1, 1, 1, maxCols).setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
        sh.setFrozenRows(1);
      }
      restored.push(name + ' (' + (padded.length - 1) + ' rows)');
    }
    return { success: true, restored: restored, exportedAt: data.exported_at || 'unknown' };
  }

  // Clears every FACILITY-DATA sheet down to just its header row — a
  // clean slate for reusing this same OXYGEN deployment at a different
  // property. Settings/Categories/Users are intentionally left in place.
  function wipeForNewFacility() {
    _requireAdmin();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cleared = [];
    FACILITY_DATA_SHEETS.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) return;
      var last = sh.getLastRow();
      if (last > 1) { sh.deleteRows(2, last - 1); cleared.push(name); }
    });
    return { success: true, cleared: cleared };
  }

  return {
    exportBackup:       exportBackup,
    restoreBackup:      restoreBackup,
    wipeForNewFacility: wipeForNewFacility
  };
})();
