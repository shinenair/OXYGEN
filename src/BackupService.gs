// ═══════════════════════════════════════════════════════════════
// BackupService — full-database backup (single downloadable JSON
// file), restore from that file, and a "wipe for a new facility"
// reset. All three are Administrator-only and irreversible except
// via a fresh backup.
//
// DYNAMIC BY DESIGN: backup and wipe both walk EVERY sheet in the
// spreadsheet at run time — never a hand-typed list — so they can
// never fall behind as new features add new sheets. (The old fixed
// list had silently missed 15 sheets, including the Corpus Fund, the
// second bank account, the caretaker ledger and occupancy status.)
// ═══════════════════════════════════════════════════════════════
var BackupService = (function() {

  // The only sheets KEPT when wiping for a new facility — the app's
  // configuration and access, so the deployment stays working and no
  // Administrator is locked out. EVERYTHING else is association data and
  // gets cleared, whatever new sheets exist now or in future.
  var WIPE_KEEP = { 'Settings': true, 'Categories': true, 'Users': true };

  function _requireAdmin() {
    var email = Session.getActiveUser().getEmail();
    if (UsersService.getRole(email) !== 'admin') throw new Error('Administrator access required.');
  }

  // Returns a JSON STRING: { version, exported_at, spreadsheet_name,
  //   sheets: { name: { values:[[...]], colFormats:[...], frozenRows:n } } }
  // For every sheet we capture its values PLUS one number-format per column.
  // OXYGEN formats columns uniformly (e.g. the "@" text format that stops
  // '2024-01 from being auto-converted to a date), so re-applying those on
  // restore rebuilds an identical, correctly-typed database.
  function exportBackup() {
    _requireAdmin();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var out = { version: 2, exported_at: new Date().toISOString(), spreadsheet_name: ss.getName(), sheets: {} };
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      var sh = sheets[s];
      var name = sh.getName();
      var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
      if (lastRow < 1 || lastCol < 1) {
        out.sheets[name] = { values: [], colFormats: [], frozenRows: sh.getFrozenRows() };
        continue;
      }
      var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
      // One representative number-format per column, taken from the first data
      // row (or the header if that's all there is) — columns are formatted
      // uniformly, so this captures the "text vs date vs number" intent cheaply.
      var fmtRow = lastRow >= 2 ? 2 : 1;
      var colFormats = sh.getRange(fmtRow, 1, 1, lastCol).getNumberFormats()[0];
      out.sheets[name] = { values: values, colFormats: colFormats, frozenRows: sh.getFrozenRows() };
    }
    return JSON.stringify(out);
  }

  // Rebuilds every sheet in the backup EXACTLY as captured: clears it, applies
  // the per-column formats FIRST (so text columns stay text and values are not
  // silently re-typed on write), then writes the values back verbatim.
  // Understands both the new (v2 { values, colFormats }) and the old (v1 plain
  // array) backup shapes. A sheet not present in the backup is left untouched.
  function restoreBackup(jsonText) {
    _requireAdmin();
    var data;
    try { data = JSON.parse(jsonText); }
    catch (e) { throw new Error('This file is not a valid OXYGEN backup (could not parse JSON).'); }
    if (!data || !data.sheets) throw new Error('This file is not a valid OXYGEN backup.');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var restored = [];
    for (var name in data.sheets) {
      var entry = data.sheets[name];
      var rows, colFormats = null, frozen = 1;
      if (Object.prototype.toString.call(entry) === '[object Array]') {
        rows = entry;                              // v1: values only
      } else if (entry && entry.values) {
        rows = entry.values;                       // v2: { values, colFormats, frozenRows }
        colFormats = entry.colFormats || null;
        if (typeof entry.frozenRows === 'number') frozen = entry.frozenRows;
      } else {
        continue;                                  // empty / unrecognised entry
      }
      if (!rows || !rows.length) continue;

      var sh = ss.getSheetByName(name);
      if (!sh) sh = ss.insertSheet(name);
      sh.clear();                                  // wipe values AND formats for a clean rebuild

      var maxCols = 0;
      for (var i = 0; i < rows.length; i++) maxCols = Math.max(maxCols, rows[i].length);
      if (maxCols < 1) continue;
      // Pad ragged rows so setValues() gets a rectangular array.
      var padded = rows.map(function(r) {
        var row = r.slice();
        while (row.length < maxCols) row.push('');
        return row;
      });

      // Formats FIRST — a text-formatted ('@') column keeps "2024-01" as text
      // instead of letting the subsequent write coerce it into a Date.
      if (colFormats && colFormats.length) {
        for (var c = 0; c < maxCols; c++) {
          var f = colFormats[c] || '';
          if (f) sh.getRange(1, c + 1, padded.length, 1).setNumberFormat(f);
        }
      }
      sh.getRange(1, 1, padded.length, maxCols).setValues(padded);
      sh.getRange(1, 1, 1, maxCols).setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      sh.setFrozenRows(frozen > 0 ? frozen : 1);
      restored.push(name + ' (' + (padded.length - 1) + ' rows)');
    }
    return { success: true, restored: restored, exportedAt: data.exported_at || 'unknown' };
  }

  // Clears the DATA (keeps each sheet's header row) of EVERY sheet except the
  // keep-list — a clean slate for reusing this deployment at a different
  // property. Because it walks every sheet dynamically, nothing association-
  // specific is left behind: corpus, committee, caretaker, both bank accounts,
  // occupancy, planner, bank portions, screenshots, fee schedule — all cleared.
  function wipeForNewFacility() {
    _requireAdmin();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var cleared = [];
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      var sh = sheets[s], name = sh.getName();
      if (WIPE_KEEP[name]) continue;
      var last = sh.getLastRow();
      if (last > 1) { sh.deleteRows(2, last - 1); cleared.push(name); }
    }
    return { success: true, cleared: cleared, kept: Object.keys(WIPE_KEEP) };
  }

  return {
    exportBackup:       exportBackup,
    restoreBackup:      restoreBackup,
    wipeForNewFacility: wipeForNewFacility
  };
})();
