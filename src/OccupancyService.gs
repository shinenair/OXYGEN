// ═══════════════════════════════════════════════════════════════
// OccupancyService.gs — Monthly Occupancy & LPG Mode Timeline.
//
// THE PROBLEM THIS SOLVES: a unit's Occupied/Unoccupied flag (and its
// LPG mode) used to be a single, current-state value — applied as a
// blanket assumption to every month being viewed, with no memory of
// when it actually changed. That meant Jan-Unoccupied /
// Feb-Occupied / Mar-Unoccupied-again couldn't be represented at all
// without manually re-marking individual months, and nothing showed
// the history.
//
// THE MODEL: a sparse list of "as of this month, the value became X"
// entries per unit per attribute (occupancy, lpg_mode). A change
// carries forward automatically until the next explicit entry — so
// marking a unit Unoccupied starting January means every month after
// January is ALSO Unoccupied by default, exactly as the association
// asked for, and any single month can be edited individually to
// override that (e.g. Feb back to Occupied), which itself then
// carries forward until the next change. Units with no entries yet
// fall back to their existing static occupancy/lpg_mode fields, so
// nothing already relied upon changes until the timeline is used.
//
// RULE: presence for even one day in a month counts that whole month
// as Occupied — this engine only ever stores whole-month values, by
// design; there is no partial-month state to represent.
// ═══════════════════════════════════════════════════════════════
var OccupancyService = (function() {
  var SHEET = 'UnitMonthlyStatus';
  var HEADERS = ['entry_id', 'unit_id', 'attribute', 'effective_month', 'value', 'created_at', 'updated_at'];
  var C = { ID: 0, UNIT: 1, ATTR: 2, MONTH: 3, VALUE: 4, CREATED: 5, UPDATED: 6 };

  var OCC_VALUES = ['Occupied', 'Unoccupied'];
  var LPG_VALUES = ['Using', 'Own Cylinder', 'Not Using'];

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

  function _allEntries() {
    ensureSheet();
    return Database.getAll(SHEET).map(function(r) {
      return { entry_id: String(r[C.ID] || ''), unit_id: String(r[C.UNIT] || '').toUpperCase(),
               attribute: String(r[C.ATTR] || ''), effective_month: String(r[C.MONTH] || '').replace(/^'/, ''),
               value: String(r[C.VALUE] || '') };
    }).filter(function(e) { return e.entry_id; });
  }

  // Every timeline entry for one unit+attribute, sorted chronologically.
  function getTimeline(unitId, attribute) {
    var unit = String(unitId).toUpperCase();
    return _allEntries()
      .filter(function(e) { return e.unit_id === unit && e.attribute === attribute; })
      .sort(function(a, b) { return a.effective_month < b.effective_month ? -1 : a.effective_month > b.effective_month ? 1 : 0; });
  }

  // The value that was in effect for a given month: the most recent
  // entry at or before that month, or the fallback (the unit's
  // existing static field) if no entry exists yet at or before it.
  function _effectiveFor(timeline, monthKey, fallback) {
    var value = fallback;
    for (var i = 0; i < timeline.length; i++) {
      if (timeline[i].effective_month <= monthKey) value = timeline[i].value;
      else break; // sorted — nothing further back matters once we've passed monthKey
    }
    return value;
  }

  // All 12 months of a year for one unit, both attributes at once —
  // this is what the Fees grid and Follow-up Board actually need, one
  // call per unit rather than 24 individual lookups.
  function getYearView(unitId, year) {
    var unit = String(unitId).toUpperCase();
    var unitRec = UnitsService.getUnitById(unit) || {};
    var occTimeline = getTimeline(unit, 'occupancy');
    var lpgTimeline = getTimeline(unit, 'lpg_mode');
    var occFallback = unitRec.occupancy || 'Occupied';
    var lpgFallback = unitRec.lpg_mode || 'Using';

    var months = [];
    for (var m = 1; m <= 12; m++) {
      var key = year + '-' + (m < 10 ? '0' + m : m);
      months.push({
        month: key,
        occupancy: _effectiveFor(occTimeline, key, occFallback),
        lpg_mode: _effectiveFor(lpgTimeline, key, lpgFallback),
        occupancy_explicit: occTimeline.some(function(e) { return e.effective_month === key; }),
        lpg_mode_explicit: lpgTimeline.some(function(e) { return e.effective_month === key; })
      });
    }
    return { unit_id: unit, months: months };
  }

  // Batched: every unit's full-year view in ONE call — what the Fees
  // Received grid and Follow-up Board use, so rendering 136 units never
  // costs 136 round-trips.
  function getYearViewForAllUnits(year) {
    var units = UnitsService.getAllUnits();
    var allEntries = _allEntries();
    var byUnit = {};
    allEntries.forEach(function(e) {
      if (!byUnit[e.unit_id]) byUnit[e.unit_id] = { occupancy: [], lpg_mode: [] };
      if (byUnit[e.unit_id][e.attribute]) byUnit[e.unit_id][e.attribute].push(e);
    });
    for (var k in byUnit) {
      byUnit[k].occupancy.sort(function(a, b) { return a.effective_month < b.effective_month ? -1 : 1; });
      byUnit[k].lpg_mode.sort(function(a, b) { return a.effective_month < b.effective_month ? -1 : 1; });
    }

    var out = {};
    units.forEach(function(u) {
      var unit = String(u.unit_id).toUpperCase();
      var entries = byUnit[unit] || { occupancy: [], lpg_mode: [] };
      var occFallback = u.occupancy || 'Occupied';
      var lpgFallback = u.lpg_mode || 'Using';
      var months = [];
      for (var m = 1; m <= 12; m++) {
        var key = year + '-' + (m < 10 ? '0' + m : m);
        months.push({
          occupancy: _effectiveFor(entries.occupancy, key, occFallback),
          lpg_mode: _effectiveFor(entries.lpg_mode, key, lpgFallback)
        });
      }
      out[unit] = months; // index 0 = Jan .. 11 = Dec
    });
    return out;
  }

  // Sets (or replaces) the value effective from a given month onward —
  // this is the one write operation. A month can only have ONE entry
  // per attribute; re-setting the same month overwrites it rather than
  // creating a duplicate.
  function setStatus(unitId, attribute, effectiveMonth, value) {
    if (attribute !== 'occupancy' && attribute !== 'lpg_mode') throw new Error('Unknown attribute: ' + attribute);
    var validValues = attribute === 'occupancy' ? OCC_VALUES : LPG_VALUES;
    if (validValues.indexOf(value) === -1) throw new Error(attribute + ' must be one of: ' + validValues.join(', '));
    if (!/^\d{4}-\d{2}$/.test(effectiveMonth)) throw new Error('effectiveMonth must be yyyy-MM.');

    var unit = String(unitId).toUpperCase();
    var sh = ensureSheet();
    var rows = Database.getAll(SHEET);
    var now = new Date().toISOString();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.UNIT]).toUpperCase() === unit && String(rows[i][C.ATTR]) === attribute &&
          String(rows[i][C.MONTH]).replace(/^'/, '') === effectiveMonth) {
        rows[i][C.VALUE] = value;
        rows[i][C.UPDATED] = now;
        Database.updateRow(SHEET, i + 2, rows[i]);
        return { success: true, entry_id: String(rows[i][C.ID]), created: false };
      }
    }
    var id = Database.generateId('UMS');
    sh.appendRow([id, unit, attribute, "'" + effectiveMonth, value, now, now]);
    return { success: true, entry_id: id, created: true };
  }

  // Removes one explicit entry — the month then simply inherits
  // whatever was in effect immediately before it again.
  function deleteEntry(entryId) {
    var found = Database.findByColumn(SHEET, C.ID, entryId);
    if (!found) throw new Error('Timeline entry not found: ' + entryId);
    ensureSheet().deleteRow(found.rowIndex);
    return { success: true };
  }

  // Every raw entry ever set for this unit, both attributes, ALL years —
  // the ground truth behind the computed carry-forward view. Exists so
  // unexpected behavior (a value not carrying forward as expected) can
  // be diagnosed by SEEING exactly what's stored, rather than guessing.
  function getAllEntriesForUnit(unitId) {
    var unit = String(unitId).toUpperCase();
    return _allEntries()
      .filter(function(e) { return e.unit_id === unit; })
      .sort(function(a, b) {
        if (a.effective_month !== b.effective_month) return a.effective_month < b.effective_month ? -1 : 1;
        return a.attribute < b.attribute ? -1 : 1;
      });
  }

  return {
    ensureSheet: ensureSheet,
    getTimeline: getTimeline,
    getYearView: getYearView,
    getYearViewForAllUnits: getYearViewForAllUnits,
    setStatus: setStatus,
    deleteEntry: deleteEntry,
    getAllEntriesForUnit: getAllEntriesForUnit
  };
})();
