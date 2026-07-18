// ============================================================
// UnitsService.gs — Unit Management
// ============================================================

var UnitsService = (function() {
  var SHEET = 'Units';

  var C = {
    UNIT_ID:      0,
    TOWER:        1,
    FLOOR:        2,
    UNIT_NUMBER:  3,
    STATUS:       4,
    CREATED_AT:   5,
    OCCUPANCY:     6,  // 'Occupied' / 'Unoccupied' — manual flag; Unoccupied => WMF & LPG not applicable
    REGISTRATION:  7,  // 'Registered' / 'Not-registered' — MF still applicable; paid lump-sum at registration
    CORPUS_PAID:     8,  // 'Yes' / 'No' — corpus fund paid?
    CORPUS_AMOUNT:   9,  // corpus fund amount (default ₹10,000)
    OWNER_IS_TENANT: 10, // 'Yes' = owner occupies the flat themselves; 'No' = rented to a tenant
    LPG_MODE:        11  // 'Using' (default) / 'Own Cylinder' / 'Not Using'
  };
  var TOTAL_COLS = 12;

  // Make sure the Units sheet has the two newer columns.
  function ensureColumns() {
    var sheet = Database.getSheet(SHEET);
    if (sheet.getMaxColumns() < TOTAL_COLS) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), TOTAL_COLS - sheet.getMaxColumns());
    }
    var head = sheet.getRange(1, 1, 1, TOTAL_COLS).getValues()[0];
    if (String(head[C.OCCUPANCY]) !== 'occupancy' || String(head[C.REGISTRATION]) !== 'registration' ||
        String(head[C.CORPUS_PAID]) !== 'corpus_paid' || String(head[C.CORPUS_AMOUNT]) !== 'corpus_amount' ||
        String(head[C.OWNER_IS_TENANT]) !== 'owner_is_tenant' || String(head[C.LPG_MODE]) !== 'lpg_mode') {
      sheet.getRange(1, C.OCCUPANCY + 1, 1, 6).setValues([['occupancy', 'registration', 'corpus_paid', 'corpus_amount', 'owner_is_tenant', 'lpg_mode']]);
    }
  }

  function seedUnits() {
    var existing = Database.getAll(SHEET);
    if (existing.length > 0) { Logger.log('Units already seeded.'); return; }

    var now  = new Date().toISOString();
    var rows = [];

    // Tower A: 8 floors x 8 units
    for (var fa = 1; fa <= 8; fa++) {
      for (var ua = 1; ua <= 8; ua++) {
        var numA = ua < 10 ? '0' + ua : String(ua);
        rows.push(['A' + fa + numA, 'A', fa, numA, 'Vacant', now]);
      }
    }
    // Tower B: 9 floors x 8 units
    for (var fb = 1; fb <= 9; fb++) {
      for (var ub = 1; ub <= 8; ub++) {
        var numB = ub < 10 ? '0' + ub : String(ub);
        rows.push(['B' + fb + numB, 'B', fb, numB, 'Vacant', now]);
      }
    }

    var sheet = Database.getSheet(SHEET);
    rows.forEach(function(row) { sheet.appendRow(row); });
    Logger.log('Seeded ' + rows.length + ' units.');
  }

  function getAllUnits() {
    var rows = Database.getAll(SHEET);
    return rows.map(function(row) { return _toObj(row); })
               .sort(function(a, b) { return a.unit_id.localeCompare(b.unit_id); });
  }

  function getUnitById(unitId) {
    var result = Database.findByColumn(SHEET, C.UNIT_ID, unitId);
    return result ? _toObj(result.data) : null;
  }

  function getUnitsByTower(tower) {
    return getAllUnits().filter(function(u) { return u.tower === tower; });
  }

  function getUnitsByStatus(status) {
    return getAllUnits().filter(function(u) { return u.status === status; });
  }

  function updateUnitStatus(unitId, status) {
    var result = Database.findByColumn(SHEET, C.UNIT_ID, unitId);
    if (!result) throw new Error('Unit not found: ' + unitId);
    var row = result.data;
    row[C.STATUS] = status;
    Database.updateRow(SHEET, result.rowIndex, row);
    return true;
  }

  // Set the manual occupancy / registration flags on a unit.
  function updateUnitFlags(unitId, data) {
    ensureColumns();
    var result = Database.findByColumn(SHEET, C.UNIT_ID, unitId);
    if (!result) throw new Error('Unit not found: ' + unitId);
    var row = result.data;
    while (row.length < TOTAL_COLS) row.push('');
    if (data.occupancy !== undefined && data.occupancy !== null && data.occupancy !== '') {
      if (data.occupancy !== 'Occupied' && data.occupancy !== 'Unoccupied') throw new Error('occupancy must be Occupied or Unoccupied');
      row[C.OCCUPANCY] = data.occupancy;
    }
    if (data.registration !== undefined && data.registration !== null && data.registration !== '') {
      if (data.registration !== 'Registered' && data.registration !== 'Not-registered') throw new Error('registration must be Registered or Not-registered');
      row[C.REGISTRATION] = data.registration;
    }
    if (data.corpus_paid !== undefined && data.corpus_paid !== null && data.corpus_paid !== '') {
      if (data.corpus_paid !== 'Yes' && data.corpus_paid !== 'No') throw new Error('corpus_paid must be Yes or No');
      row[C.CORPUS_PAID] = data.corpus_paid;
    }
    if (data.corpus_amount !== undefined && data.corpus_amount !== null && data.corpus_amount !== '') {
      var ca = Number(data.corpus_amount);
      if (isNaN(ca) || ca < 0) throw new Error('corpus_amount must be a number');
      row[C.CORPUS_AMOUNT] = ca;
    }
    if (data.owner_is_tenant !== undefined && data.owner_is_tenant !== null) {
      if (data.owner_is_tenant !== 'Yes' && data.owner_is_tenant !== 'No' && data.owner_is_tenant !== '') throw new Error('owner_is_tenant must be Yes, No or blank');
      row[C.OWNER_IS_TENANT] = data.owner_is_tenant;
    }
    if (data.lpg_mode !== undefined && data.lpg_mode !== null && data.lpg_mode !== '') {
      if (['Using','Own Cylinder','Not Using'].indexOf(data.lpg_mode) === -1) throw new Error('lpg_mode must be Using, Own Cylinder or Not Using');
      row[C.LPG_MODE] = data.lpg_mode;
    }
    Database.updateRow(SHEET, result.rowIndex, row);
    return { success: true, occupancy: row[C.OCCUPANCY], registration: row[C.REGISTRATION], corpus_paid: row[C.CORPUS_PAID], corpus_amount: row[C.CORPUS_AMOUNT], lpg_mode: row[C.LPG_MODE] };
  }

  /**
   * markOccupiedBatch — set status 'Occupied' for many units in one
   * read + one write (called by the tenants bulk import).
   */
  function markOccupiedBatch(unitIds) {
    if (!unitIds || !unitIds.length) return { updated: 0 };
    var want = {};
    for (var i = 0; i < unitIds.length; i++) want[String(unitIds[i]).toUpperCase()] = true;

    var rows = Database.getAll(SHEET);
    var touched = 0;
    for (var i = 0; i < rows.length; i++) {
      var uid = String(rows[i][0]).toUpperCase();
      if (want[uid] && String(rows[i][4]) !== 'Occupied') {
        rows[i][4] = 'Occupied';
        touched++;
      }
    }
    if (touched > 0 && rows.length > 0) {
      var sheet = Database.getSheet(SHEET);
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      SpreadsheetApp.flush();
    }
    return { updated: touched };
  }

  function getUnitStats() {
    var all      = getAllUnits();
    var occupied = all.filter(function(u) { return u.status === 'Occupied'; }).length;
    var vacant   = all.length - occupied;
    return {
      total:         all.length,
      occupied:      occupied,
      vacant:        vacant,
      towerA:        all.filter(function(u) { return u.tower === 'A'; }).length,
      towerB:        all.filter(function(u) { return u.tower === 'B'; }).length,
      occupancyRate: all.length > 0 ? Math.round((occupied / all.length) * 100) : 0
    };
  }

  function _toObj(row) {
    function safe(i) {
      var v = row[i];
      if (v === undefined || v === null) return '';
      if (v instanceof Date) return v.toISOString();
      return String(v);
    }
    return {
      unit_id:      safe(C.UNIT_ID),
      tower:        safe(C.TOWER),
      floor:        safe(C.FLOOR),
      unit_number:  safe(C.UNIT_NUMBER),
      status:       safe(C.STATUS),
      created_at:   safe(C.CREATED_AT),
      occupancy:     safe(C.OCCUPANCY)    || 'Occupied',
      registration:  safe(C.REGISTRATION) || 'Registered',
      corpus_paid:     safe(C.CORPUS_PAID)  || 'No',
      corpus_amount:   safe(C.CORPUS_AMOUNT) || '10000',
      owner_is_tenant: safe(C.OWNER_IS_TENANT),  // '' = not set
      lpg_mode: safe(C.LPG_MODE) || 'Using'
    };
  }

  return {
    seedUnits:        seedUnits,
    getAllUnits:       getAllUnits,
    getUnitById:      getUnitById,
    getUnitsByTower:  getUnitsByTower,
    getUnitsByStatus: getUnitsByStatus,
    updateUnitStatus: updateUnitStatus,
    updateUnitFlags:  updateUnitFlags,
    ensureColumns:    ensureColumns,
    markOccupiedBatch: markOccupiedBatch,
    getUnitStats:     getUnitStats
  };
})();
