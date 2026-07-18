// ============================================================
// TenantsService.gs — Tenant Management (Full Schema v2)
// 41 columns — matches Tenants.xlsx incl. CREDAI, bank statement
// names, vehicle tracking and Jan–Dec waste payment columns
// ============================================================

var TenantsService = (function() {
  var SHEET = 'Tenants';

  var C = {
    TENANT_ID:        0,
    UNIT_ID:          1,
    NAME:             2,   // Name01-02
    NAME2:            3,   // Name03-04
    EMAIL:            4,
    PHONE:            5,
    PHONE2:           6,
    ADDRESS:          7,
    WORK:             8,
    MOVE_IN_DATE:     9,
    MOVE_OUT_DATE:    10,
    ADULTS:           11,
    KIDS:             12,
    CAUTION_DEPOSIT:  13,  // Yes/No
    CAUTION_AMOUNT:   14,
    ID_PROOF:         15,  // Yes/No
    WHATSAPP_GROUP:   16,  // Yes/No
    WA_PHONE1:        17,
    WA_PHONE2:        18,
    BANK_NAME:        19,  // consolidated TENANT NAME ON BANK STMT aliases
    VEHICLE_4W:       20,  // JSON
    VEHICLE_2W:       21,  // JSON
    PROFILE_PICTURE:  22,
    COMMENT:          23,
    CREDAI:           24,
    NUM_4W:           25,
    NUM_2W:           26,
    PAY_JAN:          27, PAY_FEB: 28, PAY_MAR: 29, PAY_APR: 30,
    PAY_MAY:          31, PAY_JUN: 32, PAY_JUL: 33, PAY_AUG: 34,
    PAY_SEP:          35, PAY_OCT: 36, PAY_NOV: 37, PAY_DEC: 38,
    CREATED_AT:       39,
    UPDATED_AT:       40,
    OWNER_IS_TENANT:  41   // 'Yes' = the resident on record is the owner themselves
  };

  var TENANT_HEADERS = [
    'tenant_id','unit_id','name','name2',
    'email','phone','phone2','address','work',
    'move_in_date','move_out_date','adults','kids',
    'caution_deposit','caution_amount','id_proof','whatsapp_group',
    'wa_phone1','wa_phone2',
    'bank_name','vehicle_4w','vehicle_2w',
    'profile_picture','comment',
    'credai','num_4w','num_2w',
    'pay_jan','pay_feb','pay_mar','pay_apr','pay_may','pay_jun',
    'pay_jul','pay_aug','pay_sep','pay_oct','pay_nov','pay_dec',
    'created_at','updated_at',
    'owner_is_tenant'
  ];

  function addTenant(data) {
    _validate(data);
    var now      = new Date().toISOString();
    var tenantId = String(_nextSerial());
    var r = [];
    for (var i = 0; i < TENANT_HEADERS.length; i++) r.push('');
    _fillRow(r, data, now, tenantId, now);
    if (!r[C.MOVE_IN_DATE]) r[C.MOVE_IN_DATE] = now.split('T')[0];
    Database.insert(SHEET, r);
    UnitsService.updateUnitStatus(data.unit_id, 'Occupied');
    return { success: true, tenant_id: tenantId };
  }

  function getTenantsByUnit(unitId) {
    return Database.findAllByColumn(SHEET, C.UNIT_ID, unitId)
      .map(function(r) { return _toObj(r.data); });
  }

  function getTenantById(tenantId) {
    var r = Database.findByColumn(SHEET, C.TENANT_ID, tenantId);
    return r ? _toObj(r.data) : null;
  }


  // Make sure the sheet physically has all columns (older sheets have 41).
  function ensureColumns() {
    var sheet = Database.getSheet(SHEET);
    if (sheet.getMaxColumns() < TENANT_HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), TENANT_HEADERS.length - sheet.getMaxColumns());
    }
    var head = sheet.getRange(1, 1, 1, TENANT_HEADERS.length).getValues()[0];
    if (String(head[C.OWNER_IS_TENANT]) !== 'owner_is_tenant') {
      sheet.getRange(1, C.OWNER_IS_TENANT + 1).setValue('owner_is_tenant');
    }
  }

  function updateTenant(tenantId, data) {
    ensureColumns();
    var result = Database.findByColumn(SHEET, C.TENANT_ID, tenantId);
    if (!result) throw new Error('Tenant not found: ' + tenantId);
    var now = new Date().toISOString();
    var row = result.data;
    while (row.length < TENANT_HEADERS.length) row.push('');
    var keys = Object.keys(data);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (k === 'tenant_id' || k === 'created_at' || k === 'updated_at') continue;
      if (C.hasOwnProperty(k.toUpperCase()) && data[k] !== undefined) {
        row[C[k.toUpperCase()]] = data[k];
      }
    }
    row[C.UPDATED_AT] = now;
    Database.updateRow(SHEET, result.rowIndex, row);
    return { success: true };
  }

  // Wipes EVERY resident record (does not touch HISTORY_Residents archive).
  function deleteAllTenants() {
    var sheet = Database.getSheet(SHEET);
    var last = sheet.getLastRow();
    var deleted = Math.max(0, last - 1);
    if (last > 1) sheet.deleteRows(2, last - 1);
    return { success: true, deleted: deleted };
  }

  function removeTenant(tenantId) {
    var result = Database.findByColumn(SHEET, C.TENANT_ID, tenantId);
    if (!result) throw new Error('Tenant not found: ' + tenantId);
    var unitId = result.data[C.UNIT_ID];
    Database.deleteRow(SHEET, result.rowIndex);
    if (getTenantsByUnit(unitId).length === 0) {
      UnitsService.updateUnitStatus(unitId, 'Vacant');
    }
    return { success: true };
  }

  var HISTORY_SHEET = 'HISTORY_Residents';

  function _ensureHistorySheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(HISTORY_SHEET);
    if (!sh) {
      sh = ss.insertSheet(HISTORY_SHEET);
      var headers = TENANT_HEADERS.concat(['archived_at', 'archived_by']);
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    return sh;
  }

  // Move a resident's full record into HISTORY_Residents, then remove it from
  // the live Tenants sheet so the unit's resident section resets to blank.
  function archiveTenant(tenantId, archivedBy) {
    var result = Database.findByColumn(SHEET, C.TENANT_ID, tenantId);
    if (!result) throw new Error('Tenant not found: ' + tenantId);
    var unitId = result.data[C.UNIT_ID];
    var sh  = _ensureHistorySheet();
    var row = result.data.slice(0);
    while (row.length < TENANT_HEADERS.length) row.push('');
    row.push(new Date().toISOString());
    row.push(archivedBy || 'Manager');
    sh.appendRow(row);
    Database.deleteRow(SHEET, result.rowIndex);
    if (getTenantsByUnit(unitId).length === 0) {
      UnitsService.updateUnitStatus(unitId, 'Vacant');
    }
    return { success: true, archived_name: result.data[C.NAME] || '' };
  }

  // Archived residents for a unit (newest first), as plain objects.
  function getTenantHistory(unitId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(HISTORY_SHEET);
    if (!sh || sh.getLastRow() < 2) return [];
    var values = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (String(row[C.UNIT_ID]) !== String(unitId)) continue;
      var o = _toObj(row);
      o.archived_at = row[TENANT_HEADERS.length]     ? String(row[TENANT_HEADERS.length])     : '';
      o.archived_by = row[TENANT_HEADERS.length + 1] ? String(row[TENANT_HEADERS.length + 1]) : '';
      out.push(o);
    }
    out.sort(function(a, b) { return String(b.archived_at).localeCompare(String(a.archived_at)); });
    return out;
  }

  function getAllTenants() {
    return Database.getAll(SHEET).map(function(row) { return _toObj(row); });
  }

  // ALL archived residents across every unit, read in a single pass — used
  // when resolving "who lived here in month X" system-wide, so it never
  // has to scan the archive sheet once per unit (136 unnecessary scans).
  function getAllTenantHistory() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(HISTORY_SHEET);
    if (!sh || sh.getLastRow() < 2) return [];
    var values = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var o = _toObj(row);
      o.archived_at = row[TENANT_HEADERS.length]     ? String(row[TENANT_HEADERS.length])     : '';
      o.archived_by = row[TENANT_HEADERS.length + 1] ? String(row[TENANT_HEADERS.length + 1]) : '';
      out.push(o);
    }
    return out;
  }

  function searchTenants(query) {
    if (!query) return getAllTenants();
    var q = query.toLowerCase();
    return getAllTenants().filter(function(t) {
      return (t.name      && t.name.toLowerCase().indexOf(q)      > -1) ||
             (t.name2     && t.name2.toLowerCase().indexOf(q)     > -1) ||
             (t.email     && t.email.toLowerCase().indexOf(q)     > -1) ||
             (t.phone     && String(t.phone).indexOf(q)           > -1) ||
             (t.bank_name && t.bank_name.toLowerCase().indexOf(q) > -1) ||
             (t.unit_id   && t.unit_id.toLowerCase().indexOf(q)   > -1);
    });
  }

  function getDefaulters(month) {
    var targetMonth = month || _currentMonth();
    var allTenants  = getAllTenants();
    var defaulters  = [];
    var types = ['Maintenance', 'Waste Management'];

    allTenants.forEach(function(tenant) {
      var payments  = PaymentsService.getPaymentsByTenantAndMonth(tenant.tenant_id, targetMonth);
      var verified  = payments.filter(function(p) { return p.status === 'Verified'; });
      var paidTypes = verified.map(function(p) { return p.payment_type; });
      var missing   = types.filter(function(t) { return paidTypes.indexOf(t) === -1; });
      if (missing.length > 0) {
        defaulters.push({ tenant: tenant, unit_id: tenant.unit_id, month: targetMonth, missing_types: missing });
      }
    });
    return defaulters;
  }

  /**
   * Bulk import from Tenants.xlsx — fully batched (mirrors owners import).
   * Consolidates TENANT NAME ON BANK STMT011..020 into bank_name and
   * 4-Wheeler-01..05 / 2-Wheeler-01..05 into JSON. Jan–Dec are the
   * Waste Management amounts. Serial '#' becomes tenant_id.
   */
  function bulkImportFromCsv(rows) {
    if (!rows || rows.length < 2) return { imported:0, updated:0, errors:['No data rows found.'] };

    var headers = rows[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var errors = [], imported = 0, updated = 0;

    function col(name) {
      var idx = headers.indexOf(name.toLowerCase());
      return idx >= 0 ? idx : -1;
    }

    // Bank alias columns: stmt011..020
    var bankCols = [];
    for (var b = 11; b <= 20; b++) {
      var bc = col('tenant name on bank stmt0' + b);
      if (bc >= 0) bankCols.push(bc);
    }
    var v4Cols = [], v2Cols = [];
    for (var v = 1; v <= 5; v++) {
      var c4 = col('4-wheeler-0' + v);
      var c2 = col('2-wheeler-0' + v);
      if (c4 >= 0) v4Cols.push(c4);
      if (c2 >= 0) v2Cols.push(c2);
    }

    var unitIds = {};
    var unitRows = Database.getAll('Units');
    for (var i = 0; i < unitRows.length; i++) {
      unitIds[String(unitRows[i][0]).trim().toUpperCase()] = true;
    }

    var TOTAL_COLS = TENANT_HEADERS.length; // 41
    var tenantRows = Database.getAll(SHEET);
    var byUnit     = {};
    for (var i = 0; i < tenantRows.length; i++) {
      while (tenantRows[i].length < TOTAL_COLS) tenantRows[i].push('');
      byUnit[String(tenantRows[i][C.UNIT_ID]).trim().toUpperCase()] = i;
    }

    var now = new Date().toISOString();
    var newRows = [];
    var anyUpdates = false;
    var occupiedUnits = [];

    var flatIdx   = col('flat no.') >= 0 ? col('flat no.') : 2;
    var serialIdx = col('#');
    var serialSeq = 0;

    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row || row.length < 3) continue;

      var unitId = String(row[flatIdx] || '').trim().toUpperCase();
      if (!unitId) { errors.push('Row ' + (i+1) + ': No unit ID found'); continue; }
      if (!unitIds[unitId]) { errors.push('Row ' + (i+1) + ': Unit not found: ' + unitId); continue; }

      serialSeq++;
      var serial = serialIdx >= 0 ? _cv(row, serialIdx) : '';
      if (!serial) serial = String(serialSeq);

      var bankNames = [];
      for (var b = 0; b < bankCols.length; b++) {
        var bn = _cv(row, bankCols[b]);
        if (bn && bankNames.indexOf(bn) < 0) bankNames.push(bn);
      }
      var v4 = [];
      for (var v = 0; v < v4Cols.length; v++) {
        var vo = _vehicleObj(_cv(row, v4Cols[v]));
        if (vo) v4.push(vo);
      }
      var v2 = [];
      for (var v = 0; v < v2Cols.length; v++) {
        var vo2 = _vehicleObj(_cv(row, v2Cols[v]));
        if (vo2) v2.push(vo2);
      }

      var d = {
        unit_id:         unitId,
        name:            _cv(row, col('name01-02')) || _cv(row, col('name01')),
        name2:           _cv(row, col('name03-04')) || _cv(row, col('name02')),
        phone:           _cv(row, col('phone')),
        email:           _cv(row, col('email')),
        work:            _cv(row, col('work')),
        credai:          _cv(row, col('credai')),
        move_in_date:    _isoDate(_cv(row, col('move-in date'))),
        move_out_date:   _isoDate(_cv(row, col('move-out date'))),
        caution_deposit: _cv(row, col('caution deposit (cd) paid?')),
        caution_amount:  _cv(row, col('amount paid as cd?'), true),
        id_proof:        _cv(row, col('id proof submitted?')),
        whatsapp_group:  _cv(row, col('added in whatsapp group?')),
        wa_phone1:       _cv(row, col('phone no.01 in whatsapp group?')),
        wa_phone2:       _cv(row, col('phone no.02 in whatsapp group?')),
        profile_picture: _cv(row, col('profile picture')),
        comment:         _cv(row, col('comment 01')) || _cv(row, col('comment')),
        bank_name:       bankNames.join(', '),
        num_4w:          _cv(row, col('no:of 4-wheelers')),
        num_2w:          _cv(row, col('no:of 2-wheelers')),
        vehicle_4w:      JSON.stringify(v4),
        vehicle_2w:      JSON.stringify(v2),
        pay_jan: _cv(row, col('jan'), true), pay_feb: _cv(row, col('feb'), true),
        pay_mar: _cv(row, col('mar'), true), pay_apr: _cv(row, col('apr'), true),
        pay_may: _cv(row, col('may'), true), pay_jun: _cv(row, col('jun'), true),
        pay_jul: _cv(row, col('jul'), true), pay_aug: _cv(row, col('aug'), true),
        pay_sep: _cv(row, col('sep'), true), pay_oct: _cv(row, col('oct'), true),
        pay_nov: _cv(row, col('nov'), true), pay_dec: _cv(row, col('dec'), true)
      };

      if (!d.name) { errors.push('Row ' + (i+1) + ' (' + unitId + '): No tenant name'); continue; }

      if (byUnit.hasOwnProperty(unitId)) {
        var r = tenantRows[byUnit[unitId]];
        _fillRow(r, d, now, serial, r[C.CREATED_AT]);
        anyUpdates = true;
        updated++;
      } else {
        var nr = [];
        for (var c = 0; c < TOTAL_COLS; c++) nr.push('');
        _fillRow(nr, d, now, serial, now);
        newRows.push(nr);
        byUnit[unitId] = tenantRows.length;
        tenantRows.push(nr);
        imported++;
      }
      occupiedUnits.push(unitId);
    }

    // Write back — batched
    var sheet = Database.getSheet(SHEET);
    if (sheet.getMaxColumns() < TOTAL_COLS) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), TOTAL_COLS - sheet.getMaxColumns());
    }
    var r1c1 = String(sheet.getRange(1, 1).getValue() || '');
    if (r1c1 === '' || r1c1 === 'tenant_id') {
      sheet.getRange(1, 1, 1, TOTAL_COLS).setValues([TENANT_HEADERS]);
      sheet.getRange(1, 1, 1, TOTAL_COLS)
        .setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
    }
    var existingCount = tenantRows.length - newRows.length;
    if (anyUpdates && existingCount > 0) {
      sheet.getRange(2, 1, existingCount, TOTAL_COLS).setValues(tenantRows.slice(0, existingCount));
    }
    if (newRows.length > 0) {
      sheet.getRange(existingCount + 2, 1, newRows.length, TOTAL_COLS).setValues(newRows);
    }
    SpreadsheetApp.flush();

    // Mark all imported units as Occupied — batched
    var unitsMarked = 0;
    try {
      var mo = UnitsService.markOccupiedBatch(occupiedUnits);
      unitsMarked = mo && mo.updated ? mo.updated : 0;
    } catch (e) {}

    return { imported: imported, updated: updated, errors: errors,
             unitsMarked: unitsMarked,
             schema: 'Tenants v2 — 41 columns (bank names + vehicles + CREDAI + waste payments)' };
  }

  /**
   * applyMonthlyPayBatch — write pay_<mon> (waste amounts) for many units.
   * syncMap: { 'A101': { '2026-01': 170 }, ... } — batched, one read/write.
   */
  function applyMonthlyPayBatch(syncMap) {
    var units = syncMap ? Object.keys(syncMap) : [];
    if (!units.length) return { updated: 0 };
    var TOTAL_COLS = TENANT_HEADERS.length;
    var rows   = Database.getAll(SHEET);
    var byUnit = {};
    for (var i = 0; i < rows.length; i++) {
      while (rows[i].length < TOTAL_COLS) rows[i].push('');
      byUnit[String(rows[i][C.UNIT_ID]).trim().toUpperCase()] = i;
    }
    var monthCols = [C.PAY_JAN, C.PAY_FEB, C.PAY_MAR, C.PAY_APR, C.PAY_MAY, C.PAY_JUN,
                     C.PAY_JUL, C.PAY_AUG, C.PAY_SEP, C.PAY_OCT, C.PAY_NOV, C.PAY_DEC];
    var touched = 0;
    for (var u = 0; u < units.length; u++) {
      var unit = String(units[u]).toUpperCase();
      if (!byUnit.hasOwnProperty(unit)) continue;
      var row  = rows[byUnit[unit]];
      var mons = syncMap[units[u]];
      var keys = Object.keys(mons);
      var any  = false;
      for (var k = 0; k < keys.length; k++) {
        var m = keys[k].match(/^(\d{4})-(\d{2})$/);
        if (!m) continue;
        var mi = Number(m[2]) - 1;
        if (mi < 0 || mi > 11) continue;
        row[monthCols[mi]] = mons[keys[k]];
        any = true;
      }
      if (any) { row[C.UPDATED_AT] = new Date().toISOString(); touched++; }
    }
    if (touched > 0 && rows.length > 0) {
      var sheet = Database.getSheet(SHEET);
      if (sheet.getMaxColumns() < TOTAL_COLS) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), TOTAL_COLS - sheet.getMaxColumns());
      }
      sheet.getRange(2, 1, rows.length, TOTAL_COLS).setValues(rows);
      SpreadsheetApp.flush();
    }
    return { updated: touched };
  }

  function _fillRow(r, d, now, tenantId, createdAt) {
    r[C.TENANT_ID]       = tenantId;
    r[C.UNIT_ID]         = d.unit_id;
    r[C.NAME]            = d.name            || '';
    r[C.NAME2]           = d.name2           || '';
    r[C.EMAIL]           = d.email           || '';
    r[C.PHONE]           = d.phone           || '';
    if (d.phone2  !== undefined) r[C.PHONE2]  = d.phone2;
    if (d.address !== undefined) r[C.ADDRESS] = d.address;
    r[C.WORK]            = d.work            || '';
    if (d.move_in_date  !== undefined && d.move_in_date  !== '') r[C.MOVE_IN_DATE]  = d.move_in_date;
    if (d.move_out_date !== undefined) r[C.MOVE_OUT_DATE] = d.move_out_date || '';
    if (d.adults !== undefined) r[C.ADULTS] = d.adults;
    if (d.kids   !== undefined) r[C.KIDS]   = d.kids;
    r[C.CAUTION_DEPOSIT] = d.caution_deposit || r[C.CAUTION_DEPOSIT] || 'No';
    if (d.caution_amount !== undefined) r[C.CAUTION_AMOUNT] = d.caution_amount;
    r[C.ID_PROOF]        = d.id_proof        || r[C.ID_PROOF] || 'No';
    r[C.WHATSAPP_GROUP]  = d.whatsapp_group  || r[C.WHATSAPP_GROUP] || 'No';
    if (d.wa_phone1 !== undefined) r[C.WA_PHONE1] = d.wa_phone1;
    if (d.wa_phone2 !== undefined) r[C.WA_PHONE2] = d.wa_phone2;
    if (d.bank_name !== undefined) r[C.BANK_NAME] = d.bank_name;
    if (d.vehicle_4w !== undefined) r[C.VEHICLE_4W] = d.vehicle_4w;
    if (!r[C.VEHICLE_4W]) r[C.VEHICLE_4W] = '[]';
    if (d.vehicle_2w !== undefined) r[C.VEHICLE_2W] = d.vehicle_2w;
    if (!r[C.VEHICLE_2W]) r[C.VEHICLE_2W] = '[]';
    if (d.profile_picture !== undefined) r[C.PROFILE_PICTURE] = d.profile_picture;
    if (d.comment !== undefined) r[C.COMMENT] = d.comment;
    if (d.credai  !== undefined) r[C.CREDAI]  = d.credai;
    if (d.num_4w  !== undefined) r[C.NUM_4W]  = d.num_4w;
    if (d.num_2w  !== undefined) r[C.NUM_2W]  = d.num_2w;
    r[C.PAY_JAN] = d.pay_jan || ''; r[C.PAY_FEB] = d.pay_feb || '';
    r[C.PAY_MAR] = d.pay_mar || ''; r[C.PAY_APR] = d.pay_apr || '';
    r[C.PAY_MAY] = d.pay_may || ''; r[C.PAY_JUN] = d.pay_jun || '';
    r[C.PAY_JUL] = d.pay_jul || ''; r[C.PAY_AUG] = d.pay_aug || '';
    r[C.PAY_SEP] = d.pay_sep || ''; r[C.PAY_OCT] = d.pay_oct || '';
    r[C.PAY_NOV] = d.pay_nov || ''; r[C.PAY_DEC] = d.pay_dec || '';
    r[C.CREATED_AT] = createdAt || now;
    r[C.UPDATED_AT] = now;
  }

  function _vehicleObj(s) {
    if (!s) return null;
    var p = String(s).split('/');
    for (var i = 0; i < p.length; i++) p[i] = p[i].trim();
    return {
      reg:    p[0] || '',
      make:   p.length > 1 ? p[1] : '',
      model:  p.length > 2 ? p[2] : '',
      colour: p.length > 3 ? p.slice(3).join('/') : ''
    };
  }

  // 'DD/MM/YYYY' or 'DD-MM-YYYY' -> 'YYYY-MM-DD'; ISO passes through
  function _isoDate(s) {
    if (!s) return '';
    var m = String(s).match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) {
      return m[3] + '-' + (m[2].length < 2 ? '0' + m[2] : m[2]) + '-' + (m[1].length < 2 ? '0' + m[1] : m[1]);
    }
    return String(s);
  }

  function _nextSerial() {
    var rows = Database.getAll(SHEET);
    var max  = 0;
    for (var i = 0; i < rows.length; i++) {
      var n = parseInt(rows[i][C.TENANT_ID], 10);
      if (!isNaN(n) && n > max) max = n;
    }
    return max + 1;
  }

  function _cv(row, colIdx, isMoney) {
    if (colIdx < 0 || colIdx >= row.length) return '';
    var v = String(row[colIdx] === null || row[colIdx] === undefined ? '' : row[colIdx]).trim();
    if (isMoney) v = v.replace(/[₹,\s]/g, '').trim();
    return v;
  }

  function _currentMonth() {
    var now = new Date();
    var m = now.getMonth() + 1;
    return now.getFullYear() + '-' + (m < 10 ? '0' + m : String(m));
  }

  function _validate(data) {
    if (!data.unit_id) throw new Error('unit_id is required.');
    if (!data.name)    throw new Error('Tenant name is required.');
    if (!UnitsService.getUnitById(data.unit_id)) throw new Error('Unit not found: ' + data.unit_id);
  }


  // Format a Date cell using the spreadsheet's own timezone — never UTC —
  // so a date typed as 10 Feb 2026 never shifts to 09 Feb on read.
  var _CELL_TZ = null;
  function _fmtCellDate(v) {
    if (!_CELL_TZ) _CELL_TZ = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var s = Utilities.formatDate(v, _CELL_TZ, 'yyyy-MM-dd HH:mm:ss');
    return s.slice(11) === '00:00:00' ? s.slice(0, 10) : s;
  }

  function _toObj(row) {
    function s(i) {
      var v = row[i];
      if (v === undefined || v === null) return '';
      if (v instanceof Date) return _fmtCellDate(v);
      return String(v);
    }
    return {
      tenant_id:       s(C.TENANT_ID),
      unit_id:         s(C.UNIT_ID),
      name:            s(C.NAME),
      name2:           s(C.NAME2),
      email:           s(C.EMAIL),
      phone:           s(C.PHONE),
      phone2:          s(C.PHONE2),
      address:         s(C.ADDRESS),
      work:            s(C.WORK),
      move_in_date:    s(C.MOVE_IN_DATE),
      move_out_date:   s(C.MOVE_OUT_DATE),
      adults:          s(C.ADULTS),
      kids:            s(C.KIDS),
      caution_deposit: s(C.CAUTION_DEPOSIT),
      caution_amount:  s(C.CAUTION_AMOUNT),
      id_proof:        s(C.ID_PROOF),
      whatsapp_group:  s(C.WHATSAPP_GROUP),
      wa_phone1:       s(C.WA_PHONE1),
      wa_phone2:       s(C.WA_PHONE2),
      bank_name:       s(C.BANK_NAME),
      vehicle_4w:      s(C.VEHICLE_4W) || '[]',
      vehicle_2w:      s(C.VEHICLE_2W) || '[]',
      profile_picture: s(C.PROFILE_PICTURE),
      comment:         s(C.COMMENT),
      credai:          s(C.CREDAI),
      num_4w:          s(C.NUM_4W),
      num_2w:          s(C.NUM_2W),
      pay_jan: s(C.PAY_JAN), pay_feb: s(C.PAY_FEB),
      pay_mar: s(C.PAY_MAR), pay_apr: s(C.PAY_APR),
      pay_may: s(C.PAY_MAY), pay_jun: s(C.PAY_JUN),
      pay_jul: s(C.PAY_JUL), pay_aug: s(C.PAY_AUG),
      pay_sep: s(C.PAY_SEP), pay_oct: s(C.PAY_OCT),
      pay_nov: s(C.PAY_NOV), pay_dec: s(C.PAY_DEC),
      created_at:      s(C.CREATED_AT),
      updated_at:      s(C.UPDATED_AT),
      owner_is_tenant: s(C.OWNER_IS_TENANT) || 'No'
    };
  }

  return {
    addTenant:            addTenant,
    getTenantsByUnit:     getTenantsByUnit,
    getTenantById:        getTenantById,
    updateTenant:         updateTenant,
    removeTenant:         removeTenant,
    deleteAllTenants:     deleteAllTenants,
    ensureColumns:        ensureColumns,
    archiveTenant:        archiveTenant,
    getTenantHistory:     getTenantHistory,
    getAllTenants:         getAllTenants,
    getAllTenantHistory:   getAllTenantHistory,
    searchTenants:        searchTenants,
    getDefaulters:        getDefaulters,
    bulkImportFromCsv:    bulkImportFromCsv,
    applyMonthlyPayBatch: applyMonthlyPayBatch
  };
})();
