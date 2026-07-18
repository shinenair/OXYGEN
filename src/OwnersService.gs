// ============================================================
// OwnersService.gs — Owner Management (Full Schema v3)
// 45 columns — matches updated Owners.xlsx incl. corpus fund,
// bank statement names, and vehicle tracking
// ============================================================

var OwnersService = (function() {
  var SHEET = 'Owners';

  // Column indices — Database.gs headers MUST match this order
  var C = {
    OWNER_ID:            0,
    UNIT_ID:             1,
    NAME:                2,   // Name01
    EMAIL:               3,   // Email ID 01
    PHONE:               4,   // Phone No:01
    ADDRESS:             5,   // legacy
    NAME2:               6,
    NAME3:               7,
    PHONE2:              8,
    PHONE3:              9,
    EMAIL2:              10,
    EXEC_MEMBER:         11,
    EXEC_MEMBER_TITLE:   12,
    CAR_PARKING_SLOT:    13,
    LIVING_STATUS:       14,
    RESIDING_COUNTRY:    15,
    ADDRESS_INDIA:       16,
    ADDRESS_ABROAD:      17,
    EMPLOYER:            18,
    WORK_DESIGNATION:    19,
    WORK_ADDRESS:        20,
    TC_NUMBER:           21,
    KSEB_CONSUMER_NO:    22,
    NOTES:               23,
    PROFILE_PICTURE:     24,
    VEHICLES_4W:         25,  // JSON array [{reg,make,model,colour}]
    VEHICLES_2W:         26,  // JSON array
    BANK_NAMES:          27,  // comma-separated (from NAME ON BANK STMT01..10)
    CORPUS_FUND_PENDING: 28,  // CORPUS FUND PENDING
    NUM_4W:              29,  // No:of 4-Wheelers
    NUM_2W:              30,  // No:of 2-Wheelers
    PAY_JAN:             31,
    PAY_FEB:             32,
    PAY_MAR:             33,
    PAY_APR:             34,
    PAY_MAY:             35,
    PAY_JUN:             36,
    PAY_JUL:             37,
    PAY_AUG:             38,
    PAY_SEP:             39,
    PAY_OCT:             40,
    PAY_NOV:             41,
    PAY_DEC:             42,
    CREATED_AT:          43,
    UPDATED_AT:          44
  };

  // Canonical header row — order matches C indices exactly
  var OWNER_HEADERS = [
    'owner_id','unit_id',
    'name','email','phone','address',
    'name2','name3','phone2','phone3','email2',
    'exec_member','exec_member_title',
    'car_parking_slot','living_status','residing_country',
    'address_india','address_abroad',
    'employer','work_designation','work_address',
    'tc_number','kseb_consumer_no',
    'notes','profile_picture',
    'vehicles_4w','vehicles_2w','bank_names',
    'corpus_fund_pending','num_4w','num_2w',
    'pay_jan','pay_feb','pay_mar','pay_apr','pay_may','pay_jun',
    'pay_jul','pay_aug','pay_sep','pay_oct','pay_nov','pay_dec',
    'created_at','updated_at'
  ];

  function addOwner(data) {
    _validate(data);
    var now     = new Date().toISOString();
    var ownerId = String(_nextSerial());
    var r = [];
    for (var i = 0; i < OWNER_HEADERS.length; i++) r.push('');
    _fillRow(r, data, now, ownerId, now);
    Database.insert(SHEET, r);
    return { success: true, owner_id: ownerId };
  }

  function getOwnersByUnit(unitId) {
    return Database.findAllByColumn(SHEET, C.UNIT_ID, unitId)
      .map(function(r) { return _toObj(r.data); });
  }

  function getOwnerById(ownerId) {
    var r = Database.findByColumn(SHEET, C.OWNER_ID, ownerId);
    return r ? _toObj(r.data) : null;
  }

  function updateOwner(ownerId, data) {
    var result = Database.findByColumn(SHEET, C.OWNER_ID, ownerId);
    if (!result) throw new Error('Owner not found: ' + ownerId);
    var now = new Date().toISOString();
    var row = result.data;
    while (row.length < OWNER_HEADERS.length) row.push('');
    var keys = Object.keys(data);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (k === 'owner_id' || k === 'created_at' || k === 'updated_at') continue;
      if (C.hasOwnProperty(k.toUpperCase()) && data[k] !== undefined) {
        row[C[k.toUpperCase()]] = data[k];
      }
    }
    row[C.UPDATED_AT] = now;
    Database.updateRow(SHEET, result.rowIndex, row);
    return { success: true };
  }

  // Wipes EVERY owner record (does not touch HISTORY_Owners archive).
  // Used for a clean start on a brand-new facility, or a full data reset.
  function deleteAllOwners() {
    var sheet = Database.getSheet(SHEET);
    var last = sheet.getLastRow();
    var deleted = Math.max(0, last - 1);
    if (last > 1) sheet.deleteRows(2, last - 1);
    return { success: true, deleted: deleted };
  }

  function deleteOwner(ownerId) {
    var r = Database.findByColumn(SHEET, C.OWNER_ID, ownerId);
    if (!r) throw new Error('Owner not found: ' + ownerId);
    Database.deleteRow(SHEET, r.rowIndex);
    return { success: true };
  }

  var HISTORY_SHEET = 'HISTORY_Owners';

  // Create the history sheet (owner columns + archive metadata) if missing.
  function _ensureHistorySheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(HISTORY_SHEET);
    if (!sh) {
      sh = ss.insertSheet(HISTORY_SHEET);
      var headers = OWNER_HEADERS.concat(['archived_at', 'archived_by']);
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    return sh;
  }

  // Move an owner's full record into HISTORY_Owners, then remove it from the
  // live Owners sheet so the unit's owner section resets to blank.
  function archiveOwner(ownerId, archivedBy) {
    var r = Database.findByColumn(SHEET, C.OWNER_ID, ownerId);
    if (!r) throw new Error('Owner not found: ' + ownerId);
    var sh  = _ensureHistorySheet();
    var row = r.data.slice(0);
    while (row.length < OWNER_HEADERS.length) row.push('');
    row.push(new Date().toISOString());       // archived_at
    row.push(archivedBy || 'Manager');        // archived_by
    sh.appendRow(row);
    Database.deleteRow(SHEET, r.rowIndex);
    return { success: true, archived_name: r.data[C.NAME] || '' };
  }

  // Archived owners for a unit (newest first), as plain objects.
  function getOwnerHistory(unitId) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(HISTORY_SHEET);
    if (!sh || sh.getLastRow() < 2) return [];
    var values = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (String(row[C.UNIT_ID]) !== String(unitId)) continue;
      var o = _toObj(row);
      o.archived_at = row[OWNER_HEADERS.length]     ? String(row[OWNER_HEADERS.length])     : '';
      o.archived_by = row[OWNER_HEADERS.length + 1] ? String(row[OWNER_HEADERS.length + 1]) : '';
      out.push(o);
    }
    out.sort(function(a, b) { return String(b.archived_at).localeCompare(String(a.archived_at)); });
    return out;
  }

  function getAllOwners() {
    return Database.getAll(SHEET).map(function(row) { return _toObj(row); });
  }

  function searchOwners(query) {
    if (!query) return getAllOwners();
    var q = query.toLowerCase();
    return getAllOwners().filter(function(o) {
      return (o.name       && o.name.toLowerCase().indexOf(q)       > -1) ||
             (o.name2      && o.name2.toLowerCase().indexOf(q)      > -1) ||
             (o.email      && o.email.toLowerCase().indexOf(q)      > -1) ||
             (o.phone      && String(o.phone).indexOf(q)            > -1) ||
             (o.tc_number  && String(o.tc_number).indexOf(q)        > -1) ||
             (o.bank_names && o.bank_names.toLowerCase().indexOf(q) > -1) ||
             (o.unit_id    && o.unit_id.toLowerCase().indexOf(q)    > -1);
    });
  }

  /**
   * Bulk import — fully batched (reads once, writes in <=3 operations).
   * Consolidates NAME ON BANK STMT01..10 into bank_names and
   * 4-Wheeler-01..05 / 2-Wheeler-01..05 into JSON vehicle lists.
   */
  function bulkImportFromCsv(rows) {
    if (!rows || rows.length < 2) return { imported:0, updated:0, errors:['No data rows found.'] };

    var headers = rows[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var errors  = [];
    var imported = 0;
    var updated  = 0;

    function col(name) {
      var idx = headers.indexOf(name.toLowerCase());
      return idx >= 0 ? idx : -1;
    }

    // Pre-resolve multi-column groups
    var bankCols = [];
    for (var b = 1; b <= 10; b++) {
      var bc = col('name on bank stmt' + (b < 10 ? '0' + b : b));
      if (bc >= 0) bankCols.push(bc);
    }
    var v4Cols = [];
    var v2Cols = [];
    for (var v = 1; v <= 5; v++) {
      var c4 = col('4-wheeler-0' + v);
      var c2 = col('2-wheeler-0' + v);
      if (c4 >= 0) v4Cols.push(c4);
      if (c2 >= 0) v2Cols.push(c2);
    }

    // 1. All valid unit IDs in one read
    var unitIds = {};
    var unitRows = Database.getAll('Units');
    for (var i = 0; i < unitRows.length; i++) {
      unitIds[String(unitRows[i][0]).trim().toUpperCase()] = true;
    }

    // 2. All existing owner rows in one read
    var TOTAL_COLS = OWNER_HEADERS.length; // 45
    var ownerRows  = Database.getAll(SHEET);
    var byUnit     = {};
    for (var i = 0; i < ownerRows.length; i++) {
      while (ownerRows[i].length < TOTAL_COLS) ownerRows[i].push('');
      byUnit[String(ownerRows[i][C.UNIT_ID]).trim().toUpperCase()] = i;
    }

    var now        = new Date().toISOString();
    var newRows    = [];
    var anyUpdates = false;

    var flatIdx   = col('flat no.') >= 0 ? col('flat no.') : col('flat no') >= 0 ? col('flat no') : 2;
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

      // Consolidate NAME ON BANK STMT01..10 -> comma-separated list
      var bankNames = [];
      for (var b = 0; b < bankCols.length; b++) {
        var bn = _cv(row, bankCols[b]);
        if (bn && bankNames.indexOf(bn) < 0) bankNames.push(bn);
      }

      // Vehicles: "REG/MAKE/MODEL/COLOUR" -> {reg,make,model,colour}
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
        unit_id:           unitId,
        name:              _cv(row, col('name01')),
        name2:             _cv(row, col('name02')),
        name3:             _cv(row, col('name03')),
        phone:             _cv(row, col('phone no:01')) || _cv(row, col('phone no:1')),
        phone2:            _cv(row, col('phone no:02')) || _cv(row, col('phone no:2')),
        phone3:            _cv(row, col('phone no:03')) || _cv(row, col('phone no:3')),
        email:             _cv(row, col('email id 01')) || _cv(row, col('email id 1')) || _cv(row, col('email')),
        email2:            _cv(row, col('email id 02')) || _cv(row, col('email id 2')),
        exec_member:       _cv(row, col('exec member')),
        exec_member_title: _cv(row, col('exec member title')),
        car_parking_slot:  _cv(row, col('car parking')),
        living_status:     _cv(row, col('living status')),
        residing_country:  _cv(row, col('residing country')),
        address_india:     _cv(row, col('address (india)')),
        address_abroad:    _cv(row, col('address (abroad)')),
        employer:          _cv(row, col('employer')),
        work_designation:  _cv(row, col('work designation')),
        work_address:      _cv(row, col('work address')),
        tc_number:         _cv(row, col('tc no:')) || _cv(row, col('tc no')),
        kseb_consumer_no:  _cv(row, col('kseb consumer no:')) || _cv(row, col('kseb consumer no')),
        notes:             _cv(row, col('notes')),
        profile_picture:   _cv(row, col('profile picture')),
        bank_names:          bankNames.join(', '),
        corpus_fund_pending: _cv(row, col('corpus fund pending'), true),
        num_4w:              _cv(row, col('no:of 4-wheelers')),
        num_2w:              _cv(row, col('no:of 2-wheelers')),
        vehicles_4w:         JSON.stringify(v4),
        vehicles_2w:         JSON.stringify(v2),
        pay_jan: _cv(row, col('jan'), true), pay_feb: _cv(row, col('feb'), true),
        pay_mar: _cv(row, col('mar'), true), pay_apr: _cv(row, col('apr'), true),
        pay_may: _cv(row, col('may'), true), pay_jun: _cv(row, col('jun'), true),
        pay_jul: _cv(row, col('jul'), true), pay_aug: _cv(row, col('aug'), true),
        pay_sep: _cv(row, col('sep'), true), pay_oct: _cv(row, col('oct'), true),
        pay_nov: _cv(row, col('nov'), true), pay_dec: _cv(row, col('dec'), true)
      };

      if (!d.name) { errors.push('Row ' + (i+1) + ' (' + unitId + '): No owner name'); continue; }

      if (byUnit.hasOwnProperty(unitId)) {
        var r = ownerRows[byUnit[unitId]];
        _fillRow(r, d, now, serial, r[C.CREATED_AT]);
        anyUpdates = true;
        updated++;
      } else {
        var nr = [];
        for (var c = 0; c < TOTAL_COLS; c++) nr.push('');
        _fillRow(nr, d, now, serial, now);
        newRows.push(nr);
        byUnit[unitId] = ownerRows.length;
        ownerRows.push(nr);
        imported++;
      }
    }

    // 3. Write back — at most 3 sheet operations
    var sheet = Database.getSheet(SHEET);
    if (sheet.getMaxColumns() < TOTAL_COLS) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), TOTAL_COLS - sheet.getMaxColumns());
    }

    // Always (re)write the header row (only over blank row 1 or an old header)
    var r1c1 = String(sheet.getRange(1, 1).getValue() || '');
    if (r1c1 === '' || r1c1 === 'owner_id') {
      sheet.getRange(1, 1, 1, TOTAL_COLS).setValues([OWNER_HEADERS]);
      sheet.getRange(1, 1, 1, TOTAL_COLS)
        .setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
    }

    var existingCount = ownerRows.length - newRows.length;
    if (anyUpdates && existingCount > 0) {
      sheet.getRange(2, 1, existingCount, TOTAL_COLS).setValues(ownerRows.slice(0, existingCount));
    }
    if (newRows.length > 0) {
      sheet.getRange(existingCount + 2, 1, newRows.length, TOTAL_COLS).setValues(newRows);
    }
    SpreadsheetApp.flush();

    return { imported: imported, updated: updated, errors: errors, schema: 'v3 — 45 columns (bank names + vehicles + corpus)' };
  }

  // "KL01CV4044/Maruti/WagonR/Red" -> {reg,make,model,colour}
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

  // Write all fields of d into row array (in place)
  function _fillRow(r, d, now, ownerId, createdAt) {
    r[C.OWNER_ID]          = ownerId;
    r[C.UNIT_ID]           = d.unit_id;
    r[C.NAME]              = d.name              || '';
    r[C.EMAIL]             = d.email             || '';
    r[C.PHONE]             = d.phone             || '';
    if (d.address !== undefined) r[C.ADDRESS] = d.address;
    r[C.NAME2]             = d.name2             || '';
    r[C.NAME3]             = d.name3             || '';
    r[C.PHONE2]            = d.phone2            || '';
    r[C.PHONE3]            = d.phone3            || '';
    r[C.EMAIL2]            = d.email2            || '';
    r[C.EXEC_MEMBER]       = d.exec_member       || 'No';
    r[C.EXEC_MEMBER_TITLE] = d.exec_member_title || '';
    r[C.CAR_PARKING_SLOT]  = d.car_parking_slot  || '';
    r[C.LIVING_STATUS]     = d.living_status     || 'Non-Resident';
    r[C.RESIDING_COUNTRY]  = d.residing_country  || '';
    r[C.ADDRESS_INDIA]     = d.address_india     || '';
    r[C.ADDRESS_ABROAD]    = d.address_abroad    || '';
    r[C.EMPLOYER]          = d.employer          || '';
    r[C.WORK_DESIGNATION]  = d.work_designation  || '';
    r[C.WORK_ADDRESS]      = d.work_address      || '';
    r[C.TC_NUMBER]         = d.tc_number         || '';
    r[C.KSEB_CONSUMER_NO]  = d.kseb_consumer_no  || '';
    r[C.NOTES]             = d.notes             || '';
    r[C.PROFILE_PICTURE]   = d.profile_picture   || '';
    if (d.vehicles_4w !== undefined) r[C.VEHICLES_4W] = d.vehicles_4w;
    if (!r[C.VEHICLES_4W]) r[C.VEHICLES_4W] = '[]';
    if (d.vehicles_2w !== undefined) r[C.VEHICLES_2W] = d.vehicles_2w;
    if (!r[C.VEHICLES_2W]) r[C.VEHICLES_2W] = '[]';
    if (d.bank_names !== undefined)          r[C.BANK_NAMES]          = d.bank_names;
    if (d.corpus_fund_pending !== undefined) r[C.CORPUS_FUND_PENDING] = d.corpus_fund_pending;
    if (d.num_4w !== undefined)              r[C.NUM_4W]              = d.num_4w;
    if (d.num_2w !== undefined)              r[C.NUM_2W]              = d.num_2w;
    r[C.PAY_JAN] = d.pay_jan || ''; r[C.PAY_FEB] = d.pay_feb || '';
    r[C.PAY_MAR] = d.pay_mar || ''; r[C.PAY_APR] = d.pay_apr || '';
    r[C.PAY_MAY] = d.pay_may || ''; r[C.PAY_JUN] = d.pay_jun || '';
    r[C.PAY_JUL] = d.pay_jul || ''; r[C.PAY_AUG] = d.pay_aug || '';
    r[C.PAY_SEP] = d.pay_sep || ''; r[C.PAY_OCT] = d.pay_oct || '';
    r[C.PAY_NOV] = d.pay_nov || ''; r[C.PAY_DEC] = d.pay_dec || '';
    r[C.CREATED_AT] = createdAt || now;
    r[C.UPDATED_AT] = now;
  }

  /**
   * applyMonthlyPayBatch — write pay_<mon> values for many units at once.
   * syncMap: { 'A101': { '2026-01': 2000, '2026-02': 2000 }, ... }
   * Only months of the current year land in the Jan–Dec columns.
   * One read, one write.
   */
  function applyMonthlyPayBatch(syncMap) {
    var units = syncMap ? Object.keys(syncMap) : [];
    if (!units.length) return { updated: 0 };

    var TOTAL_COLS = OWNER_HEADERS.length;
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

  // Next sequential owner number: max numeric owner_id + 1
  function _nextSerial() {
    var rows = Database.getAll(SHEET);
    var max  = 0;
    for (var i = 0; i < rows.length; i++) {
      var n = parseInt(rows[i][C.OWNER_ID], 10);
      if (!isNaN(n) && n > max) max = n;
    }
    return max + 1;
  }

  // Safe cell value helper. isMoney=true strips ₹ and thousand-separators.
  function _cv(row, colIdx, isMoney) {
    if (colIdx < 0 || colIdx >= row.length) return '';
    var v = String(row[colIdx] === null || row[colIdx] === undefined ? '' : row[colIdx]).trim();
    if (isMoney) v = v.replace(/[₹,\s]/g, '').trim();
    return v;
  }

  function _validate(data) {
    if (!data.unit_id) throw new Error('unit_id is required.');
    if (!data.name)    throw new Error('Owner name is required.');
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
      owner_id:            s(C.OWNER_ID),
      unit_id:             s(C.UNIT_ID),
      name:                s(C.NAME),
      email:               s(C.EMAIL),
      phone:               s(C.PHONE),
      address:             s(C.ADDRESS),
      name2:               s(C.NAME2),
      name3:               s(C.NAME3),
      phone2:              s(C.PHONE2),
      phone3:              s(C.PHONE3),
      email2:              s(C.EMAIL2),
      exec_member:         s(C.EXEC_MEMBER),
      exec_member_title:   s(C.EXEC_MEMBER_TITLE),
      car_parking_slot:    s(C.CAR_PARKING_SLOT),
      living_status:       s(C.LIVING_STATUS),
      residing_country:    s(C.RESIDING_COUNTRY),
      address_india:       s(C.ADDRESS_INDIA),
      address_abroad:      s(C.ADDRESS_ABROAD),
      employer:            s(C.EMPLOYER),
      work_designation:    s(C.WORK_DESIGNATION),
      work_address:        s(C.WORK_ADDRESS),
      tc_number:           s(C.TC_NUMBER),
      kseb_consumer_no:    s(C.KSEB_CONSUMER_NO),
      notes:               s(C.NOTES),
      profile_picture:     s(C.PROFILE_PICTURE),
      vehicles_4w:         s(C.VEHICLES_4W) || '[]',
      vehicles_2w:         s(C.VEHICLES_2W) || '[]',
      bank_names:          s(C.BANK_NAMES),
      corpus_fund_pending: s(C.CORPUS_FUND_PENDING),
      num_4w:              s(C.NUM_4W),
      num_2w:              s(C.NUM_2W),
      pay_jan: s(C.PAY_JAN), pay_feb: s(C.PAY_FEB),
      pay_mar: s(C.PAY_MAR), pay_apr: s(C.PAY_APR),
      pay_may: s(C.PAY_MAY), pay_jun: s(C.PAY_JUN),
      pay_jul: s(C.PAY_JUL), pay_aug: s(C.PAY_AUG),
      pay_sep: s(C.PAY_SEP), pay_oct: s(C.PAY_OCT),
      pay_nov: s(C.PAY_NOV), pay_dec: s(C.PAY_DEC),
      created_at:          s(C.CREATED_AT),
      updated_at:          s(C.UPDATED_AT)
    };
  }

  return {
    addOwner:          addOwner,
    getOwnersByUnit:   getOwnersByUnit,
    getOwnerById:      getOwnerById,
    updateOwner:       updateOwner,
    deleteOwner:       deleteOwner,
    deleteAllOwners:   deleteAllOwners,
    archiveOwner:      archiveOwner,
    getOwnerHistory:   getOwnerHistory,
    getAllOwners:       getAllOwners,
    searchOwners:      searchOwners,
    bulkImportFromCsv: bulkImportFromCsv,
    applyMonthlyPayBatch: applyMonthlyPayBatch
  };
})();
