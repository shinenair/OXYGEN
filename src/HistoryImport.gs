// ═══════════════════════════════════════════════════════════════
// HistoryImport.gs — one-time bulk import of HISTORIC owners &
// residents into the HISTORY_Owners / HISTORY_Residents archive
// sheets, which surface on each Unit Profile under the
// "🕓 View Unit Owner History" / "🕓 View Resident History" buttons.
//
// Called directly (not through ApiRouter):
//   google.script.run.importHistoryFile({ content: <base64 xlsx>, filename })
//
// The uploaded workbook has two tabs — "Owners History" and
// "Residents History" — in the exact review layout OXYGEN generated
// (columns mapped BY HEADER NAME, so extra columns the reviewer added,
// e.g. Work/Employer and Vehicle, are picked up automatically).
//
// PRIVACY: the residents'/owners' personal data travels only from the
// admin's browser into this Apps Script and on to the private Google
// Sheet. Nothing personal is stored in the (public) code repository —
// only this parser.
//
// IDEMPOTENT: every row it writes is tagged archived_by = IMPORT_TAG.
// A re-run first deletes all previously-imported rows (that tag only),
// then re-adds from the file — so re-importing never duplicates, and
// never touches history that was archived manually from the app.
// ═══════════════════════════════════════════════════════════════

var HISTORY_IMPORT_TAG = 'Historic import';

function importHistoryFile(payload) {
  try {
    UsersService.requireAdmin(); // admin-only; throws otherwise

    var content = (payload && payload.content) || '';
    if (!content) return { success: false, error: 'No file received.' };
    var filename = (payload && payload.filename) || 'history.xlsx';

    // Convert the uploaded .xlsx to a temporary Google Sheet so both
    // tabs can be read by name, then delete the temp file.
    var bytes = Utilities.base64Decode(content);
    var blob  = Utilities.newBlob(bytes,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename);

    var tempId = null, result;
    try {
      var created = Drive.Files.create(
        { name: 'OXYGEN History Import (temp)', mimeType: 'application/vnd.google-apps.spreadsheet' },
        blob
      );
      tempId = created.id;
      var ss = SpreadsheetApp.openById(tempId);
      result = _importHistoryFromWorkbook(ss);
    } finally {
      if (tempId) { try { Drive.Files.remove(tempId); } catch (eDel) {} }
    }
    return { success: true, data: result };
  } catch (err) {
    return { success: false, error: (err && err.message) ? err.message : String(err) };
  }
}

function _histNorm(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
}

function _histDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v == null ? '' : v).trim();
}

// Ensure a HISTORY sheet exists (mirrors the live sheet's own columns
// plus archived_at / archived_by) and return { sheet, headers }.
function _histEnsureSheet(liveName, histName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hist = ss.getSheetByName(histName);
  if (!hist) {
    var live = ss.getSheetByName(liveName);
    var base = live ? live.getRange(1, 1, 1, live.getLastColumn()).getValues()[0] : [];
    hist = ss.insertSheet(histName);
    hist.appendRow(base.concat(['archived_at', 'archived_by']));
  }
  return { sheet: hist, headers: hist.getRange(1, 1, 1, hist.getLastColumn()).getValues()[0] };
}

// Delete every previously-imported row (our tag only) from a HISTORY
// sheet, so a re-import replaces rather than piles up.
function _histClearImported(sheet, headers) {
  var byIdx = -1;
  for (var i = 0; i < headers.length; i++) if (_histNorm(headers[i]) === 'archived_by') byIdx = i;
  if (byIdx === -1) return 0;
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var vals = sheet.getRange(2, byIdx + 1, last - 1, 1).getValues();
  var removed = 0;
  for (var r = vals.length - 1; r >= 0; r--) {
    if (String(vals[r][0]) === HISTORY_IMPORT_TAG) { sheet.deleteRow(r + 2); removed++; }
  }
  return removed;
}

// Turn a { field_name: value } map into a row array positioned by the
// HISTORY sheet's own header order (unknown fields ignored, missing
// ones left blank).
function _histRow(headers, map) {
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    var key = _histNorm(headers[i]).replace(/ /g, '_');
    row.push(map.hasOwnProperty(key) ? map[key] : '');
  }
  return row;
}

function _importHistoryFromWorkbook(ss) {
  var now = new Date().toISOString();
  var sheets = ss.getSheets();
  var owners = { units: {}, count: 0 }, residents = { units: {}, count: 0 };
  var ownerRows = [], residentRows = [];

  sheets.forEach(function(sh) {
    var last = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (last < 2) return;
    var data = sh.getRange(1, 1, last, lastCol).getValues();
    var head = data[0].map(_histNorm);

    function col() {
      var args = arguments, a, c;
      // Exact header match wins first — so 'name' resolves to the "Name"
      // column, never to "Name 2".
      for (a = 0; a < args.length; a++)
        for (c = 0; c < head.length; c++) if (head[c] === args[a]) return c;
      // Then substring — lets 'vehicle' find "resident's vehicle details".
      for (a = 0; a < args.length; a++)
        for (c = 0; c < head.length; c++) if (head[c].indexOf(args[a]) > -1) return c;
      return -1;
    }

    var isResident = col('tenancy #', 'move-in') !== -1;
    var isOwner    = col('owner #') !== -1;
    if (!isResident && !isOwner) return;

    var cUnit = col('unit');

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var unit = String(row[cUnit] == null ? '' : row[cUnit]).trim().toUpperCase();
      var nameC = col('name');
      var name = nameC === -1 ? '' : String(row[nameC] == null ? '' : row[nameC]).trim();
      if (!unit || !name) continue; // skip blanks

      if (isOwner) {
        var oPhone2 = _cellAt(row, col('phone 2'));
        var oName2  = _cellAt(row, col('name 2'));
        var oName3  = _cellAt(row, col('name 3'));
        var oExec   = /^y/i.test(_cellAt(row, col('exec')));
        var map = {
          owner_id: 'HISTO-' + unit + '-' + (owners.count + 1),
          unit_id: unit, name: name,
          name2: oName2, name3: oName3,
          phone: _cellAt(row, col('phone')), phone2: oPhone2,
          email: _cellAt(row, col('email')), email2: _cellAt(row, col('email 2')),
          exec_member: oExec ? 'Yes' : '', exec_member_title: oExec ? 'Exec Member' : '',
          notes: _cellAt(row, col('notes')),
          created_at: now, updated_at: now,
          archived_at: now, archived_by: HISTORY_IMPORT_TAG
        };
        ownerRows.push(map);
        owners.count++; owners.units[unit] = true;
      } else {
        var work   = _cellAt(row, col('work'));
        var tName2 = _cellAt(row, col('name 2'));
        // Employer accidentally left in Name 2 (== Work) → drop it.
        if (work && tName2 && _histNorm(tName2) === _histNorm(work)) tName2 = '';
        var map2 = {
          tenant_id: 'HISTT-' + unit + '-' + (residents.count + 1),
          unit_id: unit, name: name, name2: tName2,
          email: _cellAt(row, col('email')),
          phone: _cellAt(row, col('phone')), phone2: _cellAt(row, col('phone 2')),
          work: work,
          vehicle_4w: _cellAt(row, col('vehicle')),
          move_in_date: _histDate(_rawAt(row, col('move-in'))),
          move_out_date: _histDate(_rawAt(row, col('move-out'))),
          comment: _cellAt(row, col('comment')),
          created_at: now, updated_at: now,
          archived_at: now, archived_by: HISTORY_IMPORT_TAG
        };
        residentRows.push(map2);
        residents.count++; residents.units[unit] = true;
      }
    }
  });

  // Write owners
  if (ownerRows.length) {
    var oH = _histEnsureSheet('Owners', 'HISTORY_Owners');
    _histClearImported(oH.sheet, oH.headers);
    var oOut = ownerRows.map(function(m) { return _histRow(oH.headers, m); });
    oH.sheet.getRange(oH.sheet.getLastRow() + 1, 1, oOut.length, oH.headers.length).setValues(oOut);
  }
  // Write residents
  if (residentRows.length) {
    var tH = _histEnsureSheet('Residents', 'HISTORY_Residents');
    _histClearImported(tH.sheet, tH.headers);
    var tOut = residentRows.map(function(m) { return _histRow(tH.headers, m); });
    tH.sheet.getRange(tH.sheet.getLastRow() + 1, 1, tOut.length, tH.headers.length).setValues(tOut);
  }

  return {
    owners: ownerRows.length, ownerUnits: _histKeys(owners.units).length,
    residents: residentRows.length, residentUnits: _histKeys(residents.units).length
  };
}

function _cellAt(row, idx) { return idx === -1 ? '' : String(row[idx] == null ? '' : row[idx]).trim(); }
function _rawAt(row, idx)  { return idx === -1 ? '' : row[idx]; }
function _histKeys(o) { var k = []; for (var x in o) if (o.hasOwnProperty(x)) k.push(x); return k; }
