// ═══════════════════════════════════════════════════════════════
// FormsService.gs — reads the community Google Forms' response
// spreadsheets LIVE (no imports needed): Move-In/Out, Party Hall
// Rental, and Vehicle Tracking. Each form's responses land in a
// linked Google Sheet; its URL is configured once on the Settings
// page, and OXYGEN reads it directly on every page load — new
// responses appear as soon as the page is refreshed.
// ═══════════════════════════════════════════════════════════════
var FormsService = (function() {

  // kind -> which Settings key holds the response spreadsheet URL
  var KIND_SETTINGS = {
    movein:    'form_movein_responses_url',
    partyhall: 'form_partyhall_responses_url',
    vehicle:   'form_vehicle_responses_url'
  };

  // Open the responses spreadsheet from whatever was pasted in Settings:
  // the responses SPREADSHEET's URL (the intended value), a bare
  // spreadsheet file ID, or the FORM's own edit URL — for a form URL the
  // linked responses spreadsheet is resolved automatically. The form's
  // PUBLIC link (…/forms/d/e/…/viewform) carries no usable ID, so that
  // one still fails — with a message saying exactly what to paste.
  function openResponsesSpreadsheet_(url) {
    if (/\/forms\//i.test(url)) {
      if (url.indexOf('/forms/d/e/') > -1) {
        throw new Error('this is the form’s PUBLIC link, which cannot lead back to its responses');
      }
      var destId = FormApp.openByUrl(url).getDestinationId(); // throws if no linked sheet / no access
      return SpreadsheetApp.openById(destId);
    }
    var m = url.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]{10,})/);
    if (m) return SpreadsheetApp.openById(m[1]);
    if (/^[A-Za-z0-9_-]{20,}$/.test(url)) return SpreadsheetApp.openById(url); // bare file ID
    return SpreadsheetApp.openByUrl(url);
  }

  function getResponses(kind) {
    var settingKey = KIND_SETTINGS[kind];
    if (!settingKey) throw new Error('Unknown form kind: ' + kind);

    var url = String(SettingsService.get(settingKey) || '').trim();
    if (!url) return { configured: false };

    var ss;
    try {
      ss = openResponsesSpreadsheet_(url);
    } catch (e) {
      throw new Error('Could not open the response spreadsheet. Check the URL saved in Settings → Community Forms (it must be the URL of the responses SPREADSHEET — in the form’s edit page: Responses tab → green Sheets icon → copy THAT tab’s address), and that this account has at least view access to it. Details: ' + ((e && e.message) ? e.message : e));
    }

    // Prefer the standard linked-responses tab; otherwise the first sheet.
    var sh = ss.getSheetByName('Form Responses 1') || ss.getSheets()[0];
    var values = sh.getDataRange().getDisplayValues();
    if (!values.length) return { configured: true, headers: [], rows: [], count: 0 };

    // Header row = the first row containing "Timestamp" (scanning the
    // top 10 rows). A live linked response sheet has it in row 1; the
    // association's own exported copies have a title band above it and
    // the header at row 4 — this handles both without configuration.
    var headerIdx = -1;
    for (var i = 0; i < Math.min(values.length, 10); i++) {
      for (var c = 0; c < values[i].length; c++) {
        if (String(values[i][c]).trim().toLowerCase() === 'timestamp') { headerIdx = i; break; }
      }
      if (headerIdx > -1) break;
    }
    if (headerIdx === -1) headerIdx = 0; // no Timestamp column at all — treat row 1 as headers

    var rawHeaders = values[headerIdx];
    // Drop columns whose header is blank or just "#" (the export's own
    // row-number column) — and remember which columns survive so data
    // rows keep aligned with their headers.
    var keepCols = [];
    var headers = [];
    for (var h = 0; h < rawHeaders.length; h++) {
      var name = String(rawHeaders[h]).trim();
      if (!name || name === '#') continue;
      keepCols.push(h);
      headers.push(name);
    }

    var rows = [];
    for (var r = headerIdx + 1; r < values.length; r++) {
      var row = [];
      var hasContent = false;
      for (var k = 0; k < keepCols.length; k++) {
        var cell = String(values[r][keepCols[k]] === undefined ? '' : values[r][keepCols[k]]).trim();
        if (cell) hasContent = true;
        row.push(cell);
      }
      if (hasContent) rows.push(row);
    }

    // Duplicate removal: two rows identical in EVERY column except the
    // Timestamp are the same submission entered twice (double-taps of
    // the form's Submit, and overlap between pasted-in historical rows
    // and live responses). The EARLIEST submission is kept, later
    // repeats dropped. The timestamp is deliberately excluded from the
    // comparison — it differs on every submission by definition.
    var tsCol = -1;
    for (var tc = 0; tc < headers.length; tc++) {
      if (headers[tc].toLowerCase() === 'timestamp') { tsCol = tc; break; }
    }
    var seen = {};
    var deduped = [];
    var duplicatesRemoved = 0;
    for (var dr = 0; dr < rows.length; dr++) {
      var sig = rows[dr].filter(function(_, ci) { return ci !== tsCol; })
                        .join('\u0001').toLowerCase();
      if (seen[sig]) { duplicatesRemoved++; continue; }
      seen[sig] = true;
      deduped.push(rows[dr]);
    }
    rows = deduped;

    // LIKELY duplicates — same Flat + Name (+ Move-In/Out direction when
    // that column exists) but differing in some detail. These are almost
    // always a resubmission with corrections (e.g. a fixed date), but
    // deleting one automatically would mean silently guessing which
    // version is right — and the same person CAN legitimately appear
    // twice (moved in, moved out, moved in again years later). So they
    // are flagged for human review, never removed.
    var flatCol = -1, nameCol = -1, dirCol = -1;
    for (var fc = 0; fc < headers.length; fc++) {
      var hl = headers[fc].toLowerCase();
      if (flatCol === -1 && hl === 'flat') flatCol = fc;
      if (nameCol === -1 && hl.indexOf('name') > -1) nameCol = fc;
      if (dirCol === -1 && hl.indexOf('move-in or move-out') > -1) dirCol = fc;
    }
    var likelyDup = [];
    if (flatCol > -1 && nameCol > -1) {
      var groups = {};
      rows.forEach(function(row, ri) {
        var gk = (row[flatCol] + '|' + row[nameCol] + (dirCol > -1 ? '|' + row[dirCol] : '')).toLowerCase();
        if (!groups[gk]) groups[gk] = [];
        groups[gk].push(ri);
      });
      likelyDup = rows.map(function() { return false; });
      for (var gk2 in groups) {
        if (groups[gk2].length > 1 && gk2.replace(/\|/g, '').trim()) {
          groups[gk2].forEach(function(ri2) { likelyDup[ri2] = true; });
        }
      }
    }

    // Newest first — form responses append chronologically, and the most
    // recent submissions are what the person opening the page wants.
    rows.reverse();
    likelyDup.reverse();

    return { configured: true, headers: headers, rows: rows, count: rows.length,
             duplicatesRemoved: duplicatesRemoved, likelyDup: likelyDup,
             sheetName: sh.getName(), spreadsheetName: ss.getName() };
  }

  // The internal vehicle registry — every vehicle recorded on owner and
  // tenant records (the Unit Profile data), flattened to one row per
  // vehicle. This is the always-current source: it updates the moment a
  // unit's profile changes, with no form submission needed. Vehicles are
  // stored on each record as JSON arrays of {reg, make, model, colour}.
  function getVehicleRegistry() {
    var out = [];

    function parseList(json) {
      try { var a = JSON.parse(json || '[]'); return (a && a.length) ? a : []; }
      catch (e) { return []; }
    }
    function pushAll(list, unitId, who, name, wheels) {
      list.forEach(function(v) {
        if (!v || (!v.reg && !v.make && !v.model)) return;
        out.push({ unit_id: unitId, who: who, name: name, wheels: wheels,
                   reg: v.reg || '', make: v.make || '', model: v.model || '', colour: v.colour || '' });
      });
    }

    OwnersService.getAllOwners().forEach(function(o) {
      pushAll(parseList(o.vehicles_4w), o.unit_id, 'Owner', o.name, '4-Wheeler');
      pushAll(parseList(o.vehicles_2w), o.unit_id, 'Owner', o.name, '2-Wheeler');
    });
    TenantsService.getAllTenants().forEach(function(t) {
      pushAll(parseList(t.vehicle_4w), t.unit_id, 'Tenant', t.name, '4-Wheeler');
      pushAll(parseList(t.vehicle_2w), t.unit_id, 'Tenant', t.name, '2-Wheeler');
    });

    out.sort(function(a, b) { return String(a.unit_id).localeCompare(String(b.unit_id)); });
    return { vehicles: out, count: out.length };
  }

  return {
    getResponses: getResponses,
    getVehicleRegistry: getVehicleRegistry
  };
})();
