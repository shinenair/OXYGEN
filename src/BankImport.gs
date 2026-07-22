// ============================================================
// BankImport.gs — Direct google.script.run entry point
// Called as: google.script.run.importBankFile(payload)
// payload = { type: 'csv'|'xls', content: string, filename: string }
// ============================================================

/**
 * Main entry point called directly from the frontend via google.script.run.
 * Handles both CSV text and base64-encoded XLS/XLSX files.
 */
function importBankFile(payload) {
  try {
    // Which account this statement belongs to: '2' = the older IOB LPG
    // account (its own ledger sheet), anything else = the main account.
    var svc = (payload && String(payload.account) === '2') ? Bank2Service : BankService;
    var accountTag = (payload && String(payload.account) === '2') ? 'IOB-LPG ' : '';
    svc.ensureSheet();

    var type    = payload.type     || 'csv';
    var content = payload.content  || '';
    var filename= payload.filename || 'statement.xls';

    var csvText = '';

    if (type === 'csv') {
      // Already plain text — use directly
      csvText = content;

    } else {
      // XLS or XLSX — decode base64, save to Drive, export as CSV
      var mimeType = filename.toLowerCase().indexOf('.xlsx') > -1
        ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        : 'application/vnd.ms-excel';

      var bytes    = Utilities.base64Decode(content);
      var blob     = Utilities.newBlob(bytes, mimeType, filename);

      // Upload to Drive (auto-converts to Google Sheets)
      var file     = DriveApp.createFile(blob);
      var fileId   = file.getId();

      // ── Get a NATIVE Google Sheet, then read RAW cell values ──
      // The upload may or may not be auto-converted to a Google Sheet. If
      // opening it as a spreadsheet fails, EXPLICITLY convert the .xls with
      // the Drive API. Reading the real cell values (not a CSV export) is what
      // lets us fix the DD/MM date swap deterministically below.
      var sheetId = null, tempSheetId = null;
      try { SpreadsheetApp.openById(fileId).getName(); sheetId = fileId; } catch (eOpen) {}
      if (!sheetId) {
        try {
          var converted = Drive.Files.copy(
            { name: 'CONVERT ' + filename, mimeType: 'application/vnd.google-apps.spreadsheet' },
            fileId
          );
          if (converted && converted.id) { sheetId = converted.id; tempSheetId = converted.id; }
        } catch (eConv) {}
      }

      if (sheetId) {
        try {
          var conv  = SpreadsheetApp.openById(sheetId);
          var sh0   = conv.getSheets()[0];
          var lastR = sh0.getLastRow(), lastC = sh0.getLastColumn();
          if (lastR >= 1 && lastC >= 1) {
            var grid = sh0.getRange(1, 1, lastR, lastC).getValues();
            // Locale-independent date fix. Google may parse the statement's
            // DD/MM/YYYY dates with a MONTH-FIRST (US) locale, swapping day and
            // month (01/02 = 1 Feb becomes 2 Jan). Detect it: any converted real
            // Date whose day-of-month is > 12 proves DAY-FIRST (leave as-is); if
            // EVERY converted Date has day <= 12 it was month-first, so swap day
            // and month back. Dates that stayed text (day > 12 can't be a US
            // month) are handled day-first downstream. Output unambiguous ISO.
            var sawDate = false, dayFirstConv = false;
            for (var r = 0; r < grid.length; r++) {
              for (var c = 0; c < grid[r].length; c++) {
                var cv = grid[r][c];
                if (cv instanceof Date && !isNaN(cv.getTime())) { sawDate = true; if (cv.getDate() > 12) dayFirstConv = true; }
              }
            }
            var swap = sawDate && !dayFirstConv;
            var pad2 = function (n) { return n < 10 ? '0' + n : String(n); };
            var cellStr = function (v) {
              if (v instanceof Date && !isNaN(v.getTime())) {
                var mo = v.getMonth() + 1, dy = v.getDate();
                if (swap) { var t = mo; mo = dy; dy = t; }
                return v.getFullYear() + '-' + pad2(mo) + '-' + pad2(dy);
              }
              return (v === null || v === undefined) ? '' : String(v);
            };
            csvText = grid.map(function (row) { return row.map(function (v) { return _csvQuote(cellStr(v)); }).join(','); }).join('\n');
          }
        } catch (eRead) { csvText = ''; }
        if (tempSheetId) { try { DriveApp.getFileById(tempSheetId).setTrashed(true); } catch (eTr) {} }
      }

      if (!csvText) {
        // Last resort: locale-blind CSV export (may mis-order ambiguous dates).
        var exportUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=csv&sheet=0';
        var token     = ScriptApp.getOAuthToken();
        var response  = UrlFetchApp.fetch(exportUrl, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
        if (response.getResponseCode() !== 200) {
          try { file.setTrashed(true); } catch (e0) {}
          return { success: false, error: 'Could not read the statement (HTTP ' + response.getResponseCode() + '). Please save your bank statement as CSV and try again.' };
        }
        csvText = response.getContentText();
      }

      // Archive the ORIGINAL statement in the permanent repository folder —
      // an untouched reference copy for independent verification.
      try {
        var repo = _bankRepoFolder();
        file.setName('IMPORT ' + accountTag + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH.mm') + ' — ' + filename);
        file.moveTo(repo);
      } catch (mvErr) { try { file.setTrashed(true); } catch (e1) {} }
    }

    if (!csvText || csvText.trim().length < 20) {
      return { success: false, error: 'File appears to be empty or unreadable.' };
    }

    // CSV uploads are archived too (XLS originals are archived above)
    if (type === 'csv') {
      try {
        _bankRepoFolder().createFile(Utilities.newBlob(csvText, 'text/csv',
          'IMPORT ' + accountTag + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH.mm') + ' — ' + filename));
      } catch (csvErr) {}
    }

    var result = svc.importCsv(csvText);
    return { success: true, data: result };

  } catch (err) {
    Logger.log('importBankFile error: ' + err.message);
    return { success: false, error: err.message };
  }
}


// Quote one cell for CSV assembly from a 2-D array of display values:
// wrap in quotes and double any embedded quotes when it contains a comma,
// quote or newline.
function _csvQuote(v) {
  var s = (v === null || v === undefined) ? '' : String(v);
  if (s.indexOf('"') > -1 || s.indexOf(',') > -1 || s.indexOf('\n') > -1 || s.indexOf('\r') > -1) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

// The permanent Drive folder holding an untouched copy of every imported
// bank statement ("OXYGEN Bank Statement Repository").
function _bankRepoFolder() {
  var it = DriveApp.getFoldersByName('OXYGEN Bank Statement Repository');
  return it.hasNext() ? it.next() : DriveApp.createFolder('OXYGEN Bank Statement Repository');
}
