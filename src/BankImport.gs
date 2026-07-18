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

      // Export first sheet as CSV
      var exportUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=csv&sheet=0';
      var token     = ScriptApp.getOAuthToken();
      var response  = UrlFetchApp.fetch(exportUrl, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });

      // Archive the ORIGINAL statement in the permanent repository folder —
      // an untouched reference copy for independent verification.
      try {
        var repo = _bankRepoFolder();
        file.setName('IMPORT ' + accountTag + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH.mm') + ' — ' + filename);
        file.moveTo(repo);
      } catch (mvErr) { file.setTrashed(true); }

      if (response.getResponseCode() !== 200) {
        return {
          success: false,
          error: 'Could not convert XLS to CSV (HTTP ' + response.getResponseCode() +
                 '). Please save your bank statement as CSV and try again.'
        };
      }

      csvText = response.getContentText();
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


// The permanent Drive folder holding an untouched copy of every imported
// bank statement ("OXYGEN Bank Statement Repository").
function _bankRepoFolder() {
  var it = DriveApp.getFoldersByName('OXYGEN Bank Statement Repository');
  return it.hasNext() ? it.next() : DriveApp.createFolder('OXYGEN Bank Statement Repository');
}
