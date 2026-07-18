// ============================================================
// CaretakerFileImport.gs — one-click Excel import for the
// Caretaker's Monthly Expenses (from the association's own
// Caretaker's Monthly Report template). Same proven XLS->CSV
// conversion pattern used by every other importer in OXYGEN.
// ============================================================

function importCaretakerExpensesFile(payload, year, month) {
  try {
    var email = Session.getActiveUser().getEmail();
    if (UsersService.getRole(email) !== 'admin') throw new Error('Administrator access required to import Caretaker expenses.');

    var type     = payload.type     || 'xls';
    var content  = payload.content  || '';
    var filename = payload.filename || 'file.xlsx';

    var csvText = '';
    if (type === 'csv') {
      csvText = content;
    } else {
      var mimeType = filename.toLowerCase().indexOf('.xlsx') > -1
        ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        : 'application/vnd.ms-excel';
      var bytes = Utilities.base64Decode(content);
      var blob  = Utilities.newBlob(bytes, mimeType, filename);
      var file   = DriveApp.createFile(blob);
      var fileId = file.getId();
      var exportUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=csv&sheet=0';
      var token     = ScriptApp.getOAuthToken();
      var response  = UrlFetchApp.fetch(exportUrl, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });
      file.setTrashed(true);
      if (response.getResponseCode() !== 200) {
        return { success: false, error: 'Could not convert the Excel file (HTTP ' + response.getResponseCode() + '). Try Save As -> CSV and upload that instead.' };
      }
      csvText = response.getContentText();
    }

    if (!csvText || csvText.trim().length < 5) throw new Error('File appears to be empty or unreadable.');
    var rows = Utilities.parseCsv(csvText);
    if (!rows || rows.length < 1) throw new Error('No data rows found in the file.');

    var result = CaretakerService.importExpenses(rows, year, month);
    return { success: true, data: result };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

