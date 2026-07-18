// ============================================================
// TenantsImport.gs — One-click Tenants.xlsx bulk import
// Called as: google.script.run.importTenantsFile(payload)
// payload = { type: 'csv'|'xls', content: string, filename: string }
// Same proven pattern as OwnersImport.gs / BankImport.gs.
// ============================================================

function importTenantsFile(payload) {
  try {
    var type     = payload.type     || 'xls';
    var content  = payload.content  || '';
    var filename = payload.filename || 'Tenants.xlsx';

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
        return {
          success: false,
          error: 'Could not convert the Excel file (HTTP ' + response.getResponseCode() +
                 '). Please open it in Excel, use Save As -> CSV, and upload the CSV instead.'
        };
      }
      csvText = response.getContentText();
    }

    if (!csvText || csvText.trim().length < 20) {
      return { success: false, error: 'File appears to be empty or unreadable.' };
    }

    var rows = Utilities.parseCsv(csvText);
    if (!rows || rows.length < 2) {
      return { success: false, error: 'No data rows found in the file.' };
    }

    var result = TenantsService.bulkImportFromCsv(rows);
    return { success: true, data: result };

  } catch (err) {
    Logger.log('importTenantsFile error: ' + err.message);
    return { success: false, error: err.message };
  }
}

