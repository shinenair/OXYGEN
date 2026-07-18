// ============================================================
// LPGFileImport.gs — One-click LPG Readings Excel bulk import
// Called as: google.script.run.importLPGReadingsFile(payload, year, month)
// payload = { type: 'csv'|'xls', content: string, filename: string }
// Same proven conversion pattern as OwnersImport.gs / BankImport.gs.
// Expected columns (association's original template):
// # | FLAT NO. | Name01 | _ | _ | _ | Previous | Current | Reading Date |
// Volume(NM3) | Factor | Price/Kg | Amount | Payment Date | Paid | Remarks
// ============================================================

function importLPGReadingsFile(payload, year, month) {
  try {
    var type     = payload.type     || 'xls';
    var content  = payload.content  || '';
    var filename = payload.filename || 'LPG.xlsx';

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

    if (!csvText || csvText.trim().length < 10) {
      return { success: false, error: 'File appears to be empty or unreadable.' };
    }

    var rows = Utilities.parseCsv(csvText);
    if (!rows || rows.length < 2) {
      return { success: false, error: 'No data rows found in the file.' };
    }

    var result = LPGImport.importFromRows(rows, year, month);
    return { success: true, data: result };

  } catch (err) {
    return { success: false, error: err.message };
  }
}
