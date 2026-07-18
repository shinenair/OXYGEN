// ============================================================
// OwnersImport.gs — One-click Owners.xlsx bulk import
// Called as: google.script.run.importOwnersFile(payload)
// payload = { type: 'csv'|'xls', content: string, filename: string }
// Same proven pattern as BankImport.gs (no CDN libraries needed).
// ============================================================

function importOwnersFile(payload) {
  try {
    var type     = payload.type     || 'xls';
    var content  = payload.content  || '';
    var filename = payload.filename || 'Owners.xlsx';

    var csvText = '';

    if (type === 'csv') {
      csvText = content;

    } else {
      // XLS/XLSX — decode base64, upload to Drive with auto-convert,
      // export first sheet as CSV, then clean up.
      var mimeType = filename.toLowerCase().indexOf('.xlsx') > -1
        ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        : 'application/vnd.ms-excel';

      var bytes = Utilities.base64Decode(content);
      var blob  = Utilities.newBlob(bytes, mimeType, filename);

      // Upload to Drive, then export the first sheet as CSV
      // (identical to the proven BankImport.gs flow)
      var file   = DriveApp.createFile(blob);
      var fileId = file.getId();

      var exportUrl = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=csv&sheet=0';
      var token     = ScriptApp.getOAuthToken();
      var response  = UrlFetchApp.fetch(exportUrl, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });

      // Clean up the temp Drive file
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

    // Parse CSV safely (handles quoted fields with commas like "₹2,000")
    var rows = Utilities.parseCsv(csvText);
    if (!rows || rows.length < 2) {
      return { success: false, error: 'No data rows found in the file.' };
    }

    var result = OwnersService.bulkImportFromCsv(rows);
    return { success: true, data: result };

  } catch (err) {
    Logger.log('importOwnersFile error: ' + err.message);
    return { success: false, error: err.message };
  }
}

// ============================================================
// uploadOwnerPhoto — direct google.script.run entry point
// payload = { owner_id, content (base64), filename, mimeType }
// Saves the image to Drive folder "OXYGEN Owner Photos",
// makes it link-viewable, stores the URL on the owner record.
// ============================================================
function uploadOwnerPhoto(payload) {
  try {
    if (!payload || !payload.owner_id) return { success: false, error: 'No owner specified.' };
    if (!payload.content)              return { success: false, error: 'No image data received.' };

    var mimeType = payload.mimeType || 'image/jpeg';
    if (mimeType.indexOf('image/') !== 0) {
      return { success: false, error: 'Please choose an image file (JPG/PNG).' };
    }

    var bytes = Utilities.base64Decode(payload.content);
    var blob  = Utilities.newBlob(bytes, mimeType, payload.filename || ('owner_' + payload.owner_id + '.jpg'));

    // Find or create the photos folder
    var folderName = 'OXYGEN Owner Photos';
    var folders    = DriveApp.getFoldersByName(folderName);
    var folder     = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';

    // Save onto the owner record immediately
    OwnersService.updateOwner(String(payload.owner_id), { profile_picture: url });

    return { success: true, data: { url: url, file_id: file.getId() } };

  } catch (err) {
    Logger.log('uploadOwnerPhoto error: ' + err.message);
    return { success: false, error: err.message };
  }
}
