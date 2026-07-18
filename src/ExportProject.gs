function exportProjectAsZip() {
  const scriptId = ScriptApp.getScriptId();
  const url = 'https://www.googleapis.com/drive/v3/files/' + scriptId +
              '/export?mimeType=' +
              encodeURIComponent('application/vnd.google-apps.script+json');

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Export failed (' + response.getResponseCode() + '): ' +
                    response.getContentText());
  }

  const files = JSON.parse(response.getContentText()).files;

  const extensionFor = { SERVER_JS: '.gs', HTML: '.html', JSON: '.json' };

  const blobs = files.map(function (f) {
    const ext = extensionFor[String(f.type).toUpperCase()] || '.txt';
    return Utilities.newBlob(f.source, 'text/plain', f.name + ext);
  });

  const timestamp = Utilities.formatDate(new Date(),
      Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm');
  const zipBlob = Utilities.zip(blobs, 'OXYGEN_export_' + timestamp + '.zip');
  const zipFile = DriveApp.createFile(zipBlob);

  Logger.log('Exported ' + files.length + ' files.');
  Logger.log('Download here: ' + zipFile.getUrl());
}
