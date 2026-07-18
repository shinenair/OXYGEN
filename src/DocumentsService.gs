// ═══════════════════════════════════════════════════════════════
// DocumentsService.gs — lists the association's shared "Documents"
// Google Drive folder live inside OXYGEN. The folder URL is set on
// the Settings page; files added to Drive appear here on refresh.
// ═══════════════════════════════════════════════════════════════
var DocumentsService = (function() {

  function _folderIdFromUrl(url) {
    var m = String(url || '').match(/folders\/([A-Za-z0-9_-]+)/);
    return m ? m[1] : '';
  }

  function listDocuments() {
    var url = String(SettingsService.get('documents_folder_url') || '').trim();
    if (!url) return { configured: false };
    var folderId = _folderIdFromUrl(url);
    if (!folderId) throw new Error('The Documents folder URL saved in Settings does not look like a Drive folder link (it must contain /folders/<id>).');

    var folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      throw new Error('Could not open the Documents folder. Check the URL in Settings → Documents, and that this account has access to the folder.');
    }

    var items = [];

    // Subfolders first — they open in Drive directly.
    var folders = folder.getFolders();
    while (folders.hasNext()) {
      var f = folders.next();
      items.push({ kind: 'folder', name: f.getName(), url: f.getUrl(),
                   updated: Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), 'dd-MMM-yyyy'),
                   size: '' });
    }

    var files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var bytes = file.getSize();
      var size = bytes >= 1048576 ? (Math.round(bytes / 104857.6) / 10) + ' MB'
               : bytes >= 1024    ? Math.round(bytes / 1024) + ' KB'
               : bytes + ' B';
      items.push({ kind: 'file', name: file.getName(), url: file.getUrl(),
                   mime: file.getMimeType(),
                   updated: Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'dd-MMM-yyyy'),
                   size: size });
    }

    // Folders first, then files, each alphabetically.
    items.sort(function(a, b) {
      if (a.kind !== b.kind) return a.kind === 'folder' ? -1 : 1;
      return a.name.localeCompare(b.name);
    });

    return { configured: true, folderName: folder.getName(), folderUrl: folder.getUrl(),
             items: items, count: items.length };
  }

  return {
    listDocuments: listDocuments
  };
})();

