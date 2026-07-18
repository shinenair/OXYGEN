// ═══════════════════════════════════════════════════════════════
// AuthCheck.gs — one-time authorization helper.
//
// Run checkAuthorization() from the editor toolbar after ANY change
// to the oauthScopes list in appsscript.json:
//   1. Select "checkAuthorization" in the function dropdown (top bar).
//   2. Click ▶ Run.
//   3. Google shows "Authorization required" → Review permissions →
//      pick the account that OWNS this deployment → Advanced →
//      "Go to OXYGEN (unsafe)" → Allow. The list MUST include
//      "See, edit, create, and delete all your Google Sheets
//      spreadsheets" — that's the new Sheets permission.
//   4. When it finishes, open View → Logs (Ctrl+Enter). Success looks
//      like:  ✅ OK — opened "…" and read N Party Hall responses.
//
// If the log shows ✅ but the web app STILL errors, the embedded site
// is running an old deployment version: Deploy → Manage deployments →
// ✏️ Edit → Version: "New version" → Deploy (this keeps the same URL
// your Google Site embeds).
// ═══════════════════════════════════════════════════════════════
function checkAuthorization() {
  var r = FormsService.getResponses('partyhall');
  if (!r.configured) {
    Logger.log('⚠️ No Party Hall responses URL is saved in Settings → Community Forms yet.');
    return;
  }
  Logger.log('✅ OK — opened "' + r.spreadsheetName + '" and read ' + r.count + ' Party Hall responses.');
}
