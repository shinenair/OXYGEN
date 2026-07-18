// ============================================================
// Code.gs — Main Entry Point
// Confident Daffodils Property Management System
// ============================================================

// ── Run this ONCE from the editor to grant the email permission ──
// It directly calls the Session API, which forces Google to show the
// authorization dialog including the "view your email address" scope.
function authorizeEmailScope() {
  var email = Session.getActiveUser().getEmail();
  Logger.log('Authorized. Signed in as: ' + email);
  SpreadsheetApp.getActiveSpreadsheet().toast('Email scope authorized for ' + email);
}

function doGet(e) {
  try {
    var email, role;
    try {
      email = Session.getActiveUser().getEmail();
      role  = UsersService.getRole(email);
    } catch (authErr) {
      // The email scope has not been authorized yet — guide the owner.
      return HtmlService.createHtmlOutput(
        '<div style="font-family:sans-serif;max-width:560px;margin:80px auto;">' +
        '<h2>🔑 OXYGEN — One-time Authorization Needed</h2>' +
        '<p>The app now uses Google Sign-In for user roles, which needs a new permission.</p>' +
        '<p><strong>Owner:</strong> open the Apps Script editor, pick any function (e.g. <code>setupSpreadsheet</code>), click ▶ Run, and approve the authorization dialog. Then reload this page.</p>' +
        '<p style="color:#888;font-size:13px;">Technical detail: ' + authErr.message + '</p></div>'
      ).setTitle('OXYGEN — Authorize');
    }
    // INTERIM MODE: Google hides consumer visitors' emails from apps that
    // "Execute as Me", so email-based roles can only identify the owner.
    // For everyone else: check whether they already verified a PIN this
    // session (AuthService's cache, scoped automatically per visitor by
    // Apps Script itself — no token or cookie management needed here).
    // Only if that's ALSO empty does a visitor fall back to generic
    // access and see the PIN entry prompt client-side.
    if (!role && !email) {
      var cached = AuthService.getCachedIdentity();
      if (cached && cached.email) { email = cached.email; role = cached.role; }
    }
    if (!role && !email) { role = 'user'; email = 'Member (Google-verified)'; }
    if (!role) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:sans-serif;max-width:520px;margin:80px auto;text-align:center;">' +
        '<h2>🔒 OXYGEN — Access Restricted</h2>' +
        '<p>You are signed in as <strong>' + (email || 'an unknown account') + '</strong>, ' +
        'but this account has not been granted access.</p>' +
        '<p>Please ask an Administrator of the Confident Daffodils Owners Association to add your email in ' +
        '<strong>Settings → User Administration</strong>.</p></div>'
      ).setTitle('OXYGEN — Access Restricted');
    }
    var template = HtmlService.createTemplateFromFile('Index');
    template.appRole  = role;
    template.appEmail = email;
    return template.evaluate()
      .setTitle('OXYGEN')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (err) {
    Logger.log('doGet error: ' + err.message);
    return HtmlService.createHtmlOutput('<h2>Error: ' + err.message + '</h2>');
  }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var result  = ApiRouter.route(payload.action, payload.data || {});
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * One-time setup — run once from the Apps Script editor.
 * Initialises all sheets, seeds 136 units, and creates the BankStatements sheet.
 */
function setupSpreadsheet() {
  Database.initializeSheets();
  UnitsService.seedUnits();
  BankService.ensureSheet();
  SettingsService.ensureSheet();
  CategoriesService.ensureSheet();
  UnitsService.ensureColumns();
  TenantsService.ensureColumns();
  UsersService.ensureSheet();
  LPGReadingService.ensureColumns(); // LPGReadings sheet + headers (21 cols)
  Database.getSheet('LPGRates');      // ensures the rate-history sheet exists
  LPGInventoryService.ensureSheets();  // Stock Inward + Outward sheets
  ExpensePatternService.ensureSheet(); // Expense category auto-match patterns
  CaretakerService.ensureSheet();      // Caretaker petty-cash ledger
  Logger.log('Setup complete. All sheets initialized, units seeded, categories seeded.');
}
