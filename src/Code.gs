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
    // Google hides consumer visitors' emails from apps that "Execute as
    // Me", so email-based roles can only identify the owner. For
    // everyone else: check whether they already verified a PIN this
    // session (AuthService's cache, scoped automatically per visitor by
    // Apps Script itself — no token or cookie management needed here).
    if (!role && !email) {
      var cached = AuthService.getCachedIdentity();
      if (cached && cached.email) { email = cached.email; role = cached.role; }
    }
    // MEMBERS ONLY: OXYGEN is for the association's Executive Members
    // and office bearers. A visitor Google can't identify AND who has
    // no verified PIN gets the locked PIN screen — never the app. The
    // API routes enforce the same rule server-side (ApiRouter), so
    // this page is the front door, not the security boundary.
    if (!role && !email) return _lockedPinPage();
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

// The locked front door for unidentified visitors: OXYGEN-branded PIN
// screen, no app code or data anywhere in it. On a correct PIN the
// identity lands in AuthService's per-visitor cache, and the page
// re-navigates to the app URL — doGet then recognizes the visitor.
function _lockedPinPage() {
  var appUrl = ScriptApp.getService().getUrl();
  return HtmlService.createHtmlOutput(
    '<style>' +
      'html,body{margin:0;height:100%;background:#04101c;font-family:Arial,Helvetica,sans-serif;}' +
      '.wrap{min-height:100%;display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box;}' +
      '.card{background:#0E2D4A;border-radius:14px;padding:30px 28px;max-width:380px;width:100%;box-shadow:0 20px 60px rgba(0,0,0,.5);}' +
      'h2{margin:0 0 6px;font-size:18px;color:#fff;}' +
      'p{font-size:12.5px;color:#c8d9e6;line-height:1.6;margin:0 0 16px;}' +
      'input{width:100%;padding:11px 14px;font-size:15px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;color:#0d1b2a;box-sizing:border-box;margin-bottom:10px;}' +
      'button{width:100%;padding:11px;font-size:14px;font-weight:700;border:none;border-radius:8px;background:#06B6D4;color:#fff;cursor:pointer;}' +
      'button:disabled{opacity:.6;cursor:default;}' +
      '#err{color:#f87171;font-size:12px;min-height:16px;margin-bottom:10px;}' +
      '.foot{font-size:11px;color:#6E8CA3;margin-top:14px;text-align:center;line-height:1.5;}' +
    '</style>' +
    '<div class="wrap"><div class="card">' +
      '<h2>&#128274; OXYGEN &mdash; Private Application</h2>' +
      '<p>This is the internal management system of the <strong>Confident Daffodils Owners Association</strong>, ' +
      'for Executive Members and office bearers only.</p>' +
      '<p>Enter the PIN issued to you by the Administrator:</p>' +
      '<input id="pin" type="password" inputmode="numeric" placeholder="Enter your PIN" ' +
        'onkeydown="if(event.key===\'Enter\')go();">' +
      '<div id="err"></div>' +
      '<button id="btn" onclick="go()">Unlock</button>' +
      '<div class="foot">No PIN? Contact the association&#39;s Administrator.</div>' +
    '</div></div>' +
    '<script>' +
      'var APP_URL=' + JSON.stringify(appUrl) + ';' +
      'function go(){' +
        'var pin=document.getElementById("pin").value.trim();' +
        'var err=document.getElementById("err");var btn=document.getElementById("btn");' +
        'err.textContent="";' +
        'if(!pin){err.textContent="Enter your PIN.";return;}' +
        'btn.disabled=true;btn.textContent="Checking\\u2026";' +
        'google.script.run.withSuccessHandler(function(r){' +
          'if(r&&r.success){btn.textContent="\\u2713 Welcome";window.location.assign(APP_URL);}' +
          'else{err.textContent=(r&&r.error)||"PIN not recognized.";btn.disabled=false;btn.textContent="Unlock";}' +
        '}).withFailureHandler(function(e){' +
          'err.textContent=String(e&&e.message||e||"PIN not recognized.");btn.disabled=false;btn.textContent="Unlock";' +
        '}).doPost_internal("auth.verifyPin",{pin:pin});' +
      '}' +
      'setTimeout(function(){var i=document.getElementById("pin");if(i)i.focus();},60);' +
    '</scr' + 'ipt>'
  ).setTitle('OXYGEN — Private')
   .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
   .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
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
