// ═══════════════════════════════════════════════════════════════
// HelpService.gs — OXYGEN's Help documentation.
//
// Content is defined ONCE, here, as a simple block schema:
//   { t: 'h1'|'h2'|'h3'|'p'|'b'|'n'|'hr', x: 'text' }
//   h1/h2/h3 = headings, p = paragraph, b = bullet, n = numbered
//   item (numbering resets whenever a non-'n' block appears), hr = divider
// This same array renders as HTML in the Help page AND as a properly
// formatted Google Doc for the "Download PDF" button — one source,
// never maintained twice.
// ═══════════════════════════════════════════════════════════════
var HelpService = (function() {

  var VERSION = '1.0';
  var RELEASE_DATE = '15 July 2026';

  function _about() {
    return [
      { t: 'h1', x: 'OXYGEN — Property Management System' },
      { t: 'p', x: 'Version ' + VERSION + '  ·  Released ' + RELEASE_DATE },
      { t: 'hr' },
      { t: 'p', x: 'OXYGEN is the property management system built for Confident Daffodils Owners Association (CDOA) — 136 units across two towers (Tower A, A101–A808; Tower B, B101–B908).' },
      { t: 'p', x: 'It runs entirely on Google Sheets and Google Apps Script — no separate database, no external hosting, no subscription. Everything the association needs lives in one Google Sheet, with OXYGEN as the application layer on top of it: bank statement matching and reconciliation, fees tracking, reports, the EC Committee and its payment follow-up board, community forms, a shared document library, the corpus fund, an annual event planner, and administrator access control.' },
      { t: 'h2', x: 'Who this is for' },
      { t: 'p', x: 'OXYGEN is built primarily for the EC Committee and Administrators who run CDOA day to day — not for general resident self-service. Day-to-day operation (importing bank statements, recording payments, generating reports, managing the committee) is the core use case this Help documentation is written for.' },
      { t: 'h2', x: 'How it\'s accessed' },
      { t: 'p', x: 'OXYGEN is deployed as a Google Apps Script web app, embedded in a Google Site. Anyone with a Google account added to the association\'s Users list can sign in. On personal (non-Workspace) Google accounts, Google sometimes cannot reveal a visitor\'s email automatically — OXYGEN\'s PIN login (Settings → User Administration) exists specifically to identify people correctly in that situation.' },
      { t: 'h2', x: 'Built by' },
      { t: 'p', x: 'Developed collaboratively with Claude (Anthropic) over an extended series of working sessions with CDOA\'s administrator, growing from a payments tracker into the full system described in this Help section.' }
    ];
  }

  function _releaseNotes() {
    return [
      { t: 'h1', x: 'Release Notes — Version ' + VERSION },
      { t: 'p', x: 'Released ' + RELEASE_DATE + '. This is OXYGEN\'s first numbered release — the notes below are a complete tour of what the system does today, rather than a change-by-change history.' },
      { t: 'hr' },

      { t: 'h2', x: '🏦 Bank Statement / Transaction Ledger' },
      { t: 'p', x: 'Two fully independent bank accounts — the main IOB account and the older, largely dormant IOB LPG account — each with its own import, matching, and reconciliation, never mixed together.' },
      { t: 'b', x: 'Import monthly (or more frequent, partial) statements in XLS or CSV; duplicate transactions are automatically skipped on re-import.' },
      { t: 'b', x: 'Automatic name-matching against Owners and Tenants, with manual Match/Edit/Unmatch always available.' },
      { t: 'b', x: 'A narration cross-check warns when the payer typed a flat number into the payment remark that disagrees with the matched unit — catching wrong fuzzy-name matches.' },
      { t: 'b', x: '🧩 Assign Portions handles the edge cases a simple one-unit-one-type match can\'t: one payment covering two flats, or one payment covering Maintenance + Waste Management + LPG together. Portions create parked funds only — they never post to Fees Received automatically.' },
      { t: 'b', x: 'Spreadsheet-style value filtering on the Credit column.' },
      { t: 'b', x: 'Balance reconciliation, one year at a time, with an automatic flag for months that look partially imported and explicit detection of any month with zero transactions sitting between two months that do have data.' },
      { t: 'b', x: 'Danger Zone: delete a month\'s bank data cleanly for re-import, with its Fees Received records and portions cleaned up automatically.' },

      { t: 'h2', x: '📥 Fees Received' },
      { t: 'p', x: 'The monthly grid for Maintenance, Waste Management, and LPG fees, per unit.' },
      { t: 'b', x: 'Simple, one-unit one-type payments post automatically from the bank statement.' },
      { t: 'b', x: 'Money from a portioned (edge-case) transaction sits as parked funds on the Unit Profile until a human manually allocates it to specific months — the system never guesses which months a bulk payment covers.' },
      { t: 'b', x: 'A cell\'s 🔎 inspector shows every record stacked behind its total, with its origin (auto-posted, manual allocation, legacy, etc.) and a surgical single-record delete.' },
      { t: 'b', x: 'Waste Management and LPG applicability follows each unit\'s Monthly Occupancy & LPG Mode timeline — correctly handling a unit that was vacant in January, occupied in February, and vacant again in March.' },

      { t: 'h2', x: '📄 Reports' },
      { t: 'p', x: 'Six report types, each previewed on screen before generating a PDF: Treasurer\'s Monthly Report, Caretaker\'s Monthly Report, Monthly Income, Monthly Expense, Annual Income, and Annual Expense.' },
      { t: 'b', x: 'The Treasurer\'s report includes both bank accounts\' positions and a combined total.' },
      { t: 'b', x: 'Reports are cash-basis and split-aware — a combined payment\'s portions are correctly attributed to each fee type.' },

      { t: 'h2', x: '🗳️ EC Committee' },
      { t: 'b', x: 'Current committee roster with contacts, duty assignments, and each member\'s allocated units for payment follow-up.' },
      { t: 'b', x: '⚖️ Auto-Distribute splits all 136 units equally among members flagged for allocation.' },
      { t: 'b', x: 'The 🎯 Follow-up Board is a live, color-coded view of who\'s paid MF/WMF/LPG each month, per EC member — exportable as a landscape PDF for sharing.' },
      { t: 'b', x: '📦 Archive Committee preserves an outgoing committee\'s full record when a new one takes over.' },

      { t: 'h2', x: '📸 Payment Screenshots' },
      { t: 'p', x: 'The last-resort identification method: residents\' WhatsApp payment confirmations, uploaded to a Drive folder, are automatically text-recognized (OCR) and matched to unmatched bank transactions by UPI reference number — a near-certain identifier shared by both the screenshot and the bank\'s own narration.' },

      { t: 'h2', x: '📅 Year Planner' },
      { t: 'p', x: 'A color-coded annual calendar of recurring CDOA events — bills, meetings, contract renewals — reproduced from the association\'s own template, with automatic email reminders to EC Members (1 month / 1 week / 1 day before) and one-click WhatsApp-ready reminder messages. Syncs to a dedicated Google Calendar.' },

      { t: 'h2', x: '💰 Corpus Fund (Administrators only)' },
      { t: 'p', x: 'CDOA-wide owed/received summary, and any number of Fixed Deposits with IOB, each with its own statement-of-account ledger. Every Fixed Deposit\'s balance computes itself live from its own Deposit/Interest/Withdrawal entries — no manual arithmetic.' },

      { t: 'h2', x: '🏠 Community pages' },
      { t: 'p', x: 'Live views of the Move-In/Out, Party Hall Rental, and Vehicle Tracking Google Forms\' responses, plus a Vehicle Registry built directly from owner and tenant records, and a Documents page listing the association\'s shared Drive folder.' },

      { t: 'h2', x: '🔐 Access & Security' },
      { t: 'b', x: 'Administrator and User roles, managed in Settings → User Administration.' },
      { t: 'b', x: 'PIN-based fallback login for Google accounts the platform can\'t automatically identify — a real limitation of personal (non-Workspace) Google deployments, not an OXYGEN restriction.' },
      { t: 'b', x: 'Every Admin-only page (Settings, Data Imports, Corpus Fund) enforces its restriction on the server, not just by hiding the menu item.' },

      { t: 'h2', x: 'Known limitations' },
      { t: 'b', x: 'WhatsApp reminders are one-click-to-send, not fully automatic — there is no free way for a program to send WhatsApp messages directly.' },
      { t: 'b', x: 'A PIN-verified session is remembered for up to 6 hours (a Google platform limit), after which the PIN may need re-entering.' },
      { t: 'b', x: 'Apps Script projects have a 200-version ceiling with no way to delete old versions; very long-lived projects will eventually need a fresh copy of the underlying Sheet.' }
    ];
  }

  function _instructions() {
    return [
      { t: 'h1', x: 'Instructions for Use' },
      { t: 'p', x: 'Short, task-first instructions for the things you\'ll do most often. For full detail on any page, see the User Manual.' },
      { t: 'hr' },

      { t: 'h2', x: '1. Import a bank statement' },
      { t: 'n', x: 'Open 🏦 Transaction Ledger, and pick the correct account tab (Main IOB or IOB LPG).' },
      { t: 'n', x: 'Under "Import Bank Statement", choose your XLS or CSV file and click Import.' },
      { t: 'n', x: 'Names are matched automatically. Re-importing the same file is safe — duplicates are skipped.' },

      { t: 'h2', x: '2. Fix a wrong match' },
      { t: 'n', x: 'Click Edit on the transaction, then ⛓️‍💥 Unmatch to clear it, or pick the correct owner/tenant directly.' },
      { t: 'n', x: 'If the match had already posted a Fees Received record, you\'ll be shown exactly what would be removed before confirming.' },

      { t: 'h2', x: '3. Handle a payment covering more than one unit or fee type' },
      { t: 'n', x: 'Click 🧩 on the transaction and add one row per unit/fee-type/amount — they must sum to the transaction total.' },
      { t: 'n', x: 'This parks the money per unit; it does NOT post to Fees Received automatically.' },
      { t: 'n', x: 'Open the affected Unit Profile → 🅿️ Parked Funds → 📅 Allocate months, and type which months the parked money covers.' },

      { t: 'h2', x: '4. Record an LPG meter reading and see if it\'s been paid correctly' },
      { t: 'n', x: 'Go to Fees Received → LPG tab → 🔥 Mark Paid pulls the suggested amount straight from the meter reading if one exists.' },
      { t: 'n', x: 'On the Unit Profile, the LPG Reading History table compares Calculated vs. Paid automatically — "Paid" always looks at the following month\'s actual payment, matching the real billing lag.' },

      { t: 'h2', x: '5. Mark a unit vacant (or occupied again) for specific months' },
      { t: 'n', x: 'Open the Unit Profile → Monthly Occupancy & LPG Mode Timeline.' },
      { t: 'n', x: 'Click any month to set its value — every later month automatically inherits it, until you set a different value for some later month.' },

      { t: 'h2', x: '6. Generate the Treasurer\'s Report' },
      { t: 'n', x: 'Open 📄 Reports → Treasurer\'s Monthly Report, choose the year and month, review the preview, then Generate PDF.' },

      { t: 'h2', x: '7. Distribute units among EC Committee members' },
      { t: 'n', x: 'Open 🗳️ EC Committee → Current Committee, add/edit members and tick "Include in unit allocation".' },
      { t: 'n', x: 'Click ⚖️ Auto-Distribute Units to split all 136 units equally, in order.' },

      { t: 'h2', x: '8. Send the Follow-up Board to the EC group' },
      { t: 'n', x: 'Open 🎯 Follow-up Board, choose the month, and either screenshot the on-screen board or click 📄 Generate PDF.' },

      { t: 'h2', x: '9. Give someone a PIN login' },
      { t: 'n', x: 'Open ⚙️ Settings → User Administration, add the person\'s email if not already listed, and click 🔑 Set PIN.' },
      { t: 'n', x: 'Tell them the PIN directly — it cannot be looked up again later, only reset.' },

      { t: 'h2', x: '10. Add a recurring event to the Year Planner' },
      { t: 'n', x: 'Open 📅 Year Planner → ⚙️ Manage Events → ➕ Add Event, choose a pattern (fixed date, specific months, last day of month, etc.).' },
      { t: 'n', x: 'Click 🔗 Sync to Google Calendar to push it to the dedicated CDOA calendar.' },

      { t: 'h2', x: '11. Review and match a payment screenshot' },
      { t: 'n', x: 'Open 📸 Payment Screenshots → 📷 Scan for New Screenshots.' },
      { t: 'n', x: 'For any row showing "⚡ Exact UTR found", click 🔗 Match, confirm the unit and fee type, and it posts through the normal ledger match.' },

      { t: 'h2', x: '12. Add a new Corpus Fund Fixed Deposit' },
      { t: 'n', x: 'Open 💰 Corpus Fund (Administrators only) → ➕ Add New Fixed Deposit, fill in the details, then add Deposit/Interest lines to its statement as they occur — the balance computes itself.' },

      { t: 'h2', x: '13. Check that a bank account\'s numbers actually reconcile' },
      { t: 'n', x: 'Open 🏦 Transaction Ledger → 🔍 Reconcile Balances, pick the year. Red rows need investigation; 🟡 rows likely just need the rest of that month imported; ⛔ rows are months with no data at all.' },

      { t: 'h2', x: '14. Safely delete and re-import a month\'s bank data' },
      { t: 'n', x: 'Open ⚙️ Settings → Danger Zone → Delete Bank Statement Data by Month, select the month, and confirm.' },
      { t: 'n', x: 'This also removes that month\'s Fees Received records, so the two can never disagree.' }
    ];
  }

  function _manual() {
    return [
      { t: 'h1', x: 'User Manual' },
      { t: 'p', x: 'A complete reference to every page in OXYGEN, organized to match the sidebar. Written for the EC Committee and Administrators who operate the system day to day.' },
      { t: 'hr' },

      { t: 'h2', x: '📊 Dashboard' },
      { t: 'p', x: 'The landing page — a snapshot of collections, defaulters, and key figures at a glance when you first sign in.' },

      { t: 'h2', x: '🏠 Units, Owners, Tenants' },
      { t: 'p', x: 'The master records for all 136 units, their owners, and current tenants. Each unit carries flags — Occupancy, Registration status, Corpus contribution, LPG Mode (Using association LPG / Own Cylinder / Not Using) — that drive fee applicability and reporting throughout the rest of the app.' },

      { t: 'h2', x: '🏢 Unit — Full Profile' },
      { t: 'p', x: 'Opened by clicking any unit anywhere in the app. Shows owner and tenant details, all bank transactions matched to the unit (across both accounts), fee payment history, and three specialized sections:' },
      { t: 'b', x: '🅿️ Parked Funds — money assigned to this unit via a portioned bank transaction, awaiting manual month allocation. Only appears when there is parked money to show.' },
      { t: 'b', x: 'Monthly Occupancy & LPG Mode Timeline — the month-by-month history of whether the unit was occupied and how it used LPG. A value set for one month carries forward automatically until a later month changes it. The 🔍 "See all explicit entries" link shows every entry ever set, across all years, useful for tracing why a value did or didn\'t carry forward as expected.' },
      { t: 'b', x: 'LPG Reading History — meter readings with a live comparison of Calculated vs. Paid (Paid is always looked up from the following month\'s actual payment). Includes a year selector, a totals row for the whole year, and an Administrator-only 🚩 flag with a note for any month that looks wrong.' },

      { t: 'h2', x: '🏦 Bank Statement / Transaction Ledger' },
      { t: 'p', x: 'Two account tabs — Main IOB Account and IOB LPG Account — completely independent: separate transactions, separate reconciliation, separate Danger Zone entries.' },
      { t: 'b', x: 'Import: XLS or CSV, duplicate-safe on re-import.' },
      { t: 'b', x: 'Match / Edit / Unmatch: automatic name-matching with manual override always available. A ⚠️ warning appears when the payer\'s own narration disagrees with the matched unit.' },
      { t: 'b', x: '🧩 Assign Portions: for one payment covering multiple units and/or multiple fee types. Creates parked funds only — allocation to specific months happens on the Unit Profile.' },
      { t: 'b', x: 'Filters: search, match status, fee type, expense category, year/month, and a spreadsheet-style value filter on the Credit column.' },
      { t: 'b', x: '🔍 Reconcile Balances: opening + credits − debits must equal the bank\'s own stated closing balance, one year at a time. Flags likely-partial months and any month with zero transactions sitting between two months that do have data.' },
      { t: 'b', x: '🔄 Re-run Matching: re-applies automatic name-matching to whatever is currently Unmatched. Never touches a transaction that already has a real match.' },

      { t: 'h2', x: '📥 Fees Received' },
      { t: 'p', x: 'The monthly grid — Maintenance, Waste Management, LPG — per unit, per year.' },
      { t: 'b', x: 'A green cell means at least one non-Rejected payment record exists for that month; click it to open the 🔎 inspector showing every record behind the total, each with its origin and a delete option.' },
      { t: 'b', x: 'A red cell with + means nothing is recorded yet — click to record a payment manually.' },
      { t: 'b', x: 'Waste Management and LPG cells automatically show as not-applicable for a month where the unit\'s occupancy or LPG Mode timeline says so.' },

      { t: 'h2', x: '📄 Reports' },
      { t: 'p', x: 'Six report tiles, each with an on-screen preview before generating a PDF: Treasurer\'s Monthly Report, Caretaker\'s Monthly Report, Monthly Income, Monthly Expense, Annual Income, Annual Expense. All are split-aware (a portioned payment\'s pieces are attributed to their real fee types) and cash-basis (an amount belongs to the month the money actually moved).' },

      { t: 'h2', x: '📦 Move-In / Move-Out, 🎉 Party Hall Rental, 🚗 Vehicle Tracking' },
      { t: 'p', x: 'Live views of each Google Form\'s responses, searchable and numbered. Vehicle Tracking\'s primary view is a Registry built directly from Owner and Tenant vehicle records (more complete than the form, which predates move-in vehicle collection); Form Responses is available as a secondary tab.' },

      { t: 'h2', x: '📁 Documents' },
      { t: 'p', x: 'A live listing of the association\'s shared Google Drive documents folder — anything added there appears here on refresh, with a direct link to open each file.' },

      { t: 'h2', x: '🗳️ EC Committee' },
      { t: 'b', x: 'Current Committee: contacts, duty assignments, and each member\'s allocated units. ⚖️ Auto-Distribute splits all units equally among members flagged for allocation. 📦 Archive Committee preserves the outgoing roster before a new one is entered.' },
      { t: 'b', x: '🎯 Follow-up Board: a color-coded month-by-month view of who has paid MF/WMF/LPG, grouped by the EC member responsible for each unit. Exportable as a PDF for sharing.' },
      { t: 'b', x: '📜 Past Committees: every archived committee, grouped by term.' },

      { t: 'h2', x: '📸 Payment Screenshots' },
      { t: 'p', x: 'Residents\' WhatsApp payment confirmations, uploaded to a designated Drive folder (name each file with the flat number, e.g. A101_15Jan.jpg), automatically text-recognized and matched to unmatched bank transactions by UPI reference number. Every extracted field is editable before matching; matching posts through the normal ledger pipeline.' },

      { t: 'h2', x: '📅 Year Planner' },
      { t: 'p', x: 'A computed, color-coded 12-month calendar of recurring CDOA events, editable under ⚙️ Manage Events (fixed annual dates, specific-months patterns, last-day-of-month, last-Sunday-of-month, or one-off dates). Reminders go to current EC Members only, by email automatically (1 month / 1 week / 1 day before, once the one-time daily trigger is set up) and by a one-click WhatsApp-ready message. 🔗 Sync to Google Calendar pushes the whole year to a dedicated calendar.' },

      { t: 'h2', x: '💰 Corpus Fund — Administrators only' },
      { t: 'p', x: 'CDOA-wide Owed and Received summaries, and any number of Fixed Deposits with IOB. Each Fixed Deposit has its own Deposit/Interest/Withdrawal statement of account, and its balance is always computed live from those entries — never a manually maintained number.' },

      { t: 'h2', x: '⚙️ Settings — Administrators only' },
      { t: 'b', x: 'Payment Fees, Community Forms, Documents, Payment Screenshots: configuration values used elsewhere in the app.' },
      { t: 'b', x: 'User Administration: add/remove people, set their role (Administrator or User), and set or reset a PIN — the fallback login for Google accounts the platform can\'t automatically identify.' },
      { t: 'b', x: 'Danger Zone: month-by-month deletion of bank statement data, LPG readings, and stock records, each showing exactly what will be removed before confirming.' },

      { t: 'h2', x: '📥 Data Imports — Administrators only' },
      { t: 'p', x: 'Bulk import tools for Owners, Tenants, and historical data, used during initial setup or when bringing in a large batch of records at once.' },

      { t: 'h2', x: '🔐 Access, roles, and PINs' },
      { t: 'p', x: 'Everyone signs in with a Google account. Their role — Administrator or User — determines what they can see and do; Administrator-only pages and actions are enforced on the server, not just hidden from the menu. On personal (non-Workspace) Google accounts, the platform sometimes cannot reveal a visitor\'s email at all — when that happens, a PIN entry prompt appears; entering a PIN set for that person by an Administrator identifies them correctly for up to 6 hours.' }
    ];
  }

  var DOCS = { about: _about, release_notes: _releaseNotes, instructions: _instructions, manual: _manual };
  var DOC_TITLES = { about: 'About OXYGEN', release_notes: 'Release Notes', instructions: 'Instructions for Use', manual: 'User Manual' };

  function getContent(docKey) {
    var fn = DOCS[docKey];
    if (!fn) throw new Error('Unknown document: ' + docKey);
    return { title: DOC_TITLES[docKey], version: VERSION, release_date: RELEASE_DATE, blocks: fn() };
  }

  function getVersionInfo() {
    return { version: VERSION, release_date: RELEASE_DATE };
  }

  // ── PDF generation, via a temporary Google Doc (proper prose
  //    formatting — headings, paragraphs, bullets — unlike the
  //    spreadsheet-cell approach used for tabular financial reports) ──
  function generatePdf(docKey) {
    var fn = DOCS[docKey];
    if (!fn) throw new Error('Unknown document: ' + docKey);
    var blocks = fn();
    var title = DOC_TITLES[docKey];

    var doc = DocumentApp.create('OXYGEN_HELP_TEMP_' + docKey + '_' + new Date().getTime());
    try {
      var body = doc.getBody();
      body.clear();
      var numbering = 0;
      blocks.forEach(function(blk) {
        if (blk.t !== 'n') numbering = 0;
        if (blk.t === 'h1') {
          var p1 = body.appendParagraph(blk.x);
          p1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
        } else if (blk.t === 'h2') {
          var p2 = body.appendParagraph(blk.x);
          p2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        } else if (blk.t === 'h3') {
          var p3 = body.appendParagraph(blk.x);
          p3.setHeading(DocumentApp.ParagraphHeading.HEADING3);
        } else if (blk.t === 'p') {
          body.appendParagraph(blk.x).setHeading(DocumentApp.ParagraphHeading.NORMAL);
        } else if (blk.t === 'b') {
          body.appendListItem(blk.x).setGlyphType(DocumentApp.GlyphType.BULLET).setHeading(DocumentApp.ParagraphHeading.NORMAL);
        } else if (blk.t === 'n') {
          numbering++;
          body.appendListItem(numbering + '. ' + blk.x).setGlyphType(DocumentApp.GlyphType.NUMBER).setHeading(DocumentApp.ParagraphHeading.NORMAL);
        } else if (blk.t === 'hr') {
          body.appendHorizontalRule();
        }
      });
      doc.saveAndClose();

      var pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
      var filename = 'OXYGEN_' + title.replace(/[^A-Za-z0-9]+/g, '_') + '_v' + VERSION + '.pdf';
      return { filename: filename, base64: Utilities.base64Encode(pdfBlob.getBytes()) };
    } finally {
      try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch (e) {}
    }
  }

  return {
    getContent: getContent,
    getVersionInfo: getVersionInfo,
    generatePdf: generatePdf
  };
})();

// Bare global wrapper purely so this is selectable in the Apps Script
// editor's Run dropdown — running this is what triggers Google's
// Docs/Drive permission prompt the FIRST time, since this is the only
// function in the whole project that uses DocumentApp.
function testHelpPdfAuthorization() {
  var result = HelpService.generatePdf('about');
  Logger.log('Generated: ' + result.filename + ' (' + result.base64.length + ' base64 chars)');
  return result.filename;
}

