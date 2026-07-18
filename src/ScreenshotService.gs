// ═══════════════════════════════════════════════════════════════
// ScreenshotService.gs — Payment Confirmation Screenshots.
// Residents WhatsApp payment confirmations; those images go into one
// Drive folder. This module OCRs each new image (Google Drive's own
// text recognition — no external service), extracts the payment
// fields, stores them for review, and suggests matches against
// UNMATCHED bank-ledger transactions using the strongest signal
// available: the 12-digit UPI reference (UTR) that appears in BOTH
// the screenshot and the bank narration. Linking posts through the
// ledger's existing manual-match pipeline — same rules, same audit
// trail as any other match.
// ═══════════════════════════════════════════════════════════════
var ScreenshotService = (function() {
  var SHEET = 'PaymentScreenshots';
  var HEADERS = ['shot_id','file_id','file_name','file_url','unit_hint','date_text','narration','payer_name','bank','amount','notes','utr','matched_txn_id','status','ocr_text','created_at','updated_at'];
  var C = { ID:0, FILE_ID:1, FILE_NAME:2, FILE_URL:3, UNIT:4, DATE:5, NARR:6, PAYER:7, BANK:8, AMOUNT:9, NOTES:10, UTR:11, TXN:12, STATUS:13, OCR:14, CREATED:15, UPDATED:16 };
  var MAX_PER_SCAN = 12; // Apps Script execution-time safety: OCR is ~2-4s per image

  var BANKS = ['Federal Bank','State Bank of India','SBI','ICICI','HDFC','Axis','Canara','Indian Overseas Bank','IOB','Kotak','Union Bank','Punjab National Bank','PNB','Bank of Baroda','South Indian Bank','CSB','Dhanlaxmi','IDBI','IDFC','Yes Bank','IndusInd','RBL','Karur Vysya','Karnataka Bank'];

  function ensureSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET);
    if (!sh) {
      sh = ss.insertSheet(SHEET);
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  }

  function _folderIdFromUrl(url) {
    var m = String(url || '').match(/folders\/([A-Za-z0-9_-]+)/);
    return m ? m[1] : '';
  }

  // ── OCR one Drive image: copy it as a Google Doc (Drive runs text
  //    recognition during that conversion), export the doc as plain
  //    text, delete the temp doc. Uses Drive's REST API with the
  //    script's own OAuth token — no extra services to enable.
  //    Drive throttles rapid-fire conversion requests ("User rate limit
  //    exceeded", 403/429) — those are retried automatically with
  //    increasing pauses rather than surfacing as failures. ──
  function _ocrImage(fileId) {
    var token = ScriptApp.getOAuthToken();
    var copyResp;
    for (var attempt = 0; attempt < 4; attempt++) {
      if (attempt > 0) Utilities.sleep(1500 * Math.pow(2, attempt - 1)); // 1.5s, 3s, 6s
      copyResp = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/copy', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ mimeType: 'application/vnd.google-apps.document', name: 'OCR_TEMP_' + fileId }),
        muteHttpExceptions: true
      });
      var code = copyResp.getResponseCode();
      if (code < 300) break; // success
      var retryable = code === 403 || code === 429 || code >= 500;
      if (!retryable || attempt === 3) throw new Error('OCR conversion failed (' + code + '): ' + copyResp.getContentText().slice(0, 180));
    }
    var docId = JSON.parse(copyResp.getContentText()).id;
    var text = '';
    try {
      var expResp = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + docId + '/export?mimeType=text%2Fplain', {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (expResp.getResponseCode() < 300) text = expResp.getContentText();
    } finally {
      try { DriveApp.getFileById(docId).setTrashed(true); } catch (e) {}
    }
    return text;
  }

  // ── Field extraction — tuned against the association's real
  //    screenshots (GPay, PhonePe, Paytm, bank apps). Every extracted
  //    value is editable on the page afterwards; extraction assists,
  //    the reviewer confirms. ──
  function _parse(text, fileName) {
    var t = String(text || '');
    var lines = t.split(/\r?\n/).map(function(l) { return l.trim(); }).filter(Boolean);
    var joined = lines.join(' \n ');
    var out = { unit_hint: '', date_text: '', narration: '', payer_name: '', bank: '', amount: '', notes: '', utr: '' };

    // Unit from the FILENAME (e.g. "A101.jpeg", "B506 payment.png") —
    // the screenshot itself rarely says which flat paid; the person
    // uploading knows, and encodes it in the name.
    var um = String(fileName || '').match(/^([AB]\d{3})\b/i);
    if (um) out.unit_hint = um[1].toUpperCase();

    // UTR / UPI reference — the gold signal. 12 consecutive digits,
    // usually labelled "UPI transaction ID", "UTR", "UPI Ref No",
    // "Reference number". Try labelled forms first, then any bare
    // 12-digit run.
    var utrM = joined.match(/(?:UPI\s*transaction\s*ID|UTR(?:\s*(?:No|Number))?|UPI\s*Ref(?:erence)?\s*(?:No|Number)?|Ref(?:erence)?\s*(?:No|Number))\s*[:.\-]?\s*(\d{10,16})/i)
            || joined.match(/\b(\d{12})\b/);
    if (utrM) out.utr = utrM[1];

    // Amount — ₹ or Rs, take the LARGEST such value on the screen
    // (screens can also show balance snippets or fees; the payment
    // amount is virtually always the biggest highlighted figure).
    var amtRe = /(?:₹|Rs\.?|INR)\s*([\d,]+(?:\.\d{1,2})?)/gi;
    var best = 0; var am;
    while ((am = amtRe.exec(joined)) !== null) {
      var v = Number(am[1].replace(/,/g, ''));
      if (v > best) best = v;
    }
    if (best > 0) out.amount = best;

    // Date — "14 Jul 2026", "14 July 2026, 10:42 AM", "14-07-2026",
    // "14/07/26". Stored as text exactly as read; editable.
    var dM = joined.match(/\b(\d{1,2}\s+[A-Za-z]{3,9},?\s+\d{4})\b/)
          || joined.match(/\b(\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4})\b/)
          || joined.match(/\b([A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})\b/);
    if (dM) out.date_text = dM[1];

    // Bank — first known bank name found.
    for (var b = 0; b < BANKS.length; b++) {
      var re = new RegExp('\\b' + BANKS[b].replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'i');
      if (re.test(joined)) { out.bank = BANKS[b]; break; }
    }

    // Payer — the line after "From", or before "Paid via"; falls back
    // to a "<NAME>" line in caps near the top.
    for (var i = 0; i < lines.length; i++) {
      if (/^from\b/i.test(lines[i])) {
        var same = lines[i].replace(/^from\s*[:\-]?\s*/i, '');
        out.payer_name = same || (lines[i + 1] || '');
        break;
      }
    }
    if (!out.payer_name) {
      var pm = joined.match(/(?:paid\s+by|sender|debited\s+from\s+account\s+of)\s*[:\-]?\s*([A-Za-z .]{3,40})/i);
      if (pm) out.payer_name = pm[1].trim();
    }
    out.payer_name = String(out.payer_name || '').replace(/\s{2,}/g, ' ').trim();

    // Narration — a compact human line: the "Paid to …" line if present,
    // else the first line mentioning payment/success.
    for (var j = 0; j < lines.length; j++) {
      if (/paid\s+to|payment\s+to|transferred\s+to/i.test(lines[j])) { out.narration = lines[j]; break; }
    }
    if (!out.narration) {
      for (var k = 0; k < lines.length; k++) {
        if (/payment|success|completed|transferred/i.test(lines[k])) { out.narration = lines[k]; break; }
      }
    }
    return out;
  }

  // ── Scan the configured folder: OCR + parse every image not already
  //    in the sheet. Batched for execution-time safety. ──
  function scanFolder() {
    var url = String(SettingsService.get('screenshots_folder_url') || '').trim();
    if (!url) throw new Error('The screenshots folder URL is empty — set it in Settings → 📸 Payment Screenshots first.');
    var folderId = _folderIdFromUrl(url);
    if (!folderId) throw new Error('That does not look like a Drive folder URL (it must contain /folders/<id>).');

    var folder;
    try { folder = DriveApp.getFolderById(folderId); }
    catch (e) { throw new Error('Could not open the screenshots folder — check the URL and that this account has access.'); }

    var sh = ensureSheet();
    var known = {};
    Database.getAll(SHEET).forEach(function(r) { known[String(r[C.FILE_ID])] = true; });

    var files = folder.getFiles();
    var processed = 0, skippedKnown = 0, remaining = 0, errors = [];
    var now = new Date().toISOString();
    while (files.hasNext()) {
      var f = files.next();
      var mime = f.getMimeType();
      if (String(mime).indexOf('image/') !== 0) continue;
      if (known[f.getId()]) { skippedKnown++; continue; }
      if (processed >= MAX_PER_SCAN) { remaining++; continue; }
      try {
        if (processed > 0) Utilities.sleep(700); // pacing — Drive throttles rapid-fire conversions
        var text = _ocrImage(f.getId());
        var p = _parse(text, f.getName());
        sh.appendRow([Database.generateId('SS'), f.getId(), f.getName(), f.getUrl(),
                      p.unit_hint, p.date_text, p.narration, p.payer_name, p.bank,
                      p.amount, '', "'" + p.utr, '', 'New', text.slice(0, 4000), now, now]);
        processed++;
      } catch (e2) {
        errors.push(f.getName() + ': ' + e2.message);
      }
    }
    return { processed: processed, alreadyKnown: skippedKnown, remaining: remaining, errors: errors };
  }

  function _toObj(r) {
    return {
      shot_id: String(r[C.ID] || ''), file_id: String(r[C.FILE_ID] || ''),
      file_name: String(r[C.FILE_NAME] || ''), file_url: String(r[C.FILE_URL] || ''),
      unit_hint: String(r[C.UNIT] || ''), date_text: String(r[C.DATE] || ''),
      narration: String(r[C.NARR] || ''), payer_name: String(r[C.PAYER] || ''),
      bank: String(r[C.BANK] || ''), amount: Number(r[C.AMOUNT]) || 0,
      notes: String(r[C.NOTES] || ''), utr: String(r[C.UTR] || '').replace(/^'/, ''),
      matched_txn_id: String(r[C.TXN] || ''), status: String(r[C.STATUS] || 'New')
    };
  }

  // List all screenshots WITH match suggestions computed against both
  // accounts' unmatched transactions: exact UTR-in-narration first
  // (near-certain), then amount+date proximity as weaker candidates.
  function listWithSuggestions() {
    ensureSheet();
    var shots = Database.getAll(SHEET).map(_toObj).filter(function(s) { return s.shot_id; });

    // Index unmatched bank txns by the digits found in their narration.
    function collect(svc, account) {
      var out = [];
      try {
        svc.getAllTransactions(null).forEach(function(t) {
          if (Number(t.credit || 0) <= 0) return;
          out.push({ txn_id: t.txn_id, account: account, date: t.date,
                     narration: String(t.narration || ''), amount: Number(t.credit || 0),
                     matched: !!(t.match_unit || t.match_id), payment_type: t.payment_type || '' });
        });
      } catch (e) {}
      return out;
    }
    var txns = collect(BankService, '1').concat(collect(Bank2Service, '2'));

    var dupUtr = {};
    shots.forEach(function(s) { if (s.utr) dupUtr[s.utr] = (dupUtr[s.utr] || 0) + 1; });

    shots.forEach(function(s) {
      s.duplicate_utr = !!(s.utr && dupUtr[s.utr] > 1);
      s.suggestions = [];
      if (s.matched_txn_id) return;
      if (s.utr) {
        txns.forEach(function(t) {
          if (t.narration.indexOf(s.utr) > -1) {
            s.suggestions.push({ txn_id: t.txn_id, account: t.account, date: t.date, narration: t.narration,
                                 amount: t.amount, matched: t.matched, confidence: 'exact', payment_type: t.payment_type });
          }
        });
      }
      if (!s.suggestions.length && s.amount > 0) {
        txns.forEach(function(t) {
          if (t.matched) return;
          if (Math.abs(t.amount - s.amount) < 0.01) {
            s.suggestions.push({ txn_id: t.txn_id, account: t.account, date: t.date, narration: t.narration,
                                 amount: t.amount, matched: false, confidence: 'amount', payment_type: t.payment_type });
          }
        });
        if (s.suggestions.length > 6) s.suggestions = s.suggestions.slice(0, 6);
      }
    });

    // Newest first by sheet order (append order = scan order).
    shots.reverse();
    return { shots: shots, count: shots.length,
             folderUrl: String(SettingsService.get('screenshots_folder_url') || '').trim() };
  }

  function updateShot(id, d) {
    var found = Database.findByColumn(SHEET, C.ID, id);
    if (!found) throw new Error('Screenshot record not found: ' + id);
    var r = found.data;
    if (d.unit_hint !== undefined)  r[C.UNIT]   = String(d.unit_hint).toUpperCase();
    if (d.date_text !== undefined)  r[C.DATE]   = String(d.date_text);
    if (d.narration !== undefined)  r[C.NARR]   = String(d.narration);
    if (d.payer_name !== undefined) r[C.PAYER]  = String(d.payer_name);
    if (d.bank !== undefined)       r[C.BANK]   = String(d.bank);
    if (d.amount !== undefined)     r[C.AMOUNT] = Number(d.amount) || 0;
    if (d.notes !== undefined)      r[C.NOTES]  = String(d.notes);
    if (d.utr !== undefined)        r[C.UTR]    = "'" + String(d.utr).replace(/^'/, '');
    if (d.status !== undefined)     r[C.STATUS] = String(d.status);
    r[C.UPDATED] = new Date().toISOString();
    Database.updateRow(SHEET, found.rowIndex, r);
    return { success: true };
  }

  function deleteShot(id) {
    var found = Database.findByColumn(SHEET, C.ID, id);
    if (!found) throw new Error('Screenshot record not found: ' + id);
    ensureSheet().deleteRow(found.rowIndex);
    return { success: true };
  }

  // Link a screenshot to a ledger transaction AND match the unit through
  // the ledger's normal manual-match pipeline (same posting, history
  // rules, and audit trail as matching from the Transaction Ledger page).
  function linkAndMatch(d) {
    if (!d || !d.shot_id || !d.txn_id) throw new Error('shot_id and txn_id are required.');
    var svc = String(d.account) === '2' ? Bank2Service : BankService;
    if (d.unit_id) {
      svc.manualMatch(d.txn_id, 'screenshot', '', String(d.payer_name || ''), String(d.unit_id).toUpperCase(),
                      String(d.payment_type || ''), d.confirmed_historical === true);
    }
    updateShot(d.shot_id, { status: 'Matched' });
    var found = Database.findByColumn(SHEET, C.ID, d.shot_id);
    var r = found.data;
    r[C.TXN] = String(d.txn_id);
    r[C.UPDATED] = new Date().toISOString();
    Database.updateRow(SHEET, found.rowIndex, r);
    return { success: true };
  }

  return {
    ensureSheet: ensureSheet,
    scanFolder: scanFolder,
    listWithSuggestions: listWithSuggestions,
    updateShot: updateShot,
    deleteShot: deleteShot,
    linkAndMatch: linkAndMatch
  };
})();
