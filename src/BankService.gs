// ============================================================
// BankService.gs — Bank Statement Import & Matching
// Parses uploaded XLS/CSV bank statements and matches
// transactions to Owners/Tenants by name from Narration field.
// ============================================================

// Factory: builds a complete, self-contained bank-statement service bound
// to one specific sheet. The app has two real bank accounts (the main IOB
// account and the older IOB LPG account), each with its own transaction
// ledger sheet — same format, same features, different data. Everything
// inside is identical for both; only the sheet differs.
function _makeBankService(SHEET, ACCOUNT) {

  // Sortable numeric key (yyyymmdd) from any of the date formats this
  // ledger stores. Shared by getAllTransactions (display order) AND
  // reconcile() (month-by-month running balance) — both MUST treat "the
  // statement's real chronological order" identically, since reconcile's
  // opening/closing math silently breaks if it ever trusts raw sheet row
  // order instead of actual dates (e.g. after any out-of-order append).
  function _dateKey(d) {
    if (!d) return 0;
    // Sheets sometimes auto-converts a date-LOOKING string (e.g. one typed
    // as "2024-01-09" instead of the rest of the column's "01-Jan-2024")
    // into an actual Date-type cell — silently, on just that one row.
    // String(dateObject) then produces "Tue Jan 09 2024 00:00:00 GMT..."
    // which matches none of the text patterns below, so the row would be
    // silently skipped from every calculation. Reading the Date object's
    // own calendar fields directly sidesteps that entirely.
    if (d instanceof Date) {
      return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
    }
    var t = String(d).trim(), low = t.toLowerCase();
    var mn = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
    var m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return (+m[1])*10000 + (+m[2])*100 + (+m[3]);
    m = low.match(/^(\d{1,2})[-\/ ]([a-z]{3})[a-z]*[-\/ ](\d{2,4})/);
    if (m && mn[m[2]]) { var y=+m[3]; if (y<100) y+=2000; return y*10000 + mn[m[2]]*100 + (+m[1]); }
    m = t.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/);
    if (m) { var yy=+m[3]; if (yy<100) yy+=2000; var a2=+m[1], b2=+m[2], mo=b2, dd=a2; if (b2>12&&a2<=12){mo=a2;dd=b2;} return yy*10000+mo*100+dd; }
    return 0;
  }

  // Column indices in BankStatements sheet
  var C = {
    TXN_ID:        0,  // unique hash of date+narration+amount
    DATE:          1,
    VALUE_DATE:    2,
    CHQ_NO:        3,
    NARRATION:     4,
    COD:           5,
    DEBIT:         6,
    CREDIT:        7,
    BALANCE:       8,
    PARSED_NAME:   9,  // name extracted from narration
    PARSED_REMARK: 10, // remarks/flat hint from narration
    TXN_TYPE:      11, // UPI / IMPS / NEFT / OTHER
    MATCH_TYPE:    12, // 'owner' / 'tenant' / 'unmatched'
    MATCH_ID:      13, // owner_id or tenant_id
    MATCH_NAME:    14, // matched person's name
    MATCH_UNIT:    15, // matched unit_id
    PAYMENT_TYPE:  16, // Maintenance / Waste Management / LPG / Unknown (credit side)
    IMPORTED_AT:   17,
    DIRECTION:     18, // 'credit' (money received) / 'debit' (money paid out)
    EXPENSE_CAT:   19, // for debits: which association expense (Care Taker, KWA, …)
    NOTE:          20  // free-text note on any line item (purpose of payment)
  };

  // ── Sheet initialisation ─────────────────────────────────────

  function ensureSheet() {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET);
    var headers = [
      'txn_id','date','value_date','chq_no','narration','cod',
      'debit','credit','balance','parsed_name','parsed_remark',
      'txn_type','match_type','match_id','match_name','match_unit',
      'payment_type','imported_at','direction','expense_category','note'
    ];
    if (!sheet) {
      sheet = ss.insertSheet(SHEET);
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length)
        .setFontWeight('bold')
        .setBackground('#0f2744')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    } else {
      // Repair a missing/blank header row (never over data)
      if (sheet.getMaxColumns() < headers.length) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
      }
      var r1c1 = String(sheet.getRange(1, 1).getValue() || '');
      if (r1c1 === '') {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1,1,1,headers.length)
          .setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
        if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
      }
    }
    // Force DATE and VALUE_DATE columns to plain text, permanently. A
    // single row typed/imported as "2024-01-09" (instead of the rest of
    // the column's "01-Jan-2024") gets silently auto-converted by Sheets
    // into a real Date-type cell — invisible in the UI, but it broke
    // reconciliation by causing that one row to be silently excluded from
    // every date-based calculation. Plain text stops this permanently.
    sheet.getRange(2, C.DATE + 1, 20000, 1).setNumberFormat('@');
    sheet.getRange(2, C.VALUE_DATE + 1, 20000, 1).setNumberFormat('@');
    return sheet;
  }

  // ── CSV Import (called from frontend with raw CSV string) ────

  function importCsv(csvString) {
    ensureSheet();

    var rows    = _parseCsv(csvString);
    if (rows.length < 2) return { imported: 0, duplicates: 0, errors: ['No data rows found.'] };

    // Find the header row — look for 'Date' in first column
    var headerRowIdx = -1;
    for (var i = 0; i < Math.min(10, rows.length); i++) {
      if (rows[i][0] && String(rows[i][0]).toLowerCase().trim() === 'date') {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) headerRowIdx = 0;

    var headers   = rows[headerRowIdx].map(function(h) { return String(h).toLowerCase().trim(); });
    var colDate   = headers.indexOf('date');
    var colValDate= headers.indexOf('value date');
    var colChq    = headers.indexOf('chq no');
    var colNarr   = headers.indexOf('narration');
    var colCod    = headers.indexOf('cod');
    var colDebit  = headers.indexOf('debit');
    var colCredit = headers.indexOf('credit');
    var colBal    = headers.indexOf('balance');

    // Fallback by position if headers not found
    if (colDate   < 0) colDate   = 0;
    if (colValDate< 0) colValDate= 1;
    if (colChq    < 0) colChq    = 2;
    if (colNarr   < 0) colNarr   = 3;
    if (colCod    < 0) colCod    = 4;
    if (colDebit  < 0) colDebit  = 5;
    if (colCredit < 0) colCredit = 6;
    if (colBal    < 0) colBal    = 7;

    // Load existing txn_ids to avoid duplicates
    var existing   = Database.getAll(SHEET);
    var existingIds = {};
    existing.forEach(function(r) { if (r[C.TXN_ID]) existingIds[r[C.TXN_ID]] = true; });

    // Load owners and tenants for matching
    var owners  = OwnersService.getAllOwners();
    var tenants = TenantsService.getAllTenants();

    // Detect day-first vs month-first ONCE across the whole file
    var allDates = [];
    for (var d = headerRowIdx + 1; d < rows.length; d++) {
      if (rows[d] && rows[d][colDate]) allDates.push(String(rows[d][colDate]));
    }
    var dayFirst = _detectDayFirst(allDates);

    var imported   = 0;
    var duplicates = 0;
    var skippedBad = 0;
    var errors     = [];
    var now        = new Date().toISOString();
    var toPost     = [];   // matched credits to post as payments
    var newRows    = [];   // batched sheet rows
    var fees       = _getFees();

    for (var i = headerRowIdx + 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row || row.length < 4) continue;

      var dateStr  = String(row[colDate]   || '').trim();
      var narr     = String(row[colNarr]   || '').replace(/[\r\n]+/g, ' ').trim();
      var debit    = String(row[colDebit]  || '').trim();
      var credit   = String(row[colCredit] || '').trim();
      var balance  = String(row[colBal]    || '').trim();

      // Only requirement: a date-looking cell and at least one real amount
      // (credit or debit). Narration is deliberately NOT required — real
      // bank statements occasionally have a blank description on an
      // otherwise completely valid, balance-affecting transaction (e.g. a
      // bank-generated credit with no narration at all). Rejecting those
      // meant a real transaction could silently vanish from the ledger
      // entirely, breaking reconciliation with no way to see why.
      if (!dateStr) { skippedBad++; continue; }
      if (!/^\d{1,2}[-\/][A-Za-z0-9]{1,3}[-\/]\d{2,4}/.test(dateStr) &&
          !/^\d{4}-\d{2}-\d{2}/.test(dateStr)) { skippedBad++; continue; }
      var hasAmount = (credit !== '' && !isNaN(Number(String(credit).replace(/[₹,\s]/g, '')))) ||
                      (debit  !== '' && !isNaN(Number(String(debit).replace(/[₹,\s]/g, ''))));
      if (!hasAmount) { skippedBad++; continue; }

      // Deterministic ID — includes the running balance so identical-looking
      // transactions on the same day are all kept
      var txnId = _makeId(dateStr, narr, credit || debit, balance);
      if (existingIds[txnId]) { duplicates++; continue; }

      // Matching only applies to credits with a parsable name; other rows
      // (charges, debits, interest) are imported as-is and stay unmatched
      var credNum  = Number(String(credit).replace(/[₹,\s]/g, ''));
      var isCredit = credit !== '' && !isNaN(credNum) && credNum > 0;
      var debNum   = Number(String(debit).replace(/[₹,\s]/g, ''));
      var isDebit  = debit !== '' && !isNaN(debNum) && debNum > 0;

      // Direction: money received by the association vs paid out by it
      var direction = isCredit ? 'credit' : (isDebit ? 'debit' : '');

      var parsed  = _parseNarration(narr);
      var match   = { type: 'unmatched', id: '', name: '', unit: '' };
      var payType = '';
      var expenseCat = '';
      if (isCredit && !_isSystemRow(narr, debit, credit)) {
        match   = _matchPerson(parsed.name, parsed.remark, owners, tenants);
        payType = _guessPaymentType(parsed.remark, credNum, fees);
      }
      // For debits, try to classify which association expense this is
      if (isDebit) {
        expenseCat = _guessExpenseCategory(narr, debNum);
      }

      newRows.push([
        txnId,
        "'" + dateStr,                                    // text — never re-parsed by Sheets
        "'" + String(row[colValDate] || '').trim(),
        String(row[colChq]     || '').trim(),
        narr,
        String(row[colCod]     || '').trim(),
        debit  || '',
        credit || '',
        balance,
        parsed.name,
        parsed.remark,
        parsed.type,
        match.type,
        match.id,
        match.name,
        match.unit,
        payType,
        now,
        direction,
        expenseCat,
        ''
      ]);

      existingIds[txnId] = true;
      imported++;

      // Queue matched credits for automatic payment posting.
      // The month is resolved HERE, unambiguously, from the statement date.
      if (match.unit && isCredit) {
        var monthKey = _txnMonthKey(dateStr, dayFirst);
        toPost.push({
          txn_id:       txnId,
          unit_id:      match.unit,
          tenant_id:    match.type === 'tenant' ? match.id : '',
          payment_type: payType,
          amount:       credNum,
          dateStr:      monthKey || dateStr
        });
      }
    }

    // ── ONE batched write (plus header repair) ──────────────────
    var sheet = Database.getSheet(SHEET);
    var HEADERS = [
      'txn_id','date','value_date','chq_no','narration','cod',
      'debit','credit','balance','parsed_name','parsed_remark',
      'txn_type','match_type','match_id','match_name','match_unit',
      'payment_type','imported_at','direction','expense_category','note'
    ];
    if (sheet.getMaxColumns() < HEADERS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
    }
    // Repair the header row if blank or already a header (never over data)
    var r1c1 = String(sheet.getRange(1, 1).getValue() || '');
    if (r1c1 === '' || r1c1 === 'txn_id') {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
    }
    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, HEADERS.length).setValues(newRows);
      SpreadsheetApp.flush();
    }

    // ── Automatic bank -> payments pipeline (batched, idempotent) ──
    var paymentsCreated = 0;
    if (toPost.length > 0) {
      try {
        var alloc = PaymentsService.allocateFromBankBatch(toPost, fees);
        paymentsCreated = alloc.created;
        // NOTE: the Owners/Tenants sheet Jan–Dec columns are reference data
        // from the imported Excel masters and are intentionally NOT modified.
        // All payment display reads from the central Payments ledger.
      } catch (allocErr) {
        errors.push('Payment posting: ' + allocErr.message);
      }
    }

    return { imported: imported, duplicates: duplicates, skippedBad: skippedBad,
             errors: errors, paymentsCreated: paymentsCreated,
             total: rows.length - headerRowIdx - 1 };
  }

  // ── Narration Parser ─────────────────────────────────────────

  function _parseNarration(narr) {
    var name   = '';
    var remark = '';
    var type   = 'OTHER';

    if (narr.indexOf('UPI/') === 0 || narr.indexOf('UPI/') > -1) {
      // UPI/TxnID/CR/SENDER NAME/Bank/Remarks
      type = 'UPI';
      var parts = narr.split('/');
      if (parts.length >= 4) name   = parts[3].trim();
      if (parts.length >= 6) remark = parts[5].trim();
    }
    else if (narr.indexOf('IMPS/') === 0) {
      // IMPS/TxnID/TRUNCATED_NAME/Bank/AccountXXXX/-
      type = 'IMPS';
      var parts = narr.split('/');
      if (parts.length >= 3) name = parts[2].trim();
    }
    else if (narr.indexOf('IB/IMPS/') === 0) {
      // IB/IMPS/TxnID/NAME/Bank
      type = 'IMPS';
      var parts = narr.split('/');
      if (parts.length >= 4) name = parts[3].trim();
    }
    else if (/\bIMPS\b/i.test(narr)) {
      // A DIFFERENT bank's IMPS format entirely — space-delimited, not
      // slash-delimited: "TRTR 400309480040 IMPS SHIBUDS". The name is
      // whatever comes after the IMPS keyword; everything before it
      // (transfer code, reference number) is not part of the name.
      type = 'IMPS';
      var m = narr.match(/\bIMPS\b\s*(.*)$/i);
      if (m && m[1]) name = m[1].replace(/^[\/\-_\s]+/, '').trim();
    }
    else if (narr.indexOf('NEFT-') === 0) {
      // NEFT-Bank-RefNo-SENDER NAME-remarks
      type = 'NEFT';
      var parts = narr.split('-');
      // Name is usually in position 3 (after bank, refno)
      if (parts.length >= 4) name = parts[3].trim();
      if (parts.length >= 5) remark = parts.slice(4).join('-').trim();
    }
    else if (narr.indexOf('MB/') === 0) {
      type = 'MB';
      // Mobile banking — name not available
      name = '';
    }
    else {
      // Plain narration (sometimes just a name)
      type   = 'OTHER';
      name   = narr.replace(/[\/\-_]+/g, ' ').trim();
    }

    // Clean up name — remove extra spaces, normalize
    name   = name.replace(/\s+/g, ' ').trim();
    remark = remark.replace(/\s+/g, ' ').trim();

    return { name: name, remark: remark, type: type };
  }

  // ── Person Matcher ───────────────────────────────────────────

  function _matchPerson(name, remark, owners, tenants) {
    if (!name) return { type: 'unmatched', id: '', name: '', unit: '' };

    var nameU    = name.toUpperCase();
    var remarkU  = (remark || '').toUpperCase();
    var bestScore = 0;
    var bestMatch = null;
    var bestType  = '';

    // Check owners — score against the primary name AND every
    // "Name on Bank Statement" alias imported from the master sheet
    owners.forEach(function(o) {
      var score = _nameScore(nameU, o.name.toUpperCase());

      // Bank statement aliases are the exact strings that appear in
      // narrations, so they are the strongest possible signal
      if (o.bank_names) {
        var aliases = o.bank_names.split(',');
        for (var a = 0; a < aliases.length; a++) {
          var alias = aliases[a].trim().toUpperCase();
          if (!alias) continue;
          if (alias === nameU) { score = Math.max(score, 100); break; }
          var aScore = _nameScore(nameU, alias);
          if (aScore > score) score = aScore;
        }
      }

      // Boost if remark mentions the unit
      if (remarkU && o.unit_id && remarkU.indexOf(o.unit_id.toUpperCase()) > -1) score += 30;
      if (score > bestScore) { bestScore = score; bestMatch = o; bestType = 'owner'; }
    });

    // Check tenants — score against name AND every bank statement alias
    tenants.forEach(function(t) {
      var score = _nameScore(nameU, t.name.toUpperCase());
      if (t.bank_name) {
        var tAliases = t.bank_name.split(',');
        for (var a = 0; a < tAliases.length; a++) {
          var tAlias = tAliases[a].trim().toUpperCase();
          if (!tAlias) continue;
          if (tAlias === nameU) { score = Math.max(score, 100); break; }
          var tScore = _nameScore(nameU, tAlias);
          if (tScore > score) score = tScore;
        }
      }
      if (remarkU && t.unit_id && remarkU.indexOf(t.unit_id.toUpperCase()) > -1) score += 30;
      if (score > bestScore) { bestScore = score; bestMatch = t; bestType = 'tenant'; }
    });

    // Accept match only if score is reasonably high
    if (bestScore >= 40 && bestMatch) {
      return {
        type:  bestType,
        id:    bestMatch.owner_id || bestMatch.tenant_id || '',
        name:  bestMatch.name,
        unit:  bestMatch.unit_id || ''
      };
    }

    return { type: 'unmatched', id: '', name: name, unit: '' };
  }

  // Scoring: how well does txnName match personName?
  function _nameScore(txnName, personName) {
    if (!txnName || !personName) return 0;
    if (txnName === personName) return 100;

    // Full words (length > 2) score normally. Short tokens — single
    // initials like "B" in "JAYAKRISHNAN B" — used to be discarded
    // entirely, which meant two residents differing only by initial
    // (e.g. "JAYAKRISHNAN B" vs "JAYAKRISHNAN K") could score identically
    // and the real match could lose to a coincidental better full-word
    // match elsewhere. Short tokens now contribute a smaller amount
    // instead of nothing, so an initial that matches (or doesn't) can
    // still tip a close decision the right way.
    var txnWords    = txnName.split(/\s+/).filter(function(w) { return w.length > 0; });
    var personWords = personName.split(/\s+/).filter(function(w) { return w.length > 0; });

    var matchCount  = 0;
    txnWords.forEach(function(tw) {
      personWords.forEach(function(pw) {
        var short = tw.length <= 2 || pw.length <= 2;
        if (tw === pw) { matchCount += short ? 0.5 : 2; return; }
        if (!short) {
          if (tw.length >= 5 && pw.indexOf(tw) === 0) matchCount += 1.5;
          if (pw.length >= 5 && tw.indexOf(pw) === 0) matchCount += 1.5;
        }
      });
    });

    if (matchCount === 0) return 0;
    var denominator = Math.max(txnWords.length, personWords.length);
    return Math.round((matchCount / denominator) * 60);
  }

  // Guess payment type from remark and amount
  function _guessPaymentType(remark, amount, fees) {
    var r = ' ' + (remark || '').toUpperCase() + ' ';
    // Keyword checks first — most reliable
    if (r.indexOf('CAUTION') > -1 || r.indexOf('DEPOSIT') > -1) return 'Caution Deposit';
    if (r.indexOf('PARTY') > -1 || r.indexOf('HALL') > -1 || r.indexOf('RENTAL') > -1) return 'Party Hall Rental';
    if (r.indexOf('MAINT') > -1 || /[^A-Z]MC[^A-Z]/.test(r) || /[^A-Z]MMC[^A-Z]/.test(r)) return 'Maintenance';
    if (r.indexOf('WASTE') > -1 || /[^A-Z]WM[^A-Z]/.test(r) || /[^A-Z]WMF[^A-Z]/.test(r)) return 'Waste Management';
    if (r.indexOf('GAS') > -1 || r.indexOf('LPG') > -1) return 'LPG';
    // Amount-based: current fee AND historical amounts, with multiples —
    // but only up to 12 months' worth. Larger sums (corpus, lump payments)
    // are NOT auto-classified as monthly fees.
    var a = Number(amount || 0);
    // Caution Deposit is a fixed one-time amount — match EXACTLY, no multiples.
    var cd = fees && fees['Caution Deposit'];
    if (a > 0 && cd && cd.length && cd.indexOf(a) > -1) return 'Caution Deposit';
    if (a > 0 && _matchesAnyFee(a, fees && fees['Maintenance']))      return 'Maintenance';
    if (a > 0 && _matchesAnyFee(a, fees && fees['Waste Management'])) return 'Waste Management';
    return a > 0 ? 'LPG' : 'Unknown';
  }

  function _matchesAnyFee(amount, feeList) {
    var list = _feeArray(feeList);
    for (var i = 0; i < list.length; i++) {
      if (list[i] > 0 && amount % list[i] === 0 && amount / list[i] <= 12) return true;
    }
    return false;
  }

  function _feeArray(v) {
    if (!v) return [];
    if (Object.prototype.toString.call(v) === '[object Array]') return v;
    return [Number(v)];
  }

  // Read the current + historical fee amounts from the Settings page.
  // Each type maps to a LIST: current amount first, then older amounts.
  function _getFees() {
    var fees = { 'Maintenance': [2000, 1500], 'Waste Management': [170] };
    try {
      var s = SettingsService.getAll();
      // Maintenance & Waste amounts come from the Fee Schedule now (the single
      // source of truth) — every distinct amount ever in force is recognised
      // in statements, current amount first.
      var mf = SettingsService.feeAmountsForType('Maintenance');
      if (mf.length) fees['Maintenance'] = mf;

      var wf = SettingsService.feeAmountsForType('Waste Management');
      if (wf.length) fees['Waste Management'] = wf;

      if (s.fee_lpg_default && !isNaN(Number(s.fee_lpg_default))) fees['LPG'] = [Number(s.fee_lpg_default)];

      // Caution Deposit — current amount (default ₹1600) plus any older/
      // alternate amounts from settings, all used for exact-match detection
      var cd = (s.fee_caution_deposit && !isNaN(Number(s.fee_caution_deposit))) ? Number(s.fee_caution_deposit) : 1600;
      fees['Caution Deposit'] = [cd].concat(_parseFeeList(s.fee_caution_deposit_history));
    } catch (e) {}
    return fees;
  }

  // '1500, 1200' -> [1500, 1200]
  function _parseFeeList(str) {
    var out = [];
    if (!str) return out;
    var parts = String(str).split(',');
    for (var i = 0; i < parts.length; i++) {
      var n = Number(String(parts[i]).replace(/[₹\s]/g, ''));
      if (!isNaN(n) && n > 0) out.push(n);
    }
    return out;
  }

  // For a debit row, try to work out which association expense it is by
  // looking for the category name (or obvious keywords) inside the narration.
  // Returns '' when nothing matches — the user can set it manually in the UI.
  function _guessExpenseCategory(narr, amount) {
    var text = String(narr || '').toLowerCase();

    // Your own patterns (set on the Settings page) are the most specific,
    // deliberate signal — either a recurring vendor's narration text, a
    // fixed recurring amount (e.g. a cylinder delivery that's always
    // exactly ₹9,120, even when the reference number is random every
    // time), or both together — so they're checked before anything
    // automatic. Checked even when there's no narration text at all,
    // since an amount-only pattern doesn't need any.
    try {
      var explicitCat = ExpensePatternService.matchCategory(narr, amount);
      if (explicitCat) return explicitCat;
    } catch (e) {}

    if (!text) return '';
    var cats = [];
    try { cats = CategoriesService.names('debit'); } catch (e) { cats = []; }

    // A few narration keywords that map to a category name, to catch cases
    // where the bank text won't literally contain the category label.
    var HINTS = {
      'Care Taker':        ['care taker', 'caretaker'],
      'Sewage':            ['sewage', 'sewerage'],
      'Insurance':         ['insurance', 'policy', 'lic '],
      'KWA':               ['kwa', 'water authority', 'kerala water'],
      'Lift AMC':          ['lift', 'elevator'],
      'DG/Pump':           ['dg ', 'diesel', 'pump', 'generator'],
      'Incinerator AMC':   ['incinerator'],
      'Cess/Property Tax': ['cess', 'property tax', 'tax'],
      'Fire Renewal':      ['fire'],
      'House Keeping (HK)': ['house keeping', 'housekeeping', 'hk '],
      'Electrical Inspectorate': ['electrical', 'inspectorate', 'kseb'],
      'IOB Bank Charges':  ['bank charge', 'iob', 'service charge', 'sms charge'],
      'We Care (Salary)':  ['we care', 'salary', 'wecare'],
      'Facility Manager':  ['facility'],
      'LPG Reading':       ['lpg reading'],
      'LPG Inventory':     ['lpg inventory'],
      'Waste Mgmt':        ['waste', 'garbage'],
      'Misc. Purchase':    ['purchase', 'misc']
    };

    // 1) Direct name match against the actual category list
    for (var i = 0; i < cats.length; i++) {
      var nm = String(cats[i] || '').toLowerCase().trim();
      if (nm && text.indexOf(nm) > -1) return cats[i];
    }
    // 2) Keyword hints — only return a category the user actually has
    for (var c = 0; c < cats.length; c++) {
      var hints = HINTS[cats[c]];
      if (!hints) continue;
      for (var h = 0; h < hints.length; h++) {
        if (text.indexOf(hints[h]) > -1) return cats[c];
      }
    }
    return '';
  }

  function _isSystemRow(narr, debit, credit) {
    var n = narr.toUpperCase();
    // Skip outgoing debits (maintenance fund payments, bank charges)
    if (debit && Number(debit) > 0 && (!credit || Number(credit) === 0)) return true;
    if (n.indexOf('CHRGS') > -1) return true;
    if (n.indexOf('INT.PD') > -1) return true;
    if (n.indexOf('INTEREST') > -1 && !credit) return true;
    if (n.indexOf('BBPS') > -1) return true;
    if (n.indexOf('IMPS CHARGES') > -1) return true;
    return false;
  }

  // ── Date disambiguation ─────────────────────────────────────
  // Numeric dates like '01/02/2026' are ambiguous (DD/MM vs MM/DD),
  // and the Excel->CSV conversion can render either. Detect the order
  // once per data set: any first-number > 12 proves day-first; any
  // second-number > 12 proves month-first. A month of statement rows
  // always contains days > 12, so this is reliable.
  function _detectDayFirst(dateStrings) {
    for (var i = 0; i < dateStrings.length; i++) {
      var m = String(dateStrings[i] || '').trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-]\d{2,4}/);
      if (!m) continue;
      var a = Number(m[1]), b = Number(m[2]);
      if (a > 12) return true;   // first number is a day
      if (b > 12) return false;  // second number is a day
    }
    return true; // default: Indian day-first
  }

  // Any date cell (string or Date) -> unambiguous 'YYYY-MM' key ('' if unknown)
  function _txnMonthKey(v, dayFirst) {
    if (v instanceof Date) {
      return v.getFullYear() + '-' + (v.getMonth() + 1 < 10 ? '0' + (v.getMonth() + 1) : String(v.getMonth() + 1));
    }
    var s = String(v || '').trim();
    var MONTHS = { JAN:1,FEB:2,MAR:3,APR:4,MAY:5,JUN:6,JUL:7,AUG:8,SEP:9,OCT:10,NOV:11,DEC:12 };
    var m = s.match(/^(\d{1,2})[-\s]([A-Za-z]{3})[A-Za-z]*[-\s](\d{2,4})/);
    if (m && MONTHS[m[2].toUpperCase()]) return _ymKey(m[3], MONTHS[m[2].toUpperCase()]);
    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (m) {
      var mo = dayFirst ? Number(m[2]) : Number(m[1]);
      if (mo >= 1 && mo <= 12) return _ymKey(m[3], mo);
      // The "month" side exceeds 12 — the other side must be the month
      var alt = dayFirst ? Number(m[1]) : Number(m[2]);
      if (alt >= 1 && alt <= 12) return _ymKey(m[3], alt);
      return '';
    }
    m = s.match(/^(\d{4})-(\d{2})/);
    if (m) return m[1] + '-' + m[2];
    return '';
  }

  function _ymKey(yearStr, mo) {
    var y = Number(yearStr);
    if (y < 100) y += 2000;
    return y + '-' + (mo < 10 ? '0' + mo : String(mo));
  }

  function _makeId(date, narr, amount, balance) {
    // The running balance makes each physical statement row unique, so two
    // genuinely different transactions with identical date/narration/amount
    // (e.g. two ₹2000 UPI credits the same day) are NEVER merged.
    var raw = (date + '|' + narr + '|' + amount + '|' + (balance || '')).replace(/\s/g,'');
    // Simple hash
    var hash = 0;
    for (var i = 0; i < raw.length; i++) {
      hash = ((hash << 5) - hash) + raw.charCodeAt(i);
      hash |= 0;
    }
    return 'TXN_' + Math.abs(hash).toString(16).toUpperCase();
  }

  // ── Queries ──────────────────────────────────────────────────

  /**
   * resetBankData — clean slate for re-importing statements.
   * Clears all BankStatements rows AND removes all bank-sourced
   * payments (notes starting 'BANK:') from the central ledger.
   * Manual payments are kept.
   */
  function resetBankData() {
    // 1. Clear all bank transaction rows
    var sheet = ensureSheet();
    var last  = sheet.getLastRow();
    var txnsCleared = last > 1 ? last - 1 : 0;
    if (last > 1) sheet.getRange(2, 1, last - 1, sheet.getMaxColumns()).clearContent();

    // 2. Remove bank-sourced payments, keep manual ones
    var pSheet = Database.getSheet('Payments');
    var rows   = Database.getAll('Payments');
    var keep   = [];
    for (var i = 0; i < rows.length; i++) {
      while (rows[i].length < 12) rows[i].push('');
      if (String(rows[i][8] || '').indexOf('BANK:') !== 0) keep.push(rows[i]);
    }
    var removed = rows.length - keep.length;
    var pl = pSheet.getLastRow();
    if (pl > 1) pSheet.getRange(2, 1, pl - 1, pSheet.getMaxColumns()).clearContent();
    if (keep.length > 0) pSheet.getRange(2, 1, keep.length, 12).setValues(keep);
    SpreadsheetApp.flush();

    return { txnsCleared: txnsCleared, paymentsRemoved: removed };
  }

  // Lightweight, UNSORTED read for callers that only need to group debits
  // by expense category (e.g. the Caretaker ledger's "Reserve Money"
  // pull) — skips building full transaction objects for every row and,
  // critically, skips the O(n log n) chronological sort entirely, since
  // callers here don't care about order, only totals per month. On a
  // multi-year Bank Statement, getAllTransactions()'s sort is real,
  // avoidable overhead for this specific use case.
  function getDebitsByCategory(category) {
    var rows = Database.getAll(SHEET);
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.EXPENSE_CAT] || '') !== category) continue;
      var debit = Number(String(rows[i][C.DEBIT] || '').replace(/[₹,\s]/g, '')) || 0;
      if (!debit) continue;
      var dateRaw = rows[i][C.DATE];
      var dateStr;
      if (dateRaw instanceof Date) {
        dateStr = dateRaw.getFullYear() + '-' + _pad2(dateRaw.getMonth() + 1) + '-' + _pad2(dateRaw.getDate());
      } else {
        // Bank Statement dates are stored as "DD-Mon-YYYY" plain text
        // throughout this project (deliberately, to prevent Sheets from
        // silently auto-converting them) — this MUST be parsed, not used
        // as-is, or a month lookup like "2024-03" will never match a raw
        // string like "15-Mar-2024" and silently find nothing every time.
        dateStr = parseFlexibleDate(String(dateRaw || '')) || String(dateRaw || '');
      }
      out.push({ date: dateStr, narration: String(rows[i][C.NARRATION] || ''), note: String(rows[i][C.NOTE] || ''), debit: debit });
    }
    return out;
  }

  // Same purpose as getAllTransactions({unit_id}), but correct for split
  // transactions. A plain match_unit filter shows the FULL, original
  // credit amount under whichever unit the bank row happened to match —
  // for a transaction that's since been split across units/months, that
  // full amount is no longer accurate for any single unit, and a unit
  // that received a split portion without being the original bank match
  // (e.g. a second flat sharing one combined payment) wouldn't show up
  // at all. This resolves both: replaces a split transaction's amount
  // with just this unit's own share, drops it entirely if this unit
  // ended up with none of it, and adds an entry for a unit that received
  // a share despite not being the original match.
  // Transactions belonging to a unit, PORTION-aware. A transaction with
  // portions is represented by each unit's own portion (amount + type);
  // one without portions passes through untouched with its normal match.
  // Portions are the only bridge for the "two flats, one payment" case —
  // the second flat gets a synthetic entry carrying the original txn's
  // date/narration and its own share as the amount.
  function getTransactionsForUnitSplitAware(unitId) {
    var unitKey = String(unitId).toUpperCase();

    var portionMap = {}; // txnId -> [portions] (this account only)
    try {
      PortionsService.getPortionsForTxn; // existence probe
      var allP = [];
      var m = PortionsService.getPortionTxnIds(ACCOUNT);
      // fetch portions per txn lazily below via getPortionsForTxn; for a
      // single pass, pull everything once through the summary API:
      // (kept simple — portion volumes are tiny by design: edge cases only)
      for (var tid0 in m) portionMap[tid0] = PortionsService.getPortionsForTxn(tid0);
    } catch (ePor) { portionMap = {}; }

    var direct = getAllTransactions({ unit_id: unitId });
    var out = [];
    var directTxnIds = {};

    direct.forEach(function(t) {
      directTxnIds[t.txn_id] = true;
      var portions = portionMap[t.txn_id];
      if (!portions) { out.push(t); return; } // simple payment — untouched
      var mine = portions.filter(function(p) { return p.unit_id === unitKey; });
      if (!mine.length) return; // this unit has no share of the portioned txn
      mine.forEach(function(p) {
        var copy = {}; for (var k in t) copy[k] = t[k];
        copy.credit = p.amount;
        copy.payment_type = p.payment_type;
        copy.portion = true;
        out.push(copy);
      });
    });

    // Portions pointing at this unit on transactions matched elsewhere
    // (or unmatched) — synthetic entries with the unit's own share.
    var allTxns = getAllTransactions(null);
    var byTxnId = {};
    allTxns.forEach(function(t) { byTxnId[t.txn_id] = t; });
    for (var tid in portionMap) {
      if (directTxnIds[tid]) continue;
      var orig = byTxnId[tid];
      if (!orig) continue;
      portionMap[tid].forEach(function(p) {
        if (p.unit_id !== unitKey) return;
        var synthetic = {}; for (var k2 in orig) synthetic[k2] = orig[k2];
        synthetic.credit = p.amount;
        synthetic.match_unit = unitKey;
        synthetic.payment_type = p.payment_type;
        synthetic.portion = true;
        out.push(synthetic);
      });
    }

    return out.sort(function(a, b) { return _dateKey(a.date) - _dateKey(b.date); });
  }

  function formatRupeesServer(n) { return '\u20b9' + (Math.round(Number(n) * 100) / 100).toLocaleString('en-IN'); }

  function getAllTransactions(filters) {
    var rows = Database.getAll(SHEET);
    var txns = rows.map(function(r) { return _toObj(r); });

    if (filters) {
      if (filters.match_type) txns = txns.filter(function(t) { return t.match_type === filters.match_type; });
      if (filters.unit_id)    txns = txns.filter(function(t) { return t.match_unit === filters.unit_id; });
      if (filters.month) {
        // t.date is "DD-Mon-YYYY" text (the enforced convention for this
        // column) — comparing that raw string against a "YYYY-MM" key
        // via indexOf() would never match anything, silently returning
        // zero rows for every month. Must actually parse it first.
        txns = txns.filter(function(t) {
          var parsed = parseFlexibleDate(t.date);
          return parsed && parsed.indexOf(filters.month) === 0;
        });
      }
      if (filters.txn_type)   txns = txns.filter(function(t) { return t.txn_type === filters.txn_type; });
    }

    // Ascending chronological order (1 Jan → 31 Dec), parsing real dates
    // rather than comparing text, so mixed formats sort correctly.
    return txns.sort(function(a, b) { return _dateKey(a.date) - _dateKey(b.date); });
  }

  function getStats() {
    var txns      = getAllTransactions(null);
    var credits   = txns.filter(function(t) { return Number(t.credit) > 0; });
    var matched   = credits.filter(function(t) { return t.match_type !== 'unmatched'; });
    var unmatched = credits.filter(function(t) { return t.match_type === 'unmatched'; });

    var totalCredit = credits.reduce(function(s, t) { return s + Number(t.credit || 0); }, 0);

    return {
      totalTransactions: txns.length,
      totalCredits:      credits.length,
      totalMatched:      matched.length,
      totalUnmatched:    unmatched.length,
      totalCreditAmount: totalCredit,
      matchRate:         credits.length > 0 ? Math.round((matched.length / credits.length) * 100) : 0
    };
  }

  // Manual match: manager assigns a txn to an owner or tenant
  // Clears a wrong match entirely — the transaction returns to
  // "unmatched" for a correct re-match later. Two-phase, per the house
  // rule: if the wrong match auto-posted Fees records, the first call
  // returns their exact list; only an explicit confirmation removes
  // them along with the match. Nothing happens silently.
  function unmatchTransaction(txnId, confirmedHistorical, confirmRemovePosted) {
    var result = Database.findByColumn(SHEET, C.TXN_ID, txnId);
    if (!result) throw new Error('Transaction not found: ' + txnId);
    var row = result.data;
    var ym = _ymText(row[C.DATE]);
    if (ym) checkYearEditable(ym.y, confirmedHistorical);

    // Fees records this wrong match created (auto-posted, BANK:<id>)
    var posted = [];
    var pays = Database.getAll('Payments');
    var re = new RegExp('^BANK:' + txnId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b');
    for (var i = 0; i < pays.length; i++) {
      if (re.test(String(pays[i][8] || ''))) {
        posted.push({ rowIndex: i + 2, payment_id: String(pays[i][0]), unit_id: String(pays[i][1]),
                      payment_type: String(pays[i][3]), amount: Number(pays[i][4]) || 0,
                      month: String(pays[i][5] || '').replace(/^'/, '') });
      }
    }
    if (posted.length && confirmRemovePosted !== true) {
      return { needs_confirm: true,
               records: posted.map(function(p) { return { unit_id: p.unit_id, payment_type: p.payment_type, month: p.month, amount: p.amount }; }) };
    }
    if (posted.length) {
      var pSheet = Database.getSheet('Payments');
      posted.sort(function(a, b) { return b.rowIndex - a.rowIndex; });
      posted.forEach(function(p) { pSheet.deleteRow(p.rowIndex); });
    }

    row[C.MATCH_TYPE] = 'unmatched';
    row[C.MATCH_ID] = '';
    row[C.MATCH_NAME] = '';
    row[C.MATCH_UNIT] = '';
    Database.updateRow(SHEET, result.rowIndex, row);
    return { success: true, postedRemoved: posted.length };
  }

  function manualMatch(txnId, matchType, matchId, matchName, matchUnit, paymentType, confirmedHistorical) {
    var result = Database.findByColumn(SHEET, C.TXN_ID, txnId);
    if (!result) throw new Error('Transaction not found: ' + txnId);

    var row = result.data;
    var ymCheck = _ymText(row[C.DATE]);
    if (ymCheck) checkYearEditable(ymCheck.y, confirmedHistorical);

    row[C.MATCH_TYPE]    = matchType;
    row[C.MATCH_ID]      = matchId;
    row[C.MATCH_NAME]    = matchName;
    row[C.MATCH_UNIT]    = matchUnit;
    row[C.PAYMENT_TYPE]  = paymentType || row[C.PAYMENT_TYPE];

    Database.updateRow(SHEET, result.rowIndex, row);

    // Post the payment for this credit (idempotent — skipped if already posted)
    var paymentsCreated = 0;
    var credit = Number(row[C.CREDIT] || 0);
    if (matchUnit && credit > 0) {
      try {
        var mmDayFirst = row[C.DATE] instanceof Date ? true : _detectDayFirst([String(row[C.DATE] || '')]);
        var alloc = PaymentsService.allocateFromBankBatch([{
          txn_id:       txnId,
          unit_id:      matchUnit,
          tenant_id:    matchType === 'tenant' ? matchId : '',
          payment_type: row[C.PAYMENT_TYPE],
          amount:       credit,
          dateStr:      _txnMonthKey(row[C.DATE], mmDayFirst) || String(row[C.DATE] || '')
        }], _getFees());
        paymentsCreated = alloc.created;
        // NOTE: the Owners/Tenants sheet Jan–Dec columns are reference data
        // from the imported Excel masters and are intentionally NOT modified.
        // All payment display reads from the central Payments ledger.
      } catch (e) {}
    }
    return { success: true, paymentsCreated: paymentsCreated };
  }

  // Set / change the expense category on a debit transaction.
  // Manually correct a transaction's date — needed for the rare case
  // where a source file's date cell was a genuine Excel date-type cell
  // (rather than plain text like the rest of the column) and Google's own
  // Drive-based XLSX->Sheets conversion shifted it by a day interpreting
  // it in the wrong timezone, before OXYGEN's import code ever saw it.
  // Always writes with the forced-plain-text prefix so Sheets can never
  // silently re-convert this cell into a Date type again.
  function setDate(txnId, newDate, confirmedHistorical) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(newDate)) throw new Error('Date must be in YYYY-MM-DD format.');
    var result = Database.findByColumn(SHEET, C.TXN_ID, txnId);
    if (!result) throw new Error('Transaction not found: ' + txnId);
    var row = result.data;
    var ymCheck = _ymText(row[C.DATE]);
    if (ymCheck) checkYearEditable(ymCheck.y, confirmedHistorical);
    // Store in the SAME "DD-Mon-YYYY" text convention as the rest of the
    // column (e.g. "01-Jan-2024"), not the date picker's raw ISO value —
    // so a manually corrected row never stands out as a different format
    // from its neighbours.
    var parts = newDate.split('-');
    var mn = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var displayDate = parts[2] + '-' + mn[parseInt(parts[1], 10) - 1] + '-' + parts[0];
    row[C.DATE] = "'" + displayDate;
    Database.updateRow(SHEET, result.rowIndex, row);
    return { success: true, date: displayDate };
  }

  // Extracts a stable "who is this payment to/from" signature from a
  // narration, reusing the same UPI/NEFT/IMPS name-extraction logic
  // already proven for credits — the payee name sits in the same
  // position in the narration regardless of which direction the money
  // moves. Falls back to a normalized full narration (numbers/refs
  // stripped) for formats with no extractable name, so even something
  // like a fixed "ATM WDL" charge can still be recognised as recurring.
  function _debitSignature(narration) {
    var parsed = _parseNarration(narration || '');
    if (parsed.name && parsed.name.length >= 3) return parsed.name.toUpperCase();
    return String(narration || '').toUpperCase().replace(/[0-9]+/g, '#').replace(/\s+/g, ' ').trim();
  }

  // Learns expense-category patterns from every debit you've ALREADY
  // categorized (signature -> category, most recent wins on conflict),
  // then applies that same category to every debit that doesn't have one
  // yet but shares the same signature. A recurring vendor/payee — Care
  // Taker salary, an AMC payment, etc. — only needs to be categorized
  // once; every future occurrence with the same narration pattern is
  // then auto-filled. Never overwrites a category you've already set.
  //
  // ALSO checks the explicit patterns defined on the Settings page
  // (ExpensePatternService) first — those are deliberate rules you set
  // up yourself, so they take priority over inferring from history.
  function autoCategorizeExpenses() {
    var sheet = Database.getSheet(SHEET);
    var rows  = Database.getAll(SHEET);
    var explicitPatterns = ExpensePatternService.getAll();

    var learned = {}; // signature -> category
    for (var i = 0; i < rows.length; i++) {
      var cat = String(rows[i][C.EXPENSE_CAT] || '').trim();
      var isDebit = Number(rows[i][C.DEBIT] || 0) > 0;
      if (!isDebit || !cat) continue;
      var sig = _debitSignature(String(rows[i][C.NARRATION] || ''));
      if (sig) learned[sig] = cat; // later (more recent) rows overwrite earlier ones
    }

    var applied = 0, examples = [];
    for (var j = 0; j < rows.length; j++) {
      var row = rows[j];
      var existingCat = String(row[C.EXPENSE_CAT] || '').trim();
      var isDebitRow = Number(row[C.DEBIT] || 0) > 0;
      if (!isDebitRow || existingCat) continue; // never overwrite an existing category
      var narrationJ = String(row[C.NARRATION] || '');
      var amountJ = Number(row[C.DEBIT] || 0);

      // Explicit pattern (from Settings) wins over the learned signature.
      var matchedCat = ExpensePatternService.matchCategory(narrationJ, amountJ, explicitPatterns);
      if (!matchedCat) {
        var sigJ = _debitSignature(narrationJ);
        if (sigJ && learned[sigJ]) matchedCat = learned[sigJ];
      }
      if (matchedCat) {
        row[C.EXPENSE_CAT] = matchedCat;
        if (!row[C.DIRECTION]) row[C.DIRECTION] = 'debit';
        Database.updateRow(SHEET, j + 2, row);
        applied++;
        if (examples.length < 8) examples.push({ narration: narrationJ, category: matchedCat });
      }
    }
    return { success: true, applied: applied, patternsLearned: Object.keys(learned).length, examples: examples };
  }

  function setExpenseCategory(txnId, category, confirmedHistorical) {
    var result = Database.findByColumn(SHEET, C.TXN_ID, txnId);
    if (!result) throw new Error('Transaction not found: ' + txnId);
    var row = result.data;
    var ymCat = _ymText(row[C.DATE]);
    if (ymCat) checkYearEditable(ymCat.y, confirmedHistorical);
    while (row.length <= C.EXPENSE_CAT) row.push('');
    row[C.EXPENSE_CAT] = category || '';
    // Make sure direction reflects that this is a debit
    if (!row[C.DIRECTION]) {
      row[C.DIRECTION] = Number(row[C.DEBIT] || 0) > 0 ? 'debit'
                       : (Number(row[C.CREDIT] || 0) > 0 ? 'credit' : '');
    }
    Database.updateRow(SHEET, result.rowIndex, row);
    return { success: true };
  }

  // Manually set the credit "Type" (Maintenance / Waste Management / LPG /
  // Bank Interest / Miscellaneous) on a money-received transaction.
  // Restores a plain fee type on a bank row that still carries a legacy
  // "Split (N)" label from the old design (where a split overwrote the
  // row's Type — the architecture flaw that broke Type filtering and
  // editing). Under the current design the bank row's Type is never
  // touched by a split, so this only ever changes rows stuck with that
  // legacy label — a genuine type like "Maintenance" is left alone.
  function healSplitLabel(txnId, fallbackType) {
    var result = Database.findByColumn(SHEET, C.TXN_ID, txnId);
    if (!result) return;
    var cur = String(result.data[C.PAYMENT_TYPE] || '');
    if (!/^Split \(/.test(cur)) return; // real type — never touch it
    setPaymentType(txnId, fallbackType || 'Unknown', true);
  }

  function setPaymentType(txnId, paymentType, confirmedHistorical) {
    var result = Database.findByColumn(SHEET, C.TXN_ID, txnId);
    if (!result) throw new Error('Transaction not found: ' + txnId);
    var row = result.data;
    var ymPt = _ymText(row[C.DATE]);
    if (ymPt) checkYearEditable(ymPt.y, confirmedHistorical);
    while (row.length <= C.NOTE) row.push('');
    row[C.PAYMENT_TYPE] = paymentType || '';
    if (!row[C.DIRECTION]) {
      row[C.DIRECTION] = Number(row[C.CREDIT] || 0) > 0 ? 'credit'
                       : (Number(row[C.DEBIT] || 0) > 0 ? 'debit' : '');
    }
    Database.updateRow(SHEET, result.rowIndex, row);
    return { success: true };
  }

  // Set / change the free-text note on any transaction (credit or debit).
  function setNote(txnId, note, confirmedHistorical) {
    var result = Database.findByColumn(SHEET, C.TXN_ID, txnId);
    if (!result) throw new Error('Transaction not found: ' + txnId);
    var row = result.data;
    var ymNote = _ymText(row[C.DATE]);
    if (ymNote) checkYearEditable(ymNote.y, confirmedHistorical);
    while (row.length <= C.NOTE) row.push('');
    row[C.NOTE] = note || '';
    Database.updateRow(SHEET, result.rowIndex, row);
    return { success: true };
  }

  // ── Reconciliation: OXYGEN ledger vs the bank's own running balance ──
  // For every month: opening + credits − debits must equal the bank's stated
  // closing balance. Any difference means rows are missing, duplicated or
  // altered — the independent month-end spot check.
  function _ymText(s) {
    if (!s) return null;
    // Same fix as _dateKey — a cell Sheets silently auto-converted to a
    // real Date object must be read via its own calendar fields, never
    // via String(dateObject) + regex.
    if (s instanceof Date) return { y: s.getFullYear(), m: s.getMonth() + 1 };
    s = String(s).trim();
    var low = s.toLowerCase();
    var mn = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
    var m = s.match(/^(\d{4})-(\d{1,2})/);
    if (m) return { y:+m[1], m:+m[2] };
    m = low.match(/^(\d{1,2})[-\/ ]([a-z]{3})[a-z]*[-\/ ](\d{2,4})/);
    if (m && mn[m[2]]) { var y1=+m[3]; if (y1<100) y1+=2000; return { y:y1, m:mn[m[2]] }; }
    m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/);
    if (m) { var yy=+m[3]; if (yy<100) yy+=2000; var a=+m[1], b=+m[2], mo=b; if (b>12&&a<=12) mo=a; return { y:yy, m:mo }; }
    return null;
  }

  // Row-by-row detail for ONE month — walks every transaction in date
  // order, tracking a running computed balance next to the bank's own
  // stated balance on each row, so the EXACT row where they first
  // diverge is visible (rather than only knowing the month's total is
  // off). This is the tool to use when reconcile() reports a mismatch.
  function reconcileMonthDetail(monthKey) {
    var rows = Database.getAll(SHEET).slice().sort(function(a, b) {
      return _dateKey(a[C.DATE]) - _dateKey(b[C.DATE]);
    });
    var running = null, out = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var ym = _ymText(r[C.DATE]);
      if (!ym) continue;
      var key = ym.y + '-' + (ym.m < 10 ? '0' + ym.m : ym.m);
      var cr  = Number(String(r[C.CREDIT]  || '').replace(/[₹,\s]/g, '')) || 0;
      var db  = Number(String(r[C.DEBIT]   || '').replace(/[₹,\s]/g, '')) || 0;
      var bal = Number(String(r[C.BALANCE] || '').replace(/[₹,\s]/g, ''));
      var hasBal = !isNaN(bal);

      // Establish the running balance the first time we see ANY stated
      // balance, even in an earlier month, so this month's opening row
      // has something correct to compare against.
      if (running === null && hasBal) running = bal - cr + db;
      if (running !== null) running = running + cr - db;

      if (key !== monthKey) continue;
      var diff = hasBal ? Math.round(((running || 0) - bal) * 100) / 100 : null;
      out.push({
        date: r[C.DATE], narration: String(r[C.NARRATION] || ''),
        credit: cr, debit: db,
        // Raw, unparsed cell contents + their JS type — if a debit is
        // silently being read as 0, this reveals whether the cell holds
        // something the parser can't recognise (e.g. text with an unusual
        // character, or truly blank) rather than a real zero.
        debitRaw: r[C.DEBIT], debitRawType: typeof r[C.DEBIT],
        creditRaw: r[C.CREDIT], creditRawType: typeof r[C.CREDIT],
        expenseCategory: String(r[C.EXPENSE_CAT] || ''),
        computedBalance: running === null ? null : Math.round(running * 100) / 100,
        statedBalance: hasBal ? bal : null,
        diff: diff, ok: diff === null || Math.abs(diff) < 0.01
      });
      if (hasBal) running = bal; // resync running to the bank's own figure after comparing
    }
    return out;
  }

  // Every distinct month present in the Transaction Ledger, with a row
  // count — for the Danger Zone's "delete a specific month" table. Uses
  // the same robust _ymText() as reconcile(), so it correctly handles
  // both plain-text dates and any leftover Date-typed cells.
  function getMonthSummary() {
    var rows = Database.getAll(SHEET);
    var counts = {};
    for (var i = 0; i < rows.length; i++) {
      var ym = _ymText(rows[i][C.DATE]);
      if (!ym) continue;
      var key = ym.y + '-' + (ym.m < 10 ? '0' + ym.m : ym.m);
      counts[key] = (counts[key] || 0) + 1;
    }
    var out = [];
    for (var k in counts) out.push({ month: k, count: counts[k] });
    out.sort(function(a, b) { return a.month < b.month ? -1 : 1; });
    return out;
  }

  // Deletes every Transaction Ledger row for one month — AND, since Fees
  // Received payments posted from those transactions have no independent
  // meaning once their source rows are gone, cascades to delete that same
  // month's Payments records too. This keeps the two stores from ever
  // silently disagreeing (a bank-derived fee record surviving after its
  // originating transaction was deleted).
  function deleteByMonth(monthKey) {
    var sheet = Database.getSheet(SHEET);
    var rows = Database.getAll(SHEET);
    var deleted = 0;
    var deletedTxnIds = [];
    for (var i = rows.length - 1; i >= 0; i--) {
      var ym = _ymText(rows[i][C.DATE]);
      if (!ym) continue;
      var key = ym.y + '-' + (ym.m < 10 ? '0' + ym.m : ym.m);
      if (key === monthKey) {
        deletedTxnIds.push(String(rows[i][C.TXN_ID]));
        sheet.deleteRow(i + 2);
        deleted++;
      }
    }
    var paymentsResult = { deleted: 0 };
    try { paymentsResult = PaymentsService.deleteByMonth(monthKey); } catch (e) {}
    // Portions describe rows that no longer exist — they go too. Manual
    // month allocations are user-approved data and are NEVER touched by
    // this cascade; if their parked backing disappears, the Parked Funds
    // card shows the pair as over-allocated rather than deleting anything.
    var portionsResult = { deleted: 0 };
    try { portionsResult = PortionsService.deleteForTxns(deletedTxnIds); } catch (e2) {}
    return { success: true, deleted: deleted, paymentsDeleted: paymentsResult.deleted || 0, portionsDeleted: portionsResult.deleted || 0 };
  }

  // Accepts an OPTIONAL pre-fetched, pre-sorted transactions array (the
  // same shape getAllTransactions() returns). If the caller already has
  // this data — as the Dashboard does, for its own income/expense summary
  // — passing it in avoids a second full read-and-sort of the entire
  // Transaction Ledger on every single call. Standalone callers (the Bank
  // Statement page's own Reconcile button) simply omit it and it fetches
  // fresh, exactly as before.
  // Local day-of-month reader for the completeness heuristic below —
  // deliberately separate from the shared _ymText (year/month only,
  // used by many other callers) to avoid changing its contract.
  function _dayOfMonth(s) {
    if (!s) return null;
    if (s instanceof Date) return s.getDate();
    s = String(s).trim();
    var m = s.match(/^\d{4}-\d{1,2}-(\d{1,2})/);
    if (m) return +m[1];
    m = s.match(/^(\d{1,2})[-\/]/);
    if (m) return +m[1];
    return null;
  }

  // year (optional): when given, only that year's months are RETURNED —
  // but the whole ledger is still walked internally first, so each
  // shown month's opening balance still chains correctly from its true
  // prior month/year, exactly as if every year were displayed. Filtering
  // is purely a display choice; it never changes what's computed.
  function reconcile(preloadedTxns, year) {
    var txns = preloadedTxns || getAllTransactions(null); // getAllTransactions already sorts by real date
    var months = {}, order = [], prevClose = null;
    for (var i = 0; i < txns.length; i++) {
      var t  = txns[i];
      var ym = _ymText(t.date);
      if (!ym) continue;
      var cr  = Number(String(t.credit || '').replace(/[₹,\s]/g, '')) || 0;
      var db  = Number(String(t.debit  || '').replace(/[₹,\s]/g, '')) || 0;
      var bal = (t.balance === '' || t.balance === null || t.balance === undefined) ? NaN : Number(String(t.balance).replace(/[₹,\s]/g, ''));
      var key = ym.y + '-' + (ym.m < 10 ? '0' + ym.m : ym.m);
      if (!months[key]) {
        var opening = prevClose;
        if (opening === null && !isNaN(bal)) opening = bal - cr + db;  // derive from first row
        months[key] = { month: key, opening: opening || 0, credits: 0, debits: 0, closing: null, rows: 0, lastDay: 0 };
        order.push(key);
      }
      months[key].credits += cr;
      months[key].debits  += db;
      months[key].rows++;
      var dom = _dayOfMonth(t.date);
      if (dom && dom > months[key].lastDay) months[key].lastDay = dom;
      if (!isNaN(bal)) { months[key].closing = bal; prevClose = bal; }
    }

    // "Possibly partial" heuristic: a month that reconciles perfectly
    // (internally consistent) but whose last imported transaction falls
    // well short of that month's actual last day. A ✓ here only proves
    // the rows given add up among themselves — it can't distinguish a
    // complete month from a partial-but-consistent one, so this flag is
    // the difference between "correct" and "correct AND complete."
    // Never flags the current, still-in-progress month.
    var tz = Session.getScriptTimeZone();
    var todayKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
    var PARTIAL_TOLERANCE_DAYS = 3; // allow for the bank simply not posting the last day or two

    var out = [];
    for (var k = 0; k < order.length; k++) {
      var mm = months[order[k]];
      var computed = mm.opening + mm.credits - mm.debits;
      var diff = (mm.closing === null) ? null : Math.round((mm.closing - computed) * 100) / 100;
      var ok = diff !== null && Math.abs(diff) < 0.01;

      var yearNum = Number(mm.month.split('-')[0]), monthNum = Number(mm.month.split('-')[1]);
      var daysInThisMonth = new Date(yearNum, monthNum, 0).getDate();
      // The "possibly partial" heuristic assumes an ACTIVE account where
      // falling short of month-end usually means "not fully imported
      // yet." The IOB LPG account is deliberately near-dormant — a
      // month with a single interest credit and nothing else is
      // genuinely complete, not partial — so the heuristic would just
      // be noise there. It stays fully active for the main account.
      var possiblyPartial = ACCOUNT !== '2' && ok && mm.month < todayKey && (daysInThisMonth - mm.lastDay) > PARTIAL_TOLERANCE_DAYS;

      out.push({ month: mm.month, rows: mm.rows,
                 opening: Math.round(mm.opening * 100) / 100,
                 credits: Math.round(mm.credits * 100) / 100,
                 debits:  Math.round(mm.debits  * 100) / 100,
                 computed: Math.round(computed * 100) / 100,
                 stated:   mm.closing,
                 diff:     diff,
                 ok:       ok,
                 isGap:    false,
                 lastDay:  mm.lastDay,
                 daysInMonth: daysInThisMonth,
                 possiblyPartial: possiblyPartial });
    }

    // Interior gap detection: a month with ZERO transactions sitting
    // between two months that DO have data is otherwise invisible — the
    // chain just silently jumps over it, which only looks fine if the
    // real bank account had zero net movement that month. If it didn't
    // (an interest credit, a stray payment), the gap surfaces as a
    // mismatch on some LATER month instead of naming its true cause.
    // This inserts an explicit marker for the missing month itself, so
    // the real cause is visible immediately rather than discovered by
    // tracing a downstream diff backwards.
    var withGaps = [];
    for (var g = 0; g < out.length; g++) {
      if (g > 0) {
        _monthsBetweenExclusive(out[g - 1].month, out[g].month).forEach(function(mk) {
          var yy = Number(mk.slice(0, 4)), mmk = Number(mk.slice(5, 7));
          withGaps.push({ month: mk, isGap: true, rows: 0, opening: null, credits: null, debits: null,
                           computed: null, stated: null, diff: null, ok: false,
                           lastDay: 0, daysInMonth: new Date(yy, mmk, 0).getDate(), possiblyPartial: false });
        });
      }
      withGaps.push(out[g]);
    }
    out = withGaps;

    if (year) out = out.filter(function(r) { return r.month.slice(0, 4) === String(year); });
    return out;
  }

  // Every calendar month strictly between two "yyyy-MM" keys (exclusive
  // of both endpoints), in order. Used only for gap detection above.
  function _monthsBetweenExclusive(startKey, endKey) {
    var p1 = startKey.split('-').map(Number), p2 = endKey.split('-').map(Number);
    var y = p1[0], m = p1[1] + 1;
    var out = [];
    while (true) {
      if (m > 12) { m = 1; y++; }
      if (y > p2[0] || (y === p2[0] && m >= p2[1])) break;
      out.push(y + '-' + (m < 10 ? '0' + m : m));
      m++;
    }
    return out;
  }

  // Re-run matching on all unmatched transactions
  function rematch() {
    var rows    = Database.getAll(SHEET);
    var owners  = OwnersService.getAllOwners();
    var tenants = TenantsService.getAllTenants();
    var fees    = _getFees();
    var updated = 0;
    var toPost  = [];

    // Resolve date order once across the stored transactions
    var allDates = [];
    for (var ad = 0; ad < rows.length; ad++) {
      if (rows[ad][C.DATE] && !(rows[ad][C.DATE] instanceof Date)) allDates.push(String(rows[ad][C.DATE]));
    }
    var dayFirst = _detectDayFirst(allDates);

    rows.forEach(function(row, i) {
      // Any row still carrying the old 'unit_hint' category (from before
      // it was removed) is treated as needing a fresh look — it gets a
      // real chance at a proper name-based match now, or genuinely falls
      // through to 'unmatched' instead of staying stuck in a category
      // that no longer exists anywhere else in the app.
      var isStaleHint = row[C.MATCH_TYPE] === 'unit_hint';
      if (row[C.MATCH_TYPE] !== 'unmatched' && !isStaleHint) {
        // Already matched in a PREVIOUS run — the match itself still
        // stands, but the parsed_name/parsed_remark shown in the "Name
        // (Bank)" column was captured under whatever parsing logic
        // existed back then. If that's since been corrected, refresh it
        // here too, purely for accurate display — this never touches
        // the match itself, since re-matching an already-correct row is
        // unnecessary and re-parsing is cheap and side-effect-free.
        var refreshed = _parseNarration(String(row[C.NARRATION] || ''));
        if (refreshed.name && refreshed.name !== String(row[C.PARSED_NAME] || '')) {
          row[C.PARSED_NAME]   = refreshed.name;
          row[C.PARSED_REMARK] = refreshed.remark;
          Database.updateRow(SHEET, i + 2, row);
        }
        // Already matched — still queue for posting (dedup makes this safe),
        // so historic matches that never posted payments get back-filled.
        if (row[C.MATCH_UNIT] && Number(row[C.CREDIT] || 0) > 0) {
          toPost.push({
            txn_id:       String(row[C.TXN_ID]),
            unit_id:      String(row[C.MATCH_UNIT]),
            tenant_id:    row[C.MATCH_TYPE] === 'tenant' ? String(row[C.MATCH_ID]) : '',
            payment_type: String(row[C.PAYMENT_TYPE] || ''),
            amount:       Number(row[C.CREDIT]),
            dateStr:      _txnMonthKey(row[C.DATE], dayFirst) || String(row[C.DATE] || '')
          });
        }
        return;
      }
      // Re-parse the RAW narration fresh every time, rather than trusting
      // the parsed_name/parsed_remark columns cached at import time. If
      // the narration-parsing logic has since been improved (as it just
      // was, for a bank's space-delimited IMPS format), rows imported
      // before that fix would otherwise be stuck with their old, wrong
      // cached name forever — "Re-run Matching" would re-score the SAME
      // bad name against everyone and never improve. This also corrects
      // the stored columns so the Match popup and diagnostics show the
      // right thing from now on, not the stale value.
      var freshParsed = _parseNarration(String(row[C.NARRATION] || ''));
      var name    = freshParsed.name;
      var remark  = freshParsed.remark;
      row[C.PARSED_NAME]   = name;
      row[C.PARSED_REMARK] = remark;
      // Only credits that aren't bank charges/interest/BBPS participate
      if (!(Number(row[C.CREDIT] || 0) > 0)) return;
      if (_isSystemRow(String(row[C.NARRATION] || ''), String(row[C.DEBIT] || ''), String(row[C.CREDIT] || ''))) return;
      var match   = _matchPerson(name, remark, owners, tenants);

      if (match.type !== 'unmatched') {
        row[C.MATCH_TYPE]  = match.type;
        row[C.MATCH_ID]    = match.id;
        row[C.MATCH_NAME]  = match.name;
        row[C.MATCH_UNIT]  = match.unit;
        Database.updateRow(SHEET, i + 2, row);
        updated++;
        if (match.unit && Number(row[C.CREDIT] || 0) > 0) {
          toPost.push({
            txn_id:       String(row[C.TXN_ID]),
            unit_id:      String(match.unit),
            tenant_id:    match.type === 'tenant' ? String(match.id) : '',
            payment_type: String(row[C.PAYMENT_TYPE] || ''),
            amount:       Number(row[C.CREDIT]),
            dateStr:      _txnMonthKey(row[C.DATE], dayFirst) || String(row[C.DATE] || '')
          });
        }
      } else {
        // Still unmatched — save the CORRECTED name/remark anyway, so the
        // Match popup and "why didn't this match" diagnostic show the
        // real extracted name next time, instead of repeating the same
        // stale garbage indefinitely just because no match was found.
        Database.updateRow(SHEET, i + 2, row);
      }
    });

    var paymentsCreated = 0;
    var postError = '';
    if (toPost.length > 0) {
      try {
        var alloc = PaymentsService.allocateFromBankBatch(toPost, fees);
        paymentsCreated = alloc.created;
        // NOTE: the Owners/Tenants sheet Jan–Dec columns are reference data
        // from the imported Excel masters and are intentionally NOT modified.
        // All payment display reads from the central Payments ledger.
      } catch (e) { postError = e.message; }
    }

    return { updated: updated, paymentsCreated: paymentsCreated, queuedForPosting: toPost.length, postError: postError };
  }

  // ── CSV Parser (handles quoted fields) ──────────────────────

  function _parseCsv(csv) {
    // Prefer the built-in parser — it correctly handles line breaks
    // INSIDE quoted cells (bank narrations often contain them).
    try {
      var parsed = Utilities.parseCsv(csv);
      if (parsed && parsed.length) return parsed;
    } catch (e) {}

    // Fallback: full state machine, quote-aware for commas AND newlines
    var rows = [], row = [], cell = '', inQ = false;
    for (var i = 0; i < csv.length; i++) {
      var ch = csv[i];
      if (inQ) {
        if (ch === '"') {
          if (csv[i + 1] === '"') { cell += '"'; i++; }
          else inQ = false;
        } else {
          cell += ch;
        }
      } else {
        if (ch === '"') {
          inQ = true;
        } else if (ch === ',') {
          row.push(cell.trim()); cell = '';
        } else if (ch === '\n') {
          row.push(cell.trim()); rows.push(row); row = []; cell = '';
        } else if (ch !== '\r') {
          cell += ch;
        }
      }
    }
    if (cell.length || row.length) { row.push(cell.trim()); rows.push(row); }
    // Drop fully empty rows
    var out = [];
    for (var r = 0; r < rows.length; r++) {
      var any = false;
      for (var c = 0; c < rows[r].length; c++) if (rows[r][c] !== '') { any = true; break; }
      if (any) out.push(rows[r]);
    }
    return out;
  }

  function _toObj(row) {
    function s(i) {
      var v = row[i];
      if (v === undefined || v === null) return '';
      if (v instanceof Date) return v.toISOString().split('T')[0];
      return String(v);
    }
    return {
      txn_id:       s(C.TXN_ID),
      date:         s(C.DATE),
      value_date:   s(C.VALUE_DATE),
      chq_no:       s(C.CHQ_NO),
      narration:    s(C.NARRATION),
      cod:          s(C.COD),
      debit:        s(C.DEBIT),
      credit:       s(C.CREDIT),
      balance:      s(C.BALANCE),
      parsed_name:  s(C.PARSED_NAME),
      parsed_remark:s(C.PARSED_REMARK),
      txn_type:     s(C.TXN_TYPE),
      match_type:   s(C.MATCH_TYPE),
      match_id:     s(C.MATCH_ID),
      match_name:   s(C.MATCH_NAME),
      match_unit:   s(C.MATCH_UNIT),
      payment_type: s(C.PAYMENT_TYPE),
      imported_at:  s(C.IMPORTED_AT),
      direction:    s(C.DIRECTION) || (Number(s(C.CREDIT)) > 0 ? 'credit' : (Number(s(C.DEBIT)) > 0 ? 'debit' : '')),
      expense_category: s(C.EXPENSE_CAT),
      note:         s(C.NOTE)
    };
  }

  // Diagnostic: for a given extracted name, show EXACTLY how it scored
  // against every owner/tenant, sorted best-first — so a "why didn't
  // this auto-match" question can be answered by looking at real numbers
  // instead of guessing. Also flags common causes: extra whitespace, a
  // name matched on the wrong record type (owner vs tenant), or the name
  // simply not being present as a bank-statement alias anywhere on file.
  function debugMatch(name, unitFilter) {
    var owners  = OwnersService.getAllOwners();
    var tenants = TenantsService.getAllTenants();
    var nameU = String(name || '').toUpperCase().replace(/\s+/g, ' ').trim();
    var results = [];

    function scoreAgainst(person, type) {
      if (unitFilter && String(person.unit_id).toUpperCase() !== String(unitFilter).toUpperCase()) return;
      var primaryScore = _nameScore(nameU, String(person.name || '').toUpperCase());
      var bestAliasScore = 0, bestAlias = '';
      var aliasField = type === 'owner' ? person.bank_names : person.bank_name;
      if (aliasField) {
        String(aliasField).split(',').forEach(function(a) {
          var alias = a.trim().toUpperCase();
          if (!alias) return;
          var s = alias === nameU ? 100 : _nameScore(nameU, alias);
          if (s > bestAliasScore) { bestAliasScore = s; bestAlias = a.trim(); }
        });
      }
      results.push({
        type: type, unit_id: person.unit_id, name: person.name,
        primary_score: primaryScore, alias_score: bestAliasScore, matched_alias: bestAlias,
        best_score: Math.max(primaryScore, bestAliasScore)
      });
    }
    owners.forEach(function(o) { scoreAgainst(o, 'owner'); });
    tenants.forEach(function(t) { scoreAgainst(t, 'tenant'); });

    results.sort(function(a, b) { return b.best_score - a.best_score; });
    return { extracted_name: name, normalized: nameU, threshold: 40, candidates: results.slice(0, 8) };
  }

  return {
    ensureSheet:        ensureSheet,
    importCsv:          importCsv,
    getAllTransactions:  getAllTransactions,
    getTransactionsForUnitSplitAware: getTransactionsForUnitSplitAware,
    getDebitsByCategory: getDebitsByCategory,
    getFeeRates:         _getFees,
    getStats:           getStats,
    manualMatch:        manualMatch,
    unmatchTransaction: unmatchTransaction,
    setExpenseCategory: setExpenseCategory,
    autoCategorizeExpenses: autoCategorizeExpenses,
    setDate:            setDate,
    setPaymentType:     setPaymentType,
    healSplitLabel:     healSplitLabel,
    setNote:            setNote,
    rematch:            rematch,
    reconcile:          reconcile,
    debugMatch:         debugMatch,
    getMonthSummary:    getMonthSummary,
    deleteByMonth:      deleteByMonth,
    reconcileMonthDetail: reconcileMonthDetail,
    resetBankData:      resetBankData
  };
}

// Account 1 — the main IOB account (all collections today)
var BankService  = _makeBankService('BankStatements', '1');
// Account 2 — the older IOB LPG account (historical transactions,
// occasional mistaken payments, and ongoing bank interest)
var Bank2Service = _makeBankService('BankStatements2', '2');
