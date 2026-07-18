// ═══════════════════════════════════════════════════════════════
// ReportPdfService.gs — renders report data (from ReportService.gs)
// as a pixel-styled PDF, by building a temporary Google Sheet with
// the exact colors/fonts/layout from the association's own
// Treasurer's/Caretaker's Monthly Report templates, then exporting
// that sheet region as a PDF via the Sheets export endpoint.
// ═══════════════════════════════════════════════════════════════
var ReportPdfService = (function() {

  // Colors extracted directly from the association's own template file
  // (not guessed from the screenshot) — see conversation history for
  // exact hex values pulled from each cell's fill color.
  var COLORS = {
    titleBand:     '#0E2D4A', // title/subtitle header band
    balanceBar:    '#000000', // "ENDING BALANCE" black bars
    incomeTitle:   '#274E13', // "INCOME (DEPOSITS)" section title + data rows
    incomeHeader:  '#38761D', // "# Date Description Amount" sub-header (income)
    expenseTitle:  '#5B0F00', // "EXPENSES (PAID)" section title + data rows
    expenseHeader: '#85200C', // "# Date Description Amount" sub-header (expenses)
    white:         '#FFFFFF'
  };

  var TEMP_PREFIX = '__pdf_temp_';

  function _newTempSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var name = TEMP_PREFIX + Utilities.getUuid().slice(0, 8);
    var sh = ss.insertSheet(name);
    return sh;
  }

  function _exportSheetAsPdf(sh, landscape) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
      '?format=pdf&gid=' + sh.getSheetId() +
      '&size=A4&portrait=' + (landscape ? 'false' : 'true') + '&fitw=true&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4' +
      '&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED&horizontal_alignment=CENTER';
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      throw new Error('PDF export failed (HTTP ' + response.getResponseCode() + ').');
    }
    return response.getBlob();
  }

  function _cleanupTempSheet(sh) {
    try { SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sh); } catch (e) {}
  }

  // Writes one italic data row across 4 logical columns (#, Date,
  // Description, Amount), matching the template's merged-cell layout —
  // simplified to plain (non-merged) columns for reliability, since the
  // visual result is identical either way.
  function _writeDataRow(sh, r, seq, date, desc, amount, bg) {
    sh.getRange(r, 1).setValue(seq);
    sh.getRange(r, 2).setValue(date);
    sh.getRange(r, 3).setValue(desc);
    sh.getRange(r, 4).setValue(amount === '' ? '' : formatRupeesPlain(amount));
    var range = sh.getRange(r, 1, 1, 4);
    range.setBackground(bg).setFontColor(COLORS.white).setFontStyle('italic').setFontSize(9);
    sh.getRange(r, 4).setHorizontalAlignment('right');
    sh.getRange(r, 1).setHorizontalAlignment('center');
  }

  function _writeSectionHeader(sh, r, label, bg, boldOnly) {
    var range = sh.getRange(r, 1, 1, 4);
    range.merge().setValue(label).setBackground(bg).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(10);
  }

  function _writeColumnHeader(sh, r, bg) {
    sh.getRange(r, 1).setValue('#');
    sh.getRange(r, 2).setValue('Date');
    sh.getRange(r, 3).setValue('Description');
    sh.getRange(r, 4).setValue('Amount');
    var range = sh.getRange(r, 1, 1, 4);
    range.setBackground(bg).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(9);
    sh.getRange(r, 4).setHorizontalAlignment('right');
    sh.getRange(r, 1).setHorizontalAlignment('center');
  }

  function _writeTotalRow(sh, r, label, amount, bg) {
    sh.getRange(r, 1, 1, 3).merge().setValue(label).setBackground(bg).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(10);
    sh.getRange(r, 4).setValue(formatRupeesPlain(amount)).setBackground(bg).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right');
  }

  function _writeBalanceBar(sh, r, label, amount) {
    sh.getRange(r, 1, 1, 3).merge().setValue(label).setBackground(COLORS.balanceBar).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(10);
    sh.getRange(r, 4).setValue(formatRupeesPlain(amount)).setBackground(COLORS.balanceBar).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right');
  }

  function formatRupeesPlain(n) {
    n = Math.round(Number(n) || 0);
    var neg = n < 0; n = Math.abs(n);
    var intPart = String(n);
    var lastThree = intPart.slice(-3), rest = intPart.slice(0, -3);
    if (rest !== '') lastThree = ',' + lastThree;
    var formatted = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + lastThree;
    return (neg ? '-' : '') + '₹' + formatted;
  }

  // ── Treasurer's Monthly Report ──────────────────────────────────
  function buildTreasurerReportPdf(year, month) {
    var d = ReportService.getTreasurerReportData(year, month);
    var sh = _newTempSheet();
    try {
      sh.getRange(1, 1, 80, 4).setFontFamily('Courier New');
      sh.setColumnWidth(1, 30);
      sh.setColumnWidth(2, 70);
      sh.setColumnWidth(3, 340);
      sh.setColumnWidth(4, 110);

      var r = 1;
      // Title band
      sh.getRange(r, 1, 2, 3).merge().setValue("TREASURER'S MONTHLY REPORT")
        .setBackground(COLORS.titleBand).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(16);
      sh.getRange(r, 4, 2, 1).merge().setValue(d.monthName)
        .setBackground(COLORS.titleBand).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
      r += 2;
      sh.getRange(r, 1, 1, 3).merge().setValue('Confident Daffodils Owners Association (CDOA)')
        .setBackground(COLORS.titleBand).setFontColor(COLORS.white).setFontSize(9);
      sh.getRange(r, 4).setValue(d.reportDate)
        .setBackground(COLORS.titleBand).setFontColor(COLORS.white).setFontSize(9).setHorizontalAlignment('center');
      r++;
      sh.getRange(r, 1, 1, 3).merge().setValue('For the Month of   ' + d.monthName + ' ' + d.year)
        .setBackground(COLORS.titleBand).setFontColor(COLORS.white).setFontSize(9);
      r += 2;

      _writeBalanceBar(sh, r, 'ENDING BANK BALANCE AS OF (LAST DAY OF ' + d.prevMonthLabel + ' )', d.openingBalance);
      if (d.openingBreakdown && d.openingBreakdown.iobLpg) {
        r++;
        sh.getRange(r, 3).setValue('  Main account ' + formatRupeesPlain(d.openingBreakdown.main) + ' + IOB LPG account ' + formatRupeesPlain(d.openingBreakdown.iobLpg)).setFontSize(8).setFontColor('#888888').setFontStyle('italic');
      }
      r += 2;

      _writeSectionHeader(sh, r, 'INCOME (DEPOSITS)', COLORS.incomeTitle); r++;
      _writeColumnHeader(sh, r, COLORS.incomeHeader); r++;
      d.incomeRows.forEach(function(row) {
        var desc = row.account ? row.description + '  [' + row.account + ']' : row.description;
        _writeDataRow(sh, r, row.seq, row.date, desc, row.amount, COLORS.incomeTitle); r++;
      });
      _writeTotalRow(sh, r, 'INCOME (DEPOSITS)', d.incomeTotal, COLORS.incomeTitle); r++;
      _writeBalanceBar(sh, r, 'TOTAL [BANK BALANCE + INCOME]', d.totalWithIncome); r += 2;

      _writeSectionHeader(sh, r, 'EXPENSES (PAID)', COLORS.expenseTitle); r++;
      _writeColumnHeader(sh, r, COLORS.expenseHeader); r++;
      d.expenseRows.forEach(function(row) {
        var desc2 = row.account ? row.description + '  [' + row.account + ']' : row.description;
        _writeDataRow(sh, r, row.seq, row.date, desc2, row.amount, COLORS.expenseTitle); r++;
      });
      _writeTotalRow(sh, r, 'TOTAL EXPENSES', d.expenseTotal, COLORS.expenseTitle); r++;
      r++;
      _writeBalanceBar(sh, r, 'ENDING BANK BALANCE AS OF (LAST DAY OF ' + d.monthName.slice(0,3) + ' )', d.closingBalance); r++;
      if (d.closingBreakdown && d.closingBreakdown.iobLpg) {
        sh.getRange(r, 3).setValue('  Main account ' + formatRupeesPlain(d.closingBreakdown.main) + ' + IOB LPG account ' + formatRupeesPlain(d.closingBreakdown.iobLpg)).setFontSize(8).setFontColor('#888888').setFontStyle('italic');
        r++;
      }

      SpreadsheetApp.flush();
      var blob = _exportSheetAsPdf(sh);
      blob.setName('Treasurers_Report_' + d.monthName + '_' + d.year + '.pdf');
      return blob;
    } finally {
      _cleanupTempSheet(sh);
    }
  }


  // Shared title band writer — same layout as the Treasurer's report.
  function _writeTitle(sh, title, sub1, sub2, cols) {
    sh.getRange(1, 1, 2, cols - 1).merge().setValue(title)
      .setBackground(COLORS.titleBand).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(16);
    sh.getRange(1, cols, 2, 1).merge().setValue(sub2)
      .setBackground(COLORS.titleBand).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    sh.getRange(3, 1, 1, cols).merge().setValue(sub1)
      .setBackground(COLORS.titleBand).setFontColor(COLORS.white).setFontSize(9);
  }

  // ── Caretaker's Monthly Report ──────────────────────────────────
  function buildCaretakerReportPdf(year, month) {
    var d = CaretakerService.getMonthReport(year, month);
    var mnFull = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
    var sh = _newTempSheet();
    try {
      sh.getRange(1, 1, 200, 4).setFontFamily('Courier New');
      sh.setColumnWidth(1, 30); sh.setColumnWidth(2, 70); sh.setColumnWidth(3, 340); sh.setColumnWidth(4, 110);
      _writeTitle(sh, "CARETAKER'S MONTHLY REPORT", 'Confident Daffodils Owners Association (CDOA)  ·  For the Month of ' + mnFull[month] + ' ' + year, mnFull[month], 4);
      var r = 5;
      _writeBalanceBar(sh, r, 'OPENING BALANCE (CARRIED FORWARD)', d.opening); r += 2;

      _writeSectionHeader(sh, r, 'RESERVE MONEY RECEIVED', COLORS.incomeTitle); r++;
      _writeColumnHeader(sh, r, COLORS.incomeHeader); r++;
      (d.reserveMoney || []).forEach(function(row, i) {
        _writeDataRow(sh, r, i + 1, row.date, row.description, row.amount, COLORS.incomeTitle); r++;
      });
      _writeTotalRow(sh, r, 'TOTAL RESERVE MONEY RECEIVED', d.reserveTotal, COLORS.incomeTitle); r++;
      r++;

      _writeSectionHeader(sh, r, "CARETAKER'S EXPENSES", COLORS.expenseTitle); r++;
      _writeColumnHeader(sh, r, COLORS.expenseHeader); r++;
      (d.expenses || []).forEach(function(row, i) {
        _writeDataRow(sh, r, i + 1, row.date, (row.voucher_no ? row.voucher_no + ' — ' : '') + row.description, row.amount, COLORS.expenseTitle); r++;
      });
      _writeTotalRow(sh, r, 'TOTAL EXPENSES', d.expenseTotal, COLORS.expenseTitle); r++;
      r++;
      _writeBalanceBar(sh, r, 'CLOSING BALANCE (CARRIED TO NEXT MONTH)', d.closing);

      SpreadsheetApp.flush();
      var blob = _exportSheetAsPdf(sh);
      blob.setName('Caretakers_Report_' + mnFull[month] + '_' + year + '.pdf');
      return blob;
    } finally { _cleanupTempSheet(sh); }
  }

  // ── Monthly Income (per-transaction detail) ─────────────────────
  function buildMonthlyIncomePdf(year, month) {
    var d = ReportService.getMonthlyIncomeData(year, month);
    var sh = _newTempSheet();
    try {
      sh.getRange(1, 1, Math.max(60, d.rows.length + 30), 5).setFontFamily('Courier New');
      sh.setColumnWidth(1, 30); sh.setColumnWidth(2, 65); sh.setColumnWidth(3, 230); sh.setColumnWidth(4, 140); sh.setColumnWidth(5, 100);
      _writeTitle(sh, 'MONTHLY INCOME REPORT', 'Confident Daffodils Owners Association (CDOA)  ·  ' + d.monthName + ' ' + d.year + '  ·  Every credit received, by transaction', d.monthName, 5);
      var r = 5;

      _writeSectionHeader(sh, r, 'INCOME (' + d.rows.length + ' transaction' + (d.rows.length === 1 ? '' : 's') + ')', COLORS.incomeTitle); r++;
      sh.getRange(r, 1).setValue('#'); sh.getRange(r, 2).setValue('Date'); sh.getRange(r, 3).setValue('Payer'); sh.getRange(r, 4).setValue('Type / Unit'); sh.getRange(r, 5).setValue('Amount');
      sh.getRange(r, 1, 1, 5).setBackground(COLORS.incomeHeader).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(9);
      sh.getRange(r, 5).setHorizontalAlignment('right'); r++;
      d.rows.forEach(function(row, i) {
        var typeUnit = row.type + (row.unit ? ' · ' + row.unit : '') + (row.split ? ' (split)' : '') + (row.account ? '  [' + row.account + ']' : '');
        sh.getRange(r, 1).setValue(i + 1).setHorizontalAlignment('center');
        sh.getRange(r, 2).setValue(row.date);
        sh.getRange(r, 3).setValue(row.payer);
        sh.getRange(r, 4).setValue(typeUnit);
        sh.getRange(r, 5).setValue(formatRupeesPlain(row.amount)).setHorizontalAlignment('right');
        sh.getRange(r, 1, 1, 5).setBackground(COLORS.incomeTitle).setFontColor(COLORS.white).setFontStyle('italic').setFontSize(9);
        r++;
      });
      r++;
      _writeSectionHeader(sh, r, 'TOTALS BY TYPE', COLORS.incomeHeader); r++;
      d.typeTotals.forEach(function(t) {
        sh.getRange(r, 1, 1, 4).merge().setValue(t.label).setBackground(COLORS.incomeTitle).setFontColor(COLORS.white).setFontSize(9);
        sh.getRange(r, 5).setValue(formatRupeesPlain(t.total)).setBackground(COLORS.incomeTitle).setFontColor(COLORS.white).setFontSize(9).setHorizontalAlignment('right');
        r++;
      });
      sh.getRange(r, 1, 1, 4).merge().setValue('TOTAL INCOME — ' + d.monthName.toUpperCase() + ' ' + d.year).setBackground(COLORS.balanceBar).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(10);
      sh.getRange(r, 5).setValue(formatRupeesPlain(d.grandTotal)).setBackground(COLORS.balanceBar).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right');

      SpreadsheetApp.flush();
      var blob = _exportSheetAsPdf(sh);
      blob.setName('Monthly_Income_' + d.monthName + '_' + d.year + '.pdf');
      return blob;
    } finally { _cleanupTempSheet(sh); }
  }

  // ── Monthly Expense (per-transaction detail) ────────────────────
  function buildMonthlyExpensePdf(year, month) {
    var d = ReportService.getMonthlyExpenseData(year, month);
    var sh = _newTempSheet();
    try {
      sh.getRange(1, 1, Math.max(60, d.rows.length + 30), 5).setFontFamily('Courier New');
      sh.setColumnWidth(1, 30); sh.setColumnWidth(2, 65); sh.setColumnWidth(3, 150); sh.setColumnWidth(4, 230); sh.setColumnWidth(5, 100);
      _writeTitle(sh, 'MONTHLY EXPENSE REPORT', 'Confident Daffodils Owners Association (CDOA)  ·  ' + d.monthName + ' ' + d.year + '  ·  Every payment out, by transaction', d.monthName, 5);
      var r = 5;

      _writeSectionHeader(sh, r, 'EXPENSES (' + d.rows.length + ' transaction' + (d.rows.length === 1 ? '' : 's') + ')', COLORS.expenseTitle); r++;
      sh.getRange(r, 1).setValue('#'); sh.getRange(r, 2).setValue('Date'); sh.getRange(r, 3).setValue('Category'); sh.getRange(r, 4).setValue('Description'); sh.getRange(r, 5).setValue('Amount');
      sh.getRange(r, 1, 1, 5).setBackground(COLORS.expenseHeader).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(9);
      sh.getRange(r, 5).setHorizontalAlignment('right'); r++;
      d.rows.forEach(function(row, i) {
        var desc = row.description + (row.account ? '  [' + row.account + ']' : '');
        sh.getRange(r, 1).setValue(i + 1).setHorizontalAlignment('center');
        sh.getRange(r, 2).setValue(row.date);
        sh.getRange(r, 3).setValue(row.category);
        sh.getRange(r, 4).setValue(desc);
        sh.getRange(r, 5).setValue(formatRupeesPlain(row.amount)).setHorizontalAlignment('right');
        sh.getRange(r, 1, 1, 5).setBackground(COLORS.expenseTitle).setFontColor(COLORS.white).setFontStyle('italic').setFontSize(9);
        r++;
      });
      r++;
      _writeSectionHeader(sh, r, 'TOTALS BY CATEGORY', COLORS.expenseHeader); r++;
      d.categoryTotals.forEach(function(t) {
        sh.getRange(r, 1, 1, 4).merge().setValue(t.category).setBackground(COLORS.expenseTitle).setFontColor(COLORS.white).setFontSize(9);
        sh.getRange(r, 5).setValue(formatRupeesPlain(t.total)).setBackground(COLORS.expenseTitle).setFontColor(COLORS.white).setFontSize(9).setHorizontalAlignment('right');
        r++;
      });
      sh.getRange(r, 1, 1, 4).merge().setValue('TOTAL EXPENSES — ' + d.monthName.toUpperCase() + ' ' + d.year).setBackground(COLORS.balanceBar).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(10);
      sh.getRange(r, 5).setValue(formatRupeesPlain(d.grandTotal)).setBackground(COLORS.balanceBar).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(10).setHorizontalAlignment('right');

      SpreadsheetApp.flush();
      var blob = _exportSheetAsPdf(sh);
      blob.setName('Monthly_Expense_' + d.monthName + '_' + d.year + '.pdf');
      return blob;
    } finally { _cleanupTempSheet(sh); }
  }

  // ── Annual matrices (Income / Expense) — landscape, 14 columns ──
  function _buildAnnualMatrixPdf(title, rows, colTotals, grandTotal, year, labelField, sectionColor, headerColor, filenameBase) {
    var mn3 = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var sh = _newTempSheet();
    try {
      sh.getRange(1, 1, rows.length + 12, 14).setFontFamily('Courier New');
      sh.setColumnWidth(1, 170);
      for (var c = 2; c <= 13; c++) sh.setColumnWidth(c, 62);
      sh.setColumnWidth(14, 84);

      _writeTitle(sh, title, 'Confident Daffodils Owners Association (CDOA)  ·  Calendar Year ' + year + '  ·  Cash basis (by bank date)', String(year), 14);
      var r = 5;

      sh.getRange(r, 1).setValue('');
      for (var m = 0; m < 12; m++) sh.getRange(r, 2 + m).setValue(mn3[m]).setHorizontalAlignment('right');
      sh.getRange(r, 14).setValue('TOTAL').setHorizontalAlignment('right');
      sh.getRange(r, 1, 1, 14).setBackground(headerColor).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(9);
      r++;

      rows.forEach(function(row) {
        sh.getRange(r, 1).setValue(row[labelField]);
        for (var m2 = 0; m2 < 12; m2++) {
          sh.getRange(r, 2 + m2).setValue(row.months[m2] ? formatRupeesPlain(row.months[m2]) : '-').setHorizontalAlignment('right');
        }
        sh.getRange(r, 14).setValue(formatRupeesPlain(row.total)).setHorizontalAlignment('right').setFontWeight('bold');
        sh.getRange(r, 1, 1, 14).setBackground(sectionColor).setFontColor(COLORS.white).setFontSize(8.5);
        r++;
      });

      sh.getRange(r, 1).setValue('MONTH TOTAL');
      for (var m3 = 0; m3 < 12; m3++) sh.getRange(r, 2 + m3).setValue(colTotals[m3] ? formatRupeesPlain(colTotals[m3]) : '-').setHorizontalAlignment('right');
      sh.getRange(r, 14).setValue(formatRupeesPlain(grandTotal)).setHorizontalAlignment('right');
      sh.getRange(r, 1, 1, 14).setBackground(COLORS.balanceBar).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(9);

      SpreadsheetApp.flush();
      var blob = _exportSheetAsPdf(sh, true); // landscape — 14 columns
      blob.setName(filenameBase + '_' + year + '.pdf');
      return blob;
    } finally { _cleanupTempSheet(sh); }
  }

  function buildAnnualIncomePdf(year) {
    var d = ReportService.getAnnualIncomeData(year);
    return _buildAnnualMatrixPdf('ANNUAL INCOME REPORT', d.rows, d.colTotals, d.grandTotal, d.year, 'label', COLORS.incomeTitle, COLORS.incomeHeader, 'Annual_Income');
  }

  function buildAnnualExpensePdf(year) {
    var d = ReportService.getAnnualExpenseData(year);
    return _buildAnnualMatrixPdf('ANNUAL EXPENSE REPORT', d.rows, d.colTotals, d.grandTotal, d.year, 'category', COLORS.expenseTitle, COLORS.expenseHeader, 'Annual_Expense');
  }


  // ── EC Payment Follow-up Board — the association's color-coded
  //    "TEAM ALLOCATION" sheet, exact palette, for print/PDF/WhatsApp ──
  function buildFollowUpPdf(year, month) {
    var d = CommitteeService.getFollowUpBoard(year, month);
    var PAL = { frame1: '#073763', frame2: '#0E2D4A', headerBlack: '#000000',
                paid: '#34A853', pending: '#980000', uo: '#FBBC04', unreg: '#000000', white: '#FFFFFF' };
    var TYPE_SHORT = { 'Maintenance': 'MF', 'Waste Management': 'WMF', 'LPG': 'LPG' };
    var cols = 5 + d.maxUnits; // #, member, TOTAL, DEF, FEE, units...
    var sh = _newTempSheet();
    try {
      var totalRows = 7 + d.members.length * 3 + 8;
      sh.getRange(1, 1, totalRows, cols).setFontFamily('Courier New').setBackground(PAL.frame2).setFontColor(PAL.white);
      sh.setColumnWidth(1, 26); sh.setColumnWidth(2, 150); sh.setColumnWidth(3, 42); sh.setColumnWidth(4, 36); sh.setColumnWidth(5, 42);
      for (var c = 6; c <= cols; c++) sh.setColumnWidth(c, 52);

      sh.getRange(1, 1, 1, cols).merge().setValue('TEAM ALLOCATION — PAYMENT FOLLOW-UP')
        .setBackground(PAL.frame1).setFontColor(PAL.white).setFontWeight('bold').setFontSize(14);
      sh.getRange(2, 1, 1, cols).merge().setValue('Confident Daffodils Owners Association (CDOA)  ·  ' + d.monthName + ' ' + d.year +
          '  ·  Generated: ' + d.generatedDate +
          '  ·  Pending — MF: ' + d.pendingTotals['Maintenance'] + '  ·  WMF: ' + d.pendingTotals['Waste Management'] + '  ·  LPG: ' + d.pendingTotals['LPG'])
        .setBackground(PAL.frame1).setFontColor(PAL.white).setFontSize(9);

      var r = 4;
      var headerRow = r;
      sh.getRange(r, 1).setValue('#'); sh.getRange(r, 2).setValue('TEAM MEMBER');
      sh.getRange(r, 3).setValue('TOTAL'); sh.getRange(r, 4).setValue('DEF'); sh.getRange(r, 5).setValue('FEE');
      sh.getRange(r, 6, 1, d.maxUnits).merge().setValue('FLATS ASSIGNED');
      sh.getRange(r, 1, 1, cols).setBackground(PAL.headerBlack).setFontColor(PAL.white).setFontWeight('bold').setFontSize(9);
      r++;
      // Numbering row 1..N, exactly like the original sheet
      for (var n = 0; n < d.maxUnits; n++) sh.getRange(r, 6 + n).setValue(n + 1);
      sh.getRange(r, 1, 1, cols).setBackground(PAL.frame1).setFontColor(PAL.white).setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center');
      r++;

      var blockStarts = [];
      d.members.forEach(function(m, mi) {
        var frame = mi % 2 === 0 ? PAL.frame1 : PAL.frame2;
        blockStarts.push(r);
        sh.getRange(r, 1, 3, 1).merge().setValue(mi + 1).setBackground(frame).setFontColor(PAL.white).setFontWeight('bold').setVerticalAlignment('middle');
        sh.getRange(r, 2, 3, 1).merge().setValue(m.name).setBackground(frame).setFontColor(PAL.white).setFontWeight('bold').setVerticalAlignment('middle').setFontSize(10);
        sh.getRange(r, 3, 3, 1).merge().setValue(m.totalPending).setBackground(frame).setFontColor(PAL.white).setFontWeight('bold').setFontSize(13).setVerticalAlignment('middle').setHorizontalAlignment('center');
        d.types.forEach(function(ty, ti) {
          sh.getRange(r + ti, 4).setValue(m.perTypePending[ty]).setBackground(frame).setFontColor(PAL.white).setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');
          sh.getRange(r + ti, 5).setValue(TYPE_SHORT[ty]).setBackground(frame).setFontColor(PAL.white).setFontSize(8.5).setFontWeight('bold');
          m.rows[ty].forEach(function(cell, ci) {
            var bg = cell.s === 'paid' ? PAL.paid : cell.s === 'pending' ? PAL.pending : cell.s === 'uo' ? PAL.uo : PAL.unreg;
            var fg = cell.s === 'uo' ? '#000000' : PAL.white;
            sh.getRange(r + ti, 6 + ci).setValue(cell.unit).setBackground(bg).setFontColor(fg).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
          });
          for (var pad = m.unitCount; pad < d.maxUnits; pad++) {
            sh.getRange(r + ti, 6 + pad).setBackground(frame);
          }
        });
        r += 3;
      });
      var boardEnd = r - 1;

      // Dark palette — thin WHITE gridlines across the whole board...
      sh.getRange(headerRow, 1, boardEnd - headerRow + 1, cols)
        .setBorder(true, true, true, true, true, true, PAL.white, SpreadsheetApp.BorderStyle.SOLID);
      // ...and a BOLD white rectangle around each member's 3-row block,
      // so one member's flats read as one clearly framed unit.
      blockStarts.forEach(function(bs) {
        sh.getRange(bs, 1, 3, cols)
          .setBorder(true, true, true, true, null, null, PAL.white, SpreadsheetApp.BorderStyle.SOLID_THICK);
      });

      r++;
      var legend = [['PAID', PAL.paid, PAL.white], ['NOT PAID', PAL.pending, PAL.white], ['UNOCCUPIED or N/A', PAL.uo, '#000000'], ['UNREGISTERED', PAL.unreg, PAL.white]];
      legend.forEach(function(lg) {
        sh.getRange(r, 1).setBackground(lg[1]);
        sh.getRange(r, 2, 1, 3).merge().setValue(lg[0]).setBackground(PAL.frame2).setFontColor(PAL.white).setFontSize(9);
        r++;
      });

      SpreadsheetApp.flush();
      var blob = _exportSheetAsPdf(sh, true); // landscape
      blob.setName('Payment_FollowUp_' + d.monthName + '_' + d.year + '.pdf');
      return blob;
    } finally { _cleanupTempSheet(sh); }
  }

  return {
    buildTreasurerReportPdf: buildTreasurerReportPdf,
    buildCaretakerReportPdf: buildCaretakerReportPdf,
    buildMonthlyIncomePdf:   buildMonthlyIncomePdf,
    buildMonthlyExpensePdf:  buildMonthlyExpensePdf,
    buildAnnualIncomePdf:    buildAnnualIncomePdf,
    buildAnnualExpensePdf:   buildAnnualExpensePdf,
    buildFollowUpPdf:        buildFollowUpPdf
  };
})();
