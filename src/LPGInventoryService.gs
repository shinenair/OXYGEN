// ═══════════════════════════════════════════════════════════════
// LPGInventoryService — cylinder stock (Inward/Outward/Net) and the
// LPG business profit-&-loss view: what residents pay (from meter
// readings) vs what cylinders cost the association (from inventory),
// cross-checked against actual bank spending.
// ═══════════════════════════════════════════════════════════════
var LPGInventoryService = (function() {

  var IN_SHEET  = 'LPGStockInward';
  var OUT_SHEET = 'LPGStockOutward';
  var IC = { ID:0, INVOICE_NO:1, DATE:2, COST_PER_CYL:3, QTY:4, TOTAL:5, RECORDED_BY:6, RECORDED_AT:7 };
  var OC = { ID:0, DATE:1, QTY:2, NOTES:3, RECORDED_BY:4, RECORDED_AT:5 };
  var IN_HEADERS  = ['inward_id','invoice_number','date','cost_per_cylinder','quantity','total_amount','recorded_by','recorded_at'];
  var OUT_HEADERS = ['outward_id','date','quantity','notes','recorded_by','recorded_at'];

  function ensureSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    [[IN_SHEET, IN_HEADERS], [OUT_SHEET, OUT_HEADERS]].forEach(function(pair) {
      var sh = ss.getSheetByName(pair[0]);
      if (!sh) {
        sh = ss.insertSheet(pair[0]);
        sh.appendRow(pair[1]);
        sh.getRange(1, 1, 1, pair[1].length).setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
        sh.setFrozenRows(1);
      }
      // Force the DATE column to plain text, permanently — same reasoning
      // as LPGReadingService: avoids Sheets silently converting a written
      // date string into its own Date type and drifting on read-back.
      var dateCol = pair[0] === IN_SHEET ? IC.DATE + 1 : OC.DATE + 1;
      if (dateCol) sh.getRange(2, dateCol, 3000, 1).setNumberFormat('@');
    });
  }

  function _fmtCellDate(v) {
    if (!v) return '';
    if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    return String(v);
  }

  // ── Stock Inward (cylinder purchases) ───────────────────────────
  function getInward() {
    return Database.getAll(IN_SHEET).map(function(r) {
      return { inward_id: String(r[IC.ID]), invoice_number: String(r[IC.INVOICE_NO] || ''),
               date: _fmtCellDate(r[IC.DATE]), cost_per_cylinder: Number(r[IC.COST_PER_CYL]) || 0,
               quantity: Number(r[IC.QTY]) || 0, total_amount: Number(r[IC.TOTAL]) || 0,
               recorded_by: String(r[IC.RECORDED_BY] || ''), recorded_at: String(r[IC.RECORDED_AT] || '') };
    }).sort(function(a, b) { return a.date < b.date ? -1 : 1; });
  }

  function addInward(data) {
    var qty = Number(data.quantity), cost = Number(data.cost_per_cylinder);
    if (!data.date && !data.allowBlankDate) throw new Error('Date is required.');
    if (isNaN(qty) || qty <= 0) throw new Error('Quantity must be a positive number.');
    if (isNaN(cost) || cost <= 0) throw new Error('Cost per cylinder must be a positive number.');
    // Skipped during bulk Excel import (already a deliberate, admin-only
    // bulk action) — applies only to a single manual add via the UI form.
    if (data.date && !data.skipYearCheck) {
      checkYearEditable(Number(data.date.slice(0, 4)), data.confirmed_historical);
    }
    var total = Math.round(qty * cost * 100) / 100;
    var sheet = Database.getSheet(IN_SHEET);
    sheet.appendRow([Utilities.getUuid(), data.invoice_number || '', data.date || '', cost, qty, total,
                      data.recorded_by || 'Manager', new Date().toISOString()]);
    return { success: true, total_amount: total };
  }

  function deleteInward(id) {
    var sheet = Database.getSheet(IN_SHEET);
    var rows = Database.getAll(IN_SHEET);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][IC.ID]) === String(id)) { sheet.deleteRow(i + 2); return { success: true }; }
    }
    throw new Error('Inward entry not found.');
  }

  // ── Stock Outward (cylinders used) ──────────────────────────────
  function getOutward() {
    return Database.getAll(OUT_SHEET).map(function(r) {
      return { outward_id: String(r[OC.ID]), date: _fmtCellDate(r[OC.DATE]),
               quantity: Number(r[OC.QTY]) || 0, notes: String(r[OC.NOTES] || ''),
               recorded_by: String(r[OC.RECORDED_BY] || ''), recorded_at: String(r[OC.RECORDED_AT] || '') };
    }).sort(function(a, b) { return a.date < b.date ? -1 : 1; });
  }

  function addOutward(data) {
    var qty = Number(data.quantity);
    if (!data.date && !data.allowBlankDate) throw new Error('Date is required.');
    if (isNaN(qty) || qty <= 0) throw new Error('Quantity must be a positive number.');
    if (data.date && !data.skipYearCheck) {
      checkYearEditable(Number(data.date.slice(0, 4)), data.confirmed_historical);
    }
    var sheet = Database.getSheet(OUT_SHEET);
    sheet.appendRow([Utilities.getUuid(), data.date || '', qty, data.notes || '',
                      data.recorded_by || 'Manager', new Date().toISOString()]);
    return { success: true };
  }

  function deleteOutward(id) {
    var sheet = Database.getSheet(OUT_SHEET);
    var rows = Database.getAll(OUT_SHEET);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][OC.ID]) === String(id)) { sheet.deleteRow(i + 2); return { success: true }; }
    }
    throw new Error('Outward entry not found.');
  }

  // Distinct months present in Inward/Outward, with counts — for the
  // Danger Zone's per-month delete tables.
  function getInwardMonthSummary() {
    var counts = {};
    getInward().forEach(function(r) {
      if (!r.date) return;
      var k = r.date.slice(0, 7);
      counts[k] = (counts[k] || 0) + 1;
    });
    var out = []; for (var k in counts) out.push({ month: k, count: counts[k] });
    out.sort(function(a, b) { return a.month < b.month ? -1 : 1; });
    return out;
  }
  function getOutwardMonthSummary() {
    var counts = {};
    getOutward().forEach(function(r) {
      if (!r.date) return;
      var k = r.date.slice(0, 7);
      counts[k] = (counts[k] || 0) + 1;
    });
    var out = []; for (var k in counts) out.push({ month: k, count: counts[k] });
    out.sort(function(a, b) { return a.month < b.month ? -1 : 1; });
    return out;
  }
  function deleteInwardByMonth(monthKey) {
    var sheet = Database.getSheet(IN_SHEET);
    var rows = Database.getAll(IN_SHEET);
    var deleted = 0;
    for (var i = rows.length - 1; i >= 0; i--) {
      var d = String(rows[i][IC.DATE] || '');
      if (d.slice(0, 7) === monthKey) { sheet.deleteRow(i + 2); deleted++; }
    }
    return { success: true, deleted: deleted };
  }
  function deleteOutwardByMonth(monthKey) {
    var sheet = Database.getSheet(OUT_SHEET);
    var rows = Database.getAll(OUT_SHEET);
    var deleted = 0;
    for (var i = rows.length - 1; i >= 0; i--) {
      var d = String(rows[i][OC.DATE] || '');
      if (d.slice(0, 7) === monthKey) { sheet.deleteRow(i + 2); deleted++; }
    }
    return { success: true, deleted: deleted };
  }

  // ── Net stock + reorder status ───────────────────────────────────
  function getStockSummary() {
    var inward = getInward(), outward = getOutward();
    var totalIn = 0, totalOut = 0, totalCost = 0;
    inward.forEach(function(r) { totalIn += r.quantity; totalCost += r.total_amount; });
    outward.forEach(function(r) { totalOut += r.quantity; });
    var net = totalIn - totalOut;
    var settings = SettingsService.getAll();
    var minQty = Number(settings.lpg_min_cylinders) || 5;
    var avgCost = totalIn > 0 ? Math.round((totalCost / totalIn) * 100) / 100 : 0;
    return {
      totalInward: totalIn, totalOutward: totalOut, netStock: net,
      minQuantity: minQty, avgCostPerCylinder: avgCost,
      statusMessage: net <= minQty ? 'Order LPG Cylinders' : 'Good',
      statusOk: net > minQty
    };
  }

  // ── Monthly Inventory-vs-Bank comparison — always LIVE, never stored ──
  // "No. of Cylinders" / "Total Cost Inventory" = purchases (inward)
  // grouped by the calendar month of the invoice date — matching the
  // association's own template. "Total Cost IOB" = actual bank debits
  // that month tagged with the "LPG Inventory" expense category on the
  // Transaction Ledger. Recomputed fresh on every request, so it can
  // never go stale or disagree with the underlying records — there is
  // nothing to edit or manually refresh here by design.
  function getLiveComparison() {
    var inward = getInward();
    var months = {};
    inward.forEach(function(r) {
      if (!r.date) return;
      var key = r.date.slice(0, 7); // yyyy-MM
      if (!months[key]) months[key] = { key: key, cylinders: 0, costInventory: 0, costBank: 0 };
      months[key].cylinders += r.quantity;
      months[key].costInventory += r.total_amount;
    });

    // Match either "LPG Inventory" OR "LPG" — the seeded expense-category
    // list has both, which is exactly the kind of overlap that causes
    // real transactions to go uncounted depending on which one was picked
    // when tagging a debit. Until the categories are consolidated into
    // one, both spellings are treated as the same thing here.
    var LPG_EXPENSE_CATEGORIES = { 'LPG Inventory': 1, 'LPG': 1 };
    var txns = BankService.getAllTransactions({});
    txns.forEach(function(t) {
      if (!LPG_EXPENSE_CATEGORIES[t.expense_category]) return;
      var ymd = _parseAnyDate(t.date);
      if (!ymd) return;
      var key = ymd.y + '-' + (ymd.m < 10 ? '0' + ymd.m : ymd.m);
      if (!months[key]) months[key] = { key: key, cylinders: 0, costInventory: 0, costBank: 0 };
      months[key].costBank += Number(t.debit) || 0;
    });

    var out = [];
    for (var k in months) {
      var m = months[k];
      var parts = k.split('-');
      var diff = Math.round((m.costInventory - m.costBank) * 100) / 100;
      out.push({ year: Number(parts[0]), month: Number(parts[1]), cylinders: m.cylinders,
                 cost_inventory: Math.round(m.costInventory * 100) / 100,
                 cost_bank: Math.round(m.costBank * 100) / 100,
                 diff: diff, ok: Math.abs(diff) < 1 });
    }
    out.sort(function(a, b) { return (a.year * 100 + a.month) - (b.year * 100 + b.month); });
    return out;
  }

  function _parseAnyDate(s) {
    if (!s) return null;
    s = String(s).trim();
    var low = s.toLowerCase();
    var mn = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
    var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return { y:+m[1], m:+m[2] };
    m = low.match(/^(\d{1,2})[-\/ ]([a-z]{3})[a-z]*[-\/ ](\d{2,4})/);
    if (m && mn[m[2]]) { var y1=+m[3]; if (y1<100) y1+=2000; return { y:y1, m:mn[m[2]] }; }
    m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/);
    if (m) { var yy=+m[3]; if (yy<100) yy+=2000; var a=+m[1], b=+m[2], mo=b; if (b>12&&a<=12) mo=a; return { y:yy, m:mo }; }
    return null;
  }

  // ── Monthly Profit & Loss for the LPG business ──────────────────
  // Income  = what residents actually PAID that month (from meter readings)
  // Cost    = cylinders OUTWARD (used) that month × the running average
  //           cost per cylinder purchased to date (a transparent, simple
  //           costing method — not FIFO/LIFO — clearly labelled as such)
  // Profit/Loss = Income − Cost
  function getMonthlyPL(year, month) {
    var stock = getStockSummary();
    var avgCost = stock.avgCostPerCylinder;

    var outward = getOutward();
    var cylindersUsed = 0;
    outward.forEach(function(r) {
      if (!r.date) return;
      var key = r.date.slice(0, 7);
      if (key === year + '-' + (month < 10 ? '0' + month : month)) cylindersUsed += r.quantity;
    });
    var cylinderCost = Math.round(cylindersUsed * avgCost * 100) / 100;

    var readings = LPGReadingService.getMonth(year, month);
    var consumed = 0;
    for (var uid in readings) { consumed += Number(readings[uid].consumed) || 0; }

    // The reading for a month's consumption is only taken in the first
    // week of the FOLLOWING month, and residents pay around that same
    // time — so the money for January's gas is collected in February,
    // February's in March, and so on. "Collected" must therefore look at
    // next month's Payments ledger entries, not the same month's, to
    // correctly match income against the consumption it's paying for.
    var collMonth = month + 1, collYear = year;
    if (collMonth > 12) { collMonth = 1; collYear = year + 1; }
    var collectionMonthKey = collYear + '-' + (collMonth < 10 ? '0' + collMonth : collMonth);

    var collected = 0;
    PaymentsService.getAllPayments({ payment_type: 'LPG', month: collectionMonthKey }).forEach(function(p) {
      if (p.status === 'Rejected') return; // exact same exclusion rule as the Dashboard
      collected += Number(p.amount) || 0;
    });

    return {
      year: year, month: month,
      collectionYear: collYear, collectionMonth: collMonth,
      consumedNM3: Math.round(consumed * 1000) / 1000,
      collected: Math.round(collected * 100) / 100,
      cylindersUsed: cylindersUsed,
      avgCostPerCylinder: avgCost,
      cylinderCost: cylinderCost,
      profitLoss: Math.round((collected - cylinderCost) * 100) / 100
    };
  }

  // ── Bulk imports from the association's own Excel templates ────────
  // Each takes rows already converted to a 2D array (via the same
  // XLSX->CSV conversion used elsewhere), skips header/blank rows by
  // shape, and reuses the robust shared date parser — never JS's
  // ambiguous new Date(string).

  // Stock Inward: #, Invoice Number, Date, Cost/Cylinder, Quantity, Total
  function importInward(rows) {
    var imported = 0, warnings = [], errors = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r || r.length < 5) continue;
      var invoice = String(r[1] || '').trim();
      var dateRaw = r[2];
      var costRaw = String(r[3] || '').replace(/[₹,\s]/g, '');
      var qtyRaw  = String(r[4] || '').replace(/[,\s]/g, '');
      if (!invoice || isNaN(Number(qtyRaw)) || Number(qtyRaw) <= 0) continue; // not a data row

      // A row with real cylinders/cost must NEVER be dropped just because
      // its date cell couldn't be read (e.g. a typo'd year in the source
      // file) — that silently loses real stock from the count. Import it
      // with a blank date and flag it clearly for a manual fix instead.
      var date = parseFlexibleDate(dateRaw);
      if (!date) warnings.push('Invoice ' + invoice + ': date "' + dateRaw + '" could not be read — imported with no date; please set it via the ✏️ pencil.');
      try {
        addInward({ invoice_number: invoice, date: date, cost_per_cylinder: Number(costRaw), quantity: Number(qtyRaw), recorded_by: 'Import', allowBlankDate: true, skipYearCheck: true });
        imported++;
      } catch (e) { errors.push('Invoice ' + invoice + ': ' + e.message); }
    }
    return { success: true, imported: imported, errors: errors, warnings: warnings };
  }

  // Stock Outward: #, Date, Quantity Used
  function importOutward(rows) {
    var imported = 0, warnings = [], errors = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r || r.length < 3) continue;
      var idxRaw = String(r[0] || '').trim();
      if (!idxRaw || isNaN(Number(idxRaw))) continue; // header/blank row
      var dateRaw = r[1];
      var qtyRaw  = String(r[2] || '').replace(/[,\s]/g, '');
      if (isNaN(Number(qtyRaw)) || Number(qtyRaw) <= 0) continue;
      // Never drop real cylinder usage just because its date is unreadable.
      var date = parseFlexibleDate(dateRaw);
      if (!date) warnings.push('Row ' + idxRaw + ': date "' + dateRaw + '" could not be read — imported with no date; please set it via the ✏️ pencil.');
      try {
        addOutward({ date: date, quantity: Number(qtyRaw), notes: '', recorded_by: 'Import', allowBlankDate: true, skipYearCheck: true });
        imported++;
      } catch (e) { errors.push('Row ' + idxRaw + ': ' + e.message); }
    }
    return { success: true, imported: imported, errors: errors, warnings: warnings };
  }

  // Inventory v Bank: Month (name), No of Cylinders, Total Cost Inventory,
  // Total Cost IOB, Verify (ignored — recomputed from the two costs).
  // The sheet only names the month ("Jan"), so the calling year applies
  // to the whole file.
  return {
    ensureSheets:         ensureSheets,
    getInward:            getInward,
    getInwardMonthSummary: getInwardMonthSummary,
    getOutwardMonthSummary: getOutwardMonthSummary,
    deleteInwardByMonth:  deleteInwardByMonth,
    deleteOutwardByMonth: deleteOutwardByMonth,
    addInward:            addInward,
    deleteInward:         deleteInward,
    getOutward:           getOutward,
    addOutward:           addOutward,
    deleteOutward:        deleteOutward,
    getStockSummary:      getStockSummary,
    getLiveComparison:    getLiveComparison,
    importInward:         importInward,
    importOutward:        importOutward,
    getMonthlyPL:         getMonthlyPL
  };
})();
