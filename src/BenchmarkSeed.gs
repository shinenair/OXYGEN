// ============================================================
// BenchmarkSeed.gs — "Dec-2024 Ultimate Truth" benchmark seed
//
// One-time authoritative baseline for 2024, from the association's three
// corrected truth sheets (Maintenance, Waste Management, LPG). Purely
// ADDITIVE: it fills in paid records the live ledger is MISSING and reports
// every conflict for manual review — it never deletes or overwrites, so
// bank-posted records and their links are preserved. Re-running is safe
// (idempotent): records already present are skipped.
//
// Conventions:
//   • Maintenance = flat Rs.1,500/month (2024); Waste = flat Rs.170/month.
//   • LPG is the Dec-2024 cycle only (month key 2024-12, our consumption-month
//     convention); the payment's real date is stored as submitted_at.
//   • Lift AMC + per-unit Extras are booked as Miscellaneous income (Dec 2024)
//     with descriptive notes, so annual totals match without polluting the
//     monthly grids.
// ============================================================

var BenchmarkSeed = (function() {
  var YEAR = 2024, MF_AMT = 1500, WMF_AMT = 170, LPG_MONTH = '2024-12';

  // ── Embedded benchmark (generated from the truth sheets) ──
  var MF        = {A101:'111111111111',A102:'111111111111',A103:'111111111111',A104:'111111111111',A105:'111111111111',A106:'111111111111',A107:'111111111111',A108:'111111111111',A201:'000000000000',A202:'111111111111',A203:'111111111111',A204:'111111111111',A205:'111111111111',A206:'111111111111',A207:'111111111111',A208:'111111111111',A301:'111111111111',A302:'111111111111',A303:'111111111111',A304:'111111111111',A305:'111111111111',A306:'111111111111',A307:'111111111111',A308:'111111111111',A401:'111111111111',A402:'111111111111',A403:'111111111111',A404:'111111111111',A405:'111111111111',A406:'111111111111',A407:'111111111111',A408:'111111111111',A501:'111111111111',A502:'111111111111',A503:'111111111111',A504:'111111111111',A505:'111111111111',A506:'111111111111',A507:'111111111111',A508:'111111111111',A601:'111111111111',A602:'000000000000',A603:'111111111111',A604:'111111111111',A605:'111111111111',A606:'111111111111',A607:'111111111111',A608:'111111111111',A701:'111111111111',A702:'111111111111',A703:'111111111111',A704:'111111111111',A705:'111111111111',A706:'111111111111',A707:'111111111111',A708:'111111111111',A801:'111111111100',A802:'111111111111',A803:'111111111111',A804:'111111111111',A805:'111111111111',A806:'111111111111',A807:'111111111111',A808:'000000000000',B101:'111111111111',B102:'111111111111',B103:'111111111111',B104:'111111111111',B105:'111111111111',B106:'111111111111',B107:'111111111111',B108:'111111111111',B201:'111111111111',B202:'111111111111',B203:'111111111111',B204:'111111111111',B205:'111111111111',B206:'111111111111',B207:'111111111111',B208:'111111111111',B301:'111111111111',B302:'111111111111',B303:'111111111111',B304:'111111111111',B305:'111111111111',B306:'111111111111',B307:'111111111111',B308:'111111111111',B401:'111111111111',B402:'000000000000',B403:'111111111111',B404:'111111111111',B405:'111111111111',B406:'111111111111',B407:'111111111111',B408:'111111111111',B501:'111111111111',B502:'111111111111',B503:'111111111111',B504:'111111111111',B505:'111111111111',B506:'111111111111',B507:'111111111111',B508:'111111111111',B601:'111111111111',B602:'111111111111',B603:'111111111111',B604:'111111111111',B605:'111111111111',B606:'000000000000',B607:'111111111111',B608:'111111111111',B701:'111111111111',B702:'111111111111',B703:'111111111111',B704:'111111111111',B705:'111111111111',B706:'111111111111',B707:'111111111111',B708:'111111111111',B801:'111111111111',B802:'111111111111',B803:'111111111111',B804:'111111111111',B805:'111111111111',B806:'111111111111',B807:'111111111111',B808:'111111111111',B901:'111111111111',B902:'111111111111',B903:'111111111111',B904:'111111111111',B905:'111111111111',B906:'111111111111',B907:'111111111111',B908:'111111111111'};
  var LIFT_AMC  = {A101:1550,A102:1550,A103:1550,A104:1550,A105:1550,A106:1550,A107:1550,A108:1550,A202:1550,A203:1550,A204:1550,A205:1550,A206:1550,A207:1550,A208:1550,A301:1550,A302:1550,A303:1550,A304:1550,A305:1550,A306:1550,A307:1550,A401:1550,A402:1550,A403:1550,A404:1550,A405:1550,A406:1550,A407:1550,A408:1550,A501:1550,A502:1550,A503:1550,A504:1550,A505:1550,A506:1550,A507:1550,A508:1550,A601:1550,A603:1550,A604:1550,A605:1550,A606:1550,A607:1550,A608:1500,A702:1550,A703:1550,A704:1550,A705:1550,A706:1550,A707:1550,A708:1550,A802:1550,A803:1550,A804:1550,A805:1550,A806:1550,A807:1550,A808:1550,B101:1550,B102:1550,B103:1550,B104:1550,B105:1550,B106:1550,B107:1550,B108:1550,B201:1550,B202:1550,B203:1550,B204:1550,B205:1550,B206:1550,B207:1550,B208:1550,B301:1550,B302:1550,B303:1550,B304:1550,B305:1550,B306:1550,B307:1550,B308:1550,B401:1550,B403:1550,B404:1550,B405:1550,B406:1550,B407:1550,B408:1550,B501:1550,B502:1550,B504:1550,B505:1550,B506:1550,B507:1550,B508:1550,B601:1550,B602:1550,B604:1550,B605:1550,B607:1550,B608:1550,B701:1550,B702:1550,B704:1550,B705:1550,B706:1550,B707:1550,B708:1550,B801:1550,B802:1550,B803:1550,B804:1550,B805:1550,B806:1550,B807:1550,B808:1550,B901:1550,B902:1550,B904:1550,B905:1550,B906:1550,B907:1550,B908:1550};
  var MF_EXTRA  = {A106:50,A606:6000,B101:3000};
  var WMF       = {A101:'PPPPPPPPPPPP',A102:'PPPPPPPPPPPP',A103:'PPPPPPPPPPPP',A104:'PPPPPPPPPPPP',A105:'UUPPPPUUUPPP',A106:'PPPPPPPPPPPP',A107:'PPPPPPPPPPPP',A108:'PPPUPPPPPPPP',A201:'UUUUUUUUUUUU',A202:'UPUPPPPUUUPP',A203:'PPPPPPPPPPPP',A204:'PPPPPPPPPPPP',A205:'PPPPPPPPPPPP',A206:'PPPPPPPPPPPP',A207:'PPPUPPPPPPPP',A208:'PPPPPPPPPPPP',A301:'PPPUUPPPPPPP',A302:'UUUUUUUUUUUU',A303:'PPPPPPPPPPPP',A304:'PPPUPPPPPPPP',A305:'PPPPPPPPPPPP',A306:'PPPPPPPPPPPP',A307:'PPPPPPPPPPPP',A308:'UUUUUUUUUUUU',A401:'UUUUUUUUUUUU',A402:'PPPPPPPPPPPP',A403:'UUPUPPPPPPPP',A404:'PPPPPPPPPPPP',A405:'UUUUUPUUUPPP',A406:'PPPPPPPPPPPP',A407:'PPPPPPPPPPPP',A408:'PPPPPPPPPPPP',A501:'PPPPPPPPPPPP',A502:'PPPPPPPPPPPP',A503:'PPPPPPPPPPPP',A504:'UUUPPPPPPPPP',A505:'PPPPPPPUUUUU',A506:'PPPPPPPPPPPP',A507:'PPPPPPUUUUUU',A508:'PPPPPPPPPPPP',A601:'UUUUUUUUUUUU',A602:'UUUUUUUUUUUU',A603:'UUPPPPPPPPPP',A604:'PPPPPPPPPPPP',A605:'UUUUUUUUUUUU',A606:'UUUUUPPPPPPP',A607:'UUUUUUUUUPPP',A608:'PPPPPPPPPPPP',A701:'PPPPPPPP....',A702:'PPPPPPPPPPPP',A703:'PPPPPPPPPPPP',A704:'UPPPPPPPPPPP',A705:'PPPPPPPPPPPP',A706:'PPPPPPPUPPPP',A707:'PPPPPPPPPPP.',A708:'PPPPPPPPPPPP',A801:'UUUUUUUUUUUU',A802:'UUPPPPPPPPPP',A803:'UUUUUPPPPPPP',A804:'PPPPPPPPPPPP',A805:'PPUUUUUUUUUU',A806:'PPPPPPPPUUPP',A807:'PPPUUPPPPPPP',A808:'PPPPPPPPPPPP',B101:'UUUUUUUUPPPP',B102:'UPPPPPPPPPPP',B103:'PUPPPPPPPPPP',B104:'PPPPPPPPPPPP',B105:'PPPPPPPPPPPP',B106:'PPPPUUPPPPPP',B107:'UUUUUUUUUUUU',B108:'UUUUUUUUUUUU',B201:'PPPPPPPPPPPP',B202:'UUUUUUUUUUUU',B203:'PPPPPPPPPPPP',B204:'PPPPPPPPPPPP',B205:'PPPPPPPPPPPP',B206:'PPPPPPPPPPPP',B207:'UUUUUUUUUUUU',B208:'PPPPPPPPPPPP',B301:'UUUUUUPUUUUU',B302:'PPPPPPPPPPPP',B303:'PPPPPPPPPPPP',B304:'PPPPPPPPPPPP',B305:'UUUUUUUUUUUU',B306:'UUPPPPPPPPPP',B307:'PPPUUUPPPPPP',B308:'UUUUUUPPPUUU',B401:'UUUPUUUPUUUU',B402:'UUUUUUUUUUUU',B403:'PPPPPPPPPPPP',B404:'UUUUUUUUUUUU',B405:'PPPPPPPPPPPP',B406:'UUUUUUUUUUPP',B407:'PUUUUUUUPPPP',B408:'PPPPPPPPPPPP',B501:'PPPPPPPPPPPP',B502:'PPPPPPPPPPPP',B503:'UUUUUUUUUUUU',B504:'UUUUUPUUUUUU',B505:'UUPPPPPPPPPP',B506:'PPPPPPPPPPPP',B507:'PPPPPPPPPPPP',B508:'PPPPPPPPPPPP',B601:'PPPPPPPPPUPP',B602:'PPPPPPPPPPPP',B603:'UUUUUUUUUUUU',B604:'PPPPPPPPPPPP',B605:'PPPPPPPPPPPP',B606:'UUUUUUUUUUUU',B607:'PUUUUUUPPPPP',B608:'UUUPPPPPPPPP',B701:'UUPPPPPPPPPP',B702:'PPPPPPPPPPUU',B703:'UUUUUUUUUUUU',B704:'UUUUUUUUUUUU',B705:'PPPPPPPPPPPP',B706:'PPPPPPPPPPPP',B707:'PPPPPPPPPPUP',B708:'UUUUUPUPPUUU',B801:'PPPPPPPPPPPU',B802:'PPPPPPPPPPPP',B803:'PPPPPPPPPPPP',B804:'UUUUUUUUUUUU',B805:'UUUPPPPPPPPP',B806:'UUPUUPUPPPPP',B807:'PPPPPPPPPPPP',B808:'PPPPPPPPPPPP',B901:'PPPPPPPPPPPP',B902:'UUUUPPPPPPPP',B903:'PPPPPPPPPPPP',B904:'UUUUUUUUUUUU',B905:'PPPPPUUUUUUU',B906:'UUUUUUUUUUUU',B907:'PPPPPPPPPPPP',B908:'PPPPPPPPPPPP'};
  var WMF_EXTRA = {A101:150,A306:510,B705:170};
  var LPG       = {A102:['2025-01-06',25.0],A103:['2025-01-06',238.0],A104:['2024-12-30',148.0],A105:['2025-01-31',684.0],A106:['2025-01-14',205.0],A107:['2025-01-07',771.0],A108:['2025-01-02',200.0],A202:['2025-01-07',314.0],A203:['2025-01-15',12.0],A204:['2025-01-02',418.0],A205:['2025-01-14',97.0],A206:['2025-01-04',164.0],A207:['2025-01-08',214.0],A208:['2025-01-06',327.0],A301:['2025-01-02',173.0],A303:['2025-01-02',243.0],A304:['2025-01-10',37.0],A305:['2025-01-17',523.0],A307:['2025-01-25',167.0],A402:['2025-01-07',105.0],A403:['2025-01-03',270.0],A404:['2025-01-02',140.0],A405:['2025-01-04',239.0],A406:['2025-01-15',131.0],A407:['2025-01-03',792.0],A408:['2025-01-03',135.0],A501:['2025-01-04',429.0],A503:['2025-01-05',269.0],A504:['2025-01-05',174.0],A506:['2025-01-11',282.0],A508:['2025-02-11',389.0],A603:['2025-01-18',137.0],A604:['2025-01-02',19.0],A606:['2025-01-11',321.0],A607:['2025-01-02',490.0],A608:['2025-01-03',20.0],A702:['2025-01-15',89.0],A703:['2025-01-18',7.0],A704:['2025-01-01',313.0],A705:['2025-01-14',9.0],A706:['2025-01-01',548.0],A802:['2025-01-02',193.0],A803:['2025-01-04',205.0],A804:['2025-01-02',245.0],A806:['2025-01-07',285.0],A808:['2025-01-01',460.0],B101:['2025-01-14',41.0],B102:['2025-01-04',207.0],B103:['2025-01-02',286.0],B104:['2025-01-04',194.0],B105:['2025-01-06',448.0],B106:['2025-01-04',178.0],B201:['2025-01-07',40.0],B203:['2025-01-08',180.0],B205:['2025-01-02',730.0],B208:['2025-01-15',283.0],B303:['2025-01-04',283.0],B304:['2025-01-18',33.0],B306:['2025-01-23',49.0],B307:['2025-01-18',113.0],B403:['2025-01-17',63.0],B405:['2025-01-06',699.0],B406:['2025-01-07',782.0],B407:['2025-01-02',870.0],B408:['2025-01-12',611.0],B501:['2025-01-02',194.0],B502:['2025-01-08',154.0],B505:['2025-01-01',531.0],B506:['2025-01-07',329.0],B507:['2025-01-07',99.0],B508:['2025-01-04',617.0],B602:['2025-01-02',238.0],B604:['2025-01-04',254.0],B605:['2025-01-01',517.0],B607:['2025-03-19',414.0],B608:['2025-01-15',133.0],B701:['2025-01-10',269.0],B705:['2025-01-02',29.0],B706:['2025-01-04',324.0],B707:['2025-01-18',218.0],B708:['2025-01-23',253.0],B803:['2025-01-03',249.0],B807:['2025-01-08',167.0],B808:['2025-01-11',32.0],B901:['2025-01-04',78.0],B907:['2025-01-08',551.0],B908:['2025-01-04',23.0]};  // unit -> [paymentDate 'YYYY-MM-DD', amount]

  function _pad(n) { return n < 10 ? '0' + n : String(n); }
  function _mk(m) { return YEAR + '-' + _pad(m); }

  // Every record the benchmark WANTS to exist.
  function _wanted() {
    var out = [];
    var u, i;
    for (u in MF) {
      var mask = MF[u];
      for (i = 0; i < 12; i++) if (mask.charAt(i) === '1')
        out.push({ unit: u, type: 'Maintenance', month: _mk(i + 1), amount: MF_AMT, note: 'BENCHMARK MF 2024', key: u + '|Maintenance|' + _mk(i + 1) });
    }
    for (u in WMF) {
      var wm = WMF[u];
      for (i = 0; i < 12; i++) if (wm.charAt(i) === 'P')
        out.push({ unit: u, type: 'Waste Management', month: _mk(i + 1), amount: WMF_AMT, note: 'BENCHMARK WMF 2024', key: u + '|Waste Management|' + _mk(i + 1) });
    }
    for (u in LPG) {
      out.push({ unit: u, type: 'LPG', month: LPG_MONTH, amount: LPG[u][1], date: LPG[u][0], note: 'BENCHMARK LPG Dec2024', key: u + '|LPG|' + LPG_MONTH });
    }
    // Annual extras as Miscellaneous (Dec 2024), deduped by note.
    for (u in LIFT_AMC)  out.push({ unit: u, type: 'Miscellaneous', month: _mk(12), amount: LIFT_AMC[u],  note: 'BENCHMARK Lift AMC 2024',  misc: true, key: u + '|misc|BENCHMARK Lift AMC 2024' });
    for (u in MF_EXTRA)  out.push({ unit: u, type: 'Miscellaneous', month: _mk(12), amount: MF_EXTRA[u],  note: 'BENCHMARK MF Extra 2024',  misc: true, key: u + '|misc|BENCHMARK MF Extra 2024' });
    for (u in WMF_EXTRA) out.push({ unit: u, type: 'Miscellaneous', month: _mk(12), amount: WMF_EXTRA[u], note: 'BENCHMARK WMF Extra 2024', misc: true, key: u + '|misc|BENCHMARK WMF Extra 2024' });
    return out;
  }

  // Index the live ledger: which unit|type|month already have a real payment,
  // and which unit+note Misc records already exist (for idempotent extras).
  function _liveIndex() {
    var rows = Database.getAll('Payments');
    var present = {}, miscNotes = {}, live2024 = { Maintenance: {}, 'Waste Management': {}, LPG: {} };
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var st = String(r[7] || '');
      if (st === 'Rejected') continue;
      var unit = String(r[1] || '').toUpperCase();
      var type = String(r[3] || '');
      var month = _normMonthLocal(r[5]);
      if (!unit || !type || !month) continue;
      if (st === 'Verified' || st === 'Pending' || st === 'UO') {
        present[unit + '|' + type + '|' + month] = true;
      }
      var note = String(r[8] || '');
      if (type === 'Miscellaneous' && note) miscNotes[unit + '||' + note] = true;
      // Track live 2024 records of the schedulable types for conflict detection.
      if (month.indexOf('2024-') === 0 && live2024[type]) {
        if (st === 'Verified' || st === 'Pending') live2024[type][unit + '|' + month] = (Number(r[4]) || 0);
      }
    }
    return { present: present, miscNotes: miscNotes, live2024: live2024 };
  }

  function _normMonthLocal(v) {
    if (v === undefined || v === null || v === '') return '';
    if (v instanceof Date) return v.getFullYear() + '-' + _pad(v.getMonth() + 1);
    var s = String(v).trim().replace(/^\'/, '');
    var m = s.match(/^(\d{4})-(\d{1,2})/);
    if (m) return m[1] + '-' + _pad(Number(m[2]));
    return '';
  }

  // What the benchmark expects a unit to have paid in 2024 (for conflict check).
  function _benchmarkExpects(unit, type, month) {
    var mi = Number(month.split('-')[1]) - 1;
    if (type === 'Maintenance')      return MF[unit]  ? MF[unit].charAt(mi)  === '1' : false;
    if (type === 'Waste Management') return WMF[unit] ? WMF[unit].charAt(mi) === 'P' : false;
    if (type === 'LPG')              return month === LPG_MONTH ? !!LPG[unit] : true; // only Dec-2024 is defined
    return true;
  }

  function preview() {
    var wanted = _wanted();
    var idx = _liveIndex();
    var cat = {};  // category -> {toAdd, present, addAmt}
    function bump(c, field, amt) { if (!cat[c]) cat[c] = { toAdd: 0, present: 0, addAmt: 0 }; cat[c][field]++; if (field === 'toAdd') cat[c].addAmt += amt; }

    for (var i = 0; i < wanted.length; i++) {
      var w = wanted[i];
      var c = w.note;
      var exists = w.misc ? idx.miscNotes[w.unit + '||' + w.note] : idx.present[w.key];
      bump(c, exists ? 'present' : 'toAdd', w.amount);
    }

    // Conflicts: live 2024 records the benchmark says should NOT be paid.
    var conflicts = [];
    ['Maintenance', 'Waste Management', 'LPG'].forEach(function(type) {
      var lv = idx.live2024[type];
      for (var k in lv) {
        var unit = k.split('|')[0], month = k.split('|')[1];
        if (type === 'LPG' && month !== LPG_MONTH) continue; // only Dec-2024 LPG is benchmarked
        if (!_benchmarkExpects(unit, type, month)) {
          conflicts.push({ unit: unit, type: type, month: month, amount: lv[k] });
        }
      }
    });
    conflicts.sort(function(a, b) { return (a.unit + a.type + a.month) < (b.unit + b.type + b.month) ? -1 : 1; });

    var totalAdd = 0, cats = [];
    for (var cc in cat) { cats.push({ label: cc, toAdd: cat[cc].toAdd, present: cat[cc].present, addAmt: Math.round(cat[cc].addAmt) }); totalAdd += cat[cc].addAmt; }
    cats.sort(function(a, b) { return a.label < b.label ? -1 : 1; });

    return {
      categories: cats,
      totalToAdd: cats.reduce(function(s, c) { return s + c.toAdd; }, 0),
      totalAddAmount: Math.round(totalAdd),
      conflictCount: conflicts.length,
      conflicts: conflicts.slice(0, 200)
    };
  }

  function commit() {
    UsersService.requireAdmin();
    var wanted = _wanted();
    var idx = _liveIndex();
    var now = new Date().toISOString();
    var newRows = [];
    var added = {};
    for (var i = 0; i < wanted.length; i++) {
      var w = wanted[i];
      var exists = w.misc ? idx.miscNotes[w.unit + '||' + w.note] : idx.present[w.key];
      if (exists) continue;
      // Guard against duplicates within this batch too.
      var dk = w.misc ? (w.unit + '||' + w.note) : w.key;
      if (added[dk]) continue;
      added[dk] = true;
      newRows.push([
        Database.generateId('PAY'),
        w.unit, '', w.type, Number(w.amount), "'" + w.month, '',
        'Verified', w.note, (w.date || now), now, 'Benchmark Dec2024'
      ]);
    }
    if (newRows.length) {
      var sheet = Database.getSheet('Payments');
      if (sheet.getMaxColumns() < 12) sheet.insertColumnsAfter(sheet.getMaxColumns(), 12 - sheet.getMaxColumns());
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 12).setValues(newRows);
      SpreadsheetApp.flush();
    }
    return { success: true, added: newRows.length };
  }

  return { preview: preview, commit: commit };
})();
