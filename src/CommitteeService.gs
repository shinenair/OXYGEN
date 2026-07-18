// ═══════════════════════════════════════════════════════════════
// CommitteeService.gs — Office Bearers & EC Members of CDOA.
// Current committee (contacts, duties, assigned units), archived
// past committees, equal distribution of all units among members
// for payment follow-up, and a one-time seed built from the
// association's own four spreadsheets.
// ═══════════════════════════════════════════════════════════════
var CommitteeService = (function() {
  var SHEET = 'Committee';
  var HEADERS = ['member_id','status','term_label','sort_order','role','name','flat','phone','email','duty','allocate_units','units_assigned','created_at','updated_at'];
  var C = { ID:0, STATUS:1, TERM:2, SORT:3, ROLE:4, NAME:5, FLAT:6, PHONE:7, EMAIL:8, DUTY:9, ALLOC:10, UNITS:11, CREATED:12, UPDATED:13 };

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

  function _toObj(r) {
    return {
      member_id: String(r[C.ID] || ''), status: String(r[C.STATUS] || ''),
      term_label: String(r[C.TERM] || ''), sort_order: Number(r[C.SORT]) || 0,
      role: String(r[C.ROLE] || ''), name: String(r[C.NAME] || ''),
      flat: String(r[C.FLAT] || ''), phone: String(r[C.PHONE] || ''),
      email: String(r[C.EMAIL] || ''), duty: String(r[C.DUTY] || ''),
      allocate_units: String(r[C.ALLOC] || '') === 'Yes',
      units_assigned: String(r[C.UNITS] || '')
    };
  }

  function getAll() {
    ensureSheet();
    var rows = Database.getAll(SHEET);
    var all = rows.map(_toObj).filter(function(m) { return m.member_id; });
    all.sort(function(a, b) { return a.sort_order - b.sort_order || a.name.localeCompare(b.name); });
    var current = all.filter(function(m) { return m.status === 'Current'; });
    var archived = all.filter(function(m) { return m.status === 'Archived'; });
    return { current: current, archived: archived, isEmpty: all.length === 0 };
  }

  function addMember(d) {
    if (!d || !d.name) throw new Error('The member needs at least a name.');
    var sh = ensureSheet();
    var id = Database.generateId('CM');
    var now = new Date().toISOString();
    var maxSort = 0;
    Database.getAll(SHEET).forEach(function(r) {
      if (String(r[C.STATUS]) === 'Current') maxSort = Math.max(maxSort, Number(r[C.SORT]) || 0);
    });
    sh.appendRow([id, 'Current', String(d.term_label || ''), maxSort + 1,
                  String(d.role || 'Exec Mem'), String(d.name), String(d.flat || ''),
                  "'" + String(d.phone || ''), String(d.email || ''), String(d.duty || ''),
                  d.allocate_units ? 'Yes' : 'No', String(d.units_assigned || ''), now, now]);
    return { success: true, member_id: id };
  }

  function updateMember(id, d) {
    var found = Database.findByColumn(SHEET, C.ID, id);
    if (!found) throw new Error('Member not found: ' + id);
    var r = found.data;
    if (d.role !== undefined)  r[C.ROLE]  = String(d.role);
    if (d.name !== undefined)  r[C.NAME]  = String(d.name);
    if (d.flat !== undefined)  r[C.FLAT]  = String(d.flat);
    if (d.phone !== undefined) r[C.PHONE] = "'" + String(d.phone);
    if (d.email !== undefined) r[C.EMAIL] = String(d.email);
    if (d.duty !== undefined)  r[C.DUTY]  = String(d.duty);
    if (d.allocate_units !== undefined) r[C.ALLOC] = d.allocate_units ? 'Yes' : 'No';
    if (d.units_assigned !== undefined) r[C.UNITS] = String(d.units_assigned);
    if (d.term_label !== undefined) r[C.TERM] = String(d.term_label);
    r[C.UPDATED] = new Date().toISOString();
    Database.updateRow(SHEET, found.rowIndex, r);
    return { success: true };
  }

  function deleteMember(id) {
    var found = Database.findByColumn(SHEET, C.ID, id);
    if (!found) throw new Error('Member not found: ' + id);
    ensureSheet().deleteRow(found.rowIndex);
    return { success: true };
  }

  // Archives the WHOLE current committee under a term label (e.g.
  // "10 Oct 2024 → 15 Mar 2026") so a fresh committee can be entered.
  // Nothing is deleted — every archived committee stays viewable on the
  // Past Committees tab, exactly as the association's own historic
  // record keeps every term since Jan 2023.
  function archiveCurrent(termLabel) {
    if (!termLabel || !String(termLabel).trim()) throw new Error('An identifying term label is required (e.g. "10 Oct 2024 → 15 Mar 2026").');
    var sh = ensureSheet();
    var rows = Database.getAll(SHEET);
    var count = 0;
    var now = new Date().toISOString();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.STATUS]) === 'Current') {
        rows[i][C.STATUS] = 'Archived';
        rows[i][C.TERM] = String(termLabel).trim();
        rows[i][C.UPDATED] = now;
        Database.updateRow(SHEET, i + 2, rows[i]);
        count++;
      }
    }
    return { success: true, archived: count };
  }

  // Distributes EVERY unit equally among the current members flagged for
  // allocation, in consecutive blocks and member order — 136 units over
  // 9 members = eight members with 15 and one with 16, matching how the
  // association already divides them. Runs against the live Units sheet,
  // so it stays correct if the member count OR the unit count changes.
  function autoAllocateUnits() {
    var units = UnitsService.getAllUnits().map(function(u) { return u.unit_id; });
    if (!units.length) throw new Error('No units found to distribute.');
    var rows = Database.getAll(SHEET);
    var eligible = [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][C.STATUS]) === 'Current' && String(rows[i][C.ALLOC]) === 'Yes') {
        eligible.push({ idx: i, sort: Number(rows[i][C.SORT]) || 0 });
      }
    }
    if (!eligible.length) throw new Error('No current members are flagged for unit allocation. Edit a member and tick "Include in unit allocation" first.');
    eligible.sort(function(a, b) { return a.sort - b.sort; });

    var n = eligible.length;
    var base = Math.floor(units.length / n);
    var extra = units.length % n; // the LAST `extra` members get one more each,
    // matching the association's current sheet (the 16-unit block sits at the end)
    var now = new Date().toISOString();
    var cursor = 0;
    var summary = [];
    for (var e = 0; e < n; e++) {
      var take = base + (e >= n - extra ? 1 : 0);
      var slice = units.slice(cursor, cursor + take);
      cursor += take;
      var row = Database.getAll(SHEET)[eligible[e].idx];
      row[C.UNITS] = slice.join(', ');
      row[C.UPDATED] = now;
      Database.updateRow(SHEET, eligible[e].idx + 2, row);
      summary.push({ name: String(row[C.NAME]), count: slice.length, from: slice[0], to: slice[slice.length - 1] });
    }
    return { success: true, members: n, unitsDistributed: units.length, summary: summary };
  }

  // One-time seed from the association's own four spreadsheets (current
  // members, duty allocation, unit allocation, and the historic
  // office-bearers record). Refuses to run unless the sheet is empty —
  // it can never overwrite live edits.
  function seedInitial() {
    ensureSheet();
    if (!getAll().isEmpty) throw new Error('The Committee sheet already has data — seeding only runs on an empty sheet, so it can never overwrite live edits.');
    var sh = ensureSheet();
    var now = new Date().toISOString();
    var TERM_NOW = '10 Oct 2024 → Till Date';

    var seed = [
      // sort, role, name, flat, phone, email, duty, allocate, units
      [1, 'President', 'Mr.Sisilkumar A. Nayanan', 'B806', '971554559286', 'sisilkumar@gmail.com', 'ALL DEPARTMENT', 'Yes', 'B801, B802, B803, B804, B805, B806, B807, B808, B901, B902, B903, B904, B905, B906, B907, B908'],
      [2, 'Secretary', 'Mrs.Anjaly Mary Joseph', 'B408', '7994458702', 'anjalymary333@gmail.com', 'BANKING/FINANCE/GRIEVANCE OFFICER', 'Yes', 'A407, A408, A501, A502, A503, A504, A505, A506, A507, A508, A601, A602, A603, A604, A605'],
      [3, 'Treasurer', 'Mr.MIDHUN KUMAR M.G', 'A402', '9746970366', 'midhunmar23@gmail.com', 'VEHICLES/PARKING/STP', 'Yes', 'A805, A806, A807, A808, B101, B102, B103, B104, B105, B106, B107, B108, B201, B202, B203'],
      [4, 'Exec Mem', 'Mr.ASSIM SAIT A M', 'A204', '9744037779', 'assimsait@gmail.com', 'WATER/SAFETY/SECURITY OFFICER', 'Yes', 'A101, A102, A103, A104, A105, A106, A107, A108, A201, A202, A203, A204, A205, A206, A207'],
      [5, 'Exec Mem', 'Mrs.ELIZABETH DANIEL', 'B905', '9446364626', 'elizdanolivet@gmail.com', 'CONTRACTS/AMC/CULTURAL EVENTS', 'Yes', 'A208, A301, A302, A303, A304, A305, A306, A307, A308, A401, A402, A403, A404, A405, A406'],
      [6, 'Exec Mem', 'Mrs.Lakshmi V.', '', '', '', 'LADY STAFF/CREDAI', 'Yes', 'A606, A607, A608, A701, A702, A703, A704, A705, A706, A707, A708, A801, A802, A803, A804'],
      [7, 'Exec Mem', 'Mrs.REEBA THOMAS/JYOTHISH JACOB', 'B401', '9447087212', 'reebathomas77@gmail.com', 'OCCUPANCY/MOVE-IN-OUT', 'Yes', 'B204, B205, B206, B207, B208, B301, B302, B303, B304, B305, B306, B307, B308, B401, B402'],
      [8, 'Exec Mem', 'Mr.AJEEM P.', 'A807', '7510827863', 'connectajeem@gmail.com', 'LPG/ELECTRICITY/FIRE APPLIANCES', 'Yes', 'B403, B404, B405, B406, B407, B408, B501, B502, B503, B504, B505, B506, B507, B508, B601'],
      [9, 'Exec Mem', 'Mr.UNNIKRISHNAN V.T.', 'A404', '9495360329', 'unnivt@gmail.com', 'OXYGEN SUPPORT', 'Yes', 'B602, B603, B604, B605, B606, B607, B608, B701, B702, B703, B704, B705, B706, B707, B708'],
      [10, 'Caretaker', 'Mr. Sunil Kumar', '', '918129693313', 'cdoacaretaker@gmail.com', '', 'No', ''],
      [11, 'OFFICE', 'CDOA OFFICE', '', '', 'cdoabulletin@gmail.com', '', 'No', '']
    ];
    var hist = [
      ['28 Jan 2023 → 02 Apr 2023 (3 Months)', [['President','Clement Rosario'], ['Secretary','Bhagath M.'], ['Treasurer','Sheeja U.']]],
      ['02 Apr 2023 → 04 Jun 2023 (2 Months)', [['President','Ajeesh V.S.'], ['Secretary','Shanos David'], ['Treasurer','Sheeja U.']]],
      ['04 Jun 2023 → 12 Nov 2023 (5 Months)', [['President','Shibu D.S.'], ['Secretary','Shanos David'], ['Treasurer','Sheeja U./Krishnakumar']]],
      ['12 Nov 2023 → 10 Oct 2024 (11 Months)', [['President','ESK Das'], ['Secretary','Shine N.S.'], ['Treasurer','Midhunkumar M.G.']]]
    ];

    var out = [];
    seed.forEach(function(s) {
      out.push([Database.generateId('CM'), 'Current', TERM_NOW, s[0], s[1], s[2], s[3], "'" + s[4], s[5], s[6], s[7], s[8], now, now]);
    });
    var sortBase = 100;
    hist.forEach(function(h) {
      h[1].forEach(function(p, i) {
        out.push([Database.generateId('CM'), 'Archived', h[0], sortBase + i, p[0], p[1], '', "'", '', '', 'No', '', now, now]);
      });
      sortBase += 10;
    });

    sh.getRange(sh.getLastRow() + 1, 1, out.length, HEADERS.length).setValues(out);
    return { success: true, current: seed.length, archived: hist.length * 3 };
  }


  // The Payment Follow-up Board — the association's color-coded "TEAM
  // ALLOCATION" sheet, generated live for any month. Per current member
  // (in committee order): their assigned units × three fee rows
  // (MF / WMF / LPG), each unit cell carrying one status:
  //   paid    → green  #34A853  (a non-rejected payment exists)
  //   pending → red    #980000  (no payment, and the fee applies)
  //   uo      → yellow #FBBC04  (unoccupied / fee not applicable / UO marker)
  //   unreg   → black  #000000  (unit not registered with CDOA)
  // WMF and LPG are not applicable to Unoccupied units; LPG additionally
  // not applicable when the unit doesn't use association LPG.
  function getFollowUpBoard(year, month) {
    year = Number(year); month = Number(month);
    var key = year + '-' + (month < 10 ? '0' + month : String(month));
    var TYPES = ['Maintenance', 'Waste Management', 'LPG'];

    // Unit flags
    var unitFlags = {};
    UnitsService.getAllUnits().forEach(function(u) {
      unitFlags[String(u.unit_id).toUpperCase()] = {
        occupancy: String(u.occupancy || 'Occupied'),
        registration: String(u.registration || 'Registered'),
        lpg_mode: String(u.lpg_mode || 'Using')
      };
    });

    // Payment status per unit|type for the month
    var payStatus = {}; // 'UNIT|TYPE' -> 'paid' | 'uo'
    PaymentsService.getAllPayments(null).forEach(function(p) {
      if (String(p.month) !== key) return;
      if (p.status === 'Rejected') return;
      var k = String(p.unit_id).toUpperCase() + '|' + p.payment_type;
      if (p.status === 'UO') { if (!payStatus[k]) payStatus[k] = 'uo'; }
      else payStatus[k] = 'paid';
    });

    function statusFor(unit, type) {
      var f = unitFlags[unit] || { occupancy: 'Occupied', registration: 'Registered', lpg_mode: 'Using' };
      if (f.registration === 'Not-registered') return 'unreg';
      var k = unit + '|' + type;
      if (payStatus[k]) return payStatus[k];
      if (type !== 'Maintenance' && f.occupancy === 'Unoccupied') return 'uo';
      if (type === 'LPG' && f.lpg_mode !== 'Using') return 'uo';
      return 'pending';
    }

    var all = getAll();
    var members = [];
    var maxUnits = 0;
    var pendingTotals = { 'Maintenance': 0, 'Waste Management': 0, 'LPG': 0 };
    all.current.forEach(function(m) {
      if (!m.allocate_units) return;
      var units = m.units_assigned ? m.units_assigned.split(',').map(function(u) { return u.trim().toUpperCase(); }).filter(Boolean) : [];
      if (!units.length) return;
      maxUnits = Math.max(maxUnits, units.length);
      var rows = {};
      var perType = {};    // this member's defaulter count per fee type
      var memberTotal = 0; // and across all three types
      TYPES.forEach(function(ty) {
        perType[ty] = 0;
        rows[ty] = units.map(function(u) {
          var s = statusFor(u, ty);
          if (s === 'pending') { pendingTotals[ty]++; perType[ty]++; memberTotal++; }
          return { unit: u, s: s };
        });
      });
      members.push({ name: m.name, role: m.role, unitCount: units.length, rows: rows,
                     perTypePending: perType, totalPending: memberTotal });
    });

    var mnFull = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
    return { year: year, month: month, monthName: mnFull[month], monthKey: key,
             generatedDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMMMyyyy'),
             members: members, maxUnits: maxUnits, pendingTotals: pendingTotals,
             types: TYPES };
  }

  return {
    ensureSheet:    ensureSheet,
    getAll:         getAll,
    addMember:      addMember,
    updateMember:   updateMember,
    deleteMember:   deleteMember,
    archiveCurrent: archiveCurrent,
    autoAllocateUnits: autoAllocateUnits,
    seedInitial:    seedInitial,
    getFollowUpBoard: getFollowUpBoard
  };
})();
