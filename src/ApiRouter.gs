// ============================================================
// ApiRouter.gs — POST Request Router
// ============================================================

// Shared by every service that writes year-specific data (Bank Statement,
// Payments, LPG Readings, LPG Inventory). The Dashboard's Universal Year
// Selector defines which year the whole app is "currently working in" —
// editing that year works normally for anyone with the usual permissions.
// Editing any OTHER year requires Administrator rights, AND requires the
// client to have already shown a confirmation prompt (confirmedHistorical
// must be explicitly true) — never just inferred, so a stray or forged
// request can't silently slip through as "already confirmed."
function checkYearEditable(recordYear, confirmedHistorical) {
  var activeYear = SettingsService.getActiveYear();
  var y = Number(recordYear);
  if (!y || y === activeYear) return; // current year — no restriction
  var email = Session.getActiveUser().getEmail();
  if (UsersService.getRole(email) !== 'admin') {
    throw new Error('This record is from ' + y + ', not the active year (' + activeYear + '). Only an Administrator can edit historical data.');
  }
  if (confirmedHistorical !== true) {
    throw new Error('CONFIRM_HISTORICAL_EDIT'); // client shows a confirmation prompt and retries with the flag set
  }
}

var ApiRouter = (function() {

  // Picks which bank account's service a bank.* route operates on.
  // account '2' = the older IOB LPG account; anything else (including
  // absent, so every existing caller keeps working unchanged) = the
  // main IOB account.
  function _bankSvc(data) {
    return (data && String(data.account) === '2') ? Bank2Service : BankService;
  }

  // Actions reachable WITHOUT being a recognized member — only the
  // PIN bootstrap itself (how an unidentified visitor becomes
  // identified) and sign-out. Everything else requires membership.
  var PUBLIC_ACTIONS = { 'auth.verifyPin': 1, 'auth.signOut': 1 };

  // MEMBERS ONLY — the server-side gate matching doGet's locked page.
  // A caller counts as a member when Google reveals an email that's on
  // the Users sheet (the owner always qualifies), or when they hold a
  // PIN-verified identity in this session's cache. Everyone else is
  // refused EVERY route, so the data stays private even for someone
  // who bypasses the page and calls the API directly.
  function _requireMember(action) {
    var email = '';
    try { email = Session.getActiveUser().getEmail(); } catch (e) {}
    if (email && UsersService.getRole(email)) return;
    var cached = AuthService.getCachedIdentity();
    if (cached && cached.email) return;
    throw new Error('Access denied — OXYGEN is a private application of the Confident Daffodils Owners Association. Reload the page and enter your PIN, or contact an Administrator.');
  }

  // ── Response cache ─────────────────────────────────────────────
  // The slow primitive here is reading whole sheets, so the heaviest
  // read routes cache their finished response for a short window.
  // Correctness rules:
  //   • Only the whitelisted routes below are ever cached — all serve
  //     the same member-wide data to every allowed user, none are
  //     role-filtered, and auth (_requireMember) still runs on every
  //     request BEFORE the cache is consulted.
  //   • ANY successful non-read action bumps a version stamp that is
  //     part of every cache key, so an edit invalidates everything at
  //     once — a member can never read pre-edit data after a write.
  //   • The short TTL (3 min) additionally bounds staleness from edits
  //     made directly in the spreadsheet, which the app can't see.
  //   • Cache failures of any kind fall through to a normal dispatch.
  var CACHEABLE = {
    'units.getAll': 1, 'owners.getAll': 1, 'tenants.getAll': 1,
    'payments.getAll': 1, 'units.profile': 1, 'dashboard.data': 1,
    'occupancy.getYearViewForAll': 1, 'lpg.getMonth': 1, 'bank.getAll': 1
  };
  var CACHE_TTL_SEC = 180;
  var CHUNK_CHARS = 30000;      // ≤90KB even if every char were 3 UTF-8 bytes
  var MAX_CACHE_CHARS = 800000; // beyond this, just don't cache

  function _rcVer(cache) {
    return cache.get('rc_ver') || '0';
  }
  function _rcBump() {
    try { CacheService.getScriptCache().put('rc_ver', String(Date.now()), 21600); } catch (e) {}
  }
  function _rcKey(cache, action, data) {
    var sig = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(data || {})));
    return 'rc_' + _rcVer(cache) + '_' + action + '_' + sig;
  }
  function _rcGet(cache, key) {
    var n = cache.get(key + '_n');
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < Number(n); i++) keys.push(key + '_' + i);
    var got = cache.getAll(keys);
    var out = '';
    for (var j = 0; j < Number(n); j++) {
      var part = got[key + '_' + j];
      if (part === undefined || part === null) return null; // a chunk expired — treat as miss
      out += part;
    }
    return out;
  }
  function _rcPut(cache, key, str) {
    if (str.length > MAX_CACHE_CHARS) return;
    var n = Math.ceil(str.length / CHUNK_CHARS) || 1;
    var obj = {};
    for (var i = 0; i < n; i++) obj[key + '_' + i] = str.substr(i * CHUNK_CHARS, CHUNK_CHARS);
    obj[key + '_n'] = String(n);
    cache.putAll(obj, CACHE_TTL_SEC);
  }
  // Read = never bumps the version. Conservative on purpose: anything
  // not clearly a read is treated as a write (worst case: a needless
  // invalidation, never stale data).
  function _isReadAction(a) {
    return CACHEABLE[a] === 1 || PUBLIC_ACTIONS[a] === 1 ||
      /\.(get|list|history|profile|stats|reconcile|search)/.test(String(a));
  }

  function route(action, data) {
    try {
      if (!PUBLIC_ACTIONS[action]) _requireMember(action);

      if (CACHEABLE[action] === 1) {
        try {
          var cache = CacheService.getScriptCache();
          var key = _rcKey(cache, action, data);
          var hit = _rcGet(cache, key);
          if (hit) return JSON.parse(hit);
          var freshResult = _dispatch(action, data);
          var wrapped = { success: true, data: freshResult };
          try { _rcPut(cache, key, JSON.stringify(wrapped)); } catch (ePut) {}
          return wrapped;
        } catch (eCache) {
          // Any cache-layer surprise: fall through to a plain dispatch.
        }
      }

      var result = _dispatch(action, data);
      if (!_isReadAction(action)) _rcBump();
      return { success: true, data: result };
    } catch (err) {
      Logger.log('ApiRouter error [' + action + ']: ' + err.message);
      return { success: false, error: err.message };
    }
  }

  // Profile photos render straight from Drive in the VIEWER's browser,
  // so the file itself must be link-viewable. Photos uploaded through
  // the app already are (uploadOwnerPhoto sets it), but photos added by
  // pasting a private Drive link were visible only to the Administrator
  // who owns the file. This repairs sharing on the records being served,
  // so a unit's photos become viewable to everyone the first time its
  // profile is opened. A cache entry keeps it to one Drive call per
  // file per 6 hours; files the script's account can't modify (foreign
  // links) are left untouched.
  function _ensurePhotoSharing(records) {
    var cache = CacheService.getScriptCache();
    (records || []).forEach(function(r) {
      var link = r && r.profile_picture;
      if (!link || String(link).indexOf('drive.google') === -1) return;
      var m = String(link).match(/[-\w]{25,}/);
      if (!m) return;
      var key = 'photoshare_' + m[0];
      if (cache.get(key)) return;
      try {
        DriveApp.getFileById(m[0]).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) { /* not this account's file — nothing we can do */ }
      cache.put(key, '1', 21600);
    });
    return records;
  }

  // Actions only Administrators may perform — enforced on the SERVER,
  // so hiding menus in the browser is cosmetic, not the security boundary.
  var ADMIN_ACTIONS = {
    'admin.resetPaymentData': 1,
    'users.list': 1, 'users.add': 1, 'users.setRole': 1, 'users.remove': 1,
    'lpg.setRate': 1, 'lpg.deleteRate': 1, 'lpg.importReadings': 1,
    'lpgInv.addInward': 1, 'lpgInv.deleteInward': 1, 'lpgInv.addOutward': 1, 'lpgInv.deleteOutward': 1,
    'lpgInv.importInward': 1, 'lpgInv.importOutward': 1,
    'admin.deleteBankByMonth': 1, 'admin.deleteBank2ByMonth': 1, 'admin.deleteLpgReadingsByMonth': 1,
    'admin.deleteLpgInwardByMonth': 1, 'admin.deleteLpgOutwardByMonth': 1,
    'admin.deleteAllOwners': 1, 'admin.deleteAllTenants': 1,
    'admin.exportBackup': 1, 'admin.restoreBackup': 1, 'admin.wipeForNewFacility': 1,
    'settings.setActiveYear': 1,
    'expensePatterns.add': 1, 'expensePatterns.delete': 1,
    'caretaker.deleteByMonth': 1, 'caretaker.setSeed': 1,
    'bank.assignPortions': 1, 'portions.removeAllocations': 1, 'payments.deleteRecord': 1, 'bank.unmatch': 1,
    'planner.addRule': 1, 'planner.updateRule': 1, 'planner.deleteRule': 1, 'planner.seed': 1, 'planner.syncCalendar': 1,
    'committee.add': 1, 'committee.update': 1, 'committee.delete': 1,
    'committee.archive': 1, 'committee.autoAllocate': 1, 'committee.seed': 1,
    'auth.setPin': 1, 'auth.clearPin': 1, 'occupancy.setStatus': 1, 'occupancy.deleteEntry': 1,
    'settings.getAll': 1, 'settings.get': 1, 'settings.saveAll': 1, 'settings.setActiveYear': 1,
    'corpus.getOwed': 1, 'corpus.setOwed': 1, 'corpus.getReceived': 1, 'corpus.addReceivedRow': 1,
    'corpus.updateReceivedRow': 1, 'corpus.deleteReceivedRow': 1, 'corpus.getAllFDs': 1,
    'corpus.addFD': 1, 'corpus.updateFD': 1, 'corpus.deleteFD': 1, 'corpus.addLine': 1, 'corpus.deleteLine': 1, 'lpg.setFlag': 1,
  };
  function _isAdminAction(a) {
    return ADMIN_ACTIONS[a] === 1 ||
      String(a).indexOf('settings.set') === 0 ||
      String(a).indexOf('categories.add') === 0 ||
      String(a).indexOf('categories.update') === 0 ||
      String(a).indexOf('categories.remove') === 0 ||
      String(a).indexOf('categories.reset') === 0;
  }

  function _dispatch(action, data) {
    if (_isAdminAction(action)) UsersService.requireAdmin();
    switch (action) {
      // User administration (admin-only, enforced above)
      case 'users.list':    return UsersService.listUsers();
      case 'users.add':     return UsersService.addUser(data.email, data.name, data.role);
      case 'users.setRole': return UsersService.setRole(data.email, data.role);
      case 'users.remove':  return UsersService.removeUser(data.email);

      // PIN-based fallback identity (see AuthService.gs)
      case 'auth.verifyPin': return AuthService.verifyPin(data.pin);
      case 'auth.setPin':    return AuthService.setPin(data.email, data.pin);
      case 'auth.clearPin':  return AuthService.clearPin(data.email);
      case 'auth.signOut':   return AuthService.signOut();

      // Units
      case 'units.getAll':        return UnitsService.getAllUnits();
      case 'units.getById':       return UnitsService.getUnitById(data.unit_id);
      case 'units.getByTower':    return UnitsService.getUnitsByTower(data.tower);
      case 'units.getByStatus':   return UnitsService.getUnitsByStatus(data.status);
      case 'units.getStats':      return UnitsService.getUnitStats();
      case 'units.profile': {
        // "All Payments" merges BOTH bank accounts' transactions for
        // this unit — the main IOB account plus the older IOB LPG one
        // (historical collections, mistaken payments). Account-2 rows
        // are tagged so the UI can label them, then the merged list is
        // re-sorted chronologically so the two accounts interleave by
        // date instead of appearing as two separate runs.
        var profileBank = BankService.getTransactionsForUnitSplitAware(data.unit_id);
        try {
          var profileBank2 = Bank2Service.getTransactionsForUnitSplitAware(data.unit_id);
          profileBank2.forEach(function(t) { t.account = '2'; });
          profileBank = profileBank.concat(profileBank2);
          profileBank.sort(function(a, b) {
            var da = parseFlexibleDate(a.date), db = parseFlexibleDate(b.date);
            return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
          });
        } catch (e2) {} // account 2 sheet may not exist yet — main account alone is fine
        return {
          unit:     UnitsService.getUnitById(data.unit_id),
          owners:   OwnersService.getOwnersByUnit(data.unit_id),
          tenants:  TenantsService.getTenantsByUnit(data.unit_id),
          payments: PaymentsService.getPaymentsByUnit(data.unit_id),
          bank:     profileBank
        };
      }
      case 'units.updateFlags':   return UnitsService.updateUnitFlags(data.unit_id, data);

      // Owners
      case 'owners.add':          return OwnersService.addOwner(data);
      case 'owners.getAll':       return OwnersService.getAllOwners();
      case 'owners.getByUnit':    return _ensurePhotoSharing(OwnersService.getOwnersByUnit(data.unit_id));
      case 'owners.archive':      return OwnersService.archiveOwner(data.owner_id, data.archived_by);
      case 'owners.history':      return OwnersService.getOwnerHistory(data.unit_id);
      case 'owners.getById':      return OwnersService.getOwnerById(data.owner_id);
      case 'owners.update':       return OwnersService.updateOwner(data.owner_id, data);
      case 'owners.delete':       return OwnersService.deleteOwner(data.owner_id);
      case 'owners.search':       return OwnersService.searchOwners(data.query);

      // Tenants
      case 'tenants.add':         return TenantsService.addTenant(data);
      case 'tenants.getAll':      return TenantsService.getAllTenants();
      case 'tenants.getByUnit':   return _ensurePhotoSharing(TenantsService.getTenantsByUnit(data.unit_id));
      case 'tenants.archive':     return TenantsService.archiveTenant(data.tenant_id, data.archived_by);
      case 'tenants.history':     return TenantsService.getTenantHistory(data.unit_id);
      case 'tenants.getById':     return TenantsService.getTenantById(data.tenant_id);
      case 'tenants.update':      return TenantsService.updateTenant(data.tenant_id, data);
      case 'tenants.remove':      return TenantsService.removeTenant(data.tenant_id);
      case 'tenants.search':      return TenantsService.searchTenants(data.query);
      case 'tenants.getDefaulters': return TenantsService.getDefaulters(data.month);

      // Payments
      case 'payments.submit':         return PaymentsService.submitPayment(data);
      case 'payments.update':         return PaymentsService.updatePayment(data.payment_id, data);
      case 'payments.verify':         return PaymentsService.verifyPayment(data.payment_id, data.reviewed_by, data.notes);
      case 'payments.reject':         return PaymentsService.rejectPayment(data.payment_id, data.reviewed_by, data.notes);
      case 'payments.getAll':         return PaymentsService.getAllPayments(data.filters || null);
      case 'payments.getPending':     return PaymentsService.getPendingPayments();
      case 'payments.getById':        return PaymentsService.getPaymentById(data.payment_id);
      case 'payments.getByUnit':      return PaymentsService.getPaymentsByUnit(data.unit_id, data.limit || null);
      case 'payments.getMonthlyHistory': return PaymentsService.getMonthlyPaymentHistory(data.unit_id);
      case 'payments.getStats':       return PaymentsService.getPaymentStats();

      // Export
      case 'export.units':         return ExportService.exportUnitsReport();
      case 'export.owners':        return ExportService.exportOwners();
      case 'export.tenants':       return ExportService.exportTenants();
      case 'export.payments':      return ExportService.exportPayments(data.filters || null);
      case 'export.defaulters':    return ExportService.exportDefaulters(data.month);
      case 'export.monthlySummary':return ExportService.exportMonthlySummary();
      case 'export.saveToDrive': {
        var url = ExportService.saveCsvToDrive(data.csvContent, data.filename);
        return { url: url };
      }

      // Import
      case 'import.owners':  return ExportService.importOwnersCsv(data.csvString);
      case 'import.tenants': return ExportService.importTenantsCsv(data.csvString);

      // Notifications
      case 'notify.sendReminder':      return NotificationsService.sendReminderToTenant(data.tenant_id, data.payment_type, data.channel);
      case 'notify.sendBulkReminders': return NotificationsService.sendBulkReminders(data.month, data.payment_type, data.channel);
      case 'notify.paymentVerified':   return NotificationsService.notifyPaymentVerified(data.payment_id, data.channel);
      case 'notify.paymentRejected':   return NotificationsService.notifyPaymentRejected(data.payment_id, data.reason, data.channel);

      // Bank Statements
      case 'bank.importCsv':         return _bankSvc(data).importCsv(data.csvText);
      case 'bank.getAll':            return _bankSvc(data).getAllTransactions(data.filters || null);
      case 'bank.getStats':          return _bankSvc(data).getStats();
      case 'bank.manualMatch':       return _bankSvc(data).manualMatch(data.txn_id, data.match_type, data.match_id, data.match_name, data.match_unit, data.payment_type, data.confirmed_historical);
      case 'bank.unmatch':           return _bankSvc(data).unmatchTransaction(data.txn_id, data.confirmed_historical, data.confirm_remove_posted);
      case 'bank.setExpenseCategory': return _bankSvc(data).setExpenseCategory(data.txn_id, data.expense_category, data.confirmed_historical);
      case 'bank.autoCategorizeExpenses': return _bankSvc(data).autoCategorizeExpenses();
      case 'bank.setDate':            return _bankSvc(data).setDate(data.txn_id, data.date, data.confirmed_historical);
      case 'bank.setPaymentType':    return _bankSvc(data).setPaymentType(data.txn_id, data.payment_type, data.confirmed_historical);
      case 'bank.setNote':           return _bankSvc(data).setNote(data.txn_id, data.note, data.confirmed_historical);
      case 'bank.rematch':           return _bankSvc(data).rematch();
      case 'bank.debugMatch':        return _bankSvc(data).debugMatch(data.name, data.unit_id);
      case 'bank.reconcile':         return _bankSvc(data).reconcile(null, data.year);
      case 'bank.reconcileMonthDetail': return _bankSvc(data).reconcileMonthDetail(data.month);

      // LPG meter readings
      case 'lpg.getMonth':        return LPGReadingService.getMonth(data.year, data.month);
      case 'lpg.getUnitHistory':  return LPGReadingService.getUnitHistory(data.unit_id);
      case 'lpg.saveReading':     return LPGReadingService.saveReading(data);
      case 'lpg.deleteReading':   return LPGReadingService.deleteReading(data.reading_id);
      case 'lpg.saveReadingsBatch': return LPGReadingService.saveReadingsBatch(data.readings);
      case 'lpg.rateFor':         return LPGReadingService.rateFor(data.year, data.month);
      case 'lpg.listRates':       return LPGReadingService.listRates();
      case 'lpg.setRate':         return LPGReadingService.setRate(data.year, data.month, data.conv_factor, data.price_per_kg, data.set_by);
      case 'lpg.deleteRate':      return LPGReadingService.deleteRate(data.rate_id);
      case 'lpg.setFlag':         return LPGReadingService.setFlag(data.reading_id, data.flagged, data.note);

      // Help documentation (open to everyone)
      case 'help.getContent':    return HelpService.getContent(data.doc_key);
      case 'help.getVersionInfo':return HelpService.getVersionInfo();
      case 'help.generatePdf':   return HelpService.generatePdf(data.doc_key);
      case 'lpg.importReadings':  return LPGImport.importFromRows(data.rows, data.year, data.month);

      // ── Danger Zone: selective monthly deletes ──
      case 'admin.getBankMonthSummary':       return BankService.getMonthSummary();
      case 'admin.deleteBankByMonth':         return BankService.deleteByMonth(data.month);
      case 'admin.getBank2MonthSummary':      return Bank2Service.getMonthSummary();
      case 'admin.deleteBank2ByMonth':        return Bank2Service.deleteByMonth(data.month);
      case 'admin.getLpgReadingsMonthSummary': return LPGReadingService.getMonthSummary();
      case 'admin.deleteLpgReadingsByMonth':  return LPGReadingService.deleteReadingsByMonth(data.year, data.month);
      case 'admin.getLpgInwardMonthSummary':  return LPGInventoryService.getInwardMonthSummary();
      case 'admin.deleteLpgInwardByMonth':    return LPGInventoryService.deleteInwardByMonth(data.month);
      case 'admin.getLpgOutwardMonthSummary': return LPGInventoryService.getOutwardMonthSummary();
      case 'admin.deleteLpgOutwardByMonth':   return LPGInventoryService.deleteOutwardByMonth(data.month);
      case 'admin.deleteAllOwners':           return OwnersService.deleteAllOwners();
      case 'admin.deleteAllTenants':          return TenantsService.deleteAllTenants();

      // ── Danger Zone: backup / restore / full wipe ──
      case 'expensePatterns.getAll':    return ExpensePatternService.getAll();

      case 'caretaker.getAll':          return CaretakerService.getAll();
      case 'caretaker.getByMonth':      return CaretakerService.getByMonth(data.year, data.month);
      case 'caretaker.addExpense':      return CaretakerService.addExpense(data);
      case 'caretaker.updateExpense':   return CaretakerService.updateExpense(data.expense_id, data);
      case 'caretaker.deleteExpense':   return CaretakerService.deleteExpense(data.expense_id, data.confirmed_historical);
      case 'caretaker.getMonthSummary': return CaretakerService.getMonthSummary();
      case 'caretaker.deleteByMonth':   return CaretakerService.deleteByMonth(data.month);
      case 'caretaker.setSeed':         return CaretakerService.setSeed(data.opening_balance, data.as_of_month);
      case 'caretaker.getSeed':         return CaretakerService.getSeed();
      case 'caretaker.getMonthReport':  return CaretakerService.getMonthReport(data.year, data.month);
      case 'expensePatterns.add':       return ExpensePatternService.addPattern(data.pattern_text, data.amount, data.expense_category);
      case 'expensePatterns.delete':    return ExpensePatternService.deletePattern(data.pattern_id);

      case 'reports.treasurerData':     return ReportService.getTreasurerReportData(data.year, data.month);

      // Community Google Forms (Move-In/Out, Party Hall, Vehicle Tracking)
      case 'forms.getResponses':        return FormsService.getResponses(data.kind);
      case 'forms.getVehicleRegistry':  return FormsService.getVehicleRegistry();
      case 'docs.list':                 return DocumentsService.listDocuments();

      // Payment confirmation screenshots (OCR + UTR matching)
      case 'screenshots.scan':   return ScreenshotService.scanFolder();
      case 'screenshots.list':   return ScreenshotService.listWithSuggestions();
      case 'screenshots.update': return ScreenshotService.updateShot(data.shot_id, data);
      case 'screenshots.delete': return ScreenshotService.deleteShot(data.shot_id);
      case 'screenshots.link':   return ScreenshotService.linkAndMatch(data);

      // Parked-funds model: portions on bank transactions + manual month allocation
      case 'bank.assignPortions':        return PortionsService.assignPortions(data);
      case 'bank.getPortionsForTxn':     return PortionsService.getPortionsForTxn(data.txn_id);
      case 'bank.getPortionTxnIds':      return PortionsService.getPortionTxnIds(data.account || '1');
      case 'bank.getPortionsMap':        return PortionsService.getPortionsMap(data.account || '1');
      case 'units.parkedFunds':          return PortionsService.getParkedSummary(data.unit_id);
      case 'portions.allocate':          return PortionsService.allocate(data);
      case 'portions.removeAllocations': return PortionsService.removeAllocations(data);
      case 'payments.deleteRecord':      return PaymentsService.deleteRecord(data.payment_id, data.confirmed_historical);

      // Monthly Occupancy & LPG Mode Timeline
      case 'occupancy.getTimeline':       return OccupancyService.getTimeline(data.unit_id, data.attribute);
      case 'occupancy.getYearView':       return OccupancyService.getYearView(data.unit_id, data.year);
      case 'occupancy.getYearViewForAll': return OccupancyService.getYearViewForAllUnits(data.year);
      case 'occupancy.setStatus':         return OccupancyService.setStatus(data.unit_id, data.attribute, data.effective_month, data.value);
      case 'occupancy.deleteEntry':       return OccupancyService.deleteEntry(data.entry_id);
      case 'occupancy.getAllEntriesForUnit': return OccupancyService.getAllEntriesForUnit(data.unit_id);

      // Year Planner
      case 'planner.getAllRules':       return YearPlannerService.getAllRules();
      case 'planner.addRule':           return YearPlannerService.addRule(data);
      case 'planner.updateRule':        return YearPlannerService.updateRule(data.rule_id, data);
      case 'planner.deleteRule':        return YearPlannerService.deleteRule(data.rule_id);
      case 'planner.seed':              return YearPlannerService.seedInitial();
      case 'planner.getOccurrences':    return YearPlannerService.getOccurrencesForYear(data.year);
      case 'planner.getDueReminders':   return YearPlannerService.getDueReminders();
      case 'planner.sendReminderNow':   return YearPlannerService.sendReminderNow(data.rule_id, data.event_date, data.stage);
      case 'planner.markWhatsappSent':  return YearPlannerService.markWhatsappSent(data.rule_id, data.event_date, data.stage);
      case 'planner.whatsappMessage':   return YearPlannerService.whatsappMessageFor(data.rule_id, data.event_date, data.stage);
      case 'planner.syncCalendar':      return YearPlannerService.syncYearToCalendar(data.year);
      case 'planner.calendarInfo':      return YearPlannerService.getCalendarInfo();

      // Office Bearers & EC Members
      case 'committee.getAll':       return CommitteeService.getAll();
      case 'committee.add':          return CommitteeService.addMember(data);
      case 'committee.update':       return CommitteeService.updateMember(data.member_id, data);
      case 'committee.delete':       return CommitteeService.deleteMember(data.member_id);
      case 'committee.archive':      return CommitteeService.archiveCurrent(data.term_label);
      case 'committee.autoAllocate': return CommitteeService.autoAllocateUnits();
      case 'committee.seed':         return CommitteeService.seedInitial();
      case 'committee.followUpBoard': return CommitteeService.getFollowUpBoard(data.year, data.month);
      case 'committee.followUpPdf': {
        var fuBlob = ReportPdfService.buildFollowUpPdf(data.year, data.month);
        return { filename: fuBlob.getName(), base64: Utilities.base64Encode(fuBlob.getBytes()) };
      }
      case 'forms.vehicleRegistry':     return FormsService.getVehicleRegistry();

      case 'reports.treasurerPdf': {
        var trBlob = ReportPdfService.buildTreasurerReportPdf(data.year, data.month);
        return { filename: trBlob.getName(), base64: Utilities.base64Encode(trBlob.getBytes()) };
      }

      case 'reports.caretakerData':      return CaretakerService.getMonthReport(data.year, data.month);
      case 'reports.monthlyIncomeData':  return ReportService.getMonthlyIncomeData(data.year, data.month);
      case 'reports.monthlyExpenseData': return ReportService.getMonthlyExpenseData(data.year, data.month);
      case 'reports.annualIncomeData':   return ReportService.getAnnualIncomeData(data.year);
      case 'reports.annualExpenseData':  return ReportService.getAnnualExpenseData(data.year);

      case 'reports.caretakerPdf': {
        var ckBlob = ReportPdfService.buildCaretakerReportPdf(data.year, data.month);
        return { filename: ckBlob.getName(), base64: Utilities.base64Encode(ckBlob.getBytes()) };
      }
      case 'reports.monthlyIncomePdf': {
        var miBlob = ReportPdfService.buildMonthlyIncomePdf(data.year, data.month);
        return { filename: miBlob.getName(), base64: Utilities.base64Encode(miBlob.getBytes()) };
      }
      case 'reports.monthlyExpensePdf': {
        var meBlob = ReportPdfService.buildMonthlyExpensePdf(data.year, data.month);
        return { filename: meBlob.getName(), base64: Utilities.base64Encode(meBlob.getBytes()) };
      }
      case 'reports.annualIncomePdf': {
        var aiBlob = ReportPdfService.buildAnnualIncomePdf(data.year);
        return { filename: aiBlob.getName(), base64: Utilities.base64Encode(aiBlob.getBytes()) };
      }
      case 'reports.annualExpensePdf': {
        var aeBlob = ReportPdfService.buildAnnualExpensePdf(data.year);
        return { filename: aeBlob.getName(), base64: Utilities.base64Encode(aeBlob.getBytes()) };
      }

      case 'admin.exportBackup':       return BackupService.exportBackup();
      case 'admin.restoreBackup':      return BackupService.restoreBackup(data.json_text);
      case 'admin.wipeForNewFacility': return BackupService.wipeForNewFacility();
      // Per-month "Mark UO" markers (set via Fees Received) for a given
      // fee type + month — the source of truth for month-specific
      // occupancy overrides (e.g. Jan vacant, Feb occupied again).
      case 'payments.getUOUnits': return PaymentsService.getAllPayments({ payment_type: data.payment_type, month: data.month, status: 'UO' }).map(function(p){ return p.unit_id; });
      // Which resident lived in each unit during a given month — checks
      // current tenants AND the archived history, by move-in/move-out date
      // range, so the LPG Readings page shows the CORRECT historical
      // tenant per month (residents change; billing history should not).
      case 'tenants.getResidentByMonth': return _tenantByMonth(data.year, data.month);

      // LPG Inventory (cylinder stock + business P&L)
      case 'lpgInv.getInward':      return LPGInventoryService.getInward();
      case 'lpgInv.addInward':      return LPGInventoryService.addInward(data);
      case 'lpgInv.deleteInward':   return LPGInventoryService.deleteInward(data.inward_id);
      case 'lpgInv.getOutward':     return LPGInventoryService.getOutward();
      case 'lpgInv.addOutward':     return LPGInventoryService.addOutward(data);
      case 'lpgInv.deleteOutward':  return LPGInventoryService.deleteOutward(data.outward_id);
      case 'lpgInv.getStockSummary':      return LPGInventoryService.getStockSummary();
      case 'lpgInv.getComparison':        return LPGInventoryService.getLiveComparison();
      case 'lpgInv.getMonthlyPL':   return LPGInventoryService.getMonthlyPL(data.year, data.month);
      case 'admin.resetPaymentData': return _resetPaymentData();

      // Dashboard — one bundled payload for the whole dashboard
      case 'dashboard.data':   return _dashboardData(data.year);

      // Settings
      case 'settings.getAll':  return SettingsService.getAllWithSchema();
      case 'settings.get':     return SettingsService.get(data.key);
      case 'settings.getActiveYear': return SettingsService.getActiveYear();
      case 'settings.setActiveYear': return SettingsService.setActiveYear(data.year);
      case 'settings.saveAll': return SettingsService.saveAll(data);

      // Corpus Fund (Admin-only — every function also self-checks via
      // UsersService.requireAdmin(), independent of this gate below)
      case 'corpus.getOwed':          return CorpusFundService.getOwed();
      case 'corpus.setOwed':          return CorpusFundService.setOwed(data.flats_owed, data.amount_per_flat);
      case 'corpus.getReceived':      return CorpusFundService.getReceived();
      case 'corpus.addReceivedRow':   return CorpusFundService.addReceivedRow(data.year, data.flats_paid, data.amount_per_flat);
      case 'corpus.updateReceivedRow':return CorpusFundService.updateReceivedRow(data.row_id, data.year, data.flats_paid, data.amount_per_flat);
      case 'corpus.deleteReceivedRow':return CorpusFundService.deleteReceivedRow(data.row_id);
      case 'corpus.getAllFDs':        return CorpusFundService.getAllFDs();
      case 'corpus.addFD':            return CorpusFundService.addFD(data);
      case 'corpus.updateFD':         return CorpusFundService.updateFD(data.fd_id, data);
      case 'corpus.deleteFD':         return CorpusFundService.deleteFD(data.fd_id);
      case 'corpus.addLine':          return CorpusFundService.addLine(data.fd_id, data.date, data.description, data.amount);
      case 'corpus.deleteLine':       return CorpusFundService.deleteLine(data.line_id);

      // ── Editable payment categories (credit fees / debit expenses) ──
      case 'categories.getAll': return CategoriesService.getAll(data.kind || null);
      case 'categories.add':    return CategoriesService.add(data.kind, data.name);
      case 'categories.update': return CategoriesService.update(data.category_id, data);
      case 'categories.delete': return CategoriesService.remove(data.category_id);
      case 'categories.reset':  return CategoriesService.resetToDefaults();

      default:
        throw new Error('Unknown action: ' + action);
    }
  }


  // ── DANGER: wipe payment data (Bank ledger + Fees ledger) for a fresh
  // test import. Owners, Tenants, Units, Settings, Categories are untouched.
  function _resetPaymentData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var out = {};
    // BOTH bank accounts — the main IOB account and the IOB LPG account
    // (BankStatements2) — plus the fees ledger. "ALL" must mean all:
    // leaving account 2 behind kept its rows visible in unit profiles
    // and reconciliation after a "full reset".
    ['BankStatements', 'BankStatements2', 'Payments'].forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) { out[name] = 0; return; }
      var last = sh.getLastRow();
      if (last > 1) sh.deleteRows(2, last - 1);
      out[name] = Math.max(0, last - 1);
    });
    return { cleared: out };
  }

  // For every unit, find whichever resident (current OR archived) had a
  // move-in/move-out range overlapping the given month — i.e. who actually
  // lived there that month, even if someone else lives there today.
  function _tenantByMonth(year, month) {
    year = Number(year); month = Number(month);
    var monthStart = new Date(year, month - 1, 1).getTime();
    var monthEnd   = new Date(year, month, 0).getTime(); // last day of month

    function overlaps(t) {
      var moveIn  = t.move_in_date  ? new Date(t.move_in_date).getTime()  : -Infinity;
      var moveOut = t.move_out_date ? new Date(t.move_out_date).getTime() : Infinity;
      return moveIn <= monthEnd && moveOut >= monthStart;
    }

    var all = TenantsService.getAllTenants().concat(TenantsService.getAllTenantHistory());

    var out = {};
    all.forEach(function(t) {
      if (!overlaps(t)) return;
      var uid = String(t.unit_id).toUpperCase();
      // If several overlap (shouldn't normally happen), prefer the one
      // with the latest move-in date — the most recent occupant that month.
      if (!out[uid] || new Date(t.move_in_date || 0) > new Date(out[uid].move_in_date || 0)) {
        out[uid] = { name: t.name, move_in_date: t.move_in_date, move_out_date: t.move_out_date };
      }
    });

    // Fallback for units with no move-in/move-out dates recorded at all
    // (common for residents who were already there before this system
    // existed) — prefer ANY current tenant on file over showing the owner,
    // since the tenant is who actually consumes and pays for LPG/WMF.
    TenantsService.getAllTenants().forEach(function(t) {
      var uid = String(t.unit_id).toUpperCase();
      if (!out[uid] && t.name) out[uid] = { name: t.name, move_in_date: t.move_in_date, move_out_date: t.move_out_date, approximate: true };
    });
    return out;
  }

  // ── Dashboard aggregation ──────────────────────────────────────
  // Parse {y, m, d} from a stored bank-date string (ISO, DD-Mon-YYYY,
  // or numeric day-first). Returns null when unparsable.
  function _dashYMD(s) {
    if (!s) return null;
    s = String(s).trim();
    var low = s.toLowerCase();
    var mn = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
    var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return { y:+m[1], m:+m[2], d:+m[3] };
    m = low.match(/^(\d{1,2})[-\/ ]([a-z]{3})[a-z]*[-\/ ](\d{2,4})/);
    if (m && mn[m[2]]) { var y1=+m[3]; if (y1<100) y1+=2000; return { y:y1, m:mn[m[2]], d:+m[1] }; }
    m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})/);
    if (m) {
      var yy=+m[3]; if (yy<100) yy+=2000;
      var a=+m[1], b=+m[2], mm=b, dd=a;
      if (b>12 && a<=12) { mm=a; dd=b; }
      return { y:yy, m:mm, d:dd };
    }
    return null;
  }

  function _dashboardData(year) {
    // viewYear = the year this payload is filtered to (the caller may
    // ask for any year); settingsYear = the app-wide Active Year, which
    // the Dashboard banner must keep showing regardless of the view.
    var settingsYear = SettingsService.getActiveYear();
    var activeYear = Number(year) || settingsYear;

    // Fetch ONCE, unfiltered — reconcile() needs the FULL chronological
    // chain (a year's opening balance depends on the previous year's true
    // closing balance), so it must never see a year-filtered subset. The
    // year-scoped view used for the Dashboard's own tiles is then derived
    // from this SAME already-fetched array via in-memory filtering —
    // no second sheet read either way.
    var allTxns = BankService.getAllTransactions(null);
    var txns = allTxns.filter(function(t) {
      var ymd = _dashYMD(t.date);
      return ymd && ymd.y === activeYear;
    });

    var months = {}, latest = null, latestKey = -1;
    for (var i = 0; i < txns.length; i++) {
      var t = txns[i];
      var ymd = _dashYMD(t.date);
      if (!ymd) continue;
      var key = ymd.y + '-' + (ymd.m < 10 ? '0' + ymd.m : ymd.m);
      if (!months[key]) months[key] = { i: 0, e: 0 };
      var cr = Number(t.credit || 0), db = Number(t.debit || 0);
      if (cr > 0) months[key].i += cr;
      if (db > 0) months[key].e += db;
      var sortKey = ymd.y * 10000 + ymd.m * 100 + ymd.d;
      if (sortKey >= latestKey && t.balance !== '' && !isNaN(Number(t.balance))) {
        latestKey = sortKey;
        latest = { amount: Number(t.balance), date: t.date };
      }
    }

    // Payments ledger (recorded fees) — compact rows, Rejected excluded,
    // scoped to the active year the same way.
    var allPays = PaymentsService.getAllPayments(null);
    var pays = allPays.filter(function(pp) { return String(pp.month || '').indexOf(activeYear + '-') === 0; });
    var pRows = [];
    for (var p = 0; p < pays.length; p++) {
      var pp = pays[p];
      if (pp.status === 'Rejected') continue;
      pRows.push({ u: String(pp.unit_id).toUpperCase(), t: pp.payment_type, m: pp.month,
                   a: Number(pp.amount || 0), uo: pp.status === 'UO' ? 1 : 0 });
    }

    // Units with flags
    var units = UnitsService.getAllUnits().map(function(u) {
      return { id: u.unit_id, tw: u.tower, fl: Number(u.floor), num: u.unit_number,
               occ: u.occupancy, reg: u.registration,
               cp: u.corpus_paid, oit: u.owner_is_tenant };
    });

    // Lite owner/tenant info for the search box
    var owners = OwnersService.getAllOwners().map(function(o) {
      return { u: String(o.unit_id).toUpperCase(),
               n: [o.name, o.name2, o.name3].filter(function(x){return x;}).join(' / '),
               p: [o.phone, o.phone2, o.phone3].filter(function(x){return x;}).join(' '),
               e: [o.email, o.email2].filter(function(x){return x;}).join(' '),
               b: o.bank_names || '' };
    });
    var tenants = TenantsService.getAllTenants().map(function(tn) {
      return { u: String(tn.unit_id).toUpperCase(),
               n: [tn.name, tn.name2].filter(function(x){return x;}).join(' / '),
               p: [tn.phone, tn.phone2].filter(function(x){return x;}).join(' '),
               e: tn.email || '',
               b: tn.bank_name || '' };
    });

    // Reconciliation summary: does the ledger match the bank's own
    // balances? The COMPUTATION must walk the full history (opening
    // balances chain across year boundaries), but the SUMMARY counts
    // only the viewed year's months — the tile describes the year on
    // screen, not the whole database.
    var recon = { total: 0, bad: 0, badMonths: [], year: activeYear };
    try {
      var rec = BankService.reconcile(allTxns); // full history, needed for correctness — NOT the year-filtered view
      for (var rr = 0; rr < rec.length; rr++) {
        if (String(rec[rr].month).indexOf(activeYear + '-') !== 0) continue;
        recon.total++;
        if (!rec[rr].ok) recon.bad++;
        if (!rec[rr].ok) recon.badMonths.push(rec[rr].month);
      }
    } catch (recErr) {}

    var settingsAll = SettingsService.getAll();
    return { months: months, balance: latest, payments: pRows, recon: recon,
             units: units, owners: owners, tenants: tenants,
             activeYear: settingsYear, viewYear: activeYear,
             feeMaintenance: Number(settingsAll.fee_maintenance) || 0,
             feeWaste: Number(settingsAll.fee_waste) || 0 };
  }

  return { route: route };
})();
