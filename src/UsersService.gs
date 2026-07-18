// ═══════════════════════════════════════════════════════════════
// UsersService — Google-login based roles (Admin / User)
// The viewer's Google email (Session.getActiveUser) is looked up in
// the Users sheet. The script OWNER is always an Admin (failsafe, so
// you can never lock yourself out). First visitor is auto-admin.
// ═══════════════════════════════════════════════════════════════
var UsersService = (function() {

  var SHEET = 'Users';
  var C = { EMAIL: 0, NAME: 1, ROLE: 2, ADDED_BY: 3, ADDED_AT: 4, PIN_HASH: 5 };

  function ensureSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET);
    if (!sh) {
      sh = ss.insertSheet(SHEET);
      sh.appendRow(['email', 'name', 'role', 'added_by', 'added_at', 'pin_hash']);
      sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    } else if (sh.getLastColumn() < 6) {
      // Migration for sheets created before PIN login existed — adds
      // the column without touching a single existing row.
      sh.getRange(1, 6).setValue('pin_hash').setFontWeight('bold').setBackground('#0f2744').setFontColor('#ffffff');
    }
    return sh;
  }

  function _norm(e) { return String(e || '').trim().toLowerCase(); }

  function _rows() {
    var sh = ensureSheet();
    if (sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  }

  // 'admin' | 'user' | '' (unknown / not allowed)
  function getRole(email) {
    email = _norm(email);
    if (!email) return '';
    if (email === _norm(Session.getEffectiveUser().getEmail())) return 'admin'; // owner failsafe
    var rows = _rows();
    // Bootstrap: empty user list → first signed-in visitor becomes Admin
    if (!rows.length) {
      ensureSheet().appendRow([email, '', 'admin', 'bootstrap', new Date().toISOString()]);
      return 'admin';
    }
    for (var i = 0; i < rows.length; i++) {
      if (_norm(rows[i][C.EMAIL]) === email) {
        return String(rows[i][C.ROLE]).toLowerCase() === 'admin' ? 'admin' : 'user';
      }
    }
    return '';
  }

  function requireAdmin() {
    var me = Session.getActiveUser().getEmail();
    var role = getRole(me);
    if (role !== 'admin') {
      // Google couldn't (or didn't) identify this visitor as an admin
      // via email — check whether they verified a PIN instead. This is
      // what makes PIN-based Admin access actually work for gated
      // server actions, not just for what the sidebar chip displays.
      var cached = AuthService.getCachedIdentity();
      if (cached && cached.role === 'admin') return cached.email;
      throw new Error('Administrator access required.');
    }
    return me;
  }

  // Every user row INCLUDING its pin_hash — for AuthService's own PIN
  // verification only. Never exposed to the client directly; no router
  // route returns this shape.
  function listUsersWithPinHash() {
    return _rows().map(function(r) {
      return { email: _norm(r[C.EMAIL]), role: String(r[C.ROLE] || 'user').toLowerCase(),
               pin_hash: String(r[C.PIN_HASH] || '') };
    });
  }

  // Sets (or, with an empty hash, clears) one user's stored PIN hash.
  // Called only from AuthService.setPin/clearPin, which already
  // enforce requireAdmin() themselves.
  function setPinHash(email, hash) {
    email = _norm(email);
    var sh = ensureSheet(), rows = _rows();
    for (var i = 0; i < rows.length; i++) {
      if (_norm(rows[i][C.EMAIL]) === email) {
        sh.getRange(i + 2, C.PIN_HASH + 1).setValue(hash);
        return { success: true };
      }
    }
    throw new Error('User not found: ' + email + ' — add them as a user first.');
  }

  function listUsers() {
    var owner = _norm(Session.getEffectiveUser().getEmail());
    return _rows().map(function(r) {
      return { email: String(r[C.EMAIL]), name: String(r[C.NAME] || ''),
               role: String(r[C.ROLE] || 'user').toLowerCase(),
               added_by: String(r[C.ADDED_BY] || ''), added_at: String(r[C.ADDED_AT] || ''),
               is_owner: _norm(r[C.EMAIL]) === owner,
               has_pin: !!String(r[C.PIN_HASH] || '').trim() };
    });
  }

  function addUser(email, name, role) {
    var by = requireAdmin();
    email = _norm(email);
    if (!email || email.indexOf('@') < 1) throw new Error('A valid email address is required.');
    var rows = _rows();
    for (var i = 0; i < rows.length; i++) {
      if (_norm(rows[i][C.EMAIL]) === email) throw new Error('This email is already a user.');
    }
    role = String(role).toLowerCase() === 'admin' ? 'admin' : 'user';
    ensureSheet().appendRow([email, name || '', role, by, new Date().toISOString()]);
    return { success: true };
  }

  function setRole(email, role) {
    var by = requireAdmin();
    email = _norm(email);
    role = String(role).toLowerCase() === 'admin' ? 'admin' : 'user';
    // An admin may not demote themself (prevents an accidental lock-out)
    if (email === _norm(by) && role !== 'admin') throw new Error('You cannot demote your own account.');
    var sh = ensureSheet(), rows = _rows();
    for (var i = 0; i < rows.length; i++) {
      if (_norm(rows[i][C.EMAIL]) === email) {
        sh.getRange(i + 2, C.ROLE + 1).setValue(role);
        return { success: true };
      }
    }
    throw new Error('User not found: ' + email);
  }

  function removeUser(email) {
    var by = requireAdmin();
    email = _norm(email);
    if (email === _norm(by)) throw new Error('You cannot remove your own account.');
    var sh = ensureSheet(), rows = _rows();
    for (var i = 0; i < rows.length; i++) {
      if (_norm(rows[i][C.EMAIL]) === email) { sh.deleteRow(i + 2); return { success: true }; }
    }
    throw new Error('User not found: ' + email);
  }

  return {
    ensureSheet:  ensureSheet,
    getRole:      getRole,
    requireAdmin: requireAdmin,
    listUsers:    listUsers,
    addUser:      addUser,
    setRole:      setRole,
    removeUser:   removeUser,
    setPinHash:   setPinHash,
    listUsersWithPinHash: listUsersWithPinHash
  };
})();
