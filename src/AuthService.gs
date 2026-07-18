// ═══════════════════════════════════════════════════════════════
// AuthService.gs — PIN-based fallback identity.
//
// THE PROBLEM THIS SOLVES: Google will only reveal a visitor's email
// to a personal (non-Workspace) Apps Script project for the OWNER —
// every other visitor's Session.getActiveUser().getEmail() comes back
// blank. That means the Users sheet's per-person Admin/User system
// can only ever recognize the owner; everyone else falls into one
// generic bucket, no matter what's in the Users list.
//
// THE FIX: a short PIN, set once per person by an Admin, lets OXYGEN
// identify a specific visitor WITHOUT needing Google to reveal their
// email. Once verified, the identity is remembered via Apps Script's
// own per-visitor UserCache (CacheService.getUserCache()) — this is
// scoped automatically to "this browser/session" by the platform
// itself, so no token needs to be carried or attached by the client
// on every request; every existing google.script.run call anywhere
// in the app benefits automatically, with no changes needed to any
// of them.
//
// LIMITS, honestly: UserCache entries expire after at most 6 hours,
// so a PIN may need re-entering after a long gap — a real trade-off
// for not needing any client-side plumbing changes. PINs are never
// stored in plain text — only a salted one-way hash — so nobody
// with sheet access, including an Admin, can read a PIN back; only
// the Admin who set it (and told the resident directly) knows it.
// ═══════════════════════════════════════════════════════════════
var AuthService = (function() {
  var CACHE_KEY = 'oxygen_pin_identity';
  var CACHE_TTL_SEC = 21600; // 6 hours — CacheService's own maximum

  function _secret() {
    var props = PropertiesService.getScriptProperties();
    var s = props.getProperty('OXYGEN_PIN_SECRET');
    if (!s) {
      s = Utilities.getUuid() + Utilities.getUuid();
      props.setProperty('OXYGEN_PIN_SECRET', s);
    }
    return s;
  }

  function _hashPin(pin, email) {
    var raw = String(pin).trim() + ':' + String(email).trim().toLowerCase() + ':' + _secret();
    var digestBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
    return Utilities.base64Encode(digestBytes);
  }

  // Admin sets or resets one resident's PIN. The raw PIN is never
  // stored — only this salted hash — so it must be given to the
  // resident directly (WhatsApp, verbally); it cannot be looked up
  // or recovered later, only reset to a new one.
  function setPin(email, pin) {
    UsersService.requireAdmin();
    email = String(email).trim().toLowerCase();
    pin = String(pin || '').trim();
    if (pin.length < 4) throw new Error('PIN must be at least 4 characters.');
    var hash = _hashPin(pin, email);
    UsersService.setPinHash(email, hash);
    return { success: true };
  }

  function clearPin(email) {
    UsersService.requireAdmin();
    UsersService.setPinHash(String(email).trim().toLowerCase(), '');
    return { success: true };
  }

  // The PIN alone identifies which resident it belongs to — nobody
  // needs to type their email too. Only ever reached when Google
  // itself couldn't identify the visitor; anyone whose email Google
  // DOES reveal never sees this path at all.
  function verifyPin(pin) {
    pin = String(pin || '').trim();
    if (!pin) throw new Error('Enter your PIN.');
    var users = UsersService.listUsersWithPinHash();
    for (var i = 0; i < users.length; i++) {
      var u = users[i];
      if (!u.pin_hash) continue;
      if (_hashPin(pin, u.email) === u.pin_hash) {
        CacheService.getUserCache().put(CACHE_KEY, JSON.stringify({ email: u.email, role: u.role }), CACHE_TTL_SEC);
        return { success: true, email: u.email, role: u.role };
      }
    }
    throw new Error('PIN not recognized. Check with your Administrator, or ask them to set one for you in Settings → User Administration.');
  }

  // Reads back whatever verifyPin() last stored for THIS visitor's
  // browser/session — entirely automatic, no token passed by anyone.
  function getCachedIdentity() {
    try {
      var raw = CacheService.getUserCache().get(CACHE_KEY);
      if (!raw) return null;
      var parsed = JSON.parse(raw);
      if (!parsed || !parsed.email) return null;
      return parsed;
    } catch (e) { return null; }
  }

  function signOut() {
    try { CacheService.getUserCache().remove(CACHE_KEY); } catch (e) {}
    return { success: true };
  }

  return {
    setPin: setPin,
    clearPin: clearPin,
    verifyPin: verifyPin,
    getCachedIdentity: getCachedIdentity,
    signOut: signOut
  };
})();

