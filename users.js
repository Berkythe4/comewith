/*
  Come With — Auth Module (users.js)
  Include on every page that needs authentication.

  GOOGLE APPS SCRIPT: Your doGet must handle these actions:
    ?action=login&email=X&hash=Y        → returns {success,user} or {success:false,error}
    ?action=getUsers                     → returns [{email,name,role,created,lastLogin,mustChangePassword}]
    ?action=checkUser&email=X            → returns {exists,mustSetPassword} for first-time flow

  Your doPost must handle these types:
    {type:'createUser', email, name, passwordHash, role, mustChangePassword}
    {type:'updatePassword', email, newPasswordHash, mustChangePassword:false}
    {type:'updateUser', email, role}        — promote/demote
    {type:'deleteUser', email}
    {type:'updateLastLogin', email, lastLogin}
*/

var CW_AUTH = (function() {
  'use strict';

  var SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxM-cixH5nNp7u7RnXxVG5fUsIOpC0xxpkg8VC0-C_9GKxV0ZBjxpXT91l-gkMbutC8Gw/exec';

  var ROLES = {
    MASTER_ADMIN: 'master_admin',
    SUB_ADMIN: 'sub_admin',
    CUSTOMER: 'customer'
  };

  var MASTER_EMAIL = 'berky@comewith.org';

  // ── SESSION ────────────────────────────────────────────
  function getSession() {
    try {
      var raw = sessionStorage.getItem('cw_user');
      return raw ? JSON.parse(raw) : null;
    } catch(e) { return null; }
  }

  function setSession(user) {
    sessionStorage.setItem('cw_user', JSON.stringify(user));
    sessionStorage.setItem('cw_admin', '1'); // backward compat
  }

  function clearSession() {
    sessionStorage.removeItem('cw_user');
    sessionStorage.removeItem('cw_admin');
  }

  function isLoggedIn() {
    return !!getSession();
  }

  function getUser() {
    return getSession();
  }

  // ── PASSWORD HASHING (SHA-256) ─────────────────────────
  function hashPassword(password) {
    var encoder = new TextEncoder();
    var data = encoder.encode(password);
    return crypto.subtle.digest('SHA-256', data).then(function(buf) {
      var arr = Array.from(new Uint8Array(buf));
      return arr.map(function(b) { return b.toString(16).padStart(2, '0'); }).join('');
    });
  }

  // ── PASSWORD VALIDATION ────────────────────────────────
  function validatePassword(password, role) {
    var errors = [];
    if (role === ROLES.CUSTOMER) {
      if (password.length < 6) errors.push('Password must be at least 6 characters.');
    } else {
      // Admin rules
      if (password.length < 10) errors.push('Password must be at least 10 characters.');
      if (!/[0-9]/.test(password)) errors.push('Password must include at least 1 number.');
      if (!/[^a-zA-Z0-9]/.test(password)) errors.push('Password must include at least 1 special character.');
    }
    return errors;
  }

  // ── RANDOM PASSWORD ────────────────────────────────────
  function generatePassword(len) {
    len = len || 8;
    var chars = 'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    var pw = '';
    var arr = new Uint8Array(len);
    crypto.getRandomValues(arr);
    for (var i = 0; i < len; i++) {
      pw += chars[arr[i] % chars.length];
    }
    return pw;
  }

  // ── API HELPERS ────────────────────────────────────────
  function apiGet(params) {
    var qs = new URLSearchParams(params).toString();
    return fetch(SCRIPT_URL + '?' + qs)
      .then(function(r) { return r.json(); });
  }

  function apiPost(data) {
    return fetch(SCRIPT_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
  }

  // ── LOGIN ──────────────────────────────────────────────
  function login(email, password) {
    return hashPassword(password).then(function(hash) {
      return apiGet({ action: 'login', email: email.toLowerCase().trim(), hash: hash });
    }).then(function(resp) {
      if (resp && resp.success && resp.user) {
        setSession(resp.user);
        // Update last login
        apiPost({
          type: 'updateLastLogin',
          email: resp.user.email,
          lastLogin: new Date().toLocaleString('en-US', { timeZone: 'America/New_York' })
        });
        return { success: true, user: resp.user };
      }
      return { success: false, error: (resp && resp.error) || 'Invalid email or password.' };
    });
  }

  // ── LOGOUT ─────────────────────────────────────────────
  function logout() {
    clearSession();
    window.location.href = 'auth.html';
  }

  // ── CHECK USER (for first-time flow) ───────────────────
  function checkUser(email) {
    return apiGet({ action: 'checkUser', email: email.toLowerCase().trim() });
  }

  // ── CREATE USER ────────────────────────────────────────
  function createUser(email, name, password, role, mustChangePassword) {
    return hashPassword(password).then(function(hash) {
      return apiPost({
        type: 'createUser',
        email: email.toLowerCase().trim(),
        name: name,
        passwordHash: hash,
        role: role || ROLES.CUSTOMER,
        mustChangePassword: mustChangePassword !== false,
        created: new Date().toLocaleString('en-US', { timeZone: 'America/New_York' })
      });
    });
  }

  // ── UPDATE PASSWORD ────────────────────────────────────
  function updatePassword(email, newPassword) {
    return hashPassword(newPassword).then(function(hash) {
      return apiPost({
        type: 'updatePassword',
        email: email.toLowerCase().trim(),
        newPasswordHash: hash,
        mustChangePassword: false
      });
    });
  }

  // ── UPDATE USER ROLE ───────────────────────────────────
  function updateUserRole(email, newRole) {
    return apiPost({ type: 'updateUser', email: email, role: newRole });
  }

  // ── DELETE USER ────────────────────────────────────────
  function deleteUser(email) {
    return apiPost({ type: 'deleteUser', email: email });
  }

  // ── GET ALL USERS (admin) ──────────────────────────────
  function getUsers() {
    return apiGet({ action: 'getUsers' });
  }

  // ── PAGE PROTECTION ────────────────────────────────────
  // Call on page load. Options:
  //   requireAuth: true|false
  //   allowedRoles: ['master_admin','sub_admin','customer']
  //   onReady: function(user) — called when auth passes
  function init(options) {
    options = options || {};
    var requireAuth = options.requireAuth !== false;
    var allowedRoles = options.allowedRoles || null;

    if (!requireAuth) {
      if (options.onReady) options.onReady(getSession());
      return;
    }

    var user = getSession();

    if (!user) {
      window.location.href = 'auth.html?redirect=' + encodeURIComponent(window.location.href);
      return;
    }

    // Must change password check
    if (user.mustChangePassword && !window.location.pathname.match(/auth\.html/)) {
      window.location.href = 'auth.html?changePassword=1';
      return;
    }

    // Role check
    if (allowedRoles && allowedRoles.indexOf(user.role) === -1) {
      // Redirect to appropriate home
      if (user.role === ROLES.CUSTOMER) {
        window.location.href = 'customer_portal.html';
      } else {
        window.location.href = 'dashboard.html';
      }
      return;
    }

    if (options.onReady) options.onReady(user);
  }

  // ── ROLE HELPERS ───────────────────────────────────────
  function isMasterAdmin(user) {
    user = user || getSession();
    return user && user.role === ROLES.MASTER_ADMIN;
  }

  function isAdmin(user) {
    user = user || getSession();
    return user && (user.role === ROLES.MASTER_ADMIN || user.role === ROLES.SUB_ADMIN);
  }

  function isCustomer(user) {
    user = user || getSession();
    return user && user.role === ROLES.CUSTOMER;
  }

  // ── REDIRECT AFTER LOGIN ───────────────────────────────
  function getLoginRedirect(user) {
    var params = new URLSearchParams(window.location.search);
    var redirect = params.get('redirect');
    if (redirect) return redirect;
    if (user.role === ROLES.CUSTOMER) return 'customer_portal.html';
    return 'dashboard.html';
  }

  // ── PUBLIC API ─────────────────────────────────────────
  return {
    ROLES: ROLES,
    MASTER_EMAIL: MASTER_EMAIL,
    SCRIPT_URL: SCRIPT_URL,
    init: init,
    login: login,
    logout: logout,
    getUser: getUser,
    isLoggedIn: isLoggedIn,
    isMasterAdmin: isMasterAdmin,
    isAdmin: isAdmin,
    isCustomer: isCustomer,
    hashPassword: hashPassword,
    validatePassword: validatePassword,
    generatePassword: generatePassword,
    checkUser: checkUser,
    createUser: createUser,
    updatePassword: updatePassword,
    updateUserRole: updateUserRole,
    deleteUser: deleteUser,
    getUsers: getUsers,
    getLoginRedirect: getLoginRedirect,
    clearSession: clearSession,
    setSession: setSession
  };
})();
