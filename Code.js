/**
 * ============================================================
 * CAREER EDUCATION UNIT SITE - BACKEND (Code.gs)
 * Login: school email + one-time code (OTP) emailed to the user.
 * Also: session tokens, rate limiting, manual "last updated" stamp,
 * officer dashboard, user management (Admin + Officer), and the Register request form.
 *
 * Users sheet columns:
 *   A: School email (username) | B: Name (optional, used in the OTP email greeting)
 *   C: Role | D: (unused) | E: (unused) | F: School
 *
 * Project files needed in Apps Script:
 *   Code.gs, Index.html, Styles.html, JavaScript.html, Pages*.html (incl. PagesDashboard.html),
 *   OtpEmail.html
 * ============================================================
 */

/** ---------- CONFIG ---------- */
const SESSION_DURATION_SECONDS = 60 * 60; // 1 hour
const MAX_LOGIN_ATTEMPTS = 5;
const LOGIN_LOCKOUT_SECONDS = 15 * 60;    // 15 minutes

// Access requests (Register form on the login page)
const REQUESTS_SHEET_ID = '1CfP1HskKd4i08LJettrB-4Ev01Acj7ySL4_Dox6bXJc';
const REQUESTS_TAB = 'Requests';
const REQUEST_INBOX = 'upk@ds.edu.moe.bn';

// OTP settings
const OTP_TTL_SECONDS = 10 * 60;
const OTP_MAX_ATTEMPTS = 5;
const OTP_RESEND_COOLDOWN = 60;
const EMAIL_DOMAIN_REGEX = /^[^@\s]+@(?:[a-z0-9\-]+\.)*(?:edu\.moe\.bn|moe\.gov\.bn|jpi\.edu\.bn)$/i;

// Logo shown at the top of the OTP email (Google Drive file ID of the logo image)
const EMAIL_LOGO_FILE_ID = '1zPSRmfzO5pneUmYNBTP6Qh8-u7B-7RMx';

const ADMIN_VALID_ROLES_ = ["Principal", "Deputy Principal", "Head of Career Teacher", "Career Committee", "Officer", "Admin"];

// WHO CAN SEE WHAT (lowercase).
// Dashboard: Officer + Admin. User management + access requests: Officer + Admin.
// Only an Admin can assign the Admin role, edit/delete Admin accounts, or reset the visit counter.
const DASHBOARD_ROLES_ = ['officer', 'admin'];
const USER_MANAGER_ROLES_ = ['officer', 'admin'];
const STATS_CACHE_KEY_ = 'dash_stats_v2';
const STATS_CACHE_SECONDS_ = 120;

/**
 * 1. SITE PAGE LOADER
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Career Education Unit Site')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lets Index.html pull in split-out files: <?!= include('Styles'); ?> */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Adds a "CEU Site Admin" menu to the Sheet UI. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CEU Site Admin')
    .addItem('Mark Site as Updated', 'markSiteAsUpdated')
    .addItem('Trim Logs Older Than 6 Months', 'adminTrimOldLogs')
    .addToUi();
}

/** ---------- HASH HELPERS (used to store OTP codes safely) ---------- */

function generateSalt_() {
  return Utilities.getUuid();
}

function hashPassword_(value, salt) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + value,
    Utilities.Charset.UTF_8
  );
  return digest.map(function(byte) {
    var v = (byte < 0) ? byte + 256 : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

/** ---------- LOGIN RATE LIMITING ---------- */

function isLockedOut_(username) {
  var cache = CacheService.getScriptCache();
  return cache.get("lock_" + username) !== null;
}

function recordFailedAttempt_(username) {
  var cache = CacheService.getScriptCache();
  var attemptsKey = "attempts_" + username;
  var attempts = Number(cache.get(attemptsKey) || 0) + 1;
  cache.put(attemptsKey, String(attempts), LOGIN_LOCKOUT_SECONDS);

  if (attempts >= MAX_LOGIN_ATTEMPTS) {
    cache.put("lock_" + username, "1", LOGIN_LOCKOUT_SECONDS);
  }
}

function clearFailedAttempts_(username) {
  var cache = CacheService.getScriptCache();
  cache.remove("attempts_" + username);
  cache.remove("lock_" + username);
}

/** ---------- SESSION TOKENS ---------- */

function createSession_(username, role, school) {
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  cache.put("session_" + token, JSON.stringify({ username: username, role: role, school: school || "" }), SESSION_DURATION_SECONDS);
  return token;
}

/** Throws if the token is missing, invalid, or expired. */
function requireSession_(token) {
  if (!token) throw new Error("Not authenticated. Please log in again.");
  var cache = CacheService.getScriptCache();
  var raw = cache.get("session_" + token);
  if (!raw) throw new Error("Session expired. Please log in again.");
  return JSON.parse(raw); // { username, role, school }
}

/** ---------- OTP LOGIN ---------- */

function findUserByEmail_(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sheet) throw new Error("'Users' sheet not found.");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === email) {
      return {
        username: data[i][0].toString().trim(),
        name: data[i][1] ? data[i][1].toString().trim() : "",
        role: data[i][2] ? data[i][2].toString().trim() : "User",
        school: data[i][5] ? data[i][5].toString().trim() : ""
      };
    }
  }
  return null;
}

function generateOtp_() {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Utilities.getUuid() + Utilities.getUuid() + Date.now()
  );
  var code = "";
  for (var i = 0; i < 6; i++) code += String((bytes[i] & 255) % 10);
  return code;
}

/**
 * Sends the styled OTP email (template: OtpEmail.html).
 */
function sendOtpEmail_(to, name, code) {
  var expiry = new Date(Date.now() + OTP_TTL_SECONDS * 1000);
  var expiryText = Utilities.formatDate(expiry, 'Asia/Brunei', 'dd/MM/yyyy HH:mm:ss');

  // Logo is optional: if it can't be read, the email is sent without it.
  var inlineImages = {};
  var hasLogo = false;
  try {
    if (EMAIL_LOGO_FILE_ID) {
      inlineImages.ceulogo = DriveApp.getFileById(EMAIL_LOGO_FILE_ID).getBlob().setName('ceu-logo.png');
      hasLogo = true;
    }
  } catch (e) {
    Logger.log('OTP email logo not attached: ' + e);
  }

  var tpl = HtmlService.createTemplateFromFile('OtpEmail');
  tpl.hasLogo = hasLogo;
  tpl.greeting = name ? 'Hello ' + name : 'Hello,';
  tpl.code = code.split('').join(' ');   // "6 5 2 5 2 1" style spacing
  tpl.expiryText = expiryText;
  var htmlBody = tpl.evaluate().getContent();

  MailApp.sendEmail({
    to: to,
    subject: 'Your CEU Site verification code',
    body: 'Your CEU Site verification code is: ' + code +
          '\n\nIt expires in 10 minutes (' + expiryText + '). Do not share it with anyone.' +
          '\nIf you did not request this, ignore this email.' +
          '\n\nCareer Education Unit · Department of Schools',
    htmlBody: htmlBody,
    inlineImages: inlineImages,
    name: 'Career Education Unit Site'
  });
}

/** Step 1: user submits email, we send the code */
function requestLoginOtp(email) {
  var input = String(email || "").trim().toLowerCase();
  var generic = { success: true, message: "If this email is registered, a verification code has been sent." };

  if (!EMAIL_DOMAIN_REGEX.test(input)) {
    return { success: false, message: "Please use your official email (name@school.edu.moe.bn)" };
  }
  if (isLockedOut_(input)) {
    return { success: false, message: "Too many failed attempts. Try again in 15 minutes." };
  }

  var cache = CacheService.getScriptCache();
  if (cache.get("otpcool_" + input)) {
    return { success: false, message: "Please wait a minute before requesting another code." };
  }
  cache.put("otpcool_" + input, "1", OTP_RESEND_COOLDOWN);

  try {
    var user = findUserByEmail_(input);
    if (!user) { logFailedLogin_(input); return generic; } // don't reveal it's unregistered

    var code = generateOtp_();
    var salt = generateSalt_();
    cache.put("otp_" + input, JSON.stringify({ hash: hashPassword_(code, salt), salt: salt }), OTP_TTL_SECONDS);
    cache.remove("otptries_" + input);

    sendOtpEmail_(input, user.name, code);
    return generic;
  } catch (e) {
    Logger.log("OTP send error: " + e);
    return { success: false, message: "Could not send the code. Please try again." };
  }
}

/** Step 2: user submits the code */
function verifyLoginOtp(email, code) {
  var input = String(email || "").trim().toLowerCase();
  var entered = String(code || "").replace(/\D/g, "");
  var cache = CacheService.getScriptCache();

  if (isLockedOut_(input)) {
    return { success: false, message: "Too many failed attempts. Try again in 15 minutes." };
  }
  var raw = cache.get("otp_" + input);
  if (!raw) return { success: false, message: "Code expired or invalid. Please request a new one." };

  var tries = Number(cache.get("otptries_" + input) || 0) + 1;
  cache.put("otptries_" + input, String(tries), OTP_TTL_SECONDS);
  if (tries > OTP_MAX_ATTEMPTS) {
    cache.remove("otp_" + input);
    recordFailedAttempt_(input);
    return { success: false, message: "Too many incorrect codes. Please request a new one." };
  }

  var rec = JSON.parse(raw);
  if (hashPassword_(entered, rec.salt) !== rec.hash) {
    logFailedLogin_(input);
    return { success: false, message: "Incorrect code. Please try again." };
  }

  cache.remove("otp_" + input);        // single use
  cache.remove("otptries_" + input);
  clearFailedAttempts_(input);

  var user = findUserByEmail_(input);
  if (!user) return { success: false, message: "Account not found." };

  logUserActivity(user.username, user.school);
  var token = createSession_(user.username, user.role, user.school);
  return { success: true, username: user.username, role: user.role, school: user.school, token: token };
}

/**
 * 3. ACTIVITY AUDIT LOGGER
 */
function logUserActivity(username, school) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    var logSheet = ss.getSheetByName("Logs");
    if (!logSheet) {
      logSheet = ss.insertSheet("Logs");
      logSheet.appendRow(["Timestamp", "Username", "Status"]);
    }

    // School is looked up from Users!F at report time, not stored per row.
    logSheet.appendRow([new Date(), username, "Logged In Successfully"]);
  } catch(e) {
    Logger.log("Logging failed: " + e.toString());
  }
}

/** Records a failed login attempt in the "FailedLogins" sheet. */
function logFailedLogin_(username) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    var failSheet = ss.getSheetByName("FailedLogins");
    if (!failSheet) {
      failSheet = ss.insertSheet("FailedLogins");
      failSheet.appendRow(["Timestamp", "Username Attempted", "Activity Description"]);
    }

    failSheet.appendRow([new Date(), username, "Failed Login Attempt"]);
  } catch(e) {
    Logger.log("Failed-login logging failed: " + e.toString());
  }
}

/**
 * 4. SYSTEM STAMP / VERSION CONTROL
 * Stamped manually via the "Mark Site as Updated" menu item.
 */
const SYSTEM_BUILD_VERSION_FALLBACK = "16-07-2026 09:11 AM";

function getScriptEditTime() {
  var stamp = PropertiesService.getScriptProperties().getProperty("SITE_LAST_UPDATED");
  return stamp || SYSTEM_BUILD_VERSION_FALLBACK;
}

function markSiteAsUpdated() {
  var now = new Date();
  var formatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "d MMMM yyyy, h:mm a");
  PropertiesService.getScriptProperties().setProperty("SITE_LAST_UPDATED", formatted);

  var summary = "Site marked as updated: " + formatted;
  Logger.log(summary);
  try {
    SpreadsheetApp.getUi().alert(summary);
  } catch (e) {
    // Ran from the editor directly — check View > Executions for the result.
  }
}

/**
 * 5. SITE-WIDE FOOTER VISIT COUNTER
 */
function incrementSiteVisitCount() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    // Couldn't get the lock in time — skip rather than risk a lost update
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var current = Number(props.getProperty("TOTAL_VISITS") || 0) + 1;
    props.setProperty("TOTAL_VISITS", String(current));
    return current;
  } finally {
    try { lock.releaseLock(); } catch (e) { /* lock was never acquired */ }
  }
}

/** Admin-only: resets the all-time visit counter to 0. */
function resetVisitCounter(token) {
  requireAdmin_(token);
  PropertiesService.getScriptProperties().setProperty("TOTAL_VISITS", "0");
  return { success: true, totalVisits: 0 };
}

/** ADMIN TOOL — removes entries older than 6 months from Logs and FailedLogins. */
function adminTrimOldLogs() {
  var months = 6;
  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - months);

  var removedLogs = trimSheetOlderThan_("Logs", cutoff);
  var removedFailed = trimSheetOlderThan_("FailedLogins", cutoff);

  var summary = "Removed " + removedLogs + " row(s) from Logs and " +
                removedFailed + " row(s) from FailedLogins (older than " + months + " months).";
  Logger.log(summary);
  try {
    SpreadsheetApp.getUi().alert(summary);
  } catch (e) {
    // Ran from the editor directly — check View > Executions for the result.
  }
}

function trimSheetOlderThan_(sheetName, cutoffDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  var data = sheet.getDataRange().getValues();
  var removed = 0;

  // Walk bottom-to-top so deleting a row doesn't shift indices still to check
  for (var i = data.length - 1; i >= 1; i--) {
    var timestamp = data[i][0];
    if (timestamp instanceof Date && timestamp < cutoffDate) {
      sheet.deleteRow(i + 1);
      removed++;
    }
  }

  return removed;
}

/**
 * ============================================================
 * 6. DASHBOARD + USER MANAGEMENT
 * Every function re-checks the role on the server, so hiding buttons
 * in the browser is only cosmetic.
 * ============================================================
 */

function roleOf_(session) {
  return (session.role || '').toString().trim().toLowerCase();
}

function requireAdmin_(token) {
  var session = requireSession_(token);
  if (roleOf_(session) !== 'admin') {
    throw new Error('Only an Admin can do this.');
  }
  return session;
}

/** Officer + Admin: manage users and review access requests. */
function requireUserManager_(token) {
  var session = requireSession_(token);
  if (USER_MANAGER_ROLES_.indexOf(roleOf_(session)) === -1) {
    throw new Error('You do not have permission to manage users.');
  }
  return session;
}

/** Officers cannot assign the Admin role or change/delete an existing Admin account. */
function assertAdminRoleRule_(session, newRole, existingRole) {
  if (roleOf_(session) === 'admin') return;
  if (newRole === 'Admin' || existingRole === 'Admin') {
    throw new Error('Only an Admin can assign or change the Admin role.');
  }
}

function requireDashboardAccess_(token) {
  var session = requireSession_(token);
  if (DASHBOARD_ROLES_.indexOf(roleOf_(session)) === -1) {
    throw new Error('You do not have permission to view the dashboard.');
  }
  return session;
}

function getUsersSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) throw new Error("'Users' sheet not found.");
  return sheet;
}

function clearStatsCache_() {
  try { CacheService.getScriptCache().remove(STATS_CACHE_KEY_); } catch (e) {}
}

/** Called from showPage() so the dashboard can report most-viewed pages. */
function logPageView(token, pageId) {
  try {
    var session = requireSession_(token);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('PageViews');
    if (!sheet) {
      sheet = ss.insertSheet('PageViews');
      sheet.appendRow(['Timestamp', 'Username', 'PageId']);
    }
    sheet.appendRow([new Date(), session.username, pageId]);
  } catch (e) {
    Logger.log('Page view logging skipped: ' + e.toString());
  }
}

/* ------------------------------------------------------------
 * DASHBOARD
 * ------------------------------------------------------------ */

/** Main dashboard payload. Cached for 2 minutes; pass force=true to bypass (Refresh button). */
function getDashboardStats(token, force) {
  requireDashboardAccess_(token);
  var cache = CacheService.getScriptCache();

  if (!force) {
    var hit = cache.get(STATS_CACHE_KEY_);
    if (hit) return JSON.parse(hit);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();

  // ---- Users ----
  var usersData = getUsersSheet_().getDataRange().getValues();
  var roleBreakdown = {};
  var userList = [];
  var usernameToSchool = {};
  var teacherCountBySchool = {};
  var schoolDisplayName = {};

  for (var i = 1; i < usersData.length; i++) {
    var row = usersData[i];
    if (!row[0]) continue;
    var username = row[0].toString().trim();
    var role = row[2] ? row[2].toString().trim() : 'Unspecified';
    var school = row[5] ? row[5].toString().trim() : '';
    roleBreakdown[role] = (roleBreakdown[role] || 0) + 1;
    userList.push({ username: username, name: row[1] ? row[1].toString().trim() : '', role: role, school: school });
    usernameToSchool[username.toLowerCase()] = school;

    if (school) {
      var schoolKey = school.toUpperCase();
      schoolDisplayName[schoolKey] = schoolDisplayName[schoolKey] || school;
      teacherCountBySchool[schoolKey] = (teacherCountBySchool[schoolKey] || 0) + 1;
    }
  }

  // ---- Logins ----
  var logsSheet = ss.getSheetByName('Logs');
  var logsData = logsSheet ? logsSheet.getDataRange().getValues() : [];

  var now = new Date();
  var DAY_MS = 24 * 60 * 60 * 1000;
  var last7 = new Date(now.getTime() - 7 * DAY_MS);
  var last30 = new Date(now.getTime() - 30 * DAY_MS);
  var last90 = new Date(now.getTime() - 90 * DAY_MS);
  var todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');

  var lastLoginByUser = {};
  var loginCountBySchool30 = {};   // last 30 days only
  var lastActiveBySchool = {};
  var loginsToday = 0, loginsThisWeek = 0, loginsThisMonth = 0;

  var loginsByDay = {};
  for (var d = 29; d >= 0; d--) {
    loginsByDay[Utilities.formatDate(new Date(now.getTime() - d * DAY_MS), tz, 'yyyy-MM-dd')] = 0;
  }

  for (var j = 1; j < logsData.length; j++) {
    var ts = logsData[j][0];
    var luser = logsData[j][1];
    if (!(ts instanceof Date) || !luser) continue;
    var lkey = luser.toString().trim().toLowerCase();
    if (!lastLoginByUser[lkey] || ts > lastLoginByUser[lkey]) lastLoginByUser[lkey] = ts;

    var tsKey = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
    if (tsKey === todayKey) loginsToday++;
    if (ts >= last7) loginsThisWeek++;
    if (ts >= last30) loginsThisMonth++;
    if (loginsByDay.hasOwnProperty(tsKey)) loginsByDay[tsKey]++;

    var schoolRaw = usernameToSchool[lkey] || '';
    if (!schoolRaw) continue;                       // user deleted or no school set
    var schoolLabel = schoolRaw.toUpperCase();
    schoolDisplayName[schoolLabel] = schoolDisplayName[schoolLabel] || schoolRaw;
    if (ts >= last30) loginCountBySchool30[schoolLabel] = (loginCountBySchool30[schoolLabel] || 0) + 1;
    if (!lastActiveBySchool[schoolLabel] || ts > lastActiveBySchool[schoolLabel]) lastActiveBySchool[schoolLabel] = ts;
  }

  var loginTrend = Object.keys(loginsByDay).sort().map(function (k) {
    return { date: k, count: loginsByDay[k] };
  });

  var schoolActivity = Object.keys(teacherCountBySchool).map(function (k) {
    var last = lastActiveBySchool[k];
    return {
      school: schoolDisplayName[k] || k,
      teacherCount: teacherCountBySchool[k] || 0,
      logins: loginCountBySchool30[k] || 0,
      lastActive: last ? Utilities.formatDate(last, tz, 'd MMM yyyy, h:mm a') : 'Never',
      status: !last ? 'No Activity' : (last >= last7 ? 'Active' : 'Low Activity')
    };
  }).sort(function (a, b) { return b.logins - a.logins; });

  var needsOutreachCount = schoolActivity.filter(function (s) { return s.status !== 'Active'; }).length;

  // ---- User activity ----
  var active7 = 0, active30 = 0, neverLoggedIn = 0;
  var dormantList = [];

  userList.forEach(function (u) {
    var last = lastLoginByUser[u.username.toLowerCase()];
    if (!last) {
      neverLoggedIn++;
      dormantList.push({ username: u.username, name: u.name, role: u.role, school: u.school, lastLogin: 'Never' });
      return;
    }
    if (last >= last7) active7++;
    if (last >= last30) active30++;
    if (last < last90) {
      dormantList.push({ username: u.username, name: u.name, role: u.role, school: u.school, lastLogin: Utilities.formatDate(last, tz, 'd MMM yyyy') });
    }
  });
  dormantList.sort(function (a, b) { return a.school.localeCompare(b.school) || a.username.localeCompare(b.username); });

  var result = {
    totalVisits: Number(PropertiesService.getScriptProperties().getProperty('TOTAL_VISITS') || 0),
    totalUsers: userList.length,
    active7: active7,
    active30: active30,
    neverLoggedIn: neverLoggedIn,
    roleBreakdown: roleBreakdown,
    loginsToday: loginsToday,
    loginsThisWeek: loginsThisWeek,
    loginsThisMonth: loginsThisMonth,
    loginTrend: loginTrend,
    schoolActivity: schoolActivity,
    participatingSchools: schoolActivity.length,
    needsOutreachCount: needsOutreachCount,
    dormantList: dormantList,
    generatedAt: Utilities.formatDate(now, tz, 'd MMM yyyy, h:mm a'),
    lastUpdated: getScriptEditTime()
  };

  try {
    var json = JSON.stringify(result);
    if (json.length < 90000) cache.put(STATS_CACHE_KEY_, json, STATS_CACHE_SECONDS_);
  } catch (e) { /* cache is best-effort */ }

  return result;
}

/** Most-viewed pages leaderboard (client splits it into main pages vs individual pages). */
function getPageViewStats(token) {
  requireDashboardAccess_(token);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PageViews');
  if (!sheet) return { topPages: [], totalViews: 0 };

  var data = sheet.getDataRange().getValues();
  var counts = {};
  for (var i = 1; i < data.length; i++) {
    var pageId = data[i][2];
    if (!pageId) continue;
    counts[pageId] = (counts[pageId] || 0) + 1;
  }
  var topPages = Object.keys(counts)
    .map(function (k) { return { pageId: k, count: counts[k] }; })
    .sort(function (a, b) { return b.count - a.count; })
    .slice(0, 100);

  return { topPages: topPages, totalViews: Math.max(0, data.length - 1) };
}

/* ------------------------------------------------------------
 * USER MANAGEMENT (Officer + Admin)
 * ------------------------------------------------------------ */

/** Full user list with name, role, school and last login. */
function getUserManagementData(token) {
  requireUserManager_(token);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var usersData = getUsersSheet_().getDataRange().getValues();
  var logsSheet = ss.getSheetByName('Logs');
  var logsData = logsSheet ? logsSheet.getDataRange().getValues() : [];
  var tz = Session.getScriptTimeZone();

  var lastLoginByUser = {};
  for (var j = 1; j < logsData.length; j++) {
    var ts = logsData[j][0];
    var luser = logsData[j][1];
    if (!(ts instanceof Date) || !luser) continue;
    var lkey = luser.toString().trim().toLowerCase();
    if (!lastLoginByUser[lkey] || ts > lastLoginByUser[lkey]) lastLoginByUser[lkey] = ts;
  }

  var users = [];
  for (var i = 1; i < usersData.length; i++) {
    var row = usersData[i];
    if (!row[0]) continue;
    var username = row[0].toString().trim();
    var last = lastLoginByUser[username.toLowerCase()];
    users.push({
      username: username,
      name: row[1] ? row[1].toString().trim() : '',
      role: row[2] ? row[2].toString().trim() : 'Unspecified',
      school: row[5] ? row[5].toString().trim() : '',
      lastLogin: last ? Utilities.formatDate(last, tz, 'd MMM yyyy, h:mm a') : 'Never'
    });
  }
  users.sort(function (a, b) { return a.username.localeCompare(b.username); });
  return users;
}

function addUserRow_(email, name, role, school) {
  var clean = String(email || '').trim().toLowerCase();
  var cleanRole = String(role || '').trim();
  if (!EMAIL_DOMAIN_REGEX.test(clean)) throw new Error('A valid school email (name@school.edu.moe.bn) is required.');
  if (ADMIN_VALID_ROLES_.indexOf(cleanRole) === -1) throw new Error('Please select a valid role.');
  if (findUserByEmail_(clean)) throw new Error('That email already exists.');
  getUsersSheet_().appendRow([
    clean,
    name ? String(name).trim().slice(0, 120) : '',
    cleanRole,
    '', '',
    school ? String(school).trim().slice(0, 120) : ''
  ]);
  return clean;
}

function adminAddUser(token, email, role, school, name) {
  var session = requireUserManager_(token);
  assertAdminRoleRule_(session, String(role || '').trim(), '');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var clean = addUserRow_(email, name, role, school);
    clearStatsCache_();
    return { success: true, message: clean + ' added successfully.' };
  } finally {
    lock.releaseLock();
  }
}

/** Edit name, role and school in one go. */
function adminUpdateUser(token, username, name, role, school) {
  var session = requireUserManager_(token);
  var target = String(username || '').trim().toLowerCase();
  var cleanRole = String(role || '').trim();
  if (!target) throw new Error('Username is required.');
  if (ADMIN_VALID_ROLES_.indexOf(cleanRole) === -1) throw new Error('Please select a valid role.');

  var sheet = getUsersSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === target) {
      assertAdminRoleRule_(session, cleanRole, String(data[i][2] || '').trim());
      if (target === session.username.toString().trim().toLowerCase() && roleOf_(session) === 'admin' && cleanRole !== 'Admin') {
        throw new Error("You can't remove your own Admin role.");
      }
      sheet.getRange(i + 1, 2).setValue(name ? String(name).trim().slice(0, 120) : '');
      sheet.getRange(i + 1, 3).setValue(cleanRole);
      sheet.getRange(i + 1, 6).setValue(school ? String(school).trim().slice(0, 120) : '');
      clearStatsCache_();
      return { success: true, message: data[i][0] + ' updated.' };
    }
  }
  throw new Error('User not found.');
}

function adminDeleteUser(token, username) {
  var session = requireUserManager_(token);
  var target = String(username || '').trim().toLowerCase();
  if (!target) throw new Error('Username is required.');
  if (target === session.username.toString().trim().toLowerCase()) {
    throw new Error("You can't delete your own account while logged in.");
  }

  var sheet = getUsersSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === target) {
      assertAdminRoleRule_(session, '', String(data[i][2] || '').trim());
      var removed = data[i][0];
      sheet.deleteRow(i + 1);
      clearStatsCache_();
      return { success: true, message: removed + ' was deleted.' };
    }
  }
  throw new Error('User not found.');
}

/* ------------------------------------------------------------
 * ACCESS REQUESTS (Officer + Admin) - the queue behind the Register form
 * Requests sheet columns: A Timestamp | B Type | C Name | D Role | E School | F Contact Email | G Phone | H Status
 * ------------------------------------------------------------ */

function getAccessRequests(token) {
  requireUserManager_(token);
  var sh = SpreadsheetApp.openById(REQUESTS_SHEET_ID).getSheetByName(REQUESTS_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  var tz = Session.getScriptTimeZone();
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  var out = [];
  values.forEach(function (v, idx) {
    if (!v[2] && !v[5]) return;
    out.push({
      row: idx + 2,
      time: v[0] instanceof Date ? Utilities.formatDate(v[0], tz, 'd MMM yyyy, h:mm a') : String(v[0]),
      type: String(v[1]),
      name: String(v[2]),
      role: String(v[3]),
      school: String(v[4]),
      contact: String(v[5]),
      phone: String(v[6]),
      status: String(v[7] || 'New')
    });
  });
  return out.reverse();   // newest first
}

/** action: 'approve' (creates the user for registration requests) or 'reject'. */
function adminReviewRequest(token, row, action) {
  var session = requireUserManager_(token);
  var r = Number(row);
  if (!(r >= 2)) throw new Error('Invalid request.');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = SpreadsheetApp.openById(REQUESTS_SHEET_ID).getSheetByName(REQUESTS_TAB);
    if (!sh || r > sh.getLastRow()) throw new Error('Request not found.');
    var v = sh.getRange(r, 1, 1, 8).getValues()[0];
    if (String(v[7] || 'New').trim() !== 'New') throw new Error('This request was already handled.');

    if (action === 'reject') {
      sh.getRange(r, 8).setValue('Rejected');
      return { success: true, message: 'Request rejected.' };
    }
    if (action !== 'approve') throw new Error('Unknown action.');

    if (String(v[1]) !== 'Registration Request') {
      sh.getRange(r, 8).setValue('Resolved');
      return { success: true, message: 'Marked as resolved.' };
    }

    assertAdminRoleRule_(session, String(v[3] || '').trim(), '');
    var email = addUserRow_(v[5], v[2], v[3], v[4]);
    sh.getRange(r, 8).setValue('Approved');
    clearStatsCache_();

    try {
      // 1. Load the HTML template
      var htmlTemplate = HtmlService.createTemplateFromFile('RegistrationApprovedEmail');

      // 2. Pass dynamic variables to the template
      htmlTemplate.userName = v[2]; // User's name
      htmlTemplate.hasLogo = false; // Set to true if attaching an inline logo via inlineImages

      // 3. Render HTML output
      var htmlBody = htmlTemplate.evaluate().getContent();

      // 4. Plain text fallback for email clients that don't render HTML
      var textBody = 'Hello ' + v[2] + ',\n\n' +
        'Your CEU Site account has been created. Open the site, enter your registered school email and use the verification code sent to you to sign in.\n\n' +
        'Career Education Unit · Department of Schools · Ministry of Education, Brunei Darussalam';

      MailApp.sendEmail({
        to: email,
        subject: 'Your CEU Site account is ready',
        body: textBody,
        htmlBody: htmlBody,
        name: 'Career Education Unit Site'
      });

    } catch (mailErr) {
      Logger.log('Approval email not sent: ' + mailErr);
    }

    return { success: true, message: email + ' approved and added.' };

  } finally {
    lock.releaseLock();
  }
}

/** Clears all login logs, keeping headers. Run from the editor. */
function resetDashboardStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ss.getSheetByName("Logs");
  if (!logsSheet) throw new Error("Logs sheet not found.");

  var lastRow = logsSheet.getLastRow();
  if (lastRow > 1) {
    var lastCol = logsSheet.getLastColumn();
    logsSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
  clearStatsCache_();
}

/**
 * Run this ONCE from the Apps Script editor to grant permissions
 * (Mail, Spreadsheet, and Drive for the OTP email logo).
 */
function authorizeAll() {
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'CEU Site test', 'Mail permission works.');
  SpreadsheetApp.openById(REQUESTS_SHEET_ID).getSheets();
  DriveApp.getFileById(EMAIL_LOGO_FILE_ID).getName();
}

/**
 * Run this from the editor to preview the OTP email in your own inbox
 * without going through the login page.
 */
function testOtpEmail() {
  sendOtpEmail_(Session.getEffectiveUser().getEmail(), 'Hajah', '652521');
}

/**
 * ============================================================
 * 7. ACCESS REQUESTS (Register form on login page)
 * ============================================================
 */
function submitAccessRequest(data) {
  var clean = function(s) { return String(s || '').replace(/[<>]/g, '').trim().slice(0, 200); };
  var isRegister = data && data.type === 'register';
  var type    = isRegister ? 'Registration Request' : 'Sign-in Help Request';
  var name    = clean(data && data.name);
  var role    = isRegister ? clean(data.role) : '';
  var school  = clean(data && data.school);
  var contact = clean(data && data.contact);
  var phone   = clean(data && data.phone);

  var phoneDigits = phone.replace(/\D/g, '');
  var phoneOk = /^[0-9+\-()\s]+$/.test(phone) && phoneDigits.length >= 7 && phoneDigits.length <= 15;

  if (!name || !school || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(contact) || !phoneOk || (isRegister && !role)) {
    throw new Error('Please fill in all fields with a valid email and phone number.');
  }
  if (isRegister && ADMIN_VALID_ROLES_.indexOf(role) === -1) {
    throw new Error('Please select a valid role.');
  }

  var stamp = new Date();

  // 1. Log to the Sheet first, so nothing is lost if the email fails
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.openById(REQUESTS_SHEET_ID);
    var sh = ss.getSheetByName(REQUESTS_TAB);
    if (!sh) {
      sh = ss.insertSheet(REQUESTS_TAB);
      sh.appendRow(['Timestamp', 'Type', 'Name', 'Role', 'School', 'Contact Email', 'Phone', 'Status']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([stamp, type, name, role, school, contact, '', 'New']);
    sh.getRange(sh.getLastRow(), 7).setNumberFormat('@').setValue(phone);
  } finally {
    lock.releaseLock();
  }

  // 2. Email the UPK inbox
  var body = [
    'Type: ' + type,
    'Name: ' + name,
    isRegister ? 'Role: ' + role : null,
    'School: ' + school,
    'Contact email: ' + contact,
    'Phone: ' + phone,
    'Submitted: ' + stamp.toLocaleString('en-GB', { timeZone: 'Asia/Brunei' })
  ].filter(Boolean).join('\n');

  MailApp.sendEmail({
    to: REQUEST_INBOX,
    replyTo: contact,
    subject: '[CEU Site] ' + type + ' - ' + name,
    body: body
  });

  return true;
}