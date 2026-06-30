/**
 * Google Apps Script — TEST VERSION for Yoga Signup
 *
 * This is the TEST deployment of the Apps Script backend.
 * It is identical to the production script EXCEPT:
 *
 *   1. Uses "Test Yoga Signup" and "Test Yoga Waitlist" spreadsheets
 *   2. Logs emails to "Test Email Log" sheet instead of calling MailApp.sendEmail()
 *   3. Exposes test-only endpoints:
 *      - GET  ?action=read_sheet&sheet=SheetName  — Read all rows from a sheet
 *      - GET  ?action=cleanup                     — Wipe all data rows (keep headers)
 *      - GET  ?action=trigger_archive             — Run archive function on demand
 *      - POST action=seed                         — Insert test data rows directly
 *
 * SETUP:
 *   1. Create a NEW Apps Script project (separate from production)
 *   2. Paste this entire file
 *   3. Deploy as Web app (Execute as: Me, Access: Anyone)
 *   4. Create two Google Sheets manually:
 *      - "Test Yoga Signup"   (with a "Sign-Ups" tab and "Archive" tab)
 *      - "Test Yoga Waitlist" (with a "Waitlist" tab and "Waitlist Archive" tab)
 *      OR let the script auto-create them on first run
 *   5. Copy the deployment URL into test.config.js (or set TEST_APPS_SCRIPT_URL env var)
 */

// ========== TEST CONFIGURATION ==========
var TEST_SIGNUP_SS_NAME   = 'Test Yoga Signup';
var TEST_WAITLIST_SS_NAME = 'Test Yoga Waitlist';
var SITE_URL = 'https://yogawithjessica.com';
var ADMIN_EMAIL = 'xiaojing25@gmail.com';

// ============================================================================
// ========== ADMIN CONSOLE AUTH (Google ID-token verification) ==========
// ============================================================================
// Mirror of production. The admin page signs in with Google Identity Services,
// gets an ID token (JWT), and sends it with each admin request. We verify it
// server-side via Google's tokeninfo endpoint (Google checks the signature +
// expiry), then assert aud === our client ID, iss === Google, email_verified,
// and that the email is on the allowlist. NO shared secret. See verifyAdminToken_().
//
// PASTE the OAuth "Web application" Client ID here (same value the admin page
// uses). It is NOT a secret. For local/test the admin page can point its
// SHEETS_WEB_APP_URL at the test deployment instead.
var ADMIN_OAUTH_CLIENT_ID = '83041676087-iia2s4jjtb3n6je56so3mfbdin9lpe0u.apps.googleusercontent.com';
var ADMIN_ALLOWLIST = ['liu.jessica.x@gmail.com'];

// ============================================================================
// ========== SCHEDULE — SINGLE SOURCE OF TRUTH (Google Sheet) ==========
// ============================================================================
// Mirror of the production schedule module. Reads the "Schedule" and
// "Exceptions" tabs of the Test Yoga Signup spreadsheet. Run setupScheduleSheets()
// once to create + seed them.

var MEET_TZ = 'America/Los_Angeles';
var SCHEDULE_SHEET_NAME   = 'Schedule';
var EXCEPTIONS_SHEET_NAME = 'Exceptions';
var SCHEDULE_CACHE_KEY    = 'yoga_schedule_test_v1';

function setupScheduleSheets() {
  var ss = getOrCreateSpreadsheet();
  var schedule = getOrCreateScheduleSheet_(ss);
  getOrCreateExceptionsSheet_(ss);

  if (schedule.getLastRow() <= 1) {
    var seed = [
      ['sunday-online', 'Sunday Evening \u2014 Online via Zoom', 0, '18:00', 75, 'online', '', '', 'Open to Everyone', 'Yoga mat, Strap, Two blocks, Wall space, Yoga chair (ideal), Bolster (ideal)', 'TRUE'],
      ['tuesday-ccv', 'Tuesday Evening \u2014 CCV Clubhouse (In Person)', 2, '18:00', 75, 'inperson', 'CCV Clubhouse', 10, 'CCV Residents Only, In Person', 'Yoga mat, Two blocks, Strap', 'TRUE'],
      ['wednesday-restorative', 'Wednesday Evening \u2014 Restorative Yoga (Online)', 3, '20:00', 75, 'online', '', '', 'Restorative, Open to Everyone', 'Yoga mat, Bolster, Two blocks, Two blankets, Strap, Wall space, Yoga chair (ideal)', 'TRUE']
    ];
    var range = schedule.getRange(2, 1, seed.length, 11);
    range.setNumberFormat('@');
    range.setValues(seed);
  }

  bustScheduleCache();
  Logger.log('Schedule sheets ready. Class rows: ' + (schedule.getLastRow() - 1));
}

function getOrCreateScheduleSheet_(ss) {
  var sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SCHEDULE_SHEET_NAME);
    var headers = ['ID', 'Label', 'Day (0=Sun)', 'Start Time', 'Duration (min)', 'Type', 'Location', 'Capacity', 'Tags', 'Props', 'Active', 'One-Off Date'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(2, 1, 2000, headers.length).setNumberFormat('@');
  } else if (!(sheet.getRange(1, 12).getValue() || '').toString().trim()) {
    // Auto-migrate an existing 11-column sheet: add the One-Off Date header +
    // text format so older deployments don't need a manual schema change.
    sheet.getRange(1, 12).setValue('One-Off Date').setFontWeight('bold');
    sheet.getRange(2, 12, 2000, 1).setNumberFormat('@');
  }
  return sheet;
}

function getOrCreateExceptionsSheet_(ss) {
  var sheet = ss.getSheetByName(EXCEPTIONS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(EXCEPTIONS_SHEET_NAME);
    var headers = ['Class ID', 'Date (YYYY-MM-DD)', 'Status', 'New Date', 'New Start Time', 'Note'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(2, 1, 2000, headers.length).setNumberFormat('@');
  }
  return sheet;
}

function readScheduleFromSheet_() {
  var ss = getOrCreateSpreadsheet();
  var classes = [];
  var exceptions = [];

  var sSheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (sSheet && sSheet.getLastRow() > 1) {
    var sData = sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 12).getValues();
    for (var i = 0; i < sData.length; i++) {
      var row = sData[i];
      var id = (row[0] || '').toString().trim();
      if (!id) continue;
      var t = parseHM_(row[3]);
      classes.push({
        id: id,
        label: (row[1] || '').toString().trim(),
        day: parseInt(row[2], 10) || 0,
        startTime: pad2_(t.h) + ':' + pad2_(t.m),
        startH: t.h,
        startM: t.m,
        durationMins: parseInt(row[4], 10) || 75,
        type: (row[5] || 'online').toString().trim().toLowerCase(),
        location: (row[6] || '').toString().trim(),
        capacity: parseCapacity_(row[7]),
        tags: splitList_(row[8]),
        props: splitList_(row[9]),
        active: parseBool_(row[10]),
        // Col 12: blank = recurring weekly; a date = single non-recurring occurrence.
        oneOffDate: normalizeIso_(row[11])
      });
    }
  }

  var eSheet = ss.getSheetByName(EXCEPTIONS_SHEET_NAME);
  if (eSheet && eSheet.getLastRow() > 1) {
    var eData = eSheet.getRange(2, 1, eSheet.getLastRow() - 1, 6).getValues();
    for (var j = 0; j < eData.length; j++) {
      var er = eData[j];
      var cid = (er[0] || '').toString().trim();
      var dt = normalizeIso_(er[1]);
      if (!cid || !dt) continue;
      exceptions.push({
        classId: cid,
        date: dt,
        status: (er[2] || '').toString().trim().toLowerCase(),
        newDate: normalizeIso_(er[3]),
        newStartTime: normalizeTimeStr_(er[4]),
        note: (er[5] || '').toString()
      });
    }
  }

  return { classes: classes, exceptions: exceptions };
}

function getSchedule() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(SCHEDULE_CACHE_KEY);
    if (cached) return JSON.parse(cached);
    var data = readScheduleFromSheet_();
    cache.put(SCHEDULE_CACHE_KEY, JSON.stringify(data), 60);
    return data;
  } catch (e) {
    Logger.log('getSchedule error: ' + e);
    return readScheduleFromSheet_();
  }
}

function bustScheduleCache() {
  try { CacheService.getScriptCache().remove(SCHEDULE_CACHE_KEY); } catch (e) {}
}

function getScheduleClassByLabel_(label) {
  var sched = getSchedule();
  for (var i = 0; i < sched.classes.length; i++) {
    if (sched.classes[i].label === label) return sched.classes[i];
  }
  return null;
}

function getOccurrencesOnPacificDate_(pstNow) {
  var sched = getSchedule();
  var iso = Utilities.formatDate(pstNow, MEET_TZ, 'yyyy-MM-dd');
  var display = formatClassDate(pstNow);
  var weekday = pstNow.getDay();
  var occ = [];
  var byId = {};
  sched.classes.forEach(function(c) { byId[c.id] = c; });

  // Recurring occurrences for this weekday — or a one-off class on its single
  // date (unless cancelled or moved away).
  sched.classes.forEach(function(c) {
    if (!c.active) return;
    var occursToday = c.oneOffDate ? (c.oneOffDate === iso) : (c.day === weekday);
    if (!occursToday) return;
    var ex = findExceptionForDate_(sched.exceptions, c.id, iso);
    if (ex && (ex.status === 'cancelled' || ex.status === 'moved')) return;
    occ.push(makeOccurrence_(c, c.startH, c.startM, display));
  });

  // Skip inactive classes here too, so deactivating/cancelling a series also
  // hides any pending moved-in/extra date (mirrors the client expander).
  sched.exceptions.forEach(function(ex) {
    var c = byId[ex.classId];
    if (!c || !c.active) return;
    if (ex.status === 'moved' && ex.newDate === iso) {
      var t = ex.newStartTime ? parseHM_(ex.newStartTime) : { h: c.startH, m: c.startM };
      occ.push(makeOccurrence_(c, t.h, t.m, display));
    } else if (ex.status === 'extra' && ex.date === iso) {
      var t2 = ex.newStartTime ? parseHM_(ex.newStartTime) : { h: c.startH, m: c.startM };
      occ.push(makeOccurrence_(c, t2.h, t2.m, display));
    }
  });

  return occ;
}

function makeOccurrence_(c, h, m, display) {
  return {
    id: c.id, label: c.label, type: c.type,
    startH: h, startM: m, durationMins: c.durationMins,
    capacity: c.capacity, location: c.location, classDate: display
  };
}

function findExceptionForDate_(exceptions, classId, iso) {
  for (var i = 0; i < exceptions.length; i++) {
    if (exceptions[i].classId === classId && exceptions[i].date === iso) return exceptions[i];
  }
  return null;
}

function pad2_(n) { n = parseInt(n, 10) || 0; return (n < 10 ? '0' : '') + n; }

function parseHM_(v) {
  if (v instanceof Date) return { h: v.getHours(), m: v.getMinutes() };
  if (typeof v === 'number' && v > 0 && v < 1) { var mins = Math.round(v * 1440); return { h: Math.floor(mins / 60), m: mins % 60 }; }
  var s = String(v == null ? '' : v).trim();
  var m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) return { h: parseInt(m[1], 10), m: parseInt(m[2], 10) };
  var n = parseInt(s, 10);
  if (!isNaN(n) && n >= 0 && n <= 23) return { h: n, m: 0 };
  return { h: 18, m: 0 };
}

function parseCapacity_(v) {
  if (v === '' || v == null) return null;
  var n = parseInt(v, 10);
  return isNaN(n) ? null : n;
}

function splitList_(v) {
  return String(v == null ? '' : v).split(',').map(function(x) { return x.trim(); }).filter(function(x) { return x.length; });
}

function parseBool_(v) {
  if (v === '' || v == null) return true;
  if (v === true) return true;
  if (v === false) return false;
  var s = String(v).trim().toLowerCase();
  return !(s === 'false' || s === 'no' || s === '0');
}

function normalizeIso_(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) return Utilities.formatDate(v, MEET_TZ, 'yyyy-MM-dd');
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad2_(+m[2]) + '-' + pad2_(+m[3]);
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, MEET_TZ, 'yyyy-MM-dd');
  return '';
}

function normalizeTimeStr_(v) {
  if (v == null || v === '') return '';
  var t = parseHM_(v);
  return pad2_(t.h) + ':' + pad2_(t.m);
}

// ============================================================================
// ========== ADMIN CONSOLE — server-side auth + read endpoint ==========
// ============================================================================

// Small JSON responder used by the admin endpoints.
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Verify a Google ID token (JWT) and return the verified, allowlisted admin
// email, or null. Google's tokeninfo validates the signature + expiry; we then
// assert aud === our client ID, iss === Google, email_verified, not-expired,
// and email on ADMIN_ALLOWLIST. The decision is cached by a hash of the token
// for the token's remaining life so repeated actions don't re-hit tokeninfo.
function verifyAdminToken_(idToken) {
  if (!idToken || typeof idToken !== 'string') return null;

  if (!ADMIN_OAUTH_CLIENT_ID || ADMIN_OAUTH_CLIENT_ID.indexOf('PASTE_') === 0) {
    Logger.log('verifyAdminToken_: ADMIN_OAUTH_CLIENT_ID not configured');
    return null;
  }

  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'admin_tok_' + Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));

    var cached = cache.get(cacheKey);
    if (cached) {
      var c = JSON.parse(cached);
      if (c.exp && (c.exp * 1000) > Date.now()) return c.email;
    }

    var resp = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;

    var claims = JSON.parse(resp.getContentText());

    if (claims.aud !== ADMIN_OAUTH_CLIENT_ID) return null;
    if (claims.iss !== 'accounts.google.com' && claims.iss !== 'https://accounts.google.com') return null;
    var email = (claims.email || '').toString().trim().toLowerCase();
    var emailVerified = (claims.email_verified === true || claims.email_verified === 'true');
    if (!email || !emailVerified) return null;
    var exp = parseInt(claims.exp, 10);
    if (!exp || (exp * 1000) <= Date.now()) return null;
    var allowed = false;
    for (var i = 0; i < ADMIN_ALLOWLIST.length; i++) {
      if (String(ADMIN_ALLOWLIST[i]).trim().toLowerCase() === email) { allowed = true; break; }
    }
    if (!allowed) return null;

    var ttl = Math.min(21600, (exp - Math.floor(Date.now() / 1000)) - 30);
    if (ttl > 0) cache.put(cacheKey, JSON.stringify({ email: email, exp: exp }), ttl);

    return email;
  } catch (err) {
    Logger.log('verifyAdminToken_ error: ' + err);
    return null;
  }
}

// Build the admin read payload: full schedule (incl. inactive classes) + the
// next-7-days occurrences with confirmed sign-up and waitlist counts.
function buildAdminSchedulePayload_(adminEmail) {
  try {
    var sched = getSchedule();
    return {
      status: 'ok',
      adminEmail: adminEmail,
      classes: sched.classes,
      exceptions: sched.exceptions,
      upcoming: buildUpcomingAdminOccurrences_(7)
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// Expand occurrences over the next windowDays Pacific days (active + exceptions
// via getOccurrencesOnPacificDate_) and attach sign-up/waitlist counts.
function buildUpcomingAdminOccurrences_(windowDays) {
  var out = [];
  var exceptions = getSchedule().exceptions;
  var now = new Date();
  var pstNow = new Date(now.toLocaleString('en-US', { timeZone: MEET_TZ }));
  var currentTotalMin = pstNow.getHours() * 60 + pstNow.getMinutes();

  for (var offset = 0; offset <= windowDays; offset++) {
    var day = new Date(pstNow.getFullYear(), pstNow.getMonth(), pstNow.getDate() + offset);
    var iso = Utilities.formatDate(day, MEET_TZ, 'yyyy-MM-dd');
    var occ = getOccurrencesOnPacificDate_(day);
    for (var i = 0; i < occ.length; i++) {
      var o = occ[i];
      if (offset === 0 && ((o.startH * 60 + o.startM) - currentTotalMin) < -15) continue;
      var counts = countSignupsForClass_(o.label, o.classDate);
      out.push({
        classId: o.id,
        label: o.label,
        type: o.type,
        date: o.classDate,
        iso: iso,
        startH: o.startH,
        startM: o.startM,
        durationMins: o.durationMins,
        capacity: o.capacity,
        signedUp: counts.signedUp,
        waitlist: counts.waitlist,
        origin: occurrenceOrigin_(exceptions, o.id, iso)
      });
    }
  }

  // Far-future one-off classes can collect signups before the 7-day window, so
  // surface them here too (with counts + Move/Cancel). We use each one-off's
  // EFFECTIVE date(s): the base date unless it's cancelled or moved away, PLUS
  // any moved-in / extra date — so a one-off moved to a far-future date still
  // shows on its new date. Dates already within the 7-day scan above (or duplicates)
  // are skipped.
  var windowEndIso = Utilities.formatDate(
    new Date(pstNow.getFullYear(), pstNow.getMonth(), pstNow.getDate() + windowDays), MEET_TZ, 'yyyy-MM-dd');
  var seenAdminOcc = {};
  out.forEach(function (o) { seenAdminOcc[o.classId + '|' + o.iso] = true; });
  var allClasses = getSchedule().classes;
  for (var k = 0; k < allClasses.length; k++) {
    var oc = allClasses[k];
    if (!oc.active || !oc.oneOffDate) continue;

    // Build this one-off's effective dates.
    var ocDates = [];
    var baseEx = findExceptionForDate_(exceptions, oc.id, oc.oneOffDate);
    if (!(baseEx && (baseEx.status === 'cancelled' || baseEx.status === 'moved'))) {
      ocDates.push({ iso: oc.oneOffDate, startH: oc.startH, startM: oc.startM });
    }
    for (var e = 0; e < exceptions.length; e++) {
      var ex = exceptions[e];
      if (ex.classId !== oc.id) continue;
      if (ex.status === 'moved' && ex.newDate) {
        var mt = ex.newStartTime ? parseHM_(ex.newStartTime) : { h: oc.startH, m: oc.startM };
        ocDates.push({ iso: ex.newDate, startH: mt.h, startM: mt.m });
      } else if (ex.status === 'extra') {
        var xt = ex.newStartTime ? parseHM_(ex.newStartTime) : { h: oc.startH, m: oc.startM };
        ocDates.push({ iso: ex.date, startH: xt.h, startM: xt.m });
      }
    }

    for (var di = 0; di < ocDates.length; di++) {
      var od = ocDates[di];
      if (od.iso <= windowEndIso) continue;               // within window → covered by the scan above
      if (seenAdminOcc[oc.id + '|' + od.iso]) continue;   // dedup
      seenAdminOcc[oc.id + '|' + od.iso] = true;
      var odDisp = displayDateFromIso_(od.iso);
      var odCounts = countSignupsForClass_(oc.label, odDisp);
      out.push({
        classId: oc.id,
        label: oc.label,
        type: oc.type,
        date: odDisp,
        iso: od.iso,
        startH: od.startH,
        startM: od.startM,
        durationMins: oc.durationMins,
        capacity: oc.capacity,
        signedUp: odCounts.signedUp,
        waitlist: odCounts.waitlist,
        origin: occurrenceOrigin_(exceptions, oc.id, od.iso)
      });
    }
  }

  out.sort(function (a, b) {
    if (a.iso !== b.iso) return a.iso < b.iso ? -1 : 1;
    return (a.startH * 60 + a.startM) - (b.startH * 60 + b.startM);
  });
  return out;
}

// Classify how an occurrence on `iso` arose (recurring / moved / extra) so the
// admin UI only offers Move on plain recurring dates.
function occurrenceOrigin_(exceptions, classId, iso) {
  for (var i = 0; i < exceptions.length; i++) {
    var ex = exceptions[i];
    if (ex.classId !== classId) continue;
    if (ex.status === 'extra' && ex.date === iso) return 'extra';
    if (ex.status === 'moved' && ex.newDate === iso) return 'moved';
  }
  return 'recurring';
}

// Count confirmed sign-ups (incl. guests) and waitlist entries for a class
// label + display date. Uses the TEST spreadsheets.
function countSignupsForClass_(className, classDate) {
  var signedUp = 0, waitlist = 0;
  var targetMD = extractMonthDay_(classDate);

  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
      for (var i = 0; i < data.length; i++) {
        var rowClass = (data[i][4] || '').toString().trim();
        var rowDate  = (data[i][5] || '').toString().trim();
        if (rowClass === className && (rowDate === classDate || extractMonthDay_(rowDate) === targetMD)) {
          signedUp++;
          if ((data[i][8] || '').toString().trim()) signedUp++;
        }
      }
    }
  } catch (err) {
    Logger.log('countSignupsForClass_ signups error: ' + err);
  }

  try {
    var wfiles = DriveApp.getFilesByName(TEST_WAITLIST_SS_NAME);
    if (wfiles.hasNext()) {
      var wss = SpreadsheetApp.open(wfiles.next());
      var wsheet = wss.getSheetByName('Waitlist');
      if (wsheet && wsheet.getLastRow() > 1) {
        var wdata = wsheet.getRange(2, 1, wsheet.getLastRow() - 1, wsheet.getLastColumn()).getValues();
        for (var j = 0; j < wdata.length; j++) {
          var wClass = (wdata[j][4] || '').toString().trim();
          var wDate  = (wdata[j][5] || '').toString().trim();
          if (wClass === className && (wDate === classDate || extractMonthDay_(wDate) === targetMD)) {
            waitlist++;
            if ((wdata[j][7] || '').toString().trim()) waitlist++;
          }
        }
      }
    }
  } catch (err) {
    Logger.log('countSignupsForClass_ waitlist error: ' + err);
  }

  return { signedUp: signedUp, waitlist: waitlist };
}

// ============================================================================
// ========== ADMIN CONSOLE — write actions (token-gated) ==========
// ============================================================================
// Mirror of production. Each runs only after verifyAdminToken_ passed and busts
// the schedule cache. Phase 2: schedule edits only — no student emails.

function adminUpdateClass_(data) {
  try {
    var norm = normalizeAdminClass_(data['class'], { requireId: true });
    if (!norm.ok) return { status: 'error', message: norm.error };
    var c = norm.value;

    var ss = getOrCreateSpreadsheet();
    var sheet = getOrCreateScheduleSheet_(ss);
    var rowIdx = findScheduleRowById_(sheet, c.id);
    if (rowIdx < 0) return { status: 'error', message: 'Class not found: ' + c.id };

    var oldLabel = (sheet.getRange(rowIdx, 2).getValue() || '').toString().trim();
    var relabeled = 0;
    if (oldLabel && oldLabel !== c.label) {
      relabeled = relabelSignups_(oldLabel, c.label);
    }

    writeScheduleRow_(sheet, rowIdx, c);
    bustScheduleCache();
    return { status: 'ok', id: c.id, relabeled: relabeled };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function adminCreateClass_(data) {
  try {
    var norm = normalizeAdminClass_(data['class'], { requireId: false });
    if (!norm.ok) return { status: 'error', message: norm.error };
    var c = norm.value;

    var ss = getOrCreateSpreadsheet();
    var sheet = getOrCreateScheduleSheet_(ss);
    c.id = uniqueClassId_(sheet, slugify_(c.label) || 'class');

    writeScheduleRow_(sheet, sheet.getLastRow() + 1, c);
    bustScheduleCache();
    return { status: 'ok', id: c.id };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function adminSetActive_(data) {
  try {
    var id = (data.id || '').toString().trim();
    if (!id) return { status: 'error', message: 'Missing class id' };
    var active = (data.active === true || String(data.active).toLowerCase() === 'true');

    var ss = getOrCreateSpreadsheet();
    var sheet = getOrCreateScheduleSheet_(ss);
    var rowIdx = findScheduleRowById_(sheet, id);
    if (rowIdx < 0) return { status: 'error', message: 'Class not found: ' + id };

    var cell = sheet.getRange(rowIdx, 11);
    cell.setNumberFormat('@');
    cell.setValue(active ? 'TRUE' : 'FALSE');
    bustScheduleCache();
    return { status: 'ok', id: id, active: active };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function adminSetException_(data) {
  try {
    var classId = (data.classId || '').toString().trim();
    var dateIso = normalizeIso_(data.date);
    var status  = (data.status || 'moved').toString().trim().toLowerCase();
    if (!classId) return { status: 'error', message: 'Missing classId' };
    if (!dateIso) return { status: 'error', message: 'Missing or invalid date' };
    if (status !== 'cancelled' && status !== 'moved' && status !== 'extra') {
      return { status: 'error', message: 'Invalid status' };
    }

    var newDate  = normalizeIso_(data.newDate);
    var newStart = normalizeTimeStr_(data.newStartTime);
    if (status === 'moved' && !newDate) {
      return { status: 'error', message: 'A move needs a new date' };
    }

    var ss = getOrCreateSpreadsheet();
    if (findScheduleRowById_(getOrCreateScheduleSheet_(ss), classId) < 0) {
      return { status: 'error', message: 'Class not found: ' + classId };
    }

    var exSheet = getOrCreateExceptionsSheet_(ss);
    var rowIdx = findExceptionRow_(exSheet, classId, dateIso);
    if (rowIdx < 0) rowIdx = exSheet.getLastRow() + 1;
    var range = exSheet.getRange(rowIdx, 1, 1, 6);
    range.setNumberFormat('@');
    range.setValues([[classId, dateIso, status, newDate, newStart, (data.note || '').toString()]]);

    bustScheduleCache();
    return { status: 'ok', classId: classId, date: dateIso, newDate: newDate, newStartTime: newStart };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ---- admin write helpers ----

function normalizeAdminClass_(raw, opts) {
  opts = opts || {};
  if (!raw || typeof raw !== 'object') return { ok: false, error: 'Missing class data' };

  var label = (raw.label || '').toString().trim();
  if (!label) return { ok: false, error: 'Class name is required' };

  // A one-off carries a single date; the weekday is derived from it (and stored
  // for display) but the expander keys off the date, not the day. A recurring
  // class needs a valid weekday and no one-off date.
  var oneOffDate = normalizeIso_(raw.oneOffDate);
  var day;
  if (oneOffDate) {
    var dp = oneOffDate.split('-');
    day = new Date(+dp[0], +dp[1] - 1, +dp[2]).getDay();
  } else {
    day = parseInt(raw.day, 10);
    if (isNaN(day) || day < 0 || day > 6) return { ok: false, error: 'Day must be 0\u20136' };
  }

  var t = parseHM_(raw.startTime != null ? raw.startTime : (pad2_(raw.startH) + ':' + pad2_(raw.startM)));

  var dur = parseInt(raw.durationMins, 10);
  if (isNaN(dur) || dur <= 0) return { ok: false, error: 'Duration must be a positive number' };

  var type = (raw.type || 'online').toString().trim().toLowerCase();
  if (type !== 'online' && type !== 'inperson') return { ok: false, error: 'Type must be online or inperson' };

  var cap = null;
  if (raw.capacity !== '' && raw.capacity != null) {
    cap = parseInt(raw.capacity, 10);
    if (isNaN(cap) || cap < 0) return { ok: false, error: 'Capacity must be a whole number or blank' };
  }

  var tags  = (Array.isArray(raw.tags)  ? raw.tags  : splitList_(raw.tags )).map(function (x) { return x.toString().trim(); }).filter(Boolean);
  var props = (Array.isArray(raw.props) ? raw.props : splitList_(raw.props)).map(function (x) { return x.toString().trim(); }).filter(Boolean);

  var active = !(raw.active === false || String(raw.active).toLowerCase() === 'false');

  var value = {
    id: (raw.id || '').toString().trim(),
    label: label, day: day, startH: t.h, startM: t.m,
    durationMins: dur, type: type,
    location: (raw.location || '').toString().trim(),
    capacity: cap, tags: tags, props: props, active: active,
    oneOffDate: oneOffDate
  };
  if (opts.requireId && !value.id) return { ok: false, error: 'Missing class id' };
  return { ok: true, value: value };
}

function classToScheduleRow_(c) {
  return [
    c.id, c.label, c.day,
    pad2_(c.startH) + ':' + pad2_(c.startM),
    c.durationMins, c.type, c.location,
    (c.capacity == null ? '' : c.capacity),
    c.tags.join(', '), c.props.join(', '),
    c.active ? 'TRUE' : 'FALSE',
    (c.oneOffDate || '')
  ];
}

function writeScheduleRow_(sheet, rowIdx, c) {
  var range = sheet.getRange(rowIdx, 1, 1, 12);
  range.setNumberFormat('@');
  range.setValues([classToScheduleRow_(c)]);
}

function findScheduleRowById_(sheet, id) {
  var last = sheet.getLastRow();
  if (last <= 1) return -1;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if ((ids[i][0] || '').toString().trim() === id) return i + 2;
  }
  return -1;
}

function findExceptionRow_(sheet, classId, dateIso) {
  var last = sheet.getLastRow();
  if (last <= 1) return -1;
  var data = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === classId &&
        normalizeIso_(data[i][1]) === dateIso) return i + 2;
  }
  return -1;
}

function slugify_(s) {
  return (s || '').toString().toLowerCase()
    .replace(/[\u2014\u2013]/g, '-')
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/-{2,}/g, '-')
    .slice(0, 40)
    .replace(/-+$/g, '');
}

function uniqueClassId_(sheet, base) {
  base = base || 'class';
  var existing = {};
  var last = sheet.getLastRow();
  if (last > 1) {
    var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) existing[(ids[i][0] || '').toString().trim()] = true;
  }
  if (!existing[base]) return base;
  var n = 2;
  while (existing[base + '-' + n]) n++;
  return base + '-' + n;
}

function relabelSignups_(oldLabel, newLabel) {
  var count = 0;
  count += relabelInSheet_(getOrCreateSpreadsheet(), 'Sign-Ups', oldLabel, newLabel);
  try {
    var wfiles = DriveApp.getFilesByName(TEST_WAITLIST_SS_NAME);
    if (wfiles.hasNext()) {
      count += relabelInSheet_(SpreadsheetApp.open(wfiles.next()), 'Waitlist', oldLabel, newLabel);
    }
  } catch (e) {
    Logger.log('relabelSignups_ waitlist error: ' + e);
  }
  return count;
}

function relabelInSheet_(ss, sheetName, oldLabel, newLabel) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  var last = sheet.getLastRow();
  if (last <= 1) return 0;
  var rng = sheet.getRange(2, 5, last - 1, 1); // Class label is column 5
  var vals = rng.getValues();
  var changed = 0;
  for (var i = 0; i < vals.length; i++) {
    if ((vals[i][0] || '').toString().trim() === oldLabel) { vals[i][0] = newLabel; changed++; }
  }
  if (changed) rng.setValues(vals);
  return changed;
}

// ============================================================================
// ========== ADMIN CONSOLE — Phase 3: cancellation + notifications ==========
// ============================================================================
// Mirror of production. Token-gated (see doPost ADMIN_ACTIONS gate). The ONLY
// divergence from prod is the two env helpers below: openWaitlistSpreadsheetIfExists_()
// (test waitlist filename) and deliverEmail_() (test logs to "Test Email Log"
// instead of sending). Every other function here is identical to prod.

// Open the Waitlist spreadsheet if it exists, else null (never creates one).
// TEST: uses TEST_WAITLIST_SS_NAME (prod uses 'Yoga Waitlist').
function openWaitlistSpreadsheetIfExists_() {
  var files = DriveApp.getFilesByName(TEST_WAITLIST_SS_NAME);
  return files.hasNext() ? SpreadsheetApp.open(files.next()) : null;
}

// Single mail primitive for Phase 3 emails. TEST: logs to "Test Email Log"
// instead of sending, so the email templates stay byte-identical to prod.
function deliverEmail_(to, subject, htmlBody, type) {
  getOrCreateEmailLogSheet().appendRow([
    new Date().toISOString(), to, subject, htmlBody, '', type || 'Phase 3'
  ]);
}

// --- date / time helpers -------------------------------------------------

// Build the stored display date ("EEEE, MMMM d, yyyy") from an ISO date. Used
// as a fallback; the admin UI normally passes the exact display string it shows.
function displayDateFromIso_(iso) {
  var m = String(iso || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return '';
  var d = new Date(+m[1], +m[2] - 1, +m[3], 12, 0, 0); // noon avoids DST edges
  return formatClassDate(d);
}

function fmtTime12_(hhmm) {
  var t = parseHM_(hhmm);
  var ampm = t.h >= 12 ? 'PM' : 'AM';
  var hh = t.h % 12; if (hh === 0) hh = 12;
  return hh + ':' + pad2_(t.m) + ' ' + ampm;
}

// Parse a stored class-date cell ("Sunday, April 13, 2026", "April 13", or a
// Date object) into a Date. Mirrors the robust parsing in archiveSheet_.
function parseStoredClassDate_(classDateStr, pstNow) {
  var classDate;
  if (classDateStr instanceof Date) {
    classDate = classDateStr;
  } else {
    var s = String(classDateStr).trim().replace(/^[A-Za-z]+,\s*/, '');
    if (!/\d{4}/.test(s)) {
      var inferYear = pstNow.getFullYear();
      var attempt = new Date(s + ', ' + inferYear);
      if (!isNaN(attempt.getTime()) && (pstNow - attempt) > 7 * 24 * 3600 * 1000) {
        attempt = new Date(s + ', ' + (inferYear + 1));
      }
      classDate = attempt;
    } else {
      classDate = new Date(s);
    }
  }
  if (isNaN(classDate.getTime())) return null;
  return classDate;
}

// --- local-timezone class time rendering ---------------------------------
// The class is always Pacific wall-clock. These build the actual instant for a
// Pacific date + time and render it in both Pacific and the recipient's IANA
// zone, e.g. "6:00 PM PDT (9:00 PM EDT)". Used by every class email.

// ISO 'yyyy-MM-dd' from a stored/display class date ("Sunday, April 13, 2026")
// or '' if unparseable. Reads the calendar fields directly (no tz round-trip).
function isoFromDisplayDate_(dateDisplay) {
  if (!dateDisplay) return '';
  var s = String(dateDisplay).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  var pacNow = new Date(new Date().toLocaleString('en-US', { timeZone: MEET_TZ }));
  var d = parseStoredClassDate_(s, pacNow);
  if (!d || isNaN(d.getTime())) return '';
  return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
}

// The UTC instant for a Pacific wall-clock date ('yyyy-MM-dd') + time ('HH:MM').
function pacificInstant_(iso, hhmm) {
  var d = String(iso).split('-');
  if (d.length < 3) return null;
  var t = parseHM_(hhmm);
  // Treat the wall-clock components as if they were UTC, then shift by Pacific's
  // offset at that moment so the instant lands on the intended Pacific time.
  var guess = new Date(Date.UTC(+d[0], +d[1] - 1, +d[2], t.h, t.m, 0));
  var z = Utilities.formatDate(guess, MEET_TZ, 'Z'); // e.g. "-0700"
  var sign = z.charAt(0) === '-' ? -1 : 1;
  var offMin = sign * (parseInt(z.substr(1, 2), 10) * 60 + parseInt(z.substr(3, 2), 10));
  return new Date(guess.getTime() - offMin * 60000);
}

// "6:00 PM PDT (9:00 PM EDT)" for the recipient. Pacific-only when the recipient
// zone is missing/unknown or resolves to the same wall time. '' if inputs bad.
function localTimeLine_(hhmm, dateDisplay, recipientTz) {
  try {
    if (!hhmm) return '';
    var iso = isoFromDisplayDate_(dateDisplay);
    if (!iso) return '';
    var inst = pacificInstant_(iso, hhmm);
    if (!inst || isNaN(inst.getTime())) return '';
    var pac = Utilities.formatDate(inst, MEET_TZ, 'h:mm a') + ' ' + Utilities.formatDate(inst, MEET_TZ, 'z');
    var tz = (recipientTz || '').toString().trim();
    if (!tz) return pac;
    var loc;
    try {
      loc = Utilities.formatDate(inst, tz, 'h:mm a') + ' ' + Utilities.formatDate(inst, tz, 'z');
    } catch (e) {
      return pac; // unrecognized IANA zone
    }
    return (loc === pac) ? pac : (pac + ' (' + loc + ')');
  } catch (err) {
    return '';
  }
}

// True if a signup timestamp falls before the given Pacific class day (so the
// student signed up ahead of time, not the morning of). Unknown -> true.
function signedUpBeforeDay_(ts, classIso) {
  if (!ts) return true;
  var d = (ts instanceof Date) ? ts : new Date(String(ts));
  if (isNaN(d.getTime())) return true;
  return Utilities.formatDate(d, MEET_TZ, 'yyyy-MM-dd') < classIso;
}

// --- registrant / waitlist readers (full records, online + in person) ----

function getRegisteredStudentsFull_(label, dateDisplay) {
  var out = [];
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return out;
    // Read through the Timezone column (18) when present; older sheets that
    // haven't grown that column yet read fewer cols and fall back to ''.
    var nCols = Math.min(18, sheet.getLastColumn() || 11);
    var data = sheet.getRange(2, 1, lastRow - 1, nCols).getValues();
    var targetMD = extractMonthDay_(dateDisplay);
    for (var i = 0; i < data.length; i++) {
      var rowClass = (data[i][4] || '').toString().trim();
      var rowDate  = (data[i][5] || '').toString().trim();
      if (rowClass === label && (rowDate === dateDisplay || extractMonthDay_(rowDate) === targetMD)) {
        out.push({
          timestamp:  data[i][0],
          firstName:  (data[i][1] || '').toString().trim(),
          lastName:   (data[i][2] || '').toString().trim(),
          email:      (data[i][3] || '').toString().trim(),
          classType:  (data[i][6] || '').toString().trim(),
          guestFirst: (data[i][8] || '').toString().trim(),
          guestLast:  (data[i][9] || '').toString().trim(),
          timezone:   (data[i][17] || '').toString().trim()
        });
      }
    }
  } catch (err) { Logger.log('getRegisteredStudentsFull_ error: ' + err); }
  return out;
}

function getWaitlistStudentsFull_(label, dateDisplay) {
  var out = [];
  try {
    var ss = openWaitlistSpreadsheetIfExists_();
    if (!ss) return out;
    var sheet = ss.getSheetByName('Waitlist');
    if (!sheet || sheet.getLastRow() <= 1) return out;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var targetMD = extractMonthDay_(dateDisplay);
    for (var i = 0; i < data.length; i++) {
      var rowClass = (data[i][4] || '').toString().trim();
      var rowDate  = (data[i][5] || '').toString().trim();
      if (rowClass === label && (rowDate === dateDisplay || extractMonthDay_(rowDate) === targetMD)) {
        out.push({
          firstName:  (data[i][1] || '').toString().trim(),
          lastName:   (data[i][2] || '').toString().trim(),
          email:      (data[i][3] || '').toString().trim(),
          classType:  (data[i][6] || '').toString().trim(),
          guestFirst: (data[i][7] || '').toString().trim(),
          guestLast:  (data[i][8] || '').toString().trim()
        });
      }
    }
  } catch (err) { Logger.log('getWaitlistStudentsFull_ error: ' + err); }
  return out;
}

// --- targeted archive (move specific class+date rows out of the live sheets) ---

function archiveRowsForClassDate_(label, dateDisplay) {
  var signups = archiveMatchingRows_(getOrCreateSpreadsheet(), 'Sign-Ups', 'Archive', label, dateDisplay);
  var waitlist = 0;
  try {
    var wss = openWaitlistSpreadsheetIfExists_();
    if (wss) waitlist = archiveMatchingRows_(wss, 'Waitlist', 'Waitlist Archive', label, dateDisplay);
  } catch (e) { Logger.log('archiveRowsForClassDate_ waitlist error: ' + e); }
  return { signups: signups, waitlist: waitlist };
}

// Move every row matching (label, dateDisplay) from sheetName to archiveName in
// the same spreadsheet. Class label is col 5 (idx 4), Class Date col 6 (idx 5)
// in both Sign-Ups and Waitlist. Batch write + bottom-up delete, last-row safe.
function archiveMatchingRows_(ss, sheetName, archiveName, label, dateDisplay) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues(); // incl. header
  var targetMD = extractMonthDay_(dateDisplay);

  var archive = ss.getSheetByName(archiveName);
  if (!archive) {
    archive = ss.insertSheet(archiveName);
    archive.appendRow(data[0]);
    archive.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
    archive.setFrozenRows(1);
  }

  // Collect matches top-down indices but iterate bottom-up so deletes are safe.
  var matches = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var rowClass = (data[i][4] || '').toString().trim();
    var rowDate  = (data[i][5] || '').toString().trim();
    if (rowClass === label && (rowDate === dateDisplay || extractMonthDay_(rowDate) === targetMD)) {
      matches.push({ index: i, row: data[i] });
    }
  }
  if (!matches.length) return 0;

  // Batch write to archive in natural (top-down) order.
  var rows = [];
  for (var a = matches.length - 1; a >= 0; a--) rows.push(matches[a].row);
  var archLast = archive.getLastRow();
  archive.getRange(archLast + 1, 1, rows.length, rows[0].length).setValues(rows);

  // Batch delete from source bottom-up (protect the last non-frozen row).
  var totalDataRows = sheet.getLastRow() - 1;
  for (var d = 0; d < matches.length; d++) {
    if (totalDataRows <= 1) {
      sheet.getRange(matches[d].index + 1, 1, 1, sheet.getLastColumn()).clearContent();
    } else {
      sheet.deleteRow(matches[d].index + 1);
    }
    totalDataRows--;
  }
  return matches.length;
}

// --- upcoming-date discovery (for whole-series cancel + delete guard) -----

// Distinct stored display-dates for a label whose sign-up cutoff is still in the
// future, across Sign-Ups and Waitlist. These are the occurrences a series
// cancel must notify/clean up, and the rows the delete guard refuses on.
function distinctUpcomingDatesForLabel_(label, startTimes) {
  var set = {};
  collectUpcomingDatesFromSheet_(getOrCreateSpreadsheet(), 'Sign-Ups', label, startTimes, set);
  try {
    var wss = openWaitlistSpreadsheetIfExists_();
    if (wss) collectUpcomingDatesFromSheet_(wss, 'Waitlist', label, startTimes, set);
  } catch (e) { Logger.log('distinctUpcomingDatesForLabel_ waitlist error: ' + e); }
  return Object.keys(set);
}

function collectUpcomingDatesFromSheet_(ss, sheetName, label, startTimes, set) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var now = new Date();
  var pstNow = new Date(now.toLocaleString('en-US', { timeZone: MEET_TZ }));
  for (var i = 0; i < data.length; i++) {
    var rowClass = (data[i][4] || '').toString().trim();
    if (rowClass !== label) continue;
    var rowDateRaw = data[i][5];
    if (!rowDateRaw) continue;
    var parsed = parseStoredClassDate_(rowDateRaw, pstNow);
    if (!parsed) continue;
    var st = startTimes[label] || { startH: 18, startM: 0 };
    var cutoff = new Date(parsed);
    cutoff.setHours(st.startH, st.startM + 15, 0, 0);
    if (pstNow <= cutoff) set[rowDateRaw.toString().trim()] = true; // still upcoming
  }
}

// --- email templates (match the confirmation-email look) -----------------

function sendClassCancelledEmail_(student, label, dateDisplay) {
  if (!student || !student.email) return false;
  try {
    var sc = getScheduleClassByLabel_(label);
    var timeStr = sc ? localTimeLine_(sc.startTime, dateDisplay, student.timezone) : '';
    var subject = 'Class cancelled \u2014 ' + ((label || '').split(' \u2014 ')[0] || label) + ' on ' + extractMonthDay_(dateDisplay);
    var body = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
      '<div style="background:#f5f0e8;padding:24px;text-align:center;border-radius:8px 8px 0 0;">' +
        '<h1 style="margin:0;font-family:Georgia,serif;font-size:24px;">' +
          '<a href="' + SITE_URL + '" style="color:#5B7553;text-decoration:none;">yogawithjessica.com</a>' +
        '</h1>' +
        '<p style="margin:6px 0 0;color:#888;font-size:13px;">Class Cancelled</p>' +
      '</div>' +
      '<div style="padding:24px;background:#fff;border:1px solid #e8e4dc;border-top:none;">' +
        '<p style="font-size:15px;">Hi ' + escHtml(student.firstName || 'there') + ',</p>' +
        '<p style="font-size:15px;line-height:1.6;">We&rsquo;re sorry to let you know that the following class has been <strong>cancelled</strong>:</p>' +
        '<div style="background:#fdecea;padding:16px;border-radius:6px;margin:16px 0;font-size:14px;border-left:4px solid #c62828;">' +
          '<strong>' + escHtml(label) + '</strong><br>' + escHtml(dateDisplay) +
          (timeStr ? '<br>' + escHtml(timeStr) : '') +
        '</div>' +
        (student.guestFirst ? '<p style="font-size:14px;color:#555;">This also releases the spot held for your guest, ' + escHtml(student.guestFirst) + ' ' + escHtml(student.guestLast) + '.</p>' : '') +
        '<p style="font-size:14px;line-height:1.6;color:#555;">No action is needed &mdash; your registration has been removed. We hope to see you in a future class; you can view the latest schedule and sign up again anytime:</p>' +
        '<div style="text-align:center;margin:20px 0;">' +
          '<a href="' + SITE_URL + '/schedule.html" style="display:inline-block;background:#5B7553;color:#fff;padding:12px 32px;border-radius:6px;text-decoration:none;font-size:15px;font-weight:600;">View Schedule</a>' +
        '</div>' +
        '<p style="font-size:14px;color:#555;">With gratitude,<br>Jessica</p>' +
      '</div>' +
      '<div style="padding:16px;text-align:center;background:#f5f0e8;border-radius:0 0 8px 8px;">' +
        '<p style="margin:0;font-size:12px;color:#999;">Yoga with Jessica &mdash; Playa Del Rey, CA</p>' +
        '<p style="margin:6px 0 0;"><a href="' + SITE_URL + '" style="color:#5B7553;font-size:15px;font-weight:600;text-decoration:none;">yogawithjessica.com</a></p>' +
      '</div>' +
    '</div>';
    deliverEmail_(student.email, subject, body, 'Class Cancelled');
    return true;
  } catch (err) { Logger.log('sendClassCancelledEmail_ error for ' + student.email + ': ' + err); return false; }
}

function sendClassMovedEmail_(student, label, oldDateDisplay, newDateDisplay, newStartTime) {
  if (!student || !student.email) return false;
  try {
    var sc = getScheduleClassByLabel_(label);
    var hhmm = newStartTime || (sc ? sc.startTime : '');
    var tzLine = localTimeLine_(hhmm, newDateDisplay, student.timezone);
    var timeLine = tzLine ? (' at ' + escHtml(tzLine)) : '';
    var subject = 'Class rescheduled \u2014 ' + ((label || '').split(' \u2014 ')[0] || label);
    var body = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
      '<div style="background:#f5f0e8;padding:24px;text-align:center;border-radius:8px 8px 0 0;">' +
        '<h1 style="margin:0;font-family:Georgia,serif;font-size:24px;">' +
          '<a href="' + SITE_URL + '" style="color:#5B7553;text-decoration:none;">yogawithjessica.com</a>' +
        '</h1>' +
        '<p style="margin:6px 0 0;color:#888;font-size:13px;">Class Rescheduled</p>' +
      '</div>' +
      '<div style="padding:24px;background:#fff;border:1px solid #e8e4dc;border-top:none;">' +
        '<p style="font-size:15px;">Hi ' + escHtml(student.firstName || 'there') + ',</p>' +
        '<p style="font-size:15px;line-height:1.6;">Your upcoming class has been <strong>rescheduled</strong>. Your spot is reserved for the new date &mdash; no action is needed.</p>' +
        '<div style="background:#f0f5ee;padding:16px;border-radius:6px;margin:16px 0;font-size:14px;border-left:4px solid #5B7553;">' +
          '<strong>' + escHtml(label) + '</strong><br>' +
          '<span style="color:#999;text-decoration:line-through;">' + escHtml(oldDateDisplay) + '</span><br>' +
          '<strong style="color:#5B7553;">Now: ' + escHtml(newDateDisplay) + timeLine + '</strong>' +
        '</div>' +
        (student.guestFirst ? '<p style="font-size:14px;color:#555;">Your guest&rsquo;s spot (' + escHtml(student.guestFirst) + ' ' + escHtml(student.guestLast) + ') has moved with you.</p>' : '') +
        '<div style="text-align:center;margin:20px 0;">' +
          '<a href="' + SITE_URL + '/schedule.html" style="display:inline-block;background:#5B7553;color:#fff;padding:12px 32px;border-radius:6px;text-decoration:none;font-size:15px;font-weight:600;">View Schedule</a>' +
        '</div>' +
        '<p style="font-size:14px;color:#555;">See you then!<br>Jessica</p>' +
      '</div>' +
      '<div style="padding:16px;text-align:center;background:#f5f0e8;border-radius:0 0 8px 8px;">' +
        '<p style="margin:0;font-size:12px;color:#999;">Yoga with Jessica &mdash; Playa Del Rey, CA</p>' +
        '<p style="margin:6px 0 0;"><a href="' + SITE_URL + '" style="color:#5B7553;font-size:15px;font-weight:600;text-decoration:none;">yogawithjessica.com</a></p>' +
      '</div>' +
    '</div>';
    deliverEmail_(student.email, subject, body, 'Class Rescheduled');
    return true;
  } catch (err) { Logger.log('sendClassMovedEmail_ error for ' + student.email + ': ' + err); return false; }
}

// ========== MORNING-OF REMINDER EMAILS ==========
// Daily time-driven trigger (~7–8 AM PT; see setupTriggers in prod). For each
// class happening today (Pacific), email registered students who signed up
// BEFORE today — a same-day signup already knows. Deduped per class+date via a
// Script Property so a re-fire inside the trigger's hour window can't double-send.
// (TEST: no trigger is installed; invoke directly. Emails are logged, not sent.)
// forIso (optional 'yyyy-MM-dd') lets tests target a specific class day; the
// daily trigger calls this with no argument, which uses the real Pacific today.
function sendClassReminders(forIso) {
  var pstNow, todayIso;
  if (forIso) {
    var p = String(forIso).split('-');
    pstNow = new Date(+p[0], +p[1] - 1, +p[2], 12, 0, 0); // noon avoids DST edges
    todayIso = forIso;
  } else {
    var now = new Date();
    pstNow = new Date(now.toLocaleString('en-US', { timeZone: MEET_TZ }));
    todayIso = Utilities.formatDate(now, MEET_TZ, 'yyyy-MM-dd');
  }
  var props = PropertiesService.getScriptProperties();
  var occurrences = getOccurrencesOnPacificDate_(pstNow);

  Logger.log('Reminder check for ' + todayIso + ' — ' + occurrences.length + ' occurrence(s)');

  for (var c = 0; c < occurrences.length; c++) {
    var cls = occurrences[c];
    var classDate = cls.classDate;
    var sentKey = 'reminder_sent_' + cls.id + '_' + classDate;
    if (props.getProperty(sentKey)) {
      Logger.log('Reminder already sent for ' + cls.label + ' on ' + classDate);
      continue;
    }

    var students = getRegisteredStudentsFull_(cls.label, classDate);
    var hhmm = pad2_(cls.startH) + ':' + pad2_(cls.startM);
    var sent = 0, seen = {};
    for (var i = 0; i < students.length; i++) {
      var s = students[i];
      var em = (s.email || '').toLowerCase();
      if (!em || seen[em]) continue;
      seen[em] = true;
      if (!signedUpBeforeDay_(s.timestamp, todayIso)) continue; // skip same-day signups
      if (sendClassReminderEmail_(s, cls, hhmm)) sent++;
    }
    // Mark done even if nobody was eligible, so we don't re-scan this occurrence.
    props.setProperty(sentKey, new Date().toISOString());
    Logger.log('Reminder for ' + cls.label + ' on ' + classDate + ': sent ' + sent);
  }
}

function sendClassReminderEmail_(student, cls, hhmm) {
  if (!student || !student.email) return false;
  try {
    var timeStr = localTimeLine_(hhmm, cls.classDate, student.timezone);
    var subject = 'Reminder: your Yoga with Jessica class today';
    var isOnline = (cls.type === 'online');
    var detailHtml = isOnline
      ? '<p style="font-size:14px;line-height:1.6;color:#555;">This is an <strong>online</strong> class. Your Zoom link will arrive about <strong>30 minutes before</strong> class starts.</p>'
      : ('<p style="font-size:14px;line-height:1.6;color:#555;">This is an <strong>in-person</strong> class' +
         (cls.location ? ' at <strong>' + escHtml(cls.location) + '</strong>' : '') + '.</p>');
    var body = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
      '<div style="background:#f5f0e8;padding:24px;text-align:center;border-radius:8px 8px 0 0;">' +
        '<h1 style="margin:0;font-family:Georgia,serif;font-size:24px;">' +
          '<a href="' + SITE_URL + '" style="color:#5B7553;text-decoration:none;">yogawithjessica.com</a>' +
        '</h1>' +
        '<p style="margin:6px 0 0;color:#888;font-size:13px;">Class Reminder</p>' +
      '</div>' +
      '<div style="padding:24px;background:#fff;border:1px solid #e8e4dc;border-top:none;">' +
        '<p style="font-size:15px;">Hi ' + escHtml(student.firstName || 'there') + ',</p>' +
        '<p style="font-size:15px;line-height:1.6;">A friendly reminder that you&rsquo;re registered for a class <strong>today</strong>:</p>' +
        '<div style="background:#f0f5ee;padding:16px;border-radius:6px;margin:16px 0;font-size:14px;border-left:4px solid #5B7553;">' +
          '<strong>' + escHtml(cls.label) + '</strong><br>' + escHtml(cls.classDate) +
          (timeStr ? '<br>' + escHtml(timeStr) : '') +
        '</div>' +
        detailHtml +
        (student.guestFirst ? '<p style="font-size:14px;color:#555;">Your guest ' + escHtml(student.guestFirst) + ' ' + escHtml(student.guestLast) + ' is registered with you.</p>' : '') +
        '<p style="font-size:14px;color:#555;line-height:1.6;">Don&rsquo;t forget to check the <a href="' + SITE_URL + '/props.html" style="color:#5B7553;">Props page</a> for what to bring.</p>' +
        '<p style="font-size:14px;color:#555;">See you soon!<br>Jessica</p>' +
      '</div>' +
      '<div style="padding:16px;text-align:center;background:#f5f0e8;border-radius:0 0 8px 8px;">' +
        '<p style="margin:0;font-size:12px;color:#999;">Yoga with Jessica &mdash; Playa Del Rey, CA</p>' +
        '<p style="margin:6px 0 0;"><a href="' + SITE_URL + '" style="color:#5B7553;font-size:15px;font-weight:600;text-decoration:none;">yogawithjessica.com</a></p>' +
      '</div>' +
    '</div>';
    deliverEmail_(student.email, subject, body, 'Class Reminder');
    return true;
  } catch (err) { Logger.log('sendClassReminderEmail_ error for ' + student.email + ': ' + err); return false; }
}

// Admin summary for a Phase 3 action. `lines` are pre-built HTML fragments.
function sendAdminClassActionNotification_(title, lines) {
  try {
    var ts = new Date().toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }) + ' PST';
    var body = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
      '<div style="background:#fff3e0;padding:16px 24px;border-radius:8px 8px 0 0;border-left:4px solid #e07b39;">' +
        '<h2 style="margin:0;color:#b45309;font-size:18px;">' + escHtml(title) + '</h2>' +
      '</div>' +
      '<div style="padding:20px 24px;background:#fff;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px;">' +
        lines.map(function (l) { return '<p style="margin:6px 0;">' + l + '</p>'; }).join('') +
        '<p style="color:#777;font-size:13px;margin-top:16px;">' + ts + '</p>' +
      '</div>' +
    '</div>';
    deliverEmail_(ADMIN_EMAIL, '\u26A0 Admin: ' + title, body, 'Admin - Schedule Change');
  } catch (err) { Logger.log('sendAdminClassActionNotification_ error: ' + err); }
}

// --- shared send loops (dedupe by email) ---------------------------------

function sendCancelEmails_(label, dateDisplay, list) {
  var sent = 0, seen = {};
  for (var i = 0; i < list.length; i++) {
    var em = (list[i].email || '').toLowerCase();
    if (!em || seen[em]) continue;
    seen[em] = true;
    if (sendClassCancelledEmail_(list[i], label, dateDisplay)) sent++;
  }
  return sent;
}

function sendMoveEmails_(label, oldDisplay, newDisplay, newStartTime, list) {
  var sent = 0, seen = {};
  for (var i = 0; i < list.length; i++) {
    var em = (list[i].email || '').toLowerCase();
    if (!em || seen[em]) continue;
    seen[em] = true;
    if (sendClassMovedEmail_(list[i], label, oldDisplay, newDisplay, newStartTime)) sent++;
  }
  return sent;
}

// Cancel one dated occurrence: email registered students (waitlist only if
// asked) and archive ALL their rows (waitlist always archived, even when not
// emailed, so the waitlist processor can't later invite them to a dead date).
// Does NOT write the Exceptions row — callers decide.
function cancelOneDate_(label, dateDisplay, notifyWaitlist) {
  var students = getRegisteredStudentsFull_(label, dateDisplay);
  var waiters  = getWaitlistStudentsFull_(label, dateDisplay);
  var emailed         = sendCancelEmails_(label, dateDisplay, students);
  var waitlistEmailed = notifyWaitlist ? sendCancelEmails_(label, dateDisplay, waiters) : 0;
  var arch = archiveRowsForClassDate_(label, dateDisplay);
  return {
    signups: students.length, waitlist: waiters.length,
    emailed: emailed, waitlistEmailed: waitlistEmailed,
    archived: arch.signups, waitlistArchived: arch.waitlist
  };
}

// Re-stamp the stored Class Date on Sign-Ups + Waitlist rows old -> new (the
// class date is part of the join key for capacity / Zoom / archive). Returns
// the number of rows updated. Mirrors relabelSignups_ but for the date column.
function restampSignupsDate_(label, oldDisplay, newDisplay) {
  var count = 0;
  count += restampDateInSheet_(getOrCreateSpreadsheet(), 'Sign-Ups', label, oldDisplay, newDisplay);
  try {
    var wss = openWaitlistSpreadsheetIfExists_();
    if (wss) count += restampDateInSheet_(wss, 'Waitlist', label, oldDisplay, newDisplay);
  } catch (e) { Logger.log('restampSignupsDate_ waitlist error: ' + e); }
  return count;
}

function restampDateInSheet_(ss, sheetName, label, oldDisplay, newDisplay) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  var last = sheet.getLastRow();
  if (last <= 1) return 0;
  var rng = sheet.getRange(2, 5, last - 1, 2); // cols 5-6: Class, Class Date
  var vals = rng.getValues();
  var targetMD = extractMonthDay_(oldDisplay);
  var changed = 0;
  for (var i = 0; i < vals.length; i++) {
    var rc = (vals[i][0] || '').toString().trim();
    var rd = (vals[i][1] || '').toString().trim();
    if (rc === label && (rd === oldDisplay || extractMonthDay_(rd) === targetMD)) {
      vals[i][1] = newDisplay; changed++;
    }
  }
  if (changed) { rng.setNumberFormat('@'); rng.setValues(vals); }
  return changed;
}

// --- the four token-gated Phase 3 actions --------------------------------

function adminCancelOccurrence_(data) {
  try {
    var classId = (data.classId || '').toString().trim();
    var iso = normalizeIso_(data.date);
    var notifyWaitlist = (data.notifyWaitlist === true || String(data.notifyWaitlist).toLowerCase() === 'true');
    if (!classId) return { status: 'error', message: 'Missing classId' };
    if (!iso) return { status: 'error', message: 'Missing or invalid date' };

    var ss = getOrCreateSpreadsheet();
    var schedSheet = getOrCreateScheduleSheet_(ss);
    var rowIdx = findScheduleRowById_(schedSheet, classId);
    if (rowIdx < 0) return { status: 'error', message: 'Class not found: ' + classId };
    var label = (schedSheet.getRange(rowIdx, 2).getValue() || '').toString().trim();
    var dateDisplay = (data.dateDisplay || '').toString().trim() || displayDateFromIso_(iso);

    // 1. Upsert a cancelled exception so the date drops from the schedule first.
    var exSheet = getOrCreateExceptionsSheet_(ss);
    var exRow = findExceptionRow_(exSheet, classId, iso);
    if (exRow < 0) exRow = exSheet.getLastRow() + 1;
    var exRange = exSheet.getRange(exRow, 1, 1, 6);
    exRange.setNumberFormat('@');
    exRange.setValues([[classId, iso, 'cancelled', '', '', (data.note || '').toString()]]);

    // 2. Notify + archive the affected registrants.
    var res = cancelOneDate_(label, dateDisplay, notifyWaitlist);

    sendAdminClassActionNotification_('Class cancelled \u2014 ' + label, [
      '<strong>Date:</strong> ' + escHtml(dateDisplay),
      '<strong>Registered students emailed:</strong> ' + res.emailed + ' of ' + res.signups,
      '<strong>Waitlist emailed:</strong> ' + res.waitlistEmailed + ' of ' + res.waitlist,
      '<strong>Archived:</strong> ' + res.archived + ' sign-up row(s), ' + res.waitlistArchived + ' waitlist row(s)'
    ]);
    bustScheduleCache();
    return {
      status: 'ok', classId: classId, date: iso,
      signups: res.signups, waitlist: res.waitlist,
      emailed: res.emailed, waitlistEmailed: res.waitlistEmailed,
      archived: res.archived, waitlistArchived: res.waitlistArchived
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function adminCancelSeries_(data) {
  try {
    var id = (data.id || '').toString().trim();
    var notifyWaitlist = (data.notifyWaitlist === true || String(data.notifyWaitlist).toLowerCase() === 'true');
    if (!id) return { status: 'error', message: 'Missing class id' };

    var ss = getOrCreateSpreadsheet();
    var schedSheet = getOrCreateScheduleSheet_(ss);
    var rowIdx = findScheduleRowById_(schedSheet, id);
    if (rowIdx < 0) return { status: 'error', message: 'Class not found: ' + id };
    var label = (schedSheet.getRange(rowIdx, 2).getValue() || '').toString().trim();

    // 1. Deactivate the series (hides all future dates, incl. any moved-in ones).
    var cell = schedSheet.getRange(rowIdx, 11);
    cell.setNumberFormat('@');
    cell.setValue('FALSE');

    // 2. Notify + archive every upcoming date that has registrants/waitlisters.
    var dates = distinctUpcomingDatesForLabel_(label, buildStartTimeMap_());
    var totals = { signups: 0, waitlist: 0, emailed: 0, waitlistEmailed: 0, archived: 0, waitlistArchived: 0 };
    for (var i = 0; i < dates.length; i++) {
      var r = cancelOneDate_(label, dates[i], notifyWaitlist);
      totals.signups += r.signups; totals.waitlist += r.waitlist;
      totals.emailed += r.emailed; totals.waitlistEmailed += r.waitlistEmailed;
      totals.archived += r.archived; totals.waitlistArchived += r.waitlistArchived;
    }

    sendAdminClassActionNotification_('Series cancelled \u2014 ' + label, [
      'The series was deactivated and removed from the public schedule.',
      '<strong>Upcoming dates affected:</strong> ' + dates.length,
      '<strong>Registered students emailed:</strong> ' + totals.emailed + ' of ' + totals.signups,
      '<strong>Waitlist emailed:</strong> ' + totals.waitlistEmailed + ' of ' + totals.waitlist,
      '<strong>Archived:</strong> ' + totals.archived + ' sign-up row(s), ' + totals.waitlistArchived + ' waitlist row(s)'
    ]);
    bustScheduleCache();
    return {
      status: 'ok', id: id, datesAffected: dates.length,
      signups: totals.signups, waitlist: totals.waitlist,
      emailed: totals.emailed, waitlistEmailed: totals.waitlistEmailed,
      archived: totals.archived, waitlistArchived: totals.waitlistArchived
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function adminMoveOccurrence_(data) {
  try {
    var classId  = (data.classId || '').toString().trim();
    var iso      = normalizeIso_(data.date);
    var newDate  = normalizeIso_(data.newDate);
    var newStart = normalizeTimeStr_(data.newStartTime);
    var notify         = (data.notify === true || String(data.notify).toLowerCase() === 'true');
    var notifyWaitlist = (data.notifyWaitlist === true || String(data.notifyWaitlist).toLowerCase() === 'true');
    if (!classId) return { status: 'error', message: 'Missing classId' };
    if (!iso) return { status: 'error', message: 'Missing or invalid date' };
    if (!newDate) return { status: 'error', message: 'A move needs a new date' };

    var ss = getOrCreateSpreadsheet();
    var schedSheet = getOrCreateScheduleSheet_(ss);
    var rowIdx = findScheduleRowById_(schedSheet, classId);
    if (rowIdx < 0) return { status: 'error', message: 'Class not found: ' + classId };
    var label = (schedSheet.getRange(rowIdx, 2).getValue() || '').toString().trim();

    // 1. Upsert the moved exception.
    var exSheet = getOrCreateExceptionsSheet_(ss);
    var exRow = findExceptionRow_(exSheet, classId, iso);
    if (exRow < 0) exRow = exSheet.getLastRow() + 1;
    var exRange = exSheet.getRange(exRow, 1, 1, 6);
    exRange.setNumberFormat('@');
    exRange.setValues([[classId, iso, 'moved', newDate, newStart, (data.note || '').toString()]]);

    var oldDisplay = (data.dateDisplay || '').toString().trim() || displayDateFromIso_(iso);
    var newDisplay = displayDateFromIso_(newDate);

    // 2. Re-stamp registrant + waitlist rows to the new date (join key upkeep).
    var restamped = restampSignupsDate_(label, oldDisplay, newDisplay);

    // 3. Optionally email students the new date/time (rows are now on newDisplay).
    var emailed = 0, waitlistEmailed = 0;
    if (notify) {
      emailed = sendMoveEmails_(label, oldDisplay, newDisplay, newStart, getRegisteredStudentsFull_(label, newDisplay));
      if (notifyWaitlist) {
        waitlistEmailed = sendMoveEmails_(label, oldDisplay, newDisplay, newStart, getWaitlistStudentsFull_(label, newDisplay));
      }
    }

    sendAdminClassActionNotification_('Class moved \u2014 ' + label, [
      '<strong>From:</strong> ' + escHtml(oldDisplay),
      '<strong>To:</strong> ' + escHtml(newDisplay) + (newStart ? ' at ' + escHtml(fmtTime12_(newStart)) : ''),
      '<strong>Registration rows re-dated:</strong> ' + restamped,
      '<strong>Registered students emailed:</strong> ' + emailed,
      '<strong>Waitlist emailed:</strong> ' + waitlistEmailed
    ]);
    bustScheduleCache();
    return {
      status: 'ok', classId: classId, date: iso, newDate: newDate, newStartTime: newStart,
      restamped: restamped, emailed: emailed, waitlistEmailed: waitlistEmailed
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function adminDeleteClass_(data) {
  try {
    var id = (data.id || '').toString().trim();
    if (!id) return { status: 'error', message: 'Missing class id' };

    var ss = getOrCreateSpreadsheet();
    var schedSheet = getOrCreateScheduleSheet_(ss);
    var rowIdx = findScheduleRowById_(schedSheet, id);
    if (rowIdx < 0) return { status: 'error', message: 'Class not found: ' + id };
    var label = (schedSheet.getRange(rowIdx, 2).getValue() || '').toString().trim();

    var upcoming = distinctUpcomingDatesForLabel_(label, buildStartTimeMap_());
    if (upcoming.length) {
      return {
        status: 'error', code: 'has_upcoming', upcoming: upcoming.length,
        message: 'This series still has upcoming sign-ups on ' + upcoming.length + ' date(s). Cancel the series first (to notify and archive them), then delete.'
      };
    }

    schedSheet.deleteRow(rowIdx);
    var exDeleted = deleteExceptionsForClass_(ss, id);
    bustScheduleCache();
    return { status: 'ok', id: id, label: label, exceptionsDeleted: exDeleted };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function deleteExceptionsForClass_(ss, classId) {
  var sheet = getOrCreateExceptionsSheet_(ss);
  var last = sheet.getLastRow();
  if (last <= 1) return 0;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  var rows = [];
  for (var i = 0; i < ids.length; i++) {
    if ((ids[i][0] || '').toString().trim() === classId) rows.push(i + 2);
  }
  rows.sort(function (a, b) { return b - a; });
  var total = last - 1;
  for (var d = 0; d < rows.length; d++) {
    if (total <= 1) sheet.getRange(rows[d], 1, 1, 6).clearContent();
    else sheet.deleteRow(rows[d]);
    total--;
  }
  return rows.length;
}

// ========== WRITE SIGN-UPS ==========

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // ----- Admin console (all actions require a verified Google ID token) ---
    // Read = getScheduleAdmin. Writes = updateClass / createClass / setActive /
    // setException (single-date Move). Every write busts the schedule cache.
    var ADMIN_ACTIONS = { getScheduleAdmin: 1, updateClass: 1, createClass: 1, setActive: 1, setException: 1 };
    if (data.action && ADMIN_ACTIONS[data.action]) {
      var adminEmail = verifyAdminToken_(data.idToken);
      if (!adminEmail) {
        return jsonOut_({ status: 'unauthorized', message: 'Sign in with an authorized admin account.' });
      }
      switch (data.action) {
        case 'getScheduleAdmin': return jsonOut_(buildAdminSchedulePayload_(adminEmail));
        case 'updateClass':      return jsonOut_(adminUpdateClass_(data));
        case 'createClass':      return jsonOut_(adminCreateClass_(data));
        case 'setActive':        return jsonOut_(adminSetActive_(data));
        case 'setException':     return jsonOut_(adminSetException_(data));
      }
    }

    var rows = data.rows;

    // Test-only: seed action — insert rows directly into any sheet
    if (data.action === 'seed') {
      return handleSeed(data);
    }

    // Waitlist submission
    if (data.action === 'waitlist') {
      return handleWaitlist(rows);
    }

    var ss = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss);

    // Generate a cancellation token for this sign-up batch
    var cancelToken = generateCancelToken();

    // Append each class as its own row (using setValues to prevent date auto-formatting)
    rows.forEach(function(row) {
      var newRow = sheet.getLastRow() + 1;
      var values = [[
        row.timestamp,
        row.firstName,
        row.lastName,
        row.email,
        row.className,
        row.classDate,
        row.classType,
        row.liabilityWaiver,
        row.guestFirstName || '',
        row.guestLastName || '',
        row.guestOf || '',
        cancelToken,
        row.device || '',
        row.browser || '',
        row.city || '',
        row.state || '',
        row.zip || '',
        row.timezone || ''
      ]];
      var range = sheet.getRange(newRow, 1, 1, 18);
      range.setNumberFormat('@'); // Force plain text to prevent date conversion
      range.setValues(values);
    });

    // Ensure headers exist for all columns
    var header12 = sheet.getRange(1, 12).getValue();
    if (!header12) {
      sheet.getRange(1, 12).setValue('Cancel Token').setFontWeight('bold');
    }
    var header13 = sheet.getRange(1, 13).getValue();
    if (!header13) {
      sheet.getRange(1, 13, 1, 5).setValues([['Device', 'Browser', 'City', 'State', 'Zip Code']]).setFontWeight('bold');
    }
    var header18 = sheet.getRange(1, 18).getValue();
    if (!header18) {
      sheet.getRange(1, 18).setValue('Timezone').setFontWeight('bold');
    }

    // Check if any online classes in this sign-up are starting within 30 minutes
    // In test mode, this uses a mock that checks Script Properties for simulated Meet state
    var meetLink = checkAndCreateMeetForLateSignup(rows);

    // Log confirmation email instead of sending (with Meet link if applicable)
    logConfirmationEmail(rows, cancelToken, meetLink);

    // Log admin sign-up notification
    logAdminSignupNotification(rows);

    var result = { status: 'ok', cancelToken: cancelToken };
    if (meetLink) result.meetLink = meetLink;

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== CANCEL TOKEN ==========
function generateCancelToken() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var token = '';
  for (var i = 0; i < 24; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

// ========== EMAIL LOGGING (replaces MailApp.sendEmail) ==========
// Instead of sending real emails, logs them to a "Test Email Log" sheet
// in the Test Yoga Signup spreadsheet.

function getOrCreateEmailLogSheet() {
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName('Test Email Log');
  if (!sheet) {
    sheet = ss.insertSheet('Test Email Log');
    sheet.appendRow([
      'Timestamp',
      'To',
      'Subject',
      'Body (HTML)',
      'Cancel Token',
      'Type'
    ]);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function logConfirmationEmail(rows, cancelToken, meetLink) {
  if (!rows || rows.length === 0) return;

  var firstName = rows[0].firstName;
  var lastName  = rows[0].lastName;
  var email     = rows[0].email;
  var guestFirst = rows[0].guestFirstName || '';
  var guestLast  = rows[0].guestLastName || '';
  var hasGuest   = !!(guestFirst);

  // Build class list. Time is shown in Pacific + the signer's local zone (the
  // browser timezone captured at signup), e.g. "6:00 PM PDT (9:00 PM EDT)".
  var classLines = '';
  var hasInPerson = false;
  var recipientTz = rows[0].timezone || '';
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var icon = r.classType === 'In-Person' ? '&#x1F3E0;' : '&#x1F4BB;';
    var sc = getScheduleClassByLabel_(r.className);
    var timeStr = sc ? localTimeLine_(sc.startTime, r.classDate, recipientTz) : '';
    classLines += '<tr>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;">' + icon + ' ' + escHtml(r.className) + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;">' + escHtml(r.classDate) + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;">' + (timeStr ? escHtml(timeStr) : '&mdash;') + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;">' + escHtml(r.classType) + '</td>' +
      '</tr>';
    if (r.classType === 'In-Person') hasInPerson = true;
  }

  var cancelUrl = SITE_URL + '/cancel.html?token=' + cancelToken;
  var subject = 'Yoga with Jessica \u2014 Sign-Up Confirmation';

  var body = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
    '<div style="background:#f5f0e8;padding:24px;text-align:center;border-radius:8px 8px 0 0;">' +
      '<h1 style="margin:0;font-family:Georgia,serif;font-size:24px;">' +
        '<a href="' + SITE_URL + '" style="color:#5B7553;text-decoration:none;">yogawithjessica.com</a>' +
      '</h1>' +
      '<p style="margin:6px 0 0;color:#888;font-size:13px;">Sign-Up Confirmation</p>' +
    '</div>' +
    '<div style="padding:24px;background:#fff;border:1px solid #e8e4dc;border-top:none;">' +
      '<p style="font-size:15px;">Hi ' + escHtml(firstName) + ',</p>' +
      '<p style="font-size:15px;line-height:1.6;">Thank you for signing up! Here are your confirmed classes:</p>' +
      '<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:14px;">' +
        '<tr style="background:#f5f0e8;">' +
          '<th style="padding:8px 12px;text-align:left;">Class</th>' +
          '<th style="padding:8px 12px;text-align:left;">Date</th>' +
          '<th style="padding:8px 12px;text-align:left;">Time</th>' +
          '<th style="padding:8px 12px;text-align:left;">Type</th>' +
        '</tr>' +
        classLines +
      '</table>' +
      (hasGuest ?
        '<div style="background:#f9f7f2;padding:12px 16px;border-radius:6px;margin:16px 0;font-size:14px;">' +
          '<strong>Guest:</strong> ' + escHtml(guestFirst) + ' ' + escHtml(guestLast) +
          '<br><span style="color:#777;">Your guest will need to sign a liability waiver upon arrival to class.</span>' +
        '</div>'
      : '') +
      (meetLink ?
        '<div style="background:#e8f5e9;padding:16px;border-radius:6px;margin:16px 0;font-size:14px;border-left:4px solid #5B7553;">' +
          '<strong>&#x1F4F9; Your Zoom link is ready</strong><br>' +
          '<p style="margin:8px 0;">Class is starting soon &mdash; join here:</p>' +
          '<div style="text-align:center;margin:12px 0;">' +
            '<a href="' + meetLink + '" style="display:inline-block;background:#5B7553;color:#fff;padding:12px 32px;border-radius:6px;text-decoration:none;font-size:15px;font-weight:600;">Join Zoom</a>' +
          '</div>' +
          '<p style="margin:8px 0 0;color:#555;">Please have your camera on with good lighting. Microphones will be muted to minimize noise.</p>' +
        '</div>'
      :
        '<div style="background:#f0f5ee;padding:12px 16px;border-radius:6px;margin:16px 0;font-size:14px;">' +
          '<strong>For online classes:</strong> A Zoom link will be sent to you 30 minutes before class. ' +
          'Please have your camera on with good lighting. Microphones will be muted to minimize noise.' +
        '</div>'
      ) +
      '<p style="font-size:14px;color:#555;line-height:1.6;">' +
        'Don\'t forget to check the <a href="' + SITE_URL + '/props.html" style="color:#5B7553;">Props page</a> for recommended props to bring.' +
      '</p>' +
      '<hr style="border:none;border-top:1px solid #e8e4dc;margin:24px 0;" />' +
      '<div style="text-align:center;margin:16px 0;">' +
        '<p style="font-size:14px;color:#555;margin-bottom:12px;">Need to cancel? Click below to cancel your registration' +
          (hasGuest ? ' (this will cancel both your spot and your guest\'s spot).' : '.') +
        '</p>' +
        '<a href="' + cancelUrl + '" style="display:inline-block;background:#b44;color:#fff;padding:10px 28px;border-radius:6px;text-decoration:none;font-size:14px;font-weight:600;">Cancel Registration</a>' +
      '</div>' +
    '</div>' +
    '<div style="padding:16px;text-align:center;background:#f5f0e8;border-radius:0 0 8px 8px;">' +
      '<p style="margin:0;font-size:12px;color:#999;">Yoga with Jessica &mdash; Playa Del Rey, CA</p>' +
      '<p style="margin:6px 0 0;"><a href="' + SITE_URL + '" style="color:#5B7553;font-size:15px;font-weight:600;text-decoration:none;">yogawithjessica.com</a></p>' +
    '</div>' +
  '</div>';

  // LOG to sheet instead of sending
  var logSheet = getOrCreateEmailLogSheet();
  logSheet.appendRow([
    new Date().toISOString(),
    email,
    subject,
    body,
    cancelToken,
    'Confirmation'
  ]);
}

function logWaitlistNotificationEmail(email, firstName, className, classDate) {
  var logSheet = getOrCreateEmailLogSheet();
  var subject = 'Yoga with Jessica \u2014 A Spot Opened Up';
  var body = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
    '<div style="background:#f5f0e8;padding:24px;text-align:center;border-radius:8px 8px 0 0;">' +
      '<h1 style="margin:0;font-family:Georgia,serif;color:#5B7553;font-size:24px;">Yoga with Jessica</h1>' +
    '</div>' +
    '<div style="padding:24px;background:#fff;border:1px solid #e8e4dc;border-top:none;">' +
      '<p style="font-size:15px;">Hi ' + escHtml(firstName) + ',</p>' +
      '<p style="font-size:15px;line-height:1.6;">A spot has opened up for <strong>' +
        escHtml(className) + '</strong> on <strong>' + escHtml(classDate) + '</strong>.</p>' +
      '<p style="font-size:15px;line-height:1.6;">Spots are first come, first served, so sign up soon before it fills up again.</p>' +
      '<div style="text-align:center;margin:24px 0;">' +
        '<a href="' + SITE_URL + '/schedule.html" style="display:inline-block;background:#5B7553;color:#fff;padding:12px 32px;border-radius:6px;text-decoration:none;font-size:15px;font-weight:600;">Sign Up Now</a>' +
      '</div>' +
    '</div>' +
    '<div style="padding:16px;text-align:center;font-size:12px;color:#999;background:#f5f0e8;border-radius:0 0 8px 8px;">' +
      '<p style="margin:0;">Yoga with Jessica &mdash; Playa Del Rey, CA</p>' +
    '</div>' +
  '</div>';

  logSheet.appendRow([
    new Date().toISOString(),
    email,
    subject,
    body,
    '',
    'Waitlist Notification'
  ]);
}

function escHtml(str) {
  return (str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ========== ADMIN NOTIFICATION LOGGING ==========
// Test version: logs to Test Email Log sheet instead of sending via MailApp

function logAdminSignupNotification(rows) {
  if (!rows || rows.length === 0) return;
  try {
    var firstName = rows[0].firstName || '';
    var lastName  = rows[0].lastName || '';
    var email     = rows[0].email || '';
    var guestFirst = rows[0].guestFirstName || '';
    var guestLast  = rows[0].guestLastName || '';
    var hasGuest   = !!(guestFirst);
    var now = new Date();
    var timestamp  = now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }) + ' PST';

    var classLines = '';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var icon = r.classType === 'In-Person' ? '&#x1F3E0;' : '&#x1F4BB;';
      classLines += '<li>' + icon + ' ' + escHtml(r.className) + ' &mdash; ' + escHtml(r.classDate) + ' (' + escHtml(r.classType) + ')</li>';
    }

    var subject = '\uD83E\uDDD8 New Sign-Up: ' + firstName + ' ' + lastName + ' \u2014 ' + (rows[0].className || '').split(' \u2014 ')[0];

    var body = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
      '<div style="background:#e8f5e9;padding:16px 24px;border-radius:8px 8px 0 0;border-left:4px solid #5B7553;">' +
        '<h2 style="margin:0;color:#5B7553;font-size:18px;">New Sign-Up</h2>' +
      '</div>' +
      '<div style="padding:20px 24px;background:#fff;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px;">' +
        '<p><strong>Student:</strong> ' + escHtml(firstName) + ' ' + escHtml(lastName) + '</p>' +
        '<p><strong>Email:</strong> ' + escHtml(email) + '</p>' +
        '<p><strong>Classes:</strong></p>' +
        '<ul style="margin:4px 0 16px;">' + classLines + '</ul>' +
        (hasGuest ? '<p><strong>Guest:</strong> ' + escHtml(guestFirst) + ' ' + escHtml(guestLast) + '</p>' : '') +
        '<p style="color:#777;font-size:13px;">Signed up at: ' + timestamp + '</p>' +
      '</div>' +
    '</div>';

    var logSheet = getOrCreateEmailLogSheet();
    logSheet.appendRow([
      new Date().toISOString(),
      ADMIN_EMAIL,
      subject,
      body,
      '',
      'Admin - Signup'
    ]);
  } catch (err) {
    Logger.log('Admin signup notification log error: ' + err.toString());
  }
}

function logAdminCancelNotification(studentName, studentEmail, guestName, cancelledClasses, count) {
  try {
    var now = new Date();
    var timestamp = now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' }) + ' PST';

    var classLines = '';
    for (var i = 0; i < cancelledClasses.length; i++) {
      var c = cancelledClasses[i];
      classLines += '<li>' + escHtml(c.className) + ' &mdash; ' + escHtml(c.classDate) + ' (' + escHtml(c.classType) + ')</li>';
    }

    var subject = '\u274C Cancellation: ' + studentName + ' \u2014 ' + (cancelledClasses[0] ? cancelledClasses[0].className : '').split(' \u2014 ')[0];

    var body = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
      '<div style="background:#ffebee;padding:16px 24px;border-radius:8px 8px 0 0;border-left:4px solid #c62828;">' +
        '<h2 style="margin:0;color:#c62828;font-size:18px;">Cancellation</h2>' +
      '</div>' +
      '<div style="padding:20px 24px;background:#fff;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 8px 8px;">' +
        '<p><strong>Student:</strong> ' + escHtml(studentName) + '</p>' +
        '<p><strong>Email:</strong> ' + escHtml(studentEmail) + '</p>' +
        '<p><strong>Cancelled classes:</strong></p>' +
        '<ul style="margin:4px 0 16px;">' + classLines + '</ul>' +
        (guestName ? '<p><strong>Guest also cancelled:</strong> ' + escHtml(guestName) + '</p>' : '') +
        '<p><strong>Rows removed:</strong> ' + count + '</p>' +
        '<p style="color:#777;font-size:13px;">Cancelled at: ' + timestamp + '</p>' +
      '</div>' +
    '</div>';

    var logSheet = getOrCreateEmailLogSheet();
    logSheet.appendRow([
      new Date().toISOString(),
      ADMIN_EMAIL,
      subject,
      body,
      '',
      'Admin - Cancel'
    ]);
  } catch (err) {
    Logger.log('Admin cancel notification log error: ' + err.toString());
  }
}

// ========== WAITLIST HANDLER ==========
function handleWaitlist(rows) {
  try {
    var ss = getOrCreateWaitlistSpreadsheet();
    var sheet = getOrCreateWaitlistSheet(ss);

    rows.forEach(function(row) {
      var newRow = sheet.getLastRow() + 1;
      var values = [[
        row.timestamp,
        row.firstName,
        row.lastName,
        row.email,
        row.className,
        row.classDate,
        row.classType,
        row.guestFirstName || '',
        row.guestLastName || '',
        '', // Status
        '', // Notified At
        row.device || '',
        row.browser || '',
        row.city || '',
        row.state || '',
        row.zip || ''
      ]];
      var range = sheet.getRange(newRow, 1, 1, 16);
      range.setNumberFormat('@');
      range.setValues(values);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== CANCELLATION HANDLER ==========
function handleCancellation(token) {
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss);
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return { status: 'not_found', message: 'No registrations found.' };
    }

    var numCols = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

    var rowsToDelete = [];
    var cancelledClasses = [];
    var studentName = '';
    var studentEmail = '';
    var guestName = '';

    for (var i = 0; i < data.length; i++) {
      var rowToken = (data[i][11] || '').toString().trim();
      if (rowToken === token) {
        rowsToDelete.push(i + 2);
        cancelledClasses.push({
          className: (data[i][4] || '').toString(),
          classDate: (data[i][5] || '').toString(),
          classType: (data[i][6] || '').toString()
        });
        if (!studentName) {
          studentName = (data[i][1] || '') + ' ' + (data[i][2] || '');
          studentEmail = (data[i][3] || '').toString();
        }
        var gf = (data[i][8] || '').toString().trim();
        var gl = (data[i][9] || '').toString().trim();
        if (gf && !guestName) guestName = gf + ' ' + gl;
      }
    }

    if (rowsToDelete.length === 0) {
      return { status: 'not_found', message: 'This cancellation link has already been used or the registration was not found.' };
    }

    // Delete rows bottom-up to preserve indices
    // Google Sheets throws "not possible to delete all non-frozen rows" if
    // deleting would leave zero non-frozen rows. Clear content instead for the last row.
    rowsToDelete.sort(function(a, b) { return b - a; });
    var totalDataRows = sheet.getLastRow() - 1; // exclude header
    for (var d = 0; d < rowsToDelete.length; d++) {
      if (totalDataRows <= 1) {
        // Last data row — clear content instead of deleting to avoid Sheets error
        sheet.getRange(rowsToDelete[d], 1, 1, sheet.getLastColumn()).clearContent();
      } else {
        sheet.deleteRow(rowsToDelete[d]);
      }
      totalDataRows--;
    }

    // Log admin cancel notification
    logAdminCancelNotification(studentName.trim(), studentEmail, guestName, cancelledClasses, rowsToDelete.length);

    return {
      status: 'cancelled',
      studentName: studentName.trim(),
      studentEmail: studentEmail,
      guestName: guestName,
      classes: cancelledClasses,
      count: rowsToDelete.length
    };

  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ========== DUPLICATE CHECK & ALL GET HANDLERS ==========
// ========== MOCK MEET LINK FOR TESTING ==========
// In test mode, we simulate Zoom behavior using Script Properties.
// Tests can set up Zoom state via ?action=set_meet_link and ?action=clear_meet_state
// The mock checks if a class is online, starts within 30 min, and has simulated Zoom state.

var LATE_SIGNUP_WINDOW_MIN = 40; // keep >= sendMeetInvites upper window (prod parity)
function checkAndCreateMeetForLateSignup(rows) {
  if (!rows || rows.length === 0) return '';

  var now = new Date();
  var pstNow = new Date(now.toLocaleString('en-US', { timeZone: MEET_TZ }));
  var currentTotalMin = pstNow.getHours() * 60 + pstNow.getMinutes();
  var cache = PropertiesService.getScriptProperties();
  var meetLink = '';

  // Also check for test-forced Zoom state (set via ?action=set_meet_link)
  var forcedLink = cache.getProperty('test_force_meet_link');
  if (forcedLink) {
    Logger.log('Test: forced Zoom link found: ' + forcedLink);
    return forcedLink;
  }

  var occurrences = getOccurrencesOnPacificDate_(pstNow);

  for (var c = 0; c < occurrences.length; c++) {
    var cls = occurrences[c];

    if (cls.type !== 'online') continue;

    var minutesUntilClass = (cls.startH * 60 + cls.startM) - currentTotalMin;

    if (minutesUntilClass > LATE_SIGNUP_WINDOW_MIN || minutesUntilClass < -15) continue;

    // Check if this student signed up for this class
    var classDate = cls.classDate;
    var signedUpForThis = false;
    for (var r = 0; r < rows.length; r++) {
      var rowType = (rows[r].classType || '').toString().trim().toLowerCase();
      var rowDate = (rows[r].classDate || '').toString().trim();
      var dateMatches = rowDate && (rowDate === classDate ||
                                    extractMonthDay_(rowDate) === extractMonthDay_(classDate));
      if (rowType === 'online' && dateMatches) {
        signedUpForThis = true;
        break;
      }
    }
    if (!signedUpForThis) continue;

    // Check for existing mock Zoom link
    var linkKey = 'meet_link_' + cls.id + '_' + classDate;
    var existingLink = cache.getProperty(linkKey);

    if (existingLink) {
      meetLink = existingLink;
      Logger.log('Test: found existing mock Zoom link: ' + meetLink);
    } else {
      // Create a mock Zoom link
      meetLink = 'https://zoom.us/j/test-yoga-' + cls.id + '-' + Date.now();
      cache.setProperty(linkKey, meetLink);
      var sentKey = 'meet_sent_' + cls.id + '_' + classDate;
      cache.setProperty(sentKey, new Date().toISOString());
      Logger.log('Test: created mock Zoom link: ' + meetLink);
    }
  }

  return meetLink;
}

// Returns "Sunday, April 13, 2026" — matches what signup.html stores.
function formatClassDate(date) {
  return Utilities.formatDate(date, MEET_TZ, "EEEE, MMMM d, yyyy");
}

// Tolerant date extraction — returns "April 13" from any stored format.
function extractMonthDay_(s) {
  if (s instanceof Date) {
    var fullMonths = ['January','February','March','April','May','June',
                      'July','August','September','October','November','December'];
    return fullMonths[s.getMonth()] + ' ' + s.getDate();
  }
  var FULL = 'January|February|March|April|May|June|July|August|September|October|November|December';
  var SHORT = 'Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec';
  var shortToFull = {Jan:'January',Feb:'February',Mar:'March',Apr:'April',May:'May',
                     Jun:'June',Jul:'July',Aug:'August',Sep:'September',
                     Oct:'October',Nov:'November',Dec:'December'};
  var m = String(s).match(new RegExp('\\b(' + FULL + ')\\s+(\\d{1,2})\\b', 'i'));
  if (m) return m[1] + ' ' + parseInt(m[2]);
  m = String(s).match(new RegExp('\\b(' + SHORT + ')\\s+(\\d{1,2})\\b', 'i'));
  if (m) return shortToFull[m[1]] + ' ' + parseInt(m[2]);
  return String(s).trim();
}

function doGet(e) {
  var params = e ? e.parameter : {};

  // Public read endpoint — the schedule as JSON for the website.
  if (params.action === 'schedule') {
    try {
      var sched = getSchedule();
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', classes: sched.classes, exceptions: sched.exceptions }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ===== TEST-ONLY: Read sheet data =====
  if (params.action === 'read_sheet') {
    return handleReadSheet(params);
  }

  // ===== TEST-ONLY: Cleanup all test sheets =====
  if (params.action === 'cleanup') {
    return handleCleanup();
  }

  // ===== TEST-ONLY: Trigger archive on demand =====
  if (params.action === 'trigger_archive') {
    return handleTriggerArchive();
  }

  // ===== TEST-ONLY: Set a forced Zoom link (simulates Zoom meeting existing) =====
  if (params.action === 'set_meet_link') {
    var link = (params.link || '').trim();
    if (!link) link = 'https://zoom.us/j/test-forced-' + Date.now();
    PropertiesService.getScriptProperties().setProperty('test_force_meet_link', link);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', meetLink: link }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ===== TEST-ONLY: Clear all Meet state =====
  if (params.action === 'clear_meet_state') {
    var cache = PropertiesService.getScriptProperties();
    var allKeys = cache.getKeys();
    for (var k = 0; k < allKeys.length; k++) {
      if (allKeys[k].indexOf('meet_') === 0 || allKeys[k] === 'test_force_meet_link') {
        cache.deleteProperty(allKeys[k]);
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'Meet state cleared' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ===== TEST-ONLY: Fire the morning-of reminder pass =====
  // ?action=trigger_reminders[&date=YYYY-MM-DD]  — date targets a class day so
  // the pass is deterministic (the real daily trigger runs with no date).
  if (params.action === 'trigger_reminders') {
    try {
      sendClassReminders(params.date || null);
      return jsonOut_({ status: 'ok', date: params.date || null });
    } catch (err) {
      return jsonOut_({ status: 'error', message: err.toString() });
    }
  }

  // ===== TEST-ONLY: Drive Phase-3 admin actions WITHOUT a Google ID token =====
  // (prod gates these behind verifyAdminToken_; the test harness can't get a
  //  real token, so it calls the same helper functions directly.)
  // ?action=test_cancel_occurrence&classId=..&date=YYYY-MM-DD[&dateDisplay=..][&notifyWaitlist=true]
  if (params.action === 'test_cancel_occurrence') {
    return jsonOut_(adminCancelOccurrence_(params));
  }
  // ?action=test_move_occurrence&classId=..&date=..&newDate=..[&newStartTime=..][&notify=true][&notifyWaitlist=true][&dateDisplay=..]
  if (params.action === 'test_move_occurrence') {
    return jsonOut_(adminMoveOccurrence_(params));
  }

  // Cancel preview
  if (params.action === 'cancel_preview') {
    var token = (params.token || '').trim();
    if (!token) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'not_found', message: 'Missing token' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    try {
      var ss = getOrCreateSpreadsheet();
      var sheet = getOrCreateSheet(ss);
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'not_found' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      var classes = [];
      var studentName = '', guestName = '';
      for (var i = 0; i < data.length; i++) {
        if ((data[i][11] || '').toString().trim() === token) {
          if (!studentName) {
            studentName = ((data[i][1] || '') + ' ' + (data[i][2] || '')).trim();
          }
          var gf = (data[i][8] || '').toString().trim();
          var gl = (data[i][9] || '').toString().trim();
          if (gf && !guestName) guestName = (gf + ' ' + gl).trim();
          classes.push({
            className: (data[i][4] || '').toString(),
            classDate: (data[i][5] || '').toString(),
            classType: (data[i][6] || '').toString()
          });
        }
      }
      if (classes.length === 0) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'not_found', message: 'This cancellation link has already been used or the registration was not found.' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'found', studentName: studentName, guestName: guestName, classes: classes }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Cancel (actually deletes rows)
  if (params.action === 'cancel') {
    var token = (params.token || '').trim();
    if (!token) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Missing token' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var result = handleCancellation(token);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Capacity check
  if (params.action === 'capacity') {
    try {
      var className = (params.className || '').trim();
      var classDate = (params.classDate || '').trim();

      var ss = getOrCreateSpreadsheet();
      var sheet = getOrCreateSheet(ss);
      var lastRow = sheet.getLastRow();
      var count = 0;

      if (lastRow > 1) {
        var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
        for (var i = 0; i < data.length; i++) {
          var rowClass = (data[i][4] || '').toString().trim();
          var rowDate  = (data[i][5] || '').toString().trim();
          if (rowClass === className && rowDate === classDate) {
            count++;
            var guestName = (data[i][8] || '').toString().trim();
            if (guestName) count++;
          }
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'ok',
          count: count,
          className: className,
          classDate: classDate
        }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString(), count: 0 }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Duplicate / waiver check
  if (params.action === 'check') {
    try {
      var firstName = (params.firstName || '').trim().toLowerCase();
      var lastName  = (params.lastName || '').trim().toLowerCase();
      var email     = (params.email || '').trim().toLowerCase();
      var classDates = (params.classDates || '').split(';;').filter(Boolean);

      var ss = getOrCreateSpreadsheet();
      var sheet = getOrCreateSheet(ss);

      var lastRow = sheet.getLastRow();
      var duplicates = [];
      var hasPriorWaiver = false;

      if (lastRow > 1) {
        var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

        classDates.forEach(function(cd) {
          var parts = cd.split('|');
          var checkClass = (parts[0] || '').trim();
          var checkDate  = (parts[1] || '').trim();

          for (var i = 0; i < data.length; i++) {
            var rowFirst = (data[i][1] || '').toString().trim().toLowerCase();
            var rowLast  = (data[i][2] || '').toString().trim().toLowerCase();
            var rowEmail = (data[i][3] || '').toString().trim().toLowerCase();
            var rowClass = (data[i][4] || '').toString().trim();
            var rowDate  = (data[i][5] || '').toString().trim();

            if (rowFirst === firstName &&
                rowLast === lastName &&
                rowEmail === email &&
                rowClass === checkClass &&
                rowDate === checkDate) {
              duplicates.push(checkClass + ' \u2014 ' + checkDate);
              break;
            }
          }
        });

        for (var j = 0; j < data.length; j++) {
          var rFirst  = (data[j][1] || '').toString().trim().toLowerCase();
          var rLast   = (data[j][2] || '').toString().trim().toLowerCase();
          var rEmail  = (data[j][3] || '').toString().trim().toLowerCase();
          var rWaiver = (data[j][7] || '').toString().trim();

          if (rFirst === firstName &&
              rLast === lastName &&
              rEmail === email &&
              rWaiver.indexOf('YES') === 0) {
            hasPriorWaiver = true;
            break;
          }
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'ok',
          hasDuplicates: duplicates.length > 0,
          duplicates: duplicates,
          hasPriorWaiver: hasPriorWaiver
        }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Default GET response
  return ContentService
    .createTextOutput('Test Yoga Signup endpoint is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ========== TEST-ONLY HANDLERS ==========

/**
 * Read all rows from a named sheet.
 * GET ?action=read_sheet&sheet=Sign-Ups
 * GET ?action=read_sheet&sheet=Archive
 * GET ?action=read_sheet&sheet=Waitlist
 * GET ?action=read_sheet&sheet=Waitlist Archive
 * GET ?action=read_sheet&sheet=Test Email Log
 */
function handleReadSheet(params) {
  try {
    var sheetName = (params.sheet || '').trim();
    if (!sheetName) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Missing sheet parameter' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Determine which spreadsheet to look in
    var ss;
    if (sheetName === 'Waitlist' || sheetName === 'Waitlist Archive') {
      ss = getOrCreateWaitlistSpreadsheet();
    } else {
      ss = getOrCreateSpreadsheet();
    }

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', headers: [], rows: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', headers: headers, rows: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];
    var rows = [];

    for (var i = 1; i < allData.length; i++) {
      var rowObj = {};
      for (var j = 0; j < headers.length; j++) {
        rowObj[headers[j]] = (allData[i][j] || '').toString();
      }
      rows.push(rowObj);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', headers: headers, rows: rows }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Cleanup: wipe all data rows from all test sheets (keep headers).
 * GET ?action=cleanup
 */
function handleCleanup() {
  try {
    var signupCounts = { signups: 0, archive: 0, emailLog: 0 };
    var waitlistCounts = { waitlist: 0, waitlistArchive: 0 };

    // Clean signup spreadsheet sheets
    var signupSS = getOrCreateSpreadsheet();
    signupCounts.signups = clearSheetData_(signupSS, 'Sign-Ups');
    signupCounts.archive = clearSheetData_(signupSS, 'Archive');
    signupCounts.emailLog = clearSheetData_(signupSS, 'Test Email Log');
    // Exceptions is mutable test state (cancel/move write here); reset it too so
    // a leftover cancelled/moved row can't suppress a later occurrence.
    signupCounts.exceptions = clearSheetData_(signupSS, 'Exceptions');

    // Clean waitlist spreadsheet sheets
    var waitlistSS = getOrCreateWaitlistSpreadsheet();
    waitlistCounts.waitlist = clearSheetData_(waitlistSS, 'Waitlist');
    waitlistCounts.waitlistArchive = clearSheetData_(waitlistSS, 'Waitlist Archive');

    // Clear reminder dedupe properties so reminder tests start fresh.
    var remProps = PropertiesService.getScriptProperties();
    var allRemProps = remProps.getProperties();
    for (var pk in allRemProps) {
      if (pk.indexOf('reminder_sent_') === 0) remProps.deleteProperty(pk);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'ok',
        signups: signupCounts.signups,
        archive: signupCounts.archive,
        emailLog: signupCounts.emailLog,
        exceptions: signupCounts.exceptions,
        waitlist: waitlistCounts.waitlist,
        waitlistArchive: waitlistCounts.waitlistArchive
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Clear all data rows from a sheet, preserving headers.
 * Returns number of rows deleted.
 */
function clearSheetData_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  var rowsDeleted = lastRow - 1;

  // Google Sheets throws "not possible to delete all non-frozen rows" if
  // deleting rows would leave zero non-frozen rows (even with a frozen header).
  // Fix: keep one blank data row, delete the rest, then clear it.
  if (rowsDeleted === 1) {
    // Only one data row — just clear its content (don't delete it)
    sheet.getRange(2, 1, 1, sheet.getLastColumn()).clearContent();
  } else {
    // Delete all but one data row, then clear the remaining one
    sheet.deleteRows(3, rowsDeleted - 1);
    sheet.getRange(2, 1, 1, sheet.getLastColumn()).clearContent();
  }
  return rowsDeleted;
}

/**
 * Seed: insert test data rows directly into a named sheet.
 * POST { action: 'seed', sheet: 'Sign-Ups', rows: [...] }
 */
function handleSeed(data) {
  try {
    var sheetName = (data.sheet || '').trim();
    var rows = data.rows || [];

    if (!sheetName || rows.length === 0) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Missing sheet or rows' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Determine spreadsheet
    var ss;
    if (sheetName === 'Waitlist' || sheetName === 'Waitlist Archive') {
      ss = getOrCreateWaitlistSpreadsheet();
    } else {
      ss = getOrCreateSpreadsheet();
    }

    // Get or create sheet with appropriate headers
    var sheet;
    if (sheetName === 'Sign-Ups') {
      sheet = getOrCreateSheet(ss);
    } else if (sheetName === 'Waitlist') {
      sheet = getOrCreateWaitlistSheet(ss);
    } else {
      sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      }
    }

    // Append rows based on sheet type (using setValues to prevent date auto-formatting)
    rows.forEach(function(row) {
      var newRow = sheet.getLastRow() + 1;
      if (sheetName === 'Sign-Ups' || sheetName === 'Archive') {
        var values = [[
          row.timestamp || new Date().toISOString(),
          row.firstName || '',
          row.lastName || '',
          row.email || '',
          row.className || '',
          row.classDate || '',
          row.classType || 'In-Person',
          row.liabilityWaiver || 'YES \u2014 Accepted',
          row.guestFirstName || '',
          row.guestLastName || '',
          row.guestOf || '',
          row.cancelToken || '',
          row.device || '',
          row.browser || '',
          row.city || '',
          row.state || '',
          row.zip || '',
          row.timezone || ''
        ]];
        var range = sheet.getRange(newRow, 1, 1, 18);
        range.setNumberFormat('@');
        range.setValues(values);
      } else if (sheetName === 'Waitlist' || sheetName === 'Waitlist Archive') {
        var values = [[
          row.timestamp || new Date().toISOString(),
          row.firstName || '',
          row.lastName || '',
          row.email || '',
          row.className || '',
          row.classDate || '',
          row.classType || 'In-Person',
          row.guestFirstName || '',
          row.guestLastName || '',
          row.status || '',
          row.notifiedAt || '',
          row.device || '',
          row.browser || '',
          row.city || '',
          row.state || '',
          row.zip || ''
        ]];
        var range = sheet.getRange(newRow, 1, 1, 16);
        range.setNumberFormat('@');
        range.setValues(values);
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', rowsInserted: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Trigger archive on demand (for testing).
 * GET ?action=trigger_archive
 */
function handleTriggerArchive() {
  try {
    archivePastSignups();
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'Archive triggered' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== WAITLIST PROCESSOR ==========

function processWaitlistForClass(className, classDate) {
  var signupSS = getOrCreateSpreadsheet();
  var signupSheet = getOrCreateSheet(signupSS);
  var signupLastRow = signupSheet.getLastRow();
  var confirmedCount = 0;
  var signupData = [];

  if (signupLastRow > 1) {
    signupData = signupSheet.getRange(2, 1, signupLastRow - 1, 11).getValues();
    for (var i = 0; i < signupData.length; i++) {
      var rowClass = (signupData[i][4] || '').toString().trim();
      var rowDate  = (signupData[i][5] || '').toString().trim();
      if (rowClass === className && rowDate === classDate) {
        confirmedCount++;
        var guestName = (signupData[i][8] || '').toString().trim();
        if (guestName) confirmedCount++;
      }
    }
  }

  var waitlistSS = getOrCreateWaitlistSpreadsheet();
  var waitlistSheet = getOrCreateWaitlistSheet(waitlistSS);
  var waitlistLastRow = waitlistSheet.getLastRow();

  if (waitlistLastRow <= 1) return { notified: 0 };

  var headers = waitlistSheet.getRange(1, 1, 1, waitlistSheet.getLastColumn()).getValues()[0];
  var statusColIndex = headers.indexOf('Status');

  if (statusColIndex === -1) {
    var newCol = waitlistSheet.getLastColumn() + 1;
    waitlistSheet.getRange(1, newCol).setValue('Status').setFontWeight('bold');
    statusColIndex = newCol - 1;
  }

  var numCols = waitlistSheet.getLastColumn();
  var waitlistData = waitlistSheet.getRange(2, 1, waitlistLastRow - 1, numCols).getValues();

  // Clean up: delete waitlist entries for people who already signed up
  var rowsToDelete = [];
  for (var e = 0; e < waitlistData.length; e++) {
    var eClass = (waitlistData[e][4] || '').toString().trim();
    var eDate  = (waitlistData[e][5] || '').toString().trim();
    if (eClass !== className || eDate !== classDate) continue;

    var eName  = (waitlistData[e][1] || '').toString().trim().toLowerCase();
    var eLast  = (waitlistData[e][2] || '').toString().trim().toLowerCase();
    var eEmail = (waitlistData[e][3] || '').toString().trim().toLowerCase();

    for (var s = 0; s < signupData.length; s++) {
      var sFirst = (signupData[s][1] || '').toString().trim().toLowerCase();
      var sLast  = (signupData[s][2] || '').toString().trim().toLowerCase();
      var sEmail = (signupData[s][3] || '').toString().trim().toLowerCase();
      var sClass = (signupData[s][4] || '').toString().trim();
      var sDate  = (signupData[s][5] || '').toString().trim();

      if (sFirst === eName && sLast === eLast && sEmail === eEmail &&
          sClass === className && sDate === classDate) {
        rowsToDelete.push(e + 2);
        break;
      }
    }
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  var wlDataRows = waitlistSheet.getLastRow() - 1;
  for (var d = 0; d < rowsToDelete.length; d++) {
    if (wlDataRows <= 1) {
      waitlistSheet.getRange(rowsToDelete[d], 1, 1, waitlistSheet.getLastColumn()).clearContent();
    } else {
      waitlistSheet.deleteRow(rowsToDelete[d]);
    }
    wlDataRows--;
  }

  var capClass = getScheduleClassByLabel_(className);
  var capacity = capClass ? capClass.capacity : null;
  var availableSpots = (capacity == null) ? Number.MAX_SAFE_INTEGER : (capacity - confirmedCount);
  if (availableSpots <= 0) return { notified: 0 };

  waitlistLastRow = waitlistSheet.getLastRow();
  if (waitlistLastRow <= 1) return { notified: 0 };

  numCols = waitlistSheet.getLastColumn();
  waitlistData = waitlistSheet.getRange(2, 1, waitlistLastRow - 1, numCols).getValues();

  var notifiedCount = 0;

  for (var w = 0; w < waitlistData.length; w++) {
    var wClass  = (waitlistData[w][4] || '').toString().trim();
    var wDate   = (waitlistData[w][5] || '').toString().trim();
    var wStatus = (waitlistData[w][statusColIndex] || '').toString().trim();

    if (wClass !== className || wDate !== classDate) continue;
    if (wStatus === 'Notified') continue;

    var rowNum = w + 2;
    waitlistSheet.getRange(rowNum, statusColIndex + 1).setValue('Notified');

    var firstName = (waitlistData[w][1] || '').toString().trim();
    var email     = (waitlistData[w][3] || '').toString().trim();

    // LOG email instead of sending (test mode)
    logWaitlistNotificationEmail(email, firstName, wClass, wDate);

    notifiedCount++;
  }

  return { notified: notifiedCount, available: availableSpots };
}

function processAllWaitlists() {
  var waitlistSS;
  try {
    waitlistSS = getOrCreateWaitlistSpreadsheet();
  } catch (e) {
    return;
  }

  var waitlistSheet = getOrCreateWaitlistSheet(waitlistSS);
  var lastRow = waitlistSheet.getLastRow();
  if (lastRow <= 1) return;

  var headers = waitlistSheet.getRange(1, 1, 1, waitlistSheet.getLastColumn()).getValues()[0];
  var statusColIndex = headers.indexOf('Status');
  var numCols = waitlistSheet.getLastColumn();
  var data = waitlistSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  var classDatePairs = {};
  for (var i = 0; i < data.length; i++) {
    var status = statusColIndex >= 0 ? (data[i][statusColIndex] || '').toString().trim() : '';
    if (status === 'Notified') continue;

    var cls = (data[i][4] || '').toString().trim();
    var dt  = (data[i][5] || '').toString().trim();
    if (cls && dt) {
      classDatePairs[cls + '|||' + dt] = { className: cls, classDate: dt };
    }
  }

  var keys = Object.keys(classDatePairs);
  for (var k = 0; k < keys.length; k++) {
    var pair = classDatePairs[keys[k]];
    processWaitlistForClass(pair.className, pair.classDate);
  }
}

// ========== AUTO-ARCHIVE PAST SIGN-UPS ==========
var ARCHIVE_TZ = 'America/Los_Angeles';

function archivePastSignups() {
  var startTimes = buildStartTimeMap_();
  archiveSheet_(TEST_SIGNUP_SS_NAME, 'Sign-Ups', 'Archive', 6, startTimes);
  archiveSheet_(TEST_WAITLIST_SS_NAME, 'Waitlist', 'Waitlist Archive', 6, startTimes);
}

function buildStartTimeMap_() {
  var map = {};
  try {
    var sched = getSchedule();
    sched.classes.forEach(function(c) { map[c.label] = { startH: c.startH, startM: c.startM }; });
  } catch (e) {
    Logger.log('buildStartTimeMap_ error: ' + e);
  }
  return map;
}

function archiveSheet_(ssName, sheetName, archiveName, dateCol, startTimes) {
  startTimes = startTimes || {};
  var files = DriveApp.getFilesByName(ssName);
  if (!files.hasNext()) return;
  var ss = SpreadsheetApp.open(files.next());
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var archive = ss.getSheetByName(archiveName);
  if (!archive) {
    archive = ss.insertSheet(archiveName);
    archive.appendRow(data[0]);
    archive.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
    archive.setFrozenRows(1);
  }

  var now = new Date();
  var pstNow = new Date(now.toLocaleString('en-US', { timeZone: ARCHIVE_TZ }));

  var rowsToArchive = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var classDateStr = data[i][dateCol - 1];
    var className = data[i][dateCol - 2] || '';

    if (!classDateStr) continue;

    // Parse the class date string robustly.
    // Stored as "Sunday, April 13, 2026" (new format) or "Sunday, April 13" (old, no year).
    // new Date() chokes on the weekday prefix, so strip it first.
    var classDate;
    if (classDateStr instanceof Date) {
      classDate = classDateStr;
    } else {
      var s = String(classDateStr).trim();
      s = s.replace(/^[A-Za-z]+,\s*/, ''); // strip weekday prefix
      if (!/\d{4}/.test(s)) {
        var inferYear = pstNow.getFullYear();
        var attempt = new Date(s + ', ' + inferYear);
        if (!isNaN(attempt.getTime()) && (pstNow - attempt) > 7 * 24 * 3600 * 1000) {
          attempt = new Date(s + ', ' + (inferYear + 1));
        }
        classDate = attempt;
      } else {
        classDate = new Date(s);
      }
    }
    if (isNaN(classDate.getTime())) continue;

    var startH = 18, startM = 0;
    var st = startTimes[(className || '').toString()];
    if (st) { startH = st.startH; startM = st.startM; }

    var cutoff = new Date(classDate);
    cutoff.setHours(startH, startM + 15, 0, 0);

    if (pstNow > cutoff) {
      rowsToArchive.push({ index: i, row: data[i] });
    }
  }

  var archiveDataRows = sheet.getLastRow() - 1;
  for (var j = 0; j < rowsToArchive.length; j++) {
    archive.appendRow(rowsToArchive[j].row);
    if (archiveDataRows <= 1) {
      sheet.getRange(rowsToArchive[j].index + 1, 1, 1, sheet.getLastColumn()).clearContent();
    } else {
      sheet.deleteRow(rowsToArchive[j].index + 1);
    }
    archiveDataRows--;
  }
}

// ========== HELPERS ==========
// These point to TEST spreadsheets (not production!)

function getOrCreateSpreadsheet() {
  var files = DriveApp.getFilesByName(TEST_SIGNUP_SS_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  return SpreadsheetApp.create(TEST_SIGNUP_SS_NAME);
}

function getOrCreateWaitlistSpreadsheet() {
  var files = DriveApp.getFilesByName(TEST_WAITLIST_SS_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  return SpreadsheetApp.create(TEST_WAITLIST_SS_NAME);
}

function getOrCreateSheet(ss) {
  var sheet = ss.getSheetByName('Sign-Ups');
  if (!sheet) {
    sheet = ss.insertSheet('Sign-Ups');
    sheet.appendRow([
      'Timestamp',
      'First Name',
      'Last Name',
      'Email',
      'Class',
      'Class Date',
      'Class Type',
      'Liability Waiver',
      'Guest First Name',
      'Guest Last Name',
      'Guest Of',
      'Cancel Token',
      'Device',
      'Browser',
      'City',
      'State',
      'Zip Code'
    ]);
    sheet.getRange(1, 1, 1, 17).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateWaitlistSheet(ss) {
  var sheet = ss.getSheetByName('Waitlist');
  if (!sheet) {
    sheet = ss.insertSheet('Waitlist');
    sheet.appendRow([
      'Timestamp',
      'First Name',
      'Last Name',
      'Email',
      'Class',
      'Class Date',
      'Class Type',
      'Guest First Name',
      'Guest Last Name',
      'Status',
      'Notified At',
      'Device',
      'Browser',
      'City',
      'State',
      'Zip Code'
    ]);
    sheet.getRange(1, 1, 1, 16).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
