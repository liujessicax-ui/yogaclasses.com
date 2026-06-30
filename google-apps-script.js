/**
 * Google Apps Script — Yoga Signup Sheet Writer
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com and create a new project
 * 2. Paste this entire file into the editor (replace any existing code)
 * 3. Click "Deploy" → "New deployment"
 * 4. Choose type: "Web app"
 * 5. Set "Execute as": Me (your Google account)
 * 6. Set "Who has access": Anyone
 * 7. Click "Deploy" and authorize when prompted
 * 8. Copy the web app URL
 * 9. Paste it into signup.html where it says:
 *      const SHEETS_WEB_APP_URL = '';
 *
 * The script will automatically create a spreadsheet called "Yoga Signup"
 * in your Google Drive if one doesn't already exist.
 *
 * FEATURES:
 * - POST: Writes sign-up rows to the spreadsheet
 * - POST with action='waitlist': Writes to Yoga Waitlist spreadsheet
 * - GET with ?action=check: Checks for duplicate registrations
 * - GET with ?action=capacity: Returns current sign-up count for a class+date
 */

// ========== WRITE SIGN-UPS ==========
// YOUR WEBSITE URL — update this when you deploy the site
var SITE_URL = 'https://yogawithjessica.com'; // change to your actual domain
var ADMIN_EMAIL = 'xiaojing25@gmail.com';

// ============================================================================
// ========== ADMIN CONSOLE AUTH (Google ID-token verification) ==========
// ============================================================================
// The admin page (yogawithjessica.com/admin/) signs in with Google Identity
// Services, gets an ID token (JWT), and sends it with each admin request. We
// verify it server-side via Google's tokeninfo endpoint — Google checks the
// signature + expiry — then assert aud === our client ID, iss === Google,
// email_verified, and that the email is on the allowlist below. NO shared
// secret; the client-sent email is never trusted. See verifyAdminToken_().
//
// PASTE the OAuth "Web application" Client ID here (same value the admin page
// uses). It is NOT a secret — it ships to the browser too.
var ADMIN_OAUTH_CLIENT_ID = '83041676087-iia2s4jjtb3n6je56so3mfbdin9lpe0u.apps.googleusercontent.com';
var ADMIN_ALLOWLIST = ['liu.jessica.x@gmail.com'];

// ============================================================================
// ========== SCHEDULE — SINGLE SOURCE OF TRUTH (Google Sheet) ==========
// ============================================================================
// The class schedule lives in two tabs of the "Yoga Signup" spreadsheet:
//   - "Schedule":   one row per recurring class series
//   - "Exceptions": one row per single-date cancel / move / extra
// The public site and the Zoom/archive triggers all read from here, so the
// schedule can be changed from the Sheet (or the admin console) with no code
// edits or redeploy.
//
// ONE-TIME SETUP: run setupScheduleSheets() once from the editor to create and
// seed the tabs from the 3 current classes. Safe to re-run (won't duplicate).

var SCHEDULE_SHEET_NAME   = 'Schedule';
var EXCEPTIONS_SHEET_NAME = 'Exceptions';
var SCHEDULE_CACHE_KEY    = 'yoga_schedule_v1';

// Run ONCE manually from the Apps Script editor (Run → setupScheduleSheets).
function setupScheduleSheets() {
  var ss = getOrCreateSpreadsheet();
  var schedule = getOrCreateScheduleSheet_(ss);
  getOrCreateExceptionsSheet_(ss);

  // Seed only if there are no data rows yet (idempotent).
  if (schedule.getLastRow() <= 1) {
    var seed = [
      ['sunday-online', 'Sunday Evening \u2014 Online via Zoom', 0, '18:00', 75, 'online', '', '', 'Open to Everyone', 'Yoga mat, Strap, Two blocks, Wall space, Yoga chair (ideal), Bolster (ideal)', 'TRUE'],
      ['tuesday-ccv', 'Tuesday Evening \u2014 CCV Clubhouse (In Person)', 2, '18:00', 75, 'inperson', 'CCV Clubhouse', 10, 'CCV Residents Only, In Person', 'Yoga mat, Two blocks, Strap', 'TRUE'],
      ['wednesday-restorative', 'Wednesday Evening \u2014 Restorative Yoga (Online)', 3, '20:00', 75, 'online', '', '', 'Restorative, Open to Everyone', 'Yoga mat, Bolster, Two blocks, Two blankets, Strap, Wall space, Yoga chair (ideal)', 'TRUE']
    ];
    var range = schedule.getRange(2, 1, seed.length, 11);
    range.setNumberFormat('@'); // store everything as plain text (so "18:00" stays a string)
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
    // Keep text columns (id/label/time/tags/props/active/date) from auto-converting.
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
    sheet.getRange(2, 1, 2000, headers.length).setNumberFormat('@'); // dates/times as text
  }
  return sheet;
}

// Read schedule + exceptions directly from the Sheet (uncached).
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

// Cached read (~60s). Cache is busted on any admin write.
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

// Return the effective class occurrences for the given Pacific date, applying
// the active flag and Exceptions (cancelled = removed, moved = away/in, extra =
// added). Used by the Zoom triggers. classDate is the display string that
// matches what sign-up rows store ("EEEE, MMMM d, yyyy").
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

  // Moved-in and one-off "extra" occurrences landing on this date.
  // Skip inactive classes here too, so deactivating/cancelling a series also
  // hides any pending moved-in/extra date (mirrors the client expander, which
  // returns nothing for inactive classes).
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

// ---- small parsers shared by the schedule reader ----
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
  if (v === '' || v == null) return true; // blank defaults to active
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
// email, or null if anything fails. Method: forward the token to Google's
// tokeninfo endpoint (Google validates the signature + expiry and returns the
// decoded claims), then assert aud === our client ID, iss === Google, the email
// is verified, the token isn't expired, and the email is on ADMIN_ALLOWLIST.
// The decision is cached by a hash of the token for the token's remaining life
// so repeated admin actions don't re-hit tokeninfo.
function verifyAdminToken_(idToken) {
  if (!idToken || typeof idToken !== 'string') return null;

  // Refuse to run if the client ID hasn't been configured (fail closed).
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

    // aud must be OUR OAuth client ID (token was minted for this app).
    if (claims.aud !== ADMIN_OAUTH_CLIENT_ID) return null;
    // iss must be Google.
    if (claims.iss !== 'accounts.google.com' && claims.iss !== 'https://accounts.google.com') return null;
    // Email must be present and verified.
    var email = (claims.email || '').toString().trim().toLowerCase();
    var emailVerified = (claims.email_verified === true || claims.email_verified === 'true');
    if (!email || !emailVerified) return null;
    // Not expired (tokeninfo already rejects expired tokens; double-check anyway).
    var exp = parseInt(claims.exp, 10);
    if (!exp || (exp * 1000) <= Date.now()) return null;
    // Email must be on the allowlist.
    var allowed = false;
    for (var i = 0; i < ADMIN_ALLOWLIST.length; i++) {
      if (String(ADMIN_ALLOWLIST[i]).trim().toLowerCase() === email) { allowed = true; break; }
    }
    if (!allowed) return null;

    // Cache the decision until ~30s before the token expires (Apps Script caps
    // cache TTL at 6h). Keyed by the token hash so it can't be reused for a
    // different token, and exp is re-checked on every read above.
    var ttl = Math.min(21600, (exp - Math.floor(Date.now() / 1000)) - 30);
    if (ttl > 0) cache.put(cacheKey, JSON.stringify({ email: email, exp: exp }), ttl);

    return email;
  } catch (err) {
    Logger.log('verifyAdminToken_ error: ' + err);
    return null;
  }
}

// Build the admin read payload: the full schedule (including inactive classes,
// which the public expander hides) plus the next-7-days occurrences with
// confirmed sign-up and waitlist counts per occurrence.
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

// Expand occurrences over the next windowDays Pacific days (applying active +
// exceptions via getOccurrencesOnPacificDate_) and attach sign-up/waitlist
// counts. Mirrors the public 15-min sign-up cutoff so already-passed classes
// today drop off.
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
      // Drop occurrences whose sign-up cutoff (start + 15 min) has already passed today.
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

// Classify how an occurrence on `iso` arose, so the admin UI only offers Move
// on plain recurring dates (moving an already-moved/extra date is ambiguous and
// is managed from the Sheet for now).
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
// label + display date. Tolerant date matching mirrors the capacity endpoint.
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
          if ((data[i][8] || '').toString().trim()) signedUp++; // guest takes a spot
        }
      }
    }
  } catch (err) {
    Logger.log('countSignupsForClass_ signups error: ' + err);
  }

  try {
    var wfiles = DriveApp.getFilesByName('Yoga Waitlist');
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
            if ((wdata[j][7] || '').toString().trim()) waitlist++; // guest
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
// Each function below runs ONLY after verifyAdminToken_ passed (see doPost
// gate) and busts the schedule cache so the public site + triggers pick up the
// change within ~60s. Phase 2: schedule edits only — no student emails.

// Save edits to an existing series. If the label (the sign-up join key) changed,
// migrate future Sign-Ups + Waitlist rows so capacity/Zoom/archive keep matching.
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

// Create a new recurring series (appends a Schedule row with a generated id).
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

// Flip a series active on/off (the inline toggle). Quiet edit — does NOT email
// students; cancelling a series WITH notifications is the Phase 3 flow.
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

// Move a single occurrence (Exceptions status=moved). Upserts on (classId,date)
// so re-moving the same original date updates its row instead of duplicating.
// Phase 2: schedule only — registered students are NOT emailed yet (Phase 3).
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

// Validate + coerce an incoming class object from the admin form.
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

// Serialize a normalized class to its 11-column Schedule row (text-formatted so
// "18:00" / ids / TRUE-FALSE don't get auto-converted by Sheets).
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

// Migrate future Sign-Ups + Waitlist rows from an old class label to a new one
// (the label is the join key). Archives are historical and left untouched.
function relabelSignups_(oldLabel, newLabel) {
  var count = 0;
  count += relabelInSheet_(getOrCreateSpreadsheet(), 'Sign-Ups', oldLabel, newLabel);
  try {
    var wfiles = DriveApp.getFilesByName('Yoga Waitlist');
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
// Token-gated (see doPost ADMIN_ACTIONS gate). Unlike Phase 2, these DO email
// students: cancel a single date, cancel a whole series, move-with-notify, and
// (guarded) delete a series. Each busts the schedule cache. Designed to be DRY
// across prod/test — the only env-specific helpers are openWaitlistSpreadsheetIfExists_()
// and deliverEmail_() (test logs to a sheet instead of sending).

// Open the Waitlist spreadsheet if it exists, else null (never creates one).
// TEST MIRROR: swaps 'Yoga Waitlist' -> TEST_WAITLIST_SS_NAME.
function openWaitlistSpreadsheetIfExists_() {
  var files = DriveApp.getFilesByName('Yoga Waitlist');
  return files.hasNext() ? SpreadsheetApp.open(files.next()) : null;
}

// Single mail primitive for Phase 3 emails. TEST MIRROR: logs to "Test Email
// Log" instead of sending. Keeping all the templates routed through this means
// the email functions themselves are byte-identical across prod/test.
function deliverEmail_(to, subject, htmlBody, type) {
  MailApp.sendEmail({
    to: to,
    subject: subject,
    body: stripHtml(htmlBody),
    htmlBody: htmlBody,
    name: 'Yoga with Jessica',
    replyTo: ADMIN_EMAIL
  });
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
// Does NOT write the Exceptions row — callers decide (single-date writes a
// cancelled exception; series cancel relies on Active=FALSE).
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

// Cancel a single dated occurrence: write a cancelled Exceptions row, email +
// archive registrants (waitlist optional to email, always archived), notify admin.
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

// Cancel a whole series WITH notifications: set Active=FALSE and email + archive
// every upcoming registrant across the series' upcoming dates. Distinct from the
// quiet Phase-2 setActive toggle.
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

// Move a single occurrence WITH optional notification. Writes the moved
// exception (as Phase 2), re-stamps registrant rows old date -> new date so the
// join stays intact, and (if notify) emails students the new date/time.
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

// Guarded delete: remove the Schedule row (and the class's Exceptions rows) ONLY
// when no upcoming registrants remain. Otherwise refuse with code 'has_upcoming'
// so the UI can prompt the admin to Cancel the series first. Archives (historical
// rows) keep their text label and are untouched.
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

// ========== EMAIL HELPER ==========

function stripHtml(html) {
  return (html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<li[^>]*>/gi, '  - ')
    .replace(/<[^>]+>/g, '')
    .replace(/&mdash;/g, '\u2014')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#x1F3E0;/g, '')
    .replace(/&#x1F4BB;/g, '')
    .replace(/&#x1F4F9;/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // ----- Admin console (all actions require a verified Google ID token) ---
    // Read = getScheduleAdmin. Phase-2 writes = updateClass / createClass /
    // setActive / setException (silent single-date Move). Phase-3 writes (these
    // DO email students) = cancelOccurrence / cancelSeries / moveOccurrence /
    // deleteClass. Every write busts the schedule cache.
    var ADMIN_ACTIONS = {
      getScheduleAdmin: 1, updateClass: 1, createClass: 1, setActive: 1, setException: 1,
      cancelOccurrence: 1, cancelSeries: 1, moveOccurrence: 1, deleteClass: 1
    };
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
        case 'cancelOccurrence': return jsonOut_(adminCancelOccurrence_(data));
        case 'cancelSeries':     return jsonOut_(adminCancelSeries_(data));
        case 'moveOccurrence':   return jsonOut_(adminMoveOccurrence_(data));
        case 'deleteClass':      return jsonOut_(adminDeleteClass_(data));
      }
    }

    var rows = data.rows;

    // Check if this is a waitlist submission
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
    // If so, ensure a Meet event exists and add the student to it
    var meetLink = checkAndCreateMeetForLateSignup(rows);

    // Send confirmation email with cancel link (and Meet link if applicable)
    sendConfirmationEmail(rows, cancelToken, meetLink);

    // Notify admin of new sign-up
    sendAdminSignupNotification(rows);

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

// ========== CONFIRMATION EMAIL ==========
function sendConfirmationEmail(rows, cancelToken, meetLink) {
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

  // Cancel link — points to cancel.html on the website
  var cancelUrl = SITE_URL + '/cancel.html?token=' + cancelToken;

  var subject = 'Yoga with Jessica \u2014 Sign-Up Confirmation';

  var body = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +

    // Header
    '<div style="background:#f5f0e8;padding:24px;text-align:center;border-radius:8px 8px 0 0;">' +
      '<h1 style="margin:0;font-family:Georgia,serif;font-size:24px;">' +
        '<a href="' + SITE_URL + '" style="color:#5B7553;text-decoration:none;">yogawithjessica.com</a>' +
      '</h1>' +
      '<p style="margin:6px 0 0;color:#888;font-size:13px;">Sign-Up Confirmation</p>' +
    '</div>' +

    // Body
    '<div style="padding:24px;background:#fff;border:1px solid #e8e4dc;border-top:none;">' +

      '<p style="font-size:15px;">Hi ' + escHtml(firstName) + ',</p>' +

      '<p style="font-size:15px;line-height:1.6;">Thank you for signing up! Here are your confirmed classes:</p>' +

      // Class table
      '<table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:14px;">' +
        '<tr style="background:#f5f0e8;">' +
          '<th style="padding:8px 12px;text-align:left;">Class</th>' +
          '<th style="padding:8px 12px;text-align:left;">Date</th>' +
          '<th style="padding:8px 12px;text-align:left;">Time</th>' +
          '<th style="padding:8px 12px;text-align:left;">Type</th>' +
        '</tr>' +
        classLines +
      '</table>' +

      // Guest info
      (hasGuest ?
        '<div style="background:#f9f7f2;padding:12px 16px;border-radius:6px;margin:16px 0;font-size:14px;">' +
          '<strong>Guest:</strong> ' + escHtml(guestFirst) + ' ' + escHtml(guestLast) +
          '<br><span style="color:#777;">Your guest will need to sign a liability waiver upon arrival to class.</span>' +
        '</div>'
      : '') +

      // Online class note — with or without immediate Zoom link
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

      // Props reminder
      '<p style="font-size:14px;color:#555;line-height:1.6;">' +
        'Don\'t forget to check the <a href="' + SITE_URL + '/props.html" style="color:#5B7553;">Props page</a> for recommended props to bring.' +
      '</p>' +

      // Divider
      '<hr style="border:none;border-top:1px solid #e8e4dc;margin:24px 0;" />' +

      // Cancel section
      '<div style="text-align:center;margin:16px 0;">' +
        '<p style="font-size:14px;color:#555;margin-bottom:12px;">Need to cancel? Click below to cancel your registration' +
          (hasGuest ? ' (this will cancel both your spot and your guest\'s spot).' : '.') +
        '</p>' +
        '<a href="' + cancelUrl + '" style="display:inline-block;background:#b44;color:#fff;padding:10px 28px;border-radius:6px;text-decoration:none;font-size:14px;font-weight:600;">Cancel Registration</a>' +
      '</div>' +

    '</div>' +

    // Footer
    '<div style="padding:16px;text-align:center;background:#f5f0e8;border-radius:0 0 8px 8px;">' +
      '<p style="margin:0;font-size:12px;color:#999;">Yoga with Jessica &mdash; Playa Del Rey, CA</p>' +
      '<p style="margin:6px 0 0;"><a href="' + SITE_URL + '" style="color:#5B7553;font-size:15px;font-weight:600;text-decoration:none;">yogawithjessica.com</a></p>' +
    '</div>' +

  '</div>';

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: stripHtml(body),
    htmlBody: body,
    name: 'Yoga with Jessica',
    replyTo: ADMIN_EMAIL
  });
}

function escHtml(str) {
  return (str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ========== ADMIN NOTIFICATION EMAILS ==========

function sendAdminSignupNotification(rows) {
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

    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject,
      body: stripHtml(body),
      htmlBody: body,
      name: 'Yoga with Jessica',
      replyTo: ADMIN_EMAIL
    });
  } catch (err) {
    Logger.log('Admin signup notification error: ' + err.toString());
  }
}

function sendAdminCancelNotification(studentName, studentEmail, guestName, cancelledClasses, count) {
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

    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject,
      body: stripHtml(body),
      htmlBody: body,
      name: 'Yoga with Jessica',
      replyTo: ADMIN_EMAIL
    });
  } catch (err) {
    Logger.log('Admin cancel notification error: ' + err.toString());
  }
}

// ========== WAITLIST HANDLER ==========
function handleWaitlist(rows) {
  try {
    var ss = getOrCreateWaitlistSpreadsheet();
    var sheet = getOrCreateWaitlistSheet(ss);

    rows.forEach(function(row) {
      sheet.appendRow([
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
      ]);
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
// Called via GET with action=cancel&token=ABC123
// Deletes ALL rows matching the cancel token (covers registrant + guest, all classes in that batch)
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

    // Find rows matching this token (column 12 = index 11)
    var rowsToDelete = [];
    var cancelledClasses = [];
    var studentName = '';
    var studentEmail = '';
    var guestName = '';

    for (var i = 0; i < data.length; i++) {
      var rowToken = (data[i][11] || '').toString().trim();
      if (rowToken === token) {
        rowsToDelete.push(i + 2); // +2 for header and 0-index
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

    // Delete rows bottom-up to preserve indices (protect last non-frozen row)
    rowsToDelete.sort(function(a, b) { return b - a; });
    var totalDataRows = sheet.getLastRow() - 1;
    for (var d = 0; d < rowsToDelete.length; d++) {
      if (totalDataRows <= 1) {
        sheet.getRange(rowsToDelete[d], 1, 1, sheet.getLastColumn()).clearContent();
      } else {
        sheet.deleteRow(rowsToDelete[d]);
      }
      totalDataRows--;
    }

    // Notify admin of cancellation
    sendAdminCancelNotification(studentName.trim(), studentEmail, guestName, cancelledClasses, rowsToDelete.length);

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

// ========== DUPLICATE CHECK ==========
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

  // If action=cancel_preview, look up the token without deleting
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

  // If action=cancel, process a cancellation request (actually deletes rows)
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

  // If action=capacity, return confirmed sign-up count for a class+date.
  // Only counts the Sign-Ups sheet (no held spots — waitlist is first-come-first-served).
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
        var targetMonthDay = extractMonthDay_(classDate);
        for (var i = 0; i < data.length; i++) {
          var rowClass = (data[i][4] || '').toString().trim();
          var rowDate  = (data[i][5] || '').toString().trim();
          var dateMatch = (rowDate === classDate) || (extractMonthDay_(rowDate) === targetMonthDay);
          if (rowClass === className && dateMatch) {
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

  // If action=check, look for duplicates
  if (params.action === 'check') {
    try {
      var firstName = (params.firstName || '').trim().toLowerCase();
      var lastName  = (params.lastName || '').trim().toLowerCase();
      var email     = (params.email || '').trim().toLowerCase();
      // classDates is a ;;-separated list of "className|classDate" pairs
      var classDates = (params.classDates || '').split(';;').filter(Boolean);

      var ss = getOrCreateSpreadsheet();
      var sheet = getOrCreateSheet(ss);

      var lastRow = sheet.getLastRow();
      var duplicates = [];
      var hasPriorWaiver = false;

      if (lastRow > 1) {
        // Read all data rows (skip header)
        // Columns: Timestamp(0), First(1), Last(2), Email(3), Class(4), Date(5), Type(6), Waiver(7), GuestName(8), GuestOf(9)
        var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

        // Check for duplicate class registrations
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

        // Check if this person has EVER signed the liability waiver before
        // (same first name + last name + email, waiver column = "YES — Accepted")
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
    .createTextOutput('Yoga Signup endpoint is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ========== WAITLIST PROCESSOR ==========
// Call processAllWaitlists() on a time-driven trigger (e.g., every 10 min).
//
// SIMPLIFIED LOGIC (no holds, no party-size matching):
// 1. Count confirmed sign-ups for the class+date
// 2. Clean up waitlist: delete entries for people who already signed up
// 3. If spots are available, notify ALL remaining waitlisted people for that class
// 4. First-come-first-served: whoever signs up first gets the spot(s)
//
// Waitlist statuses:
//   (blank)     = waiting, not yet notified
//   Notified    = has been told a spot is open (no time limit)

function processWaitlistForClass(className, classDate) {

  // 1. Count confirmed registrations from the Sign-Ups sheet
  var signupSS = getOrCreateSpreadsheet();
  var signupSheet = getOrCreateSheet(signupSS);
  var signupLastRow = signupSheet.getLastRow();
  var confirmedCount = 0;
  var signupData = [];

  if (signupLastRow > 1) {
    signupData = signupSheet.getRange(2, 1, signupLastRow - 1, 11).getValues();
    var targetMD = extractMonthDay_(classDate);
    for (var i = 0; i < signupData.length; i++) {
      var rowClass = (signupData[i][4] || '').toString().trim();
      var rowDate  = (signupData[i][5] || '').toString().trim();
      var dateMatch = (rowDate === classDate) || (extractMonthDay_(rowDate) === targetMD);
      if (rowClass === className && dateMatch) {
        confirmedCount++;
        var guestName = (signupData[i][8] || '').toString().trim();
        if (guestName) confirmedCount++;
      }
    }
  }

  // 2. Read the waitlist
  var waitlistSS = getOrCreateWaitlistSpreadsheet();
  var waitlistSheet = getOrCreateWaitlistSheet(waitlistSS);
  var waitlistLastRow = waitlistSheet.getLastRow();

  if (waitlistLastRow <= 1) {
    Logger.log('Waitlist is empty for ' + className + ' on ' + classDate);
    return { notified: 0 };
  }

  var headers = waitlistSheet.getRange(1, 1, 1, waitlistSheet.getLastColumn()).getValues()[0];
  var statusColIndex = headers.indexOf('Status');

  if (statusColIndex === -1) {
    var newCol = waitlistSheet.getLastColumn() + 1;
    waitlistSheet.getRange(1, newCol).setValue('Status').setFontWeight('bold');
    statusColIndex = newCol - 1;
  }

  var numCols = waitlistSheet.getLastColumn();
  var waitlistData = waitlistSheet.getRange(2, 1, waitlistLastRow - 1, numCols).getValues();

  // 3. Clean up: delete waitlist entries for people who already signed up
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
        Logger.log('Removing from waitlist (already signed up): ' + waitlistData[e][1] + ' ' + waitlistData[e][2]);
        break;
      }
    }
  }

  // Delete bottom-up (protect last non-frozen row)
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

  // 4. Check if spots are available (per-class capacity from the schedule;
  //    blank capacity = unlimited, so there's always room).
  var capClass = getScheduleClassByLabel_(className);
  var capacity = capClass ? capClass.capacity : null;
  var availableSpots = (capacity == null) ? Number.MAX_SAFE_INTEGER : (capacity - confirmedCount);
  if (availableSpots <= 0) {
    Logger.log('No spots for ' + className + ' on ' + classDate + ' (confirmed: ' + confirmedCount + ')');
    return { notified: 0 };
  }

  // 5. Re-read waitlist after cleanup and notify ALL remaining people for this class
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
    if (wStatus === 'Notified') continue; // Already notified, skip

    // Notify this person
    var rowNum = w + 2;
    waitlistSheet.getRange(rowNum, statusColIndex + 1).setValue('Notified');

    var firstName = (waitlistData[w][1] || '').toString().trim();
    var lastName  = (waitlistData[w][2] || '').toString().trim();
    var email     = (waitlistData[w][3] || '').toString().trim();

    // Send notification email
    try {
      var waitlistHtml = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
          '<div style="background:#f5f0e8;padding:24px;text-align:center;border-radius:8px 8px 0 0;">' +
            '<h1 style="margin:0;font-family:Georgia,serif;font-size:24px;">' +
              '<a href="' + SITE_URL + '" style="color:#5B7553;text-decoration:none;">yogawithjessica.com</a>' +
            '</h1>' +
          '</div>' +
          '<div style="padding:24px;background:#fff;border:1px solid #e8e4dc;border-top:none;">' +
            '<p style="font-size:15px;">Hi ' + escHtml(firstName) + ',</p>' +
            '<p style="font-size:15px;line-height:1.6;">A spot has opened up for <strong>' +
              escHtml(wClass) + '</strong> on <strong>' + escHtml(wDate) + '</strong>.</p>' +
            '<p style="font-size:15px;line-height:1.6;">Spots are first come, first served, so sign up soon before it fills up again.</p>' +
            '<div style="text-align:center;margin:24px 0;">' +
              '<a href="' + SITE_URL + '/schedule.html" style="display:inline-block;background:#5B7553;color:#fff;padding:12px 32px;border-radius:6px;text-decoration:none;font-size:15px;font-weight:600;">Sign Up Now</a>' +
            '</div>' +
          '</div>' +
          '<div style="padding:16px;text-align:center;font-size:12px;color:#999;background:#f5f0e8;border-radius:0 0 8px 8px;">' +
            '<p style="margin:0;">Yoga with Jessica &mdash; Playa Del Rey, CA</p>' +
          '</div>' +
        '</div>';
      MailApp.sendEmail({
        to: email,
        subject: 'Yoga with Jessica \u2014 A Spot Opened Up',
        body: stripHtml(waitlistHtml),
        htmlBody: waitlistHtml,
        name: 'Yoga with Jessica',
        replyTo: ADMIN_EMAIL
      });
    } catch (mailErr) {
      Logger.log('Failed to email ' + email + ': ' + mailErr.toString());
    }

    notifiedCount++;
    Logger.log('Notified: ' + firstName + ' ' + lastName + ' (' + email + ')');
  }

  Logger.log('Waitlist processing done for ' + className + ' on ' + classDate +
             '. Available: ' + availableSpots + ', Notified: ' + notifiedCount);

  return { notified: notifiedCount, available: availableSpots };
}

// Process waitlist for ALL upcoming in-person classes.
// Set up as a time-driven trigger (e.g., every 10 minutes).
function processAllWaitlists() {
  var waitlistSS;
  try {
    waitlistSS = getOrCreateWaitlistSpreadsheet();
  } catch (e) {
    Logger.log('No waitlist spreadsheet found. Nothing to process.');
    return;
  }

  var waitlistSheet = getOrCreateWaitlistSheet(waitlistSS);
  var lastRow = waitlistSheet.getLastRow();
  if (lastRow <= 1) return;

  var headers = waitlistSheet.getRange(1, 1, 1, waitlistSheet.getLastColumn()).getValues()[0];
  var statusColIndex = headers.indexOf('Status');
  var numCols = waitlistSheet.getLastColumn();
  var data = waitlistSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  // Get unique class+date combinations from the waitlist (skip already-notified)
  var classDatePairs = {};
  for (var i = 0; i < data.length; i++) {
    var status = statusColIndex >= 0 ? (data[i][statusColIndex] || '').toString().trim() : '';
    // Process classes that have at least one un-notified entry
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

// ========== LATE SIGN-UP ZOOM LINK ==========
// Called at sign-up time. For each online class in the sign-up, checks if class
// starts soon (within LATE_SIGNUP_WINDOW_MIN). If so, ensures a Zoom meeting
// exists (creating one if needed) and returns the join URL for the confirmation
// email.
//
// This window must be >= the upper bound of sendMeetInvites' send window (35
// min). Otherwise a student who signs up after the periodic trigger has already
// emailed the link (and set the meet_sent_ flag) but before the old 30-min
// cutoff falls into a dead zone and never receives a link. Use 40 for margin.
var LATE_SIGNUP_WINDOW_MIN = 40;

function checkAndCreateMeetForLateSignup(rows) {
  if (!rows || rows.length === 0) return '';

  var now = new Date();
  var pstNow = new Date(now.toLocaleString('en-US', { timeZone: MEET_TZ }));
  var currentTotalMin = pstNow.getHours() * 60 + pstNow.getMinutes();
  var cache = PropertiesService.getScriptProperties();
  var meetLink = '';

  var occurrences = getOccurrencesOnPacificDate_(pstNow);

  for (var c = 0; c < occurrences.length; c++) {
    var cls = occurrences[c];

    // Only care about online classes happening today
    if (cls.type !== 'online') continue;

    var minutesUntilClass = (cls.startH * 60 + cls.startM) - currentTotalMin;

    // Only trigger for sign-ups close to class start
    // (and not after class has already started by more than 15 min)
    if (minutesUntilClass > LATE_SIGNUP_WINDOW_MIN || minutesUntilClass < -15) continue;

    // Check if this student actually signed up for this online class today
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

    Logger.log('Late sign-up for ' + cls.label + ' (' + minutesUntilClass + ' min away) — checking Zoom meeting');

    var linkKey = 'meet_link_' + cls.id + '_' + classDate;
    var existingLink = cache.getProperty(linkKey);

    if (existingLink) {
      // Zoom meeting already exists — return the cached link
      Logger.log('Zoom meeting exists, returning link for late sign-up');
      meetLink = existingLink;
    } else {
      // No Zoom meeting yet — create one now
      Logger.log('No Zoom meeting exists yet — creating for late sign-up');
      try {
        var result = createZoomMeeting(cls, pstNow, cls.durationMins);
        if (result.joinUrl) {
          meetLink = result.joinUrl;
          // Cache the link + event so sendMeetInvites reuses this meeting instead
          // of creating a duplicate. Do NOT set meet_sent_ here: that flag means
          // "the bulk email went out to all registered students," which this path
          // does not do (it only puts the link in this one confirmation email).
          // Setting it would make sendMeetInvites skip and silently drop the link
          // for everyone else registered for the class.
          cache.setProperty(linkKey, result.joinUrl);
          cache.setProperty('meet_event_' + cls.id + '_' + classDate, result.meetingId);
          Logger.log('Created Zoom meeting for late sign-up: ' + meetLink);
        }
      } catch (createErr) {
        Logger.log('Error creating Zoom meeting for late sign-up: ' + createErr.toString());
      }
    }
  }

  return meetLink;
}

// Get a Zoom API access token via Server-to-Server OAuth
function getZoomAccessToken() {
  var props = PropertiesService.getScriptProperties();
  var accountId    = props.getProperty('ZOOM_ACCOUNT_ID');
  var clientId     = props.getProperty('ZOOM_CLIENT_ID');
  var clientSecret = props.getProperty('ZOOM_CLIENT_SECRET');

  var credentials = Utilities.base64Encode(clientId + ':' + clientSecret);
  var response = UrlFetchApp.fetch(
    'https://zoom.us/oauth/token?grant_type=account_credentials&account_id=' + accountId,
    {
      method: 'POST',
      headers: {
        'Authorization': 'Basic ' + credentials,
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      muteHttpExceptions: true
    }
  );
  var data = JSON.parse(response.getContentText());
  if (!data.access_token) {
    throw new Error('Zoom token error: ' + response.getContentText());
  }
  return data.access_token;
}

// Create a scheduled Zoom meeting and return { joinUrl, meetingId }
function createZoomMeeting(cls, dateRef, durationMins) {
  var token = getZoomAccessToken();

  var startTime = new Date(dateRef);
  startTime.setHours(cls.startH, cls.startM, 0, 0);
  var startStr = Utilities.formatDate(startTime, MEET_TZ, "yyyy-MM-dd'T'HH:mm:ss");

  var payload = {
    topic: 'Yoga with Jessica \u2014 ' + (cls.label || cls.name),
    type: 2,
    start_time: startStr,
    duration: durationMins || 75,
    timezone: MEET_TZ,
    settings: {
      join_before_host: false,
      waiting_room: true,
      host_video: true,
      participant_video: true
    }
  };

  var response = UrlFetchApp.fetch('https://api.zoom.us/v2/users/me/meetings', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var meeting = JSON.parse(response.getContentText());
  if (!meeting.join_url) {
    throw new Error('Zoom meeting creation failed: ' + response.getContentText());
  }
  Logger.log('Created Zoom meeting: ' + meeting.join_url);
  return { joinUrl: meeting.join_url, meetingId: String(meeting.id) };
}

// Email the Zoom join link to all registered students for a class
function sendZoomLinkToStudents(students, cls, zoomLink) {
  var subject = 'Your Zoom link for today\'s Yoga with Jessica class';
  var hhmm = pad2_(cls.startH) + ':' + pad2_(cls.startM);
  for (var i = 0; i < students.length; i++) {
    try {
      var s = students[i];
      var timeStr = localTimeLine_(hhmm, cls.classDate, s.timezone);
      var timeHtml = timeStr
        ? '<p style="font-size:14px;color:#555;line-height:1.6;margin:0 0 12px;">Class time: <strong>' + escHtml(timeStr) + '</strong></p>'
        : '';
      var body =
        '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
        '<div style="background:#f5f0e8;padding:24px;text-align:center;border-radius:8px 8px 0 0;">' +
          '<h1 style="margin:0;font-family:Georgia,serif;font-size:24px;">' +
            '<a href="' + SITE_URL + '" style="color:#5B7553;text-decoration:none;">yogawithjessica.com</a>' +
          '</h1>' +
          '<p style="margin:6px 0 0;color:#888;font-size:13px;">Class Starting Soon</p>' +
        '</div>' +
        '<div style="padding:24px;background:#fff;border:1px solid #e8e4dc;border-top:none;">' +
          '<p style="font-size:15px;">Hi there,</p>' +
          '<p style="font-size:15px;line-height:1.6;">Your <strong>' + escHtml(cls.label || cls.name) + '</strong> class starts in about 30 minutes. Here\'s your Zoom link:</p>' +
          timeHtml +
          '<div style="background:#e8f5e9;padding:16px;border-radius:6px;margin:16px 0;font-size:14px;border-left:4px solid #5B7553;">' +
            '<strong>&#x1F4F9; Your Zoom link is ready</strong>' +
            '<div style="text-align:center;margin:12px 0;">' +
              '<a href="' + zoomLink + '" style="display:inline-block;background:#5B7553;color:#fff;padding:12px 32px;border-radius:6px;text-decoration:none;font-size:15px;font-weight:600;">Join Zoom</a>' +
            '</div>' +
            '<p style="margin:8px 0 0;color:#555;">Please have your camera on with good lighting. Microphones will be muted to minimize noise.</p>' +
          '</div>' +
          '<p style="font-size:14px;color:#555;">See you soon! &mdash; Jessica</p>' +
        '</div>' +
        '<div style="padding:16px;text-align:center;font-size:12px;color:#999;background:#f5f0e8;border-radius:0 0 8px 8px;">' +
          '<p style="margin:0;">Yoga with Jessica &mdash; Playa Del Rey, CA</p>' +
          '<p style="margin:4px 0 0;"><a href="' + SITE_URL + '" style="color:#5B7553;">yogawithjessica.com</a></p>' +
        '</div>' +
        '</div>';
      MailApp.sendEmail({
        to: s.email,
        subject: subject,
        body: stripHtml(body),
        htmlBody: body,
        name: 'Yoga with Jessica',
        replyTo: ADMIN_EMAIL
      });
      Logger.log('Sent Zoom link to: ' + s.email);
    } catch (mailErr) {
      Logger.log('Error sending Zoom email to ' + s.email + ': ' + mailErr.toString());
    }
  }
}

// ========== ZOOM INVITE AUTOMATION ==========
// Set up as a time-driven trigger (every 5 minutes).
// Checks if any class is starting within 30 minutes, then creates a
// Zoom meeting and emails the join link to all registered students.
//
// REQUIRES: Set ZOOM_ACCOUNT_ID, ZOOM_CLIENT_ID, ZOOM_CLIENT_SECRET
//           in Apps Script Project Settings → Script Properties
//
// Class schedule now lives in the "Schedule" Sheet tab (see getSchedule()).

var MEET_TZ = 'America/Los_Angeles';

function sendMeetInvites() {
  var now = new Date();

  // Convert "now" to PST to figure out what day/time it is in class timezone
  var pstNow = new Date(now.toLocaleString('en-US', { timeZone: MEET_TZ }));
  var currentTotalMin = pstNow.getHours() * 60 + pstNow.getMinutes();

  Logger.log('Zoom invite check at PST: ' + pstNow.toLocaleString());

  var cache = PropertiesService.getScriptProperties();
  var occurrences = getOccurrencesOnPacificDate_(pstNow);

  for (var c = 0; c < occurrences.length; c++) {
    var cls = occurrences[c];

    // Only process online classes
    if (cls.type !== 'online') continue;

    var minutesUntilClass = (cls.startH * 60 + cls.startM) - currentTotalMin;

    // Send invite when class is 25-35 minutes away (covers the 5-min trigger interval)
    if (minutesUntilClass < 25 || minutesUntilClass > 35) {
      Logger.log('Skipping ' + cls.label + ': ' + minutesUntilClass + ' min away (not in 25-35 min window)');
      continue;
    }

    Logger.log('Class ' + cls.label + ' starts in ' + minutesUntilClass + ' min — preparing Zoom invite');

    // Build the class date string to match what's in the spreadsheet
    var classDate = cls.classDate;

    // Check if we already sent an invite for this class+date (avoid duplicates)
    var sentKey = 'meet_sent_' + cls.id + '_' + classDate;
    if (cache.getProperty(sentKey)) {
      Logger.log('Already sent Zoom invite for ' + cls.label + ' on ' + classDate);
      continue;
    }

    // Get registered students for this class
    var students = getRegisteredStudents(cls.label, classDate);
    if (students.length === 0) {
      Logger.log('No students registered for ' + cls.label + ' on ' + classDate);
      continue;
    }

    Logger.log('Found ' + students.length + ' student(s) for ' + cls.label);

    // Create Zoom meeting and email the link to all registered students
    try {
      // Reuse a meeting already created by an earlier late sign-up (cached link)
      // so we never spin up a duplicate meeting; only create one if none exists.
      var linkKey = 'meet_link_' + cls.id + '_' + classDate;
      var joinUrl = cache.getProperty(linkKey);
      if (joinUrl) {
        Logger.log('Reusing existing Zoom meeting for ' + cls.label + ': ' + joinUrl);
      } else {
        var result = createZoomMeeting(cls, pstNow, cls.durationMins);
        joinUrl = result.joinUrl;
        cache.setProperty(linkKey, joinUrl);
        cache.setProperty('meet_event_' + cls.id + '_' + classDate, result.meetingId);
        Logger.log('Created Zoom meeting: ' + joinUrl + ' for ' + students.length + ' students');
      }

      // Mark the bulk email as sent so a later trigger fire doesn't re-send it.
      cache.setProperty(sentKey, new Date().toISOString());

      // Email all registered students the Zoom link
      sendZoomLinkToStudents(students, cls, joinUrl);

    } catch (zoomErr) {
      Logger.log('Error creating Zoom meeting for ' + cls.label + ': ' + zoomErr.toString());
    }
  }
}

// ========== MORNING-OF REMINDER EMAILS ==========
// Daily time-driven trigger (~7–8 AM PT; see setupTriggers). For each class
// happening today (Pacific), email registered students who signed up BEFORE
// today — a same-day signup already knows. Deduped per class+date via a Script
// Property so a re-fire inside the trigger's hour window can't double-send.
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

// Get all registered student emails for a class name + date
function getRegisteredStudents(className, classDate) {
  var out = [];
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return out;

    // Read through the Timezone column (18) when present so the Zoom email can
    // show the start time in each student's local zone.
    var nCols = Math.min(18, sheet.getLastColumn() || 11);
    var data = sheet.getRange(2, 1, lastRow - 1, nCols).getValues();
    var seen = {};

    for (var i = 0; i < data.length; i++) {
      var rowClass = (data[i][4] || '').toString().trim();
      var rowDate  = (data[i][5] || '').toString().trim();
      var rowEmail = (data[i][3] || '').toString().trim().toLowerCase();

      // Match by class date — use tolerant comparison to handle both old
      // ("Sunday, April 13") and new ("Sunday, April 13, 2026") stored formats.
      var dateMatch = (rowDate === classDate) ||
                      (extractMonthDay_(rowDate) === extractMonthDay_(classDate));
      if (dateMatch && rowEmail && !seen[rowEmail]) {
        // Check the class type matches — only send Meet invites for online classes
        var rowType = (data[i][6] || '').toString().trim().toLowerCase();
        if (rowType === 'online') {
          out.push({
            email: rowEmail,
            firstName: (data[i][1] || '').toString().trim(),
            timezone: (data[i][17] || '').toString().trim()
          });
          seen[rowEmail] = true;
        }
      }
    }
  } catch (err) {
    Logger.log('Error getting registered students: ' + err.toString());
  }
  return out;
}

// Format a date to match the spreadsheet format.
// Returns "Sunday, April 13, 2026" — matches what signup.html stores.
function formatClassDate(date) {
  return Utilities.formatDate(date, MEET_TZ, "EEEE, MMMM d, yyyy");
}

// Extract a canonical "MonthName Day" string from any of the date formats
// we have stored: "Sunday, April 13, 2026", "Sunday, April 13", "Sun, Apr 13".
// Used for tolerant matching in getRegisteredStudents.
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

// ========== HELPERS ==========
function getOrCreateSpreadsheet() {
  var files = DriveApp.getFilesByName('Yoga Signup');
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  return SpreadsheetApp.create('Yoga Signup');
}

function getOrCreateWaitlistSpreadsheet() {
  var files = DriveApp.getFilesByName('Yoga Waitlist');
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  return SpreadsheetApp.create('Yoga Waitlist');
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

// ========== AUTO-ARCHIVE PAST SIGN-UPS ==========
// Run this on a time-driven trigger every 10 minutes.
// Moves rows to Archive / Waitlist Archive once the class cutoff has passed.

var ARCHIVE_TZ = 'America/Los_Angeles';

function archivePastSignups() {
  var startTimes = buildStartTimeMap_();
  archiveSheet_('Yoga Signup', 'Sign-Ups', 'Archive', 6, startTimes);   // Class Date in col 6
  archiveSheet_('Yoga Waitlist', 'Waitlist', 'Waitlist Archive', 6, startTimes);
}

// Map of class label -> { startH, startM } from the schedule, used to compute
// each row's sign-up cutoff. Falls back to 6:00 PM for unknown/renamed classes.
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
  if (data.length <= 1) return; // header only

  // Get or create the archive sheet
  var archive = ss.getSheetByName(archiveName);
  if (!archive) {
    archive = ss.insertSheet(archiveName);
    archive.appendRow(data[0]); // copy headers
    archive.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
    archive.setFrozenRows(1);
  }

  // Current time in Pacific
  var now = new Date();
  var pstNow = new Date(now.toLocaleString('en-US', { timeZone: ARCHIVE_TZ }));

  // Check rows bottom-to-top so deletions don't shift indices
  var rowsToArchive = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var classDateStr = data[i][dateCol - 1]; // 0-indexed
    var className = data[i][dateCol - 2] || ''; // Class name column (col 5)

    if (!classDateStr) continue;

    // Parse the class date string robustly.
    // Stored as "Sunday, April 13, 2026" (new format) or "Sunday, April 13" (old, no year).
    // new Date() chokes on the weekday prefix, so strip it first.
    var classDate;
    if (classDateStr instanceof Date) {
      // Sheets auto-converted the cell to a Date object — use it directly
      classDate = classDateStr;
    } else {
      var s = String(classDateStr).trim();
      // Strip leading weekday (e.g. "Sunday, " or "Mon, ")
      s = s.replace(/^[A-Za-z]+,\s*/, '');
      // s is now "April 13, 2026" or "April 13"
      if (!/\d{4}/.test(s)) {
        // No year stored — infer it: use current year, but if that date is
        // already more than a week in the past, try next year instead
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

    // Determine start time from the schedule (match the row's class label)
    var startH = 18, startM = 0; // default 6:00 PM
    var st = startTimes[(className || '').toString()];
    if (st) { startH = st.startH; startM = st.startM; }

    // Build cutoff: class date + start time + 15 minutes (sign-up cutoff)
    var cutoff = new Date(classDate);
    cutoff.setHours(startH, startM + 15, 0, 0);

    // If cutoff has passed, archive this row
    if (pstNow > cutoff) {
      rowsToArchive.push({ index: i, row: data[i] });
    }
  }

  if (rowsToArchive.length === 0) return;

  // BATCH WRITE to archive (much faster than appendRow in a loop)
  var archiveData = [];
  for (var j = 0; j < rowsToArchive.length; j++) {
    archiveData.push(rowsToArchive[j].row);
  }
  var archiveLastRow = archive.getLastRow();
  archive.getRange(archiveLastRow + 1, 1, archiveData.length, archiveData[0].length).setValues(archiveData);

  // BATCH DELETE from source (bottom-to-top, handling last-row protection)
  var totalDataRows = sheet.getLastRow() - 1; // exclude header
  for (var j = 0; j < rowsToArchive.length; j++) {
    if (totalDataRows <= 1) {
      // Can't delete the last non-frozen row — clear it instead
      sheet.getRange(rowsToArchive[j].index + 1, 1, 1, sheet.getLastColumn()).clearContent();
    } else {
      sheet.deleteRow(rowsToArchive[j].index + 1);
    }
    totalDataRows--;
  }
}

// ========== TRIGGER SETUP ==========
// Run this function ONCE manually (from Apps Script editor: Run → setupTriggers)
// to install all required time-driven triggers. Safe to re-run — deletes old
// copies of the same functions before creating new ones to avoid duplicates.
function setupTriggers() {
  var TRIGGER_FUNCTIONS = ['sendMeetInvites', 'processAllWaitlists', 'archivePastSignups', 'sendClassReminders'];

  // Delete any existing triggers for our functions
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (TRIGGER_FUNCTIONS.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // sendMeetInvites — every 5 minutes (Zoom link email 30 min before online class)
  ScriptApp.newTrigger('sendMeetInvites')
    .timeBased().everyMinutes(5).create();

  // processAllWaitlists — every 10 minutes (promote waitlist when spots open)
  ScriptApp.newTrigger('processAllWaitlists')
    .timeBased().everyMinutes(10).create();

  // archivePastSignups — every hour (move past classes to Archive tab)
  ScriptApp.newTrigger('archivePastSignups')
    .timeBased().everyHours(1).create();

  // sendClassReminders — daily, fires sometime in the 7–8 AM Pacific hour
  // (morning-of reminder to students who signed up before today)
  ScriptApp.newTrigger('sendClassReminders')
    .timeBased().atHour(7).everyDays(1).inTimezone(MEET_TZ).create();

  Logger.log('Triggers installed: sendMeetInvites (5 min), processAllWaitlists (10 min), archivePastSignups (1 hour), sendClassReminders (daily ~7 AM PT)');
}
