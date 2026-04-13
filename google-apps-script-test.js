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

// ========== WRITE SIGN-UPS ==========

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
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
        row.zip || ''
      ]];
      var range = sheet.getRange(newRow, 1, 1, 17);
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

  // Build class list
  var classLines = '';
  var hasInPerson = false;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var icon = r.classType === 'In-Person' ? '&#x1F3E0;' : '&#x1F4BB;';
    classLines += '<tr>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;">' + icon + ' ' + escHtml(r.className) + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;">' + escHtml(r.classDate) + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #eee;">' + escHtml(r.classType) + '</td>' +
      '</tr>';
    if (r.classType === 'In-Person') hasInPerson = true;
  }

  var cancelUrl = SITE_URL + '/cancel.html?token=' + cancelToken;
  var subject = 'Yoga with Jessica — Sign-Up Confirmation';

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
          '<strong>&#x1F4F9; Your Zoom link is ready!</strong><br>' +
          '<p style="margin:8px 0;">Class is starting soon — join here:</p>' +
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
  var subject = 'Yoga with Jessica — A Spot Opened Up!';
  var body = '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
    '<div style="background:#f5f0e8;padding:24px;text-align:center;border-radius:8px 8px 0 0;">' +
      '<h1 style="margin:0;font-family:Georgia,serif;color:#5B7553;font-size:24px;">Yoga with Jessica</h1>' +
    '</div>' +
    '<div style="padding:24px;background:#fff;border:1px solid #e8e4dc;border-top:none;">' +
      '<p style="font-size:15px;">Hi ' + escHtml(firstName) + ',</p>' +
      '<p style="font-size:15px;line-height:1.6;">A spot has opened up for <strong>' +
        escHtml(className) + '</strong> on <strong>' + escHtml(classDate) + '</strong>!</p>' +
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
      classLines += '<li>' + icon + ' ' + escHtml(r.className) + ' — ' + escHtml(r.classDate) + ' (' + escHtml(r.classType) + ')</li>';
    }

    var subject = '\uD83E\uDDD8 New Sign-Up: ' + firstName + ' ' + lastName + ' — ' + (rows[0].className || '').split(' — ')[0];

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
      classLines += '<li>' + escHtml(c.className) + ' — ' + escHtml(c.classDate) + ' (' + escHtml(c.classType) + ')</li>';
    }

    var subject = '\u274C Cancellation: ' + studentName + ' — ' + (cancelledClasses[0] ? cancelledClasses[0].className : '').split(' — ')[0];

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

var TEST_MEET_TZ = 'America/Los_Angeles';

var TEST_CLASS_SCHEDULE = [
  { day: 0, name: 'Sunday Evening — Online via Zoom',           startH: 18, startM: 0, endH: 19, endM: 15, type: 'online' },
  { day: 2, name: 'Tuesday Evening — CCV Clubhouse (In Person)', startH: 18, startM: 0, endH: 19, endM: 15, type: 'in-person' },
  { day: 3, name: 'Wednesday Evening — Restorative Yoga (Online)', startH: 18, startM: 0, endH: 19, endM: 15, type: 'online' }
];

function checkAndCreateMeetForLateSignup(rows) {
  if (!rows || rows.length === 0) return '';

  var now = new Date();
  var pstNow = new Date(now.toLocaleString('en-US', { timeZone: TEST_MEET_TZ }));
  var currentDay = pstNow.getDay();
  var currentTotalMin = pstNow.getHours() * 60 + pstNow.getMinutes();
  var cache = PropertiesService.getScriptProperties();
  var meetLink = '';

  // Also check for test-forced Zoom state (set via ?action=set_meet_link)
  var forcedLink = cache.getProperty('test_force_meet_link');
  if (forcedLink) {
    Logger.log('Test: forced Zoom link found: ' + forcedLink);
    return forcedLink;
  }

  for (var c = 0; c < TEST_CLASS_SCHEDULE.length; c++) {
    var cls = TEST_CLASS_SCHEDULE[c];

    if (cls.type !== 'online' || cls.day !== currentDay) continue;

    var classStartMin = cls.startH * 60 + cls.startM;
    var minutesUntilClass = classStartMin - currentTotalMin;

    if (minutesUntilClass > 30 || minutesUntilClass < -15) continue;

    // Check if this student signed up for this class
    var signedUpForThis = false;
    for (var r = 0; r < rows.length; r++) {
      if (rows[r].classType === 'Online') {
        signedUpForThis = true;
        break;
      }
    }
    if (!signedUpForThis) continue;

    // Check for existing mock Zoom link
    var classDate = formatClassDate(pstNow);
    var linkKey = 'meet_link_' + cls.day + '_' + classDate;
    var existingLink = cache.getProperty(linkKey);

    if (existingLink) {
      meetLink = existingLink;
      Logger.log('Test: found existing mock Zoom link: ' + meetLink);
    } else {
      // Create a mock Zoom link
      meetLink = 'https://zoom.us/j/test-yoga-' + cls.day + '-' + Date.now();
      cache.setProperty(linkKey, meetLink);
      var sentKey = 'meet_sent_' + cls.day + '_' + classDate;
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
              duplicates.push(checkClass + ' — ' + checkDate);
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

    // Clean waitlist spreadsheet sheets
    var waitlistSS = getOrCreateWaitlistSpreadsheet();
    waitlistCounts.waitlist = clearSheetData_(waitlistSS, 'Waitlist');
    waitlistCounts.waitlistArchive = clearSheetData_(waitlistSS, 'Waitlist Archive');

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'ok',
        signups: signupCounts.signups,
        archive: signupCounts.archive,
        emailLog: signupCounts.emailLog,
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
          row.liabilityWaiver || 'YES — Accepted',
          row.guestFirstName || '',
          row.guestLastName || '',
          row.guestOf || '',
          row.cancelToken || '',
          row.device || '',
          row.browser || '',
          row.city || '',
          row.state || '',
          row.zip || ''
        ]];
        var range = sheet.getRange(newRow, 1, 1, 17);
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
var MAX_CAPACITY = 10;

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

  var availableSpots = MAX_CAPACITY - confirmedCount;
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

var CLASS_START_TIMES = {
  'Sunday':    { startH: 18, startM: 0 },
  'Tuesday':   { startH: 18, startM: 0 },
  'Wednesday': { startH: 18, startM: 0 }
};

function archivePastSignups() {
  archiveSheet_(TEST_SIGNUP_SS_NAME, 'Sign-Ups', 'Archive', 6);
  archiveSheet_(TEST_WAITLIST_SS_NAME, 'Waitlist', 'Waitlist Archive', 6);
}

function archiveSheet_(ssName, sheetName, archiveName, dateCol) {
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
    for (var key in CLASS_START_TIMES) {
      if (className.indexOf(key) !== -1) {
        startH = CLASS_START_TIMES[key].startH;
        startM = CLASS_START_TIMES[key].startM;
        break;
      }
    }

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
