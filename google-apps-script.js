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

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
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

  // Cancel link — points to cancel.html on the website
  var cancelUrl = SITE_URL + '/cancel.html?token=' + cancelToken;

  var subject = 'Yoga with Jessica — Sign-Up Confirmation';

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
    htmlBody: body
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

    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject,
      htmlBody: body
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

    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject,
      htmlBody: body
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
              duplicates.push(checkClass + ' — ' + checkDate);
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

var MAX_CAPACITY = 10;

function processWaitlistForClass(className, classDate) {

  // 1. Count confirmed registrations from the Sign-Ups sheet
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

  // 4. Check if spots are available
  var availableSpots = MAX_CAPACITY - confirmedCount;
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
      MailApp.sendEmail({
        to: email,
        subject: 'Yoga with Jessica — A Spot Opened Up!',
        htmlBody: '<div style="font-family:Calibri,Arial,sans-serif;max-width:600px;margin:0 auto;color:#333;">' +
          '<div style="background:#f5f0e8;padding:24px;text-align:center;border-radius:8px 8px 0 0;">' +
            '<h1 style="margin:0;font-family:Georgia,serif;font-size:24px;">' +
              '<a href="' + SITE_URL + '" style="color:#5B7553;text-decoration:none;">yogawithjessica.com</a>' +
            '</h1>' +
          '</div>' +
          '<div style="padding:24px;background:#fff;border:1px solid #e8e4dc;border-top:none;">' +
            '<p style="font-size:15px;">Hi ' + escHtml(firstName) + ',</p>' +
            '<p style="font-size:15px;line-height:1.6;">A spot has opened up for <strong>' +
              escHtml(wClass) + '</strong> on <strong>' + escHtml(wDate) + '</strong>!</p>' +
            '<p style="font-size:15px;line-height:1.6;">Spots are first come, first served, so sign up soon before it fills up again.</p>' +
            '<div style="text-align:center;margin:24px 0;">' +
              '<a href="' + SITE_URL + '/schedule.html" style="display:inline-block;background:#5B7553;color:#fff;padding:12px 32px;border-radius:6px;text-decoration:none;font-size:15px;font-weight:600;">Sign Up Now</a>' +
            '</div>' +
          '</div>' +
          '<div style="padding:16px;text-align:center;font-size:12px;color:#999;background:#f5f0e8;border-radius:0 0 8px 8px;">' +
            '<p style="margin:0;">Yoga with Jessica &mdash; Playa Del Rey, CA</p>' +
          '</div>' +
        '</div>'
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
// starts within 30 minutes. If so, ensures a Zoom meeting exists (creating one if
// needed) and returns the join URL for the confirmation email.

function checkAndCreateMeetForLateSignup(rows) {
  if (!rows || rows.length === 0) return '';

  var now = new Date();
  var pstNow = new Date(now.toLocaleString('en-US', { timeZone: MEET_TZ }));
  var currentDay = pstNow.getDay();
  var currentTotalMin = pstNow.getHours() * 60 + pstNow.getMinutes();
  var cache = PropertiesService.getScriptProperties();
  var meetLink = '';

  for (var c = 0; c < CLASS_SCHEDULE.length; c++) {
    var cls = CLASS_SCHEDULE[c];

    // Only care about online classes happening today
    if (cls.type !== 'online' || cls.day !== currentDay) continue;

    var classStartMin = cls.startH * 60 + cls.startM;
    var minutesUntilClass = classStartMin - currentTotalMin;

    // Only trigger for sign-ups within 30 minutes of class start
    // (and not after class has already started by more than 15 min)
    if (minutesUntilClass > 30 || minutesUntilClass < -15) continue;

    // Check if this student actually signed up for this class
    var classDate = formatClassDate(pstNow);
    var signedUpForThis = false;
    for (var r = 0; r < rows.length; r++) {
      if (rows[r].classType === 'Online' && rows[r].classDate &&
          rows[r].className && rows[r].className.indexOf(cls.name.split(' ')[0]) !== -1) {
        signedUpForThis = true;
        break;
      }
    }
    if (!signedUpForThis) continue;

    Logger.log('Late sign-up for ' + cls.name + ' (' + minutesUntilClass + ' min away) — checking Zoom meeting');

    var linkKey = 'meet_link_' + cls.day + '_' + classDate;
    var existingLink = cache.getProperty(linkKey);

    if (existingLink) {
      // Zoom meeting already exists — return the cached link
      Logger.log('Zoom meeting exists, returning link for late sign-up');
      meetLink = existingLink;
    } else {
      // No Zoom meeting yet — create one now
      Logger.log('No Zoom meeting exists yet — creating for late sign-up');
      try {
        var result = createZoomMeeting(cls, pstNow, 75);
        if (result.joinUrl) {
          meetLink = result.joinUrl;
          // Save for future late sign-ups and the regular trigger
          cache.setProperty(linkKey, result.joinUrl);
          cache.setProperty('meet_event_' + cls.day + '_' + classDate, result.meetingId);
          // Mark as sent so the regular trigger doesn't create a second meeting
          cache.setProperty('meet_sent_' + cls.day + '_' + classDate, new Date().toISOString());
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
    topic: 'Yoga with Jessica \u2014 ' + cls.name,
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
function sendZoomLinkToStudents(emails, cls, zoomLink) {
  var subject = 'Your Zoom link for today\'s Yoga with Jessica class';
  for (var i = 0; i < emails.length; i++) {
    try {
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
          '<p style="font-size:15px;line-height:1.6;">Your <strong>' + escHtml(cls.name) + '</strong> class starts in about 30 minutes. Here\'s your Zoom link:</p>' +
          '<div style="background:#e8f5e9;padding:16px;border-radius:6px;margin:16px 0;font-size:14px;border-left:4px solid #5B7553;">' +
            '<strong>&#x1F4F9; Your Zoom link is ready!</strong>' +
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
      MailApp.sendEmail({ to: emails[i], subject: subject, htmlBody: body });
      Logger.log('Sent Zoom link to: ' + emails[i]);
    } catch (mailErr) {
      Logger.log('Error sending Zoom email to ' + emails[i] + ': ' + mailErr.toString());
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
// Class schedule (PST / America/Los_Angeles):
//   Sunday    6:00 PM - 7:15 PM  Online
//   Tuesday   6:00 PM - 7:15 PM  In-Person (CCV)
//   Wednesday 6:00 PM - 7:15 PM  Online

var CLASS_SCHEDULE = [
  { day: 0, name: 'Sunday Evening — Online via Zoom',           startH: 18, startM: 0, endH: 19, endM: 15, type: 'online' },
  { day: 2, name: 'Tuesday Evening — CCV Clubhouse (In Person)', startH: 18, startM: 0, endH: 19, endM: 15, type: 'in-person' },
  { day: 3, name: 'Wednesday Evening — Restorative Yoga (Online)', startH: 18, startM: 0, endH: 19, endM: 15, type: 'online' }
];

var MEET_TZ = 'America/Los_Angeles';

function sendMeetInvites() {
  var now = new Date();

  // Convert "now" to PST to figure out what day/time it is in class timezone
  var pstNow = new Date(now.toLocaleString('en-US', { timeZone: MEET_TZ }));
  var currentDay = pstNow.getDay();       // 0=Sun, 1=Mon, ...
  var currentHour = pstNow.getHours();
  var currentMin = pstNow.getMinutes();
  var currentTotalMin = currentHour * 60 + currentMin;

  Logger.log('Zoom invite check at PST: ' + pstNow.toLocaleString() + ' (day=' + currentDay + ', time=' + currentHour + ':' + currentMin + ')');

  for (var c = 0; c < CLASS_SCHEDULE.length; c++) {
    var cls = CLASS_SCHEDULE[c];

    // Only process online classes today
    if (currentDay !== cls.day || cls.type !== 'online') continue;

    var classStartMin = cls.startH * 60 + cls.startM;
    var minutesUntilClass = classStartMin - currentTotalMin;

    // Send invite when class is 25-35 minutes away (covers the 5-min trigger interval)
    if (minutesUntilClass < 25 || minutesUntilClass > 35) {
      Logger.log('Skipping ' + cls.name + ': ' + minutesUntilClass + ' min away (not in 25-35 min window)');
      continue;
    }

    Logger.log('Class ' + cls.name + ' starts in ' + minutesUntilClass + ' min — preparing Zoom invite');

    // Build the class date string to match what's in the spreadsheet (e.g., "Sun, Apr 6")
    var classDate = formatClassDate(pstNow);

    // Check if we already sent an invite for this class+date (avoid duplicates)
    var sentKey = 'meet_sent_' + cls.day + '_' + classDate;
    var cache = PropertiesService.getScriptProperties();
    if (cache.getProperty(sentKey)) {
      Logger.log('Already sent Zoom invite for ' + cls.name + ' on ' + classDate);
      continue;
    }

    // Get registered students for this class
    var students = getRegisteredStudents(cls.name, classDate);
    if (students.length === 0) {
      Logger.log('No students registered for ' + cls.name + ' on ' + classDate);
      continue;
    }

    Logger.log('Found ' + students.length + ' student(s) for ' + cls.name);

    // Create Zoom meeting and email the link to all registered students
    try {
      var result = createZoomMeeting(cls, pstNow, 75);

      Logger.log('Created Zoom meeting: ' + result.joinUrl + ' for ' + students.length + ' students');

      // Save Zoom link so late sign-ups can use it
      var linkKey = 'meet_link_' + cls.day + '_' + classDate;
      cache.setProperty(linkKey, result.joinUrl);
      cache.setProperty('meet_event_' + cls.day + '_' + classDate, result.meetingId);

      // Mark as sent to avoid duplicates
      cache.setProperty(sentKey, new Date().toISOString());

      // Email all registered students the Zoom link
      sendZoomLinkToStudents(students, cls, result.joinUrl);

    } catch (zoomErr) {
      Logger.log('Error creating Zoom meeting for ' + cls.name + ': ' + zoomErr.toString());
    }
  }
}

// Get all registered student emails for a class name + date
function getRegisteredStudents(className, classDate) {
  var emails = [];
  try {
    var ss = getOrCreateSpreadsheet();
    var sheet = getOrCreateSheet(ss);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return emails;

    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    var seen = {};

    for (var i = 0; i < data.length; i++) {
      var rowClass = (data[i][4] || '').toString().trim();
      var rowDate  = (data[i][5] || '').toString().trim();
      var rowEmail = (data[i][3] || '').toString().trim().toLowerCase();

      // Match by class date (the class name in the sheet may be abbreviated)
      if (rowDate === classDate && rowEmail && !seen[rowEmail]) {
        // Check the class type matches — only send Meet invites for online classes
        var rowType = (data[i][6] || '').toString().trim().toLowerCase();
        if (rowType === 'online') {
          emails.push(rowEmail);
          seen[rowEmail] = true;
        }
      }
    }
  } catch (err) {
    Logger.log('Error getting registered students: ' + err.toString());
  }
  return emails;
}

// Format a date to match the spreadsheet format (e.g., "Sun, Apr 6")
function formatClassDate(date) {
  var days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return days[date.getDay()] + ', ' + months[date.getMonth()] + ' ' + date.getDate();
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

// Class start times by keyword (used to determine cutoff)
var CLASS_START_TIMES = {
  'Sunday':    { startH: 18, startM: 0 },
  'Tuesday':   { startH: 18, startM: 0 },
  'Wednesday': { startH: 18, startM: 0 }
};

function archivePastSignups() {
  archiveSheet_('Yoga Signup', 'Sign-Ups', 'Archive', 6);   // Class Date in col 6
  archiveSheet_('Yoga Waitlist', 'Waitlist', 'Waitlist Archive', 6);
}

function archiveSheet_(ssName, sheetName, archiveName, dateCol) {
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

    // Determine start time from class name
    var startH = 18, startM = 0; // default
    for (var key in CLASS_START_TIMES) {
      if (className.indexOf(key) !== -1) {
        startH = CLASS_START_TIMES[key].startH;
        startM = CLASS_START_TIMES[key].startM;
        break;
      }
    }

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
