/**
 * ══════════════════════════════════════════════════════════════════
 *  Come With — Google Apps Script (Production)
 *  Handles all form submissions, user auth, and data retrieval.
 *  Master admin: berky@comewith.org
 * ══════════════════════════════════════════════════════════════════
 *
 *  GOOGLE SHEET TAB STRUCTURE:
 *
 *  Tab: "Bookings Intake"
 *  Columns: timestamp | firstName | lastName | email | phone | service | eventDate | instagram | message | servicesSelected | serviceNotes | status
 *
 *  Tab: "Agreements"
 *  Columns: timestamp | type | agreementNum | inquiryEmail | clientName | email | phone | instagram | services | eventType | eventDate | venueName | setStart | setEnd | genre | totalFee | depositOpt | depositAmt | payment | promoRights | clientSigName | clientDate | notes | status
 *
 *  Tab: "Rental Intake"
 *  Columns: timestamp | agreementNum | inquiryEmail | renterName | djName | email | phone | instagram | equipment | numDays | avAddon | pickupDate | pickupTime | returnDate | returnTime | intendedUse | rentalFee | totalFee | deposit | depositAmount | payment | lateFee | renterSigName | renterDate | notes | status
 *
 *  Tab: "Users"
 *  Columns: email | name | passwordHash | role | mustChangePassword | created | lastLogin
 *
 * ══════════════════════════════════════════════════════════════════
 */

var MASTER_EMAIL = 'berky@comewith.org';

// ═══════════════════════════════════════════════════════════════════
//  UTILITIES
// ═══════════════════════════════════════════════════════════════════

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * SHEET LAYOUT: Row 1 = title, Row 2 = subtitle, Row 3 = headers, Row 4+ = data.
 * All read functions use HEADER_ROW (3) as the header source.
 */
var HEADER_ROW = 3; // 1-based row number where column headers live
var DATA_START_ROW = 4; // 1-based row number where data begins

function sheetToArray(sheet) {
  return sheetToArrayWithLayout(sheet, HEADER_ROW, DATA_START_ROW);
}

/**
 * Generalized sheet reader. Lets callers specify a custom header row and
 * data start row for sheets that don't use the default layout.
 *
 * Example: the Income sheet uses rows 1-4 as title rows, row 5 as headers,
 * and row 6 as the first data row — call sheetToArrayWithLayout(sheet, 5, 6).
 */
function sheetToArrayWithLayout(sheet, headerRow, dataStartRow) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1 || lastRow < dataStartRow) return []; // no data rows yet

  // Read headers explicitly from headerRow
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  // Read data explicitly starting from dataStartRow — this guarantees
  // the header row is never returned as a data record.
  var numDataRows = lastRow - dataStartRow + 1;
  var dataRows = sheet.getRange(dataStartRow, 1, numDataRows, lastCol).getValues();

  // Build a normalized lowercase set of header values so we can defensively
  // skip any row that accidentally contains the headers (belt-and-suspenders).
  var headerFingerprint = headers.map(function(h) { return String(h).toLowerCase().trim(); }).join('|');

  var rows = [];
  for (var i = 0; i < dataRows.length; i++) {
    var rowFingerprint = dataRows[i].map(function(v) { return String(v).toLowerCase().trim(); }).join('|');
    if (rowFingerprint === headerFingerprint) continue; // skip an accidental duplicate of the header row

    var obj = {};
    var hasContent = false;
    for (var j = 0; j < headers.length; j++) {
      var h = String(headers[j]).trim();
      if (h) {
        obj[h] = dataRows[i][j];
        if (String(dataRows[i][j]).trim() !== '') hasContent = true;
      }
    }
    if (hasContent) rows.push(obj);
  }
  return rows;
}

/**
 * CORS NOTE: Google Apps Script web apps deployed as "Anyone" automatically
 * include Access-Control-Allow-Origin: * on responses. However, if the
 * script throws an uncaught error, Google returns an HTML error page
 * which the browser interprets as a CORS failure. The doGet function
 * below is wrapped in try/catch to ensure it ALWAYS returns valid JSON.
 *
 * ContentService does not support custom headers — CORS is handled by
 * the Apps Script infrastructure when deployed correctly:
 *   Deploy → Web app → Execute as: Me → Who has access: Anyone
 *
 * If you get CORS errors, redeploy as a NEW deployment (not edit existing).
 */
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Find a row index (1-based) by matching a column value.
 * Returns -1 if not found.
 */
function findRowByColumn(sheet, colIndex, value) {
  if (!sheet || !value) return -1;
  var data = sheet.getDataRange().getValues();
  var target = String(value).toLowerCase().trim();
  for (var i = DATA_START_ROW - 1; i < data.length; i++) {
    if (String(data[i][colIndex]).toLowerCase().trim() === target) {
      return i + 1; // 1-based row number
    }
  }
  return -1;
}

/**
 * Get column index (0-based) by header name.
 */
function getColIndex(sheet, headerName) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  var headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().trim() === headerName.toLowerCase().trim()) {
      return i;
    }
  }
  return -1;
}

/**
 * Ensure headers exist in HEADER_ROW. If that row is empty, write them.
 */
function ensureHeaders(sheet, headers) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers]);
    return;
  }
  var existing = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  var hasHeaders = existing.some(function(h) { return String(h).trim() !== ''; });
  if (!hasHeaders) {
    sheet.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  }
}

/**
 * Filter rows by email across a sheet.
 */
function filterByEmail(rows, email) {
  var target = String(email).toLowerCase().trim();
  return rows.filter(function(r) {
    for (var k in r) {
      if (k.toLowerCase() === 'email' && String(r[k]).toLowerCase().trim() === target) {
        return true;
      }
    }
    return false;
  });
}


// ═══════════════════════════════════════════════════════════════════
//  doGet — ALL READ OPERATIONS
// ═══════════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    return _doGetInner(e);
  } catch (err) {
    return jsonResponse({ error: 'Server error: ' + String(err) });
  }
}

function _doGetInner(e) {
  var action = (e && e.parameter && e.parameter.action) || '';

  // ── Get Inquiries ──────────────────────────────────────
  if (action === 'getInquiries') {
    var sheet = getSheet('Bookings Intake');
    return jsonResponse(sheetToArray(sheet));
  }

  // ── Get Agreements ─────────────────────────────────────
  if (action === 'getAgreements') {
    var sheet = getSheet('Agreements');
    return jsonResponse(sheetToArray(sheet));
  }

  // ── Get Rentals ────────────────────────────────────────
  if (action === 'getRentals') {
    var sheet = getSheet('Rental Intake');
    return jsonResponse(sheetToArray(sheet));
  }

  // ── Get Income ─────────────────────────────────────────
  // Income sheet layout: rows 1-4 are title rows, row 5 is headers,
  // data starts at row 6. Use the generalized reader with custom offsets.
  if (action === 'getIncome') {
    var sheet = getSheet('Income');
    return jsonResponse(sheetToArrayWithLayout(sheet, 5, 6));
  }

  // ── Get Agreement Links ────────────────────────────────
  if (action === 'getAgreementLinks') {
    var sheet = getSheet('Agreement Links');
    // Strip the prefill blob before returning — it can be large and isn't
    // needed in the dashboard list view.
    var rows = sheetToArray(sheet).map(function(r) {
      var copy = {};
      for (var k in r) if (k !== 'prefill') copy[k] = r[k];
      return copy;
    });
    return jsonResponse(rows);
  }

  // ── Get Agreement Data by ID (for client pre-fill) ─────
  if (action === 'getAgreementData') {
    var id = (e.parameter.id || '').trim();
    if (!id) return jsonResponse({ error: 'Missing id parameter.' });
    var sheet = getSheet('Agreement Links');
    var idCol = getColIndex(sheet, 'agreementId');
    if (idCol === -1) return jsonResponse({ error: 'Agreement Links sheet missing agreementId column.' });
    var rowNum = findRowByColumn(sheet, idCol, id);
    if (rowNum === -1) return jsonResponse({ error: 'Agreement not found: ' + id });
    var lastCol = sheet.getLastColumn();
    var rowData = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    var obj = {};
    for (var i = 0; i < headers.length; i++) {
      obj[String(headers[i]).trim()] = rowData[i];
    }
    return jsonResponse(obj);
  }

  // ── Update Agreement Status (GET convenience) ──────────
  // e.g. ?action=updateAgreementStatus&id=ESA-20260422-X7K2P1&status=Signed
  if (action === 'updateAgreementStatus') {
    var id = (e.parameter.id || '').trim();
    var newStatus = (e.parameter.status || '').trim();
    if (!id || !newStatus) return jsonResponse({ error: 'id and status required.' });
    var ok = updateAgreementLinkStatus(id, newStatus);
    return jsonResponse({ success: ok });
  }

  // ── Get All ────────────────────────────────────────────
  if (action === 'getAll') {
    return jsonResponse({
      inquiries: sheetToArray(getSheet('Bookings Intake')),
      agreements: sheetToArray(getSheet('Agreements')),
      rentals: sheetToArray(getSheet('Rental Intake')),
      income: sheetToArrayWithLayout(getSheet('Income'), 5, 6)
    });
  }

  // ── Get Customer Data (filtered by email) ──────────────
  if (action === 'getCustomerData') {
    var email = (e.parameter.email || '').toLowerCase().trim();
    return jsonResponse({
      inquiries: filterByEmail(sheetToArray(getSheet('Bookings Intake')), email),
      agreements: filterByEmail(sheetToArray(getSheet('Agreements')), email),
      rentals: filterByEmail(sheetToArray(getSheet('Rental Intake')), email)
    });
  }

  // ── Get Users ──────────────────────────────────────────
  if (action === 'getUsers') {
    var sheet = getSheet('Users');
    var rows = sheetToArray(sheet);
    // Strip passwordHash from response
    rows = rows.map(function(r) {
      return {
        email: r.email || '',
        name: r.name || '',
        role: r.role || 'customer',
        mustChangePassword: r.mustChangePassword,
        created: r.created || '',
        lastLogin: r.lastLogin || ''
      };
    });
    return jsonResponse(rows);
  }

  // ── Login ──────────────────────────────────────────────
  if (action === 'login') {
    var email = (e.parameter.email || '').toLowerCase().trim();
    var hash = (e.parameter.hash || '').trim();

    if (!email || !hash) {
      return jsonResponse({ success: false, error: 'Email and password are required.' });
    }

    var sheet = getSheet('Users');
    ensureHeaders(sheet, ['email', 'name', 'passwordHash', 'role', 'mustChangePassword', 'created', 'lastLogin']);

    var emailCol = getColIndex(sheet, 'email');
    var rowNum = findRowByColumn(sheet, emailCol, email);

    if (rowNum === -1) {
      return jsonResponse({ success: false, error: 'No account found for this email.' });
    }

    var rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];

    var user = {};
    for (var i = 0; i < headers.length; i++) {
      user[String(headers[i])] = rowData[i];
    }

    var storedHash = String(user.passwordHash || '').trim();

    if (storedHash === '' || storedHash !== hash) {
      return jsonResponse({ success: false, error: 'Invalid email or password.' });
    }

    // Return user object (without hash)
    var mcp = user.mustChangePassword;
    var mustChange = (mcp === true || mcp === 'true' || mcp === 'TRUE' || mcp === 1);

    return jsonResponse({
      success: true,
      user: {
        email: String(user.email || '').toLowerCase().trim(),
        name: user.name || '',
        role: user.role || 'customer',
        mustChangePassword: mustChange,
        created: user.created || '',
        lastLogin: user.lastLogin || ''
      }
    });
  }

  // ── Check User ─────────────────────────────────────────
  if (action === 'checkUser') {
    var email = (e.parameter.email || '').toLowerCase().trim();

    if (!email) {
      return jsonResponse({ exists: false });
    }

    var sheet = getSheet('Users');
    ensureHeaders(sheet, ['email', 'name', 'passwordHash', 'role', 'mustChangePassword', 'created', 'lastLogin']);

    var emailCol = getColIndex(sheet, 'email');
    var rowNum = findRowByColumn(sheet, emailCol, email);

    if (rowNum === -1) {
      return jsonResponse({ exists: false });
    }

    var rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];

    var user = {};
    for (var i = 0; i < headers.length; i++) {
      user[String(headers[i])] = rowData[i];
    }

    var storedHash = String(user.passwordHash || '').trim();
    var mcp = user.mustChangePassword;
    var mustSet = (storedHash === '' || mcp === true || mcp === 'true' || mcp === 'TRUE' || mcp === 1);

    return jsonResponse({
      exists: true,
      mustSetPassword: mustSet,
      role: user.role || 'customer'
    });
  }

  return jsonResponse({ error: 'Unknown action: ' + action });
} // end _doGetInner


// ═══════════════════════════════════════════════════════════════════
//  doPost — ALL WRITE OPERATIONS
// ═══════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    return _doPostInner(e);
  } catch (err) {
    return jsonResponse({ error: 'Server error: ' + String(err) });
  }
}

function _doPostInner(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'Invalid JSON payload.' });
  }

  var type = body.type || '';

  // ═════════════════════════════════════════════════════════
  //  BOOKING INQUIRY (no type field — from index.html)
  // ═════════════════════════════════════════════════════════
  if (!type && body.firstName !== undefined) {
    return handleBookingInquiry(body);
  }

  // ═════════════════════════════════════════════════════════
  //  SERVICES SELECTION
  // ═════════════════════════════════════════════════════════
  if (type === 'Services Selection') {
    return handleServicesSelection(body);
  }

  // ═════════════════════════════════════════════════════════
  //  EVENTS SERVICES AGREEMENT
  // ═════════════════════════════════════════════════════════
  if (type === 'Events Services Agreement') {
    return handleEventsAgreement(body);
  }

  // ═════════════════════════════════════════════════════════
  //  EQUIPMENT RENTAL AGREEMENT
  // ═════════════════════════════════════════════════════════
  if (type === 'Equipment Rental Agreement') {
    return handleEquipmentRental(body);
  }

  // ═════════════════════════════════════════════════════════
  //  UPDATE STATUS (from dashboard "Mark Complete")
  // ═════════════════════════════════════════════════════════
  if (type === 'updateStatus') {
    return handleUpdateStatus(body);
  }

  // ═════════════════════════════════════════════════════════
  //  AGREEMENT LINK (admin Send Agreement)
  // ═════════════════════════════════════════════════════════
  if (type === 'storeAgreementLink') {
    return handleStoreAgreementLink(body);
  }

  // ═════════════════════════════════════════════════════════
  //  USER MANAGEMENT
  // ═════════════════════════════════════════════════════════
  if (type === 'createUser')      return handleCreateUser(body);
  if (type === 'updatePassword')  return handleUpdatePassword(body);
  if (type === 'updateUser')      return handleUpdateUser(body);
  if (type === 'deleteUser')      return handleDeleteUser(body);
  if (type === 'updateLastLogin') return handleUpdateLastLogin(body);

  return jsonResponse({ error: 'Unknown type: ' + type });
}


// ═══════════════════════════════════════════════════════════════════
//  HANDLERS — FORMS
// ═══════════════════════════════════════════════════════════════════

function handleBookingInquiry(d) {
  var sheet = getSheet('Bookings Intake');
  // Kept legacy 'service' and 'eventDate' columns for backwards compatibility with prior data.
  // New simplified inquiry form sends: firstName, lastName, email, phone, instagram, message, (optional) servicesSelected.
  var headers = ['timestamp', 'firstName', 'lastName', 'email', 'phone', 'service', 'eventDate', 'instagram', 'message', 'servicesSelected', 'serviceNotes', 'status'];
  ensureHeaders(sheet, headers);

  sheet.appendRow([
    d.timestamp || new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }),
    d.firstName || '',
    d.lastName || '',
    d.email || '',
    d.phone || '',
    d.service || '',                       // legacy, usually empty now
    d.eventDate || '',                     // legacy, usually empty now
    d.instagram || '',
    d.message || '',
    d.servicesSelected || '',              // may be filled immediately or by Services Selection step
    d.serviceNotes || '',
    'New'
  ]);

  // Notify admin
  try {
    MailApp.sendEmail({
      to: MASTER_EMAIL,
      subject: 'New Inquiry — ' + (d.firstName || '') + ' ' + (d.lastName || ''),
      htmlBody: '<h2>New Booking Inquiry</h2>' +
        '<p><strong>Name:</strong> ' + (d.firstName || '') + ' ' + (d.lastName || '') + '</p>' +
        '<p><strong>Email:</strong> ' + (d.email || '') + '</p>' +
        '<p><strong>Phone:</strong> ' + (d.phone || '') + '</p>' +
        '<p><strong>Instagram:</strong> ' + (d.instagram || '') + '</p>' +
        '<p><strong>Message:</strong> ' + (d.message || '') + '</p>' +
        (d.servicesSelected ? '<p><strong>Services Interested:</strong> ' + d.servicesSelected + '</p>' : '')
    });
  } catch (emailErr) {
    Logger.log('Email error: ' + emailErr);
  }

  return jsonResponse({ success: true });
}


function handleServicesSelection(d) {
  var sheet = getSheet('Bookings Intake');
  var headers = ['timestamp', 'firstName', 'lastName', 'email', 'phone', 'service', 'eventDate', 'instagram', 'message', 'servicesSelected', 'serviceNotes', 'status'];
  ensureHeaders(sheet, headers);

  var email = (d.email || '').toLowerCase().trim();
  var emailCol = getColIndex(sheet, 'email');
  var svcCol = getColIndex(sheet, 'servicesSelected');
  var notesCol = getColIndex(sheet, 'serviceNotes');

  if (emailCol === -1 || svcCol === -1) {
    // Columns don't exist — append as new row
    sheet.appendRow([
      d.timestamp || '', '', '', d.email || '', '', '', '', '', '',
      d.servicesSelected || '', d.serviceNotes || '', 'New'
    ]);
    return jsonResponse({ success: true });
  }

  // Find the most recent row matching this email (search from bottom)
  var data = sheet.getDataRange().getValues();
  var targetRow = -1;
  for (var i = data.length - 1; i >= DATA_START_ROW - 1; i--) {
    if (String(data[i][emailCol]).toLowerCase().trim() === email) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    // No existing row — append new
    sheet.appendRow([
      d.timestamp || '', d.name || '', '', d.email || '', '', '', '', '', '',
      d.servicesSelected || '', d.serviceNotes || '', 'New'
    ]);
  } else {
    // Update existing row
    sheet.getRange(targetRow, svcCol + 1).setValue(d.servicesSelected || '');
    sheet.getRange(targetRow, notesCol + 1).setValue(d.serviceNotes || '');
  }

  return jsonResponse({ success: true });
}


function handleEventsAgreement(d) {
  var sheet = getSheet('Agreements');
  var headers = ['timestamp', 'type', 'agreementNum', 'inquiryEmail', 'clientName', 'email', 'phone', 'instagram', 'services', 'eventType', 'eventDate', 'venueName', 'setStart', 'setEnd', 'genre', 'totalFee', 'depositOpt', 'depositAmt', 'payment', 'promoRights', 'clientSigName', 'clientDate', 'notes', 'status'];
  ensureHeaders(sheet, headers);

  sheet.appendRow([
    d.timestamp || new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }),
    d.type || 'Events Services Agreement',
    d.agreementNum || '',
    d.inquiryEmail || '',
    d.clientName || '',
    d.email || '',
    d.phone || '',
    d.instagram || '',
    d.services || '',
    d.eventType || '',
    d.eventDate || '',
    d.venueName || '',
    d.setStart || '',
    d.setEnd || '',
    d.genre || '',
    d.totalFee || '',
    d.depositOpt || '',
    d.depositAmt || '',
    d.payment || '',
    d.promoRights || '',
    d.clientSigName || '',
    d.clientDate || '',
    d.notes || '',
    'Submitted'
  ]);

  // Also update Bookings Intake status if inquiryEmail exists
  if (d.inquiryEmail || d.email) {
    updateInquiryStatus(d.inquiryEmail || d.email, 'Agreement Sent');
  }

  // If this submission came from a ?id=... link, flip the link's status to Signed
  if (d.agreementId) {
    updateAgreementLinkStatus(d.agreementId, 'Signed');
  }

  // Notify admin
  try {
    MailApp.sendEmail({
      to: MASTER_EMAIL,
      subject: 'Events Agreement Submitted — ' + (d.clientName || d.email || ''),
      htmlBody: '<h2>Events Services Agreement</h2>' +
        '<p><strong>Client:</strong> ' + (d.clientName || '') + '</p>' +
        '<p><strong>Email:</strong> ' + (d.email || '') + '</p>' +
        '<p><strong>Event:</strong> ' + (d.eventType || '') + ' on ' + (d.eventDate || '') + '</p>' +
        '<p><strong>Total:</strong> ' + (d.totalFee || '') + '</p>'
    });
  } catch (emailErr) {
    Logger.log('Email error: ' + emailErr);
  }

  return jsonResponse({ success: true });
}


function handleEquipmentRental(d) {
  var sheet = getSheet('Rental Intake');
  var headers = ['timestamp', 'agreementNum', 'inquiryEmail', 'renterName', 'djName', 'email', 'phone', 'instagram', 'equipment', 'numDays', 'avAddon', 'pickupDate', 'pickupTime', 'returnDate', 'returnTime', 'intendedUse', 'rentalFee', 'totalFee', 'deposit', 'depositAmount', 'payment', 'lateFee', 'renterSigName', 'renterDate', 'notes', 'status'];
  ensureHeaders(sheet, headers);

  sheet.appendRow([
    d.timestamp || new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }),
    d.agreementNum || '',
    d.inquiryEmail || '',
    d.renterName || '',
    d.djName || '',
    d.email || '',
    d.phone || '',
    d.instagram || '',
    d.equipment || '',
    d.numDays || '',
    d.avAddon || '',
    d.pickupDate || '',
    d.pickupTime || '',
    d.returnDate || '',
    d.returnTime || '',
    d.intendedUse || '',
    d.rentalFee || '',
    d.totalFee || '',
    d.deposit || '',
    d.depositAmount || '',
    d.payment || '',
    d.lateFee || '',
    d.renterSigName || '',
    d.renterDate || '',
    d.notes || '',
    'Submitted'
  ]);

  // Update Bookings Intake status if inquiryEmail exists
  if (d.inquiryEmail || d.email) {
    updateInquiryStatus(d.inquiryEmail || d.email, 'Rental Agreement Sent');
  }

  // If this submission came from a ?id=... link, flip the link's status to Signed
  if (d.agreementId) {
    updateAgreementLinkStatus(d.agreementId, 'Signed');
  }

  // Notify admin
  try {
    MailApp.sendEmail({
      to: MASTER_EMAIL,
      subject: 'Equipment Rental Agreement — ' + (d.renterName || d.email || ''),
      htmlBody: '<h2>Equipment Rental Agreement</h2>' +
        '<p><strong>Renter:</strong> ' + (d.renterName || '') + '</p>' +
        '<p><strong>Email:</strong> ' + (d.email || '') + '</p>' +
        '<p><strong>Equipment:</strong> ' + (d.equipment || '') + '</p>' +
        '<p><strong>Pickup:</strong> ' + (d.pickupDate || '') + ' at ' + (d.pickupTime || '') + '</p>' +
        '<p><strong>Return:</strong> ' + (d.returnDate || '') + ' at ' + (d.returnTime || '') + '</p>' +
        '<p><strong>Total:</strong> ' + (d.totalFee || '') + '</p>'
    });
  } catch (emailErr) {
    Logger.log('Email error: ' + emailErr);
  }

  // Send confirmation email to renter if flag is set
  if (d.sendConfirmationEmail && d.email) {
    try {
      MailApp.sendEmail({
        to: d.email,
        subject: 'Come With — Equipment Rental Agreement Confirmation',
        htmlBody: '<div style="font-family:Arial,sans-serif;max-width:600px;">' +
          '<h2 style="color:#1A1410;">Equipment Rental Agreement Confirmed</h2>' +
          '<p>Hi ' + (d.renterName || 'there') + ',</p>' +
          '<p>Your Equipment Rental Agreement has been submitted and received by Come With.</p>' +
          '<table style="width:100%;border-collapse:collapse;margin:16px 0;">' +
          '<tr><td style="padding:8px;border-bottom:1px solid #ddd;color:#888;font-size:12px;">AGREEMENT #</td><td style="padding:8px;border-bottom:1px solid #ddd;">' + (d.agreementNum || '—') + '</td></tr>' +
          '<tr><td style="padding:8px;border-bottom:1px solid #ddd;color:#888;font-size:12px;">EQUIPMENT</td><td style="padding:8px;border-bottom:1px solid #ddd;">' + (d.equipment || '—') + '</td></tr>' +
          '<tr><td style="padding:8px;border-bottom:1px solid #ddd;color:#888;font-size:12px;">RENTAL PERIOD</td><td style="padding:8px;border-bottom:1px solid #ddd;">' + (d.pickupDate || '') + ' — ' + (d.returnDate || '') + ' (' + (d.numDays || '1') + ' days)</td></tr>' +
          '<tr><td style="padding:8px;border-bottom:1px solid #ddd;color:#888;font-size:12px;">TOTAL</td><td style="padding:8px;border-bottom:1px solid #ddd;font-weight:bold;">' + (d.totalFee || '—') + '</td></tr>' +
          '<tr><td style="padding:8px;border-bottom:1px solid #ddd;color:#888;font-size:12px;">PAYMENT</td><td style="padding:8px;border-bottom:1px solid #ddd;">' + (d.payment || '—') + '</td></tr>' +
          '</table>' +
          '<p>If you have any questions, reply to this email or contact us at berky@comewith.org.</p>' +
          '<p style="color:#888;font-size:12px;margin-top:24px;">Come With · Brooklyn, NY · comewith.org</p>' +
          '</div>'
      });
    } catch (confirmErr) {
      Logger.log('Confirmation email error: ' + confirmErr);
    }
  }

  return jsonResponse({ success: true });
}


/**
 * Update the status column on the most recent matching inquiry.
 */
function updateInquiryStatus(email, newStatus) {
  var sheet = getSheet('Bookings Intake');
  var emailCol = getColIndex(sheet, 'email');
  var statusCol = getColIndex(sheet, 'status');

  if (emailCol === -1 || statusCol === -1) return;

  var data = sheet.getDataRange().getValues();
  var target = String(email).toLowerCase().trim();

  for (var i = data.length - 1; i >= DATA_START_ROW - 1; i--) {
    if (String(data[i][emailCol]).toLowerCase().trim() === target) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      return;
    }
  }
}


function handleUpdateStatus(d) {
  var email = (d.email || '').toLowerCase().trim();
  var status = d.status || 'Complete';

  if (!email) return jsonResponse({ error: 'Email required.' });

  updateInquiryStatus(email, status);
  return jsonResponse({ success: true });
}


// ═══════════════════════════════════════════════════════════════════
//  HANDLERS — USER MANAGEMENT
// ═══════════════════════════════════════════════════════════════════

function handleCreateUser(d) {
  var sheet = getSheet('Users');
  var headers = ['email', 'name', 'passwordHash', 'role', 'mustChangePassword', 'created', 'lastLogin'];
  ensureHeaders(sheet, headers);

  var email = (d.email || '').toLowerCase().trim();
  if (!email) return jsonResponse({ error: 'Email required.' });

  // Check if user already exists
  var emailCol = getColIndex(sheet, 'email');
  var existingRow = findRowByColumn(sheet, emailCol, email);

  if (existingRow !== -1) {
    // User exists — don't overwrite, just return success
    return jsonResponse({ success: true, note: 'User already exists.' });
  }

  sheet.appendRow([
    email,
    d.name || '',
    d.passwordHash || '',
    d.role || 'customer',
    d.mustChangePassword !== false ? 'true' : 'false',
    d.created || new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }),
    ''
  ]);

  return jsonResponse({ success: true });
}


function handleUpdatePassword(d) {
  var sheet = getSheet('Users');
  var headers = ['email', 'name', 'passwordHash', 'role', 'mustChangePassword', 'created', 'lastLogin'];
  ensureHeaders(sheet, headers);

  var email = (d.email || '').toLowerCase().trim();
  if (!email) return jsonResponse({ error: 'Email required.' });

  var emailCol = getColIndex(sheet, 'email');
  var hashCol = getColIndex(sheet, 'passwordHash');
  var mcpCol = getColIndex(sheet, 'mustChangePassword');
  var rowNum = findRowByColumn(sheet, emailCol, email);

  if (rowNum === -1) return jsonResponse({ error: 'User not found.' });

  if (hashCol !== -1) sheet.getRange(rowNum, hashCol + 1).setValue(d.newPasswordHash || '');
  if (mcpCol !== -1) sheet.getRange(rowNum, mcpCol + 1).setValue(d.mustChangePassword === false ? 'false' : 'true');

  return jsonResponse({ success: true });
}


function handleUpdateUser(d) {
  var sheet = getSheet('Users');
  var headers = ['email', 'name', 'passwordHash', 'role', 'mustChangePassword', 'created', 'lastLogin'];
  ensureHeaders(sheet, headers);

  var email = (d.email || '').toLowerCase().trim();
  if (!email) return jsonResponse({ error: 'Email required.' });

  var emailCol = getColIndex(sheet, 'email');
  var roleCol = getColIndex(sheet, 'role');
  var rowNum = findRowByColumn(sheet, emailCol, email);

  if (rowNum === -1) return jsonResponse({ error: 'User not found.' });
  if (roleCol !== -1) sheet.getRange(rowNum, roleCol + 1).setValue(d.role || 'customer');

  return jsonResponse({ success: true });
}


function handleDeleteUser(d) {
  var sheet = getSheet('Users');
  var email = (d.email || '').toLowerCase().trim();
  if (!email) return jsonResponse({ error: 'Email required.' });

  // Prevent deleting master admin
  if (email === MASTER_EMAIL.toLowerCase()) {
    return jsonResponse({ error: 'Cannot delete master admin.' });
  }

  var emailCol = getColIndex(sheet, 'email');
  var rowNum = findRowByColumn(sheet, emailCol, email);

  if (rowNum === -1) return jsonResponse({ error: 'User not found.' });

  sheet.deleteRow(rowNum);
  return jsonResponse({ success: true });
}


function handleUpdateLastLogin(d) {
  var sheet = getSheet('Users');
  var headers = ['email', 'name', 'passwordHash', 'role', 'mustChangePassword', 'created', 'lastLogin'];
  ensureHeaders(sheet, headers);

  var email = (d.email || '').toLowerCase().trim();
  if (!email) return jsonResponse({ success: true }); // Silent no-op

  var emailCol = getColIndex(sheet, 'email');
  var loginCol = getColIndex(sheet, 'lastLogin');
  var rowNum = findRowByColumn(sheet, emailCol, email);

  if (rowNum === -1) return jsonResponse({ success: true });
  if (loginCol !== -1) sheet.getRange(rowNum, loginCol + 1).setValue(d.lastLogin || '');

  return jsonResponse({ success: true });
}


// ═══════════════════════════════════════════════════════════════════
//  AGREEMENT LINKS — admin pre-generates a shareable, unique agreement URL
// ═══════════════════════════════════════════════════════════════════

/**
 * Agreement Links tab layout (default: headers row 3, data row 4+).
 * Columns: agreementId, type, clientName, clientEmail, clientPhone,
 *          clientInstagram, createdDate, link, prefill, status
 *
 * 'prefill' is a JSON string containing the fields the client form will
 * populate on load. 'status' is "Pending" until the client submits the
 * agreement, at which point updateAgreementLinkStatus flips it to "Signed".
 */
function handleStoreAgreementLink(d) {
  var sheet = getSheet('Agreement Links');
  var headers = ['agreementId', 'type', 'clientName', 'clientEmail',
                 'clientPhone', 'clientInstagram', 'createdDate', 'link',
                 'prefill', 'status'];
  ensureHeaders(sheet, headers);

  if (!d.agreementId) return jsonResponse({ error: 'agreementId required.' });

  // Don't double-write if an admin clicks Generate twice — return success.
  var idCol = getColIndex(sheet, 'agreementId');
  if (findRowByColumn(sheet, idCol, d.agreementId) !== -1) {
    return jsonResponse({ success: true, note: 'Agreement already exists.' });
  }

  sheet.appendRow([
    d.agreementId,
    d.agreementType || '',
    d.clientName || '',
    d.clientEmail || '',
    d.clientPhone || '',
    d.clientInstagram || '',
    d.createdDate || new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }),
    d.link || '',
    d.prefill || '',
    'Pending'
  ]);

  // Email the client from berky@comewith.org
  try {
    sendAgreementEmail(d);
  } catch (emailErr) {
    Logger.log('sendAgreementEmail error: ' + emailErr);
  }

  return jsonResponse({ success: true, agreementId: d.agreementId });
}

/**
 * Find the Agreement Links row for a given agreementId and set its status.
 * Returns true if a row was updated, false otherwise.
 */
function updateAgreementLinkStatus(agreementId, newStatus) {
  var sheet = getSheet('Agreement Links');
  var idCol = getColIndex(sheet, 'agreementId');
  var statusCol = getColIndex(sheet, 'status');
  if (idCol === -1 || statusCol === -1) return false;

  var rowNum = findRowByColumn(sheet, idCol, agreementId);
  if (rowNum === -1) return false;

  sheet.getRange(rowNum, statusCol + 1).setValue(newStatus);
  return true;
}

/**
 * Sends the agreement link email to the client from berky@comewith.org.
 *
 * Note on the From alias: MailApp doesn't support custom From headers
 * for non-Workspace accounts. GmailApp.sendEmail supports a `from` option
 * but only if the alias is configured on the sending account's Gmail
 * settings. We try GmailApp with the alias and fall back to MailApp
 * (which sends from the Apps Script owner's address — berky@comewith.org
 * if the project is owned by that account).
 */
function sendAgreementEmail(d) {
  var clientEmail = (d.clientEmail || '').trim();
  if (!clientEmail) return;

  var clientName = d.clientName || 'there';
  var agreementType = d.agreementType || 'Agreement';
  var link = d.link || '';
  var agreementId = d.agreementId || '';
  var fromAddress = 'berky@comewith.org';

  var subject = 'Your ' + agreementType + ' from Come With';

  var htmlBody =
    '<div style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;max-width:560px;color:#1A1410;line-height:1.6;">' +
      '<div style="border-bottom:2px solid #1A1410;padding-bottom:16px;margin-bottom:24px;">' +
        '<div style="font-family:\'Bebas Neue\',Helvetica,sans-serif;font-size:28px;letter-spacing:0.06em;color:#1A1410;">COME WITH</div>' +
        '<div style="font-size:11px;letter-spacing:0.18em;text-transform:uppercase;color:#8A7F72;margin-top:4px;">' + escapeHtml(agreementType) + '</div>' +
      '</div>' +

      '<p style="margin:0 0 14px;">Hi ' + escapeHtml(clientName) + ',</p>' +

      '<p style="margin:0 0 14px;">Thanks for working with us. Your ' + escapeHtml(agreementType) +
      ' is ready for your review and signature. Most fields are already filled in — you\'ll just need to look it over, add any final details, sign, and submit.</p>' +

      '<div style="background:#F2EDE6;border-left:3px solid #C13B2A;padding:16px 18px;margin:20px 0;">' +
        '<div style="font-size:10px;letter-spacing:0.14em;text-transform:uppercase;color:#8A7F72;margin-bottom:6px;">Agreement ID</div>' +
        '<div style="font-family:\'Courier New\',monospace;font-size:14px;color:#C13B2A;letter-spacing:0.02em;">' + escapeHtml(agreementId) + '</div>' +
      '</div>' +

      '<div style="margin:28px 0;">' +
        '<a href="' + escapeHtml(link) + '" style="display:inline-block;background:#C13B2A;color:#F2EDE6;font-family:\'Bebas Neue\',Helvetica,sans-serif;font-size:15px;letter-spacing:0.12em;text-decoration:none;padding:14px 28px;border-radius:0;">REVIEW &amp; SIGN &rarr;</a>' +
      '</div>' +

      '<p style="margin:0 0 10px;font-size:13px;color:#8A7F72;">Or copy &amp; paste this link into your browser:</p>' +
      '<p style="margin:0 0 24px;font-size:12px;word-break:break-all;color:#1A1410;"><a href="' + escapeHtml(link) + '" style="color:#C13B2A;">' + escapeHtml(link) + '</a></p>' +

      '<p style="margin:0 0 14px;">If anything looks off or you have questions, just reply to this email — I\'ll be right here.</p>' +

      '<p style="margin:0 0 6px;">— Berky</p>' +
      '<p style="margin:0 0 28px;font-size:12px;color:#8A7F72;">Come With · Brooklyn, NY</p>' +

      '<div style="border-top:1px solid rgba(26,20,16,0.12);padding-top:14px;font-size:11px;color:#8A7F72;line-height:1.5;">' +
        'This link is unique to you. Keep it private. If you didn\'t request this agreement, you can ignore this email.' +
      '</div>' +
    '</div>';

  var plainBody =
    'Hi ' + clientName + ',\n\n' +
    'Thanks for working with us. Your ' + agreementType + ' is ready for your review and signature. Most fields are already filled in — you\'ll just need to look it over, sign, and submit.\n\n' +
    'Agreement ID: ' + agreementId + '\n\n' +
    'Review & sign: ' + link + '\n\n' +
    'Reply to this email if you have any questions.\n\n' +
    '— Berky\n' +
    'Come With · Brooklyn, NY';

  // Try GmailApp with the alias; fall back to MailApp on error.
  try {
    GmailApp.sendEmail(clientEmail, subject, plainBody, {
      htmlBody: htmlBody,
      from: fromAddress,
      name: 'Berky — Come With'
    });
  } catch (gmailErr) {
    Logger.log('GmailApp alias failed, falling back to MailApp: ' + gmailErr);
    MailApp.sendEmail({
      to: clientEmail,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      name: 'Berky — Come With'
    });
  }
}

/**
 * Minimal HTML escaper for email template fields.
 */
function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
