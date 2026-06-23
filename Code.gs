// ═══════════════════════════════════════════════════════════════
//  SECTOR 40 INCIDENT TRACKER — Apps Script Backend (v7)
// ═══════════════════════════════════════════════════════════════
//
//  SHEET TABS REQUIRED:
//  1. "User Credentials"      — columns: Phone | Name | IsAdmin | Email | BlockHouseNo | PIN
//  2. "Form Responses 1"      — auto-created by Google Forms link
//  3. "Pending Registrations" — auto-created by this script on first use
//
//  The Incidents data comes from Google Forms → "Form Responses 1"
//  Column order (as created by the Form):
//  A: Timestamp  B: refId  C: logDate  D: name  E: phone
//  F: category   G: description  H: geo  I: lat  J: lng
//  K: photo (Cloudinary URL)   L: Status (added manually or by script)
//
//  STATUS COLUMN: After linking the form to the sheet, manually
//  add a column header "Status" in column L of Form Responses 1.
//  New submissions will not have it — the getOpenIssues function
//  treats blank status as "Open".
//
//  SELF-REGISTRATION: residents register themselves via the app
//  (Name, Phone, Email, Block/House No, 4-digit PIN) and verify
//  with a 6-digit email OTP. On success, a new row is appended
//  automatically to User Credentials. Re-registration with an
//  already-registered phone is rejected.
//
//  LOGIN: requires Phone + PIN together (not phone alone), since
//  phone numbers are not secret. Wrong-PIN attempts are tracked
//  per-phone in "Login Lockouts" (auto-created) — after 5 wrong
//  attempts, that phone is locked for 60 seconds.
//
//  EXISTING RESIDENTS (registered before PIN existed): the admin
//  must manually add a PIN value to their row in User Credentials
//  for them to be able to log in. There is no auto-migration.
//
//  EMAIL DELIVERY: OTP emails are sent via the Brevo transactional
//  email API (not MailApp), sent from BREVO_SENDER_EMAIL below.
//  This requires:
//   1. A free Brevo account (brevo.com) with an API key
//   2. BREVO_SENDER_EMAIL verified as a sender in Brevo
//      (Senders & IP → Domains → verify noidasector40.in, or
//       Senders & IP → Senders → verify the single address)
//
// ═══════════════════════════════════════════════════════════════

const SHEET_CREDS    = 'User Credentials';
const SHEET_INCIDENTS = 'Form Responses 1';
const SHEET_PENDING   = 'Pending Registrations';
const SHEET_LOCKOUTS  = 'Login Lockouts';
const OTP_EXPIRY_MINUTES = 10;
const MAX_LOGIN_ATTEMPTS = 5;
const LOCKOUT_SECONDS     = 60;

// ── Brevo (transactional email) configuration ──
const BREVO_API_KEY    = 'YOUR_BREVO_API_KEY_HERE'; // Set this in the Apps Script editor — never commit the real key
const BREVO_SENDER_EMAIL = 'admin@noidasector40.in'; // must be verified in Brevo before this works
const BREVO_SENDER_NAME  = 'Sector 40 Community';

const COL_REFID      = 2;   // B
const COL_LOGDATE    = 3;   // C
const COL_NAME       = 4;   // D
const COL_PHONE      = 5;   // E
const COL_CATEGORY   = 6;   // F
const COL_DESC       = 7;   // G
const COL_GEO        = 8;   // H
const COL_LAT        = 9;   // I
const COL_LNG        = 10;  // J
const COL_PHOTO      = 11;  // K
const COL_STATUS     = 12;  // L

// ── Entry point ───────────────────────────────────────────────

function doGet(e) {
  const action   = e.parameter.action   || '';
  const callback = e.parameter.callback || '';

  let result;
  try {
    if      (action === 'auth')             result = handleAuth(e);
    else if (action === 'getOpenIssues')    result = handleGetOpenIssues();
    else if (action === 'closeIssue')       result = handleCloseIssue(e);
    else if (action === 'requestOtp')       result = handleRequestOtp(e);
    else if (action === 'verifyOtp')        result = handleVerifyOtp(e);
    else result = { success: false, error: 'Unknown action' };
  } catch(err) {
    result = { success: false, error: err.message };
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Auth (Phone + PIN, with lockout protection) ────────────────

function handleAuth(e) {
  const phone = (e.parameter.phone || '').trim();
  const pin   = (e.parameter.pin   || '').trim();
  if (!phone) return { success: false, error: 'No phone provided' };
  if (!pin)   return { success: false, error: 'PIN is required' };

  // Check lockout before anything else
  const lockoutCheck = checkLockout(phone);
  if (lockoutCheck.locked) {
    return { success: false, error: 'Too many incorrect attempts. Please try again in ' + lockoutCheck.secondsLeft + ' seconds.', locked: true, secondsLeft: lockoutCheck.secondsLeft };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CREDS);
  if (!sheet) return { success: false, error: 'Credentials sheet not found' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizePhone(data[i][0]) === normalizePhone(phone)) {
      const name      = String(data[i][1] || 'Resident').trim();
      const adminFlag = String(data[i][2] || '').trim().toUpperCase();
      const isAdmin   = (adminFlag === 'TRUE' || adminFlag === 'YES');
      const storedPin = String(data[i][5] || '').trim();

      if (!storedPin) {
        return { success: false, error: 'No PIN set for this account yet. Please contact your Sector 40 Community admin.' };
      }
      if (storedPin !== pin) {
        recordFailedAttempt(phone);
        return { success: false, error: 'Incorrect PIN. Please try again.' };
      }

      // Success — clear any lockout history for this phone
      clearLockout(phone);
      return { success: true, name, isAdmin };
    }
  }
  return { success: false, error: 'Phone not registered' };
}

// ── Lockout helpers (Login Lockouts sheet: Phone | FailCount | LockedUntil) ──

function checkLockout(phone) {
  const sheet = getOrCreateSheet(SHEET_LOCKOUTS);
  ensureHeaders(sheet, ['Phone', 'FailCount', 'LockedUntil']);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (normalizePhone(data[i][0]) === normalizePhone(phone)) {
      const lockedUntil = data[i][2] ? new Date(data[i][2]) : null;
      if (lockedUntil && new Date() < lockedUntil) {
        const secondsLeft = Math.ceil((lockedUntil.getTime() - new Date().getTime()) / 1000);
        return { locked: true, secondsLeft };
      }
      return { locked: false };
    }
  }
  return { locked: false };
}

function recordFailedAttempt(phone) {
  const sheet = getOrCreateSheet(SHEET_LOCKOUTS);
  ensureHeaders(sheet, ['Phone', 'FailCount', 'LockedUntil']);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (normalizePhone(data[i][0]) === normalizePhone(phone)) {
      let failCount = Number(data[i][1] || 0) + 1;
      const row = i + 1;
      sheet.getRange(row, 2).setValue(failCount);
      if (failCount >= MAX_LOGIN_ATTEMPTS) {
        const lockedUntil = new Date(new Date().getTime() + LOCKOUT_SECONDS * 1000);
        sheet.getRange(row, 3).setValue(lockedUntil);
        sheet.getRange(row, 2).setValue(0); // reset counter once locked
      }
      return;
    }
  }
  // No existing row — create one
  sheet.appendRow([phone, 1, '']);
  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 1).setNumberFormat('@STRING@').setValue(phone);
}

function clearLockout(phone) {
  const sheet = getOrCreateSheet(SHEET_LOCKOUTS);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (normalizePhone(data[i][0]) === normalizePhone(phone)) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ── Get Open Issues ───────────────────────────────────────────

function handleGetOpenIssues() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INCIDENTS);
  if (!sheet) return { success: false, error: 'Incidents sheet not found. Please link your Google Form to this spreadsheet.' };

  const data = sheet.getDataRange().getValues();
  const issues = [];

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const status = String(row[COL_STATUS - 1] || '').trim();
    if (status === 'Closed') continue;

    issues.push({
      refId:       String(row[COL_REFID    - 1] || ''),
      logDate:     String(row[COL_LOGDATE  - 1] || ''),
      name:        String(row[COL_NAME     - 1] || ''),
      phone:       String(row[COL_PHONE    - 1] || ''),
      category:    String(row[COL_CATEGORY - 1] || ''),
      description: String(row[COL_DESC     - 1] || ''),
      geo:         String(row[COL_GEO      - 1] || ''),
      photo:       String(row[COL_PHOTO    - 1] || ''),
      status:      status || 'Open'
    });
  }

  // Most recent first
  issues.reverse();
  return { success: true, issues };
}

// ── Close Issue ───────────────────────────────────────────────

function handleCloseIssue(e) {
  const refId = (e.parameter.refId || '').trim();
  const phone = (e.parameter.phone || '').trim();
  if (!refId || !phone) return { success: false, error: 'Missing refId or phone' };

  const isAdmin = isAdminPhone(phone);

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INCIDENTS);
  if (!sheet) return { success: false, error: 'Incidents sheet not found' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowRefId = String(data[i][COL_REFID - 1]).trim();
    const rowPhone = String(data[i][COL_PHONE - 1]).trim();
    if (rowRefId === refId) {
      if (normalizePhone(rowPhone) !== normalizePhone(phone) && !isAdmin) {
        return { success: false, error: 'Not authorised — only the reporter or an admin can close this issue' };
      }
      sheet.getRange(i + 1, COL_STATUS).setValue('Closed');
      return { success: true };
    }
  }
  return { success: false, error: 'Issue not found' };
}

// ── Check if a phone number is flagged as Admin in User Credentials ──

function isAdminPhone(phone) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CREDS);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizePhone(data[i][0]) === normalizePhone(phone)) {
      const flag = String(data[i][2] || '').trim().toUpperCase();
      return (flag === 'TRUE' || flag === 'YES');
    }
  }
  return false;
}

// ═══════════════════════════════════════════════════════════════
//  SELF-REGISTRATION via Email OTP
// ═══════════════════════════════════════════════════════════════
//
//  SHEET TAB REQUIRED: "Pending Registrations"
//  Columns: Phone | Name | Email | OTP | ExpiresAt | CreatedAt | BlockHouseNo | PIN
//
//  Flow:
//  1. requestOtp  — validates phone isn't already registered,
//                    generates a 6-digit OTP, emails it, stores
//                    it in Pending Registrations with an expiry.
//  2. verifyOtp   — checks the OTP matches and hasn't expired,
//                    writes the new row to User Credentials,
//                    deletes the pending row, returns success.
//
// ═══════════════════════════════════════════════════════════════

function handleRequestOtp(e) {
  const name         = (e.parameter.name         || '').trim();
  const phone        = (e.parameter.phone        || '').trim();
  const email        = (e.parameter.email        || '').trim();
  const blockHouseNo = (e.parameter.blockHouseNo || '').trim();
  const pin          = (e.parameter.pin          || '').trim();

  if (!name)  return { success: false, error: 'Name is required' };
  if (!phone) return { success: false, error: 'Phone is required' };
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, error: 'A valid email is required' };
  }
  if (!blockHouseNo) return { success: false, error: 'Block/House number is required' };
  if (!/^\d{4}$/.test(pin)) return { success: false, error: 'PIN must be exactly 4 digits' };

  // Reject if phone is already registered
  const credsSheet = getOrCreateSheet(SHEET_CREDS);
  const credsData  = credsSheet.getDataRange().getValues();
  for (let i = 1; i < credsData.length; i++) {
    if (normalizePhone(credsData[i][0]) === normalizePhone(phone)) {
      return { success: false, error: 'This number is already registered. Please log in instead.' };
    }
  }

  const otp = String(Math.floor(100000 + Math.random() * 900000)); // 6-digit
  const now = new Date();
  const expiresAt = new Date(now.getTime() + OTP_EXPIRY_MINUTES * 60000);

  const pendingSheet = getOrCreateSheet(SHEET_PENDING);
  ensureHeaders(pendingSheet, ['Phone', 'Name', 'Email', 'OTP', 'ExpiresAt', 'CreatedAt', 'BlockHouseNo', 'PIN']);

  // Remove any prior pending row for this phone (e.g. retry / resend)
  removePendingRowsForPhone(pendingSheet, phone);

  pendingSheet.appendRow([phone, name, email, otp, expiresAt, now, blockHouseNo, pin]);
  // Force column A to plain-text format so Sheets doesn't strip the leading '+'
  // or reinterpret the value as a number on this or future writes.
  const newRow = pendingSheet.getLastRow();
  pendingSheet.getRange(newRow, 1).setNumberFormat('@STRING@').setValue(phone);

  // Send the OTP email via Brevo
  try {
    sendOtpEmailViaBrevo(email, name, otp);
  } catch (err) {
    return { success: false, error: 'Could not send verification email: ' + err.message };
  }

  return { success: true, message: 'OTP sent to ' + maskEmail(email) };
}

// ── Send transactional email via Brevo API ──

function sendOtpEmailViaBrevo(toEmail, toName, otp) {
  const payload = {
    sender:  { name: BREVO_SENDER_NAME, email: BREVO_SENDER_EMAIL },
    to:      [{ email: toEmail, name: toName }],
    subject: 'Sector 40 Incident Tracker — Your verification code',
    htmlContent:
      '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:24px;">' +
      '<h2 style="color:#1a5c3a;">Sector 40 Community</h2>' +
      '<p>Hello ' + toName + ',</p>' +
      '<p>Your verification code is:</p>' +
      '<div style="font-size:32px;font-weight:700;letter-spacing:6px;color:#1a5c3a;background:#e8f4ee;padding:16px;text-align:center;border-radius:8px;margin:16px 0;">' + otp + '</div>' +
      '<p style="color:#666;font-size:13px;">This code expires in ' + OTP_EXPIRY_MINUTES + ' minutes. If you did not request this, you can safely ignore this email.</p>' +
      '<p style="color:#666;font-size:13px;">— Sector 40 Community</p>' +
      '</div>'
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'api-key': BREVO_API_KEY,
      'accept': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  if (responseCode !== 201 && responseCode !== 200) {
    throw new Error('Brevo API error (' + responseCode + '): ' + responseBody);
  }
}

function handleVerifyOtp(e) {
  const phone = (e.parameter.phone || '').trim();
  const otp   = (e.parameter.otp   || '').trim();
  if (!phone || !otp) return { success: false, error: 'Missing phone or OTP' };

  const pendingSheet = getOrCreateSheet(SHEET_PENDING);
  const data = pendingSheet.getDataRange().getValues();

  // Collect all pending rows for this phone (there may be duplicates from
  // repeated OTP requests); we'll check the most recent one first.
  let matchIndex = -1;
  for (let i = data.length - 1; i >= 1; i--) {
    if (normalizePhone(data[i][0]) === normalizePhone(phone)) {
      matchIndex = i;
      break; // most recent matching row (rows are appended in order)
    }
  }

  if (matchIndex === -1) {
    return { success: false, error: 'No pending registration found for this number. Please request a new OTP.' };
  }

  const i = matchIndex;
  const rowName        = String(data[i][1] || '').trim();
  const rowEmail       = String(data[i][2] || '').trim();
  const rowOtp         = String(data[i][3] || '').trim();
  const rowExpiresAt   = new Date(data[i][4]);
  const rowBlockHouse  = String(data[i][6] || '').trim();
  const rowPin         = String(data[i][7] || '').trim();

  if (new Date() > rowExpiresAt) {
    pendingSheet.deleteRow(i + 1);
    return { success: false, error: 'OTP expired. Please request a new code.' };
  }
  if (rowOtp !== otp) {
    return { success: false, error: 'Incorrect OTP. Please try again.' };
  }

  // Double-check phone still isn't registered (race condition safety)
  const credsSheet = getOrCreateSheet(SHEET_CREDS);
  const credsData  = credsSheet.getDataRange().getValues();
  for (let j = 1; j < credsData.length; j++) {
    if (normalizePhone(credsData[j][0]) === normalizePhone(phone)) {
      removePendingRowsForPhone(pendingSheet, phone);
      return { success: false, error: 'This number is already registered. Please log in instead.' };
    }
  }

  // Success — write to User Credentials: Phone | Name | IsAdmin | Email | BlockHouseNo | PIN
  credsSheet.appendRow([phone, rowName, '', rowEmail, rowBlockHouse, rowPin]);
  // Force phone column to plain text so it isn't reformatted/stripped of '+'
  const newCredRow = credsSheet.getLastRow();
  credsSheet.getRange(newCredRow, 1).setNumberFormat('@STRING@').setValue(phone);

  removePendingRowsForPhone(pendingSheet, phone);

  return { success: true, name: rowName };
}

function removePendingRowsForPhone(sheet, phone) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (normalizePhone(data[i][0]) === normalizePhone(phone)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function maskEmail(email) {
  const parts = email.split('@');
  if (parts.length !== 2) return email;
  const name = parts[0];
  const masked = name.length <= 2 ? name[0] + '*' : name[0] + '***' + name[name.length - 1];
  return masked + '@' + parts[1];
}

// Normalize a phone string for comparison — strips everything except digits,
// so '+919815448144', '919815448144', and '9.198154e+11' (Sheets number drift)
// all compare equal on their digit content.
function normalizePhone(phone) {
  return String(phone || '').replace(/[^0-9]/g, '');
}

// ── Sheet helpers ───────────────────────────────────────────────

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a5c3a');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}
