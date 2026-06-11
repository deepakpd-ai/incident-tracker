// ═══════════════════════════════════════════════════════════════
//  SECTOR 40 INCIDENT TRACKER — Apps Script Backend (Final)
// ═══════════════════════════════════════════════════════════════
//
//  SHEET TABS REQUIRED:
//  1. "User Credentials"  — columns: Phone | Name
//  2. "Form Responses 1"  — auto-created by Google Forms link
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
// ═══════════════════════════════════════════════════════════════

const SHEET_CREDS    = 'User Credentials';
const SHEET_INCIDENTS = 'Form Responses 1';
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
    if      (action === 'auth')          result = handleAuth(e);
    else if (action === 'getOpenIssues') result = handleGetOpenIssues();
    else if (action === 'closeIssue')    result = handleCloseIssue(e);
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

// ── Auth ──────────────────────────────────────────────────────

function handleAuth(e) {
  const phone = (e.parameter.phone || '').trim();
  if (!phone) return { success: false, error: 'No phone provided' };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CREDS);
  if (!sheet) return { success: false, error: 'Credentials sheet not found' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === phone) {
      return { success: true, name: String(data[i][1] || 'Resident').trim() };
    }
  }
  return { success: false, error: 'Phone not registered' };
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

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INCIDENTS);
  if (!sheet) return { success: false, error: 'Incidents sheet not found' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowRefId = String(data[i][COL_REFID - 1]).trim();
    const rowPhone = String(data[i][COL_PHONE - 1]).trim();
    if (rowRefId === refId) {
      if (rowPhone !== phone) return { success: false, error: 'Not authorised — only the reporter can close this issue' };
      sheet.getRange(i + 1, COL_STATUS).setValue('Closed');
      return { success: true };
    }
  }
  return { success: false, error: 'Issue not found' };
}
