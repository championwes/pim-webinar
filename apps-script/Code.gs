/**
 * Receives webinar registrations from the landing page form and appends them
 * to the "Precision Flowline Webinar Registrations" Google Sheet.
 *
 * Setup: open the sheet > Extensions > Apps Script, paste this file in,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Put the web app URL in REGISTRATION_ENDPOINT in new-page/index.html.
 */

const SHEET_ID = '1sbi7DUSk7nJTtp5pS-Op071m-UON-lNTqktcCC_yHbA';
const FIELDS = ['full_name', 'organization', 'phone', 'email', 'page_url'];

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Bots fill the hidden "website" field; real visitors never see it.
  if (p.website) return json({ ok: true });

  const values = FIELDS.map(f => clean(p[f]));
  if (!values[0] || !values[3]) return json({ ok: false, error: 'missing required fields' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    sheet.appendRow([new Date()].concat(values));
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}

// Trim, cap length, and stop values like "=HYPERLINK(...)" from running as formulas.
function clean(v) {
  const s = String(v || '').trim().slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
