/**
 * AVM Labs Marketing Studio — client-code login check.
 *
 * Sheet layout (tab named SHEET_NAME, or else the first tab):
 *   Row 1 = headers:  Client Code | Client Name | Active
 *   Row 2+ = one client per row. "Active" is optional — leave blank or TRUE/YES to allow,
 *            set FALSE/NO to block without deleting the row.
 *
 * Add a row  -> that client can log in immediately.
 * Delete the row (or set Active = NO) -> that client is logged out the next time
 * the app checks (on page load and every few minutes).
 *
 * Endpoints:
 *   ?code=XYZ  -> { ok, code, name }        single-code check
 *   ?list=1    -> { ok, salt, items:[{h,n}] } salted SHA-256 hashes of every active code,
 *                 so the app can check a code instantly without waiting on this script.
 *
 * Deploy: Deploy > Manage deployments > Edit > Version: New version  (keeps the same /exec URL)
 *   Execute as: Me    Who has access: Anyone
 */

const SHEET_NAME = 'Clients';
// Must match AUTH_SALT in Ready-to-Use Designs.dc.html.
const SALT = 'avm-ms-v1:';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.list) return json({ ok: true, salt: SALT, items: activeClients().map(c => ({ h: sha256(SALT + c.code), n: c.name })) });

  const code = normalize(p.code || '');
  if (!code) return json({ ok: false, error: 'missing_code' });
  const hit = activeClients().find(c => c.code === code);
  return hit ? json({ ok: true, code: code, name: hit.name }) : json({ ok: false, error: 'not_found' });
}

function activeClients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (!sheet) return [];
  const rows = sheet.getDataRange().getDisplayValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const code = normalize(rows[i][0]);
    if (!code) continue;
    const active = String(rows[i][2] || '').trim().toUpperCase();
    if (active === 'FALSE' || active === 'NO' || active === 'N' || active === '0') continue;
    out.push({ code: code, name: String(rows[i][1] || '').trim() });
  }
  return out;
}

function sha256(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2))
    .join('');
}

function normalize(v) {
  return String(v).trim().toUpperCase();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
