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
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me    Who has access: Anyone
 * Copy the /exec URL into AUTH_URL in Ready-to-Use Designs.dc.html.
 */

const SHEET_NAME = 'Clients';

function doGet(e) {
  const code = normalize((e && e.parameter && e.parameter.code) || '');
  if (!code) return json({ ok: false, error: 'missing_code' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (!sheet) return json({ ok: false, error: 'sheet_not_found' });

  const rows = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (normalize(row[0]) !== code) continue;
    const active = String(row[2] || '').trim().toUpperCase();
    if (active === 'FALSE' || active === 'NO' || active === 'N' || active === '0') {
      return json({ ok: false, error: 'inactive' });
    }
    return json({ ok: true, code: code, name: String(row[1] || '').trim() });
  }
  return json({ ok: false, error: 'not_found' });
}

function normalize(v) {
  return String(v).trim().toUpperCase();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
