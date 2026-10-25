/**
 * Google Apps Script — sheet proxy for Music Library Browser
 *
 * Setup:
 * 1. Open https://script.google.com → New project
 * 2. Paste this entire file into Code.gs
 * 3. Set DRIVE_FILE_ID (from the Sheets/Drive URL)
 * 4. Deploy → New deployment → Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web app URL (.../exec) into js/config.js → driveProxyUrl
 *
 * Uses SpreadsheetApp only (no UrlFetchApp) — works with Google Sheets.
 * After you edit the sheet, no redeploy is needed.
 */

/** @type {string} Google Sheets file ID of Albums List */
var DRIVE_FILE_ID = 'PASTE_YOUR_DRIVE_FILE_ID_HERE';

/**
 * Serves the first sheet as JSON rows for the static site.
 * @param {Object} e - unused
 * @return {TextOutput}
 */
function doGet(e) {
  try {
    if (!DRIVE_FILE_ID || DRIVE_FILE_ID.indexOf('PASTE_') === 0) {
      return jsonResponse_({
        ok: false,
        error: 'Set DRIVE_FILE_ID in Code.gs to your Google Sheets file ID.'
      });
    }

    var ss = SpreadsheetApp.openById(DRIVE_FILE_ID);
    var sheet = ss.getSheets()[0];
    var values = sheet.getDataRange().getDisplayValues();

    return jsonResponse_({
      ok: true,
      format: 'rows',
      name: ss.getName(),
      rows: values
    });
  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: String(err && err.message ? err.message : err)
    });
  }
}

/**
 * @param {Object} obj
 * @return {TextOutput}
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
