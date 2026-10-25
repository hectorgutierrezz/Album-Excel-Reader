/**
 * Google Apps Script — Excel proxy for Music Library Browser
 *
 * Setup:
 * 1. Open https://script.google.com → New project
 * 2. Paste this entire file into Code.gs (replace the default content)
 * 3. Set DRIVE_FILE_ID below (from the Drive/Sheets URL)
 * 4. Deploy → New deployment → Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web app URL (.../exec) into js/config.js → driveProxyUrl
 *
 * Supports both native Google Sheets and uploaded .xlsx files.
 * After you edit the sheet on Drive, no redeploy is needed —
 * each page load fetches the current data.
 */

/** @type {string} Google Drive / Sheets file ID of Albums List */
var DRIVE_FILE_ID = 'PASTE_YOUR_DRIVE_FILE_ID_HERE';

/**
 * Serves the spreadsheet as Excel (base64 JSON) for the static site.
 * @param {Object} e - Apps Script event (unused)
 * @return {TextOutput}
 */
function doGet(e) {
  try {
    if (!DRIVE_FILE_ID || DRIVE_FILE_ID.indexOf('PASTE_') === 0) {
      return jsonResponse_({
        ok: false,
        error: 'Set DRIVE_FILE_ID in Code.gs to your Google Drive Excel/Sheets file ID.'
      });
    }

    var file = DriveApp.getFileById(DRIVE_FILE_ID);
    var mime = file.getMimeType();
    var blob;
    var name = file.getName();

    if (mime === MimeType.GOOGLE_SHEETS) {
      // Native Sheets → export as real .xlsx (getBlob() would return PDF)
      var exportUrl =
        'https://docs.google.com/spreadsheets/d/' +
        DRIVE_FILE_ID +
        '/export?format=xlsx';
      var token = ScriptApp.getOAuthToken();
      var response = UrlFetchApp.fetch(exportUrl, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() !== 200) {
        throw new Error(
          'Sheets export failed (HTTP ' + response.getResponseCode() + '). ' +
          'Re-authorize the web app after updating the script.'
        );
      }
      blob = response.getBlob().setName(name + '.xlsx');
      name = name + '.xlsx';
    } else {
      blob = file.getBlob();
    }

    return jsonResponse_({
      ok: true,
      name: name,
      mimeType: blob.getContentType(),
      modified: file.getLastUpdated().toISOString(),
      data: Utilities.base64Encode(blob.getBytes())
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
