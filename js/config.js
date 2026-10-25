/**
 * Site configuration — edit after deploying the Google Apps Script proxy.
 *
 * Browsers cannot fetch Google Drive/Sheets directly (CORS).
 * Apps Script reads your sheet and returns the rows as JSON.
 * Setup: README.md and drive-proxy/Code.gs
 */
window.APP_CONFIG = {
  /**
   * URL of the deployed Apps Script web app (.../exec).
   * Required — the album data lives on Google Sheets, not in this repo.
   */
  driveProxyUrl: 'https://script.google.com/macros/s/AKfycbwt36nmu9uezVs75nJEGAw9h8vAbvpeqsVBIBLyCjVSHTk2zRYdRM8m38kxA0qF417a/exec'
};
