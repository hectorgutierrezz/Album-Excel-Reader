/**
 * Site configuration — edit this file after deploying the Google Apps Script proxy.
 *
 * Why a proxy? Browsers cannot fetch Google Drive files directly (CORS).
 * A tiny Apps Script web app reads your private Excel and returns it as JSON+base64.
 * Setup steps are in README.md and drive-proxy/Code.gs.
 */
window.APP_CONFIG = {
  /**
   * URL of the deployed Apps Script web app (.../exec).
   * Leave empty to load the local "Albums List.xlsx" (useful for offline/dev).
   */
  driveProxyUrl: 'https://script.google.com/macros/s/AKfycbwt36nmu9uezVs75nJEGAw9h8vAbvpeqsVBIBLyCjVSHTk2zRYdRM8m38kxA0qF417a/exec',

  /** Local fallback when driveProxyUrl is empty or the proxy fails. */
  localExcelPath: 'Albums List.xlsx'
};
