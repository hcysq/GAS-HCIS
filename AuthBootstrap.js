/**
 * Jalankan fungsi ini SEKALI untuk memicu authorization OAuth,
 * khususnya izin UrlFetchApp (external_request).
 */
function authorizeHCIS() {
  // Panggilan dummy ke UrlFetchApp untuk memicu permission
  const resp = UrlFetchApp.fetch('https://www.google.com', {
    muteHttpExceptions: true
  });

  Logger.log('Authorization OK. Status: ' + resp.getResponseCode());
  return 'Authorization OK';
}
