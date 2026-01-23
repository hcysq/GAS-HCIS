/*************************************************
 * HCIS Sabilul Qur'an - Global Config
 *************************************************/

const CFG = {
  SHEET_USERS: 'Users',
  SHEET_MASTERDATA: 'Masterdata',
  SHEET_SALDO: 'Cuti_Saldo',
  SHEET_CUTI: 'Cuti_Pengajuan',
  SHEET_ATASAN: 'AtasanMap',

  SESSION_TOKEN_KEY: 'HCIS_TOKEN',
  SESSION_TTL_KEY: 'SESSION_TTL_SECONDS',

  UP_KEYS: {
    nip: 'HCIS_NIP',
    nama: 'HCIS_NAMA',
    role: 'HCIS_ROLE',
    email: 'HCIS_EMAIL',
    userId: 'HCIS_USER_ID'
  }
};

function ss_() {
  return SpreadsheetApp.getActive();
}

/**
 * TEST: Cek config Profil (jalankan di Script Editor console)
 * Contoh: testProfilConfig()
 */
function testProfilConfig() {
  const result = validateProfilConfig();
  Logger.log('=== VALIDASI PROFIL CONFIG ===');
  Logger.log(`Status: ${result.ok ? '✅ VALID' : '❌ ADA MASALAH'}`);
  Logger.log(`Summary: ${result.summary}`);
  Logger.log('\nDetail:');
  result.suggestions.forEach(s => Logger.log(s));
  Logger.log('\n');
  Logger.log('Full check object:');
  Logger.log(JSON.stringify(result.checks, null, 2));
  return result;
}
