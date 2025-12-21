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
  SESSION_TTL_SECONDS: 21600, // 6 jam

  UP_KEYS: {
    nip: 'HCIS_NIP',
    nama: 'HCIS_NAMA',
    role: 'HCIS_ROLE',
    email: 'HCIS_EMAIL'
  }
};

function ss_() {
  return SpreadsheetApp.getActive();
}
