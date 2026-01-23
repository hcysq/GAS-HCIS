/*************************************************
 * HCIS Sabilul Qur'an - Global Config
 *************************************************/

const CFG = {
  SHEET_USERS: 'Users',
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

/**
 * DEBUG: Test profil retrieval untuk user saat ini
 * Jalankan: testProfilRetrieval()
 */
function testProfilRetrieval() {
  Logger.log('=== TEST PROFIL RETRIEVAL ===');
  
  // Test getProfilUsersDetail()
  const profil = getProfilUsersDetail();
  Logger.log('getProfilUsersDetail() result:');
  Logger.log(JSON.stringify(profil, null, 2));
  
  if (profil.ok && profil.data) {
    Logger.log('\n✅ Data ditemukan! Structure:');
    Logger.log(JSON.stringify({
      summary: profil.data.summary,
      contact: profil.data.contact,
      personal: profil.data.personal
    }, null, 2));
  } else {
    Logger.log(`\n❌ Error: ${profil.msg}`);
  }
  
  return profil;
}

/**
 * DEBUG: Cek headers di Users sheet
 * Jalankan: testUsersSheetHeaders()
 */
function testUsersSheetHeaders() {
  Logger.log('=== TEST USERS SHEET HEADERS ===');
  
  const { sheet: sh, error: err } = getUsersSheetByConfig_();
  if (!sh) {
    Logger.log(`❌ Sheet error: ${err}`);
    return { ok: false, msg: err };
  }
  
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Logger.log(`Sheet: ${sh.getName()}`);
  Logger.log(`Total columns: ${headers.length}`);
  Logger.log('\nHeaders:');
  headers.forEach((h, i) => {
    Logger.log(`  [${i}] ${h}`);
  });
  
  Logger.log('\n✅ Headers retrieved successfully');
  return { ok: true, count: headers.length, headers };
}

/**
 * DEBUG: Cek data row pertama di Users sheet (sample)
 * Jalankan: testUsersSheetSampleRow()
 */
function testUsersSheetSampleRow() {
  Logger.log('=== TEST USERS SHEET SAMPLE ROW ===');
  
  const { sheet: sh, error: err } = getUsersSheetByConfig_();
  if (!sh) {
    Logger.log(`❌ Sheet error: ${err}`);
    return { ok: false, msg: err };
  }
  
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  
  if (lastRow < 2) {
    Logger.log('❌ Tidak ada data di sheet Users (hanya header)');
    return { ok: false, msg: 'No data rows' };
  }
  
  const values = sh.getRange(1, 1, 2, lastCol).getValues();
  const headers = values[0];
  const firstRow = values[1];
  
  Logger.log(`\nFirst data row (Row 2):`);
  headers.forEach((h, i) => {
    const val = firstRow[i];
    if (val || i < 10) { // Tampilkan kolom pertama 10 atau yang berisi data
      Logger.log(`  ${h}: "${val}"`);
    }
  });
  
  Logger.log(`\n✅ Sample row retrieved`);
  return { ok: true, headers, firstRow };
}
