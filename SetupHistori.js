/*************************************************
 * HCIS - Setup Histori Kepegawaian (FAST, One-time Run)
 * Anti-timeout: batasi formatting/validations ke LIMIT_ROWS
 *************************************************/

const HIST_CFG = {
  SHEET_USERS: 'Users',
  SHEET_HISTORI: 'Histori_Kepegawaian',
  SHEET_DOK: 'Histori_Dokumen',
  SHEET_CFG: 'HCIS_Config',

  // Batas baris untuk data validation & number format (anti-timeout)
  LIMIT_ROWS: 2000, // bisa kamu naikkan kalau mau (mis. 5000)

  // opsional: buat folder Drive khusus histori
  CREATE_DRIVE_FOLDER: true,
  DRIVE_FOLDER_NAME: 'HCIS - Dokumen Histori Kepegawaian',

  enums: {
    jenisHistori: [
      'MUTASI','PROMOSI','DEMOSI','SK_TUGAS','SP','TEGURAN','PENGANGKATAN','KONTRAK','LAINNYA'
    ],
    statusHistori: ['DRAFT','AKTIF','DIBATALKAN'],
    visibility: ['PEGAWAI','ADMIN_ONLY'],
    jenisDok: ['SK','SP','SURAT_TUGAS','NOTA_DINAS','LAINNYA']
  }
};

function setupHistoriKepegawaian() {
  const ss = SpreadsheetApp.getActive();

  // Validasi Users minimal ada
  const shUsers = ss.getSheetByName(HIST_CFG.SHEET_USERS);
  if (!shUsers) throw new Error(`Sheet "${HIST_CFG.SHEET_USERS}" tidak ditemukan. Jalankan setup Users dulu.`);

  // 1) Create/normalize sheets (header saja, aman run ulang)
  const shHist = createWithHeader_(ss, HIST_CFG.SHEET_HISTORI, getHeaderHistori_());
  const shDok  = createWithHeader_(ss, HIST_CFG.SHEET_DOK,     getHeaderHistoriDok_());

  // 2) Config sheet
  const shCfg  = getOrCreateSheet_(ss, HIST_CFG.SHEET_CFG);
  ensureCfgHeader_(shCfg);

  // 3) Optional: ensure Drive folder config (dibuat hanya jika belum ada)
  if (HIST_CFG.CREATE_DRIVE_FOLDER) {
    ensureDriveFolderConfig_(shCfg);
  }

  // 4) Apply validations (dibatasi LIMIT_ROWS)
  applyHistoriValidations_(shHist, shDok, HIST_CFG.LIMIT_ROWS);

  // 5) Basic formatting (ringan)
  [shHist, shDok, shCfg].forEach(sh => formatSheetFast_(sh));

  SpreadsheetApp.getUi().alert(
    'Setup Histori Kepegawaian selesai ✅\n' +
    `- Sheet: ${HIST_CFG.SHEET_HISTORI}, ${HIST_CFG.SHEET_DOK}\n` +
    `- Validasi/format dibatasi sampai ${HIST_CFG.LIMIT_ROWS} baris (anti-timeout)\n` +
    (HIST_CFG.CREATE_DRIVE_FOLDER ? '- Folder Drive histori dicek/diisi di HCIS_Config\n' : '')
  );
}

/** ===================== HEADERS ===================== */

function getHeaderHistori_(){
  return [
    'HistoriID',
    'UserID',
    'NIP',
    'TanggalEfektif',
    'Jenis',
    'Judul',
    'Keterangan',
    'NoSurat',
    'TanggalSurat',
    'Status',
    'Visibility',
    'CreatedAt',
    'CreatedBy',
    'UpdatedAt',
    'UpdatedBy'
  ];
}

function getHeaderHistoriDok_(){
  return [
    'DocID',
    'HistoriID',
    'UserID',
    'NIP',
    'JenisDokumen',
    'NamaFile',
    'FileURL',
    'UploadedAt',
    'UploadedBy'
  ];
}

/** ===================== CREATE / FORMAT ===================== */

function createWithHeader_(ss, name, header) {
  const sh = getOrCreateSheet_(ss, name);

  // Pastikan minimal kolom cukup
  if (sh.getMaxColumns() < header.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), header.length - sh.getMaxColumns());
  }

  // Tulis header row 1
  const current = sh.getRange(1, 1, 1, header.length).getValues()[0];
  const curStr = current.join('|').trim();
  const targetStr = header.join('|').trim();

  if (curStr !== targetStr) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }

  return sh;
}

function getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// formatting ringan, tidak autoResize (sering bikin lambat)
function formatSheetFast_(sh) {
  sh.setFrozenRows(1);
  const lc = sh.getLastColumn();
  if (lc > 0) sh.getRange(1, 1, 1, lc).setFontWeight('bold');
}

/** ===================== VALIDATIONS ===================== */

function applyHistoriValidations_(shHist, shDok, limitRows){
  // Histori enums
  setValidationListByHeader_(shHist, 'Jenis', HIST_CFG.enums.jenisHistori, limitRows);
  setValidationListByHeader_(shHist, 'Status', HIST_CFG.enums.statusHistori, limitRows);
  setValidationListByHeader_(shHist, 'Visibility', HIST_CFG.enums.visibility, limitRows);

  // Dokumen enums
  setValidationListByHeader_(shDok, 'JenisDokumen', HIST_CFG.enums.jenisDok, limitRows);

  // Plain text IDs & NIP (dibatasi)
  setPlainTextByHeader_(shHist, 'HistoriID', limitRows);
  setPlainTextByHeader_(shHist, 'UserID', limitRows);
  setPlainTextByHeader_(shHist, 'NIP', limitRows);

  setPlainTextByHeader_(shDok, 'DocID', limitRows);
  setPlainTextByHeader_(shDok, 'HistoriID', limitRows);
  setPlainTextByHeader_(shDok, 'UserID', limitRows);
  setPlainTextByHeader_(shDok, 'NIP', limitRows);
}

function setValidationListByHeader_(sh, headerName, items, limitRows) {
  const col = findColByHeader_(sh, headerName);
  if (!col) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(items, true)
    .setAllowInvalid(true)
    .build();

  const rows = Math.max(1, Number(limitRows || 2000));
  sh.getRange(2, col, rows, 1).setDataValidation(rule);
}

function setPlainTextByHeader_(sh, headerName, limitRows){
  const col = findColByHeader_(sh, headerName);
  if (!col) return;

  const rows = Math.max(1, Number(limitRows || 2000));
  sh.getRange(2, col, rows, 1).setNumberFormat('@');
}

function findColByHeader_(sh, headerName) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x||'').trim());
  const idx = headers.indexOf(headerName);
  return idx === -1 ? null : idx + 1;
}

/** ===================== CONFIG SHEET & DRIVE FOLDER ===================== */

function ensureCfgHeader_(shCfg){
  const needed = ['Key','Value','Note'];
  if (shCfg.getLastRow() === 0) {
    shCfg.getRange(1,1,1,needed.length).setValues([needed]);
    return;
  }
  const h = shCfg.getRange(1,1,1,Math.max(shCfg.getLastColumn(), needed.length)).getValues()[0].map(x => String(x||'').trim());
  const ok = needed.every((k,i)=> h[i] === k);
  if (!ok) {
    shCfg.clear({contentsOnly:true});
    shCfg.getRange(1,1,1,needed.length).setValues([needed]);
  }
}

function ensureDriveFolderConfig_(shCfg){
  const key = 'HISTORI_FOLDER_ID';

  const lastRow = shCfg.getLastRow();
  if (lastRow < 2) {
    // belum ada data config sama sekali -> buat folder langsung
    const folderId = DriveApp.createFolder(HIST_CFG.DRIVE_FOLDER_NAME).getId();
    shCfg.appendRow([key, folderId, 'Folder Drive untuk dokumen histori kepegawaian']);
    return;
  }

  const data = shCfg.getRange(2, 1, lastRow - 1, 3).getValues();
  for (const r of data) {
    if (String(r[0] || '').trim() === key && String(r[1] || '').trim()) {
      return; // sudah ada
    }
  }

  const folderId = DriveApp.createFolder(HIST_CFG.DRIVE_FOLDER_NAME).getId();
  shCfg.appendRow([key, folderId, 'Folder Drive untuk dokumen histori kepegawaian']);
}
