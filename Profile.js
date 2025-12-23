/*************************************************
 * Profile (Masterdata) - Robust + Debug
 *************************************************/

function getProfilMasterdataSaya() {
  try {
    const s = requireLogin_();
    const nipSession = String(s.nip || '').trim();
    const userIdSession = String(s.userId || '').trim();
    const nipKey = normalizeNIP_(nipSession);
    const userIdKey = userIdSession;
    if (!nipKey && !userIdKey) return { ok:false, msg:'Session tidak memiliki NIP atau USER_ID. Coba logout lalu login ulang.' };

    // Pastikan sheet/tab Masterdata benar-benar ada (mendukung konfigurasi GID)
    const { sheet: sh, error: sheetErr } = getMasterdataSheet_();
    if (!sh) return { ok:false, msg: sheetErr || 'Sheet Masterdata tidak ditemukan.' };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { ok:false, msg:'Sheet Masterdata kosong atau tidak ada data.' };

    // Baca header (row 1)
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h||'').trim());
    const idxNip = findHeaderIndex_(headers, 'NIP'); // 0-based
    const idxUserId = findHeaderIndex_(headers, 'USER_ID');
    if (idxNip < 0 && idxUserId < 0) return { ok:false, msg:'Header "NIP" atau "USER_ID" tidak ditemukan di baris 1 sheet Masterdata.' };

    // Baca data rows (row 2..last)
    const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    for (const row of rows) {
      const nipCellKey = idxNip >= 0 ? normalizeNIP_(row[idxNip]) : '';
      const userIdCell = idxUserId >= 0 ? String(row[idxUserId] || '').trim() : '';
      const nipMatches = nipKey && nipCellKey && nipCellKey === nipKey;
      const userIdMatches = userIdKey && userIdCell && userIdCell === userIdKey;

      const userIdAllowed = userIdMatches && (!nipCellKey || !nipKey || nipCellKey === nipKey);

      if (nipMatches || userIdAllowed) {
        const obj = {};
        headers.forEach((k, i) => obj[k] = row[i]);
        return { ok:true, data: obj };
      }
    }

    return { ok:false, msg:`Profil tidak ketemu. Pencarian memakai USER_ID session=${userIdSession || '-'} dan NIP session=${nipSession} (key=${nipKey}). Cek apakah data Masterdata sudah terisi.` };

  } catch (e) {
    return { ok:false, msg:`Error Profile: ${e && e.message ? e.message : e}` };
  }
}

/**
 * DEBUG: panggil ini dari browser / console via google.script.run
 * untuk melihat apa yang kebaca dari Masterdata & session.
 */
function debugProfilMasterdataSaya() {
  try {
    const s = requireLogin_();
    const nipSession = String(s.nip || '').trim();
    const nipKey = normalizeNIP_(nipSession);

    const { sheet: sh, error: sheetErr } = getMasterdataSheet_();
    if (!sh) return { ok:false, msg: sheetErr || 'Sheet Masterdata tidak ditemukan.' };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h||'').trim());
    const idxNip = findHeaderIndex_(headers, 'NIP');

    // ambil contoh 10 NIP pertama untuk cek format
    const sample = [];
    if (lastRow >= 2 && idxNip >= 0) {
      const n = Math.min(10, lastRow - 1);
      const vals = sh.getRange(2, 1, n, lastCol).getValues();
      for (let i = 0; i < vals.length; i++) {
        sample.push({
          row: i + 2,
          raw: vals[i][idxNip],
          normalized: normalizeNIP_(vals[i][idxNip])
        });
      }
    }

    return {
      ok:true,
      session: { nip: nipSession, nipKey },
      sheet: { name: CFG.SHEET_MASTERDATA, lastRow, lastCol },
      headerNIPIndex0: idxNip,
      headersPreview: headers.slice(0, 15),
      nipSamples: sample
    };

  } catch (e) {
    return { ok:false, msg:`Error debugProfil: ${e && e.message ? e.message : e}` };
  }
}

/**
 * Ambil sheet Masterdata dengan prioritas GID (konfigurasi MASTERDATA_GID di HCIS_Config),
 * fallback ke nama tab default dari CFG.SHEET_MASTERDATA.
 */
function getMasterdataSheet_() {
  const { ss, error: ssErr } = getMasterdataSpreadsheet_();
  if (!ss) return { sheet: null, error: ssErr };

  const gidRaw = cfgGet('MASTERDATA_GID', '');
  const gid = Number(gidRaw);
  if (!isNaN(gid) && gid > 0) {
    const targetById = ss.getSheets().find(sh => sh.getSheetId() === gid);
    if (targetById) return { sheet: targetById };
    return { sheet: null, error: `Sheet Masterdata dengan GID ${gid} tidak ditemukan pada spreadsheet yang dikonfigurasi. Cek MASTERDATA_GID di HCIS_Config.` };
  }

  const sh = ss.getSheetByName(CFG.SHEET_MASTERDATA);
  if (sh) return { sheet: sh };
  return { sheet: null, error: `Sheet "${CFG.SHEET_MASTERDATA}" tidak ditemukan pada spreadsheet Masterdata.` };
}

function getMasterdataSpreadsheet_() {
  const ssId = cfgGetString('MASTERDATA_SS_ID', '');
  if (!ssId) return { ss: SpreadsheetApp.getActive() };

  try {
    return { ss: SpreadsheetApp.openById(ssId) };
  } catch (e) {
    const errMsg = e && e.message ? e.message : e;
    return { ss: null, error: `Gagal membuka spreadsheet Masterdata (MASTERDATA_SS_ID di HCIS_Config): ${errMsg}` };
  }
}

/** Header finder yang tahan spasi/case */
function findHeaderIndex_(headers, name) {
  const target = String(name||'').trim().toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]||'').trim().toLowerCase();
    if (h === target) return i;
  }
  return -1;
}

/** Normalisasi NIP biar aman dibandingkan */
function normalizeNIP_(v) {
  const s = String(v ?? '').trim();
  if (!s) return '';
  const digits = s.replace(/[^\d]/g, '');
  return digits || s;
}

/*************************************************
 * Profil Users (structured)
 *************************************************/

const USERS_SPREADSHEET_ID = '1ImaTVL7aBk3DOV5bLgIJX3XPeaUJRfkai1nWlXXhNLU';

function getProfilUsersDetail() {
  try {
    const s = requireLogin_();
    const nipSession = String(s.nip || '').trim();
    const userIdSession = String(s.userId || '').trim();
    const nipKey = normalizeNIP_(nipSession);

    if (!nipKey && !userIdSession) {
      return { ok:false, msg:'Session tidak memiliki USER_ID atau NIP untuk pencarian.' };
    }

    const { sheet: sh, error: sheetErr } = getUsersSheetById_();
    if (!sh) return { ok:false, msg: sheetErr || 'Sheet Users tidak ditemukan.' };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { ok:false, msg:'Sheet Users kosong atau belum ada data.' };

    const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const headersRow = values[0].map(h => String(h || '').trim());
    const headerMap = buildHeaderMap_(headersRow);

    const idxNip = findHeaderIdx_(headerMap, ['NIP']);
    const idxUserId = findHeaderIdx_(headerMap, ['USER_ID']);
    if (idxNip < 0 && idxUserId < 0) {
      return { ok:false, msg:'Header "NIP" atau "USER_ID" tidak ditemukan di sheet Users.' };
    }

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const nipCellKey = idxNip >= 0 ? normalizeNIP_(row[idxNip]) : '';
      const userIdCell = idxUserId >= 0 ? String(row[idxUserId] || '').trim() : '';

      const nipMatches = nipKey && nipCellKey && nipCellKey === nipKey;
      const userIdMatches = userIdSession && userIdCell && userIdCell === userIdSession;
      const userIdAllowed = userIdMatches && (!nipKey || !nipCellKey || nipCellKey === nipKey);

      if (nipMatches || userIdAllowed) {
        const data = buildStructuredProfile_(row, headerMap);
        return { ok:true, data };
      }
    }

    return { ok:false, msg:`Data Users tidak ditemukan untuk USER_ID=${userIdSession || '-'} / NIP=${nipSession || '-'}.` };

  } catch (e) {
    return { ok:false, msg:`Error Profil Users: ${e && e.message ? e.message : e}` };
  }
}

function getUsersSheetById_() {
  try {
    const ss = SpreadsheetApp.openById(USERS_SPREADSHEET_ID);
    const sh = ss.getSheetByName(CFG.SHEET_USERS);
    if (!sh) return { sheet: null, error:`Sheet "${CFG.SHEET_USERS}" tidak ditemukan pada spreadsheet Users.` };
    return { sheet: sh };
  } catch (e) {
    const errMsg = e && e.message ? e.message : e;
    return { sheet: null, error:`Gagal membuka spreadsheet Users: ${errMsg}` };
  }
}

function buildHeaderMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || '').trim();
    if (!key) return;
    map[key] = i;
    const lower = key.toLowerCase();
    if (!Object.prototype.hasOwnProperty.call(map, lower)) map[lower] = i;
  });
  return map;
}

function findHeaderIdx_(map, names) {
  for (const n of names) {
    if (Object.prototype.hasOwnProperty.call(map, n)) return map[n];
    const lower = String(n || '').toLowerCase();
    if (Object.prototype.hasOwnProperty.call(map, lower)) return map[lower];
  }
  return -1;
}

function pickCell_(row, map, names) {
  const idx = findHeaderIdx_(map, names);
  if (idx === -1) return '';
  return row[idx];
}

function buildStructuredProfile_(row, headerMap) {
  const get = (names) => pickCell_(row, headerMap, Array.isArray(names) ? names : [names]);
  const txtVal = (names) => txt(get(names));

  const tmtRaw = get(['TMT']);
  const tmtStr = formatDateLocal_(tmtRaw);
  const masaKerja = computeMasaKerjaFromDate_(tmtRaw);

  const pendidikanAkhir = txtVal(['Pendidikan_Terakhir', 'Pend_Terakhir']);

  return {
    summary: {
      nama: txtVal(['Nama']),
      nip: txtVal(['NIP']),
      jabatan: txtVal(['JABATAN', 'JABATAN STRUKTURAL', 'JABATAN FUNGSIONAL', 'Jabatan']),
      unit: txtVal(['UNIT', 'Unit']),
      status_kepeg: txtVal(['Status_Kepeg', 'Status Kepeg'] ),
      tmt: tmtStr,
      masa_kerja: masaKerja
    },
    contact: {
      hp: txtVal(['No_HP', 'HP']),
      wa: txtVal(['WhatsApp', 'WA', 'No_WA', 'WA_Number']) || txtVal(['No_HP', 'HP']),
      email: txtVal(['Email']),
      alamat_ktp: txtVal(['Alamat_KTP']),
      domisili: txtVal(['Alamat_Domisili', 'Domisili']),
      kecamatan: txtVal(['Kecamatan_Domisili', 'Kecamatan']),
      kab_kota: txtVal(['Kab_Kota_Domisili', 'Kab_Kota']),
      darurat: {
        nama: txtVal(['Darurat_Nama', 'KontakDarurat_Nama']),
        hp: txtVal(['Darurat_HP', 'KontakDarurat_HP', 'Darurat_WA']),
        hubungan: txtVal(['Darurat_Hubungan', 'KontakDarurat_Hubungan'])
      }
    },
    personal: {
      nik: txtVal(['NIK']),
      ttl: buildTTL_(txtVal(['Tempat_Lahir', 'Tempat Lahir']), get(['Tanggal_Lahir', 'Tanggal Lahir'])),
      gender: txtVal(['Gender', 'Jenis_Kelamin', 'Jenis Kelamin']),
      status_nikah: txtVal(['Status_Nikah', 'Status Nikah']),
      bpjs_kes: txtVal(['BPJS_Kes']),
      bpjs_tk: txtVal(['BPJS_TK', 'BPJS Ketenagakerjaan']),
      pendidikan_terakhir: pendidikanAkhir,
      pendidikan_str: pendidikanAkhir
    },
    edu_formal: buildFormalEdu_(row, headerMap),
    edu_nonformal: buildNonFormalEdu_(row, headerMap)
  };
}

function buildFormalEdu_(row, headerMap) {
  const levels = ['SD', 'SMP', 'SMA', 'S1', 'S2', 'S3'];
  const list = [];

  levels.forEach(lv => {
    const nama = txt(pickCell_(row, headerMap, [`Pend_${lv}`, `Pend_${lv}_Nama`]));
    const jur = txt(pickCell_(row, headerMap, [`Pend_${lv}_Jurusan`]));
    const thn = txt(pickCell_(row, headerMap, [`Pend_${lv}_Thn`, `Pend_${lv}_Tahun`]));
    const link = txt(pickCell_(row, headerMap, [`Pend_${lv}_Link`]));

    if (nama || jur || thn || link) {
      list.push({ level: lv, nama, jur, thn, link });
    }
  });

  return list;
}

function buildNonFormalEdu_(row, headerMap) {
  const list = [];
  for (let i = 1; i <= 3; i++) {
    const nama = txt(pickCell_(row, headerMap, [`NonFormal_${i}`, `NonFormal_${i}_Nama`]));
    const prog = txt(pickCell_(row, headerMap, [`NonFormal_${i}_Program`, `NonFormal_${i}_Prog`]));
    const thn = txt(pickCell_(row, headerMap, [`NonFormal_${i}_Thn`, `NonFormal_${i}_Tahun`]));
    const link = txt(pickCell_(row, headerMap, [`NonFormal_${i}_Link`]));

    if (nama || prog || thn || link) {
      list.push({ nama, prog, thn, link });
    }
  }
  return list;
}

function buildTTL_(tempat, tglRaw) {
  const tempatStr = String(tempat || '').trim();
  const tanggalStr = formatDateLocal_(tglRaw);
  if (tempatStr && tanggalStr) return `${tempatStr}, ${tanggalStr}`;
  return tempatStr || tanggalStr || '';
}

function formatDateLocal_(v) {
  if (!v) return '';
  try {
    const date = new Date(v);
    if (isNaN(date.getTime())) return '';
    const tz = Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Asia/Jakarta';
    return Utilities.formatDate(date, tz, 'dd MMM yyyy');
  } catch (e) {
    return '';
  }
}

function computeMasaKerjaFromDate_(tmtVal) {
  try {
    if (!tmtVal) return '-';
    const dt = new Date(tmtVal);
    if (isNaN(dt.getTime())) return '-';
    const now = new Date();
    let years = now.getFullYear() - dt.getFullYear();
    let months = now.getMonth() - dt.getMonth();
    if (months < 0) { years -= 1; months += 12; }
    if (years < 0) return '-';
    if (years === 0) return `${months} bulan`;
    return `${years} tahun ${months} bulan`;
  } catch (e) {
    return '-';
  }
}
