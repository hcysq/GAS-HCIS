/*************************************************
 * Profile (Masterdata) - Robust + Debug
 *************************************************/

function getProfilMasterdataSaya() {
  try {
    const s = requireLogin_();
    const nipSession = String(s.nip || '').trim();
    const userIdSession = String(s.userId || '').trim();
    const emailSession = String(s.email || '').trim();
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
    const headerMap = buildHeaderMap_(headers);
    const idxNip = findHeaderIdx_(headerMap, ['NIP']); // 0-based
    const idxUserId = findHeaderIdx_(headerMap, ['USER_ID']);
    const idxEmail = findHeaderIdx_(headerMap, ['Email', 'EMAIL']);
    if (idxNip < 0 && idxUserId < 0 && idxEmail < 0) return { ok:false, msg:'Header "NIP", "USER_ID", atau "Email" tidak ditemukan di baris 1 sheet Masterdata.' };

    // Baca data rows (row 2..last)
    const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    let matchedRow = null;

    for (const row of rows) {
      const nipCellKey = idxNip >= 0 ? normalizeNIP_(row[idxNip]) : '';
      const userIdCell = idxUserId >= 0 ? String(row[idxUserId] || '').trim() : '';
      const emailCell = idxEmail >= 0 ? String(row[idxEmail] || '').trim() : '';

      if (nipKey && nipCellKey && nipCellKey === nipKey) { matchedRow = row; break; }
      if (!matchedRow && userIdKey && userIdCell && userIdCell === userIdKey) { matchedRow = row; break; }
      if (!matchedRow && emailSession && emailCell && emailCell.toLowerCase() === emailSession.toLowerCase()) { matchedRow = row; break; }
    }

    if (matchedRow) {
      const data = buildMasterdataPayload_(matchedRow, headers, headerMap);
      return { ok:true, data };
    }

    return { ok:false, msg:`Profil tidak ketemu. Pencarian memakai USER_ID session=${userIdSession || '-'}, NIP session=${nipSession} (key=${nipKey}), dan Email session=${emailSession || '-'}. Cek apakah data Masterdata sudah terisi.` };

  } catch (e) {
    return { ok:false, msg:`Error Profile: ${e && e.message ? e.message : e}` };
  }
}

function buildMasterdataPayload_(row, headers, headerMap) {
  const getRaw = (names) => pickCell_(row, headerMap, Array.isArray(names) ? names : [names]);
  const getText = (names) => txt(getRaw(names));
  const sanitizeValue_ = (val) => {
    const t = txt(val);
    return t ? t : '-';
  };
  const sanitize = (names) => sanitizeValue_(getRaw(names));
  const hasContent = (val) => Boolean(txt(val));

  const tmtRaw = getRaw(['TMT', 'TMT MASUK', 'TMT_MASUK', 'TMT KERJA']);
  const tmtStr = sanitizeValue_(formatDateLocal_(tmtRaw));
  const masaKerja = computeMasaKerjaFromDate_(tmtRaw);

  const ttlField = getText(['TTL']);
  const ttl = ttlField
    || (() => {
      const tempat = getText(['Tempat_Lahir', 'TEMPAT LAHIR', 'Tempat Lahir', 'TEMPAT_LAHIR']);
      const tanggal = formatDateLocal_(getRaw(['Tanggal_Lahir', 'TANGGAL LAHIR', 'Tgl Lahir', 'TANGGAL_LAHIR', 'TGL LAHIR', 'TGL_LAHIR', 'DOB']));
      const parts = [tempat, tanggal].filter(hasContent);
      return parts.length ? parts.join(', ') : '';
    })();

  const hpVal = getText(['No_HP', 'HP', 'NO HP', 'NO. HP', 'NO_HP']);
  const waVal = getText(['WhatsApp', 'WA', 'No_WA', 'WA_Number', 'NO WA', 'WHATSAPP']);

  const buildEmergency = () => ({
    nama: sanitizeValue_(getText(['Darurat_Nama', 'KontakDarurat_Nama', 'KONTAK DARURAT', 'KONTAK DARURAT NAMA', 'KONTAK_DARURAT_NAMA'])),
    hp: sanitizeValue_(getText(['Darurat_HP', 'KontakDarurat_HP', 'Darurat_WA', 'KONTAK DARURAT HP', 'KONTAK_DARURAT_HP', 'HP DARURAT'])),
    hubungan: sanitizeValue_(getText(['Darurat_Hubungan', 'KontakDarurat_Hubungan', 'KONTAK DARURAT HUBUNGAN', 'KONTAK_DARURAT_HUBUNGAN']))
  });

  return {
    summary: {
      nama: sanitize(['Nama', 'NAMA']),
      nip: sanitize(['NIP']),
      jabatan: sanitize(['JABATAN', 'JABATAN STRUKTURAL', 'JABATAN FUNGSIONAL', 'Jabatan']),
      unit: sanitize(['UNIT', 'Unit', 'UNIT KERJA', 'Unit Kerja']),
      status_kepeg: sanitize(['Status_Kepeg', 'Status Kepeg', 'STATUS KEPEGAWAIAN', 'STATUS_KEPEGAWAIAN']),
      tmt: tmtStr || '-',
      masa_kerja: masaKerja || '-'
    },
    contact: {
      hp: sanitizeValue_(hpVal),
      wa: sanitizeValue_(waVal || hpVal),
      email: sanitize(['Email', 'EMAIL']),
      alamat_ktp: sanitize(['Alamat_KTP', 'ALAMAT KTP', 'Alamat KTP']),
      domisili: sanitize(['Alamat_Domisili', 'DOMISILI', 'Alamat Domisili', 'ALAMAT DOMISILI']),
      alamat_detail: sanitize(['Alamat_Detail', 'Alamat Detail', 'Alamat Domisili Detail', 'Alamat Lengkap']),
      kecamatan: sanitize(['Kecamatan_Domisili', 'Kecamatan', 'KECAMATAN']),
      kab_kota: sanitize(['Kab_Kota_Domisili', 'Kab_Kota', 'KAB/KOTA', 'KABUPATEN/KOTA', 'Kota', 'KABUPATEN']),
      darurat: buildEmergency()
    },
    personal: {
      nik: sanitize(['NIK']),
      ttl: sanitizeValue_(ttl),
      gender: sanitize(['Gender', 'Jenis_Kelamin', 'Jenis Kelamin', 'GENDER', 'JK', 'JENIS KELAMIN']),
      status_nikah: sanitize(['Status_Nikah', 'Status Nikah', 'STATUS NIKAH', 'STATUS PERNIKAHAN']),
      bpjs_kes: sanitize(['BPJS_Kes', 'BPJS KESEHATAN', 'BPJS_KES']),
      bpjs_tk: sanitize(['BPJS_TK', 'BPJS Ketenagakerjaan', 'BPJS KETENAGAKERJAAN', 'BPJSTK']),
      pendidikan_terakhir: sanitize(['Pendidikan_Terakhir', 'Pend_Terakhir', 'Pendidikan Terakhir']),
      pendidikan_str: sanitize(['Pendidikan_Terakhir', 'Pend_Terakhir', 'Pendidikan Terakhir'])
    },
    edu_formal: buildFormalEduDynamic_(row, headers),
    edu_nonformal: buildNonFormalEduDynamic_(row, headers)
  };
}

function buildFormalEduDynamic_(row, headers) {
  const groups = {};
  const order = [];
  const hasContent = (v) => Boolean(txt(v));
  const sanitizeValue_ = (val) => {
    const t = txt(val);
    return t ? t : '-';
  };

  headers.forEach((h, idx) => {
    const header = String(h || '').trim();
    const lower = header.toLowerCase();
    if (!lower.startsWith('pend_')) return;

    const remainder = header.substring(5);
    if (!remainder) return;

    const parts = remainder.split('_');
    const key = parts.shift();
    if (!key) return;
    const fieldKey = parts.join('_') || 'nama';

    if (!groups[key]) {
      groups[key] = { level: key, nama: '-', jur: '-', thn: '-', link: '-' };
      order.push(key);
    }

    const normalizedField = normalizeEduField_(fieldKey, true);
    const val = sanitizeValue_(row[idx]);

    if (normalizedField === 'jur') groups[key].jur = val;
    else if (normalizedField === 'thn') groups[key].thn = val;
    else if (normalizedField === 'link') groups[key].link = val;
    else groups[key].nama = val;
  });

  return order
    .map(k => groups[k])
    .filter(g => hasContent(g.nama));
}

function buildNonFormalEduDynamic_(row, headers) {
  const groups = {};
  const order = [];
  const hasContent = (v) => Boolean(txt(v));
  const sanitizeValue_ = (val) => {
    const t = txt(val);
    return t ? t : '-';
  };

  headers.forEach((h, idx) => {
    const header = String(h || '').trim();
    const lower = header.toLowerCase();
    if (!lower.startsWith('nonformal_')) return;

    const remainder = header.substring('nonformal_'.length);
    if (!remainder) return;

    const parts = remainder.split('_');
    const key = parts.shift();
    if (!key) return;
    const fieldKey = parts.join('_') || 'nama';

    if (!groups[key]) {
      groups[key] = { nama: '-', prog: '-', thn: '-', link: '-' };
      order.push(key);
    }

    const normalizedField = normalizeEduField_(fieldKey, false);
    const val = sanitizeValue_(row[idx]);

    if (normalizedField === 'prog') groups[key].prog = val;
    else if (normalizedField === 'thn') groups[key].thn = val;
    else if (normalizedField === 'link') groups[key].link = val;
    else groups[key].nama = val;
  });

  return order
    .map(k => groups[k])
    .filter(g => hasContent(g.nama));
}

function normalizeEduField_(fieldKey, isFormal) {
  const f = String(fieldKey || '').toLowerCase();
  if (f.includes('jur')) return 'jur';
  if (f.includes('thn') || f.includes('tahun') || f === 'th') return 'thn';
  if (f.includes('link') || f.includes('url') || f.includes('ijazah')) return 'link';
  if (!isFormal && (f.includes('prog') || f.includes('program'))) return 'prog';
  return 'nama';
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

function getProfilUsersDetail() {
  try {
    const s = requireLogin_();
    const nipSession = String(s.nip || '').trim();
    const userIdSession = String(s.userId || '').trim();
    const nipKey = normalizeNIP_(nipSession);

    if (!nipKey && !userIdSession) {
      return { ok:false, msg:'Session tidak memiliki USER_ID atau NIP untuk pencarian.' };
    }

    const { sheet: sh, error: sheetErr } = getUsersSheetByConfig_();
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

function getUsersSheetByConfig_() {
  try {
    const { ss, error: ssErr } = getMasterdataSpreadsheet_();
    if (!ss) return { sheet: null, error: ssErr || 'Spreadsheet Masterdata tidak tersedia (cek MASTERDATA_SS_ID).' };

    const gidRaw = cfgGet('MASTERDATA_GID', '');
    const gid = Number(gidRaw);
    if (!isNaN(gid) && gid > 0) {
      const byId = ss.getSheets().find(sh => sh.getSheetId() === gid);
      if (byId) return { sheet: byId };
      return { sheet: null, error:`Sheet dengan GID ${gid} (MASTERDATA_GID) tidak ditemukan di spreadsheet Masterdata.` };
    }

    const sh = ss.getSheetByName(CFG.SHEET_USERS);
    if (sh) return { sheet: sh };
    return { sheet: null, error:`Sheet "${CFG.SHEET_USERS}" tidak ditemukan pada spreadsheet Masterdata.` };
  } catch (e) {
    const errMsg = e && e.message ? e.message : e;
    return { sheet: null, error:`Gagal membuka spreadsheet Users (MASTERDATA_SS_ID): ${errMsg}` };
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
