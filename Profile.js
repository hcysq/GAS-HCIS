/*************************************************
 * Profile (Masterdata) - Robust + Debug
 *************************************************/

function getProfilMasterdataSaya(payload) {
  try {
    const s = requireLogin_(payload.nip, payload.deviceId, payload.token);
    const nipSession = String(s.nip || '').trim();
    const userIdSession = String(s.userId || '').trim();
    const emailSession = String(s.email || '').trim();
    const nipKey = normalizeNIP_(nipSession);
    const userIdKey = userIdSession;
    if (!nipKey && !userIdKey) return { ok:false, msg:'Session tidak memiliki NIP atau USER_ID. Coba logout lalu login ulang.' };

    // Ambil sheet Users (mengandung data profil lengkap)
    const { sheet: sh, error: sheetErr } = getUsersSheetByConfig_();
    if (!sh) return { ok:false, msg: sheetErr || 'Sheet Users tidak ditemukan.' };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { ok:false, msg:'Sheet Users kosong atau tidak ada data.' };

    // Baca header (row 1)
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h||'').trim());
    const headerMap = buildHeaderMap_(headers);
    const idxNip = findHeaderIdx_(headerMap, ['NIP']); // 0-based
    const idxUserId = findHeaderIdx_(headerMap, ['USER_ID']);
    const idxEmail = findHeaderIdx_(headerMap, ['Email', 'EMAIL']);
    if (idxNip < 0 && idxUserId < 0 && idxEmail < 0) return { ok:false, msg:'Header "NIP", "USER_ID", atau "Email" tidak ditemukan di baris 1 sheet Users.' };

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
      alamat: sanitize(['Alamat']),
      kelurahan_desa: sanitize(['Kelurahan_Desa']),
      kecamatan: sanitize(['Kecamatan']),
      kabupaten_kota: sanitize(['Kabupaten_Kota']),
      kode_pos: sanitize(['Kode_Pos']),
      darurat: buildEmergency()
    },
    personal: {
      nik: sanitize(['NIK']),
      ttl: sanitizeValue_(ttl),
      jenis_kelamin: sanitize(['Jenis_Kelamin']),
      status_nikah: sanitize(['Status_Nikah', 'Status Nikah', 'STATUS NIKAH', 'STATUS PERNIKAHAN']),
      no_kk: sanitize(['No_KK']),
      ayah_kandung: sanitize(['Ayah_Kandung']),
      ibu_kandung: sanitize(['Ibu_Kandung']),
      gelar_akademik_depan: sanitize(['Gelar_Akademik_Depan']),
      gelar_akademik_belakang: sanitize(['Gelar_Akademik_Belakang']),
      bpjs_kes: sanitize(['BPJS_Kes', 'BPJS KESEHATAN', 'BPJS_KES']),
      bpjs_tk: sanitize(['BPJS_TK', 'BPJS Ketenagakerjaan', 'BPJS KETENAGAKERJAAN', 'BPJSTK']),
      status_ptkp: sanitize(['Status_PTKP']),
      no_rekening: sanitize(['No._Rekening']),
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
function debugProfilMasterdataSaya(payload) {
  try {
    const s = requireLogin_(payload.nip, payload.deviceId, payload.token);
    const nipSession = String(s.nip || '').trim();
    const nipKey = normalizeNIP_(nipSession);

    const { sheet: sh, error: sheetErr } = getUsersSheetByConfig_();
    if (!sh) return { ok:false, msg: sheetErr || 'Sheet Users tidak ditemukan.' };

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
      sheet: { name: CFG.SHEET_USERS, lastRow, lastCol },
      headerNIPIndex0: idxNip,
      headersPreview: headers.slice(0, 15),
      nipSamples: sample
    };

  } catch (e) {
    return { ok:false, msg:`Error debugProfil: ${e && e.message ? e.message : e}` };
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

function getProfilUsersDetail(payload) {
  try {
    const s = requireLogin_(payload.nip, payload.deviceId, payload.token);
    const nipSession = String(s.nip || '').trim();
    const userIdSession = String(s.userId || '').trim();
    
    // ✅ SECURITY: Validate session NIP
    validateSessionNip_(nipSession, payload.deviceId, payload.token);
    
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
        
        // Tambahkan data gaji dari sheet Slip Gaji
        const nip = txt(pickCell_(row, headerMap, ['NIP']));
        if (nip && nip !== '-') {
          const salaryData = getLatestSlipGajiForProfile_(nip);
          if (salaryData && salaryData.ok) {
            Object.assign(data, salaryData.data);
          }
        }
        
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
    const ss = SpreadsheetApp.getActive();

    const gidRaw = cfgGet('USERS_GID', '');
    const gid = Number(gidRaw);
    if (!isNaN(gid) && gid > 0) {
      const byId = ss.getSheets().find(sh => sh.getSheetId() === gid);
      if (byId) return { sheet: byId };
      return { sheet: null, error:`Sheet dengan GID ${gid} (USERS_GID) tidak ditemukan di spreadsheet aktif.` };
    }

    const sh = ss.getSheetByName(CFG.SHEET_USERS);
    if (sh) return { sheet: sh };
    return { sheet: null, error:`Sheet "${CFG.SHEET_USERS}" tidak ditemukan pada spreadsheet aktif.` };
  } catch (e) {
    const errMsg = e && e.message ? e.message : e;
    return { sheet: null, error:`Gagal membuka sheet Users: ${errMsg}` };
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

/**
 * Get latest slip gaji data for a given NIP (for profile display)
 * Returns salary fields for Rincian Gaji section
 */
function getLatestSlipGajiForProfile_(nip) {
  try {
    if (!nip || nip === '-') return { ok: false };
    
    const { sheet: sh, error: sheetErr } = getSlipGajiSheet_();
    if (!sh) return { ok: false, msg: sheetErr };
    
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2) return { ok: false };
    
    const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const headersRow = values[0].map(h => String(h || '').trim());
    const headerMap = buildHeaderMap_(headersRow);
    const nipNormalized = normalizeNIP_(nip);
    
    let latestRow = null;
    let latestDate = null;
    
    // Find latest row by NIP (compare tanggal untuk mendapat yang terbaru)
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowNipKey = normalizeNIP_(pickCell_(row, headerMap, ['NIP']));
      if (rowNipKey && rowNipKey === nipNormalized) {
        latestRow = row;
        break; // Assuming data is sorted by date DESC, take first match
      }
    }
    
    if (!latestRow) return { ok: false };
    
    // Build salary payload from row
    const getRaw = (names) => pickCell_(latestRow, headerMap, Array.isArray(names) ? names : [names]);
    const getNum = (names) => {
      const v = getRaw(names);
      const num = Number(String(v || '').replace(/[^\d.-]/g, ''));
      return isNaN(num) ? 0 : num;
    };
    
    const data = {
      gajiPokok: getNum(['GAJI POKOK', 'Gaji Pokok']),
      tunjanganKinerja: getNum(['TUNJANGAN KINERJA', 'Tunjangan Kinerja']),
      tunjIstri: getNum(['TUNJ. ISTRI', 'Tunj. Istri', 'TUNJ ISTRI']),
      tunjAnak: getNum(['TUNJ. ANAK', 'Tunj. Anak', 'TUNJ ANAK']),
      tunjFungsional: getNum(['TUNJ. FUNGSIONAL', 'Tunj. Fungsional', 'TUNJ FUNGSIONAL']),
      tunjJabatan: getNum(['TUNJ. JABATAN', 'Tunj. Jabatan', 'TUNJ JABATAN']),
      tunjKualifikasiKhusus: getNum(['TUNJANGAN KUALIFIKASI KHUSUS', 'Tunjangan Kualifikasi Khusus']),
      tunjanganBpjs: getNum(['TUNJ. BPJS', 'Tunj. BPJS', 'TUNJ BPJS']),
      lembur: getNum(['LEMBUR', 'Lembur']),
      rapelGaji: getNum(['RAPEL GAJI', 'Rapel Gaji']),
      potKasbon: getNum(['POTONGAN KASBON', 'Potongan Kasbon']),
      bpjs: getNum(['BPJS', 'BPJS Potongan']),
      pendidikanAnak: getNum(['PENDIDIKAN ANAK', 'Pendidikan Anak']),
      kekuranganJam: getNum(['Kekurangan Jam', 'KEKURANGAN JAM']),
      bpjsJht: getNum(['BPJS TK (JHT)', 'BPJS TK JHT', 'BPJS JHT']),
      bpjsJp: getNum(['BPJS TK (JP)', 'BPJS TK JP', 'BPJS JP']),
      pph21: getNum(['PPH21', 'PPH 21']),
      potAbsensi: getNum(['POTONGAN ABSENSI', 'Potongan Absensi', 'POT. ABSENSI']),
      kinerjaAnnual: txt(getRaw(['KINERJA TAHUNAN', 'Kinerja Tahunan'])),
      kinerjaMonthly: txt(getRaw(['KINERJA BULANAN', 'Kinerja Bulanan'])),
      jumlahJam: txt(getRaw(['Jumlah Jam', 'JUMLAH JAM'])),
      masaBekerja: txt(getRaw(['MASA BEKERJA', 'Masa Bekerja'])),
      statusKepegawaian: txt(getRaw(['STATUS KEPEGAWAIAN', 'Status Kepegawaian'])),
      pendidikanTerakhir: txt(getRaw(['PENDIDIKAN TERAKHIR', 'Pendidikan Terakhir'])),
      suamiIstri: txt(getRaw(['SUAMI/ ISTRI', 'Suami/ Istri', 'SUAMI/ISTRI'])),
      anak: txt(getRaw(['ANAK', 'Anak'])),
      tanggal: txt(getRaw(['Tanggal', 'TANGGAL']))
    };
    
    return { ok: true, data };
  } catch (e) {
    return { ok: false, msg: `Error getLatestSlipGajiForProfile_: ${e && e.message ? e.message : e}` };
  }
}

/**
 * Get slip gaji sheet from config
 */
function getSlipGajiSheet_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const gidRaw = cfgGet('SLIP_GAJI_GID', '');
    const gid = Number(gidRaw);
    if (!isNaN(gid) && gid > 0) {
      const byId = ss.getSheets().find(sh => sh.getSheetId() === gid);
      if (byId) return { sheet: byId };
    }
    const sh = ss.getSheetByName('Slip_Gaji');
    if (sh) return { sheet: sh };
    return { sheet: null, error: `Sheet slip gaji tidak ditemukan` };
  } catch (e) {
    return { sheet: null, error: `Error getSlipGajiSheet_: ${e && e.message ? e.message : e}` };
  }
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
      alamat: txtVal(['Alamat']),
      kelurahan_desa: txtVal(['Kelurahan_Desa']),
      kecamatan: txtVal(['Kecamatan']),
      kabupaten_kota: txtVal(['Kabupaten_Kota']),
      kode_pos: txtVal(['Kode_Pos']),
      darurat: {
        nama: txtVal(['Darurat_Nama', 'KontakDarurat_Nama']),
        hp: txtVal(['Darurat_HP', 'KontakDarurat_HP', 'Darurat_WA']),
        hubungan: txtVal(['Darurat_Hubungan', 'KontakDarurat_Hubungan'])
      }
    },
    personal: {
      nik: txtVal(['NIK']),
      ttl: buildTTL_(txtVal(['Tempat_Lahir', 'Tempat Lahir']), get(['Tanggal_Lahir', 'Tanggal Lahir'])),
      jenis_kelamin: txtVal(['Jenis_Kelamin']),
      status_nikah: txtVal(['Status_Nikah', 'Status Nikah']),
      no_kk: txtVal(['No_KK']),
      ayah_kandung: txtVal(['Ayah_Kandung']),
      ibu_kandung: txtVal(['Ibu_Kandung']),
      gelar_akademik_depan: txtVal(['Gelar_Akademik_Depan']),
      gelar_akademik_belakang: txtVal(['Gelar_Akademik_Belakang']),
      bpjs_kes: txtVal(['BPJS_Kes']),
      bpjs_tk: txtVal(['BPJS_TK', 'BPJS Ketenagakerjaan']),
      status_ptkp: txtVal(['Status_PTKP']),
      no_rekening: txtVal(['No._Rekening']),
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
/*************************************************
 * HISTORI MUTASI - Pencatatan Perubahan Field
 *************************************************/

/**
 * Definisi field sensitif
 */
function getSensitiveFields_() {
  return ['NIK', 'No._Rekening'];
}

/**
 * Generate UUID v4
 */
function generateUUID_() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

/**
 * Catat perubahan field ke sheet Histori Mutasi
 * @param {object} params - { target_nip, target_nama, field_key, field_label, old_value, new_value, changed_by_nip, changed_by_nama, actor_role, consent_checked, reason }
 * @returns {object} { ok, msg, mutasi_id }
 */
function logProfilMutation_(params) {
  try {
    const {
      target_nip,
      target_nama,
      field_key,
      field_label,
      old_value,
      new_value,
      changed_by_nip,
      changed_by_nama,
      actor_role,
      consent_checked,
      reason
    } = params;

    // Validasi input
    if (!target_nip || !field_key || !changed_by_nip) {
      return { ok: false, msg: 'Parameter tidak lengkap (target_nip, field_key, changed_by_nip wajib)' };
    }

    const { sheet: sh, error: sheetErr } = getHistoriMutasiSheet_();
    if (!sh) {
      return { ok: false, msg: sheetErr || 'Sheet Histori_Mutasi tidak ditemukan' };
    }

    // Pastikan header
    ensureHistoriMutasiHeader_(sh);

    // Generate Mutasi_ID & Timestamp
    const mutasi_id = generateUUID_();
    const timestamp = new Date().toISOString();

    // Append record
    const newRow = [
      mutasi_id,                          // Mutasi_ID
      timestamp,                          // Timestamp
      target_nip,                         // Target_NIP
      target_nama || '',                  // Target_Nama
      field_key,                          // Field_Key
      field_label || field_key,           // Field_Label
      String(old_value || ''),            // Old_Value
      String(new_value || ''),            // New_Value
      changed_by_nip,                     // Changed_By_NIP
      changed_by_nama || '',              // Changed_By_Nama
      actor_role || 'pegawai',            // Actor_Role
      'profil_edit',                      // Change_Source
      reason || '',                       // Reason
      consent_checked ? 'TRUE' : 'FALSE', // Consent_Checked
      '',                                 // Client_Info
      ''                                  // Request_ID/Trace_ID
    ];

    sh.appendRow(newRow);

    return { ok: true, msg: 'Perubahan dicatat ke histori', mutasi_id };
  } catch (e) {
    return { ok: false, msg: `Error logProfilMutation_: ${e && e.message ? e.message : e}` };
  }
}

/**
 * Pastikan header di sheet Histori_Mutasi
 */
function ensureHistoriMutasiHeader_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const headerText = headers.map(h => String(h || '').trim());

  const expected = [
    'Mutasi_ID',
    'Timestamp',
    'Target_NIP',
    'Target_Nama',
    'Field_Key',
    'Field_Label',
    'Old_Value',
    'New_Value',
    'Changed_By_NIP',
    'Changed_By_Nama',
    'Actor_Role',
    'Change_Source',
    'Reason',
    'Consent_Checked',
    'Client_Info',
    'Request_ID'
  ];

  const headerOk = expected.length === headerText.length &&
    expected.every((exp, i) => String(headerText[i] || '').trim() === exp);

  if (!headerOk) {
    sh.getRange(1, 1, 1, expected.length).setValues([expected]);
    sh.setFrozenRows(1);
  }
}

/**
 * Simpan perubahan field profil dan catat ke histori
 * @param {object} payload - { field_key, field_label, old_value, new_value, consent_checked, deviceId, nip, token }
 * @returns {object} { ok, msg, mutasi_id }
 */
function saveProfilFieldChange(payload) {
  try {
    const s = requireLogin_(payload.nip, payload.deviceId, payload.token);
    const nipSession = String(s.nip || '').trim();
    const namaSession = String(s.nama || '').trim();
    
    if (!nipSession) {
      return { ok: false, msg: 'Session tidak valid (NIP tidak ditemukan)' };
    }

    const { field_key, field_label, old_value, new_value, consent_checked } = payload;

    // Validasi
    if (!field_key || new_value === undefined) {
      return { ok: false, msg: 'field_key dan new_value wajib' };
    }

    const sensitiveFields = getSensitiveFields_();
    const isSensitive = sensitiveFields.includes(field_key);

    // Jika field sensitif, pastikan consent
    if (isSensitive && !consent_checked) {
      return { ok: false, msg: 'Consent wajib untuk field sensitif' };
    }

    // Cari row di Users sheet dan update field
    const { sheet: sh, error: sheetErr } = getUsersSheetByConfig_();
    if (!sh) {
      return { ok: false, msg: sheetErr || 'Sheet Users tidak ditemukan' };
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) {
      return { ok: false, msg: 'Sheet Users kosong' };
    }

    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const headerMap = buildHeaderMap_(headers);
    const idxFieldKey = findHeaderIdx_(headerMap, [field_key]);

    if (idxFieldKey < 0) {
      return { ok: false, msg: `Field "${field_key}" tidak ditemukan di header` };
    }

    // Cari row pegawai (match NIP)
    const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const idxNip = findHeaderIdx_(headerMap, ['NIP']);

    let foundRowNum = -1;
    for (let i = 1; i < values.length; i++) {
      const nipCell = idxNip >= 0 ? normalizeNIP_(values[i][idxNip]) : '';
      if (nipCell === normalizeNIP_(nipSession)) {
        foundRowNum = i + 1; // row number (1-based)
        break;
      }
    }

    if (foundRowNum < 0) {
      return { ok: false, msg: 'Data pegawai tidak ditemukan di sheet Users' };
    }

    // Update field value
    sh.getRange(foundRowNum, idxFieldKey + 1).setValue(new_value);

    // Catat ke histori
    const histRes = logProfilMutation_({
      target_nip: nipSession,
      target_nama: namaSession,
      field_key: field_key,
      field_label: field_label || field_key,
      old_value: old_value,
      new_value: new_value,
      changed_by_nip: nipSession,
      changed_by_nama: namaSession,
      actor_role: 'pegawai',
      consent_checked: isSensitive && consent_checked,
      reason: ''
    });

    if (!histRes.ok) {
      // Log dicatat, tapi return success karena data sudah disimpan
      return {
        ok: true,
        msg: 'Perubahan disimpan (histori: ' + (histRes.msg || 'dicatat') + ')',
        mutasi_id: histRes.mutasi_id
      };
    }

    return {
      ok: true,
      msg: 'Perubahan disimpan dan dicatat dalam histori',
      mutasi_id: histRes.mutasi_id
    };

  } catch (e) {
    return { ok: false, msg: `Error saveProfilFieldChange: ${e && e.message ? e.message : e}` };
  }
}
