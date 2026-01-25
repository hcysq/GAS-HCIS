/*************************************************
 * WELFARE MODULE - Slip Gaji & Kesejahteraan
 *************************************************/

/**
 * Ambil daftar Tahun & Bulan yang tersedia untuk user
 * @returns {object} { ok, data: { tahunList, bulanPerTahun }, msg }
 */
function getAvailableSlipGajiBulan() {
  try {
    // Pastikan user login
    const s = requireLogin_();
    const nipUser = String(s.nip || '').trim();
    if (!nipUser) {
      return { ok: false, msg: 'NIP user tidak ditemukan dalam session' };
    }

    // Buka sheet Slip Gaji
    const { sheet: sh, error: sheetErr } = getSlipGajiSheet_();
    if (!sh) return { ok: false, msg: sheetErr || 'Sheet Slip Gaji tidak ditemukan' };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { ok: false, msg: 'Sheet Slip Gaji kosong' };

    // Baca header
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
    const headerMap = buildHeaderMap_(headers);

    const idxNip = findHeaderIdx_(headerMap, ['NIP']);
    const idxBulan = findHeaderIdx_(headerMap, ['Bulan']);

    if (idxNip < 0 || idxBulan < 0) {
      return { ok: false, msg: 'Header "NIP" atau "Bulan" tidak ditemukan' };
    }

    // Baca semua data dan kumpulkan tahun/bulan untuk user ini
    const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const periodeSet = new Set();

    for (let i = 0; i < values.length; i++) {
      const nipCell = String(values[i][idxNip] || '').trim();
      const bulanCell = String(values[i][idxBulan] || '').trim();

      if (nipCell === nipUser && bulanCell) {
        periodeSet.add(bulanCell);
      }
    }

    if (periodeSet.size === 0) {
      return { ok: false, msg: 'Tidak ada data slip gaji untuk user ini' };
    }

    // Parse "NamaBulan Tahun" menjadi list tahun & bulan
    const periodeList = Array.from(periodeSet).sort();
    const tahunSet = new Set();
    const bulanPerTahun = {};

    periodeList.forEach(periode => {
      const parts = periode.trim().split(/\s+/);
      if (parts.length >= 2) {
        const tahun = parts[parts.length - 1];
        const namaBulan = parts.slice(0, -1).join(' ');
        
        tahunSet.add(tahun);
        if (!bulanPerTahun[tahun]) bulanPerTahun[tahun] = [];
        
        // Jangan duplikat
        if (!bulanPerTahun[tahun].includes(namaBulan)) {
          bulanPerTahun[tahun].push(namaBulan);
        }
      }
    });

    const tahunList = Array.from(tahunSet).sort();
    
    return { ok: true, data: { tahunList, bulanPerTahun, periodeList } };

  } catch (e) {
    return { ok: false, msg: `Error getAvailableSlipGajiBulan: ${e && e.message ? e.message : e}` };
  }
}

/**
 * Ambil data Slip Gaji berdasarkan NIP user + Tahun + Bulan
 * @param {number} tahun - Tahun (YYYY)
 * @param {number} bulan - Bulan (1-12) atau nama bulan Indonesia
 * @returns {object} { ok, data, msg }
 */
function getSlipGaji(tahun, bulan) {
  try {
    // Validasi input
    tahun = Number(tahun);
    if (isNaN(tahun) || tahun < 2000 || tahun > 2099) {
      return { ok: false, msg: 'Tahun tidak valid' };
    }

    // Bulan bisa number (1-12) atau string (nama bulan Indonesia)
    let bulanName = '';
    if (typeof bulan === 'number' || !isNaN(Number(bulan))) {
      const bulanNum = Number(bulan);
      if (bulanNum < 1 || bulanNum > 12) {
        return { ok: false, msg: 'Bulan tidak valid' };
      }
      bulanName = getBulanIndonesia_(bulanNum);
    } else {
      bulanName = String(bulan).trim();
    }

    // Pastikan user login
    const s = requireLogin_();
    const nipUser = String(s.nip || '').trim();
    if (!nipUser) {
      return { ok: false, msg: 'NIP user tidak ditemukan dalam session' };
    }

    // Buka sheet Slip Gaji
    const { sheet: sh, error: sheetErr } = getSlipGajiSheet_();
    if (!sh) return { ok: false, msg: sheetErr || 'Sheet Slip Gaji tidak ditemukan' };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { ok: false, msg: 'Sheet Slip Gaji kosong atau belum ada data' };

    // Baca header (row 1)
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
    const headerMap = buildHeaderMap_(headers);

    // Cari index kolom penting
    const idxNip = findHeaderIdx_(headerMap, ['NIP']);
    const idxBulan = findHeaderIdx_(headerMap, ['Bulan']);
    const idxTanggal = findHeaderIdx_(headerMap, ['Tanggal']);

    if (idxNip < 0 || idxBulan < 0) {
      return { ok: false, msg: 'Header "NIP" atau "Bulan" tidak ditemukan di sheet Slip Gaji' };
    }

    // Form periode target: "NamaBulan Tahun"
    const periodTarget = `${bulanName} ${tahun}`;

    // Baca semua data (row 2..last)
    const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    let matchedRows = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const nipCell = String(row[idxNip] || '').trim();
      const bulanCell = String(row[idxBulan] || '').trim();

      // Cek NIP cocok dan Bulan cocok
      if (nipCell === nipUser && bulanCell === periodTarget) {
        matchedRows.push({
          index: i,
          row: row,
          tanggal: idxTanggal >= 0 ? row[idxTanggal] : null
        });
      }
    }

    if (matchedRows.length === 0) {
      return { ok: false, msg: `Slip gaji untuk periode ${periodTarget} tidak ditemukan` };
    }

    // Jika lebih dari 1 baris, ambil yang paling terbaru
    let selected = matchedRows[0];
    if (matchedRows.length > 1 && idxTanggal >= 0) {
      selected = matchedRows.reduce((latest, curr) => {
        try {
          const currDate = new Date(curr.tanggal);
          const latestDate = new Date(latest.tanggal);
          return currDate > latestDate ? curr : latest;
        } catch (e) {
          return latest;
        }
      });
    }

    // Build payload
    const payload = buildSlipGajiPayload_(selected.row, headers, headerMap);
    return { ok: true, data: payload };

  } catch (e) {
    return { ok: false, msg: `Error getSlipGaji: ${e && e.message ? e.message : e}` };
  }
}

/**
 * Helper: Buka sheet Slip Gaji berdasarkan config SLIP_GAJI_GID
 */
function getSlipGajiSheet_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const gidRaw = cfgGet('SLIP_GAJI_GID', '');
    const gid = Number(gidRaw);

    if (!isNaN(gid) && gid > 0) {
      const byId = ss.getSheets().find(sh => sh.getSheetId() === gid);
      if (byId) return { sheet: byId };
      return { sheet: null, error: `Sheet dengan GID ${gid} (SLIP_GAJI_GID) tidak ditemukan` };
    }

    const sh = ss.getSheetByName('Slip_Gaji');
    if (sh) return { sheet: sh };
    return { sheet: null, error: `Sheet "Slip_Gaji" tidak ditemukan` };
  } catch (e) {
    return { sheet: null, error: `Gagal membuka sheet Slip Gaji: ${e && e.message ? e.message : e}` };
  }
}

/**
 * Konversi bulan (1-12) ke nama bulan Indonesia
 */
function getBulanIndonesia_(bulan) {
  const bulanList = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];
  return bulanList[bulan - 1] || '';
}

/**
 * Build payload slip gaji dari row data
 * Lebih robust dengan multiple column name options
 */
function buildSlipGajiPayload_(row, headers, headerMap) {
  const getRaw = (names) => pickCell_(row, headerMap, Array.isArray(names) ? names : [names]);
  const getText = (names) => txt(getRaw(names));
  const getNum = (names) => {
    const v = getRaw(names);
    const num = Number(String(v || '').replace(/[^\d.-]/g, ''));
    return isNaN(num) ? 0 : num;
  };
  const sanitize = (val) => {
    const t = txt(val);
    return t && t !== '0' ? t : '-';
  };

  // Build jabatan dengan multiple options
  const jabatanParts = [
    getRaw('JABATAN'),
    getRaw('JABATAN STRUKTURAL'),
    getRaw('JABATAN FUNGSIONAL')
  ].filter(v => txt(v)).map(v => txt(v));
  const jabatan = jabatanParts.length > 0 ? jabatanParts.join(' / ') : '-';

  return {
    // Identitas
    nama: sanitize(getRaw(['NAMA', 'Nama'])),
    nip: sanitize(getRaw(['NIP', 'Nip'])),
    unit: sanitize(getRaw(['UNIT', 'Unit'])),
    jabatan: jabatan,
    jabatanFungsional: sanitize(getRaw(['JABATAN FUNGSIONAL', 'Jabatan Fungsional'])),
    jabatanStruktural: sanitize(getRaw(['JABATAN STRUKTURAL', 'Jabatan Struktural'])),

    // Angka Utama - try multiple column name variations
    gajiNeto: getNum(['GAJI NETO', 'Gaji Neto', 'GAJI NETO 80%', 'Gaji Netto 80%']),
    gajiBruto_80: getNum(['GAJI NETO 80%', 'Gaji Netto 80%', 'GAJI NETO']),
    totalBruto: getNum(['TOTAL BRUTO GAJI', 'Total Bruto Gaji', 'TOTAL BRUTO']),
    totalPotongan: getNum(['TOTAL POTONGAN', 'Total Potongan']),
    gajiProrata: getNum(['GAJI PRORATA', 'Gaji Prorata']),

    // Pendapatan (Earnings)
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

    // Potongan (Deductions)
    potKasbon: getNum(['POTONGAN KASBON', 'Potongan Kasbon']),
    bpjs: getNum(['BPJS', 'BPJS Potongan']),
    pendidikanAnak: getNum(['PENDIDIKAN ANAK', 'Pendidikan Anak']),
    kekuranganJam: getNum(['Kekurangan Jam', 'KEKURANGAN JAM']),
    bpjsJht: getNum(['BPJS TK (JHT)', 'BPJS TK JHT', 'BPJS JHT']),
    bpjsJp: getNum(['BPJS TK (JP)', 'BPJS TK JP', 'BPJS JP']),
    pph21: getNum(['PPH21', 'PPH 21']),
    potAbsensi: getNum(['POTONGAN ABSENSI', 'Potongan Absensi', 'POT. ABSENSI']),

    // Detail Lainnya
    kinerjaAnnual: sanitize(getRaw(['KINERJA TAHUNAN', 'Kinerja Tahunan'])),
    kinerjaMonthly: sanitize(getRaw(['KINERJA BULANAN', 'Kinerja Bulanan'])),
    jumlahJam: sanitize(getRaw(['Jumlah Jam', 'JUMLAH JAM'])),
    masaBekerja: sanitize(getRaw(['MASA BEKERJA', 'Masa Bekerja'])),
    statusKepegawaian: sanitize(getRaw(['STATUS KEPEGAWAIAN', 'Status Kepegawaian'])),
    pendidikanTerakhir: sanitize(getRaw(['PENDIDIKAN TERAKHIR', 'Pendidikan Terakhir'])),
    suamiIstri: sanitize(getRaw(['SUAMI/ ISTRI', 'Suami/ Istri', 'SUAMI/ISTRI'])),
    anak: sanitize(getRaw(['ANAK', 'Anak'])),
    tanggal: sanitize(getRaw(['Tanggal', 'TANGGAL']))
  };
}

/**
 * Format nominal ke Rupiah: "Rp 1.234.567"
 * @param {number} value - nilai numerik
 * @returns {string} format Rupiah
 */
function formatCurrencyRupiah_(value) {
  const num = Number(value) || 0;
  if (num === 0) return 'Rp 0';
  return 'Rp ' + new Intl.NumberFormat('id-ID', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0
  }).format(num);
}

/**
 * Aturan tampilan GAJI UTAMA (Neto / 80% / Prorata)
 * Prioritas:
 *  1. Jika GAJI PRORATA > 0 -> "Rp X (Prorata)"
 *  2. Else jika GAJI NETTO 80% > 0 -> "Rp X (80%)"
 *  3. Else -> "Rp X"
 * @param {object} data - slip data
 * @returns {string} formatted gaji utama
 */
function applySalaryDisplayRule_(data) {
  const prorata = Number(data.gajiProrata) || 0;
  const netto80 = Number(data.gajiBruto_80) || 0; // asumsi ada di data
  const netto = Number(data.gajiNeto) || 0;

  if (prorata > 0) {
    return formatCurrencyRupiah_(prorata) + ' (Prorata)';
  } else if (netto80 > 0) {
    return formatCurrencyRupiah_(netto80) + ' (80%)';
  } else {
    return formatCurrencyRupiah_(netto);
  }
}

/**
 * Gabung field JABATAN dari multiple columns
 * Columns: JABATAN, JABATAN FUNGSIONAL, JABATAN STRUKTURAL
 * @param {object} data - slip data
 * @returns {string} jabatan gabungan
 */
function extractJabatan_(data) {
  const parts = [data.jabatan, data.jabatanFungsional, data.jabatanStruktural]
    .map(v => String(v || '').trim())
    .filter(v => v && v !== '-' && v !== '');
  return parts.length > 0 ? parts.join(' / ') : '-';
}

/**
 * Cek apakah file slip sudah ada di folder
 * @param {string} fileName - nama file yang akan dicek (SlipGaji_YYYYMM_NIP.pdf)
 * @param {string} folderId - folder ID Drive
 * @returns {boolean} true jika sudah ada
 */
function checkSlipFileExists_(fileName, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(fileName);
    return files.hasNext();
  } catch (e) {
    Logger.log('Error checking slip file: ' + e.message);
    return false;
  }
}

/**
 * Build object placeholder replacements untuk Google Docs template
 * @param {object} data - slip data
 * @param {string} periode - "Januari 2026"
 * @returns {object} { placeholder: value }
 */
function buildPlaceholderReplacements_(data, periode) {
  const jabatan = extractJabatan_(data);
  const gajiUtama = applySalaryDisplayRule_(data);

  return {
    '{{PERIODE}}': periode,
    '{{NAMA}}': data.nama || '-',
    '{{NIP}}': data.nip || '-',
    '{{UNIT}}': data.unit || '-',
    '{{JABATAN}}': jabatan,
    '{{TOTAL_BRUTO}}': formatCurrencyRupiah_(data.totalBruto),
    '{{TOTAL_POTONGAN}}': formatCurrencyRupiah_(data.totalPotongan),
    '{{GAJI_NETO}}': gajiUtama,
    
    // Pendapatan
    '{{GAJI_POKOK}}': formatCurrencyRupiah_(data.gajiPokok),
    '{{TUNJ_KINERJA}}': formatCurrencyRupiah_(data.tunjanganKinerja),
    '{{TUNJ_ISTRI}}': formatCurrencyRupiah_(data.tunjIstri),
    '{{TUNJ_ANAK}}': formatCurrencyRupiah_(data.tunjAnak),
    '{{TUNJ_FUNGSIONAL}}': formatCurrencyRupiah_(data.tunjFungsional),
    '{{TUNJ_JABATAN}}': formatCurrencyRupiah_(data.tunjJabatan),
    '{{TUNJ_KUALIFIKASI}}': formatCurrencyRupiah_(data.tunjKualifikasiKhusus),
    '{{LEMBUR}}': formatCurrencyRupiah_(data.lembur),
    '{{RAPEL_GAJI}}': formatCurrencyRupiah_(data.rapelGaji),
    '{{TUNJ_BPJS}}': formatCurrencyRupiah_(data.tunjanganBpjs),
    
    // Potongan
    '{{POT_KASBON}}': formatCurrencyRupiah_(data.potKasbon),
    '{{BPJS}}': formatCurrencyRupiah_(data.bpjs),
    '{{PEND_ANAK}}': formatCurrencyRupiah_(data.pendidikanAnak),
    '{{KURANG_JAM}}': formatCurrencyRupiah_(data.kekuranganJam),
    '{{BPJS_JHT}}': formatCurrencyRupiah_(data.bpjsJht || 0),
    '{{BPJS_JP}}': formatCurrencyRupiah_(data.bpjsJp || 0),
    '{{PPH21}}': formatCurrencyRupiah_(data.pph21 || 0),
    '{{POT_ABSENSI}}': formatCurrencyRupiah_(data.potAbsensi || 0)
  };
}

/**
 * Generate dan Kirim Slip Gaji via PDF (dari Google Docs Template)
 * @param {number} tahun - tahun (YYYY)
 * @param {string|number} bulan - bulan (1-12 atau nama)
 * @returns {object} {ok, msg}
 */
function generateAndSaveSlipGajiPDF(tahun, bulan) {
  try {
    const s = requireLogin_();
    const nipUser = String(s.nip || '').trim();
    if (!nipUser) {
      return { ok: false, msg: 'NIP user tidak ditemukan' };
    }

    // Ambil data slip
    const slipRes = getSlipGaji(tahun, bulan);
    if (!slipRes.ok) {
      return { ok: false, msg: slipRes.msg || 'Gagal mengambil data slip' };
    }

    const data = slipRes.data;

    // Konversi bulan ke format YYYYMM untuk nama file
    const tahunNum = Number(tahun);
    let bulanNum = Number(bulan);
    if (isNaN(bulanNum)) {
      // Jika string bulan, konversi dari nama
      const bln = getBulanIndonesia_(1); // dummy untuk cek
      const months = ['januari', 'februari', 'maret', 'april', 'mei', 'juni', 
                      'juli', 'agustus', 'september', 'oktober', 'november', 'desember'];
      bulanNum = months.indexOf(String(bulan).toLowerCase()) + 1;
      if (bulanNum === 0) bulanNum = 1;
    }
    const yyyymm = String(tahunNum) + String(bulanNum).padStart(2, '0');
    const fileName = `SlipGaji_${yyyymm}_${nipUser}.pdf`;

    // Ambil folder dari config
    const folderId = cfgGet('FOLDER_SLIP', '');
    if (!folderId) {
      return { ok: false, msg: 'FOLDER_SLIP tidak dikonfigurasi' };
    }

    // CEK APAKAH FILE SUDAH ADA
    if (checkSlipFileExists_(fileName, folderId)) {
      return {
        ok: false,
        msg: 'Slip gaji periode ini sudah dibuat. Untuk permintaan ulang, silakan hubungi admin HCM.',
        alreadyExists: true
      };
    }

    // Ambil template dari config
    const templateId = cfgGet('TEMPLATE_SLIP', '');
    if (!templateId) {
      return { ok: false, msg: 'TEMPLATE_SLIP tidak dikonfigurasi' };
    }

    // Copy template
    const template = DriveApp.getFileById(templateId);
    const tempDoc = template.makeCopy('Slip_Temp_' + nipUser + '_' + Date.now());
    const docId = tempDoc.getId();

    try {
      // Buka document dan replace placeholder
      const doc = DocumentApp.openById(docId);
      const body = doc.getBody();

      // Format periode untuk display
      const bulanNama = typeof bulan === 'number' ? getBulanIndonesia_(bulan) : String(bulan).split(' ')[0];
      const periode = bulanNama + ' ' + tahunNum;

      // Build replacements
      const replacements = buildPlaceholderReplacements_(data, periode);

      // Replace semua placeholder di document
      for (const [placeholder, value] of Object.entries(replacements)) {
        body.replaceText(placeholder, String(value));
      }

      doc.saveAndClose();

      // Export ke PDF
      const pdfBlob = DocumentApp.openById(docId).getAs('application/pdf');
      pdfBlob.setName(fileName);

      // Simpan ke folder
      const folder = DriveApp.getFolderById(folderId);
      const pdfFile = folder.createFile(pdfBlob);
      pdfFile.setName(fileName);

      // Get PDF link sebelum sharing (untuk dipass ke notification function)
      const pdfLink = pdfFile.getUrl();
      Logger.log('PDF File created: ' + fileName + ', Link: ' + pdfLink);

      // Format periode untuk notifikasi
      const bulanNamaNotif = typeof bulan === 'number' ? getBulanIndonesia_(bulan) : String(bulan).split(' ')[0];
      const periodeNotif = bulanNamaNotif + ' ' + tahunNum;

      // Kirim notifikasi dan share file: email ke pegawai + WA ke pegawai + WA ke admin
      sendSlipGajiNotifications_(nipUser, data.nama, periodeNotif, pdfLink, pdfFile);

      // Delete temp document
      tempDoc.setTrashed(true);

      return {
        ok: true,
        msg: '✅ Slip berhasil dibuat. Silakan cek Email (folder inbox atau spam) dan WhatsApp anda.'
      };
    } catch (docError) {
      // Clean up temp doc
      try { tempDoc.setTrashed(true); } catch (e) {}
      throw docError;
    }

  } catch (e) {
    Logger.log('Error generateAndSaveSlipGajiPDF: ' + e.message);
    return { ok: false, msg: 'Error: ' + e.message };
  }
}

/**
 * Get user contact info (email & WA) from Users sheet by NIP
 */
function getUserContactInfo_(nip) {
  try {
    const nipNormalized = normalizeNIP_(nip);
    if (!nipNormalized) return null;

    const ss = SpreadsheetApp.getActive();
    
    // Get Users sheet
    const gidRaw = cfgGet('USERS_GID', '');
    const gid = Number(gidRaw);
    let sh = null;
    
    if (!isNaN(gid) && gid > 0) {
      sh = ss.getSheets().find(s => s.getSheetId() === gid);
    }
    if (!sh) {
      sh = ss.getSheetByName('Users');
    }
    if (!sh) return null;

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2) return null;

    const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const headersRow = values[0].map(h => String(h || '').trim());
    const headerMap = buildHeaderMap_(headersRow);

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowNip = normalizeNIP_(pickCell_(row, headerMap, ['NIP']));
      if (rowNip && rowNip === nipNormalized) {
        const email = txt(pickCell_(row, headerMap, ['Email', 'EMAIL']));
        const wa = txt(pickCell_(row, headerMap, ['No_HP', 'No_WA', 'WA', 'WhatsApp']));
        const nama = txt(pickCell_(row, headerMap, ['Nama', 'NAMA']));
        return { email: email || '', wa: wa || '', nama: nama || '' };
      }
    }

    return null;
  } catch (e) {
    Logger.log('Error getUserContactInfo_: ' + e.message);
    return null;
  }
}

/**
 * Send slip gaji notifications: Email to employee + WA to employee + WA to admin
 */
function sendSlipGajiNotifications_(nip, nama, periode, pdfLink, pdfFile) {
  try {
    // Get contact info dari Users sheet
    const contact = getUserContactInfo_(nip);
    if (!contact) {
      Logger.log('Warning: Tidak dapat menemukan info kontak untuk NIP ' + nip);
      return;
    }

    const email = contact.email;
    const wa = contact.wa;

    Logger.log('=== Slip Gaji Notification Start ===');
    Logger.log('NIP: ' + nip + ', Email: ' + email + ', WA: ' + wa);

    // SHARE PDF FILE KE EMAIL PEGAWAI (VIEWER ONLY)
    if (email && pdfFile) {
      try {
        pdfFile.addViewer(email);
        Logger.log('PDF shared (viewer) to: ' + email);
      } catch (e) {
        Logger.log('ERROR: Gagal share PDF ke ' + email + ': ' + e.message);
      }
    }

    // 1. Kirim email ke pegawai dengan template resmi
    if (email) {
      // Parse periode: ambil bulan dan tahun terpisah
      const periodeParts = String(periode || '').trim().split(' ');
      const bulanText = periodeParts[0] || 'Januari';
      const tahunText = periodeParts[1] || '2026';

      const subject = `Slip Gaji [${bulanText} ${tahunText}] [${nip}]`;
      
      const body = `Slip Gaji ${bulanText} ${tahunText}\n\n` +
        `Bismillahirrahmanirrahim.\n\n` +
        `Yth. ${nama},\n\n` +
        `Sistem telah berhasil memproses permintaan cetak Slip Gaji Anda untuk periode: 🗓️ ${bulanText} ${tahunText}\n\n` +
        `Silakan akses atau unduh dokumen melalui tautan berikut: 🔗 ${pdfLink}\n\n` +
        `Catatan Penting:\n` +
        `1. Dokumen ini bersifat RAHASIA & PRIBADI.\n` +
        `2. Mohon tidak membagikan tautan ini kepada pihak lain.\n` +
        `3. Segala risiko finansial atau hukum yang timbul akibat penggunaan dokumen ini menjadi tanggung jawab pribadi pegawai sepenuhnya.\n` +
        `Terima kasih.\n\n` +
        `======== 🤖 Pesan ini dibuat otomatis oleh Sistem HCIS Sabilul Qur'an`;

      try {
        GmailApp.sendEmail(
          email,
          subject,
          body
        );
        Logger.log('SUCCESS: Email sent to ' + email);
      } catch (e) {
        Logger.log('ERROR: Gagal kirim email ke ' + email + ': ' + e.message);
      }
    } else {
      Logger.log('WARNING: Email pegawai tidak ditemukan, skip pengiriman email');
    }

    // 2. Kirim WA ke pegawai
    if (wa) {
      const periodeParts = String(periode || '').trim().split(' ');
      const bulanText = periodeParts[0] || 'Januari';
      const tahunText = periodeParts[1] || '2026';
      
      const waMessage = `Slip Gaji ${bulanText} ${tahunText}\n\nBismillahirrahmanirrahim.\n\nYth. ${nama},\n\nSistem telah berhasil memproses permintaan cetak Slip Gaji Anda untuk periode: 🗓️ ${bulanText} ${tahunText}\n\nSilakan akses atau unduh dokumen melalui tautan berikut: 🔗 ${pdfLink}\n\nCatatan Penting:\n1. Dokumen ini bersifat RAHASIA & PRIBADI.\n2. Mohon tidak membagikan tautan ini kepada pihak lain.\n3. Segala risiko finansial atau hukum yang timbul akibat penggunaan dokumen ini menjadi tanggung jawab pribadi pegawai sepenuhnya.\n\nTerima kasih.`;
      try {
        sendWAViaStarsender_(wa, waMessage);
        Logger.log('SUCCESS: WA sent to ' + wa);
      } catch (e) {
        Logger.log('ERROR: Gagal kirim WA ke ' + wa + ': ' + e.message);
      }
    } else {
      Logger.log('WARNING: No WA number found for employee, skip WA to employee');
    }

    // 3. Kirim notifikasi ke admin via WA
    const adminWa = cfgGet('ADMIN_WA', '');
    if (adminWa) {
      const periodeParts = String(periode || '').trim().split(' ');
      const bulanText = periodeParts[0] || 'Januari';
      const tahunText = periodeParts[1] || '2026';
      
      const adminMessage = `[Slip Gaji] Periode: ${bulanText} ${tahunText}\nNIP: ${nip}\nNama: ${nama}\nStatus: Berhasil dibuat dan dikirim ke pegawai.`;
      try {
        sendWAViaStarsender_(adminWa, adminMessage);
        Logger.log('SUCCESS: Admin notification sent to WA');
      } catch (e) {
        Logger.log('ERROR: Gagal kirim notifikasi admin: ' + e.message);
      }
    } else {
      Logger.log('WARNING: ADMIN_WA not configured, skip admin notification');
    }

    Logger.log('=== Slip Gaji Notification Complete ===');
  } catch (e) {
    Logger.log('ERROR in sendSlipGajiNotifications_: ' + e.message);
  }
}

/**
 * Send message via Starsender WA API
 */
function sendWAViaStarsender_(waNumber, message) {
  try {
    const url = cfgRequireString('STARSENDER_URL');
    const apiKey = cfgRequireString('STARSENDER_APIKEY');
    const modeRaw = cfgGet('STARSENDER_MODE', '');
    const mode = String(modeRaw || '').trim().toLowerCase();

    // Format nomor: 62xxxxxxxxxx (tanpa +)
    const tujuan = String(waNumber || '').replace(/^\+/, '').replace(/[^\d]/g, '');
    if (!tujuan || tujuan.length < 10) {
      throw new Error('Format nomor WA tidak valid: ' + waNumber);
    }

    const headers = {};
    const payload = {
      tujuan: tujuan,
      message: message
    };

    if (mode === 'bearer') {
      headers.Authorization = `Bearer ${apiKey}`;
    } else if (mode === 'legacy_sendtext') {
      headers.apikey = apiKey;
    } else {
      headers.apikey = apiKey;
      payload.api_key = apiKey;
      const device = String(cfgGet('STARSENDER_DEVICE', '') || '').trim();
      if (device) payload.device = device;
    }

    const options = {
      method: 'post',
      headers,
      payload,
      muteHttpExceptions: true
    };

    const resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() >= 400) {
      throw new Error('API error: ' + resp.getResponseCode());
    }
  } catch (e) {
    throw new Error('Gagal kirim WA via Starsender: ' + e.message);
  }
}
