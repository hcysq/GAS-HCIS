/*************************************************
 * Profile (Masterdata) - Robust + Debug
 *************************************************/

function getProfilMasterdataSaya() {
  try {
    const s = requireLogin_();
    const nipSession = String(s.nip || '').trim();
    const nipKey = normalizeNIP_(nipSession);
    if (!nipKey) return { ok:false, msg:'Session NIP kosong. Coba logout lalu login ulang.' };

    // Pastikan sheet/tab Masterdata benar-benar ada (mendukung konfigurasi GID)
    const { sheet: sh, error: sheetErr } = getMasterdataSheet_();
    if (!sh) return { ok:false, msg: sheetErr || 'Sheet Masterdata tidak ditemukan.' };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { ok:false, msg:'Sheet Masterdata kosong atau tidak ada data.' };

    // Baca header (row 1)
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h||'').trim());
    const idxNip = findHeaderIndex_(headers, 'NIP'); // 0-based
    if (idxNip < 0) return { ok:false, msg:'Header "NIP" tidak ditemukan di baris 1 sheet Masterdata.' };

    // Baca data rows (row 2..last)
    const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    for (const row of rows) {
      const cellKey = normalizeNIP_(row[idxNip]);
      if (cellKey && cellKey === nipKey) {
        const obj = {};
        headers.forEach((k, i) => obj[k] = row[i]);
        return { ok:true, data: obj };
      }
    }

    return { ok:false, msg:`Profil tidak ketemu. NIP session=${nipSession} (key=${nipKey}). Cek apakah NIP di Masterdata benar-benar sama.` };

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
