/*************************************************
 * Cuti
 *************************************************/

function getApproverByNIP(nip) {
  const t = readTable_(CFG.SHEET_ATASAN);
  const h = t.headers;
  const r = t.rows;

  const cNIP = col_(h, 'NIP');
  const cApp = col_(h, 'ApproverNIP');
  const cAktif = col_(h, 'Aktif');

  for (const row of r) {
    if (txt(row[cNIP]) === nip && isTrue_(row[cAktif])) {
      return txt(row[cApp]);
    }
  }
  return '';
}

function submitCuti(data) {
  try {
    const s = requireLogin_();
    const sh = getSheet_(CFG.SHEET_CUTI);

    sh.appendRow([
      Utilities.getUuid(),
      new Date(),
      s.email,
      s.nip,
      s.nama,
      data.jenis,
      data.satuan,
      data.tglMulai,
      data.tglSelesai,
      '',
      '',
      '',
      data.alasan,
      '',
      'DIAJUKAN',
      getApproverByNIP(s.nip),
      '',
      '',
      '',
      new Date()
    ]);

    return { ok:true };
  } catch (e) {
    return { ok:false };
  }
}
