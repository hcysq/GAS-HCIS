/*************************************************
 * Dashboard
 *************************************************/

function getSaldoCutiSaya() {
  try {
    const s = requireLogin_();
    const nip = s.nip;

    const t = readTable_(CFG.SHEET_SALDO);
    const h = t.headers;
    const r = t.rows;

    const cNIP = col_(h, 'NIP');
    const cTahun = col_(h, 'Tahun');
    const cJatah = col_(h, 'Jatah');
    const cTerpakai = col_(h, 'Terpakai');
    const cSisa = col_(h, 'Sisa');

    let latest = null;
    for (const row of r) {
      if (txt(row[cNIP]) !== nip) continue;
      if (!latest || row[cTahun] > latest.Tahun) {
        latest = {
          Tahun: row[cTahun],
          Jatah: row[cJatah],
          Terpakai: row[cTerpakai],
          Sisa: row[cSisa]
        };
      }
    }

    return { ok:true, data: latest };
  } catch (e) {
    return { ok:false };
  }
}
