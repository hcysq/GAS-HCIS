/*************************************************
 * Authentication
 *************************************************/

function authLogin(nip, pin) {
  nip = txt(nip);
  pin = txt(pin);
  if (!nip || !pin) return { ok:false, msg:'NIP & PIN wajib diisi' };

  const t = readTable_(CFG.SHEET_USERS);
  const h = t.headers;
  const r = t.rows;

  const cNIP = col_(h, 'NIP');
  const cPIN = col_(h, 'PIN');
  const cAktif = col_(h, 'Aktif');
  const cNama = col_(h, 'Nama');
  const cRole = col_(h, 'Role');
  const cEmail = col_(h, 'Email');

  for (const row of r) {
    if (
      txt(row[cNIP]) === nip &&
      txt(row[cPIN]) === pin &&
      (cAktif === -1 || isTrue_(row[cAktif]))
    ) {
      setSession_({
        nip,
        nama: row[cNama] || '',
        role: row[cRole] || 'PTK',
        email: row[cEmail] || ''
      });
      return { ok:true };
    }
  }
  return { ok:false, msg:'Login gagal' };
}

function authMe() {
  const s = getSession_();
  if (!s) return { ok:false };
  return { ok:true, ...s };
}

function authLogout() {
  clearSession_();
  return { ok:true };
}
