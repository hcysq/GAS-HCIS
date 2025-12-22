const _USERS_CACHE_KEY = 'HCIS_USERS_MAP_V1';
const _USERS_CACHE_TTL = 60; // detik

/*************************************************
 * Authentication
 *************************************************/

function authLogin(nip, pin) {
  nip = txt(nip);
  pin = txt(pin);
  if (!nip || !pin) return { ok:false, msg:'NIP & PIN wajib diisi' };

  const userMap = loadUsersMap_();
  const user = userMap[nip];

  if (!user || !user.aktif) return { ok:false, msg:'Login gagal' };

  if (hashPin_(pin) !== user.pinHash) {
    return { ok:false, msg:'Login gagal' };
  }

  setSession_({
    nip,
    nama: user.nama,
    role: user.role,
    email: user.email
  });
  return { ok:true };
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

function loadUsersMap_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(_USERS_CACHE_KEY);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (_) {
      // abaikan, lanjut load ulang
    }
  }

  const t = readTable_(CFG.SHEET_USERS);
  const h = t.headers;
  const r = t.rows;

  const cNIP = col_(h, 'NIP');
  const cPIN = col_(h, 'PIN');
  const cAktif = col_(h, 'Aktif');
  const cNama = col_(h, 'Nama');
  const cRole = col_(h, 'Role');
  const cEmail = col_(h, 'Email');

  if (cNIP === -1 || cPIN === -1) {
    throw new Error('Header Users wajib punya NIP dan PIN');
  }

  const map = {};
  for (const row of r) {
    const nip = txt(row[cNIP]);
    if (!nip) continue;

    map[nip] = {
      pinHash: hashPin_(row[cPIN]),
      aktif: cAktif === -1 ? true : isTrue_(row[cAktif]),
      nama: row[cNama] || '',
      role: row[cRole] || 'PTK',
      email: row[cEmail] || ''
    };
  }

  cache.put(_USERS_CACHE_KEY, JSON.stringify(map), _USERS_CACHE_TTL);
  return map;
}

function clearUsersCache_() {
  CacheService.getScriptCache().remove(_USERS_CACHE_KEY);
}

function hashPin_(pin) {
  const raw = txt(pin);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return bytes
    .map(b => (b + 256) % 256)
    .map(b => ('0' + b.toString(16)).slice(-2))
    .join('');
}
