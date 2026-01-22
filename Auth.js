const _USERS_CACHE_KEY = 'HCIS_USERS_MAP_V1';
const _USERS_CACHE_TTL = 600; // detik
const _USERS_PASSWORD_CACHE_KEY = 'HCIS_USERS_PASSMAP_V1';
const _LOGIN_FAIL_CACHE_PREFIX = 'HCIS_LOGIN_FAIL_V1:';
const _LOGIN_BLOCK_CACHE_PREFIX = 'HCIS_LOGIN_BLOCK_V1:';
const _LOGIN_FAIL_TTL = 900; // detik
const _LOGIN_BLOCK_TTL = 300; // detik
const _LOGIN_FAIL_MAX = 5;

/*************************************************
 * Authentication
 *************************************************/

function authLogin(nip, pin) {
  try {
    nip = txt(nip);
    pin = txt(pin);
    
    if (!nip || !pin) {
      return { ok:false, msg:'NIP & PIN wajib diisi' };
    }

    if (isLoginBlocked_(nip)) {
      return { ok:false, msg:'Terlalu banyak percobaan gagal. Akun diblok sementara beberapa menit.' };
    }

    let userMap;
    try {
      userMap = loadUsersMap_();
    } catch (e) {
      Logger.log('Error loading users map: ' + (e.message || e));
      return { ok:false, msg:'Gagal membaca data user. Sheet Users mungkin tidak ada atau formatnya salah.' };
    }

    if (!userMap || Object.keys(userMap).length === 0) {
      Logger.log('Users map kosong');
      return { ok:false, msg:'Data user belum tersedia di sistem.' };
    }

    const user = userMap[nip];
    
    if (!user) {
      Logger.log('NIP tidak ditemukan: ' + nip);
      const failState = registerLoginFailure_(nip);
      if (failState.blocked) {
        return { ok:false, msg:'Terlalu banyak percobaan gagal. Akun diblok sementara beberapa menit.' };
      }
      return { ok:false, msg:'NIP atau password salah.' };
    }

    if (!user.aktif) {
      Logger.log('User tidak aktif: ' + nip);
      return { ok:false, msg:'Akun Anda tidak aktif. Hubungi admin.' };
    }

    // Hash PIN yang diinput
    const inputHash = hashPin_(pin);
    const storedHash = user.pinHash;
    
    if (inputHash !== storedHash) {
      Logger.log('Password hash tidak cocok untuk NIP: ' + nip);
      Logger.log('Input hash: ' + inputHash);
      Logger.log('Stored hash: ' + storedHash);
      const failState = registerLoginFailure_(nip);
      if (failState.blocked) {
        return { ok:false, msg:'Terlalu banyak percobaan gagal. Akun diblok sementara beberapa menit.' };
      }
      return { ok:false, msg:'NIP atau password salah.' };
    }

    clearLoginFail_(nip);

    // Parse multiple roles dari kolom boolean (PTK, KAPLA, ADMIN)
    const roleArray = parseRoles_(user);

    const token = setSession_({
      nip,
      nama: user.nama || '',
      roles: roleArray,
      role: roleArray.join(','),  // Convert array to string
      email: user.email || '',
      userId: user.userId || ''
    });
    
    Logger.log('Login berhasil untuk NIP: ' + nip + ' dengan roles: ' + roleArray.join(', '));
    return { ok:true, token };
  } catch (err) {
    Logger.log('Error di authLogin: ' + (err.message || err));
    return { ok:false, msg:'Error: ' + (err.message || err) };
  }
}

function registerLoginFailure_(nip) {
  const cache = CacheService.getScriptCache();
  const failKey = getLoginFailKey_(nip);
  const blockKey = getLoginBlockKey_(nip);
  const current = Number(cache.get(failKey) || 0);
  const nextCount = current + 1;
  cache.put(failKey, String(nextCount), _LOGIN_FAIL_TTL);
  if (nextCount >= _LOGIN_FAIL_MAX) {
    cache.put(blockKey, String(Date.now()), _LOGIN_BLOCK_TTL);
    return { blocked: true, count: nextCount };
  }
  return { blocked: false, count: nextCount };
}

function isLoginBlocked_(nip) {
  const cache = CacheService.getScriptCache();
  return Boolean(cache.get(getLoginBlockKey_(nip)));
}

function clearLoginFail_(nip) {
  const cache = CacheService.getScriptCache();
  cache.remove(getLoginFailKey_(nip));
  cache.remove(getLoginBlockKey_(nip));
}

function getLoginFailKey_(nip) {
  return `${_LOGIN_FAIL_CACHE_PREFIX}${nip}`;
}

function getLoginBlockKey_(nip) {
  return `${_LOGIN_BLOCK_CACHE_PREFIX}${nip}`;
}

function authMe(token) {
  try {
    const s = requireLogin_(token);
    return { ok:true, ...s };
  } catch (e) {
    return { ok:false };
  }
}

function authLogout(request) {
  clearSession_(request);
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
      Logger.log('Cache parsing gagal, reload dari sheet');
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const cachedAfterLock = cache.get(_USERS_CACHE_KEY);
    if (cachedAfterLock) {
      try {
        return JSON.parse(cachedAfterLock);
      } catch (_) {
        Logger.log('Cache parsing gagal setelah lock, reload dari sheet');
      }
    }

    const t = readUsersTable_(CFG.SHEET_USERS);
    const h = t.headers;
    const r = t.rows;

    const cNIP = col_(h, 'NIP');
    const cPIN = col_(h, 'PIN');
    const cAktif = col_(h, 'Aktif');
    const cNama = col_(h, 'Nama');
    const cRole = col_(h, 'Role');
    const cEmail = col_(h, 'Email');
    const cUserId = col_(h, 'USER_ID');

    if (cNIP === -1 || cPIN === -1) {
      throw new Error('Sheet Users harus punya kolom NIP dan PIN. Pastikan header row 1 ada kolom tersebut.');
    }

    if (r.length === 0) {
      Logger.log('Warning: Sheet Users kosong (tidak ada data)');
    }

    let pinCacheMap = {};
    const cachedPins = cache.get(_USERS_PASSWORD_CACHE_KEY);
    if (cachedPins) {
      try {
        pinCacheMap = JSON.parse(cachedPins);
      } catch (_) {
        Logger.log('Cache PIN parsing gagal, hitung ulang PIN hash');
        pinCacheMap = {};
      }
    }

    const map = {};
    for (const row of r) {
      const nip = txt(row[cNIP]);
      if (!nip) continue;

      const pinRaw = txt(row[cPIN]);
      if (!pinRaw) {
        Logger.log('Warning: NIP ' + nip + ' tidak punya PIN');
        continue;
      }

      let pinHash = '';
      const cachedPin = pinCacheMap[nip];
      if (cachedPin && cachedPin.pinRaw === pinRaw && cachedPin.pinHash) {
        pinHash = cachedPin.pinHash;
      } else {
        pinHash = hashPin_(pinRaw);
        pinCacheMap[nip] = { pinRaw, pinHash };
      }

      map[nip] = {
        pinHash,
        aktif: cAktif === -1 ? true : isTrue_(row[cAktif]),
        nama: row[cNama] ? txt(row[cNama]) : '',
        role: row[cRole] ? txt(row[cRole]) : 'PTK',
        email: row[cEmail] ? txt(row[cEmail]) : '',
        userId: cUserId === -1 ? '' : txt(row[cUserId])
      };
    }

    Logger.log('Loaded ' + Object.keys(map).length + ' users dari sheet');
    cache.put(_USERS_CACHE_KEY, JSON.stringify(map), _USERS_CACHE_TTL);
    cache.put(_USERS_PASSWORD_CACHE_KEY, JSON.stringify(pinCacheMap), _USERS_CACHE_TTL);
    return map;
  } catch (err) {
    throw new Error('Gagal load Users map: ' + (err.message || err));
  } finally {
    lock.releaseLock();
  }
}

function readUsersTable_(sheetName) {
  const neededHeaders = ['NIP', 'PIN', 'Aktif', 'Nama', 'Role', 'Email', 'USER_ID'];
  const sh = getSheet_(sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    return { headers: neededHeaders, rows: [] };
  }

  const lastCol = sh.getLastColumn();
  const headerRow = sh
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(h => String(h).trim());
  const numRows = lastRow - 1;

  const columns = neededHeaders.map(name => {
    const idx = headerRow.indexOf(name);
    if (idx === -1) {
      return { name, values: Array(numRows).fill('') };
    }
    const values = sh
      .getRange(2, idx + 1, numRows, 1)
      .getValues()
      .map(row => row[0]);
    return { name, values };
  });

  const rows = [];
  for (let i = 0; i < numRows; i++) {
    const row = [];
    for (const col of columns) {
      row.push(col.values[i]);
    }
    rows.push(row);
  }

  return { headers: neededHeaders, rows };
}

function clearUsersCache_() {
  const cache = CacheService.getScriptCache();
  cache.remove(_USERS_CACHE_KEY);
  cache.remove(_USERS_PASSWORD_CACHE_KEY);
}

/**
 * DEBUG: Validasi PIN untuk testing
 * Gunakan di Script Editor > Run > validatePin_(nip, pin)
 */
function validatePin_(nip, pin) {
  nip = txt(nip);
  pin = txt(pin);
  
  try {
    const userMap = loadUsersMap_();
    const user = userMap[nip];
    
    if (!user) {
      return { ok: false, msg: 'NIP tidak ditemukan', nip };
    }
    
    const inputHash = hashPin_(pin);
    const storedHash = user.pinHash;
    
    return {
      ok: inputHash === storedHash,
      nip,
      pinInput: pin,
      hashInput: inputHash,
      hashStored: storedHash,
      userAktif: user.aktif,
      userName: user.nama
    };
  } catch (err) {
    return { ok: false, error: err.message || err };
  }
}

function hashPin_(pin) {
  const raw = txt(pin);
  if (!raw) return '';
  
  try {
    // computeDigest mengembalikan signed bytes array
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
    // Convert signed bytes ke unsigned hex
    return bytes
      .map(b => {
        // Konversi signed byte (-128..127) ke unsigned (0..255)
        const unsigned = b < 0 ? 256 + b : b;
        // Format sebagai hex 2 digit
        return ('0' + unsigned.toString(16)).slice(-2);
      })
      .join('');
  } catch (err) {
    Logger.log('Error hashing pin: ' + (err.message || err));
    return '';
  }
}
