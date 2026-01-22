/*************************************************
 * PasswordService - Ganti Password via WA OTP (Starsender)
 *
 * Sheet Users memakai kolom:
 * - NIP
 * - PIN               (sekarang dipakai sebagai PASSWORD)
 * - No_HP             (format 08xxx atau 62xxx)
 * - ResetPIN_OTP
 * - ResetPIN_ExpiredAt
 * - OTP_Attempt
 * - PIN_LastChangedAt (opsional)
 *************************************************/

const OTP_EXPIRE_MIN = 5;
const OTP_MAX_ATTEMPT = 3;
const OTP_CACHE_TTL = 30; // detik

/* ===================== PUBLIC API ===================== */

// STEP 1: Request OTP (cek password lama)
function requestPasswordChange(oldPassword){
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)){
    return { ok:false, msg:'Sistem sibuk. Coba lagi.' };
  }

  try{
    const s = requireLogin_();
    const nip = String(s.nip || '').trim();
    if (!nip) return { ok:false, msg:'Session habis. Silakan login ulang.' };

    const ctx = loadUsersPasswordMap_();
    const cols = ctx.cols;
    const user = ctx.map[nip];

    if(!user) return { ok:false, msg:'User tidak ditemukan' };

    if(String(user.pass||'').trim() !== String(oldPassword||'').trim()){
      return { ok:false, msg:'Password lama tidak sesuai' };
    }

    const noHpRaw = String(user.noHp||'').trim();
    if(!noHpRaw) return { ok:false, msg:'No_HP belum terdaftar di Users' };

    const noHp = normalizeNoHP_(noHpRaw);
    if(!noHp) return { ok:false, msg:'Format No_HP tidak valid (harus 08xxx atau 62xxx)' };

    const otp = generateOTP_();
    const expireAt = new Date(Date.now() + OTP_EXPIRE_MIN * 60000);

    updateUserRow_(user.rowIndex, (row)=>{
      row[cols.cOTP] = otp;
      row[cols.cEXP] = expireAt;
      row[cols.cTRY] = 0;
    });
    clearUsersCache_();

    // kirim WA (Starsender)
    sendWAOTP_(noHp, otp);

    return { ok:true, msg:'Kode verifikasi dikirim ke No_HP' };
  }catch(e){
    // kalau Starsender error, kasih pesan lebih jelas
    const msg = String(e && e.message ? e.message : e);
    return { ok:false, msg:`Gagal mengirim OTP: ${msg}` };
  }finally{
    lock.releaseLock();
  }
}

// STEP 2: Verifikasi OTP
function verifyPasswordOTP(inputOTP){
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)){
    return { ok:false, msg:'Sistem sibuk. Coba lagi.' };
  }

  try{
    const s = requireLogin_();
    const nip = String(s.nip || '').trim();

    const ctx = loadUsersPasswordMap_();
    const cols = ctx.cols;
    const user = ctx.map[nip];
    if(!user) return { ok:false, msg:'User tidak ditemukan' };

    const exp = user.exp;
    if (!exp || new Date(exp) < new Date()){
      return { ok:false, msg:'Kode OTP sudah kedaluwarsa. Silakan kirim ulang.' };
    }

    if(String(user.otp||'').trim() !== String(inputOTP||'').trim()){
      const attempt = (Number(user.attempt)||0) + 1;
      updateUserRow_(user.rowIndex, (row)=>{
        row[cols.cTRY] = attempt;
        if(attempt >= OTP_MAX_ATTEMPT){
          clearOTPRow_(row, cols);
        }
      });
      clearUsersCache_();

      if(attempt >= OTP_MAX_ATTEMPT){
        return { ok:false, msg:'OTP salah terlalu banyak. Silakan kirim ulang.' };
      }
      return { ok:false, msg:'OTP tidak sesuai' };
    }

    return { ok:true };
  }catch(e){
    return { ok:false, msg:'Verifikasi gagal' };
  }finally{
    lock.releaseLock();
  }
}

// STEP 3: Set password baru
function updatePassword(newPass){
  const lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)){
    return { ok:false, msg:'Sistem sibuk. Coba lagi.' };
  }

  try{
    const s = requireLogin_();
    const nip = String(s.nip || '').trim();

    if(!validatePassword_(String(newPass||''))){
      return { ok:false, msg:'Password tidak memenuhi aturan: min 8, ada huruf besar, angka, simbol.' };
    }

    const ctx = loadUsersPasswordMap_();
    const cols = ctx.cols;
    const user = ctx.map[nip];
    if(!user) return { ok:false, msg:'User tidak ditemukan' };

    if(String(user.pass||'').trim() === String(newPass||'').trim()){
      return { ok:false, msg:'Password baru tidak boleh sama dengan password lama' };
    }

    updateUserRow_(user.rowIndex, (row)=>{
      row[cols.cPASS] = String(newPass);
      if(cols.cCHG !== -1){
        row[cols.cCHG] = new Date();
      }
      clearOTPRow_(row, cols);
    });

    clearUsersCache_();
    clearSession_();

    return { ok:true };
  }catch(e){
    return { ok:false, msg:'Gagal mengganti password' };
  }finally{
    lock.releaseLock();
  }
}

/* ===================== STARSENDER CONFIG ===================== */

/**
 * MODE:
 * - bearer: pakai Authorization: Bearer <token>
 * - apikey: pakai headers X-Api-Key / api_key + device
 * - legacy_sendText: endpoint lama yang butuh header apikey + payload tujuan/message
 *
 * Endpoint & format Starsender bisa beda antar akun/versi.
 * Tapi ini dibuat fleksibel: tinggal sesuaikan 3 baris payload/headers kalau perlu.
 */
function sendWAOTP_(waE164, otp){
  const url = cfgRequireString('STARSENDER_URL');
  const apiKey = cfgRequireString('STARSENDER_APIKEY');
  const modeRaw = cfgGet('STARSENDER_MODE', '');
  const mode = String(modeRaw || '').trim().toLowerCase();
  const validModes = ['bearer', 'apikey', 'legacy_sendtext'];

  // samakan format dengan script PPh
  // tujuan = 62xxxxxxxxxx (tanpa +)
  const tujuan = waE164.replace(/^\+/, '');

  const message =
    "HCIS Sabilul Qur'an\n" +
    "Kode verifikasi ganti password Anda: " + otp + "\n" +
    "Berlaku 5 menit.\n" +
    "Jangan bagikan kode ini kepada siapa pun.";

  if (!mode) {
    throw new Error('STARSENDER_MODE belum diisi di HCIS_Config. Isi dengan salah satu: bearer, apikey, legacy_sendText.');
  }

  if (validModes.indexOf(mode) === -1) {
    throw new Error(`STARSENDER_MODE tidak dikenal ("${modeRaw}"). Isi dengan salah satu: bearer, apikey, legacy_sendText.`);
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
  const code = resp.getResponseCode();
  const text = resp.getContentText() || '';

  if (code !== 200) {
    throw new Error('Starsender gagal: HTTP ' + code + ' - ' + text);
  }

  // kalau response JSON dan ada success:false
  try {
    const json = JSON.parse(text);
    if (json.success === false) {
      throw new Error('Starsender gagal (body): ' + (json.message || text));
    }
  } catch(_) {
    // abaikan kalau bukan JSON
  }

  return true;
}

/* ===================== HELPERS ===================== */

function generateOTP_(){
  return Math.floor(100000 + Math.random()*900000).toString();
}

// input: 08xxx / 62xxx → output 62xxx
function normalizeNoHP_(no){
  let n = String(no||'').trim();
  n = n.replace(/\s+/g,'').replace(/\D/g,''); // digits only
  if (n.startsWith('08')) return '62' + n.slice(1);
  if (n.startsWith('62')) return n;
  return null;
}

// password policy
function validatePassword_(p){
  const s = String(p||'');
  return (
    s.length >= 8 &&
    /[A-Z]/.test(s) &&
    /[0-9]/.test(s) &&
    /[!@#$%^&*()_+\-=]/.test(s)
  );
}

function clearOTP_(rowIndex, headers){
  const cols = {
    cOTP: col_(headers, 'ResetPIN_OTP'),
    cEXP: col_(headers, 'ResetPIN_ExpiredAt'),
    cTRY: col_(headers, 'OTP_Attempt'),
  };
  updateUserRow_(rowIndex, (row) => clearOTPRow_(row, cols));
}

function loadUsersPasswordMap_(){
  const cache = CacheService.getScriptCache();
  const cached = cache.get(_USERS_PASSWORD_CACHE_KEY);
  if(cached){
    try{
      return JSON.parse(cached);
    }catch(_){
      // ignore
    }
  }

  const t = readTable_(CFG.SHEET_USERS);
  const h = t.headers;
  const r = t.rows;

  const cols = {
    cNIP: col_(h,'NIP'),
    cPASS: col_(h,'PIN'),
    cNoHP: col_(h,'No_HP'),
    cOTP: col_(h,'ResetPIN_OTP'),
    cEXP: col_(h,'ResetPIN_ExpiredAt'),
    cTRY: col_(h,'OTP_Attempt'),
    cCHG: col_(h,'PIN_LastChangedAt'),
  };

  if (cols.cNIP === -1 || cols.cPASS === -1) throw new Error('Header Users wajib punya NIP dan PIN');
  if (cols.cNoHP === -1) throw new Error('Kolom No_HP belum ada di Users');
  if (cols.cOTP === -1 || cols.cEXP === -1 || cols.cTRY === -1) throw new Error('Kolom OTP (ResetPIN_*) belum lengkap di Users');

  const map = {};
  r.forEach((rw,i)=>{
    const nip = String(rw[cols.cNIP]||'').trim();
    if(!nip) return;
    map[nip] = {
      rowIndex: i + 2,
      pass: rw[cols.cPASS],
      noHp: rw[cols.cNoHP] || '',
      otp: cols.cOTP === -1 ? '' : rw[cols.cOTP],
      exp: cols.cEXP === -1 ? '' : rw[cols.cEXP],
      attempt: cols.cTRY === -1 ? 0 : (Number(rw[cols.cTRY])||0)
    };
  });

  const payload = { map, cols };
  cache.put(_USERS_PASSWORD_CACHE_KEY, JSON.stringify(payload), OTP_CACHE_TTL);
  return payload;
}

function updateUserRow_(rowIndex, mutator){
  const sh = getSheet_(CFG.SHEET_USERS);
  const lastCol = sh.getLastColumn();
  const range = sh.getRange(rowIndex, 1, 1, lastCol);
  const values = range.getValues();
  const row = values[0];
  mutator(row);
  range.setValues([row]);
}

function clearOTPRow_(row, cols){
  if(cols.cOTP !== -1) row[cols.cOTP] = '';
  if(cols.cEXP !== -1) row[cols.cEXP] = '';
  if(cols.cTRY !== -1) row[cols.cTRY] = '';
}


