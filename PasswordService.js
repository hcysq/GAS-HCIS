/*************************************************
 * PasswordService - Ganti Password via WA OTP (Starsender)
 *
 * Sheet Users memakai kolom:
 * - NIP
 * - PIN               (sekarang dipakai sebagai PASSWORD)
 * - WhatsApp          (format 08xxx atau 62xxx)
 * - ResetPIN_OTP
 * - ResetPIN_ExpiredAt
 * - OTP_Attempt
 * - PIN_LastChangedAt (opsional)
 *************************************************/

const OTP_EXPIRE_MIN = 5;
const OTP_MAX_ATTEMPT = 3;

/* ===================== PUBLIC API ===================== */

// STEP 1: Request OTP (cek password lama)
function requestPasswordChange(oldPassword){
  try{
    const s = requireLogin_();
    const nip = String(s.nip || '').trim();
    if (!nip) return { ok:false, msg:'Session habis. Silakan login ulang.' };

    const t = readTable_(CFG.SHEET_USERS);
    const h = t.headers;
    const r = t.rows;

    const cNIP  = col_(h,'NIP');
    const cPASS = col_(h,'PIN'); // tetap pakai kolom PIN sebagai password
    const cWA   = col_(h,'WhatsApp');
    const cOTP  = col_(h,'ResetPIN_OTP');
    const cEXP  = col_(h,'ResetPIN_ExpiredAt');
    const cTRY  = col_(h,'OTP_Attempt');

    if (cNIP === -1 || cPASS === -1) return { ok:false, msg:'Header Users wajib punya NIP dan PIN' };
    if (cWA === -1) return { ok:false, msg:'Kolom WhatsApp belum ada di Users' };
    if (cOTP === -1 || cEXP === -1 || cTRY === -1) return { ok:false, msg:'Kolom OTP (ResetPIN_*) belum lengkap di Users' };

    let rowIndex = -1; // sheet row number
    let row = null;

    r.forEach((rw,i)=>{
      if (String(rw[cNIP]||'').trim() === nip){
        rowIndex = i + 2; // + header row
        row = rw;
      }
    });
    if(!row) return { ok:false, msg:'User tidak ditemukan' };

    if(String(row[cPASS]||'').trim() !== String(oldPassword||'').trim()){
      return { ok:false, msg:'Password lama tidak sesuai' };
    }

    const waRaw = String(row[cWA]||'').trim();
    if(!waRaw) return { ok:false, msg:'Nomor WhatsApp belum terdaftar di Users' };

    const wa = normalizeWA_(waRaw);
    if(!wa) return { ok:false, msg:'Format nomor WhatsApp tidak valid (harus 08xxx atau 62xxx)' };

    const otp = generateOTP_();
    const expireAt = new Date(Date.now() + OTP_EXPIRE_MIN * 60000);

    const sh = getSheet_(CFG.SHEET_USERS);
    sh.getRange(rowIndex, cOTP+1).setValue(otp);
    sh.getRange(rowIndex, cEXP+1).setValue(expireAt);
    sh.getRange(rowIndex, cTRY+1).setValue(0);

    // kirim WA (Starsender)
    sendWAOTP_(wa, otp);

    return { ok:true, msg:'Kode verifikasi dikirim ke WhatsApp' };
  }catch(e){
    // kalau Starsender error, kasih pesan lebih jelas
    const msg = String(e && e.message ? e.message : e);
    return { ok:false, msg:`Gagal mengirim OTP: ${msg}` };
  }
}

// STEP 2: Verifikasi OTP
function verifyPasswordOTP(inputOTP){
  try{
    const s = requireLogin_();
    const nip = String(s.nip || '').trim();

    const t = readTable_(CFG.SHEET_USERS);
    const h = t.headers;
    const r = t.rows;

    const cNIP = col_(h,'NIP');
    const cOTP = col_(h,'ResetPIN_OTP');
    const cEXP = col_(h,'ResetPIN_ExpiredAt');
    const cTRY = col_(h,'OTP_Attempt');

    let rowIndex = -1;
    let row = null;
    r.forEach((rw,i)=>{
      if (String(rw[cNIP]||'').trim() === nip){
        rowIndex = i + 2;
        row = rw;
      }
    });
    if(!row) return { ok:false, msg:'User tidak ditemukan' };

    const exp = row[cEXP];
    if (!exp || new Date(exp) < new Date()){
      return { ok:false, msg:'Kode OTP sudah kedaluwarsa. Silakan kirim ulang.' };
    }

    if(String(row[cOTP]||'').trim() !== String(inputOTP||'').trim()){
      const attempt = (Number(row[cTRY])||0) + 1;
      getSheet_(CFG.SHEET_USERS).getRange(rowIndex, cTRY+1).setValue(attempt);

      if(attempt >= OTP_MAX_ATTEMPT){
        clearOTP_(rowIndex,h);
        return { ok:false, msg:'OTP salah terlalu banyak. Silakan kirim ulang.' };
      }
      return { ok:false, msg:'OTP tidak sesuai' };
    }

    return { ok:true };
  }catch(e){
    return { ok:false, msg:'Verifikasi gagal' };
  }
}

// STEP 3: Set password baru
function updatePassword(newPass){
  try{
    const s = requireLogin_();
    const nip = String(s.nip || '').trim();

    if(!validatePassword_(String(newPass||''))){
      return { ok:false, msg:'Password tidak memenuhi aturan: min 8, ada huruf besar, angka, simbol.' };
    }

    const t = readTable_(CFG.SHEET_USERS);
    const h = t.headers;
    const r = t.rows;

    const cNIP  = col_(h,'NIP');
    const cPASS = col_(h,'PIN');
    const cCHG  = col_(h,'PIN_LastChangedAt');

    let rowIndex = -1;
    let row = null;
    r.forEach((rw,i)=>{
      if (String(rw[cNIP]||'').trim() === nip){
        rowIndex = i + 2;
        row = rw;
      }
    });
    if(!row) return { ok:false, msg:'User tidak ditemukan' };

    if(String(row[cPASS]||'').trim() === String(newPass||'').trim()){
      return { ok:false, msg:'Password baru tidak boleh sama dengan password lama' };
    }

    const sh = getSheet_(CFG.SHEET_USERS);
    sh.getRange(rowIndex, cPASS+1).setValue(String(newPass));

    if(cCHG !== -1){
      sh.getRange(rowIndex, cCHG+1).setValue(new Date());
    }

    // bersihin OTP & logout
    clearOTP_(rowIndex, h);
    clearSession_();

    return { ok:true };
  }catch(e){
    return { ok:false, msg:'Gagal mengganti password' };
  }
}

/* ===================== STARSENDER CONFIG ===================== */

function getConfigValue_(key){
  const canonical = cfgGet(key, '');
  return String(canonical ?? '').trim();
}


/**
 * MODE:
 * - bearer: pakai Authorization: Bearer <token>
 * - apikey: pakai headers X-Api-Key / api_key + device
 *
 * Endpoint & format Starsender bisa beda antar akun/versi.
 * Tapi ini dibuat fleksibel: tinggal sesuaikan 3 baris payload/headers kalau perlu.
 */
function sendWAOTP_(waE164, otp){
  const url = String(cfgGet('STARSENDER_URL', '') || '').trim();
  const apiKey = String(cfgGet('STARSENDER_APIKEY', '') || '').trim();

  if (!url) throw new Error('STARSENDER_URL belum diisi di HCIS_Config');
  if (!apiKey) throw new Error('STARSENDER_APIKEY belum diisi di HCIS_Config');

  // samakan format dengan script PPh
  // tujuan = 62xxxxxxxxxx (tanpa +)
  const tujuan = waE164.replace(/^\+/, '');

  const message =
    "HCIS Sabilul Qur'an\n" +
    "Kode verifikasi ganti password Anda: " + otp + "\n" +
    "Berlaku 5 menit.\n" +
    "Jangan bagikan kode ini kepada siapa pun.";

  const options = {
    method: 'post',
    headers: {
      apikey: apiKey
    },
    payload: {
      tujuan: tujuan,
      message: message
    },
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
function normalizeWA_(no){
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
  const sh = getSheet_(CFG.SHEET_USERS);
  ['ResetPIN_OTP','ResetPIN_ExpiredAt','OTP_Attempt'].forEach(k=>{
    const c = col_(headers,k);
    if(c !== -1) sh.getRange(rowIndex, c+1).setValue('');
  });
}


