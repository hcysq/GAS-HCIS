# 🔍 ANALISA: Debug Render Data Profil

## Situasi Sekarang

### Alur Data:
```
User Login
  ↓
Session: { nip, userId, email, ... }
  ↓
Klik "Profil" tab
  ↓
renderProfil() → getProfilUsersDetail()
  ↓
Cari di MASTERDATA_SS_ID spreadsheet
  ↓
Baca tab Masterdata pakai MASTERDATA_GID
  ↓
Tampilkan data atau error
```

### Konfigurasi Sekarang:
- **HCIS_Config sheet** (GID: 1743564124) berisi:
  - Key: `USER_GID` 
  - Value: `304619888` (GID tab Users)
- **MASTERDATA_SS_ID** dan **MASTERDATA_GID** → Dikonfigurasi di HCIS_Config
- **Profile.js** → Cari user di sheet Users pakai NIP/USER_ID

---

## 🐛 Masalah yang Mungkin Terjadi

### Problem 1: USER_GID Config Tidak Dipakai
**Lokasi:** Profile.js baris ~365 (`getUsersSheetByConfig_()`)
```javascript
const gidRaw = cfgGet('MASTERDATA_GID', '');  // ← Cari MASTERDATA_GID
const gid = Number(gidRaw);
```

**Masalah:** Kode cari `MASTERDATA_GID` tapi konfigurasi Anda isi `USER_GID`!

**Solusi:** Ubah key menjadi `USER_GID` atau pastikan `MASTERDATA_GID` sudah diisi di HCIS_Config.

---

### Problem 2: Session Tidak Punya USER_ID
**Lokasi:** Profile.js baris ~327 (`getProfilUsersDetail()`)
```javascript
const userIdSession = String(s.userId || '').trim();
if (!nipKey && !userIdSession) {
  return { ok:false, msg:'Session tidak memiliki USER_ID...' };
}
```

**Masalah:** Jika `s.userId` kosong dan NIP tidak ter-normalize, fungsi langsung error.

**Solusi:** Pastikan kolom `USER_ID` ada di Users sheet dan ter-populate saat login.

---

### Problem 3: MASTERDATA_SS_ID Tidak Dikonfigurasi
**Lokasi:** Profile.js baris ~283 (`getMasterdataSpreadsheet_()`)
```javascript
const ssId = cfgGetString('MASTERDATA_SS_ID', '');
if (!ssId) return { ss: SpreadsheetApp.getActive() };  // ← Pakai aktif sheet
```

**Masalah:** Jika tidak dikonfigurasi, sistem pakai spreadsheet aktif (yang sedang dibuka di editor).

**Solusi:** Isi `MASTERDATA_SS_ID` dengan ID spreadsheet yang benar.

---

## 💡 Saran Perbaikan

### Opsi A: Pakai Script Properties (Rekomendasi ✅)
**Keuntungan:**
- ✅ Lebih cepat (tidak perlu baca sheet setiap kali)
- ✅ Aman (tidak keliatan di spreadsheet)
- ✅ Mudah di-maintain

**Implementasi:**
```javascript
// Set sekali di Properties
Properties.getScriptProperties().setProperty('MASTERDATA_SS_ID', 'xxx');
Properties.getScriptProperties().setProperty('MASTERDATA_GID', 'yyy');
Properties.getScriptProperties().setProperty('USER_GID', 'zzz');

// Baca di code
cfgGetString('MASTERDATA_SS_ID', '') // baca dari Properties
```

**Cara setup:**
1. Buka Script Properties (klik kunci 🔒 ikon di editor)
2. Tambah:
   - Key: `MASTERDATA_SS_ID`, Value: `[ID spreadsheet Masterdata]`
   - Key: `MASTERDATA_GID`, Value: `[GID tab Masterdata di spreadsheet tsb]`
   - Key: `USER_GID`, Value: `304619888` (GID tab Users)

---

### Opsi B: Perbaiki HCIS_Config Sheet
**Kalau tetap pakai sheet:**

1. Buka HCIS_Config sheet
2. Pastikan ada row:
   ```
   Key: MASTERDATA_SS_ID
   Value: [ID spreadsheet Masterdata]
   Note: ID spreadsheet yang berisi tab Masterdata & Users
   
   Key: MASTERDATA_GID
   Value: [GID tab Masterdata]
   Note: Sheet ID tab dengan data Masterdata
   
   Key: USER_GID
   Value: 304619888
   Note: Sheet ID tab Users (sudah ada)
   ```

---

### Opsi C: Ubah Kode Agar Lebih Robust
**Struktur baru:**
```javascript
// Di Config.js
function ensureRequiredConfig_() {
  const required = {
    'MASTERDATA_SS_ID': 'Spreadsheet ID Masterdata',
    'MASTERDATA_GID': 'GID tab Masterdata',
    'USER_GID': 'GID tab Users',
    'SHEET_MASTERDATA': 'Nama tab Masterdata',
    'SHEET_USERS': 'Nama tab Users'
  };
  
  const missing = [];
  for (const [key, label] of Object.entries(required)) {
    if (!cfgGet(key)) missing.push(`${key}: ${label}`);
  }
  
  if (missing.length > 0) {
    throw new Error(`Config tidak lengkap:\n${missing.join('\n')}`);
  }
}

// Di Profile.js
function getProfilUsersDetail() {
  try {
    ensureRequiredConfig_();  // ← Validasi config
    // ... lanjut
  }
}
```

---

## 📋 Checklist Debug

### Cek 1: HCIS_Config Lengkap?
```
Buka HCIS_Config sheet (GID: 1743564124)
Pastikan ada:
☐ MASTERDATA_SS_ID = [ID spreadsheet]
☐ MASTERDATA_GID = [GID Masterdata tab]
☐ USER_GID = 304619888
☐ SHEET_MASTERDATA = "Masterdata" (atau nama asli)
☐ SHEET_USERS = "Users" (atau nama asli)
```

### Cek 2: Session Punya Data?
```javascript
// Di Console browser setelah login:
console.log(state.me);
// Output harus ada:
// { nip: "...", userId: "...", ... }
```

### Cek 3: Profil API Response
```javascript
// Di Console browser di tab Profil:
google.script.run.getProfilUsersDetail();
// Buka console di Google Apps Script editor
// Cek log error
```

### Cek 4: Data di Sheet Ada?
```
Buka Users sheet (GID: 304619888)
Pastikan:
☐ Ada kolom: NIP, USER_ID, Nama, Email
☐ Ada baris data untuk user yang login
☐ Data NIP sama dengan session NIP
```

---

## 🎯 Rekomendasi Saya

### Langkah 1: Pakai Script Properties
**Paling simple & cepat:**
```javascript
// Cukup jalan 1x (bisa dari test/setup)
function setupPropertiesScriptOnce_() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('MASTERDATA_SS_ID', 'PASTE_SPREADSHEET_ID_HERE');
  props.setProperty('MASTERDATA_GID', '123456789'); // GID Masterdata tab
  props.setProperty('USER_GID', '304619888');
  props.setProperty('SHEET_MASTERDATA', 'Masterdata');
  props.setProperty('SHEET_USERS', 'Users');
}
```

### Langkah 2: Update Config.js
Tambah fallback ke Script Properties:
```javascript
function cfgGet(key, defaultValue) {
  // 1. Coba dari HCIS_Config sheet cache
  const cached = ... // kode sekarang
  
  // 2. Jika tidak ada, coba dari Script Properties
  const fromProps = PropertiesService.getScriptProperties().getProperty(key);
  if (fromProps) return fromProps;
  
  // 3. Default
  return defaultValue;
}
```

### Langkah 3: Test
1. Setup properties pakai `setupPropertiesScriptOnce_()`
2. Deploy
3. Login
4. Klik Profil
5. Harusnya data muncul

---

## 📊 Perbandingan Solusi

| Aspek | Sheet Config | Script Properties | Hybrid |
|-------|-------------|------------------|--------|
| **Setup** | Buka sheet, isi data | Code sekali | Keduanya |
| **Kecepatan** | Ada cache 5 menit | Instant | Instant |
| **Aman** | Keliatan di sheet | Tersembunyi | Tersembunyi |
| **Mudah diubah** | Mudah | Perlu redeploy | Mudah (sheet) |
| **Rekomendasih** | Untuk user | ✅ **Terbaik** | Kompromi |

---

## 🔧 Kode Perbaikan Siap Pakai

### Jika mau pakai Hybrid (Script Properties + Sheet):
```javascript
function cfgGet(key, defaultValue) {
  key = String(key || '').trim();
  if (!key) return defaultValue;

  // 1. Cek Script Properties dulu (paling cepat)
  const scriptProps = PropertiesService.getScriptProperties();
  const fromScript = scriptProps.getProperty(key);
  if (fromScript) return fromScript;

  // 2. Jika tidak ada, cek HCIS_Config sheet dengan cache
  const cache = CacheService.getScriptCache();
  const cached = cache.get(_CFG_CACHE_KEY);
  if (cached) {
    try {
      const map = JSON.parse(cached);
      if (Object.prototype.hasOwnProperty.call(map, key)) return map[key];
    } catch (e) {}
  }

  // 3. Load dari sheet, cache 5 menit
  const map = _loadCfgMap_();
  cache.put(_CFG_CACHE_KEY, JSON.stringify(map), _CFG_CACHE_TTL);
  if (Object.prototype.hasOwnProperty.call(map, key)) return map[key];
  
  return defaultValue;
}
```

---

## ✅ Kesimpulan

**Masalah utama:** Config tidak konsisten atau tidak lengkap

**Solusi terbaik:**
1. ✅ Gunakan Script Properties untuk config statis
2. ✅ Sheet HCIS_Config untuk config yang sering berubah
3. ✅ Kode cek fallback ke keduanya

**Waktu implementasi:** ~15 menit

**Tingkat kesulitan:** Mudah

---

## Langkah Selanjutnya?

Mau saya implement perbaikannya? Saya bisa:
1. Update Config.js pakai hybrid approach
2. Setup helper function untuk test config
3. Pastikan semua key penting ter-verify saat startup

Atau Anda cek sendiri dulu checklist di atas? 😊
