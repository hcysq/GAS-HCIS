# 🛠️ GUIDE: Fix Render Data Profil

## Quick Test (5 menit)

### Step 1: Cek Config Sekarang
1. Buka **Google Apps Script editor**
2. Di Console (Ctrl+Enter atau View > Logs)
3. Jalankan:
```javascript
testProfilConfig()
```

4. Lihat output di Logs:
   - ✅ Artinya config sudah benar
   - ❌ Artinya ada yang kurang

### Step 2: Setup Berdasarkan Hasil

---

## Skenario 1: Semua Config Kosong ❌

**Output testProfilConfig():**
```
❌ MASTERDATA_SS_ID: Kosong (akan pakai spreadsheet aktif)
❌ MASTERDATA_GID: Kosong (akan cari pakai SHEET_MASTERDATA)
❌ USER_GID: Kosong (akan cari pakai SHEET_USERS)
```

**Solusi:**

### A. Pakai Spreadsheet Aktif (Termudah)
Jika Masterdata & Users ada di spreadsheet yang sama dengan script:

1. Pastikan sheet bernama:
   - `Masterdata` (untuk data lengkap employee)
   - `Users` (untuk login & role)

2. Itu aja! System otomatis akan cari pakai nama sheet.

### B. Pakai HCIS_Config Sheet (Recommended)

1. Buka spreadsheet Anda
2. Buka tab **HCIS_Config**
3. Tambah row baru:
   ```
   Key: SHEET_MASTERDATA
   Value: Masterdata
   Note: Nama tab yang berisi data karyawan lengkap
   
   Key: SHEET_USERS
   Value: Users
   Note: Nama tab untuk login & akses
   ```

4. Jika Masterdata di spreadsheet berbeda:
   ```
   Key: MASTERDATA_SS_ID
   Value: [PASTE_SPREADSHEET_ID_MASTERDATA]
   Note: ID spreadsheet Masterdata
   ```
   
   Cara dapat ID:
   - Buka spreadsheet Masterdata
   - URL: `https://docs.google.com/spreadsheets/d/[ID_INI]/edit`
   - Copy bagian `[ID_INI]`

5. Test lagi: `testProfilConfig()`

---

## Skenario 2: Ada USER_GID tapi Tidak Ada MASTERDATA_GID ⚠️

**Output:**
```
✅ USER_GID: 304619888
❌ MASTERDATA_GID: Kosong
```

**Masalah:** System cari data Masterdata pakai nama sheet (`SHEET_MASTERDATA`), tapi Anda set USER_GID berarti pakai GID.

**Solusi: Sejajarkan cara referensi**

### Opsi A: Pakai Keduanya dengan GID
```
Key: MASTERDATA_GID
Value: [PASTE_MASTERDATA_GID]
Note: GID tab Masterdata

Key: USER_GID  
Value: 304619888
Note: GID tab Users (sudah ada)
```

Cara dapat GID:
1. Di spreadsheet, klik kanan tab
2. Pilih "Get sheet ID"
3. Copy nomor yang keluar

### Opsi B: Pakai Sheet Name (Lebih Simple)
```
Key: SHEET_MASTERDATA
Value: Masterdata
Note: Nama tab Masterdata

Key: SHEET_USERS
Value: Users
Note: Nama tab Users
```

Lalu hapus USER_GID (atau biarkan, tidak pakai).

---

## Skenario 3: Config Ada tapi Data Tidak Muncul di Profil 🤔

**Diagnosis:**
1. Jalankan: `testProfilConfig()`
   - Harusnya semua ✅

2. Jika OK, problem di data atau kode:
   
   a) **Session tidak ada USER_ID/NIP:**
   ```javascript
   // Di browser console setelah login
   console.log(state.me);
   // Cek ada `nip` atau `userId` tidak?
   ```
   
   b) **Data di Users/Masterdata tidak cocok:**
   ```javascript
   // Di Apps Script editor
   google.script.run.getProfilUsersDetail();
   ```
   Lihat error message di logs
   
   c) **Kolom header tidak ditemukan:**
   Buka tab Users/Masterdata, pastikan:
   - Baris 1 ada header: `NIP`, `USER_ID`, `Email`, `Nama`, dll
   - Data mulai dari baris 2

---

## Skenario 4: Ingin Pakai Script Properties (Paling Cepat)

### Setup (Jalankan 1x):
```javascript
function setupPropertiesOnce() {
  const props = PropertiesService.getScriptProperties();
  
  // Mandatory
  props.setProperty('MASTERDATA_SS_ID', ''); // Kosong = pakai aktif sheet
  props.setProperty('SHEET_MASTERDATA', 'Masterdata');
  props.setProperty('SHEET_USERS', 'Users');
  
  // Optional (kalau pakai GID)
  // props.setProperty('MASTERDATA_GID', '123456789');
  // props.setProperty('USER_GID', '304619888');
  
  Logger.log('✅ Script Properties sudah diset');
}
```

1. Jalankan: `setupPropertiesOnce()`
2. Akan otomatis isi Script Properties
3. Config bisa di-override dari HCIS_Config sheet

---

## Checklist Complete Fix

- [ ] Buka Apps Script editor
- [ ] Jalankan: `testProfilConfig()`
- [ ] Lihat hasilnya
- [ ] Setup HCIS_Config atau Properties sesuai hasil
- [ ] Jalankan: `testProfilConfig()` lagi
- [ ] Pastikan semua ✅
- [ ] Deploy
- [ ] Login & klik tab Profil
- [ ] Data harus muncul ✅

---

## Troubleshoot: Data Masih Tidak Muncul

### Check 1: Session Ada?
```javascript
// Browser console setelah login:
console.log(state.me.nip, state.me.userId);
```
Harusnya ada value, bukan kosong.

### Check 2: User Ada di Sheet?
```javascript
// Apps Script editor:
function testFindUser() {
  const s = requireLogin_();
  Logger.log('Session NIP:', s.nip);
  Logger.log('Session UserID:', s.userId);
  
  const result = getProfilUsersDetail();
  Logger.log(JSON.stringify(result, null, 2));
}
```
Jalankan & cek log.

### Check 3: Header Sheet Benar?
Buka tab `Users` dan `Masterdata`:
- Baris 1 ada header?
- Header include: `NIP`, `USER_ID`, `Email`, `Nama`?
- Data ada di baris 2 ke bawah?

### Check 4: NIP Di Session Cocok NIP Di Sheet?
```javascript
// Apps Script
function testNIPMatch() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Users');
  const data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  
  // Print semua NIP di sheet
  const nips = data.map(r => r[0]).filter(n => n);
  Logger.log('NIP di Users sheet:', nips);
}
```

Jalankan, bandingkan dengan session NIP.

---

## File-File Yang Diperbaharui

1. **Config.js**
   - Tambah: `validateProfilConfig()` function
   - Check config lengkap atau tidak

2. **code.js**
   - Tambah: `testProfilConfig()` helper
   - Untuk quick test

3. **DEBUG_PROFIL_ANALYSIS.md**
   - Analisa lengkap masalah

---

## Contact/Questions

Jika masih error:
1. Share output dari `testProfilConfig()`
2. Share output dari `testFindUser()` 
3. Buka HCIS_Config sheet & screenshot isi config
4. Buka Users sheet & screenshot header + 1 row data

Dengan info itu saya bisa fix lebih detail.

---

## Referensi Cepat

| Key | Isi Apa | Contoh |
|-----|---------|--------|
| `MASTERDATA_SS_ID` | ID spreadsheet Masterdata | `1a2b3c4d5e6f...` (optional, kosong = aktif) |
| `MASTERDATA_GID` | GID tab Masterdata | `123456789` (optional) |
| `USER_GID` | GID tab Users | `304619888` (opsional, pakai SHEET_USERS lebih simple) |
| `SHEET_MASTERDATA` | Nama tab Masterdata | `Masterdata` |
| `SHEET_USERS` | Nama tab Users | `Users` |

---

## Next: Setelah Config OK

1. Deploy code
2. Logout & login
3. Klik tab "Profil"
4. Data harus muncul
5. Done! ✅

---

Sudah coba `testProfilConfig()` dulu? Report hasilnya, saya bantu fix! 😊
