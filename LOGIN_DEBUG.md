# LOGIN TROUBLESHOOTING GUIDE

## Perbaikan yang dilakukan pada Auth.js

### 1. **Bug Hashing PIN - FIXED**
- **Masalah**: Fungsi `hashPin_()` menggunakan `(b + 256) % 256` yang salah untuk signed bytes
- **Perbaikan**: Menggunakan `b < 0 ? 256 + b : b` untuk konversi signed byte ke unsigned dengan benar
- **Impact**: PIN hash akan konsisten dan valid

### 2. **Improved Error Messages**
- Sekarang memberi error message spesifik:
  - "NIP & PIN wajib diisi" - jika form kosong
  - "Gagal membaca data user. Sheet Users mungkin tidak ada..." - jika ada error membaca sheet
  - "Data user belum tersedia di sistem." - jika Users map kosong
  - "NIP atau password salah." - jika NIP tidak ditemukan atau password tidak cocok
  - "Akun Anda tidak aktif. Hubungi admin." - jika user.aktif = false

### 3. **Better Logging**
- Added detailed logging untuk setiap step:
  - Kapan users map berhasil diload (berapa user)
  - Kapan NIP tidak ditemukan
  - Kapan password hash tidak cocok (dengan comparison value)
  - Kapan user tidak aktif

---

## Langkah Verifikasi Login

### **STEP 1: Pastikan Sheet "Users" Ada**
1. Buka spreadsheet HCIS Anda
2. Cek apakah ada sheet bernama "Users"
3. Jika tidak ada, buat sheet baru dengan nama "Users"

### **STEP 2: Pastikan Header Benar di Sheet Users**
Baris 1 harus berisi kolom (tidak harus dalam urutan ini):
- `NIP` (wajib)
- `PIN` (wajib) - ini adalah password
- `Nama` (opsional)
- `Email` (opsional)
- `Role` (opsional, default="PTK")
- `USER_ID` (opsional)
- `Aktif` (opsional, default=true)

**Contoh Header yang benar:**
```
NIP | PIN | Nama | Email | Role | USER_ID | Aktif
```

### **STEP 3: Input Data Test User**
Buat satu baris data test:
```
12345678 | 123456 | Test User | test@example.com | Admin | USR001 | TRUE
```

Pastikan:
- **NIP**: angka unik
- **PIN**: bisa kombinasi angka/huruf (ini adalah password)
- **Aktif**: harus "TRUE" atau true boolean

### **STEP 4: Clear Cache (Sangat Penting!)**
Setelah update data Users:
1. Di Apps Script Editor, buka Console (View > Logs)
2. Jalankan perintah: `clearUsersCache_()`
3. Tunggu sebentar untuk cache expire

Atau tunggu 60 detik (cache TTL default) agar otomatis clear.

### **STEP 5: Test Login Menggunakan Validation Function**
1. Di Apps Script Editor, jalankan:
   ```
   validatePin_('12345678', '123456')
   ```
   Ganti dengan NIP dan PIN test user Anda

2. Lihat hasil di Console:
   - Jika `ok: true` → PIN hash cocok, siap login
   - Jika `ok: false` → Ada masalah, lihat error message
   - Cek `hashInput` vs `hashStored` - harus sama persis

### **STEP 6: Login di Web App**
1. Deploy web app (Clasp push)
2. Buka URL web app
3. Masukkan NIP dan PIN test user
4. Klik "Masuk"

---

## Troubleshooting Common Issues

### ❌ "Gagal membaca data user. Sheet Users mungkin tidak ada..."
**Solusi:**
- Buat sheet bernama "Users" (case-sensitive)
- Pastikan header row 1 punya kolom "NIP" dan "PIN"
- Refresh cache: `clearUsersCache_()`

### ❌ "Data user belum tersedia di sistem."
**Solusi:**
- Cek sheet Users, apakah ada data user di row 2 ke bawah?
- Pastikan NIP tidak kosong di setiap row
- Pastikan PIN tidak kosong di setiap row
- Bersihkan row kosong

### ❌ "NIP atau password salah."
**Solusi:**
1. Jalankan test validation:
   ```
   validatePin_('NIP_ANDA', 'PIN_ANDA')
   ```
2. Jika output:
   - `ok: false` → PIN hash tidak cocok
     - Cek apakah PIN di sheet sudah benar
     - Coba input lagi di sheet (kemungkinan ada whitespace/karakter aneh)
   - `hashInput !== hashStored` → Ada perbedaan hashing
     - Kemungkinan bug, coba clear cache dan test ulang

### ❌ "Akun Anda tidak aktif. Hubungi admin."
**Solusi:**
- Di sheet Users, cari NIP Anda di kolom "Aktif"
- Pastikan nilai = `TRUE` atau boolean true (tidak "true" string)
- Bersihkan cache: `clearUsersCache_()`

### ❌ Login page stuck/blank
**Solusi:**
1. Buka browser DevTools (F12)
2. Cek tab Console untuk error JavaScript
3. Coba hard refresh (Ctrl+Shift+R)
4. Clear browser cache

---

## Debugging Steps untuk Technical

### 1. Check Apps Script Logs
```javascript
// Di Editor, buka View > Logs
// Akan terlihat:
// "Loaded 5 users dari sheet"
// "Login berhasil untuk NIP: 12345678"
```

### 2. Check PIN Hash
```javascript
// Hitung hash PIN manual
var hash = hashPin_('123456');
Logger.log('Hash dari PIN 123456: ' + hash);
```

### 3. Check User Map
```javascript
// Lihat isi users map
var map = loadUsersMap_();
Logger.log(JSON.stringify(map, null, 2));
```

### 4. Check Sheet Data
```javascript
// Lihat raw data dari Users sheet
var t = readTable_('Users');
Logger.log('Headers: ' + JSON.stringify(t.headers));
Logger.log('Rows: ' + JSON.stringify(t.rows.slice(0, 5))); // first 5 rows
```

---

## Deployment Checklist

- [ ] Sheet "Users" sudah dibuat
- [ ] Header row 1 punya NIP dan PIN
- [ ] Data test user sudah diinput dengan Aktif=TRUE
- [ ] Cache sudah di-clear dengan `clearUsersCache_()`
- [ ] Test validation berhasil (ok: true)
- [ ] Web app sudah di-deploy dengan `clasp push`
- [ ] Login test berhasil

---

**Last Updated:** Jan 22, 2026  
**Auth.js Version:** Fixed hash function + improved logging
