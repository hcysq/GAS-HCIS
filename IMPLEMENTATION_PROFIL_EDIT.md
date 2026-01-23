# IMPLEMENTASI FITUR EDIT PER-FIELD DI TAB PROFIL

## RINGKASAN
Fitur baru memungkinkan pegawai untuk mengedit field profil mereka secara mandiri dengan sistem safeguard yang ketat:
1. **Edit Per-Field**: Setiap field dapat diubah secara individual
2. **Konfirmasi Modal**: Dialog konfirmasi menampilkan perubahan sebelum disimpan
3. **Consent untuk Data Sensitif**: Field sensitif (NIK, No. Rekening) memerlukan persetujuan khusus
4. **Histori Mutasi Immutable**: Semua perubahan dicatat dalam tab histori mutasi (append-only)

---

## FITUR UTAMA

### A. UI/UX - EDIT PER-FIELD

#### Tampilan Default (Read-Only)
- Setiap field ditampilkan dengan:
  - Label field (kiri)
  - Nilai field (tengah)
  - **Tombol Edit (✏️ kanan)** - icon pensil kecil

#### Alur Edit Field
1. Klik tombol **✏️ Edit** pada field yang ingin diubah
2. Modal pertama muncul: **"Edit Field"**
   - Menampilkan nama field
   - Input box untuk nilai baru
   - Tombol: "Batal" dan "Lanjut Simpan"
3. Hanya **1 field dapat diedit pada saat yang sama** (prevent conflicts)

#### Tipe Input Otomatis
Berdasarkan field key:
- **Text** (default): Nama, Alamat, Ayah/Ibu, Gelar, Email, HP
- **Number**: NIK, No_KK, BPJS_Kes, BPJS_TK, Kode_Pos, No._Rekening
- **Date**: TMT (TMT_Pegawai), Tanggal_Lahir
- **Select** (read-only di profil): Jenis_Kelamin, Status_Nikah, Status_PTKP, Status_Kepeg
- **Textarea**: Alamat (untuk text panjang)

---

### B. KONFIRMASI SEBELUM SIMPAN

#### Modal Konfirmasi
Setelah klik "Lanjut Simpan" di edit modal, muncul dialog **"Konfirmasi Perubahan"** dengan:

```
┌─────────────────────────────────────┐
│   Konfirmasi Perubahan              │
├─────────────────────────────────────┤
│ ⚠️  Pastikan data yang Anda masukkan │
│     sudah benar. Perubahan akan     │
│     direkam dalam histori mutasi.   │
├─────────────────────────────────────┤
│ Field         : [Field Name]        │
│ Nilai Lama    : [Old Value]         │
│ Nilai Baru    : [New Value]         │
├─────────────────────────────────────┤
│ [Batal]  [Ya, Simpan]               │
└─────────────────────────────────────┘
```

**Aksi:**
- **Batal**: Kembali ke modal edit (tidak menyimpan)
- **Ya, Simpan**: Lanjut proses (dengan consent check jika sensitif)

---

### C. CONSENT KHUSUS UNTUK DATA SENSITIF

#### Field Sensitif
Dua field ditetapkan sebagai sensitif:
1. **NIK** (key: `NIK`)
2. **No. Rekening** (key: `No._Rekening`)

#### Mekanisme Consent
Ketika field sensitif diedit, modal konfirmasi menampilkan **checkbox consent**:

```
┌─────────────────────────────────────────┐
│   [Konfirmasi Perubahan]                │
├─────────────────────────────────────────┤
│ ... (field info seperti di atas) ...    │
├─────────────────────────────────────────┤
│ 🔴 Data Sensitif - Perlu Persetujuan    │
│                                         │
│ ☐ Saya menyatakan data yang saya       │
│   input sudah benar. Jika terjadi       │
│   kesalahan input yang merugikan,       │
│   menjadi tanggung jawab pribadi saya.  │
├─────────────────────────────────────────┤
│ [Batal]  [Ya, Simpan] (DISABLED)        │
└─────────────────────────────────────────┘
```

**Validasi:**
- Tombol **"Ya, Simpan" DISABLED sampai checkbox dicentang**
- Jika field sensitif, consent harus dicek sebelum bisa save
- Checkbox hanya muncul untuk field sensitif

---

### D. HISTORI MUTASI - PENCATATAN PERUBAHAN

#### Aturan Pencatatan
- **Append-Only**: Setiap perubahan dicatat sebagai baris baru (tidak ada delete/edit histori)
- **Mandatory Fields**: Semua field dalam tabel histori wajib terisi
- **Immutable**: Histori tidak dapat diubah setelah dicatat

#### Struktur Tabel `Histori_Mutasi`
Sheet baru bernama **`Histori_Mutasi`** dengan kolom:

| # | Kolom | Tipe | Deskripsi |
|---|-------|------|-----------|
| A | `Mutasi_ID` | UUID | Unique identifier (auto-generate: xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx) |
| B | `Timestamp` | ISO DateTime | Waktu perubahan (ISO 8601, auto-generate) |
| C | `Target_NIP` | Text | NIP pegawai yang datanya berubah |
| D | `Target_Nama` | Text | Nama pegawai (untuk baca cepat, opsional) |
| E | `Field_Key` | Text | Key field dari spreadsheet (mis: NIK, No._Rekening, Alamat) |
| F | `Field_Label` | Text | Label field untuk user (mis: "NIK", "No. Rekening") |
| G | `Old_Value` | Text | Nilai lama (stringified) |
| H | `New_Value` | Text | Nilai baru (stringified) |
| I | `Changed_By_NIP` | Text | NIP aktor yang melakukan perubahan |
| J | `Changed_By_Nama` | Text | Nama aktor (untuk baca cepat) |
| K | `Actor_Role` | Text | Role aktor: `pegawai` atau `admin` |
| L | `Change_Source` | Text | Sumber perubahan: `profil_edit` (atau `admin_panel` jika admin) |
| M | `Reason` | Text | Alasan perubahan (opsional, terutama untuk admin) |
| N | `Consent_Checked` | Boolean | TRUE/FALSE - apakah checkbox consent dicek (untuk sensitif fields) |
| O | `Client_Info` | Text | Info device/browser (opsional) |
| P | `Request_ID` | Text | Trace ID untuk debugging (opsional) |

#### Contoh Record Histori

**Perubahan NIK oleh Pegawai (dengan Consent):**
```
Mutasi_ID:       d5f4a3b2-1c0f-4e5d-9a8b-7c6f5e4d3c2b
Timestamp:       2026-01-23T14:30:45.123Z
Target_NIP:      198701151234
Target_Nama:     John Doe
Field_Key:       NIK
Field_Label:     NIK
Old_Value:       3274XXXXXXXX5678
New_Value:       3274YYYYYYYY5678
Changed_By_NIP:  198701151234
Changed_By_Nama: John Doe
Actor_Role:      pegawai
Change_Source:   profil_edit
Reason:          (kosong)
Consent_Checked: TRUE
Client_Info:     (kosong)
Request_ID:      (kosong)
```

**Perubahan Alamat oleh Pegawai (tanpa Consent):**
```
Mutasi_ID:       e6g5b4c3-2d1g-5f6e-0b9c-8d7g6f5e4d3c
Timestamp:       2026-01-23T15:45:20.456Z
Target_NIP:      198701151234
Target_Nama:     John Doe
Field_Key:       Alamat
Field_Label:     Alamat
Old_Value:       Jl. Merdeka No. 123
New_Value:       Jl. Sudirman No. 456
Changed_By_NIP:  198701151234
Changed_By_Nama: John Doe
Actor_Role:      pegawai
Change_Source:   profil_edit
Reason:          (kosong)
Consent_Checked: FALSE
Client_Info:     (kosong)
Request_ID:      (kosong)
```

#### Cara Akses Histori
- **Sheet Tab**: Buka tab **"Histori_Mutasi"** di spreadsheet HCIS
- **Sorting**: Gunakan header row yang frozen untuk sort/filter
- **Report**: Setiap baris adalah immutable record dari perubahan

---

## KONFIGURASI (CONFIG.JS)

Tambahkan ke file `Config.js`:

```javascript
// HISTORI_MUTASI_GID
// GID sheet Histori_Mutasi untuk pencatatan perubahan field profil
cfgSet('HISTORI_MUTASI_GID', '1234567890', 'GID sheet Histori_Mutasi (opsional jika pakai nama sheet)');
```

Atau gunakan nama sheet default: **`Histori_Mutasi`**

---

## BACKEND FUNCTIONS

### 1. `logProfilMutation_(params)`
Catat perubahan field ke sheet Histori_Mutasi.

**Parameter:**
```javascript
{
  target_nip: "198701151234",           // NIP pegawai
  target_nama: "John Doe",              // Nama pegawai
  field_key: "NIK",                     // Key field
  field_label: "NIK",                   // Label field
  old_value: "1234567890123456",        // Nilai lama
  new_value: "1234567890654321",        // Nilai baru
  changed_by_nip: "198701151234",       // NIP aktor
  changed_by_nama: "John Doe",          // Nama aktor
  actor_role: "pegawai",                // "pegawai" atau "admin"
  consent_checked: true,                // boolean
  reason: ""                            // (opsional)
}
```

**Return:**
```javascript
{ ok: true, msg: "...", mutasi_id: "..." }
```

### 2. `saveProfilFieldChange(params)`
Simpan perubahan field ke Users sheet DAN catat ke Histori_Mutasi.

**Parameter:**
```javascript
{
  field_key: "NIK",
  field_label: "NIK",
  old_value: "...",
  new_value: "...",
  consent_checked: true
}
```

**Return:**
```javascript
{ ok: true, msg: "...", mutasi_id: "..." }
```

**Validasi:**
- Field sensitif wajib ada `consent_checked: true`
- Perubahan dicatat ke histori otomatis

---

## BATASAN & ATURAN

### Yang TIDAK Boleh Diubah ✗
- Mapping key spreadsheet (kolom header)
- Urutan field di profil
- Struktur halaman Profil
- Logic existing selain tambah edit+save+logging
- Dashboard, Setelan, fitur lain

### Yang BISA Diubah ✓
- Nilai field profil (sesuai tipe data)
- Field baru (jika ada di spreadsheet Users)
- Tampilan form edit (styling/UX)
- Logika consent dan konfirmasi

### Data yang TIDAK Bisa Diedit
- NIP (tidak ada edit button)
- Unit, Jabatan, Status Kepegawaian, TMT (usually)
- Data pendidikan (tabel statis)
- Masa Kerja (auto-calculated)

---

## TESTING & VALIDASI

### Test Cases

#### 1. Edit Field Non-Sensitif
```
1. Masuk ke Tab Profil
2. Klik Edit ✏️ di field "Email"
3. Input: "newemail@contoh.com"
4. Klik "Lanjut Simpan"
5. Confirm modal muncul (NO consent box)
6. Klik "Ya, Simpan"
✓ Email berubah, histori tercatat, NO consent_checked
```

#### 2. Edit Field Sensitif (NIK)
```
1. Klik Edit ✏️ di field "NIK"
2. Input: NIK baru
3. Klik "Lanjut Simpan"
4. Confirm modal muncul + CONSENT BOX muncul
5. Checkbox unchecked → button DISABLED
6. Check checkbox → button ENABLED
7. Klik "Ya, Simpan"
✓ NIK berubah, histori tercatat, consent_checked: TRUE
```

#### 3. Cancel Edit
```
1. Klik Edit ✏️
2. Modal muncul
3. Klik "Batal"
✓ Modal tutup, tidak ada perubahan
```

#### 4. Histori Validation
```
1. Buka sheet Histori_Mutasi
2. Cek baris terakhir
✓ Semua kolom terisi, Mutasi_ID unique, Timestamp valid
```

---

## TROUBLESHOOTING

### Problem: Modal tidak muncul
**Solution**: 
- Pastikan `index.html` include `app.html` dan `style.html`
- Check browser console untuk error

### Problem: "Sheet Histori_Mutasi tidak ditemukan"
**Solution**:
- Buat sheet baru bernama "Histori_Mutasi" (atau set HISTORI_MUTASI_GID di config)
- Biarkan plugin auto-create header pada record pertama

### Problem: Consent checkbox tidak muncul
**Solution**:
- Field harus di `getSensitiveFieldsList_()`: `['NIK', 'No._Rekening']`
- Rebuild profil atau hard refresh browser

### Problem: Nilai berubah tapi tidak tersimpan
**Solution**:
- Check Users sheet (Users GID/name di config)
- Cek apakah field key match dengan header sheet
- Lihat browser console untuk error detail

---

## FUTURE ENHANCEMENTS

1. **Admin Edit Panel**: Interface khusus admin untuk edit profil pegawai lain
2. **Audit Report**: Dashboard untuk melihat histori mutasi per pegawai
3. **Bulk Edit**: Edit multiple fields sekaligus dengan approval workflow
4. **Photo Upload**: Support upload foto profil dengan validation
5. **Field-Level Permissions**: Kontrol field mana saja yang boleh diedit pegawai vs admin

---

## CATATAN PENGEMBANG

### File yang Diubah
1. **Config.js** - Helper untuk buka sheet Histori_Mutasi
2. **Profile.js** - Backend functions: `logProfilMutation_()`, `saveProfilFieldChange()`
3. **app.html** - UI/UX edit modal, confirm dialog, JavaScript handler
4. **style.html** - CSS untuk modal, input field, buttons, animations

### Key Functions
- `startEditField()` - Buka edit modal
- `confirmEditField()` - Validate input, lanjut ke confirm
- `showConfirmChangeModal()` - Display konfirmasi
- `saveFieldChange()` - Backend call untuk simpan
- `getSensitiveFieldsList_()` - List field sensitif
- `getFieldType_()` - Determine input type

### Global State
```javascript
window.editFieldState = {
  fieldKey: "...",
  fieldLabel: "...",
  oldValue: "...",
  newValue: "...",
  fieldType: "text",
  isSensitive: false
}
```

---

## CHECKLISTA IMPLEMENTASI

- [x] Edit button (✏️) untuk setiap field
- [x] Modal edit pertama (input field + buttons)
- [x] Modal konfirmasi (old/new values + peringatan)
- [x] Consent checkbox untuk field sensitif
- [x] Backend save function
- [x] Histori mutasi logging
- [x] Config untuk HISTORI_MUTASI_GID
- [x] UI/UX styling (modals, buttons, animations)
- [x] Validation (field wajib, tipe data)
- [x] Error handling & messages
- [x] Documentation

---

**Status**: ✅ SIAP PRODUCTION

