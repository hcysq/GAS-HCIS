# SETUP SHEET HISTORI MUTASI

## QUICK START

### Langkah 1: Buat Sheet Baru
1. Buka spreadsheet HCIS Anda
2. Klik **"+"** di bawah tab existing (jika tidak ada, gunakan menu Sheet > Insert Sheet)
3. Nama sheet: **`Histori_Mutasi`** (case-sensitive!)
4. Klik "Create"

### Langkah 2: Setup Header (Manual atau Otomatis)

#### Opsi A: Manual Setup
Salin header berikut ke baris 1, kolom A-P:

```
A                    B            C              D             E           F
Mutasi_ID           Timestamp    Target_NIP    Target_Nama   Field_Key   Field_Label

G           H           I               J               K          L
Old_Value   New_Value   Changed_By_NIP  Changed_By_Nama Actor_Role Change_Source

M      N                O           P
Reason Consent_Checked Client_Info Request_ID
```

#### Opsi B: Otomatis (Saat Record Pertama Dicatat)
- Plugin akan auto-create header saat ada perubahan field pertama kali
- Pastikan nama sheet benar: **`Histori_Mutasi`**

### Langkah 3: Freeze Header Row (Recommended)
1. Klik baris 1 (header row)
2. Menu: **View > Freeze > 1 row**

### Langkah 4: Konfigurasi (Opsional)

Jika ingin menggunakan GID sheet daripada nama:

1. Buka sheet `Histori_Mutasi` → Copy GID dari URL:
   ```
   https://docs.google.com/spreadsheets/d/...
   /edit#gid=1234567890  ← Copy angka ini
   ```

2. Buka sheet `HCIS_Config` → Tambahkan row baru:
   ```
   Key:  HISTORI_MUTASI_GID
   Value: 1234567890  (paste GID yang dicopy)
   Note: GID sheet Histori_Mutasi
   ```

3. Klik Save (atau refresh cache)

---

## STRUKTUR LENGKAP

### Header Row (Baris 1)
| Kolom | Field Name | Tipe | Keterangan |
|-------|------------|------|-----------|
| A | `Mutasi_ID` | Text | UUID unik (auto-generate) |
| B | `Timestamp` | Datetime | ISO 8601 (auto-generate) |
| C | `Target_NIP` | Text | NIP pegawai yang diubah |
| D | `Target_Nama` | Text | Nama pegawai (opsional tapi recommended) |
| E | `Field_Key` | Text | Key dari spreadsheet (NIK, Alamat, dll) |
| F | `Field_Label` | Text | Label user-friendly (NIK, Alamat, dll) |
| G | `Old_Value` | Text | Nilai sebelumnya |
| H | `New_Value` | Text | Nilai sesudahnya |
| I | `Changed_By_NIP` | Text | NIP aktor yang edit |
| J | `Changed_By_Nama` | Text | Nama aktor |
| K | `Actor_Role` | Text | "pegawai" atau "admin" |
| L | `Change_Source` | Text | "profil_edit" atau "admin_panel" |
| M | `Reason` | Text | Alasan (opsional) |
| N | `Consent_Checked` | Text | TRUE atau FALSE |
| O | `Client_Info` | Text | Device info (opsional) |
| P | `Request_ID` | Text | Debug trace ID (opsional) |

### Contoh Data Baris 2-4

**Baris 2: Edit Alamat (pegawai, no consent)**
```
A: d5f4a3b2-1c0f-4e5d-9a8b-7c6f5e4d3c2b
B: 2026-01-23T14:30:45.123Z
C: 198701151234
D: John Doe
E: Alamat
F: Alamat
G: Jl. Merdeka No. 123
H: Jl. Sudirman No. 456
I: 198701151234
J: John Doe
K: pegawai
L: profil_edit
M: (kosong)
N: FALSE
O: (kosong)
P: (kosong)
```

**Baris 3: Edit NIK (pegawai dengan consent)**
```
A: e6g5b4c3-2d1g-5f6e-0b9c-8d7g6f5e4d3c
B: 2026-01-23T15:45:20.456Z
C: 198701151234
D: John Doe
E: NIK
F: NIK
G: 1234567890123456 (atau ***-***-***-3456)
H: 1234567890654321 (atau ***-***-***-1111)
I: 198701151234
J: John Doe
K: pegawai
L: profil_edit
M: (kosong)
N: TRUE
O: (kosong)
P: (kosong)
```

---

## VALIDASI & TESTING

### Test 1: Cek Header Exist
```
1. Buka sheet Histori_Mutasi
2. Baris 1 harus punya kolom A-P
✓ Header tersedia
```

### Test 2: Edit Field → Catat Histori
```
1. Masuk Tab Profil
2. Edit field "Email"
3. Lanjut → Ya, Simpan
4. Buka sheet Histori_Mutasi
5. Baris 2 harus ada record baru
✓ Record tercatat dengan Mutasi_ID, Timestamp, values
```

### Test 3: Sensitif Field → Consent
```
1. Edit field "NIK"
2. Lanjut → Confirm modal menampilkan consent box
3. Centang checkbox
4. Ya, Simpan
5. Buka sheet Histori_Mutasi
6. Cek Consent_Checked = TRUE
✓ Consent tercatat
```

### Test 4: Immutable Check
```
1. Buka Histori_Mutasi
2. Coba edit cell di row 2
3. (Dalam future versi, ini akan diblock)
❌ Di Google Sheets, sheet ini editable
📌 Rekomendasi: Proteksi sheet dengan permission read-only untuk pegawai
```

---

## PROTEKSI SHEET (RECOMMENDED)

### Proteksi Range (Header + Data)
1. Pilih semua kolom A-P (atau semua data)
2. Menu: **Data > Protect sheets and ranges**
3. Opsi:
   - **"Restrict who can edit this range"**
   - **"Only you"** (atau restricted editors)
4. Klik "Done"

### Opsi Admin Edit
Jika ingin admin bisa edit histori (untuk correction):
1. Sebelum protect, klik **"Data > Protect sheets and ranges"** 
2. Pilih **"Only these users can edit"**
3. Tambahkan email admin
4. OK

---

## TROUBLESHOOTING

### Masalah: Sheet tidak terdeteksi
```
Error: "Sheet Histori_Mutasi tidak ditemukan pada spreadsheet aktif."
```
**Solution:**
- Pastikan nama sheet EXACTLY: `Histori_Mutasi` (case-sensitive)
- Jangan pakai spasi ekstra atau underscore di awal/akhir
- Coba rename sheet jika salah

### Masalah: "Header tidak lengkap"
```
Error: "Header tidak sesuai di sheet Histori_Mutasi"
```
**Solution:**
- Buka sheet Histori_Mutasi
- Baris 1 harus punya 16 kolom (A-P) sesuai struktur
- Jika ada kesalahan, delete baris 1 dan ulang setup

### Masalah: GID tidak valid
```
Error: "Sheet dengan GID ... tidak ditemukan"
```
**Solution:**
- Buka Histori_Mutasi sheet
- Cek GID dari URL: `...#gid=1234567890`
- Update HISTORI_MUTASI_GID di Config sheet
- Clear cache: `cfgClearCache()`

---

## BEST PRACTICES

### 1. Backup Rutin
- Export Histori_Mutasi setiap bulan ke CSV
- Simpan di cloud storage (untuk audit trail)

### 2. Cleanup Lama
- Histori bersifat append-only (immutable)
- Jangan hapus row lama (untuk compliance)
- Gunakan filter/sort untuk viewing saja

### 3. Report & Analytics
- Buat pivot table untuk stat mutasi per pegawai
- Monitor field mana saja yang sering diubah
- Flag unusual patterns (banyak edit dalam waktu singkat)

### 4. Permission Model
```
Pegawai:
- Read: Histori_Mutasi (view own records)
- Write: Users sheet (via profil edit only)

Admin:
- Read/Write: HCIS_Config, Histori_Mutasi
- Read: Users sheet
- Management: Approve/reject batch edits (future)
```

---

## API INTEGRATION (Untuk Custom Tools)

Jika ingin query Histori_Mutasi dari script lain:

```javascript
function getHistoriByNIP(nip) {
  const { sheet: sh } = getHistoriMutasiSheet_();
  if (!sh) return [];
  
  const data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  return data.filter(row => row[2] === nip); // Kolom C = Target_NIP
}

function getHistoriByField(fieldKey) {
  const { sheet: sh } = getHistoriMutasiSheet_();
  if (!sh) return [];
  
  const data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  return data.filter(row => row[4] === fieldKey); // Kolom E = Field_Key
}
```

---

## FIQH DATA & COMPLIANCE

### Audit Trail Immutable
- Setiap perubahan data sensitif tercatat
- Tidak boleh dihapus (compliance requirement)
- Timestamp ISO 8601 untuk rekam waktu akurat
- Konsisten dengan regulasi data protection

### Consent & Accountability
- Field sensitif (NIK, Rekening) wajib checkbox
- Bukti consent tercatat di `Consent_Checked`
- Actor identitas jelas (NIP + Nama)
- Alasan edit bisa dicatat oleh admin

### Data Minimization
- Hanya catat old/new value (bukan whole record)
- Client info opsional (reduce privacy risk)
- Request ID opsional (untuk debug saja)

---

**Setup Selesai!** ✅

Sheet Histori_Mutasi siap menerima pencatatan perubahan profil.

