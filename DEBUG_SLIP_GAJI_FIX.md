# DEBUG & FIX: MODUL SLIP GAJI - DYNAMIC FILTER + DATA MAPPING

**Date**: 25 Januari 2026  
**Issue**: Filter tahun/bulan statis, data tidak ditampilkan  
**Status**: ✅ FIXED

---

## 🔧 PERUBAHAN YANG DILAKUKAN

### 1. **Dynamic Filter (Ambil dari Data Aktual)**

**Function Baru: `getAvailableSlipGajiBulan()`**
- Scan sheet Slip Gaji untuk NIP user yang login
- Kumpulkan list unik Bulan (format "NamaBulan Tahun")
- Extract tahun dan bulan yang tersedia
- Return: `{ tahunList, bulanPerTahun }`

**Contoh Output:**
```json
{
  "tahunList": ["2023", "2024", "2025", "2026"],
  "bulanPerTahun": {
    "2023": ["Januari", "Februari", "Maret", ...],
    "2024": ["Januari", "Februari", ...],
    ...
  }
}
```

**Frontend Flow:**
1. Saat buka Slip Gaji → call `getAvailableSlipGajiBulan()`
2. Populate dropdown Tahun dengan list yang ada
3. Saat pilih Tahun → dynamically update dropdown Bulan
4. Tombol "Tampilkan" hanya enabled jika ada pilihan untuk keduanya

---

### 2. **Robust Data Mapping**

**Update: `buildSlipGajiPayload_()`**
- Tambah multiple column name options untuk setiap field
- Contoh: cari `GAJI NETO` atau `Gaji Neto` atau `GAJI NETO 80%`
- Lebih fleksibel terhadap variasi header

**Peningkatan Parsing:**
- `getNum()` sekarang strip non-numeric characters saat parsing
- Handling untuk nilai kosong/0 → return "-"
- Sanitize dengan check untuk `'0'` string

**Contoh Mapping:**
```javascript
gajiNeto: getNum(['GAJI NETO', 'Gaji Neto', 'GAJI NETO 80%', 'Gaji Netto 80%'])
```

---

### 3. **Update Function Calls**

**`tampilkanSlipGaji()`:**
- Pass `bulan` sebagai string (nama bulan), bukan angka
- Contoh: `getSlipGaji(2026, 'Januari')` ✅
- Bukan: `getSlipGaji(2026, 1)` ❌

**Backend `getSlipGaji()`:**
- Accept bulan sebagai number OR string
- Auto-convert jika number ke nama bulan
- Match dengan kolom "Bulan" di sheet

---

## 📋 SETUP YANG DIPERLUKAN USER

**WAJIB: Tambahkan Config**

User harus manualmenambah ke spreadsheet HCIS_Config:

```
Key            | Value              | Note
SLIP_GAJI_GID  | [GID dari sheet]   | Tab sheet slip gaji
```

**Cara mencari GID:**
1. Buka Payroll SDM YSQ spreadsheet
2. Tab "Slip_Gaji" → klik kanan
3. "Open link" → copy GID dari URL
4. Contoh GID: `135900827`

---

## 🧪 TESTING CHECKLIST

Sebelum declare done:

- [ ] User login dengan akun yang punya data dari Januari 2023+
- [ ] Buka menu Kesejahteraan → Slip Gaji
- [ ] Dropdown Tahun hanya tampil: 2023, 2024, 2025, 2026
- [ ] Pilih Tahun 2026 → Dropdown Bulan tampil bulan yang ada di 2026
- [ ] Pilih Tahun 2023 → Dropdown Bulan tampil bulan 2023 (sesuai data)
- [ ] Tombol "Tampilkan" disabled sampai pilih kedua-duanya
- [ ] Pilih periode yang ada data (contoh: Januari 2026) → tampil slip
- [ ] Verifikasi data yang tampil:
  - Nama, NIP, Unit, Jabatan cocok
  - Angka Utama (Gaji Neto, Total Bruto, Total Potongan) ada nilai
  - Rincian Pendapatan & Potongan sesuai
- [ ] Buka accordion Detail Lainnya → data lengkap
- [ ] Pilih periode yang TIDAK ada data → "Slip gaji periode ini belum tersedia"

---

## 📊 STRUKTUR DATA SPREADSHEET

**Header Row 1 (yang digunakan):**
```
KEY | NO URUT | Bulan | Tanggal | NIP | NAMA | UNIT | ... | GAJI POKOK | ... | TOTAL BRUTO GAJI | GAJI NETO | ... | GAJI PRORATA
```

**Format Bulan:**
- `"Januari 2026"`, `"Februari 2026"`, dst
- Bukan: `"1/2026"` atau `"01 2026"`

**Format Tanggal:**
- Standard date format (MM/DD/YYYY atau similar)
- Digunakan untuk sorting jika ada >1 baris periode sama

---

## 🎯 KEY IMPROVEMENTS

| Aspek | Sebelum | Sesudah |
|-------|---------|---------|
| Filter Tahun/Bulan | Static (hardcoded 5 tahun range) | Dynamic (scan actual data) |
| Bulan Input | Angka 1-12 | String "Januari"-"Desember" atau angka |
| Column Mapping | Strict (exact match) | Flexible (multiple name options) |
| Number Parsing | Simple `Number()` | Strip non-numeric, handle 0 |
| Error Feedback | Generic "tidak ditemukan" | Specific (tahun/bulan mana) |

---

## 🚀 NEXT STEPS

1. **User setup config** → Add `SLIP_GAJI_GID` ke HCIS_Config
2. **Deploy** → clasp deploy
3. **Test** → Ikuti testing checklist
4. **Report** → Jika ada error, share screenshot error message

---

## 📝 NOTES

- Filter hanya based on NIP user yang login
- Data read-only (tidak bisa edit dari UI)
- Jika ada >1 baris untuk periode sama → ambil yang paling terbaru
- Semua angka format Rupiah di frontend
- Nilai kosong/0 tampil sebagai "-" bukan "Rp 0"
