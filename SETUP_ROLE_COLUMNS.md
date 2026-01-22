# Setup Kolom Role di Users Sheet

## Struktur Baru (Lebih Simple!)

Daripada menggunakan kolom `Role` dengan string "PTK,ADMIN,KAPLA", kita pakai **3 kolom boolean terpisah**:

| NIP | Nama | Email | PTK | ADMIN | KAPLA | ... |
|-----|------|-------|-----|-------|-------|-----|
| 100 | Budi | ... | TRUE | FALSE | FALSE | ... |
| 200 | Andi | ... | TRUE | TRUE | FALSE | ... |
| 300 | Citra | ... | TRUE | FALSE | TRUE | ... |
| 666 | Super | ... | TRUE | TRUE | TRUE | ... |

## Langkah Setup

### Step 1: Buka Users Sheet
1. Buka spreadsheet HCIS Anda
2. Cari sheet "Users"

### Step 2: Tambah 3 Kolom Baru
Tambahkan di kolom mana aja (sebelum atau sesudah kolom lainnya):
- `PTK` 
- `ADMIN`
- `KAPLA`

### Step 3: Isi Nilai Boolean
Untuk setiap user:
- `PTK` = **selalu TRUE** (semua user adalah PTK)
- `ADMIN` = TRUE jika user adalah admin, FALSE jika bukan
- `KAPLA` = TRUE jika user adalah kepala unit, FALSE jika bukan

### Step 4: Contoh Setup

```
User Biasa (hanya PTK):
- NIP: 100
- Nama: Budi
- PTK: TRUE
- ADMIN: FALSE
- KAPLA: FALSE

User Admin:
- NIP: 200
- Nama: Andi
- PTK: TRUE
- ADMIN: TRUE
- KAPLA: FALSE

User Kepala Unit:
- NIP: 300
- Nama: Citra
- PTK: TRUE
- ADMIN: FALSE
- KAPLA: TRUE

Super User (PTK + ADMIN + KAPLA):
- NIP: 666
- Nama: Super Admin
- PTK: TRUE
- ADMIN: TRUE
- KAPLA: TRUE
```

## Migrasi dari Format Lama

Jika sebelumnya pakai kolom `Role` dengan string:

### Dari:
```
Role = "PTK"
Role = "PTK,ADMIN"
Role = "PTK,KAPLA"
Role = "PTK,ADMIN,KAPLA"
```

### Ke:
```
PTK = TRUE, ADMIN = FALSE, KAPLA = FALSE
PTK = TRUE, ADMIN = TRUE, KAPLA = FALSE
PTK = TRUE, ADMIN = FALSE, KAPLA = TRUE
PTK = TRUE, ADMIN = TRUE, KAPLA = TRUE
```

## Keuntungan Sistem Baru

✅ **Lebih Simple** - Tidak perlu parsing string  
✅ **Lebih Aman** - Tidak ada typo atau format error  
✅ **Lebih Jelas** - Langsung terlihat role apa aja yang dimiliki user  
✅ **Lebih Cepat** - Langsung baca boolean, tidak perlu parsing  
✅ **Lebih Fleksibel** - Mudah menambah role baru nanti  

## Sistem Baca Data

Sistem sekarang:
1. User login dengan NIP + Password
2. Cari user di Users sheet
3. Baca kolom `PTK`, `ADMIN`, `KAPLA`
4. Convert ke roles array: ['PTK'] atau ['PTK','ADMIN'] atau ['PTK','KAPLA'] atau ['PTK','ADMIN','KAPLA']
5. Simpan di session
6. Dashboard tampilkan roles dan panel yang sesuai

Jadi **tidak perlu kolom `Role` lagi** jika sudah pakai 3 kolom boolean ini.

## Testing

Setelah setup kolom:

1. Buat user test dengan NIP 999
2. Isi: Nama="Test User", PTK=TRUE, ADMIN=TRUE, KAPLA=TRUE
3. Login dengan NIP 999
4. Dashboard harus menunjukkan:
   - "Role: PTK, ADMIN, KAPLA"
   - Kedua panel (Admin dan Team) visible

## Jika Ada Pertanyaan

- **Kapan PTK isi TRUE?** → Selalu, untuk semua user
- **Bisa pakai nilai lain?** → Cukup TRUE/FALSE atau bisa juga 1/0
- **Nama kolom harus persis?** → Ya, persis: `PTK`, `ADMIN`, `KAPLA` (case-sensitive)
- **Posisi kolom penting?** → Tidak, bisa di mana aja di sheet
