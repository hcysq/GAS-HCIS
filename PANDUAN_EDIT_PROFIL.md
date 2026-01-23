# PANDUAN FITUR EDIT PROFIL KARYAWAN

## 🎯 QUICK START

### Cara Edit Data Profil Anda

1. **Masuk Tab "Profil"** di aplikasi HCIS
2. **Temukan field yang ingin diubah** (cari di salah satu section)
3. **Klik tombol ✏️ (Edit)** di sebelah kanan nilai field
4. **Masukkan nilai baru** di modal yang muncul
5. **Klik "Lanjut Simpan"** untuk melanjutkan
6. **Verifikasi perubahan** di modal konfirmasi
7. **Klik "Ya, Simpan"** untuk menyelesaikan

---

## 📋 FIELD YANG BISA DIEDIT

### Ringkasan Kepegawaian
- Nama
- NIP
- Unit
- Jabatan
- Status Kepegawaian
- TMT (Tanggal Mulai Tugas)
- *(Masa Kerja otomatis dihitung)*

### Data Pribadi
- **NIK** ⚠️ (Sensitif - butuh persetujuan)
- TTL (Tempat, Tanggal Lahir)
- Jenis Kelamin
- Status Pernikahan
- No. Kartu Keluarga
- Ayah Kandung
- Ibu Kandung
- Gelar Akademik Depan
- Gelar Akademik Belakang
- BPJS Kesehatan
- BPJS Ketenagakerjaan
- Status PTKP
- **No. Rekening** ⚠️ (Sensitif - butuh persetujuan)
- Pendidikan Terakhir

### Kontak & Alamat
- No. HP
- WhatsApp
- Email
- Alamat
- Kelurahan / Desa
- Kecamatan
- Kabupaten / Kota
- Kode Pos
- Kontak Darurat - Nama
- Kontak Darurat - Hubungan
- Kontak Darurat - HP

*(Pendidikan Formal & Non-Formal tidak dapat diedit via fitur ini)*

---

## 🔒 FIELD SENSITIF & PERSETUJUAN

### Apa itu Field Sensitif?
Beberapa field mengandung **data pribadi yang penting** seperti:
- **NIK** - Nomor Induk Kependudukan
- **No. Rekening** - Nomor Rekening Bank

### Mekanisme Persetujuan
Saat mengedit field sensitif, Anda diminta untuk **memberikan persetujuan** sebelum data disimpan:

```
☐ Saya menyatakan data yang saya input sudah benar. 
  Jika terjadi kesalahan input yang merugikan, 
  menjadi tanggung jawab pribadi saya.
```

**Wajib centang checkbox** sebelum tombol "Ya, Simpan" bisa diklik.

---

## 📱 ALUR EDIT FIELD

### Step 1: Modal Edit
```
┌─────────────────────────────┐
│  Edit Field                 │
├─────────────────────────────┤
│  Nama Field                 │
│  [Email                   ] │
├─────────────────────────────┤
│         Nilai Baru          │
│  [  newemail@contoh.com  ] │
├─────────────────────────────┤
│  [Batal]  [Lanjut Simpan]   │
└─────────────────────────────┘
```
- Masukkan nilai baru
- Klik "Batal" jika tidak jadi
- Klik "Lanjut Simpan" untuk melanjutkan

### Step 2: Modal Konfirmasi
```
┌─────────────────────────────────┐
│  Konfirmasi Perubahan           │
├─────────────────────────────────┤
│ ⚠️  Pastikan data Anda sudah    │
│     benar. Perubahan akan      │
│     direkam dalam histori.    │
├─────────────────────────────────┤
│ Field      : Email              │
│ Nilai Lama : oldemail@xxx.com   │
│ Nilai Baru : newemail@xxx.com   │
├─────────────────────────────────┤
│  [Batal]  [Ya, Simpan]          │
└─────────────────────────────────┘
```
- Verifikasi data Anda
- Jika ada kesalahan, klik "Batal" kembali ke edit
- Jika benar, klik "Ya, Simpan"

### Step 3: Persetujuan (Jika Field Sensitif)
```
┌─────────────────────────────────┐
│  Konfirmasi Perubahan           │
├─────────────────────────────────┤
│ ... (same as above) ...         │
├─────────────────────────────────┤
│ 🔴 Data Sensitif - Perlu Persetujuan
│                                 │
│ ☐ Saya menyatakan data yang    │
│   saya input sudah benar. Jika  │
│   terjadi kesalahan input yang  │
│   merugikan, menjadi tanggung   │
│   jawab pribadi saya.           │
├─────────────────────────────────┤
│  [Batal]  [Ya, Simpan] (DISABLED)
└─────────────────────────────────┘
```
- **Centang checkbox** untuk memberikan persetujuan
- Tombol "Ya, Simpan" akan aktif setelah di-check
- Klik "Ya, Simpan" untuk menyelesaikan

---

## ✅ SETELAH PERUBAHAN DISIMPAN

### Notifikasi Sukses
```
✓ Perubahan berhasil disimpan dan dicatat dalam histori
```

### Apa yang Terjadi?
1. ✅ Data Anda di spreadsheet Users diperbarui
2. ✅ Perubahan dicatat di sheet **Histori_Mutasi** (untuk audit trail)
3. ✅ Admin dapat melihat riwayat perubahan Anda
4. ✅ Profil di aplikasi otomatis refresh menampilkan nilai baru

### Histori Perubahan
- Setiap perubahan **tidak bisa dihapus** (immutable)
- Disimpan untuk compliance & audit
- Mencakup: waktu perubahan, nilai lama, nilai baru, siapa yang edit

---

## ⚠️ TIPS & PERINGATAN

### ✓ LAKUKAN
- ✅ Baca kembali nilai sebelum klik "Ya, Simpan"
- ✅ Untuk field sensitif, pastikan data **benar-benar akurat**
- ✅ Jika ragu, tanya ke HR/Admin sebelum edit
- ✅ Gunakan fitur ini untuk **koreksi data personal** saja

### ✗ JANGAN
- ❌ Jangan edit field yang bukan milik Anda
- ❌ Jangan masukkan data fiktif atau test
- ❌ Jangan edit sebelum **benar-benar yakin**
- ❌ Jangan centang persetujuan jika belum yakin dengan data

---

## 🆘 BANTUAN & TROUBLESHOOTING

### Masalah: Tombol Edit tidak terlihat
**Solusi:**
- Refresh halaman (tekan F5)
- Pastikan Anda sudah login
- Clear browser cache (Ctrl+Shift+Delete)

### Masalah: Modal tidak muncul saat klik Edit
**Solusi:**
- Coba refresh halaman
- Gunakan browser terbaru (Chrome, Safari, Firefox)
- Jika tetap bermasalah, hubungi Admin

### Masalah: Tidak bisa klik "Ya, Simpan" untuk field sensitif
**Solusi:**
- **Wajib centang checkbox** terlebih dahulu
- Baca kalimat persetujuan dengan teliti
- Baru klik checkbox setelah Anda setuju

### Masalah: Menerima error "Nilai sama dengan sebelumnya"
**Solusi:**
- Anda memasukkan nilai yang sama seperti sebelumnya
- Ubah nilai ke sesuatu yang **berbeda**
- Klik "Batal" jika tidak jadi mengedit

### Masalah: Gagal menyimpan
**Solusi:**
- Cek koneksi internet
- Tunggu beberapa saat, coba lagi
- Jika masih gagal, screenshot error dan hubungi Admin

---

## 📝 PERTANYAAN UMUM

### Q: Apakah saya bisa mengedit NIP?
**A:** Tidak. NIP adalah nomor pegawai yang tidak boleh diubah. Jika ada kesalahan NIP, hubungi Admin HC.

### Q: Apakah saya bisa mengedit Jabatan atau Unit?
**A:** Biasanya tidak (field read-only). Perubahan Jabatan/Unit harus melalui proses HR resmi.

### Q: Berapa lama data tersimpan?
**A:** Data disimpan permanen di spreadsheet. Tidak ada penghapusan otomatis.

### Q: Siapa yang bisa melihat perubahan saya?
**A:** Admin HC dapat melihat histori perubahan Anda di sheet **Histori_Mutasi**. Data pribadi tetap aman.

### Q: Apakah saya bisa undo perubahan?
**A:** Tidak ada fitur undo otomatis. Jika salah, hubungi Admin untuk dikoreksi. Perubahan akan dicatat sebagai baris baru di histori.

### Q: Apakah field lain akan ditambah kemudian?
**A:** Ya. Admin dapat menambah field baru sesuai kebutuhan. Fitur edit akan otomatis bekerja untuk field baru.

---

## 📞 HUBUNGI ADMIN

Jika ada masalah, pertanyaan, atau saran:

**Tim HC (Human Capital)**
- 📧 Email: hc@sabilulquran.id
- 📱 WhatsApp: [hubungi admin]
- 📍 Lokasi: Ruang HR

Atau buat **issue/feedback** di aplikasi HCIS.

---

## 📌 PANDUAN ADMIN

Jika Anda adalah Admin HC dan ingin:
- ✅ Setup sheet Histori_Mutasi → Lihat: **SETUP_HISTORI_MUTASI.md**
- ✅ Pahami teknis implementasi → Lihat: **IMPLEMENTATION_PROFIL_EDIT.md**
- ✅ Edit profil pegawai lain → Fitur ini akan ditambahkan kemudian

---

**Selamat menggunakan fitur Edit Profil!** 🎉

Jika ada pertanyaan, jangan ragu untuk menghubungi Tim HC.

