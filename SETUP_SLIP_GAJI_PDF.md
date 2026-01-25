# Setup Fitur Download Slip Gaji PDF

## Daftar Konfigurasi yang Diperlukan

Pastikan di spreadsheet **HCIS_Config** sudah ada konfigurasi berikut:

### 1. SLIP_GAJI_GID (sudah ada)
- **Key**: `SLIP_GAJI_GID`
- **Value**: GID (ID sheet) dari sheet yang berisi data Slip Gaji
- Contoh: `"135900827"`

### 2. FOLDER_SLIP (BARU - WAJIB DITAMBAH)
- **Key**: `FOLDER_SLIP`
- **Value**: Folder ID Google Drive untuk menyimpan file PDF slip gaji
- Cara mendapatkan Folder ID:
  1. Buka folder di Google Drive
  2. Lihat URL: `https://drive.google.com/drive/folders/FOLDER_ID_DI_SINI`
  3. Copy Folder ID
- Contoh: `"1a2b3c4d5e6f7g8h9i"`

## File Kop Surat

### File: `Kop_Surat.html`
- **Lokasi**: Folder utama proyek Google Apps Script
- **Isi**: Berisi base64 image dari kop surat perusahaan
- **Format**: `data:image/png;base64,[base64string]` atau `data:image/jpeg;base64,[...]`

### Cara Mengganti Kop Surat:
1. Siapkan gambar kop surat (PNG atau JPG, rekomendasi: lebar 800-1200px)
2. Convert ke base64:
   - Gunakan online tool: https://www.base64encode.org/ atau https://www.base64converter.io/
   - Upload gambar, dapatkan base64 string
3. Update file `Kop_Surat.html`:
   ```
   data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAY...
   ```

## Fitur Download Slip Gaji PDF

### Deskripsi
- User dapat download slip gaji dalam format PDF
- PDF disimpan ke folder Drive FOLDER_SLIP (bukan download langsung)
- Setelah sukses, tampil tombol "Buka File" untuk membuka PDF di Drive

### Tampilan PDF (A4 Portrait)
1. **Header**: Kop surat (dari Kop_Surat.html)
2. **Judul**: "SLIP GAJI"
3. **Periode**: "Periode: Januari 2026"
4. **Identitas Pegawai**: Nama, NIP, Unit, Jabatan
5. **Ringkasan Gaji**:
   - Total Bruto
   - Total Potongan
   - Gaji Neto (tegas/bold)
   - Gaji Prorata (jika ada)
6. **Rincian 2 Kolom**:
   - Kolom kiri: Pendapatan (Gaji Pokok, Tunj. Kinerja, dll)
   - Kolom kanan: Potongan (Kasbon, BPJS, Pendidikan Anak, dll)
7. **Tanda Tangan**: Area tanda tangan Ketua Yayasan (Lily Masngali)

### Penamaan File
Format: `Slip Gaji [Bln] [Tahun] [NIP] [Nama].pdf`

Contoh:
- `Slip Gaji Jan 2026 202009199411191071 M. Imadduddin Muqoyim.pdf`
- `Slip Gaji Feb 2026 202009199411191071 M. Imadduddin Muqoyim.pdf`

### Testing
1. Login sebagai user
2. Buka menu **Kesejahteraan** → **Slip Gaji**
3. Pilih Tahun dan Bulan dari dropdown
4. Klik **Tampilkan** (akan melihat data slip)
5. Klik **Download Slip (PDF)**
6. Tunggu beberapa detik
7. Jika sukses:
   - Muncul pesan hijau: "✅ Slip berhasil dibuat"
   - Tombol "📂 Buka File" yang bisa diklik
   - File PDF otomatis tersimpan di folder FOLDER_SLIP
8. Jika error:
   - Cek apakah FOLDER_SLIP sudah dikonfigurasi
   - Pastikan folder ID valid dan dapat diakses
   - Cek error message di UI

## Troubleshooting

### Error: "FOLDER_SLIP tidak dikonfigurasi"
- **Solusi**: Tambah key FOLDER_SLIP di HCIS_Config sheet dengan Folder ID yang benar

### Error: "Sheet Slip Gaji tidak ditemukan"
- **Solusi**: Pastikan SLIP_GAJI_GID sudah dikonfigurasi dengan benar di HCIS_Config

### PDF tidak tergenerate / Loading terus
- **Solusi**: 
  - Cek internet connection
  - Lihat Apps Script execution logs (View → Execution logs)
  - Pastikan Kop_Surat.html file sudah ada di folder proyek

### File PDF tersimpan tapi kosong / Format salah
- **Solusi**: 
  - Periksa data slip di sheet, pastikan ada nilai untuk field-field penting
  - Periksa browser console (F12) untuk error message

## Security Note

- ✅ Hanya user dengan NIP yang terdaftar bisa download slip miliknya
- ✅ Filter dilakukan di backend berdasarkan session NIP
- ✅ Tidak ada input form untuk NIP manual (mencegah akses slip orang lain)

## Fitur Lainnya (Coming Soon)

- Klaim
- Reimbursement
- Pinjaman

Fitur-fitur ini akan dikembangkan di tahap berikutnya.
