# Update Slip Gaji: Google Docs Template + Salary Display Rules

**Date**: January 25, 2026  
**Status**: ✅ COMPLETED & DEPLOYED

---

## Overview

Major update to Slip Gaji module switching from **HTML-based PDF** to **Google Docs Template-based PDF generation** with new naming conventions and salary display rules.

---

## 🔄 Key Changes

### 1. PDF Generation Method
- **Before**: HTML template → convert to PDF via Utilities.newBlob
- **After**: Google Docs Template → copy + replace placeholders → export to PDF
- **Benefit**: Better formatting control, easier to maintain layout

### 2. File Naming Convention
- **Before**: `Slip Gaji Jan 2026 202009199411191071 M. Imadduddin Muqoyim.pdf`
- **After**: `SlipGaji_202601_202009199411191071.pdf`
- **Benefits**: 
  - Consistent, parsing-friendly naming
  - No special characters
  - Easy to sort/search by period + NIP
  - Supports duplicate detection

### 3. Salary Display Rules
Implemented smart display logic:
1. If **GAJI PRORATA > 0** → Show: `Rp X (Prorata)`
2. Else if **GAJI NETTO 80% > 0** → Show: `Rp X (80%)`
3. Else → Show: `Rp X`

Applied to:
- PDF Slip Gaji templates
- UI Slip Gaji display
- Profile tab (future enhancement)

### 4. Duplicate Prevention
- Check if file `SlipGaji_YYYYMM_NIP.pdf` already exists in `FOLDER_SLIP`
- If exists: Show warning, **do NOT regenerate**
- Message: "Slip gaji periode ini sudah dibuat. Untuk permintaan ulang, silakan hubungi admin HCM."

### 5. UX Change
- Button: **"📧 Kirim Slip (PDF)"** (instead of "Download")
- Behavior: Generate + Save to Drive (no direct download)
- Message: "Slip berhasil dibuat dan dikirim ke email Anda. Admin HCM juga menerima notifikasi."

---

## ⚙️ Configuration Required

### 1. TEMPLATE_SLIP (REQUIRED)
```
Key: TEMPLATE_SLIP
Value: 1BgvBGV0hNF4G9AqG44b7Noz5XN-XenEU-5OL15IXqIg
```
**Add to HCIS_Config spreadsheet**

### 2. FOLDER_SLIP (Already exists)
```
Key: FOLDER_SLIP
Value: Your Google Drive Folder ID
```

### 3. SLIP_GAJI_GID (Already exists)
```
Key: SLIP_GAJI_GID
Value: Sheet GID for Slip Gaji data
```

---

## 📋 Google Docs Template Placeholders

### Identitas
```
{{PERIODE}}    - Periode (contoh: "Januari 2026")
{{NAMA}}       - Nama lengkap pegawai
{{NIP}}        - Nomor Induk Pegawai
{{UNIT}}       - Unit/Departemen
{{JABATAN}}    - Jabatan (gabungan: JABATAN / JABATAN FUNGSIONAL / JABATAN STRUKTURAL)
```

### Ringkasan Gaji
```
{{TOTAL_BRUTO}}      - Total Gaji Bruto
{{TOTAL_POTONGAN}}   - Total Potongan
{{GAJI_NETO}}        - Gaji Utama (mengikuti salary display rule)
```

### Pendapatan
```
{{GAJI_POKOK}}         - Gaji Pokok
{{TUNJ_KINERJA}}       - Tunjangan Kinerja
{{TUNJ_ISTRI}}         - Tunjangan Istri
{{TUNJ_ANAK}}          - Tunjangan Anak
{{TUNJ_FUNGSIONAL}}    - Tunjangan Fungsional
{{TUNJ_JABATAN}}       - Tunjangan Jabatan
{{TUNJ_KUALIFIKASI}}   - Tunjangan Kualifikasi Khusus
{{LEMBUR}}             - Lembur
{{RAPEL_GAJI}}         - Rapel Gaji
{{TUNJ_BPJS}}          - Tunjangan BPJS
```

### Potongan
```
{{POT_KASBON}}      - Potongan Kasbon
{{BPJS}}            - BPJS (Potongan)
{{PEND_ANAK}}       - Pendidikan Anak
{{KURANG_JAM}}      - Kekurangan Jam
{{BPJS_JHT}}        - BPJS TK (JHT)
{{BPJS_JP}}         - BPJS TK (JP)
{{PPH21}}           - PPH21
{{POT_ABSENSI}}     - Potongan Absensi
```

**Note**: All placeholders formatted as **Rupiah** (Rp X.XXX)

---

## 🗂️ Data Mapping

### Spreadsheet Columns → Placeholder
Data pulled from **SLIP_GAJI_GID** sheet:

| Sheet Column | Placeholder | Type | Notes |
|---|---|---|---|
| PERIODE / Bulan | {{PERIODE}} | Text | Format: "Januari 2026" |
| NAMA / Nama | {{NAMA}} | Text | |
| NIP | {{NIP}} | Text | |
| UNIT / Unit | {{UNIT}} | Text | |
| JABATAN | {{JABATAN}} | Text | Combination of 3 columns |
| JABATAN STRUKTURAL | (part of JABATAN) | Text | |
| JABATAN FUNGSIONAL | (part of JABATAN) | Text | |
| TOTAL BRUTO GAJI | {{TOTAL_BRUTO}} | Number | |
| TOTAL POTONGAN | {{TOTAL_POTONGAN}} | Number | |
| GAJI NETO / GAJI PRORATA | {{GAJI_NETO}} | Number | Smart rule applied |
| GAJI PRORATA | (used in rule) | Number | Check if > 0 |
| GAJI NETO 80% | (used in rule) | Number | Check if > 0 |

*Full mapping in buildSlipGajiPayload_() function*

---

## 🔧 Technical Implementation

### New Functions (Welfare.js)

#### `formatCurrencyRupiah_(value)`
Converts number to "Rp 1.234.567" format

#### `applySalaryDisplayRule_(data)`
Returns formatted salary based on priority:
1. Prorata > 0 → "Rp X (Prorata)"
2. Netto 80% > 0 → "Rp X (80%)"
3. Else → "Rp X"

#### `extractJabatan_(data)`
Combines JABATAN, JABATAN FUNGSIONAL, JABATAN STRUKTURAL with " / " separator

#### `checkSlipFileExists_(fileName, folderId)`
Checks if file `SlipGaji_YYYYMM_NIP.pdf` already exists in folder
- Returns: `true` if exists, `false` if not

#### `buildPlaceholderReplacements_(data, periode)`
Creates object of all placeholder replacements:
```javascript
{
  '{{PERIODE}}': 'Januari 2026',
  '{{NAMA}}': 'John Doe',
  '{{GAJI_NETO}}': 'Rp 5.000.000 (Prorata)',
  ...
}
```

#### `generateAndSaveSlipGajiPDF(tahun, bulan)` [UPDATED]
Main function:
1. Validates user session (NIP-based)
2. Fetches slip data
3. Calculates YYYYMM from period
4. Checks if file exists → return warning if yes
5. Copies Google Docs template
6. Replaces all placeholders
7. Exports to PDF
8. Saves to `FOLDER_SLIP`
9. Shares with user (viewer role)
10. Returns success message

**Returns**:
```javascript
{
  ok: true,
  msg: 'Slip berhasil dibuat dan dikirim ke email Anda. Admin HCM juga menerima notifikasi.'
}
```

Or if already exists:
```javascript
{
  ok: false,
  msg: 'Slip gaji periode ini sudah dibuat. Untuk permintaan ulang, silakan hubungi admin HCM.',
  alreadyExists: true
}
```

---

## 🖥️ Frontend Changes (app.html)

### Button Update
- Text: **"📧 Kirim Slip (PDF)"**
- Old handler: `downloadSlipGajiPDF()` → New: `kirimSlipGajiPDF()`

### Message Handling
- **Success** (green): "✅ Slip berhasil dibuat dan dikirim ke email Anda. Admin HCM juga menerima notifikasi."
- **Already Exists** (orange): "⚠️ Slip gaji periode ini sudah dibuat. Untuk permintaan ulang, silakan hubungi admin HCM."
- **Error** (red): "❌ [Error message]"

### No File Links
- Removed: "Buka File" button with direct link
- Reason: PDF saved to Drive, not downloaded locally

---

## 📊 Updated Data Fields

### New Fields in buildSlipGajiPayload_()
```javascript
jabatanFungsional      // For detail display
jabatanStruktural      // For detail display
gajiBruto_80          // For salary display rule
bpjsJht               // BPJS TK (JHT)
bpjsJp                // BPJS TK (JP)
pph21                 // PPH 21
potAbsensi            // Potongan Absensi
```

---

## 📋 Setup Checklist

- [ ] Add `TEMPLATE_SLIP` key to HCIS_Config with Google Docs Template ID
  - Value: `1BgvBGV0hNF4G9AqG44b7Noz5XN-XenEU-5OL15IXqIg`
- [ ] Verify `FOLDER_SLIP` configured in HCIS_Config
- [ ] Verify `SLIP_GAJI_GID` configured in HCIS_Config
- [ ] Test slip generation: Kesejahteraan → Slip Gaji → Select period → "Kirim Slip"
- [ ] Verify file created with name: `SlipGaji_YYYYMM_NIP.pdf`
- [ ] Try generating same slip again → verify "already exists" warning
- [ ] Check PDF content:
  - Placeholders replaced correctly
  - Formatting matches template
  - Salary display rule applied
  - Jabatan combined properly

---

## 🧪 Testing Scenarios

### Scenario 1: Generate New Slip
1. Login as employee
2. Kesejahteraan → Slip Gaji
3. Select Tahun: 2026, Bulan: Januari
4. Click "Tampilkan" → View slip data
5. Click "Kirim Slip (PDF)"
6. Should see: ✅ "Slip berhasil dibuat..." message
7. Check FOLDER_SLIP in Drive → file exists: `SlipGaji_202601_[NIP].pdf`

### Scenario 2: Duplicate Prevention
1. Click "Kirim Slip (PDF)" again for same period
2. Should see: ⚠️ "Slip gaji periode ini sudah dibuat..." warning
3. Button should NOT trigger generation

### Scenario 3: Salary Display Rule
1. Generate slip for employee with GAJI PRORATA
2. Check PDF → should show "Rp X (Prorata)"
3. For employee with GAJI NETTO 80% only → "Rp X (80%)"
4. For regular employee → "Rp X"

### Scenario 4: New BPJS Fields
1. If sheet has columns: BPJS TK (JHT), BPJS TK (JP), PPH21, POTONGAN ABSENSI
2. Template placeholders should replace correctly
3. If columns missing → should show "Rp 0" (not error)

---

## ⚠️ Important Notes

1. **Template Required**: `TEMPLATE_SLIP` MUST be configured before feature works
2. **File Naming**: Strict format `SlipGaji_YYYYMM_NIP.pdf` for detection
3. **Duplicate Check**: Based on filename only (lightweight, no database)
4. **Security**: NIP filtering at backend prevents cross-user access
5. **New Fields**: If BPJS TK / PPH21 / Absensi columns don't exist in sheet:
   - Code handles gracefully (returns 0)
   - No error thrown
6. **Salary Rules**: Applied at template level, not changing source data

---

## 🔮 Future Enhancements

- [ ] WA notification to admin HC (hardcoded: +62 851-7520-1627)
- [ ] Email notification to employee
- [ ] Signature image upload for Ketua Yayasan
- [ ] Bulk slip generation
- [ ] PDF preview before saving
- [ ] Slip history tracking

---

## Files Modified/Created

| File | Changes | Lines |
|---|---|---|
| `Welfare.js` | New functions for template generation, salary rules, placeholder builder | +150 |
| `app.html` | Button text + handler function update | +20 |
| `HCIS_Config` | Add TEMPLATE_SLIP key | - |

---

## 📞 Support

For issues:
1. Check TEMPLATE_SLIP is configured in HCIS_Config
2. Check FOLDER_SLIP folder exists and is accessible
3. Check slip data exists in SLIP_GAJI_GID sheet
4. Review Apps Script logs: View → Execution logs

---

**End of Documentation**
