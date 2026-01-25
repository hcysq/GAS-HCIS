# ⚡ Quick Reference: Slip Gaji v2.0

## 🎯 Immediate Action Required

**Before using feature:**

1. **Add to HCIS_Config spreadsheet:**
   ```
   Key: TEMPLATE_SLIP
   Value: 1BgvBGV0hNF4G9AqG44b7Noz5XN-XenEU-5OL15IXqIg
   ```

2. **Verify these exist:**
   ```
   FOLDER_SLIP (Drive folder for saving PDFs)
   SLIP_GAJI_GID (Slip Gaji data sheet)
   ```

---

## 📄 Google Docs Template Placeholders

Copy-paste these into your Google Docs template:

```
{{PERIODE}}
{{NAMA}}
{{NIP}}
{{UNIT}}
{{JABATAN}}
{{TOTAL_BRUTO}}
{{TOTAL_POTONGAN}}
{{GAJI_NETO}}
{{GAJI_POKOK}}
{{TUNJ_KINERJA}}
{{TUNJ_ISTRI}}
{{TUNJ_ANAK}}
{{TUNJ_FUNGSIONAL}}
{{TUNJ_JABATAN}}
{{TUNJ_KUALIFIKASI}}
{{LEMBUR}}
{{RAPEL_GAJI}}
{{TUNJ_BPJS}}
{{POT_KASBON}}
{{BPJS}}
{{PEND_ANAK}}
{{KURANG_JAM}}
{{BPJS_JHT}}
{{BPJS_JP}}
{{PPH21}}
{{POT_ABSENSI}}
```

---

## 🎬 How It Works

**User clicks "📧 Kirim Slip (PDF)":**

1. ✅ Generate PDF from Google Docs template
2. ✅ Name: `SlipGaji_202601_[NIP].pdf` (consistent)
3. ✅ Check if already exists → Warn if yes
4. ✅ Save to FOLDER_SLIP in Drive
5. ✅ Share with employee
6. ✅ Show success message

---

## 📊 Salary Display Rule

System automatically shows:

- **If Prorata > 0:** "Rp 5.000.000 (Prorata)"
- **Else if 80% Netto > 0:** "Rp 4.000.000 (80%)"
- **Else:** "Rp 3.500.000"

Applied to: PDF + UI display

---

## 🛑 Duplicate Prevention

**File exists?** → Show: ⚠️ "Slip gaji periode ini sudah dibuat..."  
**File not found?** → Generate slip normally

Check FOLDER_SLIP for: `SlipGaji_YYYYMM_NIP.pdf`

---

## 🔍 File Naming

Old format:
```
Slip Gaji Jan 2026 202009199411191071 M. Imadduddin.pdf
```

New format:
```
SlipGaji_202601_202009199411191071.pdf
```

Benefits: Consistent, parseable, duplicate-detectable

---

## 📋 New Sheet Columns (Optional)

Add these to Slip Gaji sheet if not present:

```
BPJS TK (JHT)        → {{BPJS_JHT}}
BPJS TK (JP)         → {{BPJS_JP}}
PPH21                → {{PPH21}}
POTONGAN ABSENSI     → {{POT_ABSENSI}}
```

If missing: automatically shows "Rp 0"

---

## ✅ Testing

1. Add TEMPLATE_SLIP config
2. Login as employee
3. Kesejahteraan → Slip Gaji
4. Select period
5. Click "Tampilkan"
6. Click "📧 Kirim Slip"
7. Check FOLDER_SLIP for PDF
8. Try same slip again → see warning

---

## 🚨 Troubleshooting

| Issue | Solution |
|-------|----------|
| TEMPLATE_SLIP error | Add config to HCIS_Config |
| File not created | Check FOLDER_SLIP folder ID |
| Placeholders not replaced | Check Google Docs template text |
| "Already made" warning | Expected! File exists, use different period |

---

## 📝 Changed vs Unchanged

**✅ CHANGED:**
- PDF generation method (HTML → Google Docs template)
- File naming convention
- Button text & UX
- Salary display logic

**❌ NOT CHANGED:**
- Data filtering (Tahun/Bulan/NIP)
- Authentication
- Dashboard, Profile, Settings
- Other welfare modules
- Sheet structure

---

**Last Updated**: January 25, 2026
