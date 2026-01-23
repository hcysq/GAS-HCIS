# QUICK REFERENCE: EDIT PROFIL FEATURE

## 🚀 QUICK START

### Untuk User
1. Tab Profil → Klik ✏️ Edit → Input nilai → Lanjut → Ya, Simpan
2. Field sensitif (NIK, Rekening) butuh checkbox consent
3. Perubahan dicatat otomatis di sheet Histori_Mutasi

### Untuk Dev
1. File modified: `Config.js`, `Profile.js`, `app.html`, `style.html`
2. New functions: `saveProfilFieldChange()`, `logProfilMutation_()`
3. New sheet required: `Histori_Mutasi` (auto-create jika kosong)

### Untuk Admin
1. Setup sheet `Histori_Mutasi` atau set GID di config
2. (Optional) Protect sheet dengan read-only permissions
3. Monitor perubahan di sheet, create report jika perlu

---

## 🔧 KEY FUNCTIONS

### Backend (Profile.js)

```javascript
// Save field change to Users sheet + log to history
saveProfilFieldChange({
  field_key: "NIK",                // Required
  field_label: "NIK",              // Optional
  old_value: "1234567890123456",  // Required
  new_value: "1234567890654321",  // Required
  consent_checked: true            // Required for sensitive fields
})
// Returns: { ok: true/false, msg: "...", mutasi_id: "..." }

// Log mutation to Histori_Mutasi sheet
logProfilMutation_({
  target_nip, target_nama,
  field_key, field_label,
  old_value, new_value,
  changed_by_nip, changed_by_nama,
  actor_role: "pegawai",
  consent_checked: true/false,
  reason: ""
})
// Returns: { ok: true/false, msg: "...", mutasi_id: "..." }

// Get Histori_Mutasi sheet
getHistoriMutasiSheet_()
// Returns: { sheet, error }
```

### Frontend (app.html)

```javascript
// Open edit modal for a field
startEditField(fieldKey, fieldLabel, currentValue, fieldType)

// Close edit modal
cancelEditField()

// Validate and proceed to confirm modal
confirmEditField()

// Display confirm modal with old/new values
showConfirmChangeModal(newValue)

// Close confirm modal
cancelConfirmChange()

// Save field change (backend call)
saveFieldChange()

// List sensitive fields
getSensitiveFieldsList_()
// Returns: ['NIK', 'No._Rekening']

// Get input type for field
getFieldType_(fieldKey)
// Returns: 'text' | 'number' | 'date' | 'select' | 'textarea'
```

---

## 📊 DATA STRUCTURES

### Edit Field State
```javascript
window.editFieldState = {
  fieldKey: "NIK",              // Key from sheet
  fieldLabel: "NIK",            // Label for user
  oldValue: "1234...",          // Current value
  newValue: "1234...",          // New value (set by user)
  fieldType: "text",            // Input type
  isSensitive: true             // Boolean
}
```

### Histori_Mutasi Record
```javascript
{
  Mutasi_ID: "d5f4a3b2-...",    // UUID v4
  Timestamp: "2026-01-23T...",  // ISO 8601
  Target_NIP: "198701151234",
  Target_Nama: "John Doe",
  Field_Key: "NIK",
  Field_Label: "NIK",
  Old_Value: "1234567890123456",
  New_Value: "1234567890654321",
  Changed_By_NIP: "198701151234",
  Changed_By_Nama: "John Doe",
  Actor_Role: "pegawai",
  Change_Source: "profil_edit",
  Reason: "",
  Consent_Checked: "TRUE",
  Client_Info: "",
  Request_ID: ""
}
```

---

## 🎯 FIELD MAPPING

### Editable Fields
| Field Key | Type | Label |
|-----------|------|-------|
| Nama | text | Nama |
| NIP | number | NIP |
| UNIT | text | Unit |
| JABATAN | text | Jabatan |
| Status_Kepeg | select | Status Kepegawaian |
| TMT | date | TMT |
| NIK | number | **NIK** ⚠️ |
| Jenis_Kelamin | select | Jenis Kelamin |
| Status_Nikah | select | Status Pernikahan |
| No_KK | number | No. Kartu Keluarga |
| Ayah_Kandung | text | Ayah Kandung |
| Ibu_Kandung | text | Ibu Kandung |
| Gelar_Akademik_Depan | text | Gelar Akademik Depan |
| Gelar_Akademik_Belakang | text | Gelar Akademik Belakang |
| BPJS_Kes | number | BPJS Kesehatan |
| BPJS_TK | number | BPJS Ketenagakerjaan |
| Status_PTKP | select | Status PTKP |
| No._Rekening | number | **No. Rekening** ⚠️ |
| Pendidikan_Terakhir | text | Pendidikan Terakhir |
| No_HP | number | No. HP |
| WhatsApp | text | WhatsApp |
| Email | text | Email |
| Alamat | textarea | Alamat |
| Kelurahan_Desa | text | Kelurahan / Desa |
| Kecamatan | text | Kecamatan |
| Kabupaten_Kota | text | Kabupaten / Kota |
| Kode_Pos | number | Kode Pos |
| Darurat_Nama | text | Kontak Darurat - Nama |
| Darurat_Hubungan | text | Kontak Darurat - Hubungan |
| Darurat_HP | text | Kontak Darurat - HP |

⚠️ = Sensitive field (requires consent)

### Read-Only Fields (No Edit)
- Masa Kerja (auto-calculated)
- TTL (usually immutable)
- Pendidikan Formal/Non-Formal (table structure)

---

## 🛠️ COMMON TASKS

### Add Editable Field
1. Add field key to field mapping in `renderProfilLayout_()`
2. Add `item()` call with correct fieldKey
3. If sensitive: add to `getSensitiveFieldsList_()`
4. If custom type: add to `getFieldType_()` function
5. Test edit flow

### Change Sensitive Fields List
Edit `getSensitiveFieldsList_()` in app.html:
```javascript
function getSensitiveFieldsList_() {
  return ['NIK', 'No._Rekening', 'BPJS_Kes']; // Add here
}
```

### Customize Field Type
Edit `getFieldType_()` in app.html:
```javascript
function getFieldType_(fieldKey) {
  const numberFields = ['NIK', 'No_KK', 'Kode_Pos']; // Change here
  // ...
  if (numberFields.includes(fieldKey)) return 'number';
}
```

### Change Histori Sheet Name
1. Rename sheet to something else
2. Edit `getHistoriMutasiSheet_()` in Config.js
3. Or set HISTORI_MUTASI_GID in HCIS_Config

### Add Field Validation
In `confirmEditField()` before lanjut:
```javascript
if (fieldKey === 'Email' && !newValue.includes('@')) {
  alert('Email tidak valid');
  return;
}
```

---

## 🐛 DEBUGGING

### Check Modal State
```javascript
// Browser console
console.log(window.editFieldState);

// Should show current field being edited
{
  fieldKey: "NIK",
  fieldLabel: "NIK",
  oldValue: "1234567890123456",
  newValue: "1234567890654321",
  fieldType: "number",
  isSensitive: true
}
```

### Check Session
```javascript
// Browser console
console.log(state.me);

// Should show logged-in user
{
  ok: true,
  nip: "198701151234",
  nama: "John Doe",
  email: "john@example.com",
  role: "pegawai"
}
```

### Check Histori Sheet
```javascript
// Google Apps Script console
google.script.run.withSuccessHandler(res => {
  console.log(res);
}).getHistoriMutasiSheet_();
```

### Enable Backend Logging
Add to backend functions:
```javascript
Logger.log('Field: ' + field_key + ', Value: ' + new_value);
```

---

## 📱 MODAL CLASSES

### Show Modal
```javascript
const modal = document.getElementById('editFieldModal');
modal.classList.remove('hidden');  // Show
modal.classList.add('hidden');     // Hide
```

### Check Modal Visibility
```javascript
const isVisible = !modal.classList.contains('hidden');
```

---

## ⚡ PERFORMANCE TIPS

1. **Avatar Mode**: Don't load all profile data at once
2. **Lazy Load**: Only load histori if user request
3. **Cache State**: Don't re-fetch field types
4. **Debounce**: If adding real-time validation
5. **Optimize Sheet Access**: Batch queries if possible

---

## 🔐 SECURITY NOTES

### What's Protected
✅ Session validation (only logged-in users)
✅ User isolation (can only edit own data)
✅ Consent logging (for sensitive fields)
✅ Histori append-only (immutable)

### What's NOT Protected (Yet)
❌ Sheet protection (read-only for users)
❌ Approval workflow
❌ Field-level permissions
❌ Encryption at rest

### Future Security
- Add sheet protection (read-only)
- Implement approval for sensitive fields
- Add encryption for audit log
- Rate limiting for edits

---

## 📋 TESTING CHECKLIST

- [ ] Create sheet Histori_Mutasi
- [ ] Login as pegawai
- [ ] Edit non-sensitive field (Email)
- [ ] Check histori record created
- [ ] Edit sensitive field (NIK)
- [ ] Check consent checkbox appears
- [ ] Check button disabled until checked
- [ ] Edit field with special characters
- [ ] Cancel at edit modal
- [ ] Cancel at confirm modal
- [ ] Mobile test (touch screen)
- [ ] Browser test (Chrome, Firefox, Safari)
- [ ] Long input validation
- [ ] Empty input validation
- [ ] Same value validation

---

## 📞 HELP COMMANDS

### Reset Edit State
```javascript
window.editFieldState = {};
document.getElementById('editFieldModal').classList.add('hidden');
document.getElementById('confirmChangeModal').classList.add('hidden');
```

### Force Refresh Profil
```javascript
renderProfil();
```

### Clear Session
```javascript
google.script.run.withSuccessHandler(() => {
  location.reload();
}).authLogout();
```

### Export Histori
```
Manual: Select sheet Histori_Mutasi → Download as CSV/Excel
```

---

## 🎓 LEARNING RESOURCES

- **User Guide**: PANDUAN_EDIT_PROFIL.md
- **Setup Guide**: SETUP_HISTORI_MUTASI.md
- **Technical Doc**: IMPLEMENTATION_PROFIL_EDIT.md
- **Checklist**: CHECKLIST_IMPLEMENTASI_PROFIL_EDIT.md
- **Code**: Profile.js, app.html, Config.js, style.html

---

## 📞 CONTACTS

- **Questions**: hc@sabilulquran.id
- **Bugs**: Issue tracker (if available)
- **Feature Requests**: Team discussion

---

**Last Updated**: 23 January 2026  
**Version**: 1.0  
**Status**: Production Ready ✅

