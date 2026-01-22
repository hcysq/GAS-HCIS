# 👤 ROLE-BASED ACCESS CONTROL (RBAC) SYSTEM

## Tingkatan User yang Ada

### 1. **PTK** (Pegawai Tetap/Kontrak)
- **Default role** untuk semua user baru
- **Permissions:**
  - Login ke sistem
  - Lihat profil sendiri
  - Submit pengajuan cuti
  - Lihat saldo cuti
  - Ubah password
  - View dashboard
- **Tidak bisa:**
  - Approve cuti
  - Lihat data user lain
  - Manage system settings

### 2. **KAPLA** (Kepala Unit/Departemen)
- **Permissions:** Semua PTK + 
  - Lihat daftar subordinates
  - Approve/reject pengajuan cuti dari subordinates
  - Lihat approval dashboard
  - View attendance subordinates
  - Generate laporan unit
- **Tidak bisa:**
  - Manage users
  - Manage system settings
  - Admin functions

### 3. **ADMIN** (Administrator)
- **Full permissions** - semua fitur
  - Login sebagai user lain (impersonate)
  - Manage semua users (create, edit, delete, deactivate)
  - Approve/reject semua cuti
  - View semua reports
  - Manage system config
  - Backup/export data
  - Manage holidays & kalender

---

## Setup Role di Sheet Users

### **Column di Sheet "Users":**
```
NIP | PIN | Nama | Email | Role | USER_ID | Aktif
```

### **Nilai Role yang Valid:**
- `PTK` (default jika kosong)
- `KAPLA`
- `ADMIN`

### **Contoh Data:**
```
NIP      | PIN       | Nama            | Email              | Role  | USER_ID | Aktif
---------|-----------|-----------------|-------------------|-------|---------|-------
12345678 | pass123   | Bambang Suryono | bambang@example.id | PTK   | USR001  | TRUE
23456789 | pass456   | Siti Nurhaliza  | siti@example.id    | KAPLA | USR002  | TRUE
34567890 | pass789   | Ahmad Faizal    | ahmad@example.id   | ADMIN | USR003  | TRUE
```

---

## Map Hierarki Atasan (Sheet: AtasanMap)

**Column:**
```
NIP | ApproverNIP | Aktif
```

**Maksud:**
- NIP = Pegawai
- ApproverNIP = Atasan/Approver mereka

**Contoh:**
```
NIP      | ApproverNIP | Aktif
---------|-------------|-------
12345678 | 23456789    | TRUE
87654321 | 23456789    | TRUE
23456789 | 34567890    | TRUE
```

Artinya:
- User 12345678 → atasan 23456789 (KAPLA Siti)
- User 87654321 → atasan 23456789 (KAPLA Siti)
- User 23456789 (KAPLA) → atasan 34567890 (ADMIN)

---

## Function Reference (RoleManager.js)

### **Check Role**

```javascript
// Check role user saat ini
getUserRole()  // returns: 'PTK', 'KAPLA', 'ADMIN', dll

// Check apakah user punya role tertentu
hasRole('ADMIN')                    // true/false
hasRole(['KAPLA', 'ADMIN'])         // true/false jika salah satu match

// Check spesifik
isAdmin()        // true jika ADMIN
isManager()      // true jika KAPLA atau ADMIN
```

### **Require Role**

```javascript
// Throw error jika tidak punya role
requireRole('ADMIN')                // throw error jika bukan ADMIN
requireRole(['KAPLA', 'ADMIN'])     // throw error jika bukan keduanya
```

### **Get Hierarchy**

```javascript
// Get approver seseorang
getApprovalChain('12345678')  // returns: '23456789' (NIP atasan)

// Get semua subordinates dari manager
getSubordinates('23456789')   // returns: ['12345678', '87654321']
```

### **Approval APIs**

```javascript
// Get pending approvals untuk current user (hanya KAPLA/ADMIN)
getApprovalsPending()
// returns: {
//   ok: true,
//   data: [
//     { id, nip, nama, jenis, mulai, selesai, alasan, status },
//     ...
//   ]
// }

// Approve/reject cuti (hanya KAPLA/ADMIN yang request ke mereka)
approveCuti('CUTI_ID', true, 'Disetujui')   // approved
approveCuti('CUTI_ID', false, 'Alasan tolak')  // rejected
// returns: { ok: true/false, msg: '...' }
```

---

## Implementation Examples

### **Contoh 1: Cek apakah user bisa akses admin panel**

```javascript
function renderAdminPanel() {
  if (!isAdmin()) {
    root.innerHTML = '<div class="error">Akses ditolak. Anda bukan admin.</div>';
    return;
  }
  // Show admin panel
}
```

### **Contoh 2: Backend function yang hanya KAPLA/ADMIN bisa call**

```javascript
function approveCutiForSubordinate(cutiId) {
  try {
    requireRole([ROLES.KAPLA, ROLES.ADMIN]);  // throw error jika PTK
    
    const s = getSession_();
    // Process approval
    approveCuti(cutiId, true);
    
    return { ok: true };
  } catch (err) {
    return { ok: false, msg: err.message };
  }
}
```

### **Contoh 3: Limit data berdasarkan role**

```javascript
function getApprovalsPendingForUser() {
  const s = requireLogin_();
  
  if (isAdmin()) {
    // Admin lihat semua pending approvals
    return getAllPendingApprovals();
  } else if (isManager()) {
    // KAPLA hanya lihat subordinates mereka
    return getApprovalsPending();  // dari RoleManager
  } else {
    // PTK tidak ada approvals
    return [];
  }
}
```

---

## UI Changes untuk Role-Based UI

### **Navbar/Menu** (di app.html)
```javascript
<nav class="bottombar" id="nav">
  <button class="tab" onclick="go('dash')">Dashboard</button>
  <button class="tab" onclick="go('profil')">Profil</button>
  ${isManager() ? '<button class="tab" onclick="go(\'approvals\')">Approvals</button>' : ''}
  ${isAdmin() ? '<button class="tab" onclick="go(\'admin\')">Admin</button>' : ''}
  <button class="tab" onclick="go('settings')">Setelan</button>
</nav>
```

---

## Future: Permissions Matrix

| Feature | PTK | KAPLA | ADMIN |
|---------|-----|-------|-------|
| Login | ✅ | ✅ | ✅ |
| View Own Profile | ✅ | ✅ | ✅ |
| Submit Cuti | ✅ | ✅ | ✅ |
| Approve Cuti | ❌ | ✅ (subordinates) | ✅ (all) |
| View Subordinates | ❌ | ✅ | ✅ |
| View All Users | ❌ | ❌ | ✅ |
| Manage Users | ❌ | ❌ | ✅ |
| Manage Config | ❌ | ❌ | ✅ |
| View Reports | ❌ | ✅ (unit) | ✅ (all) |
| Export Data | ❌ | ❌ | ✅ |

---

## Setup Checklist

- [ ] Sheet "Users" punya kolom "Role"
- [ ] Data users punya role: PTK, KAPLA, atau ADMIN
- [ ] Sheet "AtasanMap" filled dengan hierarki
- [ ] Coba login sebagai KAPLA → lihat Approvals di dashboard
- [ ] Coba login sebagai ADMIN → lihat Admin menu

---

**Created:** Jan 22, 2026  
**Status:** Ready for implementation  
**Next:** Integrate into app.html UI + approval dashboard
