# ⚡ QUICK SETUP ROLE SYSTEM

## 3 Langkah Setup

### **STEP 1: Update Sheet "Users" - Tambah Role Column**

Buka spreadsheet HCIS, sheet "Users":

**Sebelum:**
```
NIP | PIN | Nama | Email | Aktif
```

**Sesudah:**
```
NIP | PIN | Nama | Email | Role | USER_ID | Aktif
```

Posisi column "Role" bisa di mana saja, tapi penting ada di header.

---

### **STEP 2: Set Role untuk Setiap User**

**Untuk data existing**, tambahkan role di kolom baru:

```
NIP      | PIN   | Nama           | Email          | Role  | USER_ID | Aktif
---------|-------|----------------|----------------|-------|---------|-------
12345678 | pin1  | Bambang S.     | bambang@co.id  | PTK   | USR001  | TRUE
23456789 | pin2  | Siti Nur       | siti@co.id     | KAPLA | USR002  | TRUE
34567890 | pin3  | Ahmad F.       | ahmad@co.id    | ADMIN | USR003  | TRUE
```

**Nilai Role yang valid:**
- `PTK` = Regular employee (default)
- `KAPLA` = Kepala unit/departemen
- `ADMIN` = System administrator

---

### **STEP 3: Setup AtasanMap (Hierarki Atasan)**

Buka/buat sheet "AtasanMap" dengan struktur:

```
NIP | ApproverNIP | Aktif
```

**Contoh hierarki:**
```
NIP      | ApproverNIP | Aktif
---------|-------------|-------
12345678 | 23456789    | TRUE
87654321 | 23456789    | TRUE
23456789 | 34567890    | TRUE
99999999 | 34567890    | TRUE
```

**Arti:**
- Pegawai 12345678 & 87654321 → atasan Siti (23456789) KAPLA
- Siti (23456789) KAPLA → atasan Ahmad (34567890) ADMIN
- Pengguna lain → atasan Ahmad ADMIN

---

## Testing Roles

### **Test 1: Login sebagai PTK**
1. Login NIP=12345678
2. Dashboard show: "Role: **Pegawai**"
3. Menu: Dashboard, Profil, Settings (tidak ada Approvals/Admin)

### **Test 2: Login sebagai KAPLA**
1. Login NIP=23456789
2. Dashboard show: "Role: **Kepala Unit**"
3. Menu: Dashboard, Profil, **Approvals** ← baru!, Settings
4. Klik Approvals → lihat cuti dari subordinates (12345678, 87654321)

### **Test 3: Login sebagai ADMIN**
1. Login NIP=34567890
2. Dashboard show: "Role: **Administrator**"
3. Menu: Dashboard, Profil, **Admin**, **Approvals**, Settings
4. Klik Admin → manage users, settings, reports

---

## Cek di Apps Script Console

```javascript
// Cek users yang sudah load
clearUsersCache_();  // Clear cache dulu
var map = loadUsersMap_();
Logger.log(JSON.stringify(map, null, 2));

// Cek role structure
Logger.log(ROLES);      // { PTK: 'PTK', KAPLA: 'KAPLA', ADMIN: 'ADMIN' }
```

---

## File Baru yang Ditambah

✅ **RoleManager.js** - Core role management functions
- `hasRole(role)` - Check user role
- `isAdmin()` / `isManager()` - Quick checks
- `requireRole(role)` - Throw error jika tidak authorized
- `getApprovalsPending()` - Get pending approvals for current user
- `approveCuti(cutiId, approved, reason)` - Approve/reject cuti

✅ **ROLE_SYSTEM.md** - Complete role system documentation

---

## Yang Belum/Perlu Dikerjakan

- [ ] UI untuk Approvals page (show pending cuti, approve buttons)
- [ ] Admin panel untuk manage users
- [ ] Role selector di admin panel
- [ ] Notification ke approver ada pending cuti
- [ ] Approval history log
- [ ] Reports based on role

---

**Mulai Test Sekarang!** ✨
