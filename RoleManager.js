/*************************************************
 * Role Manager - Role-Based Access Control
 * 
 * Roles:
 * - PTK (Pegawai Tetap/Kontrak) - Regular user
 * - KAPLA (Kepala Unit) - Unit/department head
 * - ADMIN - System admin
 *************************************************/

// Role Definitions
const ROLES = {
  PTK: 'PTK',           // Pegawai Tetap/Kontrak
  KAPLA: 'KAPLA',       // Kepala Unit/Departemen
  ADMIN: 'ADMIN'        // System Administrator
};

const ROLE_LABELS = {
  PTK: 'Pegawai',
  KAPLA: 'Kepala Unit',
  ADMIN: 'Administrator'
};

/**
 * Get user role dari session
 * @returns {string} Role user (PTK, KAPLA, ADMIN, dll)
 */
function getUserRole() {
  const s = getSession_();
  return s ? (s.role || ROLES.PTK) : null;
}

/**
 * Check apakah user punya role tertentu
 * @param {string|string[]} requiredRole - Role atau array of roles
 * @returns {boolean}
 */
function hasRole(requiredRole) {
  const userRole = getUserRole();
  if (!userRole) return false;
  
  if (Array.isArray(requiredRole)) {
    return requiredRole.includes(userRole);
  }
  return userRole === requiredRole;
}

/**
 * Check apakah user adalah ADMIN
 * @returns {boolean}
 */
function isAdmin() {
  return hasRole(ROLES.ADMIN);
}

/**
 * Check apakah user adalah KAPLA atau ADMIN
 * @returns {boolean}
 */
function isManager() {
  return hasRole([ROLES.KAPLA, ROLES.ADMIN]);
}

/**
 * Require specific role - throw error jika tidak punya
 * @param {string|string[]} requiredRole - Role atau array of roles
 * @throws {Error}
 */
function requireRole(requiredRole) {
  const userRole = getUserRole();
  if (!userRole) throw new Error('SESSION_EXPIRED');
  
  const roles = Array.isArray(requiredRole) ? requiredRole : [requiredRole];
  if (!roles.includes(userRole)) {
    throw new Error('PERMISSION_DENIED');
  }
}

/**
 * Get approval chain untuk cuti/request
 * Jika user adalah KAPLA, atasan mereka adalah ADMIN/leadership
 * Jika user adalah PTK, atasan mereka adalah KAPLA
 */
function getApprovalChain(nip) {
  const t = readTable_(CFG.SHEET_ATASAN);
  const h = t.headers;
  const r = t.rows;

  const cNIP = col_(h, 'NIP');
  const cApp = col_(h, 'ApproverNIP');
  const cAktif = col_(h, 'Aktif');

  for (const row of r) {
    if (txt(row[cNIP]) === nip && isTrue_(row[cAktif])) {
      return txt(row[cApp]);
    }
  }
  return '';
}

/**
 * Get subordinates dari KAPLA/manager
 * @param {string} managerNip - NIP manager
 * @returns {string[]} Array of subordinate NIPs
 */
function getSubordinates(managerNip) {
  const t = readTable_(CFG.SHEET_ATASAN);
  const h = t.headers;
  const r = t.rows;

  const cNIP = col_(h, 'NIP');
  const cApp = col_(h, 'ApproverNIP');
  const cAktif = col_(h, 'Aktif');

  const subordinates = [];
  for (const row of r) {
    const appNip = txt(row[cApp]);
    if (appNip === managerNip && isTrue_(row[cAktif])) {
      subordinates.push(txt(row[cNIP]));
    }
  }
  return subordinates;
}

/**
 * API: Get list of approvals yang pending untuk user
 * - Jika KAPLA/ADMIN: lihat semua cuti dari subordinates yang pending
 * - Jika PTK: tidak ada (regular users tidak approve)
 */
function getApprovalsPending() {
  try {
    const s = requireLogin_();
    
    if (!isManager()) {
      return { ok: true, data: [] }; // PTK tidak punya approvals
    }

    const cutiSheet = getSheet_(CFG.SHEET_CUTI);
    const headers = cutiSheet.getRange(1, 1, 1, cutiSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    
    const cApprover = col_(headers, 'ApproverNIP');
    const cStatus = col_(headers, 'Status');
    const cNip = col_(headers, 'NIP');
    const cNama = col_(headers, 'Nama');
    const cJenis = col_(headers, 'Jenis');
    const cMulai = col_(headers, 'TglMulai');
    const cSelesai = col_(headers, 'TglSelesai');
    const cAlasan = col_(headers, 'Alasan');
    const cId = col_(headers, 'ID');

    if (cApprover < 0 || cStatus < 0) {
      return { ok: false, msg: 'Kolom ApproverNIP atau Status tidak ditemukan' };
    }

    const rows = cutiSheet.getRange(2, 1, cutiSheet.getLastRow() - 1, cutiSheet.getLastColumn()).getValues();
    const approvals = [];

    for (const row of rows) {
      const approverNip = txt(row[cApprover]);
      const status = txt(row[cStatus]);
      
      if (approverNip === s.nip && status === 'DIAJUKAN') {
        approvals.push({
          id: cId >= 0 ? txt(row[cId]) : '',
          nip: cNip >= 0 ? txt(row[cNip]) : '',
          nama: cNama >= 0 ? txt(row[cNama]) : '',
          jenis: cJenis >= 0 ? txt(row[cJenis]) : '',
          mulai: cMulai >= 0 ? row[cMulai] : '',
          selesai: cSelesai >= 0 ? row[cSelesai] : '',
          alasan: cAlasan >= 0 ? txt(row[cAlasan]) : '',
          status: status
        });
      }
    }

    return { ok: true, data: approvals };
  } catch (e) {
    Logger.log('getApprovalsPending error: ' + (e.message || e));
    return { ok: false, msg: e.message || e };
  }
}

/**
 * API: Approve/Reject cuti request
 * Only KAPLA/ADMIN can do this for requests directed to them
 */
function approveCuti(cutiId, approved, reason = '') {
  try {
    const s = requireLogin_();
    requireRole([ROLES.KAPLA, ROLES.ADMIN]);

    const cutiSheet = getSheet_(CFG.SHEET_CUTI);
    const headers = cutiSheet.getRange(1, 1, 1, cutiSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    
    const cId = col_(headers, 'ID');
    const cStatus = col_(headers, 'Status');
    const cApproveDate = col_(headers, 'ApproveDate');
    const cApproveBy = col_(headers, 'ApproveBy');
    const cApproveNote = col_(headers, 'ApproveNote');
    const cApprover = col_(headers, 'ApproverNIP');

    if (cId < 0 || cStatus < 0) {
      throw new Error('Kolom ID atau Status tidak ditemukan');
    }

    const rows = cutiSheet.getRange(2, 1, cutiSheet.getLastRow() - 1, cutiSheet.getLastColumn()).getValues();
    
    for (let i = 0; i < rows.length; i++) {
      if (txt(rows[i][cId]) === cutiId) {
        const approverNip = txt(rows[i][cApprover]);
        if (approverNip !== s.nip) {
          throw new Error('Anda bukan approver untuk cuti ini');
        }

        const lock = LockService.getDocumentLock();
        lock.waitLock(5000);
        try {
          const rowNum = i + 2; // row 2 = index 0
          const status = approved ? 'DISETUJUI' : 'DITOLAK';
          
          cutiSheet.getRange(rowNum, cStatus + 1).setValue(status);
          if (cApproveDate >= 0) cutiSheet.getRange(rowNum, cApproveDate + 1).setValue(new Date());
          if (cApproveBy >= 0) cutiSheet.getRange(rowNum, cApproveBy + 1).setValue(s.nip);
          if (cApproveNote >= 0) cutiSheet.getRange(rowNum, cApproveNote + 1).setValue(reason);
          
          Logger.log(`Cuti ${cutiId} ${status} by ${s.nip}`);
          return { ok: true, msg: `Cuti ${status.toLowerCase()}` };
        } finally {
          lock.releaseLock();
        }
      }
    }

    return { ok: false, msg: 'Cuti tidak ditemukan' };
  } catch (err) {
    Logger.log('approveCuti error: ' + (err.message || err));
    return { ok: false, msg: err.message || err };
  }
}
