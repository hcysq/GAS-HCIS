/*************************************************
 * Cuti Management with Role-Based Approval
 *************************************************/

function getApproverByNIP(nip) {
  return getApprovalChain(nip);
}

function submitCuti(request, data) {
  try {
    const s = requireLogin_(request);
    const sh = getSheet_(CFG.SHEET_CUTI);
    const rowData = [
      Utilities.getUuid(),
      new Date(),
      s.email,
      s.nip,
      s.nama,
      data.jenis,
      data.satuan,
      data.tglMulai,
      data.tglSelesai,
      '',
      '',
      '',
      data.alasan,
      '',
      'DIAJUKAN',
      getApproverByNIP(s.nip),
      '',
      '',
      '',
      new Date()
    ];

    const maxAttempts = 3;
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const lock = LockService.getDocumentLock();
      try {
        lock.waitLock(5000);
        const nextRow = sh.getLastRow() + 1;
        sh.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
        return { ok: true };
      } catch (err) {
        Logger.log(`submitCuti attempt ${attempt} failed: ${err}`);
        if (attempt === maxAttempts) {
          throw err;
        }
        Utilities.sleep(500);
      } finally {
        lock.releaseLock();
      }
    }
  } catch (e) {
    Logger.log(`submitCuti error: ${e}`);
    return { ok: false };
  }
}
