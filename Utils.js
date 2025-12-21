/*************************************************
 * Utilities
 *************************************************/

function getSheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error(`Sheet "${name}" tidak ditemukan`);
  return sh;
}

function readTable_(sheetName) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    return { headers: [], rows: [] };
  }
  return {
    headers: values[0].map(h => String(h).trim()),
    rows: values.slice(1)
  };
}

function col_(headers, name) {
  return headers.indexOf(name);
}

function isTrue_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

function txt(v) {
  return String(v ?? '').trim();
}
