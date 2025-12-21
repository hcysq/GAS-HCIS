/*************************************************
 * Config.gs - Single Source of Truth for HCIS
 * Canonical config sheet: "HCIS_Config" (Key, Value, Note)
 * Sheet GID: 1743564124
 *
 * - ensureHCISConfigSheet_(): pastikan sheet & header
 * - cfgGet(key, default): ambil config (pakai cache)
 * - cfgSet(key, value, note): set config
 * - validateHCISConfig(): cek key penting
 *************************************************/

const CONFIG_SHEET_CANONICAL = 'HCIS_Config';
const CONFIG_SHEET_GID = 1743564124; // GID tab HCIS_Config (untuk pesan error)

// Cache key
const _CFG_CACHE_KEY = 'HCIS_CFG_MAP_V1';
const _CFG_CACHE_TTL = 300; // 5 menit

function ensureHCISConfigSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CONFIG_SHEET_CANONICAL);
  if (!sh) sh = ss.insertSheet(CONFIG_SHEET_CANONICAL);

  // Pastikan header A1:C1 = Key | Value | Note
  const header = sh.getRange(1, 1, 1, 3).getValues()[0];
  const expected = ['Key', 'Value', 'Note'];

  const headerOk =
    String(header[0] || '').trim() === expected[0] &&
    String(header[1] || '').trim() === expected[1] &&
    String(header[2] || '').trim() === expected[2];

  if (!headerOk) {
    sh.getRange(1, 1, 1, 3).setValues([expected]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Ambil config dari HCIS_Config (Key/Value)
 * - pakai ScriptCache supaya cepat
 */
function cfgGet(key, defaultValue) {
  key = String(key || '').trim();
  if (!key) return defaultValue;

  const cache = CacheService.getScriptCache();
  const cached = cache.get(_CFG_CACHE_KEY);
  if (cached) {
    try {
      const map = JSON.parse(cached);
      if (Object.prototype.hasOwnProperty.call(map, key)) return map[key];
      return defaultValue;
    } catch (e) {
      // lanjut load ulang
    }
  }

  const map = _loadCfgMap_();
  cache.put(_CFG_CACHE_KEY, JSON.stringify(map), _CFG_CACHE_TTL);

  if (Object.prototype.hasOwnProperty.call(map, key)) return map[key];
  return defaultValue;
}

function cfgGetNumber(key, defaultValue) {
  const v = cfgGet(key, defaultValue);
  const n = Number(v);
  return isNaN(n) ? defaultValue : n;
}

function cfgGetString(key, defaultValue) {
  const v = cfgGet(key, defaultValue);
  return String(v ?? '').trim();
}

function cfgRequireString(key) {
  const v = cfgGetString(key, '');
  if (!v) {
    throw new Error(`${key} belum diisi di ${CONFIG_SHEET_CANONICAL} (GID ${CONFIG_SHEET_GID})`);
  }
  return v;
}

function cfgSet(key, value, note) {
  key = String(key || '').trim();
  if (!key) throw new Error('cfgSet: key kosong');

  const sh = ensureHCISConfigSheet_();
  const lastRow = sh.getLastRow();

  const rows = lastRow >= 2
    ? sh.getRange(2, 1, lastRow - 1, 3).getValues()
    : [];

  let foundRowIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === key) {
      foundRowIndex = i;
      break;
    }
  }

  if (foundRowIndex === -1) {
    sh.appendRow([key, value, note || '']);
  } else {
    const rowNumber = foundRowIndex + 2;
    sh.getRange(rowNumber, 2).setValue(value);
    if (note !== undefined) sh.getRange(rowNumber, 3).setValue(note);
  }

  // clear cache
  CacheService.getScriptCache().remove(_CFG_CACHE_KEY);
  return true;
}

function cfgClearCache() {
  CacheService.getScriptCache().remove(_CFG_CACHE_KEY);
}

/**
 * Loader internal
 */
function _loadCfgMap_() {
  const sh = ensureHCISConfigSheet_();
  const lastRow = sh.getLastRow();
  const map = {};

  if (lastRow < 2) return map;

  const values = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  values.forEach(r => {
    const k = String(r[0] || '').trim();
    if (!k) return;
    map[k] = r[1]; // Value (biarkan tipe aslinya)
  });

  return map;
}

/**
 * Helper untuk buka Spreadsheet lain berdasarkan konfigurasi ID di HCIS_Config
 */
function getSpreadsheetFromConfig_(key, featureName) {
  const ssId = cfgRequireString(key);
  try {
    return SpreadsheetApp.openById(ssId);
  } catch (e) {
    const label = featureName || key;
    const errMsg = e && e.message ? e.message : e;
    throw new Error(`Gagal membuka spreadsheet ${label} (key ${key} di ${CONFIG_SHEET_CANONICAL}): ${errMsg}`);
  }
}

function getAbsensiSpreadsheet_() {
  return getSpreadsheetFromConfig_('ABSENSI_SS_ID', 'Absensi');
}

function getWelfareSpreadsheet_() {
  return getSpreadsheetFromConfig_('WELFARE_SS_ID', 'Kesejahteraan');
}

function getProjectSpreadsheet_() {
  return getSpreadsheetFromConfig_('PROJECT_SS_ID', 'Progres Proyek');
}

/**
 * Validasi key penting (silakan tambah)
 */
function validateHCISConfig() {
  ensureHCISConfigSheet_();
  const required = [
    'SESSION_TTL_SECONDS',
    'STARSENDER_URL',
    'STARSENDER_APIKEY',
    'STARSENDER_MODE',
    'ABSENSI_SS_ID',
    'WELFARE_SS_ID',
    'PROJECT_SS_ID'
  ];

  const missing = [];
  required.forEach(k => {
    const v = cfgGet(k, '');
    if (v === '' || v === null || v === undefined) missing.push(k);
  });

  return {
    ok: missing.length === 0,
    missing: missing
  };
}

/**
 * Helper cepat untuk set beberapa key sekaligus
 */
function cfgSetMany(obj) {
  Object.keys(obj || {}).forEach(k => cfgSet(k, obj[k], 'bulk set'));
  return true;
}
