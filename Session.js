/*************************************************
 * Session Management - MULTI-DEVICE PROTECTION
 * 1 NIP = 1 login aktif saja
 *************************************************/

const _DEVICE_SESSION_PREFIX = 'HCIS_DEVICE_SESSION_';
const _ACTIVE_SESSION_PREFIX = 'HCIS_ACTIVE_SESSION_';

function setSession_(user) {
  const nip = user.nip;
  const deviceId = Utilities.getUuid();
  const token = Utilities.getUuid();
  const ttlSeconds = cfgGetNumber(
    CFG.SESSION_TTL_KEY,
    21600 // default 6 jam
  );
  
  // Store session data dalam CacheService (AMAN, tidak shared antar tab)
  const sessionData = {
    nip: nip,
    nama: user.nama || '',
    role: user.role || 'PTK',
    email: user.email || '',
    userId: user.userId || '',
    deviceId: deviceId,
    token: token,
    createdAt: new Date().getTime()
  };
  
  const cache = CacheService.getUserCache();
  
  // Store device session (untuk read di session)
  cache.put(
    _DEVICE_SESSION_PREFIX + deviceId,
    JSON.stringify(sessionData),
    ttlSeconds
  );
  
  // Store active session per NIP (untuk check single-login)
  cache.put(
    _ACTIVE_SESSION_PREFIX + nip,
    JSON.stringify(sessionData),
    ttlSeconds
  );
  
  // Pass deviceId ke frontend via return value
  // Frontend akan store di sessionStorage (per-tab isolation)
  return deviceId;
}

function clearSession_(deviceId) {
  if (deviceId) {
    clearSessionByDeviceId_(deviceId);
  }
}

function clearSessionByDeviceId_(deviceId) {
  const cache = CacheService.getUserCache();
  cache.remove(_DEVICE_SESSION_PREFIX + deviceId);
}

function getSession_(deviceId) {
  if (!deviceId) return null;
  
  const cache = CacheService.getUserCache();
  const sessionJson = cache.get(_DEVICE_SESSION_PREFIX + deviceId);

  if (!sessionJson) {
    return null;
  }
  
  try {
    return JSON.parse(sessionJson);
  } catch (e) {
    return null;
  }
}

function getActiveSessionForNip_(nip) {
  const cache = CacheService.getUserCache();
  const sessionJson = cache.get(_ACTIVE_SESSION_PREFIX + nip);
  
  if (!sessionJson) return null;
  
  try {
    return JSON.parse(sessionJson);
  } catch (e) {
    return null;
  }
}

function requireLogin_(deviceId) {
  const s = getSession_(deviceId);
  if (!s) throw new Error('SESSION_EXPIRED');
  return s;
}

function validateSessionNip_(requiredNip, deviceId) {
  const session = requireLogin_(deviceId);
  if (session.nip !== requiredNip) {
    throw new Error(`SECURITY_ERROR: Session NIP mismatch (expected ${requiredNip}, got ${session.nip})`);
  }
  return session;
}
