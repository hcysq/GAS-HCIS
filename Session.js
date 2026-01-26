/*************************************************
 * Session Management - MULTI-DEVICE PROTECTION
 * 1 NIP = 1 login aktif saja
 *************************************************/

const _DEVICE_SESSION_PREFIX = 'HCIS_DEVICE_SESSION_';
const _ACTIVE_SESSION_PREFIX = 'HCIS_ACTIVE_SESSION_';

function getSessionCache_() {
  return CacheService.getScriptCache();
}

function getDeviceSessionKey_(nip, deviceId) {
  return `${_DEVICE_SESSION_PREFIX}${nip}_${deviceId}`;
}

function getActiveSessionKey_(nip) {
  return `${_ACTIVE_SESSION_PREFIX}${nip}`;
}

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
  
  const cache = getSessionCache_();
  const deviceKey = getDeviceSessionKey_(nip, deviceId);
  const activeKey = getActiveSessionKey_(nip);
  const lock = LockService.getScriptLock();
  
  // Store device session (untuk read di session)
  cache.put(deviceKey, JSON.stringify(sessionData), ttlSeconds);
  
  // Store active session per NIP (untuk check single-login)
  lock.waitLock(30000);
  try {
    const existingActive = cache.get(activeKey);
    if (existingActive) {
      try {
        const parsed = JSON.parse(existingActive);
        if (parsed && parsed.deviceId && parsed.deviceId !== deviceId) {
          const oldDeviceKey = getDeviceSessionKey_(parsed.nip || nip, parsed.deviceId);
          cache.remove(oldDeviceKey);
        }
      } catch (e) {
        // ignore malformed cache
      }
    }

    cache.put(activeKey, JSON.stringify(sessionData), ttlSeconds);
  } finally {
    lock.releaseLock();
  }
  
  return {
    nip: nip,
    deviceId: deviceId,
    token: token
  };
}

function clearSession_(nip, deviceId) {
  if (nip && deviceId) {
    clearSessionByDeviceId_(nip, deviceId);
  }
}

function clearSessionByDeviceId_(nip, deviceId) {
  const cache = getSessionCache_();
  const deviceKey = getDeviceSessionKey_(nip, deviceId);
  const activeKey = getActiveSessionKey_(nip);
  const lock = LockService.getScriptLock();

  cache.remove(deviceKey);

  lock.waitLock(30000);
  try {
    const activeSession = cache.get(activeKey);
    if (activeSession) {
      try {
        const parsed = JSON.parse(activeSession);
        if (parsed && parsed.deviceId === deviceId) {
          cache.remove(activeKey);
        }
      } catch (e) {
        cache.remove(activeKey);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function getSession_(nip, deviceId) {
  if (!nip || !deviceId) return null;
  
  const cache = getSessionCache_();
  const sessionJson = cache.get(getDeviceSessionKey_(nip, deviceId));

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
  if (!nip) return null;

  const cache = getSessionCache_();
  const lock = LockService.getScriptLock();
  const activeKey = getActiveSessionKey_(nip);

  lock.waitLock(30000);
  let sessionJson = null;
  try {
    sessionJson = cache.get(activeKey);
  } finally {
    lock.releaseLock();
  }
  
  if (!sessionJson) return null;
  
  try {
    return JSON.parse(sessionJson);
  } catch (e) {
    return null;
  }
}

function requireLogin_(nip, deviceId, token) {
  if (!nip || !deviceId || !token) {
    throw new Error('SESSION_EXPIRED');
  }

  const s = getSession_(nip, deviceId);
  if (!s) throw new Error('SESSION_EXPIRED');
  if (s.nip !== nip || s.deviceId !== deviceId || s.token !== token) {
    throw new Error('SECURITY_ERROR: Session mismatch');
  }

  const activeSession = getActiveSessionForNip_(nip);
  if (!activeSession || activeSession.deviceId !== deviceId || activeSession.token !== token) {
    throw new Error('ALREADY_LOGGED_IN');
  }

  return s;
}

function validateSessionNip_(requiredNip, deviceId, token) {
  const session = requireLogin_(requiredNip, deviceId, token);
  if (session.nip !== requiredNip) {
    throw new Error(`SECURITY_ERROR: Session NIP mismatch (expected ${requiredNip}, got ${session.nip})`);
  }
  return session;
}
