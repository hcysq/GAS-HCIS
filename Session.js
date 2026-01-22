/*************************************************
 * Session Management with Multiple Roles Support
 *************************************************/

const SESSION_KEY_PREFIX = 'SESSION:';
const SESSION_INDEX_KEY = 'SESSION_INDEX_V1';

function buildSessionKey_(token) {
  return `${SESSION_KEY_PREFIX}${token}`;
}

function getTokenFromRequest_(request) {
  if (!request) return '';
  if (typeof request === 'string') return request.trim();

  if (typeof request === 'object') {
    if (request.token) return String(request.token).trim();
    if (request.authorization) return String(request.authorization).trim();
    if (request.headers && request.headers.Authorization) {
      const authHeader = String(request.headers.Authorization || '').trim();
      if (authHeader.toLowerCase().startsWith('bearer ')) {
        return authHeader.slice(7).trim();
      }
      return authHeader;
    }
    if (request.parameter && request.parameter.token) {
      return String(request.parameter.token).trim();
    }
    if (request.parameters && request.parameters.token && request.parameters.token.length) {
      return String(request.parameters.token[0]).trim();
    }
  }

  return '';
}

function setSession_(user) {
  const token = Utilities.getUuid();
  const ttlSeconds = cfgGetNumber(
    CFG.SESSION_TTL_KEY,
    21600 // default 6 jam jika config belum diisi
  );
  const expiresAt = Date.now() + (ttlSeconds * 1000);

  const roles = (user.roles && Array.isArray(user.roles) && user.roles.length)
    ? user.roles
    : [user.role || 'PTK'];

  const sessionPayload = {
    nip: user.nip || '',
    nama: user.nama || '',
    email: user.email || '',
    userId: user.userId || '',
    role: user.role || roles.join(','),
    roles,
    expiresAt
  };

  const key = buildSessionKey_(token);
  const payload = JSON.stringify(sessionPayload);
  CacheService.getScriptCache().put(key, payload, ttlSeconds);
  PropertiesService.getScriptProperties().setProperty(key, payload);
  trackSessionKey_(key, expiresAt);
  maybeCleanupSessions_();

  return token;
}

function clearSession_(request) {
  const token = getTokenFromRequest_(request);
  if (!token) return;
  const key = buildSessionKey_(token);
  CacheService.getScriptCache().remove(key);
  PropertiesService.getScriptProperties().deleteProperty(key);
  removeSessionKey_(key);
}

function getSession_(request) {
  const sessionToken = getTokenFromRequest_(request);
  if (!sessionToken) return null;

  const key = buildSessionKey_(sessionToken);
  const cache = CacheService.getScriptCache();
  const props = PropertiesService.getScriptProperties();

  let payload = cache.get(key);
  if (!payload) {
    payload = props.getProperty(key);
    if (!payload) return null;
  }

  try {
    const data = JSON.parse(payload);
    const expiresAt = Number(data.expiresAt || 0);
    if (expiresAt && Date.now() > expiresAt) {
      clearSession_(sessionToken);
      return null;
    }

    if (expiresAt) {
      const ttlSeconds = Math.max(1, Math.floor((expiresAt - Date.now()) / 1000));
      cache.put(key, payload, ttlSeconds);
    }

    maybeCleanupSessions_();

    const { expiresAt: _expiresAt, ...session } = data;
    return session;
  } catch (e) {
    Logger.log('Error parsing session payload: ' + (e.message || e));
    clearSession_(sessionToken);
    return null;
  }
}

function requireLogin_(request) {
  const s = getSession_(request);
  if (!s) throw new Error('SESSION_EXPIRED');
  return s;
}

function removeSessionKey_(key) {
  const index = getSessionIndex_();
  if (!index[key]) return;
  delete index[key];
  saveSessionIndex_(index);
}

function trackSessionKey_(key, expiresAt) {
  const index = getSessionIndex_();
  index[key] = expiresAt;
  if (Object.keys(index).length > 500) {
    cleanupSessions_(index);
    return;
  }
  saveSessionIndex_(index);
}

function maybeCleanupSessions_() {
  const chance = Math.random();
  if (chance > 0.03) return;
  cleanupSessions_(getSessionIndex_());
}

function getSessionIndex_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(SESSION_INDEX_KEY);
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (e) {
    Logger.log('Error parsing session index: ' + (e.message || e));
    return {};
  }
}

function saveSessionIndex_(index) {
  PropertiesService.getScriptProperties().setProperty(
    SESSION_INDEX_KEY,
    JSON.stringify(index)
  );
}

function cleanupSessions_(index) {
  const now = Date.now();
  let changed = false;
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();

  Object.entries(index).forEach(([key, expiresAt]) => {
    if (!expiresAt || now > Number(expiresAt)) {
      cache.remove(key);
      props.deleteProperty(key);
      delete index[key];
      changed = true;
    }
  });

  if (changed) {
    saveSessionIndex_(index);
  }
}
