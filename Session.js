/*************************************************
 * Session Management with Multiple Roles Support
 *************************************************/

const SESSION_KEY_PREFIX = 'SESSION:';

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

  return token;
}

function clearSession_(request) {
  const token = getTokenFromRequest_(request);
  if (!token) return;
  const key = buildSessionKey_(token);
  CacheService.getScriptCache().remove(key);
  PropertiesService.getScriptProperties().deleteProperty(key);
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
