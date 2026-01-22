/*************************************************
 * Session Management with Multiple Roles Support
 *************************************************/

const SESSION_KEY_PREFIX = 'SESSION:';

function buildSessionKey_(token) {
  return `${SESSION_KEY_PREFIX}${token}`;
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

function clearSession_(token) {
  if (!token) return;
  const key = buildSessionKey_(token);
  CacheService.getScriptCache().remove(key);
  PropertiesService.getScriptProperties().deleteProperty(key);
}

function getSession_(token) {
  const sessionToken = String(token || '').trim();
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

function requireLogin_(token) {
  const s = getSession_(token);
  if (!s) throw new Error('SESSION_EXPIRED');
  return s;
}
