/*************************************************
 * Session Management
 *************************************************/

function setSession_(user) {
  const token = Utilities.getUuid();
  CacheService.getUserCache().put(
    CFG.SESSION_TOKEN_KEY,
    token,
    CFG.SESSION_TTL_SECONDS
  );

  const up = PropertiesService.getUserProperties();
  up.setProperty(CFG.UP_KEYS.nip, user.nip || '');
  up.setProperty(CFG.UP_KEYS.nama, user.nama || '');
  up.setProperty(CFG.UP_KEYS.role, user.role || 'PTK');
  up.setProperty(CFG.UP_KEYS.email, user.email || '');
}

function clearSession_() {
  CacheService.getUserCache().remove(CFG.SESSION_TOKEN_KEY);
  PropertiesService.getUserProperties().deleteAllProperties();
}

function getSession_() {
  const token = CacheService.getUserCache().get(CFG.SESSION_TOKEN_KEY);
  if (!token) return null;

  const up = PropertiesService.getUserProperties();
  const nip = up.getProperty(CFG.UP_KEYS.nip);
  if (!nip) return null;

  return {
    nip,
    nama: up.getProperty(CFG.UP_KEYS.nama),
    role: up.getProperty(CFG.UP_KEYS.role),
    email: up.getProperty(CFG.UP_KEYS.email)
  };
}

function requireLogin_() {
  const s = getSession_();
  if (!s) throw new Error('SESSION_EXPIRED');
  return s;
}
