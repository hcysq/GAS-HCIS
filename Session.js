/*************************************************
 * Session Management with Multiple Roles Support
 *************************************************/

// Storage keys for roles
const UP_KEYS_ROLES = 'HCIS_ROLES';

function setSession_(user) {
  const token = Utilities.getUuid();
  const ttlSeconds = cfgGetNumber(
    CFG.SESSION_TTL_KEY,
    21600 // default 6 jam jika config belum diisi
  );
  CacheService.getUserCache().put(
    CFG.SESSION_TOKEN_KEY,
    token,
    ttlSeconds
  );

  const up = PropertiesService.getUserProperties();
  up.setProperty(CFG.UP_KEYS.nip, user.nip || '');
  up.setProperty(CFG.UP_KEYS.nama, user.nama || '');
  up.setProperty(CFG.UP_KEYS.email, user.email || '');
  up.setProperty(CFG.UP_KEYS.userId, user.userId || '');
  
  // Handle multiple roles
  if (user.roles && Array.isArray(user.roles)) {
    // Store roles as JSON array
    up.setProperty(UP_KEYS_ROLES, JSON.stringify(user.roles));
    up.setProperty(CFG.UP_KEYS.role, user.role || user.roles.join(','));
  } else {
    // Fallback to single role
    up.setProperty(CFG.UP_KEYS.role, user.role || 'PTK');
    up.setProperty(UP_KEYS_ROLES, JSON.stringify([user.role || 'PTK']));
  }
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

  // Get roles array
  let roles;
  try {
    const rolesJson = up.getProperty(UP_KEYS_ROLES);
    roles = rolesJson ? JSON.parse(rolesJson) : ['PTK'];
  } catch (e) {
    roles = [up.getProperty(CFG.UP_KEYS.role) || 'PTK'];
  }

  return {
    nip,
    nama: up.getProperty(CFG.UP_KEYS.nama),
    role: up.getProperty(CFG.UP_KEYS.role),
    roles: roles,  // Multiple roles as array
    email: up.getProperty(CFG.UP_KEYS.email),
    userId: up.getProperty(CFG.UP_KEYS.userId)
  };
}

function requireLogin_() {
  const s = getSession_();
  if (!s) throw new Error('SESSION_EXPIRED');
  return s;
}
