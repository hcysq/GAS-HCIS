# HCIS Google Apps Script - Developer Guide

## Project Overview
Google Apps Script web application for Human Capital Information System (HCIS) at Sabilul Qur'an. Backend runs as GAS bound to a Google Spreadsheet; frontend is a SPA in [app.html](../app.html) with routing, authentication, and role-based UI.

**Architecture**: Spreadsheet-as-Database → GAS Backend (`.js` files) → Client SPA (HTML/JS with `google.script.run`)

## Core Concepts

### 1. Spreadsheet-as-Database Pattern
All data lives in Google Sheets. Each module references sheets via:
- **CFG constants** ([code.js](../code.js)): `CFG.SHEET_USERS`, `CFG.SHEET_CUTI`, etc.
- **Config sheet** (`HCIS_Config`): Dynamic configs with `cfgGet(key, default)` ([Config.js](../Config.js))
  - Example: `SLIP_GAJI_GID` stores sheet GID for welfare module
  - Cache TTL: 5 minutes (ScriptCache)

**Data Access Pattern**:
```javascript
const t = readTable_(CFG.SHEET_USERS);  // Returns {headers, rows}
const colIdx = col_(t.headers, 'NIP');  // Find column index
```

### 2. Authentication & Session
- **Single device per NIP**: New login invalidates previous session ([Session.js](../Session.js))
- **CacheService storage**: `deviceId` + `token` stored in ScriptCache (6h TTL)
- **Client-side auth**: `sessionStorage` holds `{deviceId, token, nip}` ([app.html](../app.html))
- **Auth flow**: `authLogin(nip, pin)` → hash PIN → `setSession_()` → return deviceId/token

**Server-side guard**:
```javascript
const s = requireLogin_(payload.nip, payload.deviceId, payload.token);
// Throws if session invalid
```

### 3. Multi-Role System
Users can have multiple comma-separated roles: `"PTK,ADMIN,KAPLA"` ([RoleManager.js](../RoleManager.js))
- **PTK** (base role): All users
- **KAPLA**: Unit heads (see approvals)
- **ADMIN**: System admins

**Role parsing** (Auth.js integration):
```javascript
// In authLogin after PIN validation:
const roles = parseRoles_(user.role);  // "PTK,ADMIN" → ['PTK','ADMIN']
```

**Checking roles**:
```javascript
if (hasRole('ADMIN')) { /* admin panel */ }
if (isManager()) { /* KAPLA or ADMIN */ }
```

### 4. Module Structure
Each feature is a separate `.js` file with functions exposed to client via `google.script.run`:
- [Auth.js](../Auth.js): Login/logout/session validation
- [Profile.js](../Profile.js): User masterdata & edit (16-column history tracking)
- [Welfare.js](../Welfare.js): Salary slips (PDF generation)
- [Cuti.js](../Cuti.js): Leave requests & approvals
- [Dashboard.js](../Dashboard.js): Summary stats
- [RoleManager.js](../RoleManager.js): RBAC helpers

**Naming convention**: Module functions return `{ok: boolean, data?: any, msg?: string}`

### 5. Client-Side SPA ([app.html](../app.html))
- **Single-page routing**: `goto(route)` manages `state.route` ('dash', 'profil', 'cuti', etc.)
- **State management**: Global `state = {me, route}` where `me` is session user
- **API calls**: `google.script.run.withSuccessHandler(callback).functionName(payload)`
- **Auth wrapper**: `withAuthPayload_(data)` adds deviceId/token/nip to all API calls

**Route rendering pattern**:
```javascript
function renderDashboard() {
  root.innerHTML = `...`;  // Direct DOM manipulation
}
```

## Critical Workflows

### Deployment
1. Copy all `.js` files to Google Apps Script editor (bound to spreadsheet)
2. Deploy as Web App: Execute as "User deploying", Access "Anyone, even anonymous"
3. Test URL: `https://script.google.com/macros/s/.../exec`

**Authorization**: Run [AuthBootstrap.js](../AuthBootstrap.js):`authorizeHCIS()` once for UrlFetchApp permissions

### Testing
- **Console logs**: Both server (`Logger.log`) and client (`console.log`)
- **Quick test function** ([code.js](../code.js)):
  ```javascript
  testProfilConfig()  // Validates config sheet
  ```
- **Manual test checklist**: See [TESTING_MULTIPLE_ROLES.md](../TESTING_MULTIPLE_ROLES.md)

### Adding a New Field to Profile
1. Add column to `Users` sheet
2. Update header mapping in [Profile.js](../Profile.js):`buildMasterdataPayload_()` with aliases:
   ```javascript
   const newField = getText(['New_Field', 'NEW FIELD', 'NewField']);
   ```
3. Add to `FIELD_CONFIG` in [Profile.js](../Profile.js) for editability
4. Add 16-column audit row to `Histori_Mutasi` sheet (see [SETUP_HISTORI_MUTASI.md](../SETUP_HISTORI_MUTASI.md))

### Config Management
Use [Config.js](../Config.js) for runtime settings stored in `HCIS_Config` sheet:
```javascript
cfgRequireString('SLIP_GAJI_GID')  // Throws if missing
cfgGetNumber('SESSION_TTL_KEY', 21600)  // Default 6h
cfgSet('NEW_KEY', 'value', 'Description note')
```

## Code Conventions

### Server-Side
- **Private functions**: Suffix with `_` (e.g., `setSession_()`, `requireLogin_()`)
- **Error handling**: Return `{ok:false, msg}` instead of throwing (except auth guards)
- **Caching**: Always use `CacheService.getScriptCache()` for perf
- **Date formatting**: Use `formatDateLocal_()` for Indonesia timezone (WIB)

### Client-Side
- **No jQuery/frameworks**: Vanilla JS only
- **Inline styles**: Styles in `<style>` tags, no external CSS
- **Error display**: Use `renderError(msg)` helper
- **Loading states**: Show `Memuat...` during `google.script.run` calls

### Header Mapping Pattern
Always support multiple header aliases (Indonesian vs English, different formats):
```javascript
const idxNip = findHeaderIdx_(headerMap, ['NIP', 'No Pegawai', 'Employee ID']);
```

## Security Notes

- **PIN hashing**: Use `hashPin_(pin)` with SHA-256 ([Auth.js](../Auth.js))
- **Client role checks**: UI only - ALWAYS validate roles server-side with `requireRole()`
- **Session expiry**: 6h default (configurable via `SESSION_TTL_KEY`)
- **Sensitive fields**: Require explicit user consent before editing (see [Profile.js](../Profile.js) consent modal)

## Common Pitfalls

1. **Cache invalidation**: Call `clearUsersCache_()` after updating Users sheet
2. **Header mismatch**: Use `buildHeaderMap_()` for case-insensitive header lookups
3. **Session loss**: Client must handle 401/session errors and redirect to login
4. **GID vs sheet name**: Config uses GID (right-click sheet tab → Copy GID) for robustness
5. **Date serialization**: Sheets store dates as numbers - use `formatDateLocal_()`

## Documentation Map

### Quick Start
- [START_HERE.md](../START_HERE.md): Multiple roles implementation overview
- [DOKUMENTASI_INDEX.md](../DOKUMENTASI_INDEX.md): Full documentation index by audience

### Feature Guides
- **Profile Edit**: [PANDUAN_EDIT_PROFIL.md](../PANDUAN_EDIT_PROFIL.md) (user), [IMPLEMENTATION_PROFIL_EDIT.md](../IMPLEMENTATION_PROFIL_EDIT.md) (dev)
- **Salary Slips**: [IMPLEMENTASI_KESEJAHTERAAN_TAHAP1.md](../IMPLEMENTASI_KESEJAHTERAAN_TAHAP1.md)
- **Role System**: [ROLE_SYSTEM.md](../ROLE_SYSTEM.md), [ROLE_SETUP_GUIDE.md](../ROLE_SETUP_GUIDE.md)

### Reference
- [QUICK_REFERENCE.md](../QUICK_REFERENCE.md): Multi-role API examples
- [COMPLETION_CHECKLIST.md](../COMPLETION_CHECKLIST.md): Implementation validation

## Key Files

- [code.js](../code.js): Global config & test functions
- [WebApp.js](../WebApp.js): Entry point (`doGet()`)
- [app.html](../app.html): Main SPA (2000+ lines)
- [Utils.js](../Utils.js): `readTable_()`, `getSheet_()`, `txt()` helpers
- [appsscript.json](../appsscript.json): GAS manifest (timezone: Asia/Jakarta, runtime: V8)
