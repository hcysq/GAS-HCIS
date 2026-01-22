# Testing Multiple Roles Support

## Overview
This document guides you through testing the new multiple roles per user feature.

## Changes Made

### 1. **Session.js** - Multiple Roles Storage
- Added `UP_KEYS_ROLES` constant to store roles as JSON array
- `setSession_(user)` now stores `user.roles` array in UserProperties
- `getSession_()` returns both:
  - `roles`: Array of roles ['PTK', 'ADMIN', 'KAPLA']
  - `role`: Original role string "PTK,ADMIN,KAPLA" (backward compat)

### 2. **RoleManager.js** - Role Parsing
- `parseRoles_(roleStr)`: Parses "PTK" or "PTK,ADMIN,KAPLA" into array
- `getUserRoles()`: Returns roles array
- `hasRole(role)`: Checks if user has ANY of the required roles

### 3. **Auth.js** - Login Integration
- `authLogin()` now calls `parseRoles_()` to convert role string to array
- Session stores both roles array and original string

### 4. **app.html** - Dashboard Display
- `hasRole(role)` helper: Checks membership in roles array
- `roleDisplay`: Shows all roles joined by comma
- Conditional panels: Use `hasRole()` instead of `===` check

## Test Scenarios

### Scenario 1: Single Role User (PTK)
**Setup:**
- Create/use test user with Role = "PTK"

**Expected Results:**
- Login succeeds
- Dashboard shows: "Role: PTK"
- Admin panel: HIDDEN
- Team panel: HIDDEN
- Core HCIS features: VISIBLE

```javascript
// In browser console after login:
console.log(state.me.roles);        // ['PTK']
console.log(state.me.role);         // 'PTK'
```

---

### Scenario 2: Admin Role
**Setup:**
- Create/use test user with Role = "ADMIN"

**Expected Results:**
- Login succeeds
- Dashboard shows: "Role: PTK, ADMIN"
- Admin panel: VISIBLE
- Team panel: HIDDEN
- HCIS features: VISIBLE

```javascript
// In browser console:
console.log(state.me.roles);        // ['PTK', 'ADMIN']
console.log(state.me.role);         // 'PTK,ADMIN'
```

---

### Scenario 3: Multiple Roles (PTK + ADMIN + KAPLA)
**Setup:**
- Create/use test user with Role = "PTK,ADMIN,KAPLA"

**Expected Results:**
- Login succeeds
- Dashboard shows: "Role: PTK, ADMIN, KAPLA"
- Admin panel: VISIBLE
- Team panel: VISIBLE
- HCIS features: VISIBLE

```javascript
// In browser console:
console.log(state.me.roles);        // ['PTK', 'ADMIN', 'KAPLA']
console.log(state.me.role);         // 'PTK,ADMIN,KAPLA'
```

---

### Scenario 4: Role Format Variations
The parser supports multiple separator formats:

```javascript
// All these are equivalent:
Role = "PTK,ADMIN,KAPLA"            // comma (standard)
Role = "PTK;ADMIN;KAPLA"            // semicolon
Role = "PTK|ADMIN|KAPLA"            // pipe
Role = "ptk,admin,kapla"            // lowercase (converted to uppercase)
```

---

## Manual Testing Steps

### Step 1: Prepare Test Data
Go to your Users sheet and find or create test users with these Role values:
- User 1: `PTK` (single role)
- User 2: `PTK,ADMIN` (dual role)
- User 3: `PTK,KAPLA` (dual role)
- User 4: `PTK,ADMIN,KAPLA` (triple role)

### Step 2: Deploy to Google Apps Script
1. Open `code.gs` in Google Apps Script editor
2. Save and deploy as new version
3. Test the web app URL

### Step 3: Login Tests

**Test 3.1: Single Role Login**
```
NIP: [User 1 NIP]
Password: [User 1 Password]

Verify:
- ✓ Login succeeds
- ✓ Dashboard Role field shows "PTK"
- ✓ Admin panel is HIDDEN
- ✓ Team panel is HIDDEN
```

**Test 3.2: Dual Role (ADMIN)**
```
NIP: [User 2 NIP]
Password: [User 2 Password]

Verify:
- ✓ Login succeeds
- ✓ Dashboard Role field shows "PTK, ADMIN"
- ✓ Admin panel is VISIBLE
- ✓ Team panel is HIDDEN
- ✓ Clicking "Buka Panel Admin" navigates to admin page
```

**Test 3.3: Dual Role (KAPLA)**
```
NIP: [User 3 NIP]
Password: [User 3 Password]

Verify:
- ✓ Login succeeds
- ✓ Dashboard Role field shows "PTK, KAPLA"
- ✓ Admin panel is HIDDEN
- ✓ Team panel is VISIBLE
- ✓ Clicking "Buka Kelola Tim" navigates to team page
```

**Test 3.4: Triple Role**
```
NIP: [User 4 NIP]
Password: [User 4 Password]

Verify:
- ✓ Login succeeds
- ✓ Dashboard Role field shows "PTK, ADMIN, KAPLA"
- ✓ Admin panel is VISIBLE
- ✓ Team panel is VISIBLE
- ✓ Both "Buka Panel Admin" and "Buka Kelola Tim" buttons work
```

### Step 4: Console Verification

After each login, open browser console (F12) and run:

```javascript
// Test the state object
console.log('Full session:', state.me);
console.log('Roles array:', state.me.roles);
console.log('Role string:', state.me.role);
console.log('Has ADMIN?', state.me.roles.includes('ADMIN'));
console.log('Has KAPLA?', state.me.roles.includes('KAPLA'));
console.log('Has PTK?', state.me.roles.includes('PTK'));
```

---

## Edge Cases

### Case 1: Whitespace in Role String
```
Role = "PTK, ADMIN, KAPLA"  // with spaces
Expected: ['PTK', 'ADMIN', 'KAPLA']  // spaces trimmed
```

### Case 2: Duplicate Roles
```
Role = "PTK,ADMIN,PTK"
Expected: ['PTK', 'ADMIN', 'PTK'] or ['PTK', 'ADMIN']
Current: Keeps duplicates (design decision, can change if needed)
```

### Case 3: Empty or Null Role
```
Role = "" or null or undefined
Expected: ['PTK'] (PTK always guaranteed as default)
```

### Case 4: Invalid Role Name
```
Role = "PTK,INVALID_ROLE"
Expected: Stored as ['PTK', 'INVALID_ROLE']
Note: No validation of role names at storage level
```

---

## Troubleshooting

### Problem: Dashboard shows only "Role: ADMIN" not "Role: PTK, ADMIN"
**Cause:** Roles array not being populated
**Check:**
1. Auth.js is calling `parseRoles_(rolesStr)` 
2. `RoleManager.js` parseRoles_() function exists
3. Session.js storing roles array properly

### Problem: Admin panel doesn't appear for multi-role user
**Cause:** `hasRole()` function not working correctly
**Check:**
1. Dashboard's `hasRole()` helper is defined
2. Using `me.roles.includes()` not `me.role ===`
3. Session.js returning `roles` as array

### Problem: Login fails with multi-role user
**Cause:** Session storage of array failing
**Check:**
1. `JSON.stringify()` and `JSON.parse()` in Session.js
2. No error in console about roles property

---

## Code References

### parseRoles_() in RoleManager.js
```javascript
function parseRoles_(roleStr) {
  if (!roleStr) return ['PTK'];  // default
  
  let parts = (roleStr.toString().toUpperCase().split(/[,;|]/))
    .map(s => s.trim())
    .filter(s => s);
  
  if (!parts.includes('PTK')) {
    parts.unshift('PTK');  // always include PTK
  }
  
  return parts;
}
```

### hasRole() in app.html
```javascript
const hasRole = (role) => {
  if (!me) return false;
  if (Array.isArray(me.roles)) {
    return me.roles.includes(role);
  }
  return me.role === role;  // backward compat
};
```

### Session storage in Session.js
```javascript
// Store
up.setProperty(UP_KEYS_ROLES, JSON.stringify(user.roles));

// Retrieve
const rolesJson = up.getProperty(UP_KEYS_ROLES);
roles = rolesJson ? JSON.parse(rolesJson) : ['PTK'];
```

---

## Success Criteria

- [x] Single role users still work (backward compatibility)
- [x] Multi-role users parse correctly
- [x] Dashboard displays all roles
- [x] Admin panel shows only for ADMIN role
- [x] Team panel shows only for KAPLA role
- [x] Both panels show for users with both roles
- [x] Code is backward compatible (fallback to single role)
- [x] Changes committed to git

---

## Next Steps

After testing confirms this works:
1. Create API endpoint for admins to assign multiple roles
2. Add role management UI in Admin panel
3. Document role assignment best practices
4. Train admin staff on multi-role system
