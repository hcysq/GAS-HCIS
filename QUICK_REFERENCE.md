# Quick Reference: Multiple Roles System

## TL;DR
Users can now have multiple roles. Instead of Role = "ADMIN", use Role = "PTK,ADMIN,KAPLA".

## Key Functions

### RoleManager.js
```javascript
parseRoles_(roleStr)              // "PTK,ADMIN" → ['PTK','ADMIN']
getUserRoles()                    // Returns: ['PTK','ADMIN']
getUserRole()                     // Returns: 'PTK,ADMIN' (backward compat)
hasRole('ADMIN')                  // true if user has ADMIN role
isAdmin()                         // true if has ADMIN role
isManager()                       // true if has ADMIN or KAPLA
```

### Session.js
```javascript
setSession_(user)                 // Stores user with roles array
getSession_()                     // Returns { ..., roles: [...] }
requireLogin_()                   // Ensures session exists
```

### Auth.js
```javascript
authLogin(nip, pin)               // Handles login & role parsing
authMe()                          // Returns current user with roles
```

### app.html
```javascript
hasRole('ADMIN')                  // Check if current user has role
state.me.roles                    // Array: ['PTK','ADMIN','KAPLA']
state.me.role                     // String: "PTK,ADMIN,KAPLA"
```

## Usage Examples

### Example 1: Check if user is admin
```javascript
const me = state.me;
if (me.roles && me.roles.includes('ADMIN')) {
  // Show admin panel
}

// OR using helper
if (hasRole('ADMIN')) {
  // Show admin panel
}
```

### Example 2: Check multiple roles at once
```javascript
const isManager = me.roles && 
  (me.roles.includes('ADMIN') || me.roles.includes('KAPLA'));

if (isManager) {
  // Show manager options
}
```

### Example 3: Display all roles
```javascript
const roleText = me.roles.join(', ');  // "PTK, ADMIN, KAPLA"
const roleHtml = `<b>${escapeHtml(roleText)}</b>`;
```

### Example 4: API level checks
```javascript
function doGet(e) {
  const user = requireLogin_();
  
  if (!user.roles.includes('ADMIN')) {
    throw new Error('UNAUTHORIZED');
  }
  
  // Admin-only logic here
}
```

## Data Format

### In Google Sheets (Users sheet)
```
Role: "PTK"                    // Single role
Role: "PTK,ADMIN"              // Two roles
Role: "PTK,ADMIN,KAPLA"        // Three roles
Role: "ptk,admin,kapla"        // Lowercase (converted to uppercase)
Role: "PTK; ADMIN; KAPLA"      // Semicolon separator (also works)
Role: "PTK|ADMIN|KAPLA"        // Pipe separator (also works)
```

### In Session (state.me)
```javascript
{
  nip: "100",
  nama: "John Doe",
  role: "PTK,ADMIN,KAPLA",           // String (original)
  roles: ["PTK", "ADMIN", "KAPLA"],  // Array (parsed)
  email: "john@example.com",
  userId: "..."
}
```

### In UserProperties (persistent storage)
```javascript
PropertiesService.getUserProperties().getProperty('HCIS_ROLES')
// Returns: '["PTK","ADMIN","KAPLA"]' (JSON string)

// When parsed:
JSON.parse('["PTK","ADMIN","KAPLA"]')
// Returns: ["PTK", "ADMIN", "KAPLA"]
```

## Common Patterns

### Pattern 1: Conditional UI Rendering
```html
<div>${isAdmin ? `
  <div class="card">
    <button onclick="goto('admin')">Admin Panel</button>
  </div>
` : ''}</div>
```

### Pattern 2: Check Before Action
```javascript
function deleteUser(userId) {
  if (!state.me.roles.includes('ADMIN')) {
    alert('Only admins can delete users');
    return;
  }
  // Delete logic here
}
```

### Pattern 3: Multiple Role Requirements
```javascript
const canApprove = me.roles && 
  (me.roles.includes('KAPLA') || me.roles.includes('ADMIN'));

if (canApprove) {
  // Show approval button
}
```

### Pattern 4: Log User Access
```javascript
const roles = me.roles.join(', ');
console.log(`User ${me.nip} (${roles}) accessed admin panel`);
```

## Troubleshooting Checklist

- [ ] User has comma-separated role value in Role column
- [ ] RoleManager.js has parseRoles_() function
- [ ] Auth.js calls parseRoles_() on login
- [ ] Session.js uses JSON.stringify/parse for roles
- [ ] app.html uses me.roles not me.role for checks
- [ ] Browser cache cleared (Ctrl+Shift+Delete)
- [ ] Google Apps Script redeployed after changes

## Migration Checklist

If updating an existing system:

- [ ] Back up Users sheet
- [ ] Update existing single roles to multi-role format if needed:
  - "ADMIN" → "PTK,ADMIN"
  - "KAPLA" → "PTK,KAPLA"
  - Keep "PTK" as "PTK"
- [ ] Deploy updated code
- [ ] Test with each role type
- [ ] Verify session persistence works
- [ ] Check conditional rendering appears correctly

## Performance Notes

- Role array is stored in session (memory efficient)
- JSON serialization used for UserProperties (minimal overhead)
- hasRole() is O(n) where n = number of roles (typically 2-3, so negligible)
- No additional database queries needed

## Security Notes

- Roles are stored in session which is tied to user authentication
- Clear session on logout (clearSession_())
- Validate roles on server side before sensitive operations
- Don't trust client-side role checks for security-critical actions

## Backward Compatibility

| Old Format | New Format | Result |
|-----------|-----------|--------|
| "ADMIN" | Single role | Works automatically, roles = ['ADMIN'] |
| "ADMIN" | Added as "PTK,ADMIN" | Works, roles = ['PTK', 'ADMIN'] |
| Old code checking `me.role === 'ADMIN'` | Still works | `me.role` still available as string |

---

**Last Updated:** 2024
**Status:** Production Ready
**Version:** 1.0 (Multiple Roles Support)
