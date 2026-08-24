function loginUser(identifier) {
  identifier = safeString_(identifier, 150);
  if (!identifier) throw new Error('Consultant ID / Email required.');

  const admin = findAdmin_(identifier);
  if (admin) {
    if (!isActiveRecord_(admin)) throw new Error('This admin account is inactive.');
    // v3.8: no PIN on the login screen. If Google exposes the signed-in user,
    // enforce that it matches the configured Admin email. Some Apps Script
    // deployments return a blank ActiveUser email; in that case the active
    // Admin Users record + entered email remains the identity check.
    const activeGoogle = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    const adminEmail = String(admin.Email || '').trim().toLowerCase();
    if (activeGoogle && activeGoogle !== adminEmail) {
      throw new Error('Please open the app with the Google account configured for this Admin ID.');
    }
    return createSession_({
      role: MW.ROLES.ADMIN,
      id: safeString_(admin.Email),
      name: safeString_(admin['Display Name']) || safeString_(admin.Email),
      email: safeString_(admin.Email),
      designation: 'Administrator',
      permission: normalizeAdminPermission_(admin.Permission),
      allowedCfls: [],
      allowedDistricts: []
    });
  }

  const profile = findOperationalUser_(identifier);
  if (!profile) throw new Error('User not found in Employee Master / Cbo Master.');
  return createSession_(profile);
}

function createSession_(user) {
  user.loginAt = new Date().toISOString();
  user.activeGoogleEmail = String(Session.getActiveUser().getEmail() || '');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('MWSESS:' + token, JSON.stringify(user), MW.SESSION_TTL_SECONDS);
  audit_(user.id || user.email, 'LOGIN', 'USER', user.id || user.email, user.role + '/' + (user.permission || ''));
  return {
    token: token,
    version: MW.VERSION,
    user: publicUser_(user),
    config: (typeof clientBootstrapConfig_ === 'function' ? clientBootstrapConfig_() : {
      maxPhotos: MW.MAX_PHOTOS,
      editWindowHours: MW.EDIT_WINDOW_HOURS,
      languageDefault: 'EN'
    })
  };
}

function publicUser_(user) {
  return {
    role: user.role,
    id: user.id,
    name: user.name,
    email: user.email,
    designation: user.designation,
    permission: user.permission || '',
    allowedCflCount: (user.allowedCfls || []).length,
    allowedDistricts: user.allowedDistricts || []
  };
}

function requireSession_(token, roles) {
  const key = 'MWSESS:' + String(token || '');
  const raw = CacheService.getScriptCache().get(key);
  if (!raw) throw new Error('Session expired. Please login again.');
  const user = JSON.parse(raw);
  if (roles && !roles.includes(user.role)) throw new Error('Permission denied.');
  CacheService.getScriptCache().put(key, raw, MW.SESSION_TTL_SECONDS);
  return user;
}

function requireAdminWrite_(token) {
  const user = requireSession_(token, [MW.ROLES.ADMIN]);
  const p = normalizeAdminPermission_(user.permission);
  if (p === MW.ADMIN_PERMS.VIEW) throw new Error('View-only admin account cannot make changes.');
  return user;
}

function requireFullAdmin_(token) {
  const user = requireSession_(token, [MW.ROLES.ADMIN]);
  if (normalizeAdminPermission_(user.permission) !== MW.ADMIN_PERMS.ADMIN) throw new Error('Full Admin permission required.');
  return user;
}

function logout(token) {
  CacheService.getScriptCache().remove('MWSESS:' + String(token || ''));
  return true;
}

function findAdmin_(identifier) {
  const q = String(identifier || '').toLowerCase();
  return sheetObjects_(MW.SHEETS.ADMINS).find(r => String(r.Email || '').toLowerCase() === q) || null;
}

function findOperationalUser_(identifier) {
  const q = String(identifier || '').trim().toLowerCase();

  // v3.7: CBO access is controlled ONLY by Cbo Master -> District.
  // Ignore the legacy "Additional Cfl also" and duplicate mapping columns for access.
  // Each CBO sees every CFL in CFL Master whose District matches their Cbo Master District.
  const cboRows = sheetObjects_(MW.SHEETS.CBO).filter(r => {
    return [r['Consultant ID'], r['Official Email ID'], r['Consultant Name']]
      .some(v => String(v || '').trim().toLowerCase() === q);
  });
  if (cboRows.length) {
    const first = cboRows[0];
    return {
      role: MW.ROLES.CBO,
      id: safeString_(first['Consultant ID'] || identifier),
      name: safeString_(first['Consultant Name'] || identifier),
      email: safeString_(first['Official Email ID']),
      designation: safeString_(first['Consultant Designation']) || 'Consultant (Capacity Building Officer)',
      permission: '',
      allowedCfls: [],
      allowedDistricts: unique_(cboRows.map(r => safeString_(r.District)).filter(Boolean)),
      loginExempt: false
    };
  }

  // AM / operational employee mapping remains controlled by Employee Master.
  const empRows = sheetObjects_(MW.SHEETS.EMPLOYEE).filter(r => rowMatchesIdentity_(r, q, ['Consultant Id','Consultant ID','Email','Email Id','Name']));
  if (empRows.length) {
    const first = empRows[0];
    const desig = safeString_(first.Designation || first['Consultant Designation']);
    const role = desig.toLowerCase().includes('area manager') ? MW.ROLES.AM : MW.ROLES.CBO;
    const allowedCfls = unique_(empRows.map(r => safeString_(r['CFL Name'])).filter(Boolean));
    const allowedDistricts = unique_(empRows.map(r => safeString_(r.District)).filter(Boolean));
    return {
      role: role,
      id: safeString_(first['Consultant Id'] || first['Consultant ID'] || identifier),
      name: safeString_(first.Name || first['Consultant Name'] || identifier),
      email: safeString_(first.Email || first['Email Id'] || first['Official Email ID']),
      designation: desig,
      permission: '',
      allowedCfls: allowedCfls,
      allowedDistricts: allowedDistricts,
      loginExempt: truthy_(first['Login Exempt'])
    };
  }
  return null;
}

function rowMatchesIdentity_(row, q, fields) {
  return fields.some(f => String(row[f] || '').trim().toLowerCase() === q);
}

function cboRowMatchesIdentity_(row, q) {
  return [
    row['Consultant ID'], row['Official Email ID'], row['Consultant Name'],
    row['Consultant ID #2'], row['Consultant ID '], row['Consultant ID 2'], row['Name '], row.Name
  ].some(v => String(v || '').trim().toLowerCase() === q) ||
  String(row['Consultant ID'] || '').trim().toLowerCase() === q ||
  String(row['Name '] || '').trim().toLowerCase() === q;
}

function isPinValidForProfile_(profile, pin) {
  if (profile.loginExempt) return true;
  const rows = sheetObjects_(MW.SHEETS.PINS);
  const rec = rows.find(r => {
    const identities = [r['Consultant ID'], r['Consultant Id'], r.Email, r['Officer Name']].map(v => String(v || '').trim().toLowerCase());
    return [profile.id, profile.email, profile.name].map(v => String(v || '').trim().toLowerCase()).some(v => v && identities.includes(v));
  });
  if (!rec || !isActiveRecord_(rec)) return true;
  const pinHeader = Object.keys(rec).find(k => k.toLowerCase().includes('pin') && !k.startsWith('_'));
  const stored = pinHeader ? String(rec[pinHeader] || '').trim() : '';
  if (!stored) return true;
  return stored === String(pin || '').trim();
}

function truthy_(v) {
  return ['1','true','yes','y','active'].includes(String(v || '').trim().toLowerCase());
}

function unique_(arr) {
  return Array.from(new Set((arr || []).map(x => String(x || '').trim()).filter(Boolean)));
}

function splitList_(value) {
  return String(value || '').split(/[,;\n]+/).map(x => x.trim()).filter(Boolean);
}
