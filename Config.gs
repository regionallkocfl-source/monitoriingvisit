const MW = Object.freeze({
  VERSION: '3.9.0',
  SHEETS: {
    EMPLOYEE: 'Employee Master',
    CBO: 'Cbo Master',
    CFL: 'CFL Master',
    BLOCK_OVERRIDES: 'CFL Block Overrides',
    ADMINS: 'Admin Users',
    PINS: 'Employee PINs',
    EDIT_REQUESTS: 'Edit Requests',
    CFL_SUB: 'CFL Submissions',
    SESSION_SUB: 'Session Submissions',
    CFL_TEMPLATE: 'CFL Reporting format',
    SESSION_TEMPLATE: 'Session Monitoring Report',
    TRANSLATIONS: 'Form Translations',
    AUDIT: 'Audit Log',
    FOLLOWUP_ACTIONS: 'Follow-up Actions',
    REMINDER_LOG: 'Reminder Log'
  },
  ROLES: { ADMIN: 'ADMIN', AM: 'AM', CBO: 'CBO' },
  ADMIN_PERMS: { ADMIN: 'ADMIN', EDIT: 'EDIT', VIEW: 'VIEW' },
  REPORTS: { CFL: 'CFL', SESSION: 'SESSION' },
  EDIT_WINDOW_HOURS: 24,
  SESSION_TTL_SECONDS: 21600,
  ROOT_FOLDER_NAME: 'MoneyWise CFL Monitoring Reports',
  STAGING_FOLDER_NAME: '_Pending Uploads',
  MAX_PHOTOS: 8,
  MAX_PHOTO_BYTES: 1800000,
  TAT_GRACE_DAYS: 7
});

function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('System not initialized. Run setupMoneyWise() once from the bound spreadsheet.');
  return SpreadsheetApp.openById(id);
}

function setupMoneyWise() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this Apps Script project from the MoneyWise Google Sheet, then run setupMoneyWise().');

  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  props.setProperty('APP_VERSION', MW.VERSION);
  try { CacheService.getScriptCache().removeAll(['MW_FORM_DEF_'+MW.VERSION+'_CFL','MW_FORM_DEF_'+MW.VERSION+'_SESSION']); } catch (e) {}

  const root = getOrCreateRootFolder_();
  props.setProperty('ROOT_FOLDER_ID', root.getId());

  ensureSheet_(MW.SHEETS.TRANSLATIONS, ['Key','Report Type','Template Row','English','Hindi','Updated At']);
  ensureSheet_(MW.SHEETS.AUDIT, ['Timestamp','User','Action','Entity Type','Entity ID','Details']);
  ensureSheet_(MW.SHEETS.FOLLOWUP_ACTIONS, ['Action ID','Source Submission ID','Report Type','CFL Name','District','Visit Date','Indicator Sr','Indicator Title','Observation','Suggestion','Timeline Date','Status','Completion Date','Completed In Submission ID','Delay Days','TAT Days','TAT Status','Created By','Created At','Completed By','Updated At']);
  ensureSheet_(MW.SHEETS.REMINDER_LOG, ['Reminder ID','Date','Mode','Recipient','Consultant ID','Consultant Name','Action Count','Action IDs','Sent At','Sent By','Status','Error']);

  ensureColumns_(MW.SHEETS.ADMINS, ['Email','Permission','Added At','Active','Display Name','Admin PIN','Updated At']);
  ensureColumns_(MW.SHEETS.PINS, ['Officer Name','PIN (4 digit, blank = no PIN required)','Consultant ID','Email','Active','Updated At']);
  ensureColumns_(MW.SHEETS.EDIT_REQUESTS, ['Request ID','Submission ID','Report Type','Requested By','Reason','Status','Requested At','Approved By','Approved At','Edit Expires At','Edited At','Decision Note']);
  ensureColumns_(MW.SHEETS.CFL_SUB, ['Consultant ID','Consultant Name','Drive Folder Link','Data Backup Link','Digital Signature URL','Previous Submission ID','Follow-up Action','Follow-up Completion Date','Follow-up TAT','TAT Green Count','TAT Amber Count','TAT Red Count','TAT Pending Count','Last Edited At','Edit Count','Report Version','Data Quality Flags']);
  ensureColumns_(MW.SHEETS.SESSION_SUB, ['Consultant ID','Consultant Name','Drive Folder Link','Data Backup Link','Digital Signature URL','Previous Submission ID','Follow-up Action','Follow-up Completion Date','Follow-up TAT','TAT Green Count','TAT Amber Count','TAT Red Count','TAT Pending Count','Last Edited At','Edit Count','Report Version','Data Quality Flags']);

  bootstrapDeployerAdmin_();

  let translations = {created: 0, failed: 0};
  try { translations = initializeTranslations_(); } catch (e) { audit_('SYSTEM','TRANSLATION_SETUP_WARNING','SYSTEM','',e.message); }
  let migratedActions = 0;
  try { migratedActions = migrateLegacyFollowupActions_(); } catch (e) { audit_('SYSTEM','FOLLOWUP_MIGRATION_WARNING','SYSTEM','',e.message); }
  try { repairPhotoHyperlinks_(); } catch (e) { audit_('SYSTEM','PHOTO_LINK_REPAIR_WARNING','SYSTEM','',e.message); }

  return {
    ok: true,
    version: MW.VERSION,
    spreadsheetId: ss.getId(),
    rootFolderId: root.getId(),
    translations: translations,
    migratedFollowupActions: migratedActions,
    reminderAutomation: (typeof reminderAutomationStatus_ === 'function' ? reminderAutomationStatus_() : {enabled:false})
  };
}

function getOrCreateRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty('ROOT_FOLDER_ID');
  if (existingId) {
    try { return DriveApp.getFolderById(existingId); } catch (e) {}
  }
  const it = DriveApp.getFoldersByName(MW.ROOT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(MW.ROOT_FOLDER_NAME);
}

function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || getSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}

function ensureColumns_(sheetName, headers) {
  const sh = ensureSheet_(sheetName, headers);
  let existing = [];
  if (sh.getLastColumn() > 0) existing = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x => String(x).trim());
  headers.forEach(h => {
    if (!existing.includes(h)) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      existing.push(h);
    }
  });
  sh.setFrozenRows(1);
  return sh;
}

function bootstrapDeployerAdmin_() {
  const email = String(Session.getEffectiveUser().getEmail() || '').trim();
  if (!email) return;
  const sh = getSS_().getSheetByName(MW.SHEETS.ADMINS);
  if (!sh) return;
  const rows = sheetObjects_(MW.SHEETS.ADMINS);
  const hasFullAdmin = rows.some(r => isActiveRecord_(r) && normalizeAdminPermission_(r.Permission) === MW.ADMIN_PERMS.ADMIN);
  if (hasFullAdmin) return;

  const existing = rows.find(r => String(r.Email || '').toLowerCase() === email.toLowerCase());
  const headers = getHeaders_(sh);
  if (existing) {
    setByHeader_(sh, existing._row, headers, 'Permission', MW.ADMIN_PERMS.ADMIN);
    setByHeader_(sh, existing._row, headers, 'Active', true);
    if (!safeString_(existing['Admin PIN'])) setByHeader_(sh, existing._row, headers, 'Admin PIN', String(Math.floor(100000 + Math.random()*900000)));
    setByHeader_(sh, existing._row, headers, 'Updated At', new Date());
  } else {
    const rec = {'Email': email, 'Permission': MW.ADMIN_PERMS.ADMIN, 'Added At': new Date(), 'Active': true, 'Display Name': 'System Administrator', 'Admin PIN': String(Math.floor(100000 + Math.random()*900000)), 'Updated At': new Date()};
    sh.appendRow(headers.map(h => rec[h] !== undefined ? rec[h] : ''));
  }
}

function normalizeAdminPermission_(value) {
  const p = String(value || '').trim().toUpperCase();
  if (p === 'ADMIN' || p === 'FULL') return MW.ADMIN_PERMS.ADMIN;
  if (p === 'EDIT' || p === 'EDITOR') return MW.ADMIN_PERMS.EDIT;
  return MW.ADMIN_PERMS.VIEW;
}

function isActiveRecord_(row) {
  const v = row && row.Active;
  if (v === '' || v === undefined || v === null) return true;
  const s = String(v).trim().toLowerCase();
  return !['false','0','no','inactive','disabled'].includes(s);
}

function safeString_(value, maxLen) {
  const s = String(value == null ? '' : value).trim();
  return maxLen ? s.slice(0, maxLen) : s;
}

function parseInputDate_(value) {
  if (!value) return '';
  if (value instanceof Date) return value;
  const s = String(value).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
  const d = new Date(value);
  return isNaN(d.getTime()) ? '' : d;
}

function formatDateYmd_(value) {
  const d = parseInputDate_(value);
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
}
