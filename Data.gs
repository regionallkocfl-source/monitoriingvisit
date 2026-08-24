function getHeaders_(sh) {
  return sh && sh.getLastColumn() ? sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x => String(x).trim()) : [];
}

function sheetObjects_(sheetName) {
  const sh = getSS_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return [];
  const values = sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getDisplayValues();
  const rawHead = values.shift().map(x => String(x).trim());
  const seen = {};
  const head = rawHead.map(h => {
    if (!h) return '';
    seen[h] = (seen[h] || 0) + 1;
    return seen[h] === 1 ? h : h + ' #' + seen[h];
  });
  return values.map((r, idx) => ({r:r, rowNumber:idx+2})).filter(x => x.r.some(v => String(v).trim() !== '')).map(x => {
    const o = {_row: x.rowNumber};
    head.forEach((h, i) => { if (h) o[h] = x.r[i]; });
    return o;
  });
}

function setByHeader_(sh, row, headers, name, value) {
  const i = headers.indexOf(name);
  if (i >= 0) sh.getRange(row, i + 1).setValue(value);
}

function clientBootstrapConfig_() {
  return {
    maxPhotos: MW.MAX_PHOTOS,
    editWindowHours: MW.EDIT_WINDOW_HOURS,
    languageDefault: 'EN'
  };
}

function getBootstrap(token) {
  const user = requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  // Keep session restore lightweight. Dashboard and CFL master are loaded only when
  // the user actually opens those screens.
  return {
    version: MW.VERSION,
    user: publicUser_(user),
    cfls: [],
    deferCfls: true,
    config: clientBootstrapConfig_()
  };
}

function getDashboard(token) {
  const user = requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  return getDashboardForUser_(user);
}

function dashboardCacheKey_(user) {
  const who = safeString_((user && (user.id || user.email || user.name)) || 'unknown').replace(/[^a-zA-Z0-9_.@-]/g,'_').slice(0,70);
  return 'MW_DASH_' + MW.VERSION + '_' + safeString_(user && user.role) + '_' + who;
}

function clearDashboardCacheForUser_(user) {
  try { CacheService.getScriptCache().remove(dashboardCacheKey_(user)); } catch (e) {}
}

function getDashboardForUser_(user) {
  const cache = CacheService.getScriptCache();
  const key = dashboardCacheKey_(user);
  try {
    const hit = cache.get(key);
    if (hit) return JSON.parse(hit);
  } catch (e) {}
  const result = computeDashboardForUser_(user);
  try { cache.put(key, JSON.stringify(result), 45); } catch (e) {}
  return result;
}

function computeDashboardForUser_(user) {
  const cfl = visibleSubmissions_(MW.SHEETS.CFL_SUB, user);
  const ses = visibleSubmissions_(MW.SHEETS.SESSION_SUB, user);
  const all = cfl.map(r => toCard_(r, 'CFL')).concat(ses.map(r => toCard_(r, 'SESSION')));
  const pct = all.map(x => Number(String(x.score || '').replace('%',''))).filter(Number.isFinite);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const month = today.slice(0,7);
  const visitKey = x => normalizeDateString_(x.visitDate || x.timestamp);
  const pending = user.role === MW.ROLES.ADMIN ? sheetObjects_(MW.SHEETS.EDIT_REQUESTS).filter(r => String(r.Status) === 'Pending').length : 0;
  const scoreTrend = all.slice().filter(x => Number.isFinite(Number(String(x.score||'').replace('%','')))).sort((a,b)=>visitKey(a).localeCompare(visitKey(b))).slice(-12).map(x=>({
    date: visitKey(x), score: Number(String(x.score||'').replace('%','')), type:x.type, cfl:x.cfl, id:x.id
  }));
  const tat = dashboardTat_(user);
  return {
    counts: {
      total: all.length,
      cfl: cfl.length,
      session: ses.length,
      today: all.filter(x => visitKey(x) === today).length,
      month: all.filter(x => visitKey(x).slice(0,7) === month).length,
      pendingEdits: pending,
      openActions: tat.openActions
    },
    averageScore: pct.length ? Math.round((pct.reduce((a,b) => a+b, 0) / pct.length) * 10) / 10 : 0,
    scoreTrend: scoreTrend,
    tatSummary: tat.summary,
    dueActions: tat.dueActions,
    recent: all.sort(cardDateSortDesc_).slice(0, 12)
  };
}

function getCflOptions(token) {
  const user = requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  return getCflOptionsForUser_(user);
}

function getCflOptionsForUser_(user) {
  const masters = sheetObjects_(MW.SHEETS.CFL);
  const overrides = sheetObjects_(MW.SHEETS.BLOCK_OVERRIDES);
  let rows = masters;
  if (user.role === MW.ROLES.AM) {
    const allowed = new Set((user.allowedCfls || []).map(x => x.toLowerCase()));
    rows = masters.filter(r => allowed.has(safeString_(r['CFL Name']).toLowerCase()));
  } else if (user.role === MW.ROLES.CBO) {
    // v3.7: District comes from Cbo Master. Every CFL in the same District is accessible.
    const districts = new Set((user.allowedDistricts || []).map(x => x.toLowerCase()));
    rows = masters.filter(r => districts.has(safeString_(r.District).toLowerCase()));
  }

  return rows.map(r => {
    const name = safeString_(r['CFL Name']);
    const bcc = normalizeCode_(r['BCC Code']);
    const ov = overrides.find(o => (bcc && normalizeCode_(o['BCC Code']) === bcc) || safeString_(o['CFL Name']).toLowerCase() === name.toLowerCase());
    const adj = [
      ov && ov['Adjacent Block 1'],
      ov && ov['Adjacent Block 2'],
      r['Adjacent Block 1'],
      r['Adjacent Block 2']
    ].map(x => safeString_(x)).filter(Boolean);
    return {
      cflName: name,
      bccCode: bcc,
      phase: safeString_(r.Phase),
      state: safeString_(r.State),
      district: safeString_(r.District),
      baseBlock: safeString_(r['Base Block']),
      adjacentBlocks: unique_(adj),
      bankName: safeString_(r['Bank Name'])
    };
  }).filter(x => x.cflName).sort((a,b) => a.cflName.localeCompare(b.cflName));
}

function normalizeCode_(v) {
  const s = String(v == null ? '' : v).trim();
  return s.endsWith('.0') ? s.slice(0,-2) : s;
}

function visibleSubmissions_(sheetName, user) {
  const rows = sheetObjects_(sheetName);
  if (user.role === MW.ROLES.ADMIN) return rows;
  const keys = [user.id, user.email, user.name].map(x => safeString_(x).toLowerCase()).filter(Boolean);
  return rows.filter(r => {
    const hay = [r['Logged-in Email'], r['Logged-in Google Account'], r['Officer Name'], r['Sign-off Name']].map(x => safeString_(x).toLowerCase());
    return keys.some(k => hay.some(h => h === k));
  });
}

function getMySubmissions(token, reportType, search) {
  const user = requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  const t = String(reportType || '').toUpperCase();
  let rows = [];
  if (!t || t === 'CFL') rows = rows.concat(visibleSubmissions_(MW.SHEETS.CFL_SUB, user).map(r => toCard_(r, 'CFL')));
  if (!t || t === 'SESSION') rows = rows.concat(visibleSubmissions_(MW.SHEETS.SESSION_SUB, user).map(r => toCard_(r, 'SESSION')));
  const q = safeString_(search).toLowerCase();
  if (q) rows = rows.filter(r => [r.id,r.type,r.cfl,r.district,r.visitDate,r.officer,r.grade].some(v => safeString_(v).toLowerCase().includes(q)));
  const reqs = sheetObjects_(MW.SHEETS.EDIT_REQUESTS);
  rows.forEach(x => {
    const rr = reqs.filter(r => r['Submission ID'] === x.id).sort((a,b) => b._row - a._row)[0];
    x.editStatus = rr ? safeString_(rr.Status) : '';
    x.editExpiry = rr ? safeString_(rr['Edit Expires At']) : '';
    x.canEdit = !!rr && String(rr.Status) === 'Approved';
  });
  return rows.sort(cardDateSortDesc_).slice(0, 500);
}

function toCard_(r, type) {
  return {
    id: safeString_(r['Submission ID']),
    type: type,
    timestamp: safeString_(r.Timestamp),
    cfl: safeString_(r['CFL Name']),
    district: safeString_(r.District),
    visitDate: safeString_(r['Visit Date']),
    officer: safeString_(r['Officer Name']),
    score: safeString_(r['Score %']),
    grade: safeString_(r.Grade),
    pdf: safeString_(r['PDF Link']),
    excel: safeString_(r['Excel Link']),
    pdfDownload: driveDownloadUrl_(r['PDF Link']),
    excelDownload: driveDownloadUrl_(r['Excel Link']),
    folder: safeString_(r['Drive Folder Link']),
    quality: safeString_(r['Data Quality Flags']),
    reportVersion: safeString_(r['Report Version']) || '1'
  };
}

function cardDateSortDesc_(a,b) {
  return String(b.timestamp || b.visitDate || '').localeCompare(String(a.timestamp || a.visitDate || ''));
}

function normalizeDateString_(value) {
  if (!value) return '';
  const s = String(value).trim();
  const ymd = s.match(/(\d{4})[-/]([01]?\d)[-/]([0-3]?\d)/);
  if (ymd) return ymd[1] + '-' + String(ymd[2]).padStart(2,'0') + '-' + String(ymd[3]).padStart(2,'0');
  const d = new Date(s);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function findSubmissionRow_(sheetName, submissionId) {
  const sh = getSS_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return null;
  const headers = getHeaders_(sh);
  const idx = headers.indexOf('Submission ID');
  if (idx < 0) return null;
  const finder = sh.getRange(2, idx+1, sh.getLastRow()-1, 1).createTextFinder(String(submissionId)).matchEntireCell(true);
  const cell = finder.findNext();
  if (!cell) return null;
  const display = sh.getRange(cell.getRow(),1,1,headers.length).getDisplayValues()[0];
  const raw = sh.getRange(cell.getRow(),1,1,headers.length).getValues()[0];
  const obj = {_row: cell.getRow(), _raw: raw};
  headers.forEach((h,i) => { if (h) obj[h] = display[i]; });
  return obj;
}

function getSubmissionDetails(token, submissionId, reportType) {
  const user = requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  const type = String(reportType || '').toUpperCase();
  const sheetName = type === 'CFL' ? MW.SHEETS.CFL_SUB : MW.SHEETS.SESSION_SUB;
  const row = findSubmissionRow_(sheetName, submissionId);
  if (!row) throw new Error('Submission not found.');
  if (user.role !== MW.ROLES.ADMIN && !visibleSubmissions_(sheetName, user).some(r => r['Submission ID'] === submissionId)) throw new Error('You cannot view this submission.');
  return submissionRowToDetails_(row, type);
}

function submissionRowToDetails_(row, type) {
  const defs = type === 'CFL' ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
  return {
    card: toCard_(row, type),
    meta: {
      cflName: row['CFL Name'], district: row.District, state: row.State, bccCode: row['BCC Code'], bankName: row['Bank Name'],
      phase: row.Phase, baseBlock: row['Base Block'], adjacentBlock: row['Adjacent Block'], visitDate: row['Visit Date'],
      officerName: row['Officer Name'], officerDesignation: row['Officer Designation'], photoUrls: splitList_(row['Photo URLs']),
      submitLocation: row['Submit Location'], submitMap: row['Submit Location Map Link'], fillDuration: row['Fill Duration (min)']
    },
    indicators: defs.map(d => {
      const p = d.submissionPrefix || shortPrefix_(d.titleEn);
      const acts = actionRowsForSource_(row['Submission ID']).filter(x => String(x['Indicator Sr']) === String(d.sr));
      return {sr:d.sr, titleEn:d.titleEn, titleHi:d.titleHi, score:row[p+' - Score'], observation:row[p+' - Observation'], suggestion:row[p+' - Suggestion'], timelineDate:row[p+' - Timeline Date'], completionDate:row[p+' - Completion Date'], suggestions:acts.map(x=>({actionId:x['Action ID'],suggestion:x.Suggestion,timelineDate:normalizeDateString_(x['Timeline Date']),completionDate:normalizeDateString_(x['Completion Date']),tatStatus:x['TAT Status'],tatDays:x['TAT Days']}))};
    }),
    followup: {previousSubmissionId:row['Previous Submission ID'], action:row['Follow-up Action'], completionDate:row['Follow-up Completion Date'], tat:row['Follow-up TAT']}
  };
}

function getPreviousFollowup_(type, cflName, excludeId) {
  const sheetName = type === 'CFL' ? MW.SHEETS.CFL_SUB : MW.SHEETS.SESSION_SUB;
  const rows = sheetObjects_(sheetName).filter(r => safeString_(r['CFL Name']).toLowerCase() === safeString_(cflName).toLowerCase() && r['Submission ID'] !== excludeId);
  if (!rows.length) return {exists:false, previousSubmissionId:'', visitDate:'', observation:'No previous data / कोई पिछला डेटा उपलब्ध नहीं', suggestion:'', recommended:'', completionDate:'', tat:'', items:[]};
  rows.sort((a,b) => { const bd=parseInputDate_(b.Timestamp||b['Visit Date']), ad=parseInputDate_(a.Timestamp||a['Visit Date']); return (bd?bd.getTime():0)-(ad?ad.getTime():0); });
  const prev = rows[0];
  // v3.7: carry forward ALL unresolved action points from this CFL, not only
  // the immediately previous submission. This prevents older observations from disappearing.
  const backlog = followupBacklogForCfl_(type, cflName, new Date());
  const items = backlog.length ? backlog : followupItemsForPrevious_(type, prev, new Date());
  const completedDates = items.map(x=>x.completionDate).filter(Boolean).sort();
  return {
    exists: true,
    previousSubmissionId: prev['Submission ID'],
    visitDate: normalizeDateString_(prev['Visit Date']),
    observation: pointwiseText_(items,'observation','No major pending observation recorded.'),
    suggestion: pointwiseText_(items,'suggestion',''),
    recommended: pointwiseText_(items,'suggestion',''),
    completionDate: completedDates.length ? completedDates[completedDates.length-1] : '',
    tat: tatSummaryLabel_(items),
    items: items,
    pdf: prev['PDF Link'] || '',
    excel: prev['Excel Link'] || ''
  };
}

function getPreviousFollowup(token, reportType, cflName, excludeId) {
  const user = requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  const requested = safeString_(cflName).toLowerCase();
  if (user.role !== MW.ROLES.ADMIN) {
    const allowed = getCflOptionsForUser_(user).some(c => safeString_(c.cflName).toLowerCase() === requested);
    if (!allowed) throw new Error('This CFL is not mapped to your current access.');
  }
  return getPreviousFollowup_(String(reportType || '').toUpperCase(), cflName, excludeId || '');
}
