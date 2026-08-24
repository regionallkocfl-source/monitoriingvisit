function submitReport(token, payload) {
  const user = requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  payload = payload || {};
  const type = String(payload.reportType || '').toUpperCase();
  if (![MW.REPORTS.CFL, MW.REPORTS.SESSION].includes(type)) throw new Error('Invalid report type.');
  validatePayload_(user, type, payload);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheetName = type === MW.REPORTS.CFL ? MW.SHEETS.CFL_SUB : MW.SHEETS.SESSION_SUB;
    const sh = getSS_().getSheetByName(sheetName);
    if (!sh) throw new Error(sheetName + ' sheet missing.');
    const headers = getHeaders_(sh);
    const clientToken = safeString_(payload.clientToken || Utilities.getUuid(), 120);
    payload.clientToken = clientToken;

    const duplicate = findRowByHeaderValue_(sh, headers, 'Client Token', clientToken);
    if (duplicate && duplicate['PDF Link'] && duplicate['Excel Link']) return responseFromSubmissionRow_(duplicate, type, true);

    const id = duplicate ? duplicate['Submission ID'] : ((type === 'CFL' ? 'CFL-' : 'SESS-') + Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase());
    const previous = getPreviousFollowup_(type, payload.cflName, id);
    payload.previousFollowup = mergeFollowupUpdatesForDisplay_(previous, payload.followupItems || [], payload.visitDate);
    if (!payload.previousSubmissionId) payload.previousSubmissionId = previous.previousSubmissionId || '';
    payload.followupCompletionDate = payload.previousFollowup.completionDate || '';
    payload.followupTat = payload.previousFollowup.tat || 'Not applicable';

    const folder = getSubmissionFolder_(payload.visitDate, payload.officerName || user.name, payload.cflName, id, user);
    const photos = finalizeStagedPhotos_(clientToken, folder, user);
    payload.photoUrls = photos.length ? photos.map(x => x.url) : (duplicate ? splitList_(duplicate['Photo URLs']) : []);
    // v3.6: pass Drive IDs directly to the PDF generator for reliable image embedding.
    payload.photoFileIds = photos.length ? photos.map(x => x.id) : payload.photoUrls.map(extractDriveId_).filter(Boolean);

    const version = duplicate ? Math.max(1, Number(duplicate['Report Version'] || 1)) : 1;
    let signatureInfo = saveSignature_(payload.signatureData, folder, id, version);
    if (!signatureInfo && duplicate && duplicate['Digital Signature URL']) {
      try {
        const oldSignatureId = extractDriveId_(duplicate['Digital Signature URL']);
        if (oldSignatureId) signatureInfo = {id:oldSignatureId,url:duplicate['Digital Signature URL'],blob:DriveApp.getFileById(oldSignatureId).getBlob()};
      } catch (e) {}
    }
    payload.signatureUrl = signatureInfo ? signatureInfo.url : (duplicate ? duplicate['Digital Signature URL'] : '');

    const record = buildRecord_(type, id, user, payload, version, duplicate ? Number(duplicate['Edit Count'] || 0) : 0);
    if (duplicate) delete record['Timestamp'];
    record['Drive Folder Link'] = folder.getUrl();
    record['Digital Signature URL'] = payload.signatureUrl || '';
    const backup = saveSubmissionBackup_(folder, id, type, payload, user, version);
    record['Data Backup Link'] = backup.url;

    let rowNum;
    if (duplicate) {
      rowNum = duplicate._row;
      const current = sh.getRange(rowNum,1,1,headers.length).getValues()[0];
      const newRow = headers.map((h,i) => record[h] !== undefined ? record[h] : current[i]);
      sh.getRange(rowNum,1,1,headers.length).setValues([newRow]);
    } else {
      sh.appendRow(headers.map(h => record[h] !== undefined ? record[h] : ''));
      rowNum = sh.getLastRow();
    }
    setPhotoLinksCell_(sh, rowNum, headers, payload.photoUrls || []);

    // Point-wise follow-up: close previous action points and create current suggestions as new action points.
    applyFollowupUpdates_(payload.followupItems || [], id, user, payload.visitDate);
    upsertSubmissionActions_(type, id, payload, user);

    let links = null;
    let reportError = '';
    try {
      links = createReportFiles_(type, id, payload, user, folder, version, signatureInfo ? signatureInfo.blob : null);
      if (links.pdf) setByHeader_(sh,rowNum,headers,'PDF Link',links.pdf);
      if (links.excel) setByHeader_(sh,rowNum,headers,'Excel Link',links.excel);
      reportError = links.reportError || '';
      shareSubmissionAssets_(folder, [links.pdfId, links.excelId, backup.id].filter(Boolean).concat(signatureInfo ? [signatureInfo.id] : []).concat(photos.map(x => x.id)), user);
      if (links.pdfId && links.excelId) notifyAdminsOfSubmission_(type, id, payload, user, links, folder, version, duplicate ? 'UPDATE' : 'SUBMIT');
      if (reportError) {
        const flags = [record['Data Quality Flags'], 'REPORT_GENERATION_PARTIAL: ' + reportError].filter(Boolean).join(' | ');
        setByHeader_(sh,rowNum,headers,'Data Quality Flags',flags.slice(0,1000));
      }
    } catch (e) {
      reportError = e.message;
      const flags = [record['Data Quality Flags'], 'REPORT_GENERATION_FAILED: ' + e.message].filter(Boolean).join(' | ');
      setByHeader_(sh,rowNum,headers,'Data Quality Flags',flags.slice(0,1000));
      audit_(user.id || user.email, 'REPORT_GENERATION_FAILED', type, id, e.stack || e.message);
    }

    try { clearDashboardCacheForUser_(user); } catch(e) {}
    audit_(user.id || user.email, duplicate ? 'RETRY_SUBMIT' : 'SUBMIT', type, id, payload.cflName);
    const saved = findSubmissionRow_(sheetName, id);
    const response = responseFromSubmissionRow_(saved, type, !!duplicate);
    response.reportError = reportError;
    return response;
  } finally {
    lock.releaseLock();
  }
}

function validatePayload_(user, type, p) {
  const required = ['cflName','visitDate','officerName','officerDesignation'];
  required.forEach(k => { if (!safeString_(p[k])) throw new Error(k + ' is required.'); });
  const visit = parseInputDate_(p.visitDate);
  if (!visit) throw new Error('Valid Visit Date required.');
  const tomorrow = new Date(); tomorrow.setHours(23,59,59,999); tomorrow.setDate(tomorrow.getDate()+1);
  if (visit > tomorrow) throw new Error('Visit Date cannot be in the future.');

  const cfls = getCflOptionsForUser_(user);
  if (user.role !== MW.ROLES.ADMIN && !cfls.some(x => x.cflName.toLowerCase() === safeString_(p.cflName).toLowerCase())) throw new Error('This CFL is not mapped to your login.');

  const defs = type === 'CFL' ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
  const answers = p.indicators || [];
  if (answers.length !== defs.length) throw new Error('All indicators must be loaded before submission.');
  defs.forEach((d,i) => {
    const a = answers[i] || {};
    if (a.score === '' || a.score === null || a.score === undefined) throw new Error('Score missing for indicator ' + d.sr + '.');
    const score = Number(a.score);
    if (!Number.isFinite(score) || score < 0 || score > Number(d.maxScore)) throw new Error('Invalid score for indicator ' + d.sr + '.');
    a.observation = safeString_(a.observation, 4000);
    const sugRows = normalizedSuggestions_(a);
    sugRows.forEach((x,j) => {
      if (x.suggestion && !x.timelineDate) throw new Error('Timeline Date required for indicator ' + d.sr + ', suggestion ' + (j+1) + '.');
      const td=parseInputDate_(x.timelineDate); if(td && td < visit) throw new Error('Timeline Date cannot be before Visit Date for indicator ' + d.sr + ', suggestion ' + (j+1) + '.');
    });
    a.suggestions = sugRows;
    a.suggestion = sugRows.map(x=>x.suggestion).filter(Boolean).join('\n');
    a.timelineDate = sugRows.length ? sugRows[0].timelineDate : '';
  });

  if (type === 'SESSION') {
    if (!safeString_(p.sessionType)) throw new Error('Type of Session is required.');
    if (!safeString_(p.venue)) throw new Error('Venue of Session is required.');
  }
}

function buildRecord_(type, id, user, p, version, editCount) {
  const r = {
    'Submission ID': id,
    'Timestamp': new Date(),
    'Report Type': type,
    'CFL Name': safeString_(p.cflName),
    'District': safeString_(p.district),
    'State': safeString_(p.state),
    'BCC Code': normalizeCode_(p.bccCode),
    'Bank Name': safeString_(p.bankName),
    'Phase': safeString_(p.phase),
    'Base Block': safeString_(p.baseBlock),
    'Adjacent Block': safeString_(p.adjacentBlock),
    'Visit Date': parseInputDate_(p.visitDate),
    'Officer Name': safeString_(p.officerName || user.name),
    'Officer Designation': safeString_(p.officerDesignation || user.designation),
    'Consultant ID': safeString_(user.id),
    'Consultant Name': safeString_(user.name || p.officerName),
    'Logged-in Google Account': String(Session.getActiveUser().getEmail() || 'Not detected'),
    'Logged-in Email': safeString_(user.email || user.id),
    'Form Opened At': p.formOpenedAt ? new Date(p.formOpenedAt) : '',
    'Fill Duration (min)': Number(p.fillDuration || 0) || '',
    'Client Token': safeString_(p.clientToken,120),
    'Photo Location': safeString_(p.photoLocation,300),
    'Photo Location Map Link': mapLinkFromLocation_(p.photoLocation),
    'Submit Location': safeString_(p.submitLocation,300),
    'Submit Location Map Link': mapLinkFromLocation_(p.submitLocation),
    'Photo URLs': (p.photoUrls || []).join(', '),
    'Sign-off Name': safeString_(p.signoffName || user.name),
    'Sign-off Designation': safeString_(p.signoffDesignation || user.designation),
    'Sign-off District': safeString_(p.signoffDistrict || p.district),
    'Sign-off Zone': safeString_(p.signoffZone),
    'Recipient Emails': '',
    'Previous Submission ID': safeString_(p.previousSubmissionId),
    'Follow-up Action': safeString_(p.followupAction,4000),
    'Follow-up Completion Date': parseInputDate_(p.followupCompletionDate) || '',
    'Follow-up TAT': safeString_(p.followupTat),
    'Edit Count': Number(editCount || 0),
    'Report Version': Number(version || 1)
  };

  if (type === 'SESSION') {
    r['Consultant Staff Designation'] = safeString_(p.consultantStaffDesignation,500);
    r['Session Observer / Stakeholders'] = safeString_(p.observer,2000);
    r['Type of Session'] = safeString_(p.sessionType,200);
    r['Venue of Session'] = safeString_(p.venue,2000);
  }

  const defs = type === 'CFL' ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
  let total = 0, max = 0;
  const timelineAnswers = [];
  defs.forEach((d,i) => {
    const a = (p.indicators || [])[i] || {};
    const prefix = d.submissionPrefix || shortPrefix_(d.titleEn);
    const score = Number(a.score || 0);
    total += score;
    max += Number(d.maxScore || 1);
    r[prefix + ' - Score'] = score;
    r[prefix + ' - Observation'] = safeString_(a.observation,4000);
    const sugRows = normalizedSuggestions_(a);
    r[prefix + ' - Suggestion'] = sugRows.map((x,j)=>(sugRows.length>1?(j+1)+'. ':'')+x.suggestion).filter(Boolean).join('\n');
    r[prefix + ' - Timeline Date'] = sugRows.length ? (parseInputDate_(sugRows[0].timelineDate) || '') : '';
    // Completion date is captured during the NEXT visit follow-up, not while creating the current observation.
    r[prefix + ' - Completion Date'] = '';
    timelineAnswers.push(a);
    if (type === 'CFL' && prefix === 'Staffing') r['Staffing - Staffing Details'] = safeString_(a.staffingDetails,4000);
  });

  const pct = max ? Math.round((total / max) * 1000) / 10 : 0;
  r['Total Score'] = Math.round(total * 100) / 100;
  r['Max Score'] = max;
  r['Score %'] = pct;
  r['Grade'] = pct >= 90 ? 'A' : pct >= 75 ? 'B' : pct >= 60 ? 'C' : 'D';
  const tat = tatSummaryFromItems_((p.previousFollowup && p.previousFollowup.items) || []);
  r['TAT Green Count'] = tat.green;
  r['TAT Amber Count'] = tat.amber;
  r['TAT Red Count'] = tat.red;
  r['TAT Pending Count'] = tat.pending + tat.noTimeline;
  r['Data Quality Flags'] = buildQualityFlags_(defs, p.indicators || []);
  return r;
}

function computeTatCounts_(answers) {
  const today = new Date(); today.setHours(0,0,0,0);
  let green=0, red=0, pending=0;
  (answers || []).forEach(a => {
    const t = parseInputDate_(a.timelineDate);
    const c = parseInputDate_(a.completionDate);
    if (!t) return;
    t.setHours(0,0,0,0);
    if (c) {
      c.setHours(0,0,0,0);
      if (c <= t) green++; else red++;
    } else if (today > t) red++; else pending++;
  });
  return {green:green, red:red, pending:pending};
}

function buildQualityFlags_(defs, answers) {
  const flags = [];
  defs.forEach((d,i) => {
    const a = answers[i] || {};
    const score = Number(a.score);
    if (Number.isFinite(score) && score < Number(d.maxScore) && !safeString_(a.observation)) flags.push('I' + d.sr + ': low score without observation');
    normalizedSuggestions_(a).forEach((x,j)=>{ if (x.suggestion && !x.timelineDate) flags.push('I' + d.sr + ': suggestion '+(j+1)+' without timeline'); });
  });
  return flags.join(' | ').slice(0,1000);
}

function computeFollowupTat_(previousVisitDate, completionDate) {
  if (!previousVisitDate) return 'Not applicable';
  if (!completionDate) return 'Pending';
  return 'Completed';
}

function mapLinkFromLocation_(location) {
  const s = String(location || '');
  const m = s.match(/(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)/);
  return m ? 'https://maps.google.com/?q=' + m[1] + ',' + m[2] : '';
}

function findRowByHeaderValue_(sh, headers, header, value) {
  const idx = headers.indexOf(header);
  if (idx < 0 || sh.getLastRow() < 2 || !value) return null;
  const finder = sh.getRange(2,idx+1,sh.getLastRow()-1,1).createTextFinder(String(value)).matchEntireCell(true);
  const cell = finder.findNext();
  if (!cell) return null;
  const display = sh.getRange(cell.getRow(),1,1,headers.length).getDisplayValues()[0];
  const obj = {_row:cell.getRow()};
  headers.forEach((h,i) => { if (h) obj[h] = display[i]; });
  return obj;
}

function responseFromSubmissionRow_(row, type, duplicate) {
  const card = toCard_(row || {}, type);
  return {
    ok: true,
    duplicate: !!duplicate,
    id: card.id,
    pdf: card.pdf,
    excel: card.excel,
    pdfDownload: card.pdfDownload,
    excelDownload: card.excelDownload,
    folder: card.folder,
    score: card.score,
    grade: card.grade,
    quality: card.quality
  };
}


function setPhotoLinksCell_(sh, rowNum, headers, urls) {
  const idx = headers.indexOf('Photo URLs');
  if (idx < 0 || !rowNum) return;
  const clean = unique_((urls || []).map(x => safeString_(x)).filter(x => /^https?:\/\//i.test(x)));
  const cell = sh.getRange(rowNum, idx + 1);
  if (!clean.length) { cell.clearContent(); return; }
  const text = clean.join('\n');
  let builder = SpreadsheetApp.newRichTextValue().setText(text);
  let pos = 0;
  clean.forEach((url, i) => {
    const start = pos, end = pos + url.length;
    builder = builder.setLinkUrl(start, end, url);
    pos = end + (i < clean.length - 1 ? 1 : 0);
  });
  const style = SpreadsheetApp.newTextStyle().setForegroundColor('#1155cc').setUnderline(true).build();
  builder = builder.setTextStyle(0, text.length, style);
  cell.setRichTextValue(builder.build()).setWrap(true).setBackground('#eef6ff');
  cell.setNote(clean.map((u,i)=>(i+1)+'. '+u).join('\n'));
}

function repairPhotoHyperlinks_() {
  [MW.SHEETS.CFL_SUB, MW.SHEETS.SESSION_SUB].forEach(name => {
    const sh = getSS_().getSheetByName(name); if (!sh || sh.getLastRow() < 2) return;
    const headers = getHeaders_(sh), idx = headers.indexOf('Photo URLs'); if (idx < 0) return;
    for (let r=2; r<=sh.getLastRow(); r++) {
      const v = sh.getRange(r,idx+1).getDisplayValue();
      const urls = splitList_(v).filter(x => /^https?:\/\//i.test(x));
      if (urls.length) setPhotoLinksCell_(sh,r,headers,urls);
    }
  });
}

function getSubmissionForEdit(token, submissionId, reportType) {
  const user = requireSession_(token, [MW.ROLES.AM, MW.ROLES.CBO]);
  const type = String(reportType || '').toUpperCase();
  const req = activeEditRequest_(user, submissionId, type);
  if (!req) throw new Error('Approved edit permission is not active for this submission.');
  const sheetName = type === 'CFL' ? MW.SHEETS.CFL_SUB : MW.SHEETS.SESSION_SUB;
  const row = findSubmissionRow_(sheetName, submissionId);
  if (!row) throw new Error('Submission not found.');
  return {requestId:req['Request ID'], payload:submissionRowToPayload_(row,type)};
}

function submissionRowToPayload_(row, type) {
  const defs = type === 'CFL' ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
  const p = {
    reportType:type, cflName:row['CFL Name'], district:row.District, state:row.State, bccCode:row['BCC Code'], bankName:row['Bank Name'], phase:row.Phase,
    baseBlock:row['Base Block'], adjacentBlock:row['Adjacent Block'], visitDate:normalizeDateString_(row['Visit Date']), officerName:row['Officer Name'], officerDesignation:row['Officer Designation'],
    signoffName:row['Sign-off Name'], signoffDesignation:row['Sign-off Designation'], signoffDistrict:row['Sign-off District'], signoffZone:row['Sign-off Zone'],
    followupAction:row['Follow-up Action'], followupCompletionDate:normalizeDateString_(row['Follow-up Completion Date']), previousSubmissionId:row['Previous Submission ID'],
    photoUrls:splitList_(row['Photo URLs']), photoFileIds:splitList_(row['Photo URLs']).map(extractDriveId_).filter(Boolean)
  };
  if (type === 'SESSION') {
    p.consultantStaffDesignation=row['Consultant Staff Designation']; p.observer=row['Session Observer / Stakeholders']; p.sessionType=row['Type of Session']; p.venue=row['Venue of Session'];
  }
  p.indicators = defs.map(d => {
    const k = d.submissionPrefix || shortPrefix_(d.titleEn);
    const acts = actionRowsForSource_(row['Submission ID']).filter(x=>String(x['Indicator Sr'])===String(d.sr)).sort((x,y)=>String(x['Action ID']).localeCompare(String(y['Action ID'])));
    const a = {score:row[k+' - Score'], observation:row[k+' - Observation'], suggestion:row[k+' - Suggestion'], timelineDate:normalizeDateString_(row[k+' - Timeline Date']), completionDate:normalizeDateString_(row[k+' - Completion Date']), suggestions:acts.map(x=>({suggestion:x.Suggestion,timelineDate:normalizeDateString_(x['Timeline Date'])}))};
    if (!a.suggestions.length && a.suggestion) a.suggestions=[{suggestion:a.suggestion,timelineDate:a.timelineDate}];
    if (type === 'CFL' && k === 'Staffing') a.staffingDetails = row['Staffing - Staffing Details'] || '';
    return a;
  });
  return p;
}

function updateReport(token, submissionId, reportType, payload) {
  const user = requireSession_(token, [MW.ROLES.AM, MW.ROLES.CBO]);
  const type = String(reportType || '').toUpperCase();
  const req = activeEditRequest_(user, submissionId, type);
  if (!req) throw new Error('Edit permission missing or expired.');
  validatePayload_(user, type, payload);

  const sheetName = type === 'CFL' ? MW.SHEETS.CFL_SUB : MW.SHEETS.SESSION_SUB;
  const sh = getSS_().getSheetByName(sheetName);
  const old = findSubmissionRow_(sheetName, submissionId);
  if (!old) throw new Error('Submission not found.');
  const headers = getHeaders_(sh);
  const version = Math.max(1, Number(old['Report Version'] || 1)) + 1;
  const editCount = Number(old['Edit Count'] || 0) + 1;

  const previous = getPreviousFollowup_(type, payload.cflName, submissionId);
  payload.previousFollowup = mergeFollowupUpdatesForDisplay_(previous, payload.followupItems || [], payload.visitDate);
  payload.previousSubmissionId = payload.previousSubmissionId || previous.previousSubmissionId || '';
  payload.clientToken = old['Client Token'] || payload.clientToken || Utilities.getUuid();
  payload.photoUrls = splitList_(old['Photo URLs']);
  payload.photoFileIds = payload.photoUrls.map(extractDriveId_).filter(Boolean);
  payload.followupCompletionDate = payload.previousFollowup.completionDate || '';
  payload.followupTat = payload.previousFollowup.tat || 'Not applicable';

  const folder = getSubmissionFolder_(payload.visitDate, payload.officerName || user.name, payload.cflName, submissionId, user);
  let signatureInfo = saveSignature_(payload.signatureData, folder, submissionId, version);
  if (!signatureInfo && old['Digital Signature URL']) {
    try {
      const oldId = extractDriveId_(old['Digital Signature URL']);
      if (oldId) signatureInfo = {id:oldId,url:old['Digital Signature URL'],blob:DriveApp.getFileById(oldId).getBlob()};
    } catch (e) {}
  }
  payload.signatureUrl = signatureInfo ? signatureInfo.url : old['Digital Signature URL'];

  applyFollowupUpdates_(payload.followupItems || [], submissionId, user, payload.visitDate);
  upsertSubmissionActions_(type, submissionId, payload, user);
  const links = createReportFiles_(type, submissionId, payload, user, folder, version, signatureInfo ? signatureInfo.blob : null);
  const backup = saveSubmissionBackup_(folder, submissionId, type, payload, user, version);
  const record = buildRecord_(type, submissionId, user, payload, version, editCount);
  record.Timestamp = old._raw ? old._raw[headers.indexOf('Timestamp')] : old.Timestamp;
  record['Client Token'] = old['Client Token'];
  record['Photo URLs'] = old['Photo URLs'];
  record['Drive Folder Link'] = folder.getUrl();
  record['Data Backup Link'] = backup.url;
  record['Digital Signature URL'] = payload.signatureUrl || '';
  record['PDF Link'] = links.pdf || old['PDF Link'] || '';
  record['Excel Link'] = links.excel || old['Excel Link'] || '';
  record['Last Edited At'] = new Date();

  const current = sh.getRange(old._row,1,1,headers.length).getValues()[0];
  sh.getRange(old._row,1,1,headers.length).setValues([headers.map((h,i) => record[h] !== undefined ? record[h] : current[i])]);
  setPhotoLinksCell_(sh, old._row, headers, splitList_(record['Photo URLs'] || old['Photo URLs']));

  const er = getSS_().getSheetByName(MW.SHEETS.EDIT_REQUESTS);
  const erHeaders = getHeaders_(er);
  setByHeader_(er, req._row, erHeaders, 'Status', 'Edited');
  setByHeader_(er, req._row, erHeaders, 'Edited At', new Date());
  shareSubmissionAssets_(folder, [links.pdfId,links.excelId,backup.id].filter(Boolean).concat(signatureInfo ? [signatureInfo.id] : []), user);
  if (links.pdfId && links.excelId) notifyAdminsOfSubmission_(type, submissionId, payload, user, links, folder, version, 'UPDATE');
  try { clearDashboardCacheForUser_(user); } catch(e) {}
  audit_(user.id || user.email, 'EDIT_SUBMISSION', type, submissionId, req['Request ID']);
  const updatedResponse = responseFromSubmissionRow_(findSubmissionRow_(sheetName, submissionId), type, false);
  updatedResponse.reportError = links.reportError || '';
  return updatedResponse;
}

function activeEditRequest_(user, submissionId, type) {
  const uid = [user.id,user.email].map(x => safeString_(x).toLowerCase()).filter(Boolean);
  const now = new Date();
  return sheetObjects_(MW.SHEETS.EDIT_REQUESTS).filter(r => r['Submission ID'] === submissionId && String(r['Report Type']).toUpperCase() === type && String(r.Status) === 'Approved')
    .sort((a,b) => b._row-a._row)
    .find(r => uid.includes(safeString_(r['Requested By']).toLowerCase()) && (!r['Edit Expires At'] || new Date(r['Edit Expires At']) >= now)) || null;
}

function regenerateReport(token, submissionId, reportType) {
  const user = requireSession_(token, [MW.ROLES.ADMIN,MW.ROLES.AM,MW.ROLES.CBO]);
  if (user.role === MW.ROLES.ADMIN && normalizeAdminPermission_(user.permission) === MW.ADMIN_PERMS.VIEW) throw new Error('View-only admin cannot regenerate files.');
  const type = String(reportType || '').toUpperCase();
  const sheetName = type === 'CFL' ? MW.SHEETS.CFL_SUB : MW.SHEETS.SESSION_SUB;
  const row = findSubmissionRow_(sheetName, submissionId);
  if (!row) throw new Error('Submission not found.');
  if (user.role !== MW.ROLES.ADMIN && !visibleSubmissions_(sheetName,user).some(x => x['Submission ID'] === submissionId)) throw new Error('You cannot regenerate this submission.');
  const p = submissionRowToPayload_(row,type);
  p.previousFollowup = getPreviousFollowup_(type,p.cflName,submissionId);
  let folderUser = user;
  if (user.role === MW.ROLES.ADMIN) {
    try { folderUser = findOperationalUser_(row['Consultant ID'] || row['Logged-in Email'] || row['Officer Name']) || {id:row['Consultant ID'] || row['Logged-in Email'] || 'NO-ID', name:row['Consultant Name'] || row['Officer Name'] || 'Unknown Officer'}; } catch(e) {}
  }
  const folder = getSubmissionFolder_(p.visitDate,p.officerName,p.cflName,submissionId,folderUser);
  let signatureInfo = null;
  try { const sid=extractDriveId_(row['Digital Signature URL']); if(sid) signatureInfo={id:sid,url:row['Digital Signature URL'],blob:DriveApp.getFileById(sid).getBlob()}; } catch(e) {}
  const version = Math.max(1,Number(row['Report Version'] || 1));
  const links = createReportFiles_(type,submissionId,p,user,folder,version,signatureInfo ? signatureInfo.blob : null);
  const sh = getSS_().getSheetByName(sheetName); const headers=getHeaders_(sh);
  if (links.pdf) setByHeader_(sh,row._row,headers,'PDF Link',links.pdf); if (links.excel) setByHeader_(sh,row._row,headers,'Excel Link',links.excel); setByHeader_(sh,row._row,headers,'Drive Folder Link',folder.getUrl());
  shareSubmissionAssets_(folder,[links.pdfId,links.excelId].filter(Boolean).concat(signatureInfo?[signatureInfo.id]:[]),user);
  audit_(user.id || user.email,'REGENERATE_REPORT',type,submissionId,links.reportError || '');
  const regenerated = responseFromSubmissionRow_(findSubmissionRow_(sheetName,submissionId),type,false);
  regenerated.reportError = links.reportError || '';
  return regenerated;
}
