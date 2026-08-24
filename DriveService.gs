function getSubmissionFolder_(visitDate, officerName, cflName, submissionId, user) {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
  const day = formatDateYmd_(visitDate) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const consultantId = safeString_(user && user.id) || 'NO-ID';
  const consultantName = safeString_(user && user.name) || safeString_(officerName) || 'Unknown Officer';
  const consultantFolder = childFolder_(root, consultantId + ' - ' + consultantName);
  const dayFolder = childFolder_(consultantFolder, day);
  const cflFolder = childFolder_(dayFolder, safeFileName_(cflName || 'Unknown CFL'));
  return childFolder_(cflFolder, safeFileName_(submissionId));
}

function getStagingFolder_(clientToken) {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
  const stagingRoot = childFolder_(root, MW.STAGING_FOLDER_NAME);
  return childFolder_(stagingRoot, safeFileName_(clientToken));
}

function childFolder_(parent, name) {
  const safe = safeFileName_(name || 'Unknown');
  const it = parent.getFoldersByName(safe);
  return it.hasNext() ? it.next() : parent.createFolder(safe);
}

function safeFileName_(s) {
  return String(s || 'Unknown').replace(/[\\/:*?"<>|#%]/g,'-').replace(/\s+/g,' ').trim().slice(0,100) || 'Unknown';
}

function uploadPhoto(token, clientToken, photo) {
  const user = requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  clientToken = safeString_(clientToken, 100);
  if (!clientToken) throw new Error('Draft token missing. Reload the form.');
  photo = photo || {};
  const mime = safeString_(photo.mime, 80).toLowerCase();
  if (!mime.startsWith('image/')) throw new Error('Only image files are allowed.');
  let data = String(photo.data || '');
  data = data.replace(/^data:[^;]+;base64,/, '');
  const bytes = Utilities.base64Decode(data);
  if (!bytes.length) throw new Error('Empty image.');
  if (bytes.length > MW.MAX_PHOTO_BYTES) throw new Error('Photo is too large after compression.');
  const folder = getStagingFolder_(clientToken);
  const name = safeFileName_(photo.name || ('Photo_' + new Date().getTime() + '.jpg'));
  const blob = Utilities.newBlob(bytes, mime, name);
  const file = folder.createFile(blob);
  const desc = {uploadedBy:user.id || user.email, uploadedAt:new Date().toISOString(), location:photo.location || ''};
  try { file.setDescription(JSON.stringify(desc)); } catch (e) {}
  return {id:file.getId(), url:file.getUrl(), name:file.getName()};
}

function finalizeStagedPhotos_(clientToken, finalFolder, user) {
  if (!clientToken) return [];
  let staging;
  try { staging = getStagingFolder_(clientToken); } catch (e) { return []; }
  const files = staging.getFiles();
  const out = [];
  let count = 0;
  while (files.hasNext() && count < MW.MAX_PHOTOS) {
    const f = files.next();
    try { f.moveTo(finalFolder); } catch (e) {}
    out.push({id:f.getId(), url:f.getUrl(), name:f.getName()});
    count++;
  }
  try { if (!staging.getFiles().hasNext() && !staging.getFolders().hasNext()) staging.setTrashed(true); } catch (e) {}
  return out;
}

function saveSignature_(signatureData, folder, submissionId, version) {
  if (!signatureData) return null;
  let data = String(signatureData).replace(/^data:image\/png;base64,/, '');
  if (!data) return null;
  const bytes = Utilities.base64Decode(data);
  if (!bytes.length || bytes.length > 900000) return null;
  const name = safeFileName_(submissionId + '_v' + version + '_signature.png');
  const blob = Utilities.newBlob(bytes, 'image/png', name);
  const file = folder.createFile(blob);
  return {id:file.getId(), url:file.getUrl(), blob:blob};
}

function getActiveAdminEmails_() {
  return unique_(sheetObjects_(MW.SHEETS.ADMINS)
    .filter(isActiveRecord_)
    .map(r => safeString_(r.Email).toLowerCase())
    .filter(e => e && e.includes('@')));
}

// Share access without sending Google Drive's separate "Item shared with you" emails.
// One MoneyWise report email is sent separately by notifyAdminsOfSubmission_().
function shareSubmissionAssets_(folder, fileIds, user) {
  const emails = [];
  if (user && user.email && String(user.email).includes('@')) emails.push(user.email);
  if (user && user.activeGoogleEmail && String(user.activeGoogleEmail).includes('@')) emails.push(user.activeGoogleEmail);
  getActiveAdminEmails_().forEach(e => emails.push(e));
  unique_(emails).forEach(email => silentAddViewer_(folder.getId(), email));
}

function silentAddViewer_(fileOrFolderId, email) {
  if (!fileOrFolderId || !email) return false;
  try {
    const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileOrFolderId) + '/permissions?sendNotificationEmail=false&supportsAllDrives=true';
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
      payload: JSON.stringify({type:'user', role:'reader', emailAddress:email}),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    return code >= 200 && code < 300;
  } catch (e) { return false; }
}

function notifyAdminsOfSubmission_(type, submissionId, payload, user, links, folder, version, action) {
  const admins = getActiveAdminEmails_();
  if (!admins.length) return {sent:false, reason:'No active admin email configured'};
  if (!links || !links.pdfId || !links.excelId) return {sent:false, reason:'Report files unavailable'};

  const label = type === MW.REPORTS.CFL ? 'CFL Monitoring' : 'Session Monitoring';
  const verb = action === 'UPDATE' ? 'Updated' : 'Submitted';
  const subject = '[MoneyWise] ' + label + ' ' + verb + ' - ' + submissionId + ' - ' + safeString_(payload.cflName);
  const scoreInfo = scoreSummaryFromPayload_(type, payload);
  const submittedBy = safeString_(payload.officerName || (user && user.name) || '');
  const visitDate = safeString_(payload.visitDate);
  const district = safeString_(payload.district);
  const html = [
    '<div style="font-family:Arial,sans-serif;max-width:680px;margin:auto;color:#172033">',
    '<div style="background:#1261a6;color:#fff;padding:16px 18px;border-radius:10px 10px 0 0"><h2 style="margin:0;font-size:20px">MoneyWise CFL Monitoring</h2></div>',
    '<div style="border:1px solid #d9e2ec;border-top:0;padding:18px;border-radius:0 0 10px 10px">',
    '<p><b>' + escapeHtmlMail_(label) + '</b> report has been ' + verb.toLowerCase() + '.</p>',
    '<table style="border-collapse:collapse;width:100%;font-size:14px">',
    mailRow_('Submission ID', submissionId),
    mailRow_('CFL', payload.cflName),
    mailRow_('District', district),
    mailRow_('Visit Date', visitDate),
    mailRow_('Officer', submittedBy),
    mailRow_('Consultant ID', user && user.id),
    mailRow_('Score', scoreInfo),
    mailRow_('Report Version', String(version || 1)),
    '</table>',
    '<p style="margin-top:16px"><b>Attached:</b> PDF report + Excel report.</p>',
    '<p style="margin-top:12px">',
    '<a href="' + escapeHtmlMail_(links.pdf) + '" style="display:inline-block;background:#1261a6;color:#fff;text-decoration:none;padding:10px 14px;border-radius:7px;margin-right:8px">View PDF</a>',
    '<a href="' + escapeHtmlMail_(links.excel) + '" style="display:inline-block;background:#eef4fb;color:#1261a6;text-decoration:none;padding:10px 14px;border-radius:7px;margin-right:8px">View Excel</a>',
    '<a href="' + escapeHtmlMail_(folder.getUrl()) + '" style="display:inline-block;color:#1261a6;text-decoration:none;padding:10px 4px">Drive Folder</a>',
    '</p>',
    '<p style="color:#667085;font-size:12px;margin-bottom:0">Automatic MoneyWise report email. Drive access is shared silently, so separate Google Drive sharing emails are suppressed.</p>',
    '</div></div>'
  ].join('');

  try {
    const pdfBlob = DriveApp.getFileById(links.pdfId).getBlob();
    const excelBlob = DriveApp.getFileById(links.excelId).getBlob();
    MailApp.sendEmail({
      to: admins[0],
      bcc: admins.slice(1).join(','),
      subject: subject,
      htmlBody: html,
      body: label + ' ' + verb + '\nSubmission ID: ' + submissionId + '\nCFL: ' + safeString_(payload.cflName) + '\nPDF and Excel are attached.',
      attachments: [pdfBlob, excelBlob]
    });
    return {sent:true, recipients:admins.length};
  } catch (e) {
    audit_('SYSTEM','ADMIN_EMAIL_WARNING',type,submissionId,e.message);
    return {sent:false, reason:e.message};
  }
}

function scoreSummaryFromPayload_(type, payload) {
  try {
    const defs = type === MW.REPORTS.CFL ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
    let total=0,max=0;
    defs.forEach((d,i)=>{ total += Number(((payload.indicators||[])[i]||{}).score||0); max += Number(d.maxScore||1); });
    const pct = max ? Math.round(total/max*1000)/10 : 0;
    const grade = pct >= 90 ? 'A' : pct >= 75 ? 'B' : pct >= 60 ? 'C' : 'D';
    return total + ' / ' + max + ' (' + pct + '%) Grade ' + grade;
  } catch (e) { return ''; }
}

function mailRow_(label, value) {
  return '<tr><td style="border-bottom:1px solid #edf1f5;padding:7px 5px;color:#667085;width:155px"><b>' + escapeHtmlMail_(label) + '</b></td><td style="border-bottom:1px solid #edf1f5;padding:7px 5px">' + escapeHtmlMail_(value || '-') + '</td></tr>';
}

function escapeHtmlMail_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function driveDownloadUrl_(urlOrId) {
  const id = extractDriveId_(urlOrId);
  return id ? 'https://drive.google.com/uc?export=download&id=' + id : '';
}

function extractDriveId_(value) {
  const s = String(value || '');
  const m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/) || s.match(/\/folders\/([a-zA-Z0-9_-]{10,})/) || s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/) || s.match(/^([a-zA-Z0-9_-]{20,})$/);
  return m ? m[1] : '';
}

function deleteStagedPhoto(token, clientToken, fileId) {
  requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  const folder = getStagingFolder_(safeString_(clientToken,100));
  const targetId = safeString_(fileId,100);
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getId() === targetId) { f.setTrashed(true); return true; }
  }
  return false;
}

function saveSubmissionBackup_(folder, submissionId, type, payload, user, version) {
  const clean = JSON.parse(JSON.stringify(payload || {}));
  delete clean.signatureData;
  clean.submissionId = submissionId;
  clean.reportType = type;
  clean.reportVersion = version;
  clean.savedAt = new Date().toISOString();
  clean.savedBy = {id:user.id || '', name:user.name || '', email:user.email || '', role:user.role || ''};
  const name = safeFileName_(submissionId + '_v' + version + '_data.json');
  const blob = Utilities.newBlob(JSON.stringify(clean,null,2), 'application/json', name);
  const file = folder.createFile(blob);
  return {id:file.getId(), url:file.getUrl()};
}
