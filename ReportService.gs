function createReportFiles_(type, id, payload, user, folder, version, signatureBlob) {
  // v3.5: PDF and Excel are independent. One export failure must never hide a file
  // that was created successfully. The caller receives partial links + a clear error.
  const baseName = reportBaseName_(type, id, payload, version);
  let pdfInfo = null, xlsxInfo = null;
  const errors = [];
  try {
    pdfInfo = createDirectPdfReport_(type, id, payload, user, folder, version, baseName);
  } catch (e) {
    errors.push('PDF: ' + (e && e.message ? e.message : e));
    audit_(user && (user.id || user.email) || 'SYSTEM', 'PDF_GENERATION_FAILED', type, id, e.stack || e.message || String(e));
  }
  try {
    xlsxInfo = createExcelReport_(type, id, payload, user, folder, version, signatureBlob, baseName);
  } catch (e) {
    errors.push('Excel: ' + (e && e.message ? e.message : e));
    audit_(user && (user.id || user.email) || 'SYSTEM', 'EXCEL_GENERATION_FAILED', type, id, e.stack || e.message || String(e));
  }
  if (!pdfInfo && !xlsxInfo) throw new Error(errors.join(' | ') || 'PDF/Excel generation failed.');
  return {
    pdf: pdfInfo ? pdfInfo.file.getUrl() : '',
    excel: xlsxInfo ? xlsxInfo.file.getUrl() : '',
    pdfDownload: pdfInfo ? driveDownloadUrl_(pdfInfo.file.getId()) : '',
    excelDownload: xlsxInfo ? driveDownloadUrl_(xlsxInfo.file.getId()) : '',
    pdfId: pdfInfo ? pdfInfo.file.getId() : '',
    excelId: xlsxInfo ? xlsxInfo.file.getId() : '',
    reportError: errors.join(' | ')
  };
}

function createExcelReport_(type, id, payload, user, folder, version, signatureBlob, baseName) {
  const ss = getSS_();
  const templateName = type === 'CFL' ? MW.SHEETS.CFL_TEMPLATE : MW.SHEETS.SESSION_TEMPLATE;
  const src = ss.getSheetByName(templateName);
  if (!src) throw new Error(templateName + ' template sheet missing.');

  const out = SpreadsheetApp.create(id + '_v' + version + '_EXCEL_TEMP');
  const blank = out.getSheets()[0];
  const copied = src.copyTo(out).setName(templateName);

  // Reporting template tabs are hidden in the master workbook. copyTo() preserves that
  // hidden state. Google Sheets refuses to delete the only visible blank tab while the
  // copied report tab is hidden (the exact error seen in Audit Log on 23-Aug-2026).
  try { copied.showSheet(); } catch (e) {}
  try { out.setActiveSheet(copied); } catch (e) {}
  if (out.getSheets().length > 1) out.deleteSheet(blank);

  fillTemplate_(copied, type, payload, user, signatureBlob);
  SpreadsheetApp.flush();
  Utilities.sleep(450);

  try {
    const xlsxBlob = exportSpreadsheet_(out.getId(), copied.getSheetId(), 'xlsx').setName(baseName + '.xlsx');
    return {file: folder.createFile(xlsxBlob)};
  } finally {
    try { DriveApp.getFileById(out.getId()).setTrashed(true); } catch (e) {}
  }
}

function createDirectPdfReport_(type, id, p, user, folder, version, baseName) {
  const doc = DocumentApp.create(id + '_v' + version + '_PDF_TEMP');
  const body = doc.getBody();
  // v3.6: lock the direct report to A4 portrait and keep every table safely
  // inside the printable width. 18pt side margins leave ~559pt usable width.
  body.setPageWidth(595.28).setPageHeight(841.89);
  body.setMarginTop(18).setMarginBottom(18).setMarginLeft(18).setMarginRight(18);

  const title = type === 'CFL'
    ? 'Checklist for CFL Visit - MoneyWise CFL Project'
    : 'Checklist for Session Quality Monitoring - MoneyWise CFL Project';
  appendBlueHeading_(body, title, 16);

  const generated = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  const idLine = body.appendParagraph('Submission ID: ' + id + ' | Generated: ' + generated);
  idLine.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  styleParagraph_(idLine, 8, false, '#667085');

  appendMetaTable_(body, type, p);
  appendBlueHeading_(body, 'Checklist & Scoring', 14);
  appendScoringTable_(body, type, p);

  const score = reportScoreSummary_(type, p);
  const totalLine = body.appendParagraph('Total Score: ' + score.total + ' / ' + score.max + '   |   Score %: ' + score.pct + '%   |   Grade: ' + score.grade);
  styleParagraph_(totalLine, 10, true, '#172033');
  totalLine.setSpacingBefore(5).setSpacingAfter(6);

  appendPreviousFollowup_(body, p.previousFollowup || {});
  appendBlueHeading_(body, 'Photographs (Geotagged)', 13);
  body.appendPageBreak();
  // Prefer direct Drive file IDs from the current submission. Existing/old
  // submissions still work because appendPhotos_ also accepts Drive URLs.
  appendPhotos_(body, (p.photoFileIds && p.photoFileIds.length) ? p.photoFileIds : (p.photoUrls || []));
  appendSignoff_(body, p, user, id, generated);

  doc.saveAndClose();
  try {
    const pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF).setName(baseName + '.pdf');
    return {file: folder.createFile(pdfBlob)};
  } finally {
    try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch (e) {}
  }
}

function appendBlueHeading_(body, text, fontSize) {
  const t = body.appendTable([[String(text || '')]]);
  t.setBorderColor('#1b1b1b').setBorderWidth(0.7);
  const cell = t.getCell(0,0);
  cell.setBackgroundColor('#1266b1');
  const p = cell.getChild(0).asParagraph();
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  styleParagraph_(p, fontSize || 13, true, '#ffffff');
  p.setSpacingBefore(2).setSpacingAfter(2);
  body.appendParagraph('').setSpacingAfter(2).setSpacingBefore(0);
  return t;
}

function appendMetaTable_(body, type, p) {
  const rows = [
    ['CFL Name', p.cflName],
    ['District', p.district],
    ['State', p.state],
    ['BCC Code', p.bccCode],
    ['Bank Name', p.bankName],
    ['Phase', p.phase],
    ['Base Block', p.baseBlock],
    ['Adjacent Block', p.adjacentBlock],
    ['Visit Date', formatDateYmd_(p.visitDate)],
    ['Officer Name', p.officerName],
    ['Designation', p.officerDesignation]
  ];
  if (type === 'SESSION') {
    rows.push(['Designation of Consultant Staff', p.consultantStaffDesignation || '-']);
    rows.push(['Session Observer / Stakeholders', p.observer || '-']);
    rows.push(['Type of Session', p.sessionType || '-']);
    rows.push(['Venue of Session', p.venue || '-']);
  }
  const table = body.appendTable(rows.map(r => [safeString_(r[0]), safeString_(r[1]) || '-']));
  table.setBorderColor('#202020').setBorderWidth(0.8);
  // Total = 525pt, comfortably below the ~559pt printable A4 width.
  try { table.setColumnWidth(0, 145).setColumnWidth(1, 380); } catch(e) {}
  for (let i=0;i<rows.length;i++) {
    const left = table.getCell(i,0), right = table.getCell(i,1);
    left.setBackgroundColor('#eef0f3');
    styleCellText_(left, 9, true, '#252a31');
    styleCellText_(right, 9, false, '#252a31');
  }
  body.appendParagraph('').setSpacingAfter(2).setSpacingBefore(0);
}

function appendScoringTable_(body, type, p) {
  const defs = type === 'CFL' ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
  const data = [['Sr','Title / Full Guidance','Score','Observation','Suggestion\nRecommended','Time Line Date']];
  defs.forEach((d,i) => {
    const a = (p.indicators || [])[i] || {};
    const suggestions = normalizedSuggestions_(a);
    const suggestionText = suggestions.length
      ? suggestions.map((x,j) => (suggestions.length > 1 ? (j+1)+'. ' : '') + safeString_(x.suggestion)).join('\n')
      : '-';
    const timelineText = suggestions.length
      ? suggestions.map((x,j) => (suggestions.length > 1 ? (j+1)+'. ' : '') + (formatDateYmd_(x.timelineDate) || '-')).join('\n')
      : '-';
    const fullGuidance = [d.titleEn].concat(d.subEn || []).filter(Boolean).join('\n');
    data.push([
      String(d.sr),
      fullGuidance,
      (a.score === '' || a.score == null ? '-' : String(a.score)) + ' / ' + d.maxScore,
      safeString_(a.observation) || '-',
      suggestionText || '-',
      timelineText || '-'
    ]);
  });
  const table = body.appendTable(data);
  table.setBorderColor('#111111').setBorderWidth(0.8);
  // Total = 529pt. Previous v3.5 widths totalled 598pt and could spill
  // outside an A4 portrait page after Docs cell padding was applied.
  const widths = [24,170,55,105,105,70];
  widths.forEach((w,i)=>{ try { table.setColumnWidth(i,w); } catch(e) {} });
  for (let c=0;c<6;c++) {
    const cell = table.getCell(0,c);
    cell.setBackgroundColor('#1266b1');
    styleCellText_(cell, 7.5, true, '#ffffff');
    try { cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e) {}
  }
  for (let r=1;r<data.length;r++) {
    for (let c=0;c<6;c++) {
      const cell = table.getCell(r,c);
      styleCellText_(cell, c===1 ? 6.25 : 6.75, c===0 || c===2, '#1f2937');
      if (c===0 || c===2) try { cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e) {}
      if (c===1) {
        try {
          const txt = cell.getChild(0).asParagraph().editAsText();
          const firstLine = String(data[r][1] || '').split('\n')[0];
          if (firstLine) txt.setForegroundColor(0, Math.max(0,firstLine.length-1), '#1266b1').setBold(0, Math.max(0,firstLine.length-1), true);
        } catch(e) {}
      }
    }
  }
}

function appendPreviousFollowup_(body, f) {
  const items = f.items || [];
  const counts = tatSummaryFromItems_(items);
  const suffix = f.previousSubmissionId
    ? ('Previous: ' + f.previousSubmissionId + ' · TAT: Green ' + counts.green + ' Amber ' + counts.amber + ' Red ' + counts.red + ' Pending ' + (counts.pending + counts.noTimeline))
    : 'Pehli visit — koi previous data nahi · TAT: Green 0 Amber 0 Red 0 Pending 0';
  appendBlueHeading_(body, 'Follow-up on Previous Visit Observations & Stakeholder Feedback Status (' + suffix + ')', 11);
  const data = [['Sr','Indicator','Last Month Observation','Last Month Suggestion','Last Month Timeline','Completion Date','TAT']];
  if (items.length) {
    items.forEach((x,i)=>data.push([
      String(i+1),
      'I' + (x.indicatorSr || '-') + ' ' + (x.indicatorTitle || ''),
      x.observation || '-',
      x.suggestion || '-',
      formatDateYmd_(x.timelineDate) || '-',
      formatDateYmd_(x.completionDate) || 'Pending',
      x.tatLabel || x.tatStatus || 'Pending'
    ]));
  } else {
    data.push(['','', 'Pichli visit me koi Observation/Suggestion darj nahi tha.','','','','']);
  }
  const table = body.appendTable(data);
  table.setBorderColor('#222222').setBorderWidth(0.7);
  const widths=[22,80,100,100,65,65,55];
  widths.forEach((w,i)=>{ try { table.setColumnWidth(i,w); } catch(e) {} });
  for(let c=0;c<7;c++) {
    const cell=table.getCell(0,c); cell.setBackgroundColor('#eef0f3'); styleCellText_(cell,6.6,true,'#252a31');
    try { cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER); } catch(e) {}
  }
  for(let r=1;r<data.length;r++) for(let c=0;c<7;c++) styleCellText_(table.getCell(r,c),6.2,c===0,'#30343b');
  body.appendParagraph('').setSpacingAfter(1).setSpacingBefore(0);
}

function appendPhotos_(body, refs) {
  // v3.6: accept either Drive IDs or Drive URLs. Resolve and validate each
  // image before building the grid so a bad photo never prevents the rest.
  const ids = [];
  (refs || []).forEach(ref => {
    const id = extractDriveId_(ref) || safeString_(ref, 120);
    if (id && /^[A-Za-z0-9_-]{10,}$/.test(id) && ids.indexOf(id) < 0) ids.push(id);
  });
  const valid = ids.slice(0, MW.MAX_PHOTOS);
  if (!valid.length) {
    const p = body.appendParagraph('No photographs uploaded.');
    styleParagraph_(p, 9, false, '#667085');
    return;
  }

  for (let i = 0; i < valid.length; i += 2) {
    const table = body.appendTable([['','']]);
    table.setBorderColor('#d0d5dd').setBorderWidth(0.45);
    // Keep photo grid well inside A4 width even after cell padding.
    try { table.setColumnWidth(0, 250).setColumnWidth(1, 250); } catch(e) {}

    [0,1].forEach(col => {
      const fid = valid[i + col];
      const cell = table.getCell(0, col);
      if (!fid) {
        try { cell.setText(''); } catch(e) {}
        return;
      }
      try {
        const file = DriveApp.getFileById(fid);
        let blob = file.getBlob();
        const mime = String(blob.getContentType() || '').toLowerCase();
        if (mime.indexOf('image/') !== 0) throw new Error('Not an image file');

        // Uploaded field photos are JPEG, but the fallback makes older image
        // formats more tolerant when Google Docs can convert them.
        let img;
        try {
          img = cell.appendImage(blob);
        } catch (firstErr) {
          try {
            blob = blob.getAs('image/jpeg').setName('photo.jpg');
            img = cell.appendImage(blob);
          } catch (secondErr) {
            throw firstErr;
          }
        }

        const originalW = Math.max(1, img.getWidth());
        const originalH = Math.max(1, img.getHeight());
        const maxW = 220;
        const maxH = 165;
        const scale = Math.min(maxW / originalW, maxH / originalH, 1);
        img.setWidth(Math.max(80, Math.round(originalW * scale)));
        img.setHeight(Math.max(60, Math.round(originalH * scale)));

        try {
          const para = cell.getChild(cell.getNumChildren()-1).asParagraph();
          para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          para.setSpacingBefore(2).setSpacingAfter(2);
        } catch(e) {}
      } catch(e) {
        // Do not silently lose the evidence: identify the failed item in the
        // PDF and audit it, while allowing all other photos to render.
        cell.setText('Photo unavailable in PDF');
        styleCellText_(cell, 8, false, '#667085');
        try { audit_('SYSTEM', 'PDF_PHOTO_EMBED_WARNING', 'PHOTO', fid, e.message || String(e)); } catch(auditErr) {}
      }
    });
    body.appendParagraph('').setSpacingBefore(0).setSpacingAfter(2);
  }
}

function appendSignoff_(body, p, user, submissionId, generated) {
  appendBlueHeading_(body, 'Sign-off', 14);
  const rows = [
    ['Name', p.signoffName || user.name || ''],
    ['Designation', p.signoffDesignation || user.designation || ''],
    ['District', p.signoffDistrict || p.district || ''],
    ['Zone', p.signoffZone || ''],
    ['Digital Signature', (p.signoffName || user.name || '') + ' · digitally submitted on ' + generated + ' (Submission ID ' + submissionId + ')']
  ];
  const table = body.appendTable(rows.map(r => [r[0], safeString_(r[1]) || '']));
  table.setBorderColor('#111111').setBorderWidth(0.8);
  try { table.setColumnWidth(0,145).setColumnWidth(1,380); } catch(e) {}
  rows.forEach((_,r) => {
    table.getCell(r,0).setBackgroundColor('#eef0f3');
    styleCellText_(table.getCell(r,0),9,true,'#252a31');
    styleCellText_(table.getCell(r,1),r===4?8.5:9,false,'#252a31');
  });
}

function styleCellText_(cell, fontSize, bold, color) {
  try {
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);
    const p = cell.getChild(0).asParagraph();
    p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    styleParagraph_(p,fontSize,bold,color);
  } catch(e) {}
}

function styleParagraph_(p, fontSize, bold, color) {
  try {
    p.editAsText().setFontFamily('Arial').setFontSize(Number(fontSize || 9)).setBold(!!bold).setForegroundColor(color || '#172033');
  } catch(e) {}
  return p;
}

function reportScoreSummary_(type,p) {
  const defs = type === 'CFL' ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
  let total=0,max=0;
  defs.forEach((d,i)=>{ total += Number(((p.indicators||[])[i]||{}).score || 0); max += Number(d.maxScore || 1); });
  total = Math.round(total*100)/100;
  const pct = max ? Math.round(total/max*1000)/10 : 0;
  const grade = pct >= 90 ? 'A' : pct >= 75 ? 'B' : pct >= 60 ? 'C' : 'D';
  return {total:total,max:max,pct:pct,grade:grade};
}

function reportBaseName_(type, id, p, version) {
  const date = formatDateYmd_(p.visitDate) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return safeFileName_((type === 'CFL' ? 'CFL_Visit' : 'Session_Monitoring') + '_' + (p.cflName || 'CFL') + '_' + date + '_' + (p.officerName || 'Officer') + '_' + id + '_v' + version);
}

function fillTemplate_(sh, type, p, user, signatureBlob) {
  if (type === 'CFL') fillCflTemplate_(sh, p, user, signatureBlob);
  else fillSessionTemplate_(sh, p, user, signatureBlob);
}

function fillCflTemplate_(sh, p, user, signatureBlob) {
  const values = [
    p.cflName,p.district,p.state,p.bccCode,p.bankName,p.phase,p.baseBlock,p.adjacentBlock,
    parseInputDate_(p.visitDate),p.officerName,p.officerDesignation,user.name || p.officerName
  ];
  sh.getRange(2,3,12,1).setValues(values.map(v => [v || '']));
  sh.getRange('C10').setNumberFormat('dd-MMM-yyyy');

  const indicators = p.indicators || [];
  for (let i=0; i<17; i++) {
    const a = indicators[i] || {};
    const row = 15 + i;
    sh.getRange(row,3).setValue(a.score === '' || a.score == null ? '' : Number(a.score));
    sh.getRange(row,4).setValue(joinReportNotes_(a));
    sh.getRange(row,5).setValue(parseInputDate_(a.timelineDate) || '');
    if (a.timelineDate) sh.getRange(row,5).setNumberFormat('dd-MMM-yyyy');
  }

  const f = p.previousFollowup || {};
  sh.getRange('D33').setValue(f.observation || 'No previous data / कोई पिछला डेटा उपलब्ध नहीं');
  sh.getRange('D34').setValue(f.suggestion || '');
  sh.getRange('D35').setValue(p.followupAction || f.recommended || '');
  sh.getRange('D38').setValue(f.suggestion || '');
  sh.getRange('D39').setValue(parseInputDate_(f.visitDate) || f.visitDate || '');
  if (f.visitDate) sh.getRange('D39').setNumberFormat('dd-MMM-yyyy');
  sh.getRange('D42').setValue(followupCompletionText_(f));
  sh.getRange('D43').setValue(followupTatText_(f));

  sh.getRange('D44').setValue(p.signoffName || user.name || '');
  sh.getRange('D45').setValue(p.signoffDesignation || user.designation || '');
  sh.getRange('D46').setValue(p.signoffDistrict || p.district || '');
  sh.getRange('D47').setValue(p.signoffZone || '');
  sh.getRange('D48').setValue(signatureBlob ? '' : (p.signoffName || user.name || ''));
  if (signatureBlob) insertSignatureImage_(sh, signatureBlob, 4, 48, 150, 45);
}

function fillSessionTemplate_(sh, p, user, signatureBlob) {
  const values = [
    p.cflName,p.district,p.state,p.bccCode,p.bankName,p.phase,p.baseBlock,p.adjacentBlock,
    parseInputDate_(p.visitDate),p.officerName,p.officerDesignation,p.consultantStaffDesignation,p.observer,p.sessionType
  ];
  sh.getRange(2,4,14,1).setValues(values.map(v => [v || '']));
  sh.getRange('D10').setNumberFormat('dd-MMM-yyyy');
  sh.getRange('D17').setValue(p.venue || '');

  const indicators = p.indicators || [];
  for (let i=0; i<14; i++) {
    const a = indicators[i] || {};
    const row = 18 + i;
    sh.getRange(row,3).setValue(a.score === '' || a.score == null ? '' : Number(a.score));
    sh.getRange(row,4).setValue(joinReportNotes_(a));
  }

  const f = p.previousFollowup || {};
  sh.getRange('A34').setValue(p.signoffName || user.name || '');
  sh.getRange('A35').setValue(p.signoffDesignation || user.designation || '');
  sh.getRange('A36').setValue(p.signoffDistrict || p.district || '');
  sh.getRange('A37').setValue(p.signoffZone || '');
  sh.getRange('A38').setValue(signatureBlob ? '' : (p.signoffName || user.name || ''));
  if (signatureBlob) insertSignatureImage_(sh, signatureBlob, 1, 38, 135, 42);

  sh.getRange('D34').setValue(f.observation || 'No previous data / कोई पिछला डेटा उपलब्ध नहीं');
  sh.getRange('D35').setValue(f.suggestion || '');
  sh.getRange('D36').setValue(p.followupAction || f.recommended || '');
  sh.getRange('D39').setValue(f.suggestion || '');
  sh.getRange('D40').setValue(parseInputDate_(f.visitDate) || f.visitDate || '');
  if (f.visitDate) sh.getRange('D40').setNumberFormat('dd-MMM-yyyy');
  sh.getRange('D43').setValue(followupCompletionText_(f));
  sh.getRange('D44').setValue(followupTatText_(f));
}

function joinReportNotes_(a) {
  const lines = [];
  if (a.observation) lines.push('Observation: ' + a.observation);
  const suggestions = normalizedSuggestions_(a);
  suggestions.forEach((x,i) => lines.push((suggestions.length>1 ? 'Suggestion '+(i+1) : 'Suggestion / Recommended') + ': ' + x.suggestion + (x.timelineDate ? ' | Timeline: ' + formatDateYmd_(x.timelineDate) : '')));
  return lines.join('\n');
}

function followupCompletionText_(f) {
  const items=(f && f.items)||[];
  if(!items.length) return f && f.completionDate ? formatDateYmd_(f.completionDate) : 'Pending / Not applicable';
  return items.map((x,i)=>(i+1)+'. [I'+(x.indicatorSr||'-')+'] '+(x.completionDate?formatDateYmd_(x.completionDate):'Pending')).join('\n');
}

function followupTatText_(f) {
  const items=(f && f.items)||[];
  if(!items.length) return (f && f.tat) || 'Pending / Not applicable';
  return items.map((x,i)=>(i+1)+'. [I'+(x.indicatorSr||'-')+'] '+(x.tatLabel||x.tatStatus||'Pending')).join('\n');
}

function insertSignatureImage_(sh, blob, col, row, width, height) {
  try {
    const img = sh.insertImage(blob, col, row);
    img.setWidth(width).setHeight(height);
  } catch (e) {
    sh.getRange(row,col).setValue('Digitally signed');
  }
}

function exportSpreadsheet_(spreadsheetId, sheetId, format) {
  const token = ScriptApp.getOAuthToken();
  let url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=' + encodeURIComponent(format);
  if (format === 'pdf') {
    url += '&gid=' + sheetId + '&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false&fzr=false&top_margin=0.25&bottom_margin=0.25&left_margin=0.25&right_margin=0.25';
  }
  let lastCode = 0, lastText = '';
  for (let attempt = 1; attempt <= 3; attempt++) {
    const res = UrlFetchApp.fetch(url, {headers:{Authorization:'Bearer ' + token}, muteHttpExceptions:true, followRedirects:true});
    lastCode = res.getResponseCode();
    lastText = res.getContentText();
    if (lastCode < 300) {
      const mime = format === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      return res.getBlob().setContentType(mime);
    }
    if (![429,500,502,503,504].includes(lastCode)) break;
    Utilities.sleep(700 * attempt);
  }
  throw new Error('Report export failed (' + lastCode + '): ' + String(lastText || '').slice(0,300));
}
