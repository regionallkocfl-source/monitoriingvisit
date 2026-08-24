function followupSheet_() {
  return ensureSheet_(MW.SHEETS.FOLLOWUP_ACTIONS, ['Action ID','Source Submission ID','Report Type','CFL Name','District','Visit Date','Indicator Sr','Indicator Title','Observation','Suggestion','Timeline Date','Status','Completion Date','Completed In Submission ID','Delay Days','TAT Days','TAT Status','Created By','Created At','Completed By','Updated At']);
}

function daysBetween_(fromDate, toDate) {
  const a = parseInputDate_(fromDate), b = parseInputDate_(toDate);
  if (!a || !b) return 0;
  a.setHours(12,0,0,0); b.setHours(12,0,0,0);
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

function tatInfo_(timelineDate, completionDate, asOfDate) {
  const timeline = parseInputDate_(timelineDate);
  if (!timeline) return {delayDays:'', tatDays:'', status:'NO_TIMELINE', label:'No timeline'};
  const completion = parseInputDate_(completionDate);
  const end = completion || parseInputDate_(asOfDate) || new Date();
  const delay = daysBetween_(timeline, end);
  if (completion && delay <= 0) return {delayDays:Math.max(0,delay), tatDays:0, status:'GREEN', label:'Green · On time'};
  if (!completion && delay <= 0) return {delayDays:0, tatDays:0, status:'PENDING', label:'Pending · Within timeline'};
  if (delay <= MW.TAT_GRACE_DAYS) return {delayDays:delay, tatDays:0, status:'AMBER', label:'Amber · '+delay+' day(s) after timeline'};
  const tatDays = delay - MW.TAT_GRACE_DAYS;
  return {delayDays:delay, tatDays:tatDays, status:'RED', label:'Red · TAT '+tatDays+' day(s)'};
}

function normalizedSuggestions_(answer) {
  const a = answer || {};
  let rows = Array.isArray(a.suggestions) ? a.suggestions : [];
  rows = rows.map(x => ({suggestion:safeString_(x && (x.suggestion || x.text),4000), timelineDate:formatDateYmd_(x && x.timelineDate)}))
    .filter(x => x.suggestion || x.timelineDate);
  if (!rows.length && (safeString_(a.suggestion) || a.timelineDate)) rows = [{suggestion:safeString_(a.suggestion,4000), timelineDate:formatDateYmd_(a.timelineDate)}];
  return rows;
}

function actionRowsForSource_(submissionId) {
  if (!submissionId) return [];
  return sheetObjects_(MW.SHEETS.FOLLOWUP_ACTIONS).filter(r => safeString_(r['Source Submission ID']) === safeString_(submissionId));
}

function upsertSubmissionActions_(type, submissionId, payload, user) {
  const sh = followupSheet_();
  const headers = getHeaders_(sh);
  const defs = type === 'CFL' ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
  const existing = actionRowsForSource_(submissionId);
  const now = new Date();
  const keepIds = new Set();
  defs.forEach((d,i) => {
    const a = (payload.indicators || [])[i] || {};
    const suggestions = normalizedSuggestions_(a);
    suggestions.forEach((s,j) => {
      if (!s.suggestion) return;
      const actionId = submissionId + '-I' + d.sr + '-A' + (j+1);
      keepIds.add(actionId);
      const old = existing.find(x => x['Action ID'] === actionId);
      const rec = {
        'Action ID': actionId,
        'Source Submission ID': submissionId,
        'Report Type': type,
        'CFL Name': safeString_(payload.cflName),
        'District': safeString_(payload.district),
        'Visit Date': parseInputDate_(payload.visitDate) || '',
        'Indicator Sr': d.sr,
        'Indicator Title': safeString_(d.titleEn,1000),
        'Observation': safeString_(a.observation,4000),
        'Suggestion': s.suggestion,
        'Timeline Date': parseInputDate_(s.timelineDate) || '',
        'Status': old && old.Status ? old.Status : 'OPEN',
        'Completion Date': old ? (parseInputDate_(old['Completion Date']) || '') : '',
        'Completed In Submission ID': old ? old['Completed In Submission ID'] : '',
        'Delay Days': old ? old['Delay Days'] : '',
        'TAT Days': old ? old['TAT Days'] : '',
        'TAT Status': old ? old['TAT Status'] : '',
        'Created By': old && old['Created By'] ? old['Created By'] : safeString_(user.id || user.email),
        'Created At': old && old['Created At'] ? old['Created At'] : now,
        'Completed By': old ? old['Completed By'] : '',
        'Updated At': now
      };
      if (old) {
        sh.getRange(old._row,1,1,headers.length).setValues([headers.map(h => rec[h] !== undefined ? rec[h] : '')]);
      } else {
        sh.appendRow(headers.map(h => rec[h] !== undefined ? rec[h] : ''));
      }
    });
  });
  // Do not delete completed history. Remove only unused OPEN action points created by the same source submission.
  existing.filter(x => !keepIds.has(x['Action ID']) && String(x.Status||'OPEN').toUpperCase() !== 'COMPLETED').sort((a,b)=>b._row-a._row).forEach(x => sh.deleteRow(x._row));
}

function applyFollowupUpdates_(items, currentSubmissionId, user, asOfDate) {
  if (!Array.isArray(items) || !items.length) return;
  const sh = followupSheet_();
  const headers = getHeaders_(sh);
  const rows = sheetObjects_(MW.SHEETS.FOLLOWUP_ACTIONS);
  items.forEach(item => {
    const id = safeString_(item && item.actionId);
    if (!id) return;
    const row = rows.find(r => r['Action ID'] === id);
    if (!row) return;
    const completion = parseInputDate_(item.completionDate) || '';
    if (completion) {
      const sourceVisit=parseInputDate_(row['Visit Date']), verifyDate=parseInputDate_(asOfDate)||new Date();
      if (sourceVisit && completion < sourceVisit) throw new Error('Completion Date cannot be before the source visit for action ' + id + '.');
      if (verifyDate && completion > verifyDate) throw new Error('Completion Date cannot be after the current Visit Date for action ' + id + '.');
    }
    const info = tatInfo_(row['Timeline Date'], completion, asOfDate);
    setByHeader_(sh,row._row,headers,'Completion Date',completion);
    setByHeader_(sh,row._row,headers,'Status',completion ? 'COMPLETED' : 'OPEN');
    setByHeader_(sh,row._row,headers,'Completed In Submission ID',completion ? currentSubmissionId : '');
    setByHeader_(sh,row._row,headers,'Delay Days',info.delayDays);
    setByHeader_(sh,row._row,headers,'TAT Days',info.tatDays);
    setByHeader_(sh,row._row,headers,'TAT Status',info.status);
    setByHeader_(sh,row._row,headers,'Completed By',completion ? safeString_(user.id || user.email) : '');
    setByHeader_(sh,row._row,headers,'Updated At',new Date());
    refreshLegacyIndicatorCompletion_(row['Source Submission ID'], row['Report Type'], row['Indicator Sr']);
  });
}

function refreshLegacyIndicatorCompletion_(sourceId, type, indicatorSr) {
  const all = actionRowsForSource_(sourceId).filter(x => String(x['Indicator Sr']) === String(indicatorSr));
  if (!all.length || all.some(x => String(x.Status).toUpperCase() !== 'COMPLETED' || !x['Completion Date'])) return;
  const dates = all.map(x => parseInputDate_(x['Completion Date'])).filter(Boolean).sort((a,b)=>a-b);
  if (!dates.length) return;
  const sheetName = String(type).toUpperCase()==='CFL' ? MW.SHEETS.CFL_SUB : MW.SHEETS.SESSION_SUB;
  const row = findSubmissionRow_(sheetName, sourceId);
  if (!row) return;
  const defs = String(type).toUpperCase()==='CFL' ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
  const d = defs.find(x => String(x.sr)===String(indicatorSr));
  if (!d) return;
  const sh = getSS_().getSheetByName(sheetName), headers=getHeaders_(sh);
  setByHeader_(sh,row._row,headers,(d.submissionPrefix||shortPrefix_(d.titleEn))+' - Completion Date',dates[dates.length-1]);
}


function ensureActionsForExistingSubmission_(type, previousRow) {
  if (!previousRow || !previousRow['Submission ID']) return;
  if (actionRowsForSource_(previousRow['Submission ID']).length) return;
  const sh=followupSheet_(), headers=getHeaders_(sh), now=new Date();
  legacyActionItemsFromSubmission_(type,previousRow).forEach(r=>{
    const rec={
      'Action ID':r['Action ID'],'Source Submission ID':previousRow['Submission ID'],'Report Type':type,
      'CFL Name':previousRow['CFL Name'],'District':previousRow.District,'Visit Date':parseInputDate_(previousRow['Visit Date'])||'',
      'Indicator Sr':r['Indicator Sr'],'Indicator Title':r['Indicator Title'],'Observation':r.Observation,'Suggestion':r.Suggestion,
      'Timeline Date':parseInputDate_(r['Timeline Date'])||'','Status':r['Completion Date']?'COMPLETED':'OPEN','Completion Date':parseInputDate_(r['Completion Date'])||'',
      'Completed In Submission ID':'','Delay Days':'','TAT Days':'','TAT Status':'','Created By':previousRow['Logged-in Email']||previousRow['Officer Name']||'LEGACY',
      'Created At':parseInputDate_(previousRow.Timestamp)||now,'Completed By':'','Updated At':now
    };
    const info=tatInfo_(rec['Timeline Date'],rec['Completion Date'],new Date()); rec['Delay Days']=info.delayDays;rec['TAT Days']=info.tatDays;rec['TAT Status']=info.status;
    sh.appendRow(headers.map(h=>rec[h]!==undefined?rec[h]:''));
  });
}

function mergeFollowupUpdatesForDisplay_(previous, updates, asOfDate) {
  const f=Object.assign({},previous||{}), byId={};
  (updates||[]).forEach(x=>{if(x&&x.actionId)byId[x.actionId]=x;});
  f.items=(f.items||[]).map(x=>{
    const u=byId[x.actionId]||{}, completion=u.completionDate!==undefined?u.completionDate:x.completionDate;
    const info=tatInfo_(x.timelineDate,completion,asOfDate);
    return Object.assign({},x,{completionDate:completion||'',status:completion?'COMPLETED':'OPEN',delayDays:info.delayDays,tatDays:info.tatDays,tatStatus:info.status,tatLabel:info.label});
  });
  f.observation=pointwiseText_(f.items,'observation','No major pending observation recorded.');
  f.suggestion=pointwiseText_(f.items,'suggestion','');
  f.recommended=f.suggestion;
  const dates=f.items.map(x=>x.completionDate).filter(Boolean).sort();
  f.completionDate=dates.length?dates[dates.length-1]:'';
  f.tat=tatSummaryLabel_(f.items);
  return f;
}

function followupItemsForPrevious_(type, previousRow, asOfDate) {
  if (!previousRow) return [];
  ensureActionsForExistingSubmission_(type, previousRow);
  let rows = actionRowsForSource_(previousRow['Submission ID']);
  // Show OPEN points first on the next visit so the field officer sees what must be verified.
  rows.sort((a,b) => {
    const ac = a['Completion Date'] ? 1 : 0, bc = b['Completion Date'] ? 1 : 0;
    if (ac !== bc) return ac - bc;
    return Number(a['Indicator Sr'] || 0) - Number(b['Indicator Sr'] || 0);
  });
  return rows.map((r,idx) => {
    const completion = normalizeDateString_(r['Completion Date']);
    const info = tatInfo_(r['Timeline Date'], completion, asOfDate);
    return {
      no: idx+1,
      actionId: safeString_(r['Action ID']) || (previousRow['Submission ID']+'-LEGACY-'+(idx+1)),
      sourceSubmissionId: safeString_(r['Source Submission ID']) || safeString_(previousRow['Submission ID']),
      sourceVisitDate: normalizeDateString_(r['Visit Date'] || previousRow['Visit Date']),
      indicatorSr: safeString_(r['Indicator Sr']),
      indicatorTitle: safeString_(r['Indicator Title']),
      observation: safeString_(r.Observation),
      suggestion: safeString_(r.Suggestion),
      timelineDate: normalizeDateString_(r['Timeline Date']),
      completionDate: completion,
      status: completion ? 'COMPLETED' : 'OPEN',
      delayDays: info.delayDays,
      tatDays: info.tatDays,
      tatStatus: info.status,
      tatLabel: info.label
    };
  });
}

function legacyActionItemsFromSubmission_(type, row) {
  const defs = type === 'CFL' ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
  const out=[];
  defs.forEach(d => {
    const p=d.submissionPrefix||shortPrefix_(d.titleEn), sug=safeString_(row[p+' - Suggestion']);
    if (!sug) return;
    out.push({
      'Action ID': row['Submission ID']+'-I'+d.sr+'-A1',
      'Indicator Sr': d.sr,
      'Indicator Title': d.titleEn,
      'Observation': row[p+' - Observation'],
      'Suggestion': sug,
      'Timeline Date': row[p+' - Timeline Date'],
      'Completion Date': row[p+' - Completion Date'],
      'Status': row[p+' - Completion Date'] ? 'COMPLETED' : 'OPEN'
    });
  });
  return out;
}

function pointwiseText_(items, key, emptyText) {
  const arr=(items||[]).filter(x=>safeString_(x[key]));
  return arr.length ? arr.map((x,i)=>(i+1)+'. [I'+(x.indicatorSr||'-')+'] '+safeString_(x[key])).join('\n') : (emptyText||'');
}

function tatSummaryFromItems_(items) {
  const c={green:0,amber:0,red:0,pending:0,noTimeline:0};
  (items||[]).forEach(x=>{
    const s=String(x.tatStatus||'').toUpperCase();
    if(s==='GREEN')c.green++; else if(s==='AMBER')c.amber++; else if(s==='RED')c.red++; else if(s==='NO_TIMELINE')c.noTimeline++; else c.pending++;
  });
  return c;
}

function tatSummaryLabel_(items) {
  const c=tatSummaryFromItems_(items);
  return 'Green '+c.green+' | Amber '+c.amber+' | Red '+c.red+' | Pending '+c.pending;
}

function visibleFollowupActions_(user) {
  const rows=sheetObjects_(MW.SHEETS.FOLLOWUP_ACTIONS);
  if(user.role===MW.ROLES.ADMIN)return rows;
  const ids=new Set(
    visibleSubmissions_(MW.SHEETS.CFL_SUB,user).map(r=>r['Submission ID'])
      .concat(visibleSubmissions_(MW.SHEETS.SESSION_SUB,user).map(r=>r['Submission ID']))
  );
  return rows.filter(r=>ids.has(r['Source Submission ID']));
}

function dashboardTat_(user) {
  const today=formatDateYmd_(new Date());
  const rows=visibleFollowupActions_(user);
  const items=rows.map(r=>{
    const info=tatInfo_(r['Timeline Date'],r['Completion Date'],today);
    return Object.assign({},r,{tatStatus:info.status,tatDays:info.tatDays,tatLabel:info.label});
  });
  const summary=tatSummaryFromItems_(items);
  const open=items.filter(x=>!x['Completion Date']);
  const due=open.filter(x=>['AMBER','RED'].includes(x.tatStatus)).sort((a,b)=>String(a['Timeline Date']).localeCompare(String(b['Timeline Date']))).slice(0,6).map(x=>({
    actionId:x['Action ID'],cfl:x['CFL Name'],indicator:x['Indicator Sr'],suggestion:x.Suggestion,timelineDate:normalizeDateString_(x['Timeline Date']),tatStatus:x.tatStatus,tatDays:x.tatDays,tatLabel:x.tatLabel
  }));
  return {summary:summary,openActions:open.length,dueActions:due};
}

function migrateLegacyFollowupActions_() {
  const sh=followupSheet_(), headers=getHeaders_(sh), existing=new Set(sheetObjects_(MW.SHEETS.FOLLOWUP_ACTIONS).map(r=>r['Source Submission ID']).filter(Boolean));
  const append=[], now=new Date();
  [[MW.SHEETS.CFL_SUB,'CFL'],[MW.SHEETS.SESSION_SUB,'SESSION']].forEach(pair=>{
    sheetObjects_(pair[0]).forEach(prev=>{
      const sid=safeString_(prev['Submission ID']); if(!sid||existing.has(sid))return;
      legacyActionItemsFromSubmission_(pair[1],prev).forEach(r=>{
        const completion=parseInputDate_(r['Completion Date'])||'', info=tatInfo_(r['Timeline Date'],completion,new Date());
        const rec={'Action ID':r['Action ID'],'Source Submission ID':sid,'Report Type':pair[1],'CFL Name':prev['CFL Name'],'District':prev.District,'Visit Date':parseInputDate_(prev['Visit Date'])||'','Indicator Sr':r['Indicator Sr'],'Indicator Title':r['Indicator Title'],'Observation':r.Observation,'Suggestion':r.Suggestion,'Timeline Date':parseInputDate_(r['Timeline Date'])||'','Status':completion?'COMPLETED':'OPEN','Completion Date':completion,'Completed In Submission ID':'','Delay Days':info.delayDays,'TAT Days':info.tatDays,'TAT Status':info.status,'Created By':prev['Logged-in Email']||prev['Officer Name']||'LEGACY','Created At':parseInputDate_(prev.Timestamp)||now,'Completed By':'','Updated At':now};
        append.push(headers.map(h=>rec[h]!==undefined?rec[h]:''));
      });
      existing.add(sid);
    });
  });
  if(append.length)sh.getRange(sh.getLastRow()+1,1,append.length,headers.length).setValues(append);
  return append.length;
}

// v3.7: do not lose an older unresolved action just because another visit happened later.
// Every next visit to the same CFL sees the complete OPEN backlog for that report type.
function followupBacklogForCfl_(type, cflName, asOfDate) {
  const t=safeString_(type).toUpperCase(), c=safeString_(cflName).toLowerCase();
  return sheetObjects_(MW.SHEETS.FOLLOWUP_ACTIONS)
    .filter(r=>safeString_(r['Report Type']).toUpperCase()===t && safeString_(r['CFL Name']).toLowerCase()===c && !r['Completion Date'])
    .map(r=>{
      const info=tatInfo_(r['Timeline Date'],r['Completion Date'],asOfDate);
      return {
        no:0, actionId:safeString_(r['Action ID']), sourceSubmissionId:safeString_(r['Source Submission ID']),
        sourceVisitDate:normalizeDateString_(r['Visit Date']), indicatorSr:safeString_(r['Indicator Sr']), indicatorTitle:safeString_(r['Indicator Title']),
        observation:safeString_(r.Observation), suggestion:safeString_(r.Suggestion), timelineDate:normalizeDateString_(r['Timeline Date']),
        completionDate:'', status:'OPEN', delayDays:info.delayDays, tatDays:info.tatDays, tatStatus:info.status, tatLabel:info.label
      };
    })
    .sort((a,b)=>{
      const at=a.timelineDate||'9999-12-31', bt=b.timelineDate||'9999-12-31';
      if(at!==bt)return at.localeCompare(bt);
      return (a.sourceVisitDate||'').localeCompare(b.sourceVisitDate||'');
    })
    .map((x,i)=>Object.assign(x,{no:i+1}));
}


// v3.8 FINAL: dedicated Follow-up tab. Shows all OPEN action points in the
// user's current CFL access, not only points created by that same login.
function getFollowupList(token) {
  const user = requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  const cflMaster = sheetObjects_(MW.SHEETS.CFL);
  const cflMeta = {};
  cflMaster.forEach(r => {
    const key = safeString_(r['CFL Name']).toLowerCase();
    if (key) cflMeta[key] = {bcc:normalizeCode_(r['BCC Code']), district:safeString_(r.District)};
  });

  const allowedCfls = new Set((user.allowedCfls || []).map(x => safeString_(x).toLowerCase()).filter(Boolean));
  const allowedDistricts = new Set((user.allowedDistricts || []).map(x => safeString_(x).toLowerCase()).filter(Boolean));
  let rows = sheetObjects_(MW.SHEETS.FOLLOWUP_ACTIONS).filter(r => !r['Completion Date'] && String(r.Status || 'OPEN').toUpperCase() !== 'COMPLETED');

  if (user.role === MW.ROLES.CBO) {
    // CBO access: Cbo Master District -> all CFLs of that District.
    rows = rows.filter(r => allowedDistricts.has(safeString_(r.District).toLowerCase()));
  } else if (user.role === MW.ROLES.AM) {
    // AM access: mapped CFLs from Employee Master; district is a fallback only
    // when no explicit CFL mapping exists.
    rows = rows.filter(r => allowedCfls.size
      ? allowedCfls.has(safeString_(r['CFL Name']).toLowerCase())
      : allowedDistricts.has(safeString_(r.District).toLowerCase()));
  }

  const rank = {RED:0, AMBER:1, NO_TIMELINE:2, PENDING:3, GREEN:4};
  const items = rows.map(r => {
    const info = tatInfo_(r['Timeline Date'], r['Completion Date'], new Date());
    const meta = cflMeta[safeString_(r['CFL Name']).toLowerCase()] || {};
    return {
      actionId:safeString_(r['Action ID']), sourceSubmissionId:safeString_(r['Source Submission ID']),
      type:safeString_(r['Report Type']), cfl:safeString_(r['CFL Name']), bcc:safeString_(meta.bcc),
      district:safeString_(r.District || meta.district), visitDate:normalizeDateString_(r['Visit Date']),
      indicatorSr:safeString_(r['Indicator Sr']), indicatorTitle:safeString_(r['Indicator Title']),
      observation:safeString_(r.Observation), suggestion:safeString_(r.Suggestion),
      timelineDate:normalizeDateString_(r['Timeline Date']), tatStatus:info.status, tatDays:info.tatDays, tatLabel:info.label
    };
  }).sort((a,b) => {
    const ra = rank[a.tatStatus] === undefined ? 9 : rank[a.tatStatus], rb = rank[b.tatStatus] === undefined ? 9 : rank[b.tatStatus];
    if (ra !== rb) return ra-rb;
    const at=a.timelineDate||'9999-12-31', bt=b.timelineDate||'9999-12-31';
    if (at !== bt) return at.localeCompare(bt);
    return String(a.cfl).localeCompare(String(b.cfl));
  });
  return {
    summary:{
      cfls:new Set(items.map(x=>x.cfl).filter(Boolean)).size,
      open:items.length,
      overdue:items.filter(x=>x.tatStatus==='AMBER'||x.tatStatus==='RED').length,
      noTimeline:items.filter(x=>x.tatStatus==='NO_TIMELINE').length
    },
    items:items
  };
}
