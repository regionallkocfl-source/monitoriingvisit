/**
 * v3.7 Monitoring & reminder service
 * - Field last-visit summary
 * - Admin officer/CFL monitoring matrix
 * - Manual + optional daily consolidated follow-up reminder emails
 */

function reminderLogSheet_() {
  return ensureSheet_(MW.SHEETS.REMINDER_LOG, [
    'Reminder ID','Date','Mode','Recipient','Consultant ID','Consultant Name',
    'Action Count','Action IDs','Sent At','Sent By','Status','Error'
  ]);
}

function getFieldVisitSummary(token) {
  const user = requireSession_(token, [MW.ROLES.AM, MW.ROLES.CBO]);
  const all = visibleSubmissions_(MW.SHEETS.CFL_SUB, user).map(r => ({row:r,type:'CFL'}))
    .concat(visibleSubmissions_(MW.SHEETS.SESSION_SUB, user).map(r => ({row:r,type:'SESSION'})));
  all.sort((a,b) => submissionSortKey_(b.row).localeCompare(submissionSortKey_(a.row)));
  if (!all.length) return {exists:false,lastVisit:null,points:[],openTotal:0,overdueTotal:0};

  const latest = all[0], row = latest.row;
  const points = observationPointsForSubmission_(latest.type, row, new Date()).slice(0,12);
  const ownIds = new Set(all.map(x=>safeString_(x.row['Submission ID'])).filter(Boolean));
  const visibleActions = sheetObjects_(MW.SHEETS.FOLLOWUP_ACTIONS).filter(r=>ownIds.has(safeString_(r['Source Submission ID']))).map(r => {
    const info = tatInfo_(r['Timeline Date'], r['Completion Date'], new Date());
    return Object.assign({}, r, {__tat:info});
  });
  const open = visibleActions.filter(x => !x['Completion Date']);
  const overdue = open.filter(x => ['AMBER','RED'].includes(x.__tat.status));
  return {
    exists:true,
    lastVisit:{
      id:safeString_(row['Submission ID']), type:latest.type, cfl:safeString_(row['CFL Name']), district:safeString_(row.District),
      visitDate:normalizeDateString_(row['Visit Date']), officer:safeString_(row['Officer Name']), score:safeString_(row['Score %']),
      pdf:safeString_(row['PDF Link']), excel:safeString_(row['Excel Link'])
    },
    points:points,
    openTotal:open.length,
    overdueTotal:overdue.length
  };
}

function observationPointsForSubmission_(type, row, asOfDate) {
  if (!row) return [];
  const defs = type === 'CFL' ? buildCflDefinition_().indicators : buildSessionDefinition_().indicators;
  const actions = actionRowsForSource_(row['Submission ID']);
  const out=[];
  defs.forEach(d => {
    const p = d.submissionPrefix || shortPrefix_(d.titleEn);
    const observation = safeString_(row[p+' - Observation']);
    const acts = actions.filter(a => String(a['Indicator Sr']) === String(d.sr));
    if (acts.length) {
      acts.forEach(a => {
        const info=tatInfo_(a['Timeline Date'],a['Completion Date'],asOfDate);
        out.push({
          indicatorSr:String(d.sr), indicatorTitle:safeString_(d.titleEn), observation:observation || safeString_(a.Observation),
          suggestion:safeString_(a.Suggestion), timelineDate:normalizeDateString_(a['Timeline Date']),
          completionDate:normalizeDateString_(a['Completion Date']), tatStatus:info.status, tatDays:info.tatDays, tatLabel:info.label
        });
      });
    } else if (observation) {
      const info=tatInfo_(row[p+' - Timeline Date'],row[p+' - Completion Date'],asOfDate);
      out.push({
        indicatorSr:String(d.sr), indicatorTitle:safeString_(d.titleEn), observation:observation,
        suggestion:safeString_(row[p+' - Suggestion']), timelineDate:normalizeDateString_(row[p+' - Timeline Date']),
        completionDate:normalizeDateString_(row[p+' - Completion Date']), tatStatus:info.status, tatDays:info.tatDays, tatLabel:info.label
      });
    }
  });
  return out.sort((a,b) => {
    const ao=a.completionDate?1:0, bo=b.completionDate?1:0;
    if (ao!==bo) return ao-bo;
    return Number(a.indicatorSr||0)-Number(b.indicatorSr||0);
  });
}

function submissionSortKey_(row) {
  return normalizeDateString_(row && row['Visit Date']) + '|' + safeString_(row && row.Timestamp) + '|' + safeString_(row && row['Submission ID']);
}

function getAdminMonitoring(token) {
  requireSession_(token, [MW.ROLES.ADMIN]);
  return buildAdminMonitoring_();
}

function buildAdminMonitoring_() {
  const submissions = sheetObjects_(MW.SHEETS.CFL_SUB).map(r=>({row:r,type:'CFL'}))
    .concat(sheetObjects_(MW.SHEETS.SESSION_SUB).map(r=>({row:r,type:'SESSION'})));
  const subById={}; submissions.forEach(x=>{ if(x.row['Submission ID']) subById[x.row['Submission ID']]=x; });
  const actions = sheetObjects_(MW.SHEETS.FOLLOWUP_ACTIONS);
  const officerMap={}, cflMap={};
  const resolveIdentity = buildIdentityResolver_();

  submissions.forEach(x => {
    const id=resolveIdentity(x.row);
    const key=id.id || id.email || id.name || safeString_(x.row['Officer Name']) || 'UNKNOWN';
    if(!officerMap[key]) officerMap[key]={consultantId:id.id,name:id.name||safeString_(x.row['Officer Name']),email:id.email,districts:{},totalVisits:0,cflVisits:0,sessionVisits:0,lastVisitDate:'',lastVisitCfl:'',lastVisitType:'',actionPoints:0,completedActions:0,openActions:0,amber:0,red:0,noTimeline:0};
    const o=officerMap[key];
    if(id.email && !o.email)o.email=id.email; if(id.name && !o.name)o.name=id.name; if(id.id && !o.consultantId)o.consultantId=id.id;
    if(x.row.District)o.districts[safeString_(x.row.District)]=true;
    o.totalVisits++; if(x.type==='CFL')o.cflVisits++; else o.sessionVisits++;
    const d=normalizeDateString_(x.row['Visit Date']);
    if(d && d>=o.lastVisitDate){o.lastVisitDate=d;o.lastVisitCfl=safeString_(x.row['CFL Name']);o.lastVisitType=x.type;}

    const ckey=safeString_(x.row['CFL Name']).toLowerCase();
    if(ckey){
      if(!cflMap[ckey])cflMap[ckey]={cfl:safeString_(x.row['CFL Name']),district:safeString_(x.row.District),totalVisits:0,cflVisits:0,sessionVisits:0,lastVisitDate:'',lastOfficer:'',openActions:0,completedActions:0,amber:0,red:0,noTimeline:0,nextTimeline:''};
      const c=cflMap[ckey]; c.totalVisits++; if(x.type==='CFL')c.cflVisits++;else c.sessionVisits++;
      if(d && d>=c.lastVisitDate){c.lastVisitDate=d;c.lastOfficer=id.name||safeString_(x.row['Officer Name']);}
    }
  });

  actions.forEach(a => {
    const src=subById[a['Source Submission ID']];
    if(!src)return;
    const id=resolveIdentity(src.row);
    const key=id.id || id.email || id.name || safeString_(src.row['Officer Name']) || 'UNKNOWN';
    const o=officerMap[key]; if(!o)return;
    const info=tatInfo_(a['Timeline Date'],a['Completion Date'],new Date());
    o.actionPoints++;
    if(a['Completion Date'])o.completedActions++; else {o.openActions++; if(info.status==='AMBER')o.amber++; if(info.status==='RED')o.red++; if(info.status==='NO_TIMELINE')o.noTimeline++;}
    const c=cflMap[safeString_(a['CFL Name']).toLowerCase()];
    if(c){
      if(a['Completion Date'])c.completedActions++; else {
        c.openActions++; if(info.status==='AMBER')c.amber++; if(info.status==='RED')c.red++; if(info.status==='NO_TIMELINE')c.noTimeline++;
        const tl=normalizeDateString_(a['Timeline Date']); if(tl && (!c.nextTimeline || tl<c.nextTimeline))c.nextTimeline=tl;
      }
    }
  });

  const officers=Object.values(officerMap).map(o=>{
    o.district=Object.keys(o.districts).join(', '); delete o.districts;
    o.completionPct=o.actionPoints?Math.round(o.completedActions/o.actionPoints*1000)/10:100;
    o.canRemind=!!o.email && (o.amber+o.red)>0;
    return o;
  }).sort((a,b)=>(b.red-a.red)||(b.amber-a.amber)||(b.openActions-a.openActions)||(b.totalVisits-a.totalVisits));

  const cfls=Object.values(cflMap).sort((a,b)=>(b.red-a.red)||(b.amber-a.amber)||(b.openActions-a.openActions)||String(a.cfl).localeCompare(String(b.cfl)));
  const reminderStatus=reminderAutomationStatus_();
  return {
    summary:{
      totalVisits:submissions.length,
      officers:officers.length,
      actionPoints:officers.reduce((n,x)=>n+x.actionPoints,0),
      completedActions:officers.reduce((n,x)=>n+x.completedActions,0),
      openActions:officers.reduce((n,x)=>n+x.openActions,0),
      amber:officers.reduce((n,x)=>n+x.amber,0),
      red:officers.reduce((n,x)=>n+x.red,0),
      noTimeline:officers.reduce((n,x)=>n+x.noTimeline,0)
    },
    officers:officers,
    cfls:cfls,
    reminderAutomation:reminderStatus
  };
}

function buildIdentityResolver_() {
  const cbos=sheetObjects_(MW.SHEETS.CBO), emps=sheetObjects_(MW.SHEETS.EMPLOYEE);
  return function(row) {
    const cid=safeString_(row['Consultant ID']);
    const name=safeString_(row['Consultant Name'] || row['Officer Name']);
    const rawEmail=[row['Logged-in Email'],row['Logged-in Google Account']].map(v=>safeString_(v)).find(x=>x && x.includes('@')) || '';
    const cq=cid.toLowerCase(), nq=name.toLowerCase();
    let m=cbos.find(r=>cq && safeString_(r['Consultant ID']).toLowerCase()===cq);
    if(!m && nq)m=cbos.find(r=>safeString_(r['Consultant Name']).toLowerCase()===nq);
    if(m)return {id:safeString_(m['Consultant ID'])||cid,name:safeString_(m['Consultant Name'])||name,email:safeString_(m['Official Email ID'])||rawEmail,role:'CBO'};
    m=emps.find(r=>cq && safeString_(r['Consultant Id']||r['Consultant ID']).toLowerCase()===cq);
    if(!m && nq)m=emps.find(r=>safeString_(r.Name||r['Consultant Name']).toLowerCase()===nq);
    if(m)return {id:safeString_(m['Consultant Id']||m['Consultant ID'])||cid,name:safeString_(m.Name||m['Consultant Name'])||name,email:safeString_(m.Email||m['Email Id']||m['Official Email ID'])||rawEmail,role:'AM'};
    return {id:cid,name:name,email:rawEmail,role:''};
  };
}

function resolveOfficerIdentityFromSubmission_(row) {
  return buildIdentityResolver_()(row);
}

function overdueActionsForConsultant_(consultantId) {
  const q=safeString_(consultantId).toLowerCase();
  const submissions=sheetObjects_(MW.SHEETS.CFL_SUB).map(r=>({row:r,type:'CFL'})).concat(sheetObjects_(MW.SHEETS.SESSION_SUB).map(r=>({row:r,type:'SESSION'})));
  const resolveIdentity=buildIdentityResolver_();
  const owned={}; let identity=null;
  submissions.forEach(x=>{
    const id=resolveIdentity(x.row);
    if([id.id,id.email,id.name].map(v=>safeString_(v).toLowerCase()).includes(q)) {owned[x.row['Submission ID']]=x; identity=identity||id;}
  });
  const actions=sheetObjects_(MW.SHEETS.FOLLOWUP_ACTIONS).filter(a=>owned[a['Source Submission ID']] && !a['Completion Date']).map(a=>{
    const info=tatInfo_(a['Timeline Date'],a['Completion Date'],new Date());
    const src=owned[a['Source Submission ID']];
    return {action:a,info:info,source:src};
  }).filter(x=>['AMBER','RED'].includes(x.info.status));
  return {identity:identity,items:actions};
}

function sendOfficerFollowupReminder(token, consultantId) {
  const admin=requireAdminWrite_(token);
  const bundle=overdueActionsForConsultant_(consultantId);
  if(!bundle.identity || !bundle.identity.email) throw new Error('Officer email not found in Employee Master / Cbo Master.');
  if(!bundle.items.length) return {ok:true,sent:false,message:'No overdue follow-up points for this officer.'};
  sendFollowupReminderEmail_(bundle.identity,bundle.items,'MANUAL',safeString_(admin.email||admin.id));
  return {ok:true,sent:true,count:bundle.items.length,email:bundle.identity.email};
}

function reminderAutomationStatus_() {
  const triggers=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='runDailyFollowupReminders');
  return {enabled:triggers.length>0,count:triggers.length,hour:8};
}

function setDailyFollowupReminder(token, enabled) {
  const admin=requireFullAdmin_(token);
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='runDailyFollowupReminders').forEach(t=>ScriptApp.deleteTrigger(t));
  if(enabled) ScriptApp.newTrigger('runDailyFollowupReminders').timeBased().everyDays(1).atHour(8).create();
  audit_(admin.email||admin.id,enabled?'REMINDER_TRIGGER_ENABLED':'REMINDER_TRIGGER_DISABLED','SYSTEM','','Daily follow-up reminder');
  return reminderAutomationStatus_();
}

function runDailyFollowupReminders() {
  const monitoring=buildAdminMonitoring_();
  let sent=0, skipped=0, failed=0, totalActions=0;
  monitoring.officers.filter(o=>o.email && (o.amber+o.red)>0).forEach(o=>{
    try{
      if(autoReminderSentToday_(o.email)){skipped++;return;}
      const bundle=overdueActionsForConsultant_(o.consultantId||o.email||o.name);
      if(!bundle.items.length){skipped++;return;}
      sendFollowupReminderEmail_(bundle.identity||o,bundle.items,'AUTO','SYSTEM');
      sent++; totalActions+=bundle.items.length;
    }catch(e){failed++;audit_('SYSTEM','FOLLOWUP_REMINDER_FAILED','USER',o.consultantId||o.email,e.message);}
  });
  try { sendAdminReminderSummary_(monitoring); } catch(e) { audit_('SYSTEM','ADMIN_REMINDER_SUMMARY_FAILED','SYSTEM','',e.message); }
  return {ok:true,sent:sent,skipped:skipped,failed:failed,actionPoints:totalActions};
}

function autoReminderSentToday_(email) {
  const day=formatDateYmd_(new Date());
  return sheetObjects_(MW.SHEETS.REMINDER_LOG).some(r=>safeString_(r.Date)===day && safeString_(r.Mode)==='AUTO' && safeString_(r.Recipient).toLowerCase()===safeString_(email).toLowerCase() && safeString_(r.Status)==='SENT');
}

function sendFollowupReminderEmail_(identity, items, mode, sentBy) {
  const email=safeString_(identity.email).toLowerCase();
  if(!email || !email.includes('@')) throw new Error('Valid recipient email missing.');
  const rows=items.map(x=>{
    const a=x.action, src=x.source&&x.source.row;
    return '<tr>'+
      '<td style="border:1px solid #d0d5dd;padding:6px">'+escapeHtmlMail_(a['CFL Name'])+'</td>'+
      '<td style="border:1px solid #d0d5dd;padding:6px">'+escapeHtmlMail_(normalizeDateString_(a['Visit Date']))+'</td>'+
      '<td style="border:1px solid #d0d5dd;padding:6px">I'+escapeHtmlMail_(a['Indicator Sr'])+' '+escapeHtmlMail_(a['Indicator Title'])+'</td>'+
      '<td style="border:1px solid #d0d5dd;padding:6px">'+escapeHtmlMail_(a.Observation||'-')+'</td>'+
      '<td style="border:1px solid #d0d5dd;padding:6px">'+escapeHtmlMail_(a.Suggestion||'-')+'</td>'+
      '<td style="border:1px solid #d0d5dd;padding:6px">'+escapeHtmlMail_(normalizeDateString_(a['Timeline Date'])||'-')+'</td>'+
      '<td style="border:1px solid #d0d5dd;padding:6px"><b>'+escapeHtmlMail_(x.info.label)+'</b></td></tr>';
  }).join('');
  const subject='[MoneyWise] Follow-up Reminder - '+items.length+' pending point(s) - '+safeString_(identity.name||identity.id);
  const html='<div style="font-family:Arial,sans-serif;color:#172033;max-width:950px;margin:auto">'+
    '<div style="background:#1266b1;color:#fff;padding:14px 18px"><h2 style="margin:0">MoneyWise Monitoring Follow-up Reminder</h2></div>'+
    '<div style="padding:16px;border:1px solid #d0d5dd;border-top:0"><p>Dear <b>'+escapeHtmlMail_(identity.name||identity.id)+'</b>,</p>'+
    '<p>The following previous-visit action points have crossed their Timeline Date and are still open. Please verify them during the next visit and enter the Completion Date only after physical/record verification.</p>'+
    '<table style="border-collapse:collapse;width:100%;font-size:12px"><thead><tr style="background:#eef4fb"><th style="border:1px solid #d0d5dd;padding:6px">CFL</th><th style="border:1px solid #d0d5dd;padding:6px">Source Visit</th><th style="border:1px solid #d0d5dd;padding:6px">Indicator</th><th style="border:1px solid #d0d5dd;padding:6px">Observation</th><th style="border:1px solid #d0d5dd;padding:6px">Suggestion / Action</th><th style="border:1px solid #d0d5dd;padding:6px">Timeline</th><th style="border:1px solid #d0d5dd;padding:6px">TAT</th></tr></thead><tbody>'+rows+'</tbody></table>'+
    '<p style="font-size:12px;color:#667085">TAT rule: 1-7 days after Timeline Date = Amber grace; more than 7 days = Red/TAT.</p></div></div>';
  let status='SENT',err='';
  try{MailApp.sendEmail({to:email,subject:subject,htmlBody:html,body:'MoneyWise follow-up reminder: '+items.length+' overdue action point(s).'});}catch(e){status='FAILED';err=e.message;throw e;}finally{
    logReminder_(mode,email,identity,items,sentBy,status,err);
  }
}

function logReminder_(mode,email,identity,items,sentBy,status,error) {
  const sh=reminderLogSheet_(), headers=getHeaders_(sh);
  const rec={'Reminder ID':'REM-'+Utilities.getUuid().replace(/-/g,'').slice(0,10).toUpperCase(),'Date':formatDateYmd_(new Date()),'Mode':mode,'Recipient':email,'Consultant ID':identity.id||identity.consultantId||'','Consultant Name':identity.name||'','Action Count':items.length,'Action IDs':items.map(x=>x.action['Action ID']).join(', '),'Sent At':new Date(),'Sent By':sentBy,'Status':status,'Error':error||''};
  sh.appendRow(headers.map(h=>rec[h]!==undefined?rec[h]:''));
}

function sendAdminReminderSummary_(monitoring) {
  const admins=getActiveAdminEmails_(); if(!admins.length)return;
  const s=monitoring.summary||{};
  if(!(s.amber||s.red))return;
  const top=monitoring.officers.filter(o=>(o.amber+o.red)>0).slice(0,20);
  const rows=top.map(o=>'<tr><td style="border:1px solid #ddd;padding:6px">'+escapeHtmlMail_(o.consultantId||'-')+'</td><td style="border:1px solid #ddd;padding:6px">'+escapeHtmlMail_(o.name||'-')+'</td><td style="border:1px solid #ddd;padding:6px">'+escapeHtmlMail_(o.district||'-')+'</td><td style="border:1px solid #ddd;padding:6px">'+o.totalVisits+'</td><td style="border:1px solid #ddd;padding:6px">'+o.openActions+'</td><td style="border:1px solid #ddd;padding:6px">'+o.amber+'</td><td style="border:1px solid #ddd;padding:6px">'+o.red+'</td></tr>').join('');
  MailApp.sendEmail({to:admins[0],bcc:admins.slice(1).join(','),subject:'[MoneyWise] Daily Follow-up Summary - '+s.red+' Red / '+s.amber+' Amber',htmlBody:'<div style="font-family:Arial"><h2>MoneyWise Daily Follow-up Summary</h2><p><b>Total visits:</b> '+s.totalVisits+' | <b>Open actions:</b> '+s.openActions+' | <b>Amber:</b> '+s.amber+' | <b>Red:</b> '+s.red+'</p><table style="border-collapse:collapse;width:100%"><tr><th>ID</th><th>Officer</th><th>District</th><th>Visits</th><th>Open</th><th>Amber</th><th>Red</th></tr>'+rows+'</table></div>',body:'MoneyWise daily follow-up summary: '+s.red+' Red, '+s.amber+' Amber.'});
}
