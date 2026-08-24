function getFormDefinition(token, reportType) {
  const user = requireSession_(token, [MW.ROLES.ADMIN, MW.ROLES.AM, MW.ROLES.CBO]);
  const type = String(reportType || '').toUpperCase();
  if (type !== 'CFL' && type !== 'SESSION') throw new Error('Invalid report type.');

  // Form definitions are the same for every user. Cache the heavy template/translation
  // read, then add only the signed-in user's defaults. This makes mobile form opening faster.
  const cache = CacheService.getScriptCache();
  const cacheKey = 'MW_FORM_DEF_' + MW.VERSION + '_' + type;
  let base = null;
  try {
    const cached = cache.get(cacheKey);
    if (cached) base = JSON.parse(cached);
  } catch (e) {}
  if (!base) {
    base = type === 'CFL' ? buildCflDefinition_() : buildSessionDefinition_();
    try { cache.put(cacheKey, JSON.stringify(base), 600); } catch (e) {}
  }
  const def = JSON.parse(JSON.stringify(base));
  def.userDefaults = {
    officerName: user.name || '',
    officerDesignation: user.designation || '',
    signoffName: user.name || '',
    signoffDesignation: user.designation || ''
  };
  return def;
}

function buildCflDefinition_() {
  const sh = getSS_().getSheetByName(MW.SHEETS.CFL_TEMPLATE);
  if (!sh) throw new Error(MW.SHEETS.CFL_TEMPLATE + ' sheet missing.');
  const vals = sh.getRange(15,1,17,2).getDisplayValues();
  const translations = ensureTranslationsForRows_('CFL', vals.map((r,i) => ({row:15+i, english:r[1]})));
  const prefixes = submissionScorePrefixes_('CFL');
  return {
    type: 'CFL',
    title: 'Checklist for CFL Visit - MoneyWise CFL Project',
    titleHi: 'सीएफएल विज़िट चेकलिस्ट - मनीवाइज़ सीएफएल प्रोजेक्ट',
    meta: commonMeta_(),
    indicators: vals.map((r,i) => indicatorDef_('CFL', 15+i, r[0], r[1], 1, translations['CFL-'+(15+i)] || r[1], prefixes[i]))
  };
}

function buildSessionDefinition_() {
  const sh = getSS_().getSheetByName(MW.SHEETS.SESSION_TEMPLATE);
  if (!sh) throw new Error(MW.SHEETS.SESSION_TEMPLATE + ' sheet missing.');
  const venue = sh.getRange('B17').getDisplayValue();
  const vals = sh.getRange(18,1,14,2).getDisplayValues();
  const translations = ensureTranslationsForRows_('SESSION', vals.map((r,i) => ({row:18+i, english:r[1]})));
  const prefixes = submissionScorePrefixes_('SESSION');
  return {
    type: 'SESSION',
    title: 'Checklist for Session Quality Monitoring - MoneyWise CFL Project',
    titleHi: 'सेशन क्वालिटी मॉनिटरिंग चेकलिस्ट - मनीवाइज़ सीएफएल प्रोजेक्ट',
    meta: commonMeta_().concat([
      {key:'consultantStaffDesignation', label:'Designation of consultant Staff', labelHi:'कंसल्टेंट स्टाफ का पद', type:'text'},
      // These three fields are intentionally kept together in one compact row on laptop/tablet.
      {key:'observer', label:'Session Observer / Stakeholders', labelHi:'सेशन ऑब्जर्वर / हितधारक', type:'text', uiClass:'session-inline'},
      {key:'venue', label:'Where is the session being observed (at Base/Adjacent Block, Village, at CFL, virtual, other - please specify)', labelHi:'सेशन कहाँ मॉनिटर किया जा रहा है (बेस/एडजेसेंट ब्लॉक, गांव, CFL, वर्चुअल, अन्य - कृपया बताएं)', type:'text', required:true, uiClass:'session-inline'},
      {key:'sessionType', label:'Type of Session', labelHi:'सेशन का प्रकार', type:'select', required:true, options:['AVT','FBT','AWC','Games','Calendar Roll','CFL'], uiClass:'session-inline'}
    ]),
    indicators: vals.map((r,i) => {
      const max = /session\s+observation/i.test(String(r[1] || '')) ? 3 : 1;
      return indicatorDef_('SESSION', 18+i, r[0], r[1], max, translations['SESSION-'+(18+i)] || r[1], prefixes[i]);
    })
  };
}

function commonMeta_() {
  return [
    {key:'cflName', label:'CFL Name', labelHi:'सीएफएल नाम', type:'select', required:true},
    {key:'district', label:'District', labelHi:'जिला', type:'text', readOnly:true},
    {key:'state', label:'State', labelHi:'राज्य', type:'text', readOnly:true},
    {key:'bccCode', label:'BCC Code', labelHi:'बीसीसी कोड', type:'text', readOnly:true},
    {key:'bankName', label:'Bank Name', labelHi:'बैंक का नाम', type:'text', readOnly:true},
    {key:'phase', label:'Phase', labelHi:'फेज़', type:'text', readOnly:true},
    {key:'baseBlock', label:'Base Block', labelHi:'बेस ब्लॉक', type:'text', readOnly:true},
    {key:'adjacentBlock', label:'Visit Block / Adjacent Block', labelHi:'विज़िट ब्लॉक / एडजेसेंट ब्लॉक', type:'select', required:true},
    {key:'visitDate', label:'Visit Date', labelHi:'विज़िट की तारीख', type:'date', required:true},
    {key:'officerName', label:'Officer Name', labelHi:'अधिकारी का नाम', type:'text', required:true, userLocked:true},
    {key:'officerDesignation', label:'Officer Designation', labelHi:'अधिकारी का पद', type:'text', required:true, userLocked:true}
  ];
}

function indicatorDef_(type, row, sr, title, maxScore, hindi, submissionPrefix) {
  const enLines = splitIndicatorLines_(title);
  const hiLines = splitIndicatorLines_(hindi);
  const options = maxScore === 3 ? [0,1,2,3] : [0,0.25,0.5,0.75,1];
  const out = {
    row: row,
    sr: Number(sr) || sr,
    titleEn: enLines[0] || safeString_(title),
    titleHi: hiLines[0] || safeString_(hindi),
    subEn: enLines.slice(1),
    subHi: hiLines.slice(1),
    fullEn: safeString_(title),
    fullHi: safeString_(hindi),
    maxScore: maxScore,
    scoreOptions: options,
    submissionPrefix: submissionPrefix || shortPrefix_(title)
  };
  if (type === 'CFL' && out.submissionPrefix === 'Staffing') {
    out.extraFields = [{key:'staffingDetails',label:'Staffing Details / Vacancy details',labelHi:'स्टाफिंग / रिक्त पद का विवरण',type:'textarea'}];
  }
  return out;
}

function splitIndicatorLines_(value) {
  return String(value || '').replace(/\r/g,'').split(/\n+/).map(x => x.trim()).filter(Boolean);
}

function translationSheetMap_() {
  const sh = getSS_().getSheetByName(MW.SHEETS.TRANSLATIONS) || ensureSheet_(MW.SHEETS.TRANSLATIONS, ['Key','Report Type','Template Row','English','Hindi','Updated At']);
  if (sh.getLastRow() < 2) return {};
  const data = sh.getRange(2,1,sh.getLastRow()-1,Math.max(6,sh.getLastColumn())).getDisplayValues();
  const out = {};
  data.forEach(r => { if (r[0] && r[4]) out[r[0]] = r[4]; });
  return out;
}

function ensureTranslationsForRows_(type, rows) {
  const sh = getSS_().getSheetByName(MW.SHEETS.TRANSLATIONS) || ensureSheet_(MW.SHEETS.TRANSLATIONS, ['Key','Report Type','Template Row','English','Hindi','Updated At']);
  const map = translationSheetMap_();
  const append = [];
  rows.forEach(item => {
    const key = type + '-' + item.row;
    if (map[key]) return;
    let hi = safeString_(item.english);
    try { hi = LanguageApp.translate(String(item.english || ''), 'en', 'hi'); } catch (e) {}
    map[key] = hi;
    append.push([key,type,item.row,item.english,hi,new Date()]);
  });
  if (append.length) sh.getRange(sh.getLastRow()+1,1,append.length,6).setValues(append);
  return map;
}

function initializeTranslations_() {
  const cfl = getSS_().getSheetByName(MW.SHEETS.CFL_TEMPLATE).getRange(15,1,17,2).getDisplayValues();
  const ses = getSS_().getSheetByName(MW.SHEETS.SESSION_TEMPLATE).getRange(18,1,14,2).getDisplayValues();
  const before = Object.keys(translationSheetMap_()).length;
  ensureTranslationsForRows_('CFL', cfl.map((r,i) => ({row:15+i,english:r[1]})));
  ensureTranslationsForRows_('SESSION', ses.map((r,i) => ({row:18+i,english:r[1]})));
  const after = Object.keys(translationSheetMap_()).length;
  return {created: Math.max(0, after-before), total: after, failed: 0};
}

function refreshMissingTranslations(token) {
  requireAdminWrite_(token);
  const r = initializeTranslations_();
  try { const c=CacheService.getScriptCache(); c.remove('MW_FORM_DEF_'+MW.VERSION+'_CFL'); c.remove('MW_FORM_DEF_'+MW.VERSION+'_SESSION'); } catch(e) {}
  audit_('ADMIN','REFRESH_TRANSLATIONS','SYSTEM','',JSON.stringify(r));
  return r;
}


function submissionScorePrefixes_(type) {
  const sheetName = type === 'CFL' ? MW.SHEETS.CFL_SUB : MW.SHEETS.SESSION_SUB;
  const sh = getSS_().getSheetByName(sheetName);
  if (!sh) return [];
  return getHeaders_(sh).filter(h => / - Score$/.test(h)).map(h => h.replace(/ - Score$/, ''));
}

function shortPrefix_(title) {
  return String(title || '').split(/[:\n]/)[0].trim().replace(/[?]+$/,'').trim();
}
