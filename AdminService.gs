function getAdminOverview(token) {
  const user = requireSession_(token, [MW.ROLES.ADMIN]);
  const dash = getDashboardForUser_(user);
  const edits = sheetObjects_(MW.SHEETS.EDIT_REQUESTS).sort((a,b) => b._row-a._row);
  return {
    user: publicUser_(user),
    dashboard: dash,
    pendingEdits: edits.filter(r => r.Status === 'Pending').slice(0,100),
    masterCounts: {
      employeeMappings: sheetObjects_(MW.SHEETS.EMPLOYEE).length,
      cboRows: sheetObjects_(MW.SHEETS.CBO).length,
      cfls: sheetObjects_(MW.SHEETS.CFL).length,
      admins: sheetObjects_(MW.SHEETS.ADMINS).filter(isActiveRecord_).length
    }
  };
}

function getAdminSubmissions(token, filters) {
  requireSession_(token, [MW.ROLES.ADMIN]);
  filters = filters || {};
  let rows = sheetObjects_(MW.SHEETS.CFL_SUB).map(r => toCard_(r,'CFL')).concat(sheetObjects_(MW.SHEETS.SESSION_SUB).map(r => toCard_(r,'SESSION')));
  const q = safeString_(filters.search).toLowerCase();
  if (q) rows = rows.filter(r => [r.id,r.cfl,r.district,r.officer,r.grade,r.type].some(v => safeString_(v).toLowerCase().includes(q)));
  if (filters.type) rows = rows.filter(r => r.type === String(filters.type).toUpperCase());
  if (filters.district) rows = rows.filter(r => safeString_(r.district).toLowerCase() === safeString_(filters.district).toLowerCase());
  if (filters.cfl) rows = rows.filter(r => safeString_(r.cfl).toLowerCase() === safeString_(filters.cfl).toLowerCase());
  if (filters.officer) rows = rows.filter(r => safeString_(r.officer).toLowerCase().includes(safeString_(filters.officer).toLowerCase()));
  if (filters.grade) rows = rows.filter(r => safeString_(r.grade).toUpperCase() === safeString_(filters.grade).toUpperCase());
  if (filters.from) rows = rows.filter(r => normalizeDateString_(r.visitDate || r.timestamp) >= String(filters.from));
  if (filters.to) rows = rows.filter(r => normalizeDateString_(r.visitDate || r.timestamp) <= String(filters.to));
  return rows.sort(cardDateSortDesc_).slice(0,1000);
}

function getAdminEditRequests(token, status) {
  requireSession_(token, [MW.ROLES.ADMIN]);
  let rows = sheetObjects_(MW.SHEETS.EDIT_REQUESTS).sort((a,b) => b._row-a._row);
  if (status) rows = rows.filter(r => safeString_(r.Status).toLowerCase() === safeString_(status).toLowerCase());
  return rows.slice(0,1000);
}

function getAdminMasters(token) {
  requireSession_(token, [MW.ROLES.ADMIN]);
  return {
    employees: sheetObjects_(MW.SHEETS.EMPLOYEE),
    cbos: sheetObjects_(MW.SHEETS.CBO),
    cfls: sheetObjects_(MW.SHEETS.CFL),
    admins: sheetObjects_(MW.SHEETS.ADMINS),
    pins: sheetObjects_(MW.SHEETS.PINS),
    overrides: sheetObjects_(MW.SHEETS.BLOCK_OVERRIDES)
  };
}

function requestEdit(token, submissionId, reportType, reason) {
  const user = requireSession_(token, [MW.ROLES.AM, MW.ROLES.CBO]);
  const type = String(reportType || '').toUpperCase();
  const sheetName = type === 'CFL' ? MW.SHEETS.CFL_SUB : MW.SHEETS.SESSION_SUB;
  const row = findSubmissionRow_(sheetName, submissionId);
  if (!row) throw new Error('Submission not found.');
  if (!visibleSubmissions_(sheetName,user).some(r => r['Submission ID'] === submissionId)) throw new Error('You can request edit only for your own submission.');
  reason = safeString_(reason,2000);
  if (reason.length < 3) throw new Error('Please enter the reason for edit.');

  const existing = sheetObjects_(MW.SHEETS.EDIT_REQUESTS).filter(r => r['Submission ID'] === submissionId && safeString_(r['Requested By']).toLowerCase() === safeString_(user.id || user.email).toLowerCase() && ['Pending','Approved'].includes(r.Status)).sort((a,b)=>b._row-a._row)[0];
  if (existing) return {ok:true,id:existing['Request ID'],status:existing.Status,message:'An edit request is already active.'};

  const sh = getSS_().getSheetByName(MW.SHEETS.EDIT_REQUESTS);
  const headers = getHeaders_(sh);
  const id = 'REQ-' + Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase();
  const rec = {'Request ID':id,'Submission ID':submissionId,'Report Type':type,'Requested By':user.id || user.email,'Reason':reason,'Status':'Pending','Requested At':new Date()};
  sh.appendRow(headers.map(h => rec[h] !== undefined ? rec[h] : ''));
  audit_(user.id || user.email,'EDIT_REQUEST',type,submissionId,reason);
  return {ok:true,id:id,status:'Pending'};
}

function decideEditRequest(token, requestId, decision, note, hours) {
  const admin = requireAdminWrite_(token);
  const sh = getSS_().getSheetByName(MW.SHEETS.EDIT_REQUESTS);
  const rows = sheetObjects_(MW.SHEETS.EDIT_REQUESTS);
  const r = rows.find(x => x['Request ID'] === requestId);
  if (!r) throw new Error('Request not found.');
  if (r.Status !== 'Pending') throw new Error('This request is already ' + r.Status + '.');
  const approve = String(decision).toUpperCase() === 'APPROVE';
  const headers = getHeaders_(sh);
  const status = approve ? 'Approved' : 'Rejected';
  const validHours = Math.min(168, Math.max(1, Number(hours || MW.EDIT_WINDOW_HOURS)));
  setByHeader_(sh,r._row,headers,'Status',status);
  setByHeader_(sh,r._row,headers,'Approved By',admin.email || admin.id);
  setByHeader_(sh,r._row,headers,'Approved At',new Date());
  setByHeader_(sh,r._row,headers,'Decision Note',safeString_(note,2000));
  if (approve) setByHeader_(sh,r._row,headers,'Edit Expires At',new Date(Date.now()+validHours*3600000));
  audit_(admin.id || admin.email,'EDIT_' + status.toUpperCase(),r['Report Type'],r['Submission ID'],requestId + ' ' + safeString_(note,500));
  return {ok:true,status:status};
}

function saveAdminUser(token, data) {
  const admin = requireFullAdmin_(token);
  data = data || {};
  const email = safeString_(data.email,200).toLowerCase();
  if (!email || !email.includes('@')) throw new Error('Valid admin email required.');
  const permission = normalizeAdminPermission_(data.permission);
  const adminPin = safeString_(data.adminPin,20);
  if (adminPin && !/^\d{4,8}$/.test(adminPin)) throw new Error('Admin PIN should be 4 to 8 digits.');
  const sh = getSS_().getSheetByName(MW.SHEETS.ADMINS);
  const headers = getHeaders_(sh);
  const rows = sheetObjects_(MW.SHEETS.ADMINS);
  const existing = rows.find(r => safeString_(r.Email).toLowerCase() === email);
  const rec = {'Email':email,'Permission':permission,'Active':data.active === false ? false : true,'Display Name':safeString_(data.displayName,200),'Updated At':new Date()};
  if (adminPin) rec['Admin PIN'] = adminPin;
  if (existing) {
    Object.keys(rec).forEach(k => setByHeader_(sh,existing._row,headers,k,rec[k]));
  } else {
    rec['Added At'] = new Date();
    sh.appendRow(headers.map(h => rec[h] !== undefined ? rec[h] : ''));
  }
  audit_(admin.id || admin.email, existing ? 'UPDATE_ADMIN':'ADD_ADMIN','USER',email,permission);
  return true;
}

function saveEmployeeMapping(token, data) {
  const admin = requireAdminWrite_(token);
  data = data || {};
  const id = safeString_(data.consultantId,100);
  const name = safeString_(data.name,200);
  const cfl = safeString_(data.cflName,200);
  if (!id || !name || !cfl) throw new Error('Consultant ID, Name and CFL Name are required.');
  const sh = getSS_().getSheetByName(MW.SHEETS.EMPLOYEE);
  const headers = getHeaders_(sh);
  if (data.replaceExisting) {
    const rows = sheetObjects_(MW.SHEETS.EMPLOYEE).filter(r => safeString_(r['Consultant Id'] || r['Consultant ID']) === id).sort((a,b)=>b._row-a._row);
    rows.forEach(r => sh.deleteRow(r._row));
  }
  const rec = {
    'Consultant Id':id,'Consultant ID':id,'Name':name,'Designation':safeString_(data.designation,200) || 'Area Manager',
    'State':safeString_(data.state,100),'District':safeString_(data.district,100),'Zone':safeString_(data.zone,100),'CFL Name':cfl,
    'Email Id':safeString_(data.email,200),'Email':safeString_(data.email,200),'Login Exempt':data.loginExempt ? 1 : ''
  };
  const currentRows = sheetObjects_(MW.SHEETS.EMPLOYEE);
  const found = currentRows.find(r => safeString_(r['Consultant Id'] || r['Consultant ID']) === id && safeString_(r['CFL Name']).toLowerCase() === cfl.toLowerCase());
  if (found) sh.getRange(found._row,1,1,headers.length).setValues([headers.map(h => rec[h] !== undefined ? rec[h] : '')]);
  else sh.appendRow(headers.map(h => rec[h] !== undefined ? rec[h] : ''));
  audit_(admin.id || admin.email, data.replaceExisting ? 'TRANSFER_EMPLOYEE':'SAVE_EMPLOYEE','USER',id,JSON.stringify(rec));
  return true;
}

function saveCbo(token, data) {
  const admin = requireAdminWrite_(token);
  data = data || {};
  const id = safeString_(data.consultantId,100);
  if (!id || !safeString_(data.name)) throw new Error('Consultant ID and Name are required.');
  const sh = getSS_().getSheetByName(MW.SHEETS.CBO);
  const headers = getHeaders_(sh);
  const rows = sheetObjects_(MW.SHEETS.CBO);
  const found = rows.find(r => safeString_(r['Consultant ID']) === id);
  const rec = {
    'Consultant ID':id,'Consultant Name':safeString_(data.name,200),'Consultant Designation':safeString_(data.designation,250) || 'Consultant (Capacity Building Officer)',
    'District':safeString_(data.district,100),'Official Email ID':safeString_(data.email,200),'Mobile No.':safeString_(data.mobile,30)
  };
  const firstConsultantIdIndex = headers.indexOf('Consultant ID');
  const makeRow = current => headers.map((h,i) => {
    if (h === 'Consultant ID' && i !== firstConsultantIdIndex) return current ? current[i] : '';
    return rec[h] !== undefined ? rec[h] : (current ? current[i] : '');
  });
  if (found) {
    const current = sh.getRange(found._row,1,1,headers.length).getValues()[0];
    sh.getRange(found._row,1,1,headers.length).setValues([makeRow(current)]);
  } else sh.appendRow(makeRow(null));
  audit_(admin.id || admin.email,found?'UPDATE_CBO':'ADD_CBO','USER',id,JSON.stringify(rec));
  return true;
}

function saveUserPin(token, data) {
  const admin = requireAdminWrite_(token);
  data = data || {};
  const officer = safeString_(data.officerName,200);
  const consultantId = safeString_(data.consultantId,100);
  const email = safeString_(data.email,200);
  const pin = safeString_(data.pin,20);
  if (!officer && !consultantId && !email) throw new Error('Officer Name, Consultant ID or Email required.');
  if (pin && !/^\d{4,8}$/.test(pin)) throw new Error('PIN should be 4 to 8 digits, or blank for no PIN.');
  const sh = getSS_().getSheetByName(MW.SHEETS.PINS);
  const headers = getHeaders_(sh);
  const pinHeader = headers.find(h => h.toLowerCase().includes('pin'));
  const rows = sheetObjects_(MW.SHEETS.PINS);
  const found = rows.find(r => (consultantId && safeString_(r['Consultant ID'] || r['Consultant Id']) === consultantId) || (email && safeString_(r.Email).toLowerCase() === email.toLowerCase()) || (officer && safeString_(r['Officer Name']).toLowerCase() === officer.toLowerCase()));
  const rec = {'Officer Name':officer,'Consultant ID':consultantId,'Email':email,'Active':data.active === false ? false : true,'Updated At':new Date()};
  rec[pinHeader] = pin;
  if (found) {
    Object.keys(rec).forEach(k => setByHeader_(sh,found._row,headers,k,rec[k]));
  } else sh.appendRow(headers.map(h => rec[h] !== undefined ? rec[h] : ''));
  audit_(admin.id || admin.email,'SAVE_PIN','USER',consultantId || email || officer,'PIN configured');
  return true;
}

function saveCflMaster(token, data) {
  const admin = requireAdminWrite_(token);
  data = data || {};
  const cfl = safeString_(data.cflName,200);
  const bcc = normalizeCode_(data.bccCode);
  if (!cfl || !bcc) throw new Error('CFL Name and BCC Code are required.');
  const sh = getSS_().getSheetByName(MW.SHEETS.CFL);
  const headers = getHeaders_(sh);
  const rows = sheetObjects_(MW.SHEETS.CFL);
  const found = rows.find(r => normalizeCode_(r['BCC Code']) === bcc || safeString_(r['CFL Name']).toLowerCase() === cfl.toLowerCase());
  const rec = {
    'CFL Name':cfl,'BCC Code':bcc,'Phase':safeString_(data.phase,100),'State':safeString_(data.state,100),'District':safeString_(data.district,100),
    'Base Block':safeString_(data.baseBlock,150),'Adjacent Block 1':safeString_(data.adjacent1,150),'Adjacent Block 2':safeString_(data.adjacent2,150),'Bank Name':safeString_(data.bankName,200)
  };
  if (found) sh.getRange(found._row,1,1,headers.length).setValues([headers.map(h => rec[h] !== undefined ? rec[h] : '')]);
  else sh.appendRow(headers.map(h => rec[h] !== undefined ? rec[h] : ''));
  saveBlockOverrideInternal_(data);
  audit_(admin.id || admin.email,found?'UPDATE_CFL':'ADD_CFL','CFL',bcc,JSON.stringify(rec));
  return true;
}

function saveBlockOverrideInternal_(data) {
  const sh = getSS_().getSheetByName(MW.SHEETS.BLOCK_OVERRIDES);
  if (!sh) return;
  const headers = getHeaders_(sh);
  const rows = sheetObjects_(MW.SHEETS.BLOCK_OVERRIDES);
  const bcc = normalizeCode_(data.bccCode);
  const cfl = safeString_(data.cflName);
  const found = rows.find(r => normalizeCode_(r['BCC Code']) === bcc || safeString_(r['CFL Name']).toLowerCase() === cfl.toLowerCase());
  const rec = {'BCC Code':bcc,'CFL Name':cfl,'Adjacent Block 1':safeString_(data.adjacent1,150),'Adjacent Block 2':safeString_(data.adjacent2,150),'Updated At':new Date()};
  if (found) Object.keys(rec).forEach(k => setByHeader_(sh,found._row,headers,k,rec[k]));
  else sh.appendRow(headers.map(h => rec[h] !== undefined ? rec[h] : ''));
}


// v3.9: Admin CFL/Block tab edits ONLY the two adjacent blocks.
// All other CFL master fields are fetched from the selected CFL and are never overwritten here.
function saveCflAdjacentBlocks(token, data) {
  const admin = requireAdminWrite_(token);
  data = data || {};
  const requestedBcc = normalizeCode_(data.bccCode);
  const requestedName = safeString_(data.cflName, 200);
  if (!requestedBcc && !requestedName) throw new Error('Select a CFL first.');

  const sh = getSS_().getSheetByName(MW.SHEETS.CFL);
  const headers = getHeaders_(sh);
  const rows = sheetObjects_(MW.SHEETS.CFL);
  const found = rows.find(r =>
    (requestedBcc && normalizeCode_(r['BCC Code']) === requestedBcc) ||
    (requestedName && safeString_(r['CFL Name']).toLowerCase() === requestedName.toLowerCase())
  );
  if (!found) throw new Error('Selected CFL was not found in CFL Master.');

  const cfl = safeString_(found['CFL Name'], 200);
  const bcc = normalizeCode_(found['BCC Code']);
  const adjacent1 = safeString_(data.adjacent1, 150);
  const adjacent2 = safeString_(data.adjacent2, 150);

  setByHeader_(sh, found._row, headers, 'Adjacent Block 1', adjacent1);
  setByHeader_(sh, found._row, headers, 'Adjacent Block 2', adjacent2);
  saveBlockOverrideInternal_({cflName:cfl, bccCode:bcc, adjacent1:adjacent1, adjacent2:adjacent2});

  audit_(admin.id || admin.email, 'UPDATE_CFL_BLOCKS', 'CFL', bcc,
    JSON.stringify({cflName:cfl, adjacent1:adjacent1, adjacent2:adjacent2}));
  return {ok:true, cflName:cfl, bccCode:bcc, adjacent1:adjacent1, adjacent2:adjacent2};
}
