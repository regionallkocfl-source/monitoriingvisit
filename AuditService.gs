function audit_(user, action, entityType, entityId, details) {
  try {
    const sh = getSS_().getSheetByName(MW.SHEETS.AUDIT) || ensureSheet_(MW.SHEETS.AUDIT, ['Timestamp','User','Action','Entity Type','Entity ID','Details']);
    sh.appendRow([new Date(), safeString_(user,200), safeString_(action,100), safeString_(entityType,100), safeString_(entityId,200), safeString_(details,5000)]);
  } catch (e) {}
}
