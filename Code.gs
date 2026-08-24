function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.appVersion = MW.VERSION;
  return t.evaluate()
    .setTitle('MoneyWise CFL Monitoring')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function healthCheck() {
  const ss = getSS_();
  return {ok:true, version:MW.VERSION, spreadsheet:ss.getName(), timezone:Session.getScriptTimeZone()};
}
