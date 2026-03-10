/**
 * 機房巡檢系統 - Web App 入口
 */

/**
 * 處理 GET 請求
 */
function doGet(e) {
  const p = e.parameter;
  
  // 如果有 action 參數，根據參數決定顯示什麼
  if (p.action === 'approve') {
    return showApprovalPage(parseInt(p.week), parseInt(p.level));
  }
  if (p.action === 'confirm') {
    return handleConfirmation(parseInt(p.week), parseInt(p.level), p.opinion || '', p.decision || 'approve');
  }
  
  // 預設顯示巡檢表單
  const user = Session.getActiveUser();
  const html = HtmlService.createTemplateFromFile('index');
  html.userEmail = user.getEmail();
  html.userName = user.getEmail();
  html.items = getCheckItems();
  html.today = new Date();
  
  return html.evaluate()
    .setTitle('機房巡檢系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 處理 POST 請求（儲存巡檢資料）
 */
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const result = saveInspection(data);
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 包含其他 HTML 檔案
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
