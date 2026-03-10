/**
 * 機房巡檢系統 - Web App 入口
 */

/**
 * 處理 GET 請求（顯示巡檢表單）
 */
function doGet() {
  const user = Session.getActiveUser();
  
  // 建立 HTML 輸出
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
