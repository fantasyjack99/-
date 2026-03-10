/**
 * 機房巡檢系統 - Web App 入口
 */

/**
 * 處理 GET 請求
 */
function doGet(e) {
  const p = e.parameter;
  
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
 * 處理 POST 請求
 */
function doPost(e) {
  try {
    // 嘗試解析 JSON
    let data;
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseError) {
        // 如果不是 JSON，可能是表單數據
        data = e.parameter;
      }
    } else {
      data = e.parameter;
    }
    
    // 判斷是巡檢資料還是審核確認
    if (data.action === 'confirm') {
      // 審核確認
      return handleConfirmation(
        parseInt(data.week),
        parseInt(data.level),
        data.opinion || '',
        data.decision || 'approve'
      );
    } else if (data.items) {
      // 巡檢資料
      const result = saveInspection(data);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Unknown request' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 包含其他 HTML 檔案
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
