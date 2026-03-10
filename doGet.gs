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
    let data;
    const contents = e.postData?.contents || '';
    
    // 嘗試解析 JSON
    if (contents.startsWith('{') || contents.startsWith('[')) {
      data = JSON.parse(contents);
    } else {
      // 解析 URL 編碼的參數
      data = {};
      const params = contents.split('&');
      for (const param of params) {
        const [key, value] = param.split('=');
        if (key && value) {
          data[decodeURIComponent(key)] = decodeURIComponent(value.replace(/\+/g, ' '));
        }
      }
    }
    
    // 處理審核確認
    if (data.action === 'confirm') {
      return handleConfirmation(
        parseInt(data.week),
        parseInt(data.level),
        data.opinion || '',
        data.decision || 'approve'
      );
    }
    
    // 處理巡檢資料
    if (data.items) {
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
