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

// 以下是審核表單相關函數（從 Code.gs 複製過來）

const APPROVAL_LEVELS = {
  0: { name: '資安人員', key: 'security_officer', next: 1, nextName: '組長' },
  1: { name: '組長', key: 'team_leader', next: 2, nextName: '處長' },
  2: { name: '處長', key: 'director', next: null, nextName: '' }
};

function getWeekInspectionDataForApproval(weekNumber) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('每日紀錄');
  if (!sheet) return { startDate: '', endDate: '', records: [], totalInspections: 0, abnormalCount: 0 };
  
  const data = sheet.getDataRange().getValues();
  const weekRecords = [];
  let dates = [];
  let totalInspections = 0;
  let abnormalCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    let dateValue = data[i][0];
    if (!dateValue) continue;
    
    const dateStr = typeof dateValue === 'string' ? dateValue : Utilities.formatDate(new Date(dateValue), 'Asia/Taipei', 'yyyy-MM-dd');
    const date = new Date(dateStr);
    const week = getWeekNumber(date);
    
    if (week === weekNumber) {
      const dayOfWeek = date.getDay();
      if (dayOfWeek >= 1 && dayOfWeek <= 5) {
        dates.push(dateStr);
        const item = data[i];
        const result = item[3];
        totalInspections++;
        if (result === '異常') abnormalCount++;
        
        weekRecords.push({
          date: String(dateStr).substring(5, 10),
          item: item[2],
          status: result
        });
      }
    }
  }
  
  dates.sort();
  const startDate = dates[0] || '';
  const endDate = dates[dates.length - 1] || '';
  
  return { startDate, endDate, records: weekRecords, totalInspections, abnormalCount };
}

function getWeekNumber(date) {
  const start = new Date(date.getFullYear(), 0, 1);
  const diff = date - start;
  const oneWeek = 604800000;
  return Math.ceil((diff + start.getDay() * 86400000) / oneWeek);
}

function showApprovalPage(weekNumber, level) {
  const levelInfo = APPROVAL_LEVELS[level];
  const approverName = levelInfo.name;
  const weekData = getWeekInspectionDataForApproval(weekNumber);
  
  let recordsHtml = '';
  for (const r of weekData.records) {
    const statusClass = r.status === '異常' ? 'abnormal' : 'normal';
    recordsHtml += `<tr><td>${r.date}</td><td>${r.item}</td><td class="${statusClass}">${r.status}</td></tr>`;
  }
  
  if (weekData.records.length === 0) {
    recordsHtml = '<tr><td colspan="3" style="text-align:center;">暫無巡檢資料</td></tr>';
  }
  
  const scriptUrl = ScriptApp.getService().getUrl();
  const formAction = `${scriptUrl}?action=confirm&week=${weekNumber}&level=${level}`;
  
  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap" rel="stylesheet">
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Noto Sans TC', sans-serif; background: #f0f2f5; padding: 20px; }
    .container { max-width: 700px; margin: 0 auto; background: white; border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    .header { background: #1a73e8; color: white; padding: 20px 24px; }
    .header h1 { font-size: 20px; font-weight: 500; }
    .header .subtitle { font-size: 14px; opacity: 0.9; margin-top: 4px; }
    .period { background: #f8f9fa; padding: 12px 24px; border-bottom: 1px solid #e0e0e0; font-size: 14px; color: #5f6368; }
    .content { padding: 24px; }
    .section-title { font-size: 15px; font-weight: 600; color: #202124; margin-bottom: 12px; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 13px; }
    th, td { padding: 10px 8px; text-align: left; border-bottom: 1px solid #e8eaed; }
    th { background: #f1f3f4; font-weight: 500; color: #5f6368; }
    .normal { color: #34a853; }
    .abnormal { color: #ea4335; font-weight: 500; }
    .stats { display: flex; gap: 24px; margin-bottom: 24px; }
    .stat-item { background: #f8f9fa; padding: 16px 24px; border-radius: 8px; text-align: center; flex: 1; }
    .stat-value { font-size: 24px; font-weight: 600; color: #1a73e8; }
    .stat-value.danger { color: #ea4335; }
    .stat-label { font-size: 12px; color: #5f6368; margin-top: 4px; }
    .opinion-section { margin-bottom: 24px; }
    .opinion-label { font-size: 14px; font-weight: 500; color: #202124; margin-bottom: 8px; }
    .opinion-label span { font-weight: 400; color: #5f6368; font-size: 12px; }
    textarea { width: 100%; padding: 12px; border: 1px solid #dadce0; border-radius: 8px; font-family: inherit; font-size: 14px; min-height: 80px; }
    .decision-options { display: flex; flex-direction: column; gap: 12px; margin-bottom: 24px; }
    .decision-option { display: flex; align-items: center; gap: 12px; padding: 14px 16px; border: 2px solid #e8eaed; border-radius: 8px; cursor: pointer; }
    .decision-option.selected { border-color: #1a73e8; background: #e8f0fe; }
    .decision-option input { display: none; }
    .decision-option .radio { width: 20px; height: 20px; border: 2px solid #5f6368; border-radius: 50%; }
    .decision-option.selected .radio { border-color: #1a73e8; background: #1a73e8; }
    .sign-section { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 24px; }
    .sign-item label { display: block; font-size: 14px; font-weight: 500; margin-bottom: 8px; }
    .sign-item input { width: 100%; padding: 10px; border: 1px solid #dadce0; border-radius: 6px; }
    .submit-section { text-align: center; padding: 24px; background: #f8f9fa; border-top: 1px solid #e8eaed; }
    .submit-btn { background: #1a73e8; color: white; border: none; padding: 14px 48px; border-radius: 8px; font-size: 15px; cursor: pointer; }
  </style>
</head>
<body>
  <form action="${formAction}" method="post" id="approvalForm">
    <div class="container">
      <div class="header">
        <h1>🏢 機房週巡檢審核表 - ${approverName}</h1>
        <div class="subtitle">請審核本週機房巡檢結果</div>
      </div>
      <div class="period">📅 審核期間：${weekData.startDate} - ${weekData.endDate}</div>
      <div class="content">
        <div class="section-title">📋 本週巡檢記錄表</div>
        <table><thead><tr><th>日期</th><th>項目</th><th>狀態</th></tr></thead><tbody>${recordsHtml}</tbody></table>
        <div class="stats">
          <div class="stat-item"><div class="stat-value">${weekData.totalInspections}</div><div class="stat-label">本週巡檢次數</div></div>
          <div class="stat-item"><div class="stat-value ${weekData.abnormalCount > 0 ? 'danger' : ''}">${weekData.abnormalCount}</div><div class="stat-label">異常次數</div></div>
        </div>
        <div class="opinion-section">
          <div class="opinion-label">💬 【審核意見】（選填）</div>
          <textarea name="opinion" placeholder="如有意見請輸入..."></textarea>
        </div>
        <div class="section-title">✓ 審核決策</div>
        <div class="decision-options">
          <label class="decision-option selected" onclick="selectOption(this)">
            <input type="radio" name="decision" value="approve" checked>
            <span class="radio"></span>
            <span>✅ 同意，提交給 ${levelInfo.nextName}</span>
          </label>
          <label class="decision-option" onclick="selectOption(this)">
            <input type="radio" name="decision" value="reject">
            <span class="radio"></span>
            <span>❌ 退回，要求補巡</span>
          </label>
        </div>
        <div class="sign-section">
          <div class="sign-item"><label>👤 審核人：</label><input type="text" name="reviewer" placeholder="請輸入姓名" required></div>
          <div class="sign-item"><label>⏰ 時間：</label><input type="text" value="${new Date().toLocaleString('zh-TW')}" readonly></div>
        </div>
      </div>
      <div class="submit-section">
        <button type="submit" class="submit-btn">📤 提交審核</button>
      </div>
    </div>
  </form>
  <script>
    function selectOption(el) {
      document.querySelectorAll('.decision-option').forEach(o => o.classList.remove('selected'));
      el.classList.add('selected');
    }
  </script>
</body>
</html>`;
  
  return HtmlService.createHtmlOutput(html);
}

function handleConfirmation(week, level, opinion, decision) {
  const levelInfo = APPROVAL_LEVELS[level];
  const approverName = levelInfo.name;
  const now = new Date();
  const approveDate = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy-MM-dd');
  const approveTime = Utilities.formatDate(now, 'Asia/Taipei', 'HH:mm:ss');
  
  const weekData = getWeekInspectionDataForApproval(week);
  
  // 儲存審核記錄
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('審核記錄');
  if (!sheet) { sheet = ss.insertSheet('審核記錄'); sheet.appendRow(['週次', '開始日期', '結束日期', '層級', '審核人', '審核意見', '審核日期', '審核時間', '狀態']); }
  sheet.appendRow([week, weekData.startDate, weekData.endDate, level, approverName, opinion || '', approveDate, approveTime, decision === 'approve' ? '已確認' : '退回']);
  
  if (decision === 'reject') {
    sendToSlack({ text: `⚠️ ${approverName} 退回審核，要求補巡\n\n意見：${opinion || '無'}` });
    return HtmlService.createHtmlOutput('<h2>❌ 已退回</h2><p>審核意見：' + (opinion || '無') + '</p>');
  }
  
  const nextLevel = levelInfo.next;
  
  if (nextLevel !== null) {
    const nextApprover = APPROVAL_LEVELS[nextLevel].name;
    sendToSlack({ text: `✅ ${approverName} 已確認，轉送給 ${nextApprover}` });
    Utilities.sleep(2000);
    
    const scriptUrl = ScriptApp.getService().getUrl();
    const nextLink = `${scriptUrl}?action=approve&week=${week}&level=${nextLevel}`;
    sendToSlack({ text: `🔔 ${nextApprover}，請確認本週機房巡檢結果\n${nextLink}` });
    
    return HtmlService.createHtmlOutput('<h2>✅ 確認成功！</h2><p>已轉送給 ' + nextApprover + ' 審核</p>');
  } else {
    sendToSlack({ text: '✅ 處長已確認，開始電子歸檔...' });
    
    try {
      const folderId = getSetting('approval_folder_id');
      let folder;
      if (folderId) folder = DriveApp.getFolderById(folderId);
      else { folder = DriveApp.createFolder('機房巡檢歸檔_' + new Date().getFullYear()); saveSetting('approval_folder_id', folder.getId()); }
      
      const weekFolder = folder.createFolder('第' + week + '週_' + weekData.startDate);
      DriveApp.getFileById(SPREADSHEET_ID).makeCopy('機房巡檢_第' + week + '週_' + weekData.startDate, weekFolder);
      
      sendToSlack({ text: '✅ 電子歸檔完成！📁 ' + weekFolder.getUrl() });
      return HtmlService.createHtmlOutput('<h2>✅ 全部確認完成！</h2><p>電子歸檔已完成</p><a href="' + weekFolder.getUrl() + '" target="_blank">查看歸檔資料夾</a>');
    } catch(e) {
      sendToSlack({ text: '⚠️ 電子歸檔失敗: ' + e.toString() });
      return HtmlService.createHtmlOutput('<h2>✅ 確認成功！</h2><p>歸檔失敗: ' + e.toString() + '</p>');
    }
  }
}

function getSpreadsheet() {
  return SpreadsheetApp.openById('1r7nRnSbfRbdHOXo8KJWt8Z1F1004rCafz-3DsNHiRFI');
}

function getSetting(key) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('設定');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === key) return data[i][1]; }
  return null;
}

function saveSetting(key, value) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('設定');
  if (!sheet) { sheet = ss.insertSheet('設定'); sheet.appendRow(['鍵', '值']); }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; } }
  sheet.appendRow([key, value]);
}

function sendToSlack(payload) {
  const SLACK_WEBHOOK_URL = 'YOUR_SLACK_WEBHOOK_URL';
  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload) });
  } catch(e) { Logger.log(e); }
}
