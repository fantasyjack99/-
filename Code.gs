/**
 * 機房巡檢系統 - 主程式（含多層級審核）
 * Author: Kelly
 * Date: 2026-03-10
 * 版本：參照 Nicole 設計的審核表單
 */

const SPREADSHEET_ID = '1r7nRnSbfRbdHOXo8KJWt8Z1F1004rCafz-3DsNHiRFI';
const SLACK_WEBHOOK_URL = 'YOUR_SLACK_WEBHOOK_URL';

const APPROVAL_LEVELS = {
  0: { name: '資安人員', key: 'security_officer', next: 1, nextName: '組長' },
  1: { name: '組長', key: 'team_leader', next: 2, nextName: '處長' },
  2: { name: '處長', key: 'director', next: null, nextName: '' }
};

/**
 * 簡化時間格式，只顯示 HH:MM
 */
function formatTime(timeStr) {
  if (!timeStr) return '--:--';
  try {
    // 如果已經是 HH:MM 格式，直接返回
    if (typeof timeStr === 'string' && timeStr.match(/^\d{2}:\d{2}$/)) {
      return timeStr;
    }
    // 嘗試解析時間
    const date = new Date(timeStr);
    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(date, 'Asia/Taipei', 'HH:mm');
    }
    return String(timeStr).substring(0, 5);
  } catch (e) {
    return '--:--';
  }
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function initSettingsSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('設定');
  if (!sheet) { sheet = ss.insertSheet('設定'); sheet.appendRow(['鍵', '值']); }
  const defaults = [
    ['security_officer', 'SLACK_USER_ID'], ['team_leader', 'SLACK_USER_ID'],
    ['director', 'SLACK_USER_ID'], ['approval_folder_id', ''], ['last_approval_week', '']
  ];
  const data = sheet.getDataRange().getValues();
  const keys = data.slice(1).map(r => r[0]);
  for (const [k, v] of defaults) { if (!keys.includes(k)) sheet.appendRow([k, v]); }
  return sheet;
}

function initApprovalRecordSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('審核記錄');
  if (!sheet) { sheet = ss.insertSheet('審核記錄'); sheet.appendRow(['週次', '開始日期', '結束日期', '層級', '審核人', '審核意見', '審核日期', '審核時間', '狀態']); }
  return sheet;
}

function getCheckItems() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('檢查項目');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      items.push({ category: data[i][0], item: data[i][1], standard: data[i][2] || '', note: data[i][3] || '' });
    }
  }
  return items;
}

/**
 * 取得本週巡檢資料（用於審核表單）
 * 改進：按日期分組，顯示巡檢人員和最後巡檢時間
 * 修正：巡檢次數按日計算，顯示最後一次巡檢記錄
 */
function getWeekInspectionDataForApproval(weekNumber) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('每日紀錄');
  if (!sheet) return { startDate: '', endDate: '', dailyRecords: [], totalDays: 0, abnormalDays: 0 };
  
  const data = sheet.getDataRange().getValues();
  const dailyData = {}; // 按日期分組
  
  for (let i = 1; i < data.length; i++) {
    let dateValue = data[i][0];
    if (!dateValue) continue;
    
    const dateStr = typeof dateValue === 'string' ? dateValue : Utilities.formatDate(new Date(dateValue), 'Asia/Taipei', 'yyyy-MM-dd');
    const date = new Date(dateStr);
    const week = getWeekNumber(date);
    
    if (week === weekNumber) {
      const dayOfWeek = date.getDay();
      if (dayOfWeek >= 1 && dayOfWeek <= 5) {
        const item = data[i];
        const result = item[3];
        const inspector = item[5] || '未知';
        const timeStr = item[6] || '';
        
        // 只保留最後一次的記錄
        if (!dailyData[dateStr]) {
          dailyData[dateStr] = {
            date: String(dateStr).substring(5, 10), // MM/DD
            fullDate: dateStr,
            inspector: inspector,
            lastTime: timeStr,
            items: [],
            hasAbnormal: false
          };
        }
        
        // 更新記錄（覆蓋之前的）
        dailyData[dateStr].items = dailyData[dateStr].items.filter(existing => existing.item !== item[2]);
        dailyData[dateStr].items.push({
          item: item[2],
          status: result,
          note: item[4] || ''
        });
        
        // 更新最後巡檢時間
        if (timeStr && timeStr > dailyData[dateStr].lastTime) {
          dailyData[dateStr].lastTime = timeStr;
        }
        
        // 標記是否有異常
        if (result === '異常') {
          dailyData[dateStr].hasAbnormal = true;
        }
      }
    }
  }
  
  const sortedDays = Object.values(dailyData).sort((a, b) => a.fullDate.localeCompare(b.fullDate));
  const startDate = sortedDays[0]?.fullDate || '';
  const endDate = sortedDays[sortedDays.length - 1]?.fullDate || '';
  const totalDays = sortedDays.length;
  const abnormalDays = sortedDays.filter(d => d.hasAbnormal).length;
  
  return { 
    startDate, 
    endDate, 
    dailyRecords: sortedDays, 
    totalDays, 
    abnormalDays 
  };
}

function getWeekNumber(date) {
  const start = new Date(date.getFullYear(), 0, 1);
  const diff = date - start;
  const oneWeek = 604800000;
  return Math.ceil((diff + start.getDay() * 86400000) / oneWeek);
}

function saveInspection(data) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('每日紀錄');
  if (!sheet) { sheet = ss.insertSheet('每日紀錄'); sheet.appendRow(['日期', '類別', '項目', '結果', '備註', '巡檢人員', '時間']); }
  
  const date = new Date();
  const dateStr = Utilities.formatDate(date, 'Asia/Taipei', 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(date, 'Asia/Taipei', 'HH:mm:ss');
  const userEmail = Session.getActiveUser().getEmail();
  
  for (const item of data.items) {
    sheet.appendRow([dateStr, item.category, item.item, item.result, item.note || '', userEmail, timeStr]);
  }
  
  return { success: true, message: '巡檢資料已儲存', date: dateStr };
}

function sendSlackNotification(data, abnormalCount) {
  const date = new Date();
  const dateStr = Utilities.formatDate(date, 'Asia/Taipei', 'yyyy-MM-dd');
  const abnormalItems = data.items.filter(item => item.result === '異常');
  let abnormalList = '';
  for (const item of abnormalItems) {
    abnormalList += `• ${item.category} - ${item.item}: ${item.note || '無備註'}\n`;
  }
  
  const payload = {
    text: '🎯 機房每週巡檢已完成',
    blocks: [
      { type: 'header', text: { type: 'plain_text', text: '🎯 機房每週巡檢已完成', emoji: true } },
      { type: 'section', fields: [
        { type: 'mrkdwn', text: `*📅 巡檢日期：*\n${dateStr}` },
        { type: 'mrkdwn', text: `*⚠️ 異常項目：*\n${abnormalCount} 項` }
      ]},
      { type: 'divider' },
      { type: 'section', text: { type: 'mrkdwn', text: `*🚨 異常項目詳情：*\n${abnormalList || '無'}` } }
    ]
  };
  
  return sendToSlack(payload);
}

/**
 * 發送 Slack 審核請求
 */
function sendApprovalRequest(weekNumber, level) {
  const levelInfo = APPROVAL_LEVELS[level];
  const approverName = levelInfo.name;
  
  const weekData = getWeekInspectionDataForApproval(weekNumber);
  const scriptUrl = ScriptApp.getService().getUrl();
  const approveLink = `${scriptUrl}?action=approve&week=${weekNumber}&level=${level}`;
  
  const payload = {
    text: `🔔 ${approverName}，請確認本週機房巡檢結果`,
    blocks: [
      { type: 'header', text: { type: 'plain_text', text: `🔔 ${approverName}，請確認本週機房巡檢結果`, emoji: true } },
      { type: 'section', text: { type: 'mrkdwn', text: `*📅 審核期間：*\n${weekData.startDate} - ${weekData.endDate}` } },
      { type: 'section', fields: [
        { type: 'mrkdwn', text: `*📋 本週巡檢次數：*\n${weekData.totalDays} 次` },
        { type: 'mrkdwn', text: `*⚠️ 異常次數：*\n${weekData.abnormalDays} 次` }
      ]},
      { type: 'divider' },
      { type: 'section', text: { type: 'mrkdwn', text: `*👇 請點擊下方連結進行審核：*\n<${approveLink}|📋 打開審核表單>` } }
    ]
  };
  
  return sendToSlack(payload);
}

function sendToSlack(payload) {
  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload) });
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function doGet(e) {
  const p = e.parameter;
  if (p.action === 'approve') return showApprovalPage(parseInt(p.week), parseInt(p.level));
  if (p.action === 'confirm') return handleConfirmation(parseInt(p.week), parseInt(p.level), p.opinion || '', p.decision || 'approve');
  return HtmlService.createHtmlOutput('<h2>機房巡檢系統</h2><p>請透過 Slack 連結進行審核</p>');
}

/**
 * 顯示審核表單（參照 Nicole 設計）
 */
/**
 * 顯示審核表單（含折疊日期列）
 */
function showApprovalPage(weekNumber, level) {
  const levelInfo = APPROVAL_LEVELS[level];
  const approverName = levelInfo.name;
  const userEmail = Session.getActiveUser().getEmail();
  
  const weekData = getWeekInspectionDataForApproval(weekNumber);
  
  // 建立折疊的日期列
  let recordsHtml = '';
  for (const day of weekData.dailyRecords) {
    let itemsHtml = '';
    for (const item of day.items) {
      const statusClass = item.status === '異常' ? 'abnormal' : 'normal';
      itemsHtml += `<tr><td>${item.item}</td><td class="${statusClass}">${item.status}</td><td>${item.note || '-'}</td></tr>`;
    }
    
    recordsHtml += `
      <div class="date-row">
        <div class="date-header" onclick="toggleDetails(this)">
          <span class="expand-icon">▶</span>
          <span class="date-label">${day.date}</span>
          <span class="inspector-info">👤 ${day.inspector} | ⏰ ${formatTime(day.lastTime) || '--:--'}</span>
        </div>
        <div class="date-details">
          <table>
            <thead><tr><th>項目</th><th>狀態</th><th>備註</th></tr></thead>
            <tbody>${itemsHtml}</tbody>
          </table>
        </div>
      </div>
    `;
  }
  
  if (weekData.dailyRecords.length === 0) {
    recordsHtml = '<div style="text-align:center;padding:20px;color:#9aa0a6;">暫無巡檢資料</div>';
  }
  
  // 處長的最後選項
  const isDirector = (level === 2);
  const decisionOptions = isDirector ? `
    <label class="decision-option selected" onclick="selectOption(this)">
      <input type="radio" name="decision" value="approve" checked>
      <span class="radio"></span>
      <span>✅ 同意、歸檔</span>
    </label>
  ` : `
    <label class="decision-option selected" onclick="selectOption(this)">
      <input type="radio" name="decision" value="approve" checked>
      <span class="radio"></span>
      <span>✅ 簽核</span>
    </label>
    <label class="decision-option" onclick="selectOption(this)">
      <input type="radio" name="decision" value="reject">
      <span class="radio"></span>
      <span>❌ 退回</span>
    </label>
  `;
  
  const scriptUrl = ScriptApp.getService().getUrl();
  
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
    .date-row { margin-bottom: 8px; border: 1px solid #e8eaed; border-radius: 8px; overflow: hidden; }
    .date-header { display: flex; align-items: center; padding: 14px 16px; background: #f8f9fa; cursor: pointer; }
    .date-header:hover { background: #f1f3f4; }
    .expand-icon { font-size: 12px; margin-right: 12px; color: #5f6368; transition: transform 0.2s; }
    .date-header.expanded .expand-icon { transform: rotate(90deg); }
    .date-label { font-weight: 600; margin-right: 16px; color: #202124; }
    .inspector-info { color: #5f6368; font-size: 13px; }
    .date-details { display: none; background: white; }
    .date-details.show { display: block; }
    .date-details table { margin: 0; }
    .date-details th, .date-details td { font-size: 12px; }
    .normal { color: #34a853; }
    .abnormal { color: #ea4335; font-weight: 500; }
    .stats { display: flex; gap: 24px; margin-bottom: 24px; }
    .stat-item { background: #f8f9fa; padding: 16px 24px; border-radius: 8px; text-align: center; flex: 1; }
    .stat-value { font-size: 24px; font-weight: 600; color: #1a73e8; }
    .stat-value.danger { color: #ea4335; }
    .stat-label { font-size: 12px; color: #5f6368; margin-top: 4px; }
    .opinion-section { margin-bottom: 24px; }
    .opinion-label { font-size: 14px; font-weight: 500; color: #202124; margin-bottom: 8px; }
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
    input[readonly] { background: #f1f3f4; }
  </style>
</head>
<body>
  <form id="approvalForm">
    <input type="hidden" name="action" value="confirm">
    <input type="hidden" name="week" value="${weekNumber}">
    <input type="hidden" name="level" value="${level}">
    <div class="container">
      <div class="header">
        <h1>🏢 機房週巡檢審核表 - ${approverName}</h1>
        <div class="subtitle">請審核本週機房巡檢結果</div>
      </div>
      <div class="period">📅 審核期間：${weekData.startDate} - ${weekData.endDate}</div>
      <div class="content">
        <div class="section-title">📋 本週巡檢記錄表（點擊展開）</div>
        ${recordsHtml}
        <div class="stats">
          <div class="stat-item"><div class="stat-value">${weekData.totalDays}</div><div class="stat-label">本週巡檢次數</div></div>
          <div class="stat-item"><div class="stat-value ${weekData.abnormalDays > 0 ? 'danger' : ''}">${weekData.abnormalDays}</div><div class="stat-label">異常次數</div></div>
        </div>
        <div class="opinion-section">
          <div class="opinion-label">💬 【審核意見】（選填）</div>
          <textarea name="opinion" placeholder="如有意見請輸入..."></textarea>
        </div>
        <div class="section-title">✓ 審核決策</div>
        <div class="decision-options">
          ${decisionOptions}
        </div>
        <div class="sign-section">
          <div class="sign-item"><label>👤 審核人：</label><input type="text" name="reviewer" value="${userEmail}" readonly></div>
          <div class="sign-item"><label>⏰ 時間：</label><input type="text" value="${new Date().toLocaleDateString('zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit' })}" readonly></div>
        </div>
      </div>
      <div class="submit-section">
        <button type="button" class="submit-btn" onclick="submitForm()">📤 提交審核</button>
      </div>
    </div>
  </form>
  <script>
    function toggleDetails(el) {
      el.classList.toggle('expanded');
      el.nextElementSibling.classList.toggle('show');
    }
    function selectOption(el) {
      document.querySelectorAll('.decision-option').forEach(o => o.classList.remove('selected'));
      el.classList.add('selected');
    }
    function submitForm() {
      const form = document.getElementById('approvalForm');
      const formData = new FormData(form);
      const data = {};
      formData.forEach((v, k) => data[k] = v);
      document.body.innerHTML = '<div style="padding:40px;text-align:center;">處理中...</div>';
      google.script.run.withSuccessHandler(function(html) {
        document.body.innerHTML = html;
      }).withFailureHandler(function(err) {
        document.body.innerHTML = '<div style="padding:40px;text-align:center;color:red;">錯誤: ' + err.message + '</div>';
      }).handleConfirmation(
        parseInt(data.week),
        parseInt(data.level),
        data.opinion || '',
        data.decision || 'approve'
      );
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
  initApprovalRecordSheet();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('審核記錄');
  sheet.appendRow([week, weekData.startDate, weekData.endDate, level, approverName, opinion || '', approveDate, approveTime, decision === 'approve' ? '已確認' : '退回']);
  
  if (decision === 'reject') {
    sendToSlack({ text: `⚠️ ${approverName} 退回審核，要求補巡\n\n意見：${opinion || '無'}` });
    return HtmlService.createHtmlOutput(`
      <div style="font-family:'Noto Sans TC',sans-serif;background:#f0f2f5;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;">
        <div style="background:white;border-radius:12px;padding:48px;text-align:center;max-width:400px;">
          <div style="font-size:48px;margin-bottom:16px;">❌</div>
          <h1 style="color:#ea4335;">已退回</h1>
          <p>${approverName} 已退回審核</p>
          <p style="color:#5f6368;font-size:14px;">意見：${opinion || '無'}</p>
        </div>
      </div>
    `);
  }
  
  const nextLevel = levelInfo.next;
  
  if (nextLevel !== null) {
    const nextApprover = APPROVAL_LEVELS[nextLevel].name;
    sendToSlack({ text: `✅ ${approverName} 已確認，轉送給 ${nextApprover}` });
    Utilities.sleep(2000);
    sendApprovalRequest(week, nextLevel);
    
    return HtmlService.createHtmlOutput(`
      <div style="font-family:'Noto Sans TC',sans-serif;background:#f0f2f5;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;">
        <div style="background:white;border-radius:12px;padding:48px;text-align:center;max-width:400px;">
          <div style="font-size:48px;margin-bottom:16px;">✅</div>
          <h1 style="color:#34a853;">確認成功！</h1>
          <p>${approverName} 已確認完成</p>
          <p>系統已自動轉送給 <strong>${nextApprover}</strong> 審核</p>
        </div>
      </div>
    `);
  } else {
    sendToSlack({ text: '✅ 處長已確認，開始電子歸檔...' });
    const archiveResult = archiveToGoogleDrive(week, weekData.startDate);
    
    if (archiveResult.success) {
      sendToSlack({ text: `✅ 電子歸檔完成！📁 ${archiveResult.url}` });
      return HtmlService.createHtmlOutput(`
        <div style="font-family:'Noto Sans TC',sans-serif;background:#f0f2f5;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;">
          <div style="background:white;border-radius:12px;padding:48px;text-align:center;max-width:400px;">
            <div style="font-size:48px;margin-bottom:16px;">✅</div>
            <h1 style="color:#34a853;">全部確認完成！</h1>
            <p><strong>處長</strong> 已確認完成</p>
            <p style="color:#34a853;">✅ 電子歸檔已完成！</p>
            <a href="${archiveResult.url}" style="display:inline-block;margin-top:16px;padding:12px 24px;background:#1a73e8;color:white;text-decoration:none;border-radius:8px;" target="_blank">查看歸檔資料夾</a>
          </div>
        </div>
      `);
    } else {
      sendToSlack({ text: `⚠️ 電子歸檔失敗: ${archiveResult.error}` });
      return HtmlService.createHtmlOutput(`
        <div style="font-family:'Noto Sans TC',sans-serif;background:#f0f2f5;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:20px;">
          <div style="background:white;border-radius:12px;padding:48px;text-align:center;max-width:400px;">
            <div style="font-size:48px;margin-bottom:16px;">⚠️</div>
            <h1 style="color:#f9ab00;">確認成功！</h1>
            <p><strong>處長</strong> 已確認完成</p>
            <p style="color:#ea4335;">⚠️ 電子歸檔失敗</p>
          </div>
        </div>
      `);
    }
  }
}

function archiveToGoogleDrive(week, dateStr) {
  try {
    const ss = getSpreadsheet();
    const folderId = getSetting('approval_folder_id');
    let folder;
    if (folderId) folder = DriveApp.getFolderById(folderId);
    else { folder = DriveApp.createFolder(`機房巡檢歸檔_${new Date().getFullYear()}`); saveSetting('approval_folder_id', folder.getId()); }
    
    const weekFolder = folder.createFolder(`第${week}週_${dateStr}`);
    DriveApp.getFileById(SPREADSHEET_ID).makeCopy(`機房巡檢_第${week}週_${dateStr}`, weekFolder);
    return { success: true, url: weekFolder.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function triggerWeeklyApproval() {
  const date = new Date();
  if (date.getDay() !== 1) return;
  const weekNumber = getWeekNumber(date);
  if (getSetting('last_approval_week') == weekNumber) return;
  initSettingsSheet();
  initApprovalRecordSheet();
  sendApprovalRequest(weekNumber - 1, 0);
  saveSetting('last_approval_week', weekNumber);
}

function testApprovalFlow() {
  return sendApprovalRequest(getWeekNumber(new Date()), 0);
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

function createWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'triggerWeeklyApproval') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('triggerWeeklyApproval').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
}
