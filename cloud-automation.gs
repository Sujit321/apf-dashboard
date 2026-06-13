// ===== APF Dashboard — Cloud Automation Script =====
// Version 2.0 | Handles: data storage, alert notifications, weekly reports
// Deploy as: Web App (Execute as: Me, Who has access: Anyone)
//
// SETUP TRIGGERS:
//  1. checkAndSendAlerts  → Time-driven → Day timer → 9 AM every day
//  2. generateWeeklyReport → Time-driven → Week timer → Every Monday 8 AM

// ====================================================
// CONFIGURATION — edit these after deployment
// ====================================================
var CONFIG = {
  TELEGRAM_BOT_TOKEN: '',   // Set via the APF app settings (auto-synced)
  TELEGRAM_CHAT_ID: '',     // Set via the APF app settings (auto-synced)
  REPORT_EMAIL: '',         // Set via the APF app settings (auto-synced)
  SUMMARY_FILE: 'APF_AutomationSummary.json',
  FOLDER_NAME: 'APF Dashboard Backups'
};

// ====================================================
// MAIN ENTRY POINT
// ====================================================
function doPost(e) {
  try {
    var content = JSON.parse(e.postData.contents);
    var action = content.action;

    if (action === 'ping_automation') {
      return send({ status: 'ok', app: 'APF Dashboard Automation', version: 2 });
    }

    // Save the unencrypted summary payload (overdue tasks, stats, config)
    if (action === 'save_summary') {
      var folder = getOrCreateFolder();
      var summary = content.summary || {};
      // Persist config from app
      if (content.config) {
        var cfg = content.config;
        if (cfg.telegramToken) CONFIG.TELEGRAM_BOT_TOKEN = cfg.telegramToken;
        if (cfg.chatId) CONFIG.TELEGRAM_CHAT_ID = cfg.chatId;
        if (cfg.reportEmail) CONFIG.REPORT_EMAIL = cfg.reportEmail;
      }
      // Store the config persistently in Drive
      var cfgJson = JSON.stringify(CONFIG);
      var cfgFiles = folder.getFilesByName('APF_AutomationConfig.json');
      if (cfgFiles.hasNext()) {
        cfgFiles.next().setContent(cfgJson);
      } else {
        folder.createFile('APF_AutomationConfig.json', cfgJson, 'application/json');
      }
      // Store the summary
      var json = JSON.stringify({ ts: new Date().toISOString(), data: summary });
      var files = folder.getFilesByName(CONFIG.SUMMARY_FILE);
      if (files.hasNext()) {
        files.next().setContent(json);
      } else {
        folder.createFile(CONFIG.SUMMARY_FILE, json, 'application/json');
      }
      return send({ success: true, ts: new Date().toISOString() });
    }

    return send({ error: 'Unknown action: ' + action });
  } catch (err) {
    return send({ error: err.toString() });
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', app: 'APF Dashboard Automation v2' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ====================================================
// TIME-DRIVEN TRIGGERS (set up manually in Apps Script)
// ====================================================

// Trigger: Daily at 9 AM
function checkAndSendAlerts() {
  loadConfig();
  var summary = loadSummary();
  if (!summary) {
    Logger.log('No summary data found. Open the APF Dashboard first.');
    return;
  }

  var overdueTasks = summary.overduePlannerTasks || [];
  var overdueFollowups = summary.overdueFollowups || [];
  var overdueGoals = summary.overdueGoals || [];

  var messages = [];

  if (overdueTasks.length > 0) {
    messages.push('📋 *Overdue Planner Tasks (' + overdueTasks.length + ')*');
    overdueTasks.slice(0, 5).forEach(function(t) {
      messages.push('• ' + t.title + (t.dueDate ? ' (due ' + t.dueDate + ')' : ''));
    });
    if (overdueTasks.length > 5) messages.push('  ...and ' + (overdueTasks.length - 5) + ' more');
  }

  if (overdueFollowups.length > 0) {
    messages.push('\n🔁 *Overdue Follow-ups (' + overdueFollowups.length + ')*');
    overdueFollowups.slice(0, 5).forEach(function(f) {
      messages.push('• ' + f.title + (f.school ? ' — ' + f.school : ''));
    });
    if (overdueFollowups.length > 5) messages.push('  ...and ' + (overdueFollowups.length - 5) + ' more');
  }

  if (overdueGoals.length > 0) {
    messages.push('\n🎯 *Goals Behind Schedule (' + overdueGoals.length + ')*');
    overdueGoals.slice(0, 3).forEach(function(g) {
      messages.push('• ' + g.title + ' (' + g.progress + '% of target)');
    });
  }

  if (messages.length === 0) {
    Logger.log('No overdue items. All good!');
    sendTelegram('✅ APF Dashboard: All tasks are on track! Great work today. 🌟');
    return;
  }

  var header = '⚠️ *APF Dashboard Daily Alert*\n' + formatDate(new Date()) + '\n\n';
  var fullMsg = header + messages.join('\n');

  sendTelegram(fullMsg);
  Logger.log('Alert sent: ' + messages.length + ' lines');
}

// Trigger: Every Monday at 8 AM
function generateWeeklyReport() {
  loadConfig();
  var summary = loadSummary();
  if (!summary) {
    Logger.log('No summary data. Open APF Dashboard first.');
    return;
  }

  var stats = summary.weeklyStats || {};
  var profile = summary.userProfile || {};
  var userName = profile.name || 'Resource Person';
  var block = profile.block || '';
  var reportEmail = CONFIG.REPORT_EMAIL;

  // Build Telegram summary
  var tgMsg =
    '📊 *Weekly Report — APF Dashboard*\n' +
    formatDate(new Date()) + '\n' +
    '👤 ' + userName + (block ? ' | ' + block : '') + '\n\n' +
    '🏫 *This Week:*\n' +
    '• School Visits: ' + (stats.visitsThisWeek || 0) + '\n' +
    '• Trainings: ' + (stats.trainingsThisWeek || 0) + '\n' +
    '• Observations: ' + (stats.observationsThisWeek || 0) + '\n' +
    '• Follow-ups Closed: ' + (stats.followupsClosedThisWeek || 0) + '\n\n' +
    '📈 *Month Total:*\n' +
    '• Visits: ' + (stats.visitsThisMonth || 0) + '\n' +
    '• Trainings: ' + (stats.trainingsThisMonth || 0) + '\n' +
    '• Schools Covered: ' + (stats.schoolsCoveredThisMonth || 0) + '\n' +
    '• Teachers Reached: ' + (stats.teachersThisMonth || 0);

  sendTelegram(tgMsg);

  // Send email if configured
  if (reportEmail) {
    var emailHtml = buildWeeklyReportEmail(userName, block, stats, summary);
    MailApp.sendEmail({
      to: reportEmail,
      subject: '📊 APF Dashboard Weekly Report — ' + formatDate(new Date()),
      htmlBody: emailHtml
    });
    Logger.log('Weekly report email sent to: ' + reportEmail);
  }
}

// ====================================================
// HELPER: Build HTML Email Report
// ====================================================
function buildWeeklyReportEmail(userName, block, stats, summary) {
  var overdueTasks = summary.overduePlannerTasks || [];
  var overdueFollowups = summary.overdueFollowups || [];

  var overdueHtml = '';
  if (overdueTasks.length > 0 || overdueFollowups.length > 0) {
    overdueHtml = '<h3 style="color:#ef4444;margin:24px 0 12px;">⚠️ Pending Actions</h3><ul style="padding-left:20px;color:#374151;">';
    overdueTasks.slice(0, 5).forEach(function(t) {
      overdueHtml += '<li style="margin-bottom:6px;">📋 <b>' + t.title + '</b>' + (t.dueDate ? ' — due ' + t.dueDate : '') + '</li>';
    });
    overdueFollowups.slice(0, 5).forEach(function(f) {
      overdueHtml += '<li style="margin-bottom:6px;">🔁 <b>' + f.title + '</b>' + (f.school ? ' (' + f.school + ')' : '') + '</li>';
    });
    overdueHtml += '</ul>';
  }

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#f1f5f9;font-family:\'Segoe UI\',Arial,sans-serif;">' +
    '<div style="max-width:600px;margin:32px auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.08);">' +
    '<div style="background:linear-gradient(135deg,#6366f1,#8b5cf6);padding:32px 36px;color:white;">' +
    '<h1 style="margin:0;font-size:22px;font-weight:700;">📊 Weekly Report</h1>' +
    '<p style="margin:8px 0 0;opacity:0.85;">APF Dashboard · ' + formatDate(new Date()) + '</p>' +
    '</div>' +
    '<div style="padding:32px 36px;">' +
    '<p style="color:#374151;margin:0 0 20px;font-size:15px;">Hi <b>' + userName + '</b>' + (block ? ' (' + block + ')' : '') + ',<br>Here is your weekly activity summary:</p>' +
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:24px;">' +
    statBox('🏫', 'School Visits', stats.visitsThisWeek || 0, 'This week') +
    statBox('📚', 'Trainings', stats.trainingsThisWeek || 0, 'This week') +
    statBox('👁️', 'Observations', stats.observationsThisWeek || 0, 'This week') +
    statBox('✅', 'Follow-ups Closed', stats.followupsClosedThisWeek || 0, 'This week') +
    '</div>' +
    '<h3 style="color:#6366f1;margin:24px 0 12px;">📅 Month to Date</h3>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    tableRow('Total Visits', stats.visitsThisMonth || 0) +
    tableRow('Total Trainings', stats.trainingsThisMonth || 0) +
    tableRow('Schools Covered', stats.schoolsCoveredThisMonth || 0) +
    tableRow('Teachers Reached', stats.teachersThisMonth || 0) +
    '</table>' +
    overdueHtml +
    '</div>' +
    '<div style="padding:20px 36px;background:#f8fafc;border-top:1px solid #e2e8f0;text-align:center;color:#94a3b8;font-size:12px;">' +
    'Generated by APF Dashboard · Automated Weekly Report<br>To unsubscribe, disable Cloud Automation in the app settings.' +
    '</div></div></body></html>';
}

function statBox(emoji, label, value, sub) {
  return '<div style="background:#f8fafc;border-radius:10px;padding:16px;text-align:center;">' +
    '<div style="font-size:24px;margin-bottom:6px;">' + emoji + '</div>' +
    '<div style="font-size:26px;font-weight:700;color:#1e293b;">' + value + '</div>' +
    '<div style="font-size:12px;color:#64748b;margin-top:4px;">' + label + '</div>' +
    '<div style="font-size:11px;color:#94a3b8;">' + sub + '</div>' +
    '</div>';
}

function tableRow(label, value) {
  return '<tr style="border-bottom:1px solid #f1f5f9;">' +
    '<td style="padding:10px 0;color:#374151;">' + label + '</td>' +
    '<td style="padding:10px 0;font-weight:600;color:#1e293b;text-align:right;">' + value + '</td>' +
    '</tr>';
}

// ====================================================
// UTILITY FUNCTIONS
// ====================================================
function sendTelegram(text) {
  if (!CONFIG.TELEGRAM_BOT_TOKEN || !CONFIG.TELEGRAM_CHAT_ID) {
    Logger.log('Telegram not configured. Skipping.');
    return;
  }
  try {
    var url = 'https://api.telegram.org/bot' + CONFIG.TELEGRAM_BOT_TOKEN + '/sendMessage';
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: CONFIG.TELEGRAM_CHAT_ID,
        text: text,
        parse_mode: 'Markdown'
      })
    });
    Logger.log('Telegram message sent.');
  } catch (e) {
    Logger.log('Telegram error: ' + e.toString());
  }
}

function loadConfig() {
  var folder = getOrCreateFolder();
  var files = folder.getFilesByName('APF_AutomationConfig.json');
  if (files.hasNext()) {
    try {
      var cfg = JSON.parse(files.next().getBlob().getDataAsString());
      if (cfg.TELEGRAM_BOT_TOKEN) CONFIG.TELEGRAM_BOT_TOKEN = cfg.TELEGRAM_BOT_TOKEN;
      if (cfg.TELEGRAM_CHAT_ID) CONFIG.TELEGRAM_CHAT_ID = cfg.TELEGRAM_CHAT_ID;
      if (cfg.REPORT_EMAIL) CONFIG.REPORT_EMAIL = cfg.REPORT_EMAIL;
    } catch (e) { Logger.log('Config load error: ' + e); }
  }
}

function loadSummary() {
  var folder = getOrCreateFolder();
  var files = folder.getFilesByName(CONFIG.SUMMARY_FILE);
  if (!files.hasNext()) return null;
  try {
    var obj = JSON.parse(files.next().getBlob().getDataAsString());
    return obj.data || null;
  } catch (e) { return null; }
}

function getOrCreateFolder() {
  var folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.FOLDER_NAME);
}

function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy');
}

function send(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
