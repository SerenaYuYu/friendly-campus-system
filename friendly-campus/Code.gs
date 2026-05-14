// ╔══════════════════════════════════════════════╗
// ║  友善校園管理平台 ─ Google Apps Script 後端  ║
// ║  部署設定：                                  ║
// ║    執行身分 → 我（帳號擁有者）               ║
// ║    存取對象 → 所有人                         ║
// ╚══════════════════════════════════════════════╝

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'liff') {
    const cfg  = _lineConfig();
    const tmpl = HtmlService.createTemplateFromFile('liff');
    tmpl.liffId = cfg.liffId || '';
    tmpl.botId  = cfg.botId  || '';
    return tmpl.evaluate()
      .setTitle('LINE助手綁定')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('友善校園管理平台')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ══════════════════════════════════════════════
//  初始化試算表結構
// ══════════════════════════════════════════════
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _createSheet(ss, '學生清單',
    ['id','姓名','班級','年級','需完成總次數','已完成次數','狀態','通報老師','備註','更新時間'],
    '#dbeafe');

  _createSheet(ss, '區域清單',
    ['id','通報老師','區域','詳細地點','預估時間','期望完成日','狀態','描述','通報時間'],
    '#fef3c7');

  _createSheet(ss, '打掃登記',
    ['id','區域id','區域名稱','類型(環教/自願)','帶隊老師','學生id','學生姓名','班級','登記時間','完成狀態','完成時間'],
    '#dcfce7');

  _createSheet(ss, '意見回饋',
    ['id','提交者姓名','類型','標題','內容','提交時間','處理狀態','管理員回覆'],
    '#fce7f3');

  if (!ss.getSheetByName('系統設定')) {
    const s = ss.insertSheet('系統設定');
    s.appendRow(['設定項目', '值', '說明']);
    s.appendRow(['teacher_password', 'teacher123', '老師登入密碼']);
    s.appendRow(['admin_password',   'admin2024',  '管理員登入密碼']);
    s.appendRow(['line_channel_access_token', '', 'LINE Messaging API Channel Access Token']);
    s.appendRow(['line_bot_basic_id', '', 'LINE Bot 的 @ID（如 @abc123，供加好友連結用）']);
    s.appendRow(['web_app_url', '', '本系統的 GAS 網址（供 LINE 訊息附連結用）']);
    s.setFrozenRows(1);
    s.getRange(1,1,1,3).setBackground('#f0fdf4').setFontWeight('bold');
  } else {
    _ensureLineConfigRows();
  }

  _createSheet(ss, '教師設定', ['Google帳號(Email)','姓名','角色(admin/teacher)','負責班級','啟用狀態'], '#ede9fe');
  _createSheet(ss, 'LINE綁定', ['LINE UserId','角色','綁定時間'], '#d1fae5');
}

function _createSheet(ss, name, headers, color) {
  if (ss.getSheetByName(name)) return;
  const s = ss.insertSheet(name);
  s.appendRow(headers);
  s.setFrozenRows(1);
  s.getRange(1, 1, 1, headers.length).setBackground(color).setFontWeight('bold');
}

// ══════════════════════════════════════════════
//  教師設定 helpers
// ══════════════════════════════════════════════
function _teachers() {
  return _read('教師設定', 5, r => ({
    email: String(r[0]), name: String(r[1]), role: String(r[2]),
    myClass: String(r[3]), enabled: String(r[4])
  }));
}

function getUserInfo() {
  initSheets();
  const email = Session.getActiveUser().getEmail();
  if (!email) return { found: false, email: '' };
  const teachers = _teachers();
  const t = teachers.find(x => x.email.toLowerCase() === email.toLowerCase());
  if (!t) return { found: false, email };
  if (t.enabled === '停用') return { found: false, email, disabled: true };
  const cfg = _lineConfig();
  return {
    found: true, email: t.email, name: t.name,
    role: t.role, myClass: t.myClass || null,
    lineBotId:  cfg.botId  || '',
    lineLiffId: cfg.liffId || ''
  };
}

function getMyFeedbacks(email) {
  const t = _teachers().find(x => x.email.toLowerCase() === (email || '').toLowerCase());
  if (!t) return [];
  return _feedbacks().filter(f => f.submitterName === t.name);
}

// ══════════════════════════════════════════════
//  密碼登入驗證
// ══════════════════════════════════════════════
function loginCheck(password) {
  initSheets();
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('系統設定').getDataRange().getValues();

  let teacherPw = 'teacher123';
  let adminPw   = 'admin2024';

  data.forEach(row => {
    if (row[0] === 'teacher_password') teacherPw = String(row[1]);
    if (row[0] === 'admin_password')   adminPw   = String(row[1]);
  });

  if (password === adminPw)   return { success: true, role: 'admin' };
  if (password === teacherPw) return { success: true, role: 'teacher' };
  return { success: false };
}

// 取得班級清單（從學生清單自動產生）
function getClassList() {
  const students = _students();
  const classes = [...new Set(students.map(s => s.cls))].filter(c => c).sort();
  return classes;
}

// 管理員更改密碼
function changePasswords(adminPw, newTeacherPw, newAdminPw) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('系統設定');
  const data  = sheet.getDataRange().getValues();

  let currentAdminPw = 'admin2024';
  data.forEach(row => { if (row[0] === 'admin_password') currentAdminPw = String(row[1]); });
  if (adminPw !== currentAdminPw) return { success: false, msg: '管理員密碼錯誤' };

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'teacher_password' && newTeacherPw) sheet.getRange(i+1, 2).setValue(newTeacherPw);
    if (data[i][0] === 'admin_password'   && newAdminPw)   sheet.getRange(i+1, 2).setValue(newAdminPw);
  }
  return { success: true };
}

// ══════════════════════════════════════════════
//  內部工具
// ══════════════════════════════════════════════
function _read(name, cols, fn) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow()-1, cols).getValues()
    .filter(r => r[0]).map(fn);
}

function _students() {
  return _read('學生清單', 10, r => ({
    id: String(r[0]), name: String(r[1]), cls: String(r[2]),
    grade: String(r[3]), total: Number(r[4])||5, done: Number(r[5])||0,
    status: String(r[6]), teacher: String(r[7]),
    note: String(r[8]), updatedAt: String(r[9])
  }));
}

function _areas() {
  return _read('區域清單', 9, r => ({
    id: String(r[0]), teacher: String(r[1]), area: String(r[2]),
    location: String(r[3]), duration: String(r[4]), deadline: String(r[5]),
    status: String(r[6]), desc: String(r[7]), createdAt: String(r[8])
  }));
}

function _regs() {
  return _read('打掃登記', 11, r => ({
    id: String(r[0]), areaId: String(r[1]), areaName: String(r[2]),
    type: String(r[3]), teacherName: String(r[4]),
    studentId: String(r[5]), studentName: String(r[6]),
    cls: String(r[7]), registeredAt: String(r[8]),
    completed: String(r[9]) === 'Y', completedAt: String(r[10])
  }));
}

function _feedbacks() {
  return _read('意見回饋', 8, r => ({
    id: String(r[0]), submitterName: String(r[1]),
    category: String(r[2]), title: String(r[3]), content: String(r[4]),
    submittedAt: String(r[5]), processingStatus: String(r[6]), reply: String(r[7])
  }));
}

function _now() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
}

// ══════════════════════════════════════════════
//  全校公開統計（去識別化）
// ══════════════════════════════════════════════
function getSchoolSummary() {
  const students = _students();
  const areas    = _areas();
  const regs     = _regs();

  const needEnvEd = students.filter(s => s.status === '進行中' && s.done < s.total);
  const totalLeft = needEnvEd.reduce((sum, s) => sum + Math.max(0, s.total - s.done), 0);

  const byClass = {};
  needEnvEd.forEach(s => {
    if (!byClass[s.cls]) byClass[s.cls] = { count: 0, times: 0 };
    byClass[s.cls].count++;
    byClass[s.cls].times += (s.total - s.done);
  });

  const pendingAreas = areas.filter(a => a.status !== '已完成').map(a => {
    const ar = regs.filter(r => r.areaId === a.id && !r.completed);
    return {
      id: a.id, area: a.area, location: a.location, status: a.status,
      duration: a.duration, deadline: a.deadline, desc: a.desc,
      teacher: a.teacher, createdAt: a.createdAt,
      envEdCount:      ar.filter(r => r.type === '環教').length,
      volunteerCount:  ar.filter(r => r.type === '自願').length
    };
  });

  return {
    totalNeedEnvEd: needEnvEd.length,
    totalRemainingTimes: totalLeft,
    completedCount: students.filter(s => s.status === '已完成' || s.done >= s.total).length,
    totalStudents: students.length,
    byClass,
    pendingAreas
  };
}

function getHonorRoll() {
  return _regs().filter(r => r.type === '自願').map(r => ({
    id: r.id, studentName: r.studentName, cls: r.cls, areaName: r.areaName,
    teacherName: r.teacherName, completed: r.completed,
    completedAt: r.completedAt, registeredAt: r.registeredAt
  }));
}

function getAllStudentsForHonor() {
  return _students().map(s => ({ id: s.id, name: s.name, cls: s.cls }));
}

function addManualHonorRecord(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('打掃登記');
  const id = Date.now().toString();
  sheet.appendRow([id, 'manual', data.reason, '自願', data.teacherName,
    data.studentId || '', data.studentName, data.cls, _now(), 'N', '']);
  _notifyAdminsPending(data.studentName, data.cls, data.reason, data.teacherName);
  return { success: true, id };
}

// ══════════════════════════════════════════════
//  班級老師：自己班的資料
// ══════════════════════════════════════════════
function getMyClassData(myClass) {
  return {
    students:      _students().filter(s => s.cls === myClass),
    registrations: _regs().filter(r => r.cls === myClass)
  };
}

// ══════════════════════════════════════════════
//  搜尋需要環教的學生（任何老師可用，可跨班）
// ══════════════════════════════════════════════
function searchEnvEdStudents(keyword) {
  const kw = String(keyword || '').toLowerCase();
  return _students()
    .filter(s => s.status === '進行中' && s.done < s.total)
    .filter(s => !kw || s.name.includes(kw) || s.cls.toLowerCase().includes(kw))
    .map(s => ({ id: s.id, name: s.name, cls: s.cls, remaining: s.total - s.done }));
}

// ══════════════════════════════════════════════
//  通報待清理區域
// ══════════════════════════════════════════════
function reportArea(area) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('區域清單');
  const id = Date.now().toString();
  sheet.appendRow([id, area.teacher, area.area, area.location,
    area.duration, area.deadline, '待指派', area.desc, _now()]);
  return { success: true, id };
}

// ══════════════════════════════════════════════
//  打掃登記
// ══════════════════════════════════════════════
function registerEnvEd(areaId, areaName, teacherName, studentList) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('打掃登記');
  const ids = [];
  studentList.forEach(s => {
    const id = Date.now().toString() + Math.random().toString(36).slice(2, 5);
    sheet.appendRow([id, areaId, areaName, '環教', teacherName,
      s.id, s.name, s.cls, _now(), 'N', '']);
    ids.push(id);
    Utilities.sleep(10);
  });
  _syncAreaStatus(areaId);
  return { success: true, ids };
}

function registerVolunteer(areaId, areaName, teacherName, studentName, cls) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('打掃登記');
  const id = Date.now().toString();
  sheet.appendRow([id, areaId, areaName, '自願', teacherName,
    '', studentName, cls, _now(), 'N', '']);
  _syncAreaStatus(areaId);
  _notifyAdminsPending(studentName, cls, areaName, teacherName);
  return { success: true, id };
}

function markComplete(registrationId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('打掃登記');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(registrationId)) continue;
    sheet.getRange(i+1, 10).setValue('Y');
    sheet.getRange(i+1, 11).setValue(_now());
    if (String(data[i][3]) === '環教' && data[i][5]) _deductStudent(String(data[i][5]));
    _syncAreaStatus(String(data[i][1]));
    return { success: true };
  }
  return { success: false };
}

function cancelRegistration(regId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('打掃登記');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(regId)) continue;
    const areaId = String(data[i][1]);
    sheet.deleteRow(i+1);
    _syncAreaStatus(areaId);
    return { success: true };
  }
  return { success: false };
}

function _deductStudent(sid) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學生清單');
  const data  = sheet.getDataRange().getValues();
  for (let j = 1; j < data.length; j++) {
    if (String(data[j][0]) !== sid) continue;
    const newDone = Number(data[j][5]) + 1;
    sheet.getRange(j+1, 6).setValue(newDone);
    sheet.getRange(j+1, 10).setValue(Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd'));
    if (newDone >= Number(data[j][4])) sheet.getRange(j+1, 7).setValue('已完成');
    break;
  }
}

function _syncAreaStatus(areaId) {
  const regs = _regs().filter(r => r.areaId === areaId);
  if (!regs.length) return;
  const st = regs.every(r => r.completed) ? '已完成' : '進行中';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('區域清單');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === areaId) { sheet.getRange(i+1, 7).setValue(st); break; }
  }
}

// ══════════════════════════════════════════════
//  意見回饋
// ══════════════════════════════════════════════
function submitFeedback(fb) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('意見回饋');
  sheet.appendRow([Date.now().toString(),
    fb.anonymous ? '（匿名）' : fb.submitterName,
    fb.category, fb.title, fb.content, _now(), '待處理', '']);
  return { success: true };
}

function getAllFeedbacks() { return _feedbacks(); }

function replyFeedback(id, status, reply) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('意見回饋');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(id)) continue;
    sheet.getRange(i+1, 7).setValue(status);
    if (reply !== undefined) sheet.getRange(i+1, 8).setValue(reply);
    return { success: true };
  }
  return { success: false };
}

// ══════════════════════════════════════════════
//  管理員功能
// ══════════════════════════════════════════════
function getAdminData() {
  return {
    students:  _students(),
    areas:     _areas(),
    regs:      _regs(),
    feedbacks: _feedbacks(),
    teachers:  _teachers().map(function(t) {
      return { email: t.email, name: t.name, role: t.role, myClass: t.myClass, active: t.enabled !== '停用' };
    })
  };
}

function saveTeacher(t) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('教師設定');
  if (!sheet) return { success: false };
  const enabled = t.active ? '啟用' : '停用';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === t.email.toLowerCase()) {
      sheet.getRange(i+1, 1, 1, 5).setValues([[t.email, t.name, t.role, t.myClass || '', enabled]]);
      return { success: true, action: 'updated' };
    }
  }
  sheet.appendRow([t.email, t.name, t.role, t.myClass || '', enabled]);
  return { success: true, action: 'added' };
}

function deleteTeacher(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('教師設定');
  if (!sheet) return { success: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

function addStudent(s) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學生清單');
  const id = Date.now().toString();
  sheet.appendRow([id, s.name, s.cls, s.grade, s.total, s.done||0,
    s.status||'進行中', s.teacher||'', s.note||'',
    Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd')]);
  return { success: true, id };
}

function updateStudent(s) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學生清單');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(s.id)) continue;
    sheet.getRange(i+1, 1, 1, 10).setValues([[
      s.id, s.name, s.cls, s.grade, s.total, s.done,
      s.status, s.teacher, s.note,
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd')
    ]]);
    return { success: true };
  }
  return { success: false };
}

function deleteStudent(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學生清單');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: false };
}

function updateArea(a) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('區域清單');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(a.id)) continue;
    sheet.getRange(i+1, 1, 1, 9).setValues([[
      a.id, a.teacher, a.area, a.location,
      a.duration, a.deadline, a.status, a.desc, data[i][8]
    ]]);
    return { success: true };
  }
  return { success: false };
}

function deleteArea(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('區域清單');
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: false };
}

// ══════════════════════════════════════════════
//  LINE 整合
// ══════════════════════════════════════════════

function _ensureLineConfigRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('系統設定');
  if (!sheet) return;
  const existing = sheet.getDataRange().getValues().map(r => String(r[0]));
  const needed = [
    ['line_channel_access_token', '', 'LINE Messaging API Channel Access Token'],
    ['line_bot_basic_id',         '', 'LINE Bot 的 @ID（如 @abc123，供加好友連結用）'],
    ['web_app_url',               '', '本系統的 GAS 網址（供 LINE 訊息附連結用）'],
    ['line_liff_id',              '', 'LIFF App ID（在 LINE Developers > LIFF 取得）']
  ];
  needed.forEach(row => { if (!existing.includes(row[0])) sheet.appendRow(row); });
}

function _lineConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('系統設定');
  if (!sheet) return { token: '', botId: '', webUrl: '' };
  const cfg = {};
  sheet.getDataRange().getValues().forEach(r => { cfg[String(r[0])] = String(r[1]); });
  return {
    token:  cfg['line_channel_access_token'] || '',
    botId:  cfg['line_bot_basic_id']         || '',
    webUrl: cfg['web_app_url']               || '',
    liffId: cfg['line_liff_id']              || ''
  };
}

function _sendLineMessage(userId, messages) {
  const { token } = _lineConfig();
  if (!token || !userId) return;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ to: userId, messages: Array.isArray(messages) ? messages : [messages] }),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log('LINE send error: ' + e); }
}

function _getLineUsers(role) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LINE綁定');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow()-1, 2).getValues()
    .filter(r => r[0] && (!role || r[1] === role))
    .map(r => String(r[0]));
}

function _notifyAdminsPending(studentName, cls, reason, teacherName) {
  const admins = _getLineUsers('admin');
  if (!admins.length) return;
  const { webUrl } = _lineConfig();
  const text = '📋 新的表揚申請待審核\n\n學生：' + studentName + '（' + cls + '）\n事由：' + reason + '\n登記老師：' + teacherName;
  admins.forEach(uid => {
    const msgs = [{ type: 'text', text }];
    if (webUrl) msgs.push({ type: 'text', text: '🔗 點此前往審核：\n' + webUrl });
    _sendLineMessage(uid, msgs);
  });
}

// 產生一次性綁定 Token（30分鐘有效）
function generateBindToken(role) {
  const validRoles = ['admin', 'teacher'];
  if (!validRoles.includes(role)) return { success: false };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('臨時綁定Token');
  if (!sheet) {
    sheet = ss.insertSheet('臨時綁定Token');
    sheet.appendRow(['Token', '角色', '建立時間']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,3).setBackground('#fef9c3').setFontWeight('bold');
  }
  // 清除過期 token
  const rows = sheet.getDataRange().getValues();
  const now  = new Date();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (!isNaN(new Date(rows[i][2])) && (now - new Date(rows[i][2])) > 30 * 60 * 1000)
      sheet.deleteRow(i + 1);
  }
  const token = Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
      role + Date.now() + Math.random())
  ).slice(0, 12).replace(/[^A-Za-z0-9]/g, 'x');

  sheet.appendRow([token, role, new Date().toISOString()]);

  const cfg = _lineConfig();
  const liffUrl = cfg.liffId
    ? 'https://liff.line.me/' + cfg.liffId + '?token=' + token
    : '';
  return { success: true, token, liffUrl, botId: cfg.botId };
}

// LIFF 頁面呼叫：驗證 Token 並完成綁定
function bindFromLiff(token, userId, displayName) {
  if (!token || !userId) return { success: false, msg: '缺少必要參數' };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('臨時綁定Token');
  if (!sheet) return { success: false, msg: '無效的綁定令牌' };

  const rows = sheet.getDataRange().getValues();
  const now  = new Date();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(token)) continue;
    if ((now - new Date(rows[i][2])) > 30 * 60 * 1000) {
      sheet.deleteRow(i + 1);
      return { success: false, msg: '綁定令牌已過期，請重新從系統操作' };
    }
    const role = String(rows[i][1]);
    sheet.deleteRow(i + 1);
    _saveLineBinding(userId, role);
    return { success: true, role };
  }
  return { success: false, msg: '無效的綁定令牌' };
}

// 管理員廣播（需要管理員密碼驗證）
function broadcastToLine(adminPw, message) {
  const data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('系統設定').getDataRange().getValues();
  let adminPwStored = 'admin2024';
  data.forEach(r => { if (r[0] === 'admin_password') adminPwStored = String(r[1]); });
  if (adminPw !== adminPwStored) return { success: false, msg: '密碼錯誤' };
  const users = _getLineUsers();
  users.forEach(uid => _sendLineMessage(uid, { type: 'text', text: '📢 系統公告\n\n' + message }));
  return { success: true, count: users.length };
}

// ── LINE Webhook (doPost) ──────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    (body.events || []).forEach(_handleLineEvent);
  } catch(err) { Logger.log('LINE webhook error: ' + err); }
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function _handleLineEvent(event) {
  const userId = event.source && event.source.userId;
  if (!userId) return;

  const { webUrl } = _lineConfig();

  if (event.type === 'follow') {
    let txt = '👋 歡迎加入友善校園管理助手！\n\n請回到系統網頁，點選「加入LINE助手」按鈕完成身份綁定，即可自動接收通知。';
    if (webUrl) txt += '\n\n🔗 系統連結：\n' + webUrl;
    _sendLineMessage(userId, { type: 'text', text: txt });
    return;
  }

  if (event.type === 'message' && event.message && event.message.type === 'text') {
    const text = (event.message.text || '').trim();
    if (text === '解除綁定') {
      _removeLineBinding(userId);
      _sendLineMessage(userId, { type: 'text', text: '✅ 已解除綁定，不再收到通知。\n如需重新綁定，請回到系統網頁點選「加入LINE助手」。' });
    } else if (text === '系統連結') {
      _sendLineMessage(userId, { type: 'text', text: webUrl ? '🔗 系統連結：\n' + webUrl : '管理員尚未設定系統連結。' });
    }
  }
}

function _saveLineBinding(userId, role) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('LINE綁定') || (() => {
    const s = ss.insertSheet('LINE綁定');
    s.appendRow(['LINE UserId','角色','綁定時間']);
    s.setFrozenRows(1);
    s.getRange(1,1,1,3).setBackground('#d1fae5').setFontWeight('bold');
    return s;
  })();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === userId) {
      sheet.getRange(i+1, 2).setValue(role);
      sheet.getRange(i+1, 3).setValue(_now());
      return;
    }
  }
  sheet.appendRow([userId, role, _now()]);
}

function _removeLineBinding(userId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LINE綁定');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === userId) { sheet.deleteRow(i+1); return; }
  }
}
