/**
 * 人事管理系統後端 (Google Apps Script)
 * 部署方式:貼到綁定 Google 試算表的 Apps Script,執行一次 setup(),再部署為網頁應用程式。
 * 前端以 POST (Content-Type: text/plain) 傳送 JSON: {action, token, params}
 */

var SHEETS = {
  EMP: '員工',
  LEAVE_TYPE: '假別',
  LEAVE: '請假',
  PUNCH: '打卡',
  LOC: '地點',
  SHIFT: '排班',
  SHIFT_TYPE: '班別',
  LOGIN: '登入紀錄'
};

var HEADERS = {
  '員工': ['員編', '姓名', '部門', '職稱', '角色', '主管員編', '到職日', '密碼雜湊', '狀態', 'Email', '電話'],
  '假別': ['假別', '年度額度(天)'],
  '請假': ['單號', '員編', '姓名', '假別', '開始日期', '結束日期', '天數', '事由', '狀態', '簽核人', '簽核意見', '申請時間', '簽核時間'],
  '打卡': ['時間', '日期', '員編', '姓名', '類型', '緯度', '經度', '距離(公尺)', '地點', '結果'],
  '地點': ['地點ID', '名稱', '緯度', '經度', '半徑(公尺)', '啟用'],
  '排班': ['日期', '員編', '姓名', '班別', '上班', '下班'],
  '班別': ['班別', '上班', '下班'],
  '登入紀錄': ['時間', '員編', '姓名', '結果', 'IP', '裝置']
};

var TOKEN_TTL = 6 * 60 * 60; // 秒
var SALT_KEY = 'PWD_SALT';

/* ============ 初始化 ============ */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty(SALT_KEY)) props.setProperty(SALT_KEY, Utilities.getUuid());

  var emp = ss.getSheetByName(SHEETS.EMP);
  if (emp.getLastRow() < 2) {
    emp.appendRow(['A001', '系統管理員', '人資部', '人事管理師', 'admin', '', today_(), hash_('admin1234'), '在職', '', '']);
  }
  var lt = ss.getSheetByName(SHEETS.LEAVE_TYPE);
  if (lt.getLastRow() < 2) {
    [['特休', 10], ['事假', 14], ['病假', 30], ['家庭照顧假', 7], ['婚假', 8], ['喪假', 8], ['公假', ''], ['補休', '']].forEach(function (r) { lt.appendRow(r); });
  }
  var st = ss.getSheetByName(SHEETS.SHIFT_TYPE);
  if (st.getLastRow() < 2) {
    [['早班', '08:00', '17:00'], ['中班', '12:00', '21:00'], ['晚班', '14:00', '23:00'], ['休假', '', '']].forEach(function (r) { st.appendRow(r); });
  }
  var loc = ss.getSheetByName(SHEETS.LOC);
  if (loc.getLastRow() < 2) {
    loc.appendRow(['L001', '總公司(請修改)', 25.0330, 121.5654, 100, true]);
  }
  // 避免時間/日期被試算表自動轉型
  ss.getSheetByName(SHEETS.SHIFT_TYPE).getRange('B:C').setNumberFormat('@');
  ss.getSheetByName(SHEETS.SHIFT).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SHEETS.SHIFT).getRange('E:F').setNumberFormat('@');
  ss.getSheetByName(SHEETS.EMP).getRange('G:G').setNumberFormat('@');
  ss.getSheetByName(SHEETS.LEAVE).getRange('E:F').setNumberFormat('@');
  ss.getSheetByName(SHEETS.PUNCH).getRange('B:B').setNumberFormat('@');
  Logger.log('初始化完成。預設管理員 A001 / admin1234,登入後請立即修改密碼。');
}

/* ============ 入口 ============ */

function doGet() {
  return json_({ ok: true, message: 'HR system API is running' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var req = JSON.parse(e.postData.contents || '{}');
    var action = req.action;
    var p = req.params || {};
    if (action === 'login') return json_(login_(p));

    var user = auth_(req.token);
    if (!user) return json_({ ok: false, error: '登入已逾時,請重新登入', code: 'AUTH' });

    var handler = ROUTES[action];
    if (!handler) return json_({ ok: false, error: '未知的操作: ' + action });

    lock.waitLock(20000);
    var data = handler(user, p);
    return json_({ ok: true, data: data });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) { /* 未持有鎖 */ }
  }
}

var ROUTES = {
  'me': function (u) { return publicUser_(u); },
  'changePassword': changePassword_,
  'employees.list': employeesList_,
  'employees.save': employeesSave_,
  'employees.resetPassword': employeesReset_,
  'employees.import': employeesImport_,
  'leave.types': leaveTypes_,
  'leave.balance': leaveBalance_,
  'leave.apply': leaveApply_,
  'leave.my': leaveMy_,
  'leave.cancel': leaveCancel_,
  'leave.pending': leavePending_,
  'leave.review': leaveReview_,
  'leave.all': leaveAll_,
  'punch.do': punchDo_,
  'punch.my': punchMy_,
  'punch.report': punchReport_,
  'locations.list': locationsList_,
  'locations.save': locationsSave_,
  'locations.delete': locationsDelete_,
  'locations.geocode': locationsGeocode_,
  'locations.import': locationsImport_,
  'shifts.types': shiftTypes_,
  'shifts.list': shiftsList_,
  'shifts.save': shiftsSave_,
  'login.log': loginLog_,
  'login.my': loginMy_
};

/* ============ 認證 ============ */

function login_(p) {
  var id = String(p.id || '').trim();
  var pwd = String(p.password || '');
  var ip = cleanIp_(p.ip);
  var ua = String(p.ua || '').slice(0, 200);
  if (!id || !pwd) return { ok: false, error: '請輸入員編與密碼' };
  var cache = CacheService.getScriptCache();
  var failKey = 'fail_' + id;
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) {
    logLogin_(id, '', '已鎖定', ip, ua);
    return { ok: false, error: '登入失敗次數過多,請 15 分鐘後再試' };
  }

  var emp = findEmp_(id);
  if (!emp || emp['狀態'] !== '在職' || emp['密碼雜湊'] !== hash_(pwd)) {
    cache.put(failKey, String(fails + 1), 900);
    logLogin_(id, emp ? emp['姓名'] : '', '失敗', ip, ua);
    return { ok: false, error: '員編或密碼錯誤' };
  }
  cache.remove(failKey);
  var last = lastLogin_(id); // 先取上一次登入,再寫入這一次
  logLogin_(id, emp['姓名'], '成功', ip, ua);
  var token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('tok_' + token, id, TOKEN_TTL);
  return { ok: true, data: { token: token, user: publicUser_(emp), mustChange: emp['密碼雜湊'] === hash_('admin1234'), lastLogin: last, ip: ip } };
}

/** IP 由瀏覽器自行回報(Apps Script 看不到連線來源),只接受 IPv4 / IPv6 格式,其他一律記為「未知」 */
function cleanIp_(v) {
  v = String(v || '').trim();
  if (/^\d{1,3}(\.\d{1,3}){3}$/.test(v) || /^[0-9a-fA-F:]{2,45}$/.test(v) && v.indexOf(':') >= 0) return v;
  return '未知';
}

function logLogin_(id, name, result, ip, ua) {
  try {
    sheet_(SHEETS.LOGIN).appendRow([now_(), String(id).slice(0, 30), name, result, ip, ua]);
  } catch (e) { /* 紀錄失敗不影響登入 */ }
}

function loginRow_(r) { return { time: r['時間'], empId: r['員編'], name: r['姓名'], result: r['結果'], ip: r['IP'], ua: r['裝置'] }; }

function lastLogin_(id) {
  try {
    var sh = sheet_(SHEETS.LOGIN);
    var n = sh.getLastRow();
    if (n < 2) return null;
    var from = Math.max(2, n - 1999);
    var vals = sh.getRange(from, 1, n - from + 1, 6).getValues();
    for (var i = vals.length - 1; i >= 0; i--) if (String(vals[i][1]) === String(id) && vals[i][3] === '成功') return { time: vals[i][0], ip: vals[i][4] };
  } catch (e) { /* ignore */ }
  return null;
}

/** 管理員:查詢登入紀錄(可依月份、員編篩選),最新在前,最多 500 筆 */
function loginLog_(u, p) {
  requireAdmin_(u);
  var month = p.month ? String(p.month) : '';
  var emp = p.empId ? String(p.empId).trim() : '';
  var out = readObjs_(SHEETS.LOGIN).filter(function (r) {
    return (!month || String(r['時間']).slice(0, 7) === month) && (!emp || String(r['員編']) === emp);
  }).map(loginRow_).reverse();
  return out.slice(0, 500);
}

/** 員工:自己最近的登入紀錄(20 筆),可用來發現可疑登入 */
function loginMy_(u) {
  var out = readObjs_(SHEETS.LOGIN).filter(function (r) { return String(r['員編']) === String(u['員編']); }).map(loginRow_).reverse();
  return out.slice(0, 20);
}

function auth_(token) {
  if (!token) return null;
  var id = CacheService.getScriptCache().get('tok_' + token);
  if (!id) return null;
  var emp = findEmp_(id);
  return emp && emp['狀態'] === '在職' ? emp : null;
}

function publicUser_(e) {
  return { id: e['員編'], name: e['姓名'], dept: e['部門'], title: e['職稱'], role: e['角色'], manager: e['主管員編'] };
}

function hash_(pwd) {
  var salt = PropertiesService.getScriptProperties().getProperty(SALT_KEY) || '';
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pwd, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function changePassword_(u, p) {
  var np = String(p.newPassword || '');
  if (np.length < 8) throw new Error('新密碼至少 8 碼');
  if (u['密碼雜湊'] !== hash_(String(p.oldPassword || ''))) throw new Error('舊密碼不正確');
  setEmpField_(u['員編'], '密碼雜湊', hash_(np));
  return true;
}

function isAdmin_(u) { return u['角色'] === 'admin'; }
function isManager_(u) { return u['角色'] === 'manager' || u['角色'] === 'admin'; }
function requireAdmin_(u) { if (!isAdmin_(u)) throw new Error('需要管理員權限'); }
function requireManager_(u) { if (!isManager_(u)) throw new Error('需要主管權限'); }

/** 可管理的員編集合:admin = 全部;manager = 直屬部屬 */
function scopeIds_(u) {
  var all = readObjs_(SHEETS.EMP);
  var ids = {};
  all.forEach(function (e) {
    if (isAdmin_(u) || e['主管員編'] === u['員編']) ids[e['員編']] = true;
  });
  return ids;
}

/* ============ 員工 ============ */

function employeesList_(u) {
  requireManager_(u);
  var scope = scopeIds_(u);
  return readObjs_(SHEETS.EMP)
    .filter(function (e) { return isAdmin_(u) || scope[e['員編']]; })
    .map(function (e) {
      return { id: e['員編'], name: e['姓名'], dept: e['部門'], title: e['職稱'], role: e['角色'], manager: e['主管員編'],
        hireDate: e['到職日'], status: e['狀態'], email: e['Email'], phone: e['電話'] };
    });
}

function employeesSave_(u, p) {
  requireAdmin_(u);
  var id = String(p.id || '').trim();
  if (!id || !p.name) throw new Error('員編與姓名必填');
  if (['admin', 'manager', 'employee'].indexOf(p.role) < 0) throw new Error('角色不正確');
  var sh = sheet_(SHEETS.EMP);
  var row = findRow_(sh, 1, id);
  var rec = [id, p.name, p.dept || '', p.title || '', p.role, p.manager || '', p.hireDate || '', null, p.status || '在職', p.email || '', p.phone || ''];
  if (row) {
    var old = sh.getRange(row, 8).getValue();
    rec[7] = old;
    sh.getRange(row, 1, 1, rec.length).setValues([rec]);
  } else {
    rec[7] = hash_(String(p.password || '') || '12345678');
    sh.appendRow(rec);
  }
  return true;
}

function employeesReset_(u, p) {
  requireAdmin_(u);
  var np = String(p.password || '');
  if (np.length < 8) throw new Error('密碼至少 8 碼');
  if (!findEmp_(p.id)) throw new Error('找不到員工');
  setEmpField_(p.id, '密碼雜湊', hash_(np));
  return true;
}

var ROLE_MAP = { 'admin': 'admin', '管理員': 'admin', '人資': 'admin', 'manager': 'manager', '主管': 'manager', 'employee': 'employee', '員工': 'employee', '': 'employee' };

/**
 * 批次匯入員工。p.rows = [{id,name,dept,title,role,manager,hireDate,status,email,phone,password}]
 * 先完整驗證,有任何錯誤就一筆都不寫入(回傳 errors)。
 * 新員工沒填密碼時自動產生隨機密碼並回傳,供管理員發放。
 */
function employeesImport_(u, p) {
  requireAdmin_(u);
  var rows = p.rows || [];
  if (!rows.length) throw new Error('沒有可匯入的資料');
  if (rows.length > 500) throw new Error('單次最多匯入 500 筆');

  var sh = sheet_(SHEETS.EMP);
  var n = sh.getLastRow();
  var width = HEADERS['員工'].length;
  var values = n > 1 ? sh.getRange(2, 1, n - 1, width).getValues() : [];
  var index = {};
  values.forEach(function (r, i) {
    index[String(r[0])] = i;
    if (r[6] instanceof Date) r[6] = Utilities.formatDate(r[6], 'Asia/Taipei', 'yyyy-MM-dd');
  });

  var errors = [], seen = {}, clean = [];
  rows.forEach(function (r, i) {
    var line = r.line || (i + 2);
    var id = String(r.id == null ? '' : r.id).trim();
    var name = String(r.name == null ? '' : r.name).trim();
    var role = ROLE_MAP[String(r.role == null ? '' : r.role).trim().toLowerCase()];
    var status = String(r.status || '在職').trim();
    var hire = String(r.hireDate || '').trim();
    var pwd = String(r.password == null ? '' : r.password);
    var err = [];
    if (!id) err.push('員編必填');
    else if (!/^[A-Za-z0-9_-]{1,20}$/.test(id)) err.push('員編只能是英數字、底線、連字號(最多 20 碼)');
    else if (seen[id]) err.push('員編與第 ' + seen[id] + ' 列重複');
    if (!name) err.push('姓名必填');
    if (!role) err.push('角色必須是 員工/主管/管理員');
    if (status !== '在職' && status !== '離職') err.push('狀態必須是 在職 或 離職');
    if (hire && !validDate_(hire)) err.push('到職日需為有效日期,格式 YYYY-MM-DD');
    if (pwd && pwd.length < 8) err.push('初始密碼至少 8 碼');
    if (id && id === u['員編'] && (role !== 'admin' || status !== '在職')) err.push('不能降級或停用自己');
    if (id) seen[id] = line;
    if (err.length) errors.push({ line: line, id: id, message: err.join(';') });
    clean.push({ line: line, id: id, name: name, dept: String(r.dept || '').trim(), title: String(r.title || '').trim(), role: role, manager: String(r.manager || '').trim(),
      hireDate: hire, status: status, email: String(r.email || '').trim(), phone: String(r.phone || '').trim(), password: pwd });
  });

  // 主管員編必須存在(資料表或本批次內),且對象需為主管或管理員
  var roleOf = {};
  values.forEach(function (r) { roleOf[String(r[0])] = r[4]; });
  clean.forEach(function (c) { if (c.id) roleOf[c.id] = c.role || roleOf[c.id]; });
  clean.forEach(function (c) {
    if (!c.manager) return;
    if (c.manager === c.id) errors.push({ line: c.line, id: c.id, message: '主管不能是自己' });
    else if (!(c.manager in roleOf)) errors.push({ line: c.line, id: c.id, message: '找不到主管員編 ' + c.manager });
    else if (roleOf[c.manager] === 'employee') errors.push({ line: c.line, id: c.id, message: '主管員編 ' + c.manager + ' 的角色不是主管/管理員' });
  });

  if (errors.length) return { applied: false, errors: errors, created: [], updated: 0 };

  var created = [], updated = 0, appendRows = [];
  clean.forEach(function (c) {
    var pos = index[c.id];
    if (pos !== undefined) {
      var old = values[pos];
      var pw = c.password ? hash_(c.password) : old[7];
      values[pos] = [c.id, c.name, c.dept, c.title, c.role, c.manager, c.hireDate, pw, c.status, c.email, c.phone];
      updated++;
    } else {
      var plain = c.password || randomPassword_();
      appendRows.push([c.id, c.name, c.dept, c.title, c.role, c.manager, c.hireDate, hash_(plain), c.status, c.email, c.phone]);
      created.push({ id: c.id, name: c.name, password: plain });
    }
  });
  if (values.length) sh.getRange(2, 1, values.length, width).setValues(values);
  if (appendRows.length) sh.getRange(2 + values.length, 1, appendRows.length, width).setNumberFormat('@').setValues(appendRows);
  return { applied: true, errors: [], created: created, updated: updated };
}

function validDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  var d = new Date(s + 'T00:00:00Z');
  return !isNaN(d.getTime()) && d.toISOString().slice(0, 10) === s;
}

function randomPassword_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  var out = '';
  for (var i = 0; i < 10; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}

/* ============ 請假 ============ */

function leaveTypes_() {
  return readObjs_(SHEETS.LEAVE_TYPE).map(function (r) { return { name: r['假別'], quota: r['年度額度(天)'] === '' ? null : Number(r['年度額度(天)']) }; });
}

function leaveBalance_(u) {
  var year = new Date().getFullYear();
  var used = {};
  readObjs_(SHEETS.LEAVE).forEach(function (r) {
    if (r['員編'] === u['員編'] && (r['狀態'] === '核准' || r['狀態'] === '待簽核') && String(r['開始日期']).slice(0, 4) === String(year)) {
      used[r['假別']] = (used[r['假別']] || 0) + Number(r['天數']);
    }
  });
  return leaveTypes_().map(function (t) {
    return { type: t.name, quota: t.quota, used: used[t.name] || 0, remain: t.quota === null ? null : t.quota - (used[t.name] || 0) };
  });
}

function leaveApply_(u, p) {
  var days = Number(p.days);
  if (!p.type || !p.start || !p.end) throw new Error('請填寫假別與日期');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(p.start) || !/^\d{4}-\d{2}-\d{2}$/.test(p.end)) throw new Error('日期格式錯誤');
  if (p.end < p.start) throw new Error('結束日期不可早於開始日期');
  if (!(days > 0) || days * 2 !== Math.round(days * 2)) throw new Error('天數必須大於 0,且以 0.5 天為單位');
  var bal = leaveBalance_(u).filter(function (b) { return b.type === p.type; })[0];
  if (!bal) throw new Error('無此假別');
  if (bal.remain !== null && days > bal.remain) throw new Error('「' + p.type + '」剩餘 ' + bal.remain + ' 天,不足申請 ' + days + ' 天');
  var id = 'LV' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  sheet_(SHEETS.LEAVE).appendRow([id, u['員編'], u['姓名'], p.type, p.start, p.end, days, p.reason || '', '待簽核', '', '', now_(), '']);
  return id;
}

function leaveRow_(r) {
  return { id: r['單號'], empId: r['員編'], name: r['姓名'], type: r['假別'], start: r['開始日期'], end: r['結束日期'], days: Number(r['天數']),
    reason: r['事由'], status: r['狀態'], reviewer: r['簽核人'], comment: r['簽核意見'], appliedAt: r['申請時間'], reviewedAt: r['簽核時間'] };
}

function leaveMy_(u) {
  return readObjs_(SHEETS.LEAVE).filter(function (r) { return r['員編'] === u['員編']; }).map(leaveRow_).reverse();
}

function leaveCancel_(u, p) {
  var sh = sheet_(SHEETS.LEAVE);
  var row = findRow_(sh, 1, p.id);
  if (!row) throw new Error('找不到假單');
  var rec = sh.getRange(row, 1, 1, HEADERS['請假'].length).getValues()[0];
  if (rec[1] !== u['員編']) throw new Error('只能撤回自己的假單');
  if (rec[8] !== '待簽核') throw new Error('僅「待簽核」假單可撤回');
  sh.getRange(row, 9).setValue('撤回');
  return true;
}

function leavePending_(u) {
  requireManager_(u);
  var scope = scopeIds_(u);
  return readObjs_(SHEETS.LEAVE)
    .filter(function (r) { return r['狀態'] === '待簽核' && scope[r['員編']] && r['員編'] !== u['員編']; })
    .map(leaveRow_);
}

function leaveAll_(u, p) {
  requireManager_(u);
  var scope = scopeIds_(u);
  return readObjs_(SHEETS.LEAVE)
    .filter(function (r) { return scope[r['員編']] && (!p.month || String(r['開始日期']).slice(0, 7) === p.month); })
    .map(leaveRow_).reverse();
}

function leaveReview_(u, p) {
  requireManager_(u);
  if (p.decision !== '核准' && p.decision !== '駁回') throw new Error('簽核結果不正確');
  var sh = sheet_(SHEETS.LEAVE);
  var row = findRow_(sh, 1, p.id);
  if (!row) throw new Error('找不到假單');
  var rec = sh.getRange(row, 1, 1, HEADERS['請假'].length).getValues()[0];
  if (rec[8] !== '待簽核') throw new Error('此假單已處理');
  if (rec[1] === u['員編']) throw new Error('不可簽核自己的假單');
  if (!scopeIds_(u)[rec[1]]) throw new Error('非您的簽核範圍');
  sh.getRange(row, 9, 1, 5).setValues([[p.decision, u['姓名'], p.comment || '', rec[11], now_()]]);
  return true;
}

/* ============ GPS 打卡 ============ */

function punchDo_(u, p) {
  var lat = Number(p.lat), lng = Number(p.lng);
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) throw new Error('定位資料無效');
  if (p.type !== '上班' && p.type !== '下班') throw new Error('打卡類型錯誤');
  var best = null;
  readObjs_(SHEETS.LOC).forEach(function (l) {
    if (l['啟用'] !== true && String(l['啟用']).toUpperCase() !== 'TRUE') return;
    var d = distance_(lat, lng, Number(l['緯度']), Number(l['經度']));
    if (!best || d < best.d) best = { d: d, loc: l };
  });
  if (!best) throw new Error('尚未設定任何打卡地點,請聯絡管理員');
  var ok = best.d <= Number(best['loc']['半徑(公尺)']);
  var dist = Math.round(best.d);
  sheet_(SHEETS.PUNCH).appendRow([now_(), today_(), u['員編'], u['姓名'], p.type, lat, lng, dist, best.loc['名稱'], ok ? '成功' : '超出範圍']);
  if (!ok) throw new Error('距離「' + best.loc['名稱'] + '」' + dist + ' 公尺,超出打卡範圍(' + best.loc['半徑(公尺)'] + ' 公尺)');
  return { time: now_(), place: best.loc['名稱'], distance: dist };
}

function punchRow_(r) {
  return { time: r['時間'], date: r['日期'], empId: r['員編'], name: r['姓名'], type: r['類型'], distance: r['距離(公尺)'], place: r['地點'], result: r['結果'] };
}

function punchMy_(u, p) {
  var month = p.month || today_().slice(0, 7);
  return readObjs_(SHEETS.PUNCH)
    .filter(function (r) { return r['員編'] === u['員編'] && String(r['日期']).slice(0, 7) === month; })
    .map(punchRow_).reverse();
}

function punchReport_(u, p) {
  requireManager_(u);
  var month = p.month || today_().slice(0, 7);
  var scope = scopeIds_(u);
  return readObjs_(SHEETS.PUNCH)
    .filter(function (r) { return scope[r['員編']] && String(r['日期']).slice(0, 7) === month; })
    .map(punchRow_).reverse();
}

function distance_(lat1, lon1, lat2, lon2) {
  var R = 6371000, rad = Math.PI / 180;
  var dLat = (lat2 - lat1) * rad, dLon = (lon2 - lon1) * rad;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

/* ============ 打卡地點 ============ */

function locationsList_(u) {
  var all = readObjs_(SHEETS.LOC).map(function (l) {
    return { id: l['地點ID'], name: l['名稱'], lat: Number(l['緯度']), lng: Number(l['經度']), radius: Number(l['半徑(公尺)']), active: l['啟用'] === true || String(l['啟用']).toUpperCase() === 'TRUE' };
  });
  return isAdmin_(u) ? all : all.filter(function (l) { return l.active; });
}

function locationsSave_(u, p) {
  requireAdmin_(u);
  var lat = Number(p.lat), lng = Number(p.lng), radius = Number(p.radius);
  if (!p.name) throw new Error('請輸入地點名稱');
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) throw new Error('經緯度無效');
  if (!(radius >= 10 && radius <= 5000)) throw new Error('半徑需介於 10~5000 公尺');
  var sh = sheet_(SHEETS.LOC);
  var id = p.id || ('L' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMddHHmmss'));
  var rec = [id, p.name, lat, lng, radius, p.active !== false];
  var row = p.id ? findRow_(sh, 1, p.id) : 0;
  if (row) sh.getRange(row, 1, 1, rec.length).setValues([rec]); else sh.appendRow(rec);
  return id;
}

function locationsDelete_(u, p) {
  requireAdmin_(u);
  var sh = sheet_(SHEETS.LOC);
  var row = findRow_(sh, 1, p.id);
  if (!row) throw new Error('找不到地點');
  sh.deleteRow(row);
  return true;
}

/** 以名稱或地址查詢座標(使用 Google 地圖地理編碼),回傳最多 5 筆候選 */
function locationsGeocode_(u, p) {
  requireAdmin_(u);
  var q = String(p.query || '').trim();
  if (!q) throw new Error('請輸入要搜尋的地點名稱或地址');
  if (q.length > 200) throw new Error('搜尋文字過長');
  var res = Maps.newGeocoder().setLanguage('zh-TW').setRegion('tw').geocode(q);
  if (res.status !== 'OK' || !res.results) return [];
  return res.results.slice(0, 5).map(function (r) {
    return { name: r.formatted_address, lat: r.geometry.location.lat, lng: r.geometry.location.lng, source: 'google' };
  });
}

/**
 * 批次匯入打卡地點。p.rows = [{line,name,lat,lng,radius,active}]
 * 以「名稱」為鍵:名稱已存在就更新,否則新增。全部通過驗證才寫入。
 */
function locationsImport_(u, p) {
  requireAdmin_(u);
  var rows = p.rows || [];
  if (!rows.length) throw new Error('沒有可匯入的資料');
  if (rows.length > 300) throw new Error('單次最多匯入 300 筆');
  var sh = sheet_(SHEETS.LOC);
  var n = sh.getLastRow();
  var values = n > 1 ? sh.getRange(2, 1, n - 1, 6).getValues() : [];
  var index = {};
  values.forEach(function (r, i) { index[String(r[1]).trim()] = i; });

  var errors = [], seen = {}, clean = [];
  rows.forEach(function (r, i) {
    var line = r.line || (i + 2), name = String(r.name == null ? '' : r.name).trim();
    var lat = Number(r.lat), lng = Number(r.lng), radius = Number(r.radius === '' || r.radius == null ? 100 : r.radius), err = [];
    if (!name) err.push('名稱必填');
    else if (seen[name]) err.push('名稱與第 ' + seen[name] + ' 列重複');
    if (!isFinite(lat) || !isFinite(lng) || r.lat === '' || r.lng === '' || Math.abs(lat) > 90 || Math.abs(lng) > 180) err.push('經緯度無效');
    if (!(radius >= 10 && radius <= 5000)) err.push('半徑需介於 10~5000 公尺');
    if (name) seen[name] = line;
    if (err.length) errors.push({ line: line, name: name, message: err.join(';') });
    clean.push({ name: name, lat: lat, lng: lng, radius: radius, active: r.active !== false });
  });
  if (errors.length) return { applied: false, errors: errors, created: 0, updated: 0 };

  var created = 0, updated = 0, stamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMddHHmmss');
  clean.forEach(function (c, i) {
    var pos = index[c.name];
    if (pos !== undefined) { values[pos] = [values[pos][0], c.name, c.lat, c.lng, c.radius, c.active]; updated++; }
    else { values.push(['L' + stamp + ('0' + i).slice(-2), c.name, c.lat, c.lng, c.radius, c.active]); index[c.name] = values.length - 1; created++; }
  });
  sh.getRange(2, 1, values.length, 6).setValues(values);
  return { applied: true, errors: [], created: created, updated: updated };
}

/* ============ 排班 ============ */

function shiftTypes_() {
  return readObjs_(SHEETS.SHIFT_TYPE).map(function (r) { return { name: r['班別'], start: r['上班'], end: r['下班'] }; });
}

function shiftsList_(u, p) {
  var month = p.month || today_().slice(0, 7);
  var scope = isManager_(u) ? scopeIds_(u) : null;
  return readObjs_(SHEETS.SHIFT)
    .filter(function (r) {
      if (String(r['日期']).slice(0, 7) !== month) return false;
      return scope ? (scope[r['員編']] || r['員編'] === u['員編']) : r['員編'] === u['員編'];
    })
    .map(function (r) { return { date: r['日期'], empId: r['員編'], name: r['姓名'], shift: r['班別'], start: r['上班'], end: r['下班'] }; });
}

/** p.items = [{date, empId, shift}]  shift 為空字串表示清除該日排班 */
function shiftsSave_(u, p) {
  requireManager_(u);
  var items = p.items || [];
  if (!items.length) return 0;
  if (items.length > 500) throw new Error('單次最多 500 筆');
  var scope = scopeIds_(u);
  var types = {};
  shiftTypes_().forEach(function (t) { types[t.name] = t; });
  var emps = {};
  readObjs_(SHEETS.EMP).forEach(function (e) { emps[e['員編']] = e['姓名']; });

  var sh = sheet_(SHEETS.SHIFT);
  var values = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];
  var index = {};
  values.forEach(function (r, i) { index[r[0] + '|' + r[1]] = i; });
  var removed = {};
  items.forEach(function (it) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(it.date)) throw new Error('日期格式錯誤');
    if (!scope[it.empId]) throw new Error('無權限排班: ' + it.empId);
    var key = it.date + '|' + it.empId;
    if (!it.shift) {
      if (key in index) removed[index[key]] = true;
      return;
    }
    var t = types[it.shift];
    if (!t) throw new Error('無此班別: ' + it.shift);
    var rec = [it.date, it.empId, emps[it.empId] || '', t.name, t.start, t.end];
    if (key in index) values[index[key]] = rec; else { index[key] = values.length; values.push(rec); }
  });
  var out = values.filter(function (_, i) { return !removed[i]; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, 6).setNumberFormat('@').setValues(out);
  return items.length;
}

/* ============ 共用工具 ============ */

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh && name === SHEETS.LOGIN) { // 登入紀錄表第一次使用時自動建立,不需要重新執行 setup()
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
  }
  if (!sh) throw new Error('缺少工作表「' + name + '」,請先執行 setup()');
  return sh;
}

function readObjs_(name) {
  var sh = sheet_(name);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var head = HEADERS[name];
  var vals = sh.getRange(2, 1, n - 1, head.length).getValues();
  return vals.map(function (row) {
    var o = {};
    head.forEach(function (h, i) {
      var v = row[i];
      if (v instanceof Date) v = Utilities.formatDate(v, 'Asia/Taipei', h.indexOf('時間') >= 0 ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
      o[h] = v;
    });
    return o;
  });
}

function findRow_(sh, col, value) {
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var vals = sh.getRange(2, col, n - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(value)) return i + 2;
  return 0;
}

function findEmp_(id) {
  var list = readObjs_(SHEETS.EMP);
  for (var i = 0; i < list.length; i++) if (String(list[i]['員編']) === String(id)) return list[i];
  return null;
}

function setEmpField_(id, field, value) {
  var sh = sheet_(SHEETS.EMP);
  var row = findRow_(sh, 1, id);
  if (!row) throw new Error('找不到員工');
  sh.getRange(row, HEADERS['員工'].indexOf(field) + 1).setValue(value);
}

function now_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd'); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
