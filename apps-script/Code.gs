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
  LOGIN: '登入紀錄',
  SLOT: '名額',
  SIGNUP: '報名'
};

var HEADERS = {
  '員工': ['員編', '姓名', '部門', '職稱', '角色', '主管員編', '到職日', '密碼雜湊', '狀態', 'Email', '電話'],
  '假別': ['假別', '年度額度(天)'],
  '請假': ['單號', '員編', '姓名', '假別', '開始日期', '結束日期', '天數', '事由', '狀態', '簽核人', '簽核意見', '申請時間', '簽核時間'],
  '打卡': ['時間', '日期', '員編', '姓名', '類型', '緯度', '經度', '距離(公尺)', '地點', '結果'],
  '地點': ['地點ID', '名稱', '緯度', '經度', '半徑(公尺)', '啟用'],
  '排班': ['日期', '員編', '姓名', '班別', '上班', '下班'],
  '班別': ['班別', '上班', '下班', '單位'],
  '登入紀錄': ['時間', '員編', '姓名', '結果', 'IP', '裝置'],
  '名額': ['日期', '單位', '班別', '需求人數'],
  '報名': ['日期', '員編', '姓名', '單位', '班別', '狀態', '報名時間']
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
  ss.getSheetByName(SHEETS.SLOT).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SHEETS.SIGNUP).getRange('A:A').setNumberFormat('@');
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
  'signup.open': signupOpen_,
  'signup.apply': signupApply_,
  'signup.cancel': signupCancel_,
  'signup.list': signupList_,
  'signup.confirm': signupConfirm_,
  'signup.quota.save': signupQuotaSave_,
  'sync.drive': syncDriveRoute_,
  'qr.token': qrToken,
  'login.log': loginLog_,
  'login.my': loginMy_
};

/* ============ 認證 ============ */

function login_(p) {
  var id = String(p.id || '').trim();
  var pwd = String(p.password || '');
  var ip = cleanIp_(p.ip);
  var ua = String(p.ua || '').slice(0, 200);
  if (!id || !pwd) return { ok: false, error: '請輸入帳號與密碼' };
  var cache = CacheService.getScriptCache();
  var failKey = 'fail_' + id;
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) {
    logLogin_(id, '', '已鎖定', ip, ua);
    return { ok: false, error: '登入失敗次數過多,請 15 分鐘後再試' };
  }

  var emp = findLoginEmp_(id, pwd);
  if (!emp || emp['狀態'] !== '在職' || emp['密碼雜湊'] !== hash_(pwd)) {
    cache.put(failKey, String(fails + 1), 900);
    logLogin_(id, emp ? emp['姓名'] : '', '失敗', ip, ua);
    return { ok: false, error: '帳號或密碼錯誤' };
  }
  cache.remove(failKey);
  var empId = String(emp['員編']); // 輸入的帳號可能是姓名,token 與登入紀錄一律用員編
  var last = lastLogin_(empId); // 先取上一次登入,再寫入這一次
  logLogin_(empId, emp['姓名'], '成功', ip, ua);
  var token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('tok_' + token, empId, TOKEN_TTL);
  return { ok: true, data: { token: token, user: publicUser_(emp), mustChange: emp['密碼雜湊'] === hash_('admin1234'), lastLogin: last, ip: ip } };
}

/** 登入帳號可輸入「員編」或「姓名」。姓名可能同名,所以在符合的在職員工中找密碼也對的那一位;
 *  同名又同密碼(例如生日相同)無法分辨時一律當作登入失敗,請改用員編登入。 */
function findLoginEmp_(account, pwd) {
  var list = readObjs_(SHEETS.EMP).filter(function (e) { return e['狀態'] === '在職'; });
  var byId = list.filter(function (e) { return String(e['員編']) === account; })[0];
  if (byId) return byId;
  var h = hash_(pwd);
  var hit = list.filter(function (e) { return String(e['姓名']).trim() === account && e['密碼雜湊'] === h; });
  if (hit.length === 1) return hit[0];
  return hit.length > 1 ? null : (list.filter(function (e) { return String(e['姓名']).trim() === account; })[0] || null);
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
    if (pwd && pwd.length < 6) err.push('初始密碼至少 6 碼'); // 初始密碼可用生日 6 碼;員工自行改密碼時仍要求至少 8 碼
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

/* ---- QR 掃碼打卡:靜態 token(每個地點固定一組,簽章防偽造,可列印張貼) ---- */
function qrToken_(locId) {
  var secret = PropertiesService.getScriptProperties().getProperty(SALT_KEY) || '';
  var sig = Utilities.computeHmacSha256Signature('qr:' + locId + ':static', secret);
  return sig.slice(0, 8).map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function qrValid_(locId, token) { return qrToken_(locId) === String(token || ''); }

function findActiveLoc_(id) {
  var found = null;
  readObjs_(SHEETS.LOC).forEach(function (l) {
    if (String(l['地點ID']) === String(id) && (l['啟用'] === true || String(l['啟用']).toUpperCase() === 'TRUE')) found = l;
  });
  return found;
}

/** 管理員:取得某地點的固定 QR token(用來產生可列印的 QR) */
function qrToken(u, p) {
  requireAdmin_(u);
  var loc = findActiveLoc_(p.id);
  if (!loc) throw new Error('找不到啟用中的地點');
  return { token: qrToken_(loc['地點ID']), name: loc['名稱'] };
}

function punchByQr_(u, p) {
  var loc = findActiveLoc_(p.qr && p.qr.loc);
  if (!loc) throw new Error('QR 對應的地點不存在或已停用');
  if (!qrValid_(loc['地點ID'], p.qr.token)) throw new Error('QR 無效,請掃描現場張貼的正確 QR');
  // 靜態 QR 可能被拍照轉傳:有定位時仍檢查距離,超出範圍就拒絕;沒有定位則放行並留空
  var lat = Number(p.lat), lng = Number(p.lng), dist = '', ok = true;
  var hasPos = p.lat !== undefined && p.lat !== null && p.lat !== '' && isFinite(lat) && isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180;
  if (hasPos) { dist = Math.round(distance_(lat, lng, Number(loc['緯度']), Number(loc['經度']))); ok = dist <= Number(loc['半徑(公尺)']); }
  var place = loc['名稱'] + '(QR)';
  sheet_(SHEETS.PUNCH).appendRow([now_(), today_(), u['員編'], u['姓名'], p.type, hasPos ? lat : '', hasPos ? lng : '', dist, place, ok ? '成功' : '超出範圍']);
  if (!ok) throw new Error('距離「' + loc['名稱'] + '」' + dist + ' 公尺,超出打卡範圍(' + loc['半徑(公尺)'] + ' 公尺)');
  return { time: now_(), place: place, distance: dist };
}

function punchDo_(u, p) {
  if (p.type !== '上班' && p.type !== '下班') throw new Error('打卡類型錯誤');
  if (p.qr) return punchByQr_(u, p);
  var lat = Number(p.lat), lng = Number(p.lng);
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) throw new Error('定位資料無效');
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
  return readObjs_(SHEETS.SHIFT_TYPE).map(function (r) { return { name: r['班別'], start: r['上班'], end: r['下班'], unit: r['單位'] || '' }; });
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

/* ============ 排班報名 ============ */
/* 名額:每個「日期 + 單位 + 班別」一列,需求人數 = 可報名人數。
 * 報名:員工只能報名自己「部門」(= 派駐單位)的名額;主管錄取後寫入「排班」,狀態改為已錄取。 */

function validDate_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(d); }
/** 試算表若把日期欄轉成日期格式,直接讀到的是 Date 物件;統一轉回 yyyy-MM-dd 字串再比對 */
function dstr_(v) { return v instanceof Date ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd') : String(v); }

/** 員工:自己單位、今天(含)以後的名額;主管/管理員:全部單位,可用 p.month / p.unit 篩選 */
function signupOpen_(u, p) {
  var month = String(p.month || '');
  var mgr = isManager_(u);
  var unit = mgr ? String(p.unit || '') : String(u['部門'] || '');
  if (!mgr && !unit) return [];
  var t = today_();
  var cnt = {}, mine = {};
  readObjs_(SHEETS.SIGNUP).forEach(function (r) {
    var k = r['日期'] + '|' + r['單位'] + '|' + r['班別'];
    cnt[k] = (cnt[k] || 0) + 1;
    if (String(r['員編']) === String(u['員編'])) mine[k] = r['狀態'];
  });
  var types = {};
  shiftTypes_().forEach(function (s) { types[s.name] = s; });
  return readObjs_(SHEETS.SLOT)
    .filter(function (s) {
      var d = String(s['日期']);
      if (!validDate_(d)) return false;
      if (month && d.slice(0, 7) !== month) return false;
      if (unit && s['單位'] !== unit) return false;
      return mgr || d >= t;
    })
    .map(function (s) {
      var k = s['日期'] + '|' + s['單位'] + '|' + s['班別'], ty = types[s['班別']] || {};
      return { date: s['日期'], unit: s['單位'], shift: s['班別'], start: ty.start || '', end: ty.end || '',
        quota: Number(s['需求人數']) || 0, count: cnt[k] || 0, mine: mine[k] || '' };
    })
    .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.unit + a.shift < b.unit + b.shift ? -1 : 1); });
}

function findSlot_(date, unit, shift) {
  var list = readObjs_(SHEETS.SLOT);
  for (var i = 0; i < list.length; i++) {
    if (String(list[i]['日期']) === date && list[i]['單位'] === unit && list[i]['班別'] === shift) return list[i];
  }
  return null;
}

/** p = {date, unit, shift} */
function signupApply_(u, p) {
  var date = String(p.date || ''), unit = String(p.unit || ''), shift = String(p.shift || '');
  if (!validDate_(date)) throw new Error('日期格式錯誤');
  if (date < today_()) throw new Error('已過報名日期');
  if (!unit || String(u['部門'] || '') !== unit) throw new Error('只能報名自己所屬單位的班');
  var slot = findSlot_(date, unit, shift);
  if (!slot) throw new Error('此日期未開放報名');
  var rows = readObjs_(SHEETS.SIGNUP), count = 0;
  rows.forEach(function (r) {
    if (String(r['員編']) === String(u['員編']) && String(r['日期']) === date) throw new Error('這一天已經報名過了');
    if (String(r['日期']) === date && r['單位'] === unit && r['班別'] === shift) count++;
  });
  if (count >= (Number(slot['需求人數']) || 0)) throw new Error('名額已滿');
  sheet_(SHEETS.SIGNUP).appendRow([date, u['員編'], u['姓名'], unit, shift, '已報名', now_()]);
  return true;
}

/** 員工取消自己尚未錄取的報名。p = {date, unit, shift} */
function signupCancel_(u, p) {
  var sh = sheet_(SHEETS.SIGNUP);
  var n = sh.getLastRow();
  if (n < 2) throw new Error('找不到報名紀錄');
  var vals = sh.getRange(2, 1, n - 1, 7).getValues();
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (dstr_(r[0]) === String(p.date) && String(r[1]) === String(u['員編']) && r[3] === p.unit && r[4] === p.shift) {
      if (r[5] === '已錄取') throw new Error('已錄取的班無法自行取消,請聯絡主管');
      sh.deleteRow(i + 2);
      return true;
    }
  }
  throw new Error('找不到報名紀錄');
}

/** 主管/管理員:報名名單(管理員看全部,主管只看直屬部屬)。p = {month, unit} */
function signupList_(u, p) {
  requireManager_(u);
  var month = String(p.month || ''), unit = String(p.unit || '');
  var scope = isAdmin_(u) ? null : scopeIds_(u);
  return readObjs_(SHEETS.SIGNUP)
    .filter(function (r) {
      if (month && String(r['日期']).slice(0, 7) !== month) return false;
      if (unit && r['單位'] !== unit) return false;
      return !scope || scope[r['員編']];
    })
    .map(function (r) { return { date: r['日期'], empId: r['員編'], name: r['姓名'], unit: r['單位'], shift: r['班別'], status: r['狀態'], time: r['報名時間'] }; })
    .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.unit + a.shift < b.unit + b.shift ? -1 : 1); });
}

/** 錄取:寫入排班並把報名標為已錄取。p.items = [{date, empId, unit, shift}] */
function signupConfirm_(u, p) {
  requireManager_(u);
  var items = p.items || [];
  if (!items.length) return 0;
  if (items.length > 300) throw new Error('單次最多 300 筆');
  var sh = sheet_(SHEETS.SIGNUP);
  var n = sh.getLastRow();
  var vals = n > 1 ? sh.getRange(2, 1, n - 1, 7).getValues() : [];
  var pos = {};
  vals.forEach(function (r, i) { pos[dstr_(r[0]) + '|' + r[1] + '|' + r[3] + '|' + r[4]] = i; });
  var shifts = items.map(function (it) {
    if (!(it.date + '|' + it.empId + '|' + it.unit + '|' + it.shift in pos)) throw new Error('找不到報名: ' + it.empId + ' ' + it.date);
    return { date: it.date, empId: it.empId, shift: it.shift };
  });
  shiftsSave_(u, { items: shifts }); // 先寫排班(含權限與班別檢查),成功才改狀態
  items.forEach(function (it) { sh.getRange(pos[it.date + '|' + it.empId + '|' + it.unit + '|' + it.shift] + 2, 6).setValue('已錄取'); });
  return items.length;
}

/** 設定名額。p.items = [{date, unit, shift, quota}],quota 為 0 或空白代表移除該名額 */
function signupQuotaSave_(u, p) {
  requireManager_(u);
  var items = p.items || [];
  if (!items.length) return 0;
  if (items.length > 300) throw new Error('單次最多 300 筆');
  var types = {};
  shiftTypes_().forEach(function (t) { types[t.name] = true; });
  var sh = sheet_(SHEETS.SLOT);
  var n = sh.getLastRow();
  var values = n > 1 ? sh.getRange(2, 1, n - 1, 4).getValues() : [];
  var index = {};
  values.forEach(function (r, i) { r[0] = dstr_(r[0]); index[r[0] + '|' + r[1] + '|' + r[2]] = i; });
  var removed = {};
  items.forEach(function (it) {
    var date = String(it.date || ''), unit = String(it.unit || '').trim(), shift = String(it.shift || '');
    if (!validDate_(date)) throw new Error('日期格式錯誤');
    if (!unit) throw new Error('請填單位');
    if (!types[shift]) throw new Error('無此班別: ' + shift);
    var q = Math.floor(Number(it.quota) || 0);
    if (q < 0 || q > 500) throw new Error('需求人數不合理');
    var key = date + '|' + unit + '|' + shift;
    if (!q) { if (key in index) removed[index[key]] = true; return; }
    if (key in index) values[index[key]][3] = q; else { index[key] = values.length; values.push([date, unit, shift, q]); }
  });
  var out = values.filter(function (_, i) { return !removed[i]; });
  if (n > 1) sh.getRange(2, 1, n - 1, 4).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, 4).setNumberFormat('@').setValues(out);
  return items.length;
}

/* ============ 從 Google 雲端硬碟班表同步 ============ */
/* 讀取雲端硬碟資料夾裡的「寶雅班表.xlsx」「鷹世服班表.xlsx」(班表是 Excel),把指定月份的
 * 已排班別與「需求/預計支援人數」寫進「排班」「名額」。單向:Drive → 系統,不會改動 Drive 檔案。名額 = 缺額(需求人數 − 已排人數),缺額為 0 的班不開放報名。
 * 需求:Apps Script 啟用進階服務 Drive(appsscript.json 已宣告)、部署者對該資料夾有讀取權限。
 * 另可在 Apps Script 手動建立時間驅動觸發程序,每天執行 syncDriveDaily() 自動同步本月。 */

var DRIVE_FOLDER_ID = '1UDKmu8_Ns_4PlZu6Q771iX55PMZP0HaR';

var SYNC_UNITS = {
  '寶雅': { file: '寶雅班表.xlsx', shiftOf: function (label) { return /^夜/.test(label) ? '寶雅夜班' : '寶雅早班'; } },
  '鷹世服': { file: '鷹世服班表.xlsx', shiftOf: function () { return '鷹世服早班'; } }
};

function syncDriveRoute_(u, p) {
  requireAdmin_(u);
  var month = /^\d{4}-\d{2}$/.test(String(p.month || '')) ? p.month : today_().slice(0, 7);
  var unit = String(p.unit || '');
  if (!SYNC_UNITS[unit]) throw new Error('未知的單位: ' + unit);
  return syncDriveUnit_(unit, month);
}

/** 給時間驅動觸發程序用:同步本月兩個單位 */
function syncDriveDaily() {
  var month = today_().slice(0, 7);
  Object.keys(SYNC_UNITS).forEach(function (unit) {
    try { Logger.log(JSON.stringify(syncDriveUnit_(unit, month))); } catch (e) { Logger.log(unit + ' 同步失敗: ' + e.message); }
  });
}

/** 給時間驅動觸發程序用(建議每 15 分鐘):只有班表 .xlsx 的「最後修改時間」變了才同步,沒變就直接結束。
 *  Apps Script 無法直接監聽雲端硬碟檔案變更,所以用定時檢查修改時間的方式達到「有變更就自動同步」。 */
function syncDriveIfChanged() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return; // 有人正在操作系統,下一輪再檢查
  try {
    var props = PropertiesService.getScriptProperties();
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var month = today_().slice(0, 7);
    Object.keys(SYNC_UNITS).forEach(function (unit) {
      var it = folder.getFilesByName(SYNC_UNITS[unit].file);
      if (!it.hasNext()) return;
      var stamp = month + '|' + it.next().getLastUpdated().getTime(); // 跨月也會重新同步
      var key = 'sync_' + unit;
      if (props.getProperty(key) === stamp) return;
      try { Logger.log(JSON.stringify(syncDriveUnit_(unit, month))); props.setProperty(key, stamp); }
      catch (e) { Logger.log(unit + ' 同步失敗: ' + e.message); }
    });
  } finally {
    try { lock.releaseLock(); } catch (x) { /* 未持有鎖 */ }
  }
}

function syncDriveUnit_(unit, month) {
  var cfg = SYNC_UNITS[unit];
  var it = DriveApp.getFolderById(DRIVE_FOLDER_ID).getFilesByName(cfg.file);
  if (!it.hasNext()) throw new Error('資料夾裡找不到「' + cfg.file + '」');
  var src = it.next();
  var want = month.replace('-', '');
  var sheetRows = readXlsxSheet_(src.getBlob(), function (n) { return /^班表/.test(n) && n.replace(/\D/g, '') === want; }, 9);
  if (!sheetRows) return { unit: unit, month: month, sheet: '', shifts: 0, slots: 0, unmatched: [], note: '班表檔裡沒有 ' + month + ' 的分頁' };
  var parsed = parseScheduleRows_(sheetRows.rows, 'Asia/Taipei', cfg.shiftOf);
  parsed.sheetName = sheetRows.name;

  // 寫入前先確認用到的班別都已在「班別」工作表建立,避免寫到一半才報錯
  var have = {}, need = {};
  shiftTypes_().forEach(function (t) { have[t.name] = true; });
  parsed.shifts.concat(parsed.slots).forEach(function (x) { if (!have[x.shift]) need[x.shift] = true; });
  if (Object.keys(need).length) throw new Error('「班別」工作表缺少班別: ' + Object.keys(need).join('、') + '(請先貼入匯入檔的「班別」分頁)');

  // 姓名 → 員編:先找同單位,再找其他單位(同時支援兩個單位的人);同名多人無法對應就列入 unmatched
  var emps = readObjs_(SHEETS.EMP).filter(function (e) { return e['狀態'] === '在職'; });
  function findId(name) {
    var a = emps.filter(function (e) { return String(e['姓名']).trim() === name && e['部門'] === unit; });
    if (!a.length) a = emps.filter(function (e) { return String(e['姓名']).trim() === name; });
    return a.length === 1 ? String(a[0]['員編']) : '';
  }
  var items = [], unmatched = {}, seen = {};
  parsed.shifts.forEach(function (s) {
    var id = findId(s.name);
    if (!id) { unmatched[s.name] = true; return; }
    var k = s.date + '|' + id;
    if (seen[k]) return; seen[k] = true;
    items.push({ date: s.date, empId: id, shift: s.shift });
  });
  var admin = { '角色': 'admin', '員編': 'SYSTEM' };
  for (var i = 0; i < items.length; i += 400) shiftsSave_(admin, { items: items.slice(i, i + 400) });
  for (var j = 0; j < parsed.slots.length; j += 250) signupQuotaSave_(admin, { items: parsed.slots.slice(j, j + 250).map(function (q) { return { date: q.date, unit: unit, shift: q.shift, quota: q.quota }; }) });
  return { unit: unit, month: month, sheet: parsed.sheetName, shifts: items.length, slots: parsed.slots.length, unmatched: Object.keys(unmatched) };
}

/** 直接讀 .xlsx(zip + XML),不需要 Drive 進階服務。回傳 {name, rows} 或 null(找不到分頁)。
 *  數字儲存格維持數字(日期是 Excel 序號,由 parseScheduleRows_ 轉換)。 */
function readXlsxSheet_(blob, pick, maxCols) {
  var NS = XmlService.getNamespace('http://schemas.openxmlformats.org/spreadsheetml/2006/main');
  var RNS = XmlService.getNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
  var files = {};
  Utilities.unzip(blob.setContentType('application/zip')).forEach(function (f) { files[f.getName()] = f; });
  function xml(name) { if (!files[name]) return null; return XmlService.parse(files[name].getDataAsString('UTF-8')).getRootElement(); }
  var wb = xml('xl/workbook.xml');
  var rels = xml('xl/_rels/workbook.xml.rels');
  var target = null, sheetName = '';
  wb.getChild('sheets', NS).getChildren('sheet', NS).forEach(function (sh) {
    var n = sh.getAttribute('name').getValue();
    if (target || !pick(n)) return;
    var rid = sh.getAttribute('id', RNS).getValue();
    rels.getChildren().forEach(function (rl) {
      if (rl.getAttribute('Id').getValue() === rid) { target = rl.getAttribute('Target').getValue().replace(/^\/?(xl\/)?/, 'xl/'); sheetName = n; }
    });
  });
  if (!target) return null;
  var strings = [];
  var sst = xml('xl/sharedStrings.xml');
  if (sst) sst.getChildren('si', NS).forEach(function (si) {
    var t = '';
    si.getChildren().forEach(function (ch) {
      if (ch.getName() === 't') t += ch.getText();
      else if (ch.getName() === 'r') { var tt = ch.getChild('t', NS); if (tt) t += tt.getText(); }
    });
    strings.push(t);
  });
  function colIdx(ref) { var m = ref.match(/^([A-Z]+)/)[1], n = 0; for (var i = 0; i < m.length; i++) n = n * 26 + (m.charCodeAt(i) - 64); return n - 1; }
  var rows = [];
  xml(target).getChild('sheetData', NS).getChildren('row', NS).forEach(function (rw) {
    var rn = Number(rw.getAttribute('r').getValue()) - 1;
    var arr = [];
    for (var k = 0; k < maxCols; k++) arr.push('');
    rw.getChildren('c', NS).forEach(function (c) {
      var ci = colIdx(c.getAttribute('r').getValue());
      if (ci >= maxCols) return;
      var t = c.getAttribute('t') ? c.getAttribute('t').getValue() : 'n';
      var v = c.getChild('v', NS), val = '';
      if (t === 'inlineStr') { var is = c.getChild('is', NS), tt = is && is.getChild('t', NS); val = tt ? tt.getText() : ''; }
      else if (v) {
        var raw = v.getText();
        if (t === 's') val = strings[Number(raw)] || '';
        else if (t === 'str') val = raw;
        else if (t === 'b' || t === 'e') val = '';
        else val = raw === '' ? '' : Number(raw);
      }
      arr[ci] = val;
    });
    while (rows.length < rn) rows.push(new Array(maxCols).fill(''));
    rows[rn] = arr;
  });
  return { name: sheetName, rows: rows };
}

/** 解析班表分頁:每週一組「日期」列,下面是各班別的人名列,最後是需求/預計支援人數列。
 *  名額 = 缺額 = 需求人數 − 已排人數(同一天同一類班別,最少 0);缺額 0 的班不開放報名。 */
function parseScheduleRows_(rows, tz, shiftOf) {
  var shifts = [], demand = {}, assigned = {}, dates = null;
  function ds(v) {
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    if (typeof v === 'number' && v > 30000 && v < 80000) return Utilities.formatDate(new Date(Math.floor(v - 25569) * 86400000), 'UTC', 'yyyy-MM-dd'); // Excel 日期序號
    var m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[0] : '';
  }
  rows.forEach(function (r) {
    var label = String(r[0] == null ? '' : r[0]).trim();
    if (!label) return;
    if (label === '日期') { dates = r.slice(1).map(ds); return; }
    if (!dates || label === '星期') return;
    var m = label.match(/^(早|中|晚|夜)\s*\d{1,2}:\d{2}-\d{1,2}:\d{2}$/);
    if (m) {
      var kind0 = m[1] === '夜' || m[1] === '晚' ? '夜' : '早';
      for (var c = 1; c < r.length; c++) {
        var d = dates[c - 1], cell = r[c];
        if (!d || cell === '' || cell == null) continue;
        var name = String(cell).trim().replace(/[\(（][^\)）]*[\)）]?$/, '').replace(/\d+$/, '').trim();
        if (!name || /^[0-9.\-X×]+$/.test(name)) continue;
        shifts.push({ date: d, name: name, shift: shiftOf(label) });
        var ak = d + '|' + kind0 + '|' + name;
        if (!assigned[ak]) assigned[ak] = true;
      }
      return;
    }
    var dm = label.match(/^(早|中|晚|夜)班(需求|預計)支援人數/);
    if (dm) {
      var kind = dm[1] === '晚' || dm[1] === '夜' ? '夜' : '早';
      for (var c2 = 1; c2 < r.length; c2++) {
        var d2 = dates[c2 - 1], n = Number(r[c2]);
        if (d2 && n > 0) demand[d2 + '|' + kind] = { date: d2, kind: kind, need: Math.floor(n) };
      }
    }
  });
  var count = {};
  Object.keys(assigned).forEach(function (k) { var p = k.split('|'); var kk = p[0] + '|' + p[1]; count[kk] = (count[kk] || 0) + 1; });
  var slots = [];
  Object.keys(demand).forEach(function (k) {
    var dd = demand[k], open = dd.need - (count[k] || 0);
    if (open > 0) slots.push({ date: dd.date, kind: dd.kind, need: dd.need, assigned: count[k] || 0, quota: open, shift: shiftOf(dd.kind === '夜' ? '夜 20:00-5:00' : '早 9:00-18:00') });
  });
  return { shifts: shifts, slots: slots };
}

/* ============ 共用工具 ============ */

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh && (name === SHEETS.LOGIN || name === SHEETS.SLOT || name === SHEETS.SIGNUP)) { // 這幾張表第一次使用時自動建立,不需要重新執行 setup()
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
    if (name === SHEETS.SLOT || name === SHEETS.SIGNUP) sh.getRange('A:A').setNumberFormat('@'); // 日期維持文字,避免被轉成日期格式
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
