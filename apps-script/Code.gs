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
  SIGNUP: '報名',
  REG: '註冊申請',
  CANCEL: '取消紀錄',
  PRIORITY: '優先名單',
  MFA: '二階段'
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
  '報名': ['日期', '員編', '姓名', '單位', '班別', '狀態', '報名時間'],
  '取消紀錄': ['時間', '日期', '員編', '姓名', '單位', '班別', '原因', '來源', '登入IP', '登入時間'],
  '優先名單': ['姓名', '說明'],
  '二階段': ['員編', '密鑰', '狀態', '備用碼雜湊', '最後使用週期', '啟用時間'],
  '註冊申請': ['申請編號', '姓名', '身分證字號', '生日', '電話', 'Email', '單位', '證件正面', '證件反面', '存摺', '狀態', '申請時間', '審核人', '審核時間', '備註', '員編']
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
    var initPwd = Utilities.getUuid().replace(/-/g, '').slice(0, 10); // 隨機初始密碼,只顯示在執行記錄,程式碼裡不放固定密碼
    emp.appendRow(['A001', '系統管理員', '人資部', '人事管理師', 'admin', '', today_(), hash_(initPwd), '在職', '', '']);
    props.setProperty('INIT_PWD_HASH', hash_(initPwd)); // 登入時用來判斷「還在用初始密碼,要強制改密碼」
    Logger.log('管理員 A001 的初始密碼:' + initPwd + '(請立即登入並修改,看完後可清除執行記錄)');
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
  Logger.log('初始化完成。');
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
    if (action === 'login.pick') return json_(loginPick_(p));
    if (action === 'login.mfa') return json_(loginMfa_(p));
    if (action === 'register.submit' || action === 'register.status') { // 新人註冊:不需登入
      lock.waitLock(20000);
      return json_({ ok: true, data: action === 'register.submit' ? registerSubmit_(p) : registerStatus_(p) });
    }

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
  'mfa.status': mfaStatus_,
  'mfa.setup': mfaSetup_,
  'mfa.enable': mfaEnable_,
  'mfa.disable': mfaDisable_,
  'mfa.reset': mfaReset_,
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
  'shifts.cancel': shiftsCancelOwn_,
  'cancel.log': cancelLog_,
  'signup.open': signupOpen_,
  'signup.apply': signupApply_,
  'signup.cancel': signupCancel_,
  'signup.list': signupList_,
  'signup.confirm': signupConfirm_,
  'signup.quota.save': signupQuotaSave_,
  'register.list': registerList_,
  'register.file': registerFile_,
  'register.approve': registerApprove_,
  'register.reject': registerReject_,
  'register.purge': registerPurge_,
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

  var cands = findLoginEmps_(id, pwd);
  var emp = cands[0];
  if (!emp || emp['狀態'] !== '在職' || emp['密碼雜湊'] !== hash_(pwd)) {
    cache.put(failKey, String(fails + 1), 900);
    logLogin_(id, emp ? emp['姓名'] : '', '失敗', ip, ua);
    return { ok: false, error: '帳號或密碼錯誤' };
  }
  cache.remove(failKey);
  if (cands.length > 1) { // 同名且同密碼(同一個人在多個單位):密碼已驗證通過,讓他自己選要進哪個單位
    var ticket = Utilities.getUuid() + Utilities.getUuid();
    cache.put('pick_' + ticket, JSON.stringify({ ids: cands.map(function (e) { return String(e['員編']); }), ip: ip, ua: ua }), 300);
    return { ok: true, data: { choose: true, ticket: ticket, options: cands.map(function (e) { return { empId: String(e['員編']), dept: String(e['部門'] || '') }; }) } };
  }
  return afterPassword_(emp, ip, ua);
}

/** 密碼(與選單位)通過後:有啟用二階段認證就先要驗證碼,否則直接登入 */
function afterPassword_(emp, ip, ua) {
  var rec = mfaRec_(String(emp['員編']));
  if (rec && rec.status === '啟用') {
    var ticket = Utilities.getUuid() + Utilities.getUuid();
    CacheService.getScriptCache().put('mfa_' + ticket, JSON.stringify({ id: String(emp['員編']), ip: ip, ua: ua }), 300);
    return { ok: true, data: { mfa: true, ticket: ticket } };
  }
  return finishLogin_(emp, ip, ua);
}

/** 登入成功後的共同收尾:寫登入紀錄、發 token */
function finishLogin_(emp, ip, ua) {
  var cache = CacheService.getScriptCache();
  var empId = String(emp['員編']); // 輸入的帳號可能是姓名,token 與登入紀錄一律用員編
  var last = lastLogin_(empId); // 先取上一次登入,再寫入這一次
  logLogin_(empId, emp['姓名'], '成功', ip, ua);
  var token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('tok_' + token, empId, TOKEN_TTL);
  return { ok: true, data: { token: token, user: publicUser_(emp), mustChange: !!PropertiesService.getScriptProperties().getProperty('INIT_PWD_HASH') && emp['密碼雜湊'] === PropertiesService.getScriptProperties().getProperty('INIT_PWD_HASH'), lastLogin: last, ip: ip } };
}

/** 同名同密碼的人選好單位後完成登入。p = {ticket, empId}(ticket 5 分鐘內有效、只能用一次) */
function loginPick_(p) {
  var cache = CacheService.getScriptCache();
  var key = 'pick_' + String(p.ticket || '');
  var raw = cache.get(key);
  if (!raw) return { ok: false, error: '選擇已逾時,請重新登入' };
  var d = JSON.parse(raw), empId = String(p.empId || '');
  if (d.ids.indexOf(empId) < 0) return { ok: false, error: '選擇無效' };
  cache.remove(key);
  var emp = findEmp_(empId);
  if (!emp || emp['狀態'] !== '在職') return { ok: false, error: '帳號無法使用' };
  return afterPassword_(emp, d.ip, d.ua);
}

/** 登入帳號可輸入「員編」或「姓名」。姓名可能同名,所以在符合的在職員工中找密碼也對的人;
 *  同名又同密碼(例如同一個人在寶雅和鷹世服各有一筆員編)會全部回傳,讓登入時選單位。
 *  回傳陣列;找不到人時回傳空陣列;姓名對但密碼都不對時回傳第一位(供登入失敗時記錄)。 */
function findLoginEmps_(account, pwd) {
  var list = readObjs_(SHEETS.EMP).filter(function (e) { return e['狀態'] === '在職'; });
  var byId = list.filter(function (e) { return String(e['員編']) === account; })[0];
  if (byId) return [byId];
  var h = hash_(pwd);
  var hit = list.filter(function (e) { return String(e['姓名']).trim() === account && e['密碼雜湊'] === h; });
  if (hit.length) return hit;
  var any = list.filter(function (e) { return String(e['姓名']).trim() === account; })[0];
  return any ? [any] : [];
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
  if (u['角色'] === 'employee') throw new Error('員工不能自行修改密碼,請聯絡人資重設');
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
  var prio = priorityNames_(), nowMs = Date.now();
  return readObjs_(SHEETS.SLOT)
    .filter(function (s) {
      var d = String(s['日期']);
      if (!validDate_(d)) return false;
      if (month && d.slice(0, 7) !== month) return false;
      if (unit && s['單位'] !== unit) return false;
      return mgr || d > t; // 員工看不到今天(含)以前的班
    })
    .map(function (s) {
      var k = s['日期'] + '|' + s['單位'] + '|' + s['班別'], ty = types[s['班別']] || {};
      var openAt = signupOpenAt_(s['單位'], String(s['日期']), u['姓名'], prio);
      return { date: s['日期'], unit: s['單位'], shift: s['班別'], start: ty.start || '', end: ty.end || '',
        quota: Number(s['需求人數']) || 0, count: cnt[k] || 0, mine: mine[k] || '',
        open: nowMs >= openAt, openAt: openAt ? openAtLabel_(openAt) : '' };
    })
    .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.unit + a.shift < b.unit + b.shift ? -1 : 1); });
}

/** 名額表「需求人數」= 班表目前尚餘缺額;報名/取消後立即加減(0 時保留列,讓員工仍看得到自己的報名)。下次同步會再以班表為準。 */
function slotAdjust_(date, unit, shift, delta) {
  var sh = sheet_(SHEETS.SLOT), n = sh.getLastRow();
  var vals = n > 1 ? sh.getRange(2, 1, n - 1, 4).getValues() : [];
  for (var i = 0; i < vals.length; i++) {
    if (dstr_(vals[i][0]) === date && vals[i][1] === unit && vals[i][2] === shift) {
      sh.getRange(i + 2, 4).setValue(Math.max(0, (Number(vals[i][3]) || 0) + delta));
      return;
    }
  }
  if (delta > 0) { sh.appendRow([date, unit, shift, delta]); sh.getRange(sh.getLastRow(), 1).setNumberFormat('@').setValue(date); }
}

/** 該日期所屬那一週(週一開頭)的週一日期,如 2026-11-04 → 2026-11-02 */
/* ============ 二階段認證(TOTP:手機驗證器 App,每 30 秒換一組 6 位數) ============ */

var B32 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';

function b32encode_(bytes) {
  var out = '', bits = 0, val = 0;
  for (var i = 0; i < bytes.length; i++) {
    val = (val << 8) | (bytes[i] & 255); bits += 8;
    while (bits >= 5) { out += B32.charAt((val >>> (bits - 5)) & 31); bits -= 5; }
  }
  if (bits > 0) out += B32.charAt((val << (5 - bits)) & 31);
  return out;
}

function b32decode_(str) {
  var out = [], bits = 0, val = 0;
  str = String(str).replace(/[\s=]/g, '').toUpperCase();
  for (var i = 0; i < str.length; i++) {
    var idx = B32.indexOf(str.charAt(i));
    if (idx < 0) continue;
    val = (val << 5) | idx; bits += 5;
    if (bits >= 8) { out.push((val >>> (bits - 8)) & 255); bits -= 8; }
  }
  return out;
}

function toSigned_(arr) { return arr.map(function (b) { b = b & 255; return b > 127 ? b - 256 : b; }); }

/** RFC 6238 / RFC 4226:HMAC-SHA1、30 秒、6 位數 */
function totpCode_(secretB32, counter) {
  var msg = [], c = counter;
  for (var i = 7; i >= 0; i--) { msg[i] = c % 256; c = Math.floor(c / 256); }
  var sig = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, toSigned_(msg), toSigned_(b32decode_(secretB32)));
  var off = sig[19] & 15;
  var bin = ((sig[off] & 0x7f) * 16777216) + ((sig[off + 1] & 255) << 16) + ((sig[off + 2] & 255) << 8) + (sig[off + 3] & 255);
  return ('000000' + (bin % 1000000)).slice(-6);
}

function mfaRec_(empId) {
  var sh = sheet_(SHEETS.MFA), row = findRow_(sh, 1, empId);
  if (!row) return null;
  var v = sh.getRange(row, 1, 1, 6).getValues()[0];
  return { sh: sh, row: row, secret: String(v[1]), status: String(v[2]), backup: String(v[3] || ''), last: Number(v[4]) || 0 };
}

/** 驗證 6 位數驗證碼(前後各容許 30 秒誤差),同一組驗證碼不能重複使用;回傳該週期,失敗回傳 -1 */
function verifyTotp_(secret, code, last) {
  code = String(code || '').replace(/\s/g, '');
  if (!/^\d{6}$/.test(code)) return -1;
  var c = Math.floor(Date.now() / 30000);
  for (var d = -1; d <= 1; d++) if (totpCode_(secret, c + d) === code && c + d > last) return c + d;
  return -1;
}

function newBackupCodes_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', out = [];
  for (var i = 0; i < 8; i++) {
    var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + i), code = '';
    for (var j = 0; j < 8; j++) code += chars.charAt((raw[j] & 255) % chars.length);
    out.push(code.slice(0, 4) + '-' + code.slice(4));
  }
  return out;
}

/** 驗證驗證碼或備用碼(備用碼用過就作廢)。失敗 5 次鎖 15 分鐘。 */
function mfaCheck_(empId, input) {
  var cache = CacheService.getScriptCache(), key = 'mfafail_' + empId, fails = Number(cache.get(key) || 0);
  if (fails >= 5) throw new Error('驗證失敗次數過多,請 15 分鐘後再試');
  var rec = mfaRec_(empId);
  if (!rec) return false;
  var ok = false, txt = String(input || '').trim();
  var per = rec.status === '啟用' || rec.status === '待驗證' ? verifyTotp_(rec.secret, txt, rec.last) : -1;
  if (per >= 0) { rec.sh.getRange(rec.row, 5).setValue(per); ok = true; }
  else if (rec.status === '啟用' && /^[A-Za-z0-9]{4}-?[A-Za-z0-9]{4}$/.test(txt)) {
    var h = hash_(txt.replace('-', '').toUpperCase()), list = rec.backup ? rec.backup.split(',') : [], k = list.indexOf(h);
    if (k >= 0) { list.splice(k, 1); rec.sh.getRange(rec.row, 4).setValue(list.join(',')); ok = true; }
  }
  if (ok) cache.remove(key); else cache.put(key, String(fails + 1), 900);
  return ok;
}

function mfaStatus_(u) {
  var rec = mfaRec_(String(u['員編']));
  var left = rec && rec.backup ? rec.backup.split(',').length : 0;
  return { allowed: isAdmin_(u), enabled: !!(rec && rec.status === '啟用'), backupLeft: left };
}

/** 管理員開始設定:產生密鑰(尚未啟用,要輸入一次正確驗證碼才算) */
function mfaSetup_(u) {
  requireAdmin_(u);
  var id = String(u['員編']), rec = mfaRec_(id);
  if (rec && rec.status === '啟用') throw new Error('已經啟用了,要重新設定請先停用');
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid() + Date.now());
  var secret = b32encode_(raw.slice(0, 20));
  var row = [id, secret, '待驗證', '', 0, ''];
  if (rec) rec.sh.getRange(rec.row, 1, 1, 6).setValues([row]); else sheet_(SHEETS.MFA).appendRow(row);
  var issuer = '東慧人事管理系統';
  return { secret: secret, uri: 'otpauth://totp/' + encodeURIComponent(issuer + ':' + u['姓名']) + '?secret=' + secret + '&issuer=' + encodeURIComponent(issuer) + '&algorithm=SHA1&digits=6&period=30' };
}

/** 輸入手機 App 顯示的驗證碼確認 → 啟用,並一次給 8 組備用碼(只顯示這一次) */
function mfaEnable_(u, p) {
  requireAdmin_(u);
  var id = String(u['員編']), rec = mfaRec_(id);
  if (!rec || rec.status !== '待驗證') throw new Error('請先按「啟用二階段認證」產生 QR code');
  if (!mfaCheck_(id, p.code)) throw new Error('驗證碼不正確,請確認手機時間正確後再試');
  var codes = newBackupCodes_();
  var hashes = codes.map(function (c) { return hash_(c.replace('-', '')); });
  rec = mfaRec_(id);
  rec.sh.getRange(rec.row, 3, 1, 4).setValues([['啟用', hashes.join(','), rec.last, now_()]]);
  return { codes: codes };
}

/** 停用自己的二階段認證:要輸入目前的驗證碼或備用碼 */
function mfaDisable_(u, p) {
  requireAdmin_(u);
  var id = String(u['員編']), rec = mfaRec_(id);
  if (!rec) return true;
  if (rec.status === '啟用' && !mfaCheck_(id, p.code)) throw new Error('驗證碼或備用碼不正確');
  rec = mfaRec_(id);
  rec.sh.deleteRow(rec.row);
  return true;
}

/** 管理員幫「別的管理員」重設(手機遺失又沒有備用碼時)。不能重設自己。p = {empId} */
function mfaReset_(u, p) {
  requireAdmin_(u);
  var id = String(p.empId || '').trim();
  if (id === String(u['員編'])) throw new Error('自己的請用「停用」');
  var rec = mfaRec_(id);
  if (rec) rec.sh.deleteRow(rec.row);
  return true;
}

/** 登入第二步:p = {ticket, code}。ticket 5 分鐘內有效,錯 5 次作廢 */
function loginMfa_(p) {
  var cache = CacheService.getScriptCache(), key = 'mfa_' + String(p.ticket || '');
  var raw = cache.get(key);
  if (!raw) return { ok: false, error: '驗證已逾時,請重新登入' };
  var d = JSON.parse(raw), tk = 'mfan_' + String(p.ticket), tries = Number(cache.get(tk) || 0);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var emp = findEmp_(d.id);
    if (!emp || emp['狀態'] !== '在職') return { ok: false, error: '帳號無法使用' };
    var ok = false;
    try { ok = mfaCheck_(d.id, p.code); } catch (e) { return { ok: false, error: e.message }; }
    if (!ok) {
      logLogin_(d.id, emp['姓名'], '二階段失敗', d.ip, d.ua);
      cache.put(tk, String(tries + 1), 300);
      if (tries + 1 >= 5) { cache.remove(key); return { ok: false, error: '驗證失敗次數過多,請重新登入' }; }
      return { ok: false, error: '驗證碼不正確' };
    }
    cache.remove(key);
    return finishLogin_(emp, d.ip, d.ua);
  } finally { lock.releaseLock(); }
}

/* ============ 寶雅報名開放時間(優先名單先報) ============ */

var PRIORITY_DEFAULT = ['王逾奧', '高意涵', '施嘉欣', '曾靖紋', '許智超', '吳淑真', '陳明輝', '林美華', '朱建誠', '鄭宇倫', '何建弘', '郭承軒', '林麗珍', '王仲尉', '楊易合'];

/** 優先名單(「優先名單」分頁的姓名欄,人資可直接在試算表增減)。第一次使用時自動填入預設 15 人。 */
function priorityNames_() {
  var sh = sheet_(SHEETS.PRIORITY), props = PropertiesService.getScriptProperties();
  if (sh.getLastRow() < 2 && !props.getProperty('PRIORITY_SEEDED')) {
    sh.getRange(2, 1, PRIORITY_DEFAULT.length, 2).setValues(PRIORITY_DEFAULT.map(function (n) { return [n, '寶雅排班優先名單']; }));
    props.setProperty('PRIORITY_SEEDED', '1');
  }
  var n = sh.getLastRow(), out = {};
  if (n > 1) sh.getRange(2, 1, n - 1, 1).getValues().forEach(function (r) { var v = String(r[0]).trim(); if (v) out[v] = true; });
  return out;
}

/** 寶雅某一天的班,這位員工從什麼時候開始可以報名(毫秒)。
 *  排班週 = 週一到週日。優先名單:排班週的上個星期一 0:00 開放;其他人:排班週前一個星期六 8:00 開放。
 *  例:11/9-11/15 這週,優先名單 11/2 0:00,其他人 11/7 8:00。其他單位不限制,回傳 0。 */
function signupOpenAt_(unit, date, name, prio) {
  if (unit !== '寶雅') return 0;
  var monday = new Date(weekKey_(date) + 'T00:00:00+08:00').getTime();
  return prio[String(name).trim()] ? monday - 7 * 86400000 : monday - 2 * 86400000 + 8 * 3600000;
}

function openAtLabel_(ms) {
  return Utilities.formatDate(new Date(ms), 'Asia/Taipei', 'MM/dd HH:mm');
}

function weekKey_(d) {
  var t = new Date(d + 'T00:00:00Z'), dow = (t.getUTCDay() + 6) % 7;
  return Utilities.formatDate(new Date(t.getTime() - dow * 86400000), 'UTC', 'yyyy-MM-dd');
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
  if (date <= today_()) throw new Error('當天(含今天)已截止報名,請報名明天以後的班');
  if (!unit || String(u['部門'] || '') !== unit) throw new Error('只能報名自己所屬單位的班');
  var slot = findSlot_(date, unit, shift);
  if (!slot) throw new Error('此日期未開放報名');
  if (isBlacklisted_(unit, u['姓名'], u['Email'])) throw new Error('很抱歉,您目前無法報名此單位的班。如有疑問請洽人資。'); // 只擋「該公司」黑名單,其他公司不受影響
  var openAt = signupOpenAt_(unit, date, u['姓名'], priorityNames_());
  if (Date.now() < openAt) throw new Error('這週的班要到 ' + openAtLabel_(openAt) + ' 才開放您報名');
  var rows = readObjs_(SHEETS.SIGNUP), count = 0;
  if (unit === '寶雅') { // 寶雅一週(週一到週日)最多報名 3 天
    var wk = weekKey_(date), mine = 0;
    rows.forEach(function (r) { if (String(r['員編']) === String(u['員編']) && r['單位'] === '寶雅' && weekKey_(String(r['日期'])) === wk) mine++; });
    if (mine >= 3) throw new Error('寶雅一週最多只能報名 3 天(週一到週日),這週已報名 ' + mine + ' 天');
  }
  rows.forEach(function (r) {
    if (String(r['員編']) === String(u['員編']) && String(r['日期']) === date) throw new Error('這一天已經報名過了');
    if (String(r['日期']) === date && r['單位'] === unit && r['班別'] === shift) count++;
  });
  if ((Number(slot['需求人數']) || 0) <= 0) throw new Error('名額已滿'); // 「需求人數」欄 = 班表目前尚餘缺額
  var st = null;
  shiftTypes_().forEach(function (t) { if (t.name === shift) st = t; });
  if (!st) throw new Error('無此班別: ' + shift);
  writeBackSchedule_(unit, date, shift, u['姓名'], false); // 先把名字寫進雲端硬碟班表(找不到空格會丟錯,此時不建立報名)
  // 報名成功即直接錄取:寫入排班(同人同日覆蓋),報名狀態記為已錄取
  var shs = sheet_(SHEETS.SHIFT), sn = shs.getLastRow(), srow = -1;
  if (sn > 1) shs.getRange(2, 1, sn - 1, 2).getValues().forEach(function (r, i) { if (dstr_(r[0]) === date && String(r[1]) === String(u['員編'])) srow = i + 2; });
  if (srow < 0) srow = sn + 1;
  shs.getRange(srow, 1, 1, 6).setNumberFormat('@').setValues([[date, u['員編'], u['姓名'], st.name, st.start, st.end]]);
  sheet_(SHEETS.SIGNUP).appendRow([date, u['員編'], u['姓名'], unit, shift, '已錄取', now_()]);
  slotAdjust_(date, unit, shift, -1); // 班表已多一人,尚餘缺額立刻減 1
  return true;
}

/** 員工取消自己尚未錄取的報名。p = {date, unit, shift} */
function signupCancel_(u, p) {
  var reason = String(p.reason || '').trim();
  if (reason.length < 2) throw new Error('請填寫取消原因');
  if (reason.length > 200) throw new Error('取消原因請在 200 字以內');
  var sh = sheet_(SHEETS.SIGNUP);
  var n = sh.getLastRow();
  if (n < 2) throw new Error('找不到報名紀錄');
  var vals = sh.getRange(2, 1, n - 1, 7).getValues();
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (dstr_(r[0]) === String(p.date) && String(r[1]) === String(u['員編']) && r[3] === p.unit && r[4] === p.shift) {
      if (r[5] === '已錄取') return shiftsCancelOwn_(u, { date: String(p.date), reason: reason }, '排班報名'); // 上班前 30 分鐘前可取消
      try { writeBackSchedule_(p.unit, String(p.date), p.shift, u['姓名'], true); } catch (e) { Logger.log('取消報名:班表清除失敗 ' + e.message); }
      sh.deleteRow(i + 2);
      slotAdjust_(String(p.date), p.unit, p.shift, 1);
      recordCancel_(u, String(p.date), p.unit, p.shift, reason, '排班報名');
      return true;
    }
  }
  throw new Error('找不到報名紀錄');
}

/** 員工取消自己的班,上班前 30 分鐘以前都可以(不論是網站報名的,還是本來就排在班表上的):
 *  從雲端硬碟班表清掉名字(寫前備份)、刪除排班、移除對應報名紀錄、尚餘缺額加回 1。p = {date} */
function shiftsCancelOwn_(u, p, src) {
  var date = String(p.date || ''), reason = String(p.reason || '').trim();
  if (!validDate_(date)) throw new Error('日期格式錯誤');
  if (reason.length < 2) throw new Error('請填寫取消原因');
  if (reason.length > 200) throw new Error('取消原因請在 200 字以內');
  var sh = sheet_(SHEETS.SHIFT), n = sh.getLastRow();
  var vals = n > 1 ? sh.getRange(2, 1, n - 1, 6).getValues() : [], hit = -1;
  vals.forEach(function (r, i) { if (dstr_(r[0]) === date && String(r[1]) === String(u['員編'])) hit = i; });
  if (hit < 0) throw new Error('找不到這一天的排班');
  var shift = String(vals[hit][3]), unit = '';
  Object.keys(SYNC_UNITS).forEach(function (k) { if (shift.indexOf(k) === 0) unit = k; });
  if (!unit) throw new Error('這個班別不能自行取消,請聯絡主管');
  // 上班前 30 分鐘(含)以前都能取消;上班時間取「班別」表,找不到就用排班列上的上班時間
  var start = '';
  shiftTypes_().forEach(function (t) { if (t.name === shift) start = String(t.start || ''); });
  var sm = (start || String(vals[hit][4] || '')).match(/(\d{1,2}):(\d{2})/);
  if (!sm) throw new Error('這個班沒有上班時間,無法自行取消,請聯絡主管');
  var deadline = new Date(date + 'T' + ('0' + sm[1]).slice(-2) + ':' + sm[2] + ':00+08:00').getTime() - 30 * 60000;
  if (Date.now() > deadline) throw new Error('已超過取消時限(上班前 30 分鐘),請聯絡主管');
  writeBackSchedule_(unit, date, shift, u['姓名'], true); // 先清班表(失敗就整個中止,不會半套)
  sh.deleteRow(hit + 2);
  var ss = sheet_(SHEETS.SIGNUP), m = ss.getLastRow();
  if (m > 1) {
    var rows = ss.getRange(2, 1, m - 1, 7).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      if (dstr_(rows[i][0]) === date && String(rows[i][1]) === String(u['員編'])) ss.deleteRow(i + 2);
    }
  }
  slotAdjust_(date, unit, shift, 1);
  recordCancel_(u, date, unit, shift, reason, src || '排班狀況');
  return true;
}


/* ============ 匯出員工名單.xlsx(放在 排班班表\公司) ============ */

/** 把二維陣列(第一列是標題)做成單一分頁的 .xlsx;全部存成文字,員編不會被改成數字。 */
function buildXlsxBlob_(sheetName, rows, fileName) {
  var cols = function (i) { var c = ''; i++; while (i > 0) { var m = (i - 1) % 26; c = String.fromCharCode(65 + m) + c; i = Math.floor((i - 1) / 26); } return c; };
  var body = rows.map(function (r, ri) {
    return '<row r="' + (ri + 1) + '">' + r.map(function (v, ci) {
      return '<c r="' + cols(ci) + (ri + 1) + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEsc_(v) + '</t></is></c>';
    }).join('') + '</row>';
  }).join('');
  var widths = rows[0].map(function (_, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="14" customWidth="1"/>'; }).join('');
  var head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  var parts = [
    Utilities.newBlob(head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>', 'application/xml', '[Content_Types].xml'),
    Utilities.newBlob(head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>', 'application/xml', '_rels/.rels'),
    Utilities.newBlob(head + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + xmlEsc_(sheetName) + '" sheetId="1" r:id="rId1"/></sheets></workbook>', 'application/xml', 'xl/workbook.xml'),
    Utilities.newBlob(head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>', 'application/xml', 'xl/_rels/workbook.xml.rels'),
    Utilities.newBlob(head + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>' + widths + '</cols><sheetData>' + body + '</sheetData></worksheet>', 'application/xml', 'xl/worksheets/sheet1.xml')
  ];
  return Utilities.zip(parts, fileName).setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}

/** 把「員工」分頁匯出成 排班班表\公司\員工名單.xlsx(不含密碼雜湊);檔案已存在就直接更新內容。在編輯器手動執行即可。 */
function exportEmployeesXlsx() {
  var keep = ['員編', '姓名', '部門', '職稱', '角色', '主管員編', '到職日', '狀態', 'Email', '電話'];
  var sh = sheet_(SHEETS.EMP), n = sh.getLastRow();
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idx = keep.map(function (k) { return hdr.indexOf(k); });
  var vals = n > 1 ? sh.getRange(2, 1, n - 1, hdr.length).getValues() : [];
  var rows = [keep];
  vals.forEach(function (r) {
    if (!String(r[0]).trim()) return;
    rows.push(idx.map(function (i) { var v = i < 0 ? '' : r[i]; return v instanceof Date ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd') : v; }));
  });
  var dir = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var name = '員工名單.xlsx', blob = buildXlsxBlob_('員工名單', rows, name);
  var it = dir.getFilesByName(name);
  if (it.hasNext()) {
    var f = it.next();
    var mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + f.getId() + '?uploadType=media', {
      method: 'patch', contentType: mime, payload: blob.getBytes(), headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('更新員工名單.xlsx 失敗(' + res.getResponseCode() + ')');
  } else {
    dir.createFile(blob.setName(name));
  }
  Logger.log('已匯出 ' + (rows.length - 1) + ' 位員工到 公司/' + name);
  return rows.length - 1;
}

/* ============ 取消紀錄:寫入分頁 + 追加到雲端硬碟「公司」資料夾的 取消排名.xlsx ============ */

var CANCEL_XLSX = '取消排名.xlsx';
var CANCEL_XLSX_HEAD = ['來源', '取消時間', '員工姓名', '取消日期', '班別', '取消理由', '登錄IP', '登入時間'];

function recordCancel_(u, date, unit, shift, reason, source) {
  var login = lastLogin_(u['員編']) || {};
  var ip = login.ip || '未知', ltime = login.time ? String(login.time) : '';
  var t = now_();
  var sh = sheet_(SHEETS.CANCEL);
  if (!sh.getRange(1, 8).getValue()) sh.getRange(1, 8, 1, 3).setValues([['來源', '登入IP', '登入時間']]).setFontWeight('bold');
  sh.appendRow([t, date, u['員編'], u['姓名'], unit, shift, reason, source, ip, ltime]);
  try {
    appendCancelXlsx_([source, t, u['姓名'], date, unit + ' ' + shift, reason, ip, ltime]);
  } catch (e) { Logger.log('取消排名.xlsx 寫入失敗:' + e.message); } // 失敗不影響取消本身,紀錄仍留在「取消紀錄」分頁
}

function xmlEsc_(v) { return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, ''); }

function cancelRowXml_(n, vals) {
  var cols = 'ABCDEFGH';
  return '<row r="' + n + '">' + vals.map(function (v, i) {
    return '<c r="' + cols[i] + n + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEsc_(v) + '</t></is></c>';
  }).join('') + '</row>';
}

/** 全新的 取消排名.xlsx(只有標題列) */
function newCancelXlsxBlob_() {
  var mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var parts = [
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>', 'application/xml', '[Content_Types].xml'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>', 'application/xml', '_rels/.rels'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="取消排名" sheetId="1" r:id="rId1"/></sheets></workbook>', 'application/xml', 'xl/workbook.xml'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>', 'application/xml', 'xl/_rels/workbook.xml.rels'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="12" customWidth="1"/><col min="2" max="2" width="20" customWidth="1"/><col min="3" max="3" width="12" customWidth="1"/><col min="4" max="4" width="12" customWidth="1"/><col min="5" max="5" width="18" customWidth="1"/><col min="6" max="6" width="40" customWidth="1"/><col min="7" max="7" width="18" customWidth="1"/><col min="8" max="8" width="20" customWidth="1"/></cols><sheetData>' + cancelRowXml_(1, CANCEL_XLSX_HEAD) + '</sheetData></worksheet>', 'application/xml', 'xl/worksheets/sheet1.xml')
  ];
  return Utilities.zip(parts, CANCEL_XLSX).setContentType(mime);
}

/** 在 排班班表\公司\取消排名.xlsx 最後面追加一列;檔案不存在就建立。 */
function appendCancelXlsx_(vals) {
  var mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var dir = DriveApp.getFolderById(DRIVE_FOLDER_ID); // 這個資料夾就是 排班班表\公司,取消排名.xlsx 直接放在裡面
  var it = dir.getFilesByName(CANCEL_XLSX);
  if (!it.hasNext()) {
    var blob = newCancelXlsxBlob_();
    var f = dir.createFile(blob.setName(CANCEL_XLSX));
    it = dir.getFilesByName(CANCEL_XLSX);
  }
  var file = it.next();
  var parts = Utilities.unzip(file.getBlob().setContentType('application/zip'));
  var pi = -1;
  parts.forEach(function (b, i) { if (b.getName() === 'xl/worksheets/sheet1.xml') pi = i; });
  if (pi < 0) throw new Error('取消排名.xlsx 結構異常(找不到 sheet1)');
  var xml = parts[pi].getDataAsString('UTF-8');
  var last = 0, m, re = /<row [^>]*?r="(\d+)"/g;
  while ((m = re.exec(xml))) last = Math.max(last, Number(m[1]));
  var rowXml = cancelRowXml_(last + 1, vals);
  if (/<sheetData\s*\/>/.test(xml)) xml = xml.replace(/<sheetData\s*\/>/, '<sheetData>' + rowXml + '</sheetData>');
  else xml = xml.replace('</sheetData>', rowXml + '</sheetData>');
  xml = xml.replace(/<dimension [^>]*\/>/, '');
  parts[pi] = Utilities.newBlob(xml, 'application/xml', 'xl/worksheets/sheet1.xml');
  var out = Utilities.zip(parts, CANCEL_XLSX).setContentType(mime);
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + file.getId() + '?uploadType=media', {
    method: 'patch', contentType: mime, payload: out.getBytes(),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('更新取消排名.xlsx 失敗(' + res.getResponseCode() + ')');
}

/** 管理員:取消紀錄(含原因)。p = {month} */
function cancelLog_(u, p) {
  requireAdmin_(u);
  var month = String(p.month || '');
  return readObjs_(SHEETS.CANCEL).filter(function (r) { return !month || String(r['日期']).slice(0, 7) === month; })
    .map(function (r) { return { time: r['時間'], date: r['日期'], empId: r['員編'], name: r['姓名'], unit: r['單位'], shift: r['班別'], reason: r['原因'] }; }).reverse();
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
  // 班表上已經沒有缺額的班(補滿了)→ 名額歸 0 移除,網頁上的「尚餘缺額」才會和班表一致
  var keep = {};
  parsed.slots.forEach(function (q) { keep[q.date + '|' + q.shift] = true; });
  var stale = readObjs_(SHEETS.SLOT).filter(function (r) { return r['單位'] === unit && String(r['日期']).slice(0, 7) === month && !keep[r['日期'] + '|' + r['班別']]; })
    .map(function (r) { return { date: String(r['日期']), unit: unit, shift: r['班別'], quota: 0 }; });
  for (var k = 0; k < stale.length; k += 250) signupQuotaSave_(admin, { items: stale.slice(k, k + 250) });
  for (var i = 0; i < items.length; i += 400) shiftsSave_(admin, { items: items.slice(i, i + 400) });
  for (var j = 0; j < parsed.slots.length; j += 250) signupQuotaSave_(admin, { items: parsed.slots.slice(j, j + 250).map(function (q) { return { date: q.date, unit: unit, shift: q.shift, quota: q.quota }; }) });
  var pr = pruneRemovedShifts_(unit, month, parsed.shifts); // 班表上被手動拿掉的人,網站的排班/報名也一併移除
  return { unit: unit, month: month, sheet: parsed.sheetName, shifts: items.length, slots: parsed.slots.length, unmatched: Object.keys(unmatched), pruned: pr.removed, pruneNote: pr.note || '' };
}

/** 班表(Drive)為準:本月「明天以後」的班,網站「排班」有、班表上卻已經沒有這個名字的 → 刪掉排班,並刪掉同人同日同單位的報名紀錄。
 *  名額不在這裡加減(同步時名額本來就依班表重算)。保護:班表讀不到任何人名、或要刪的超過 30 筆(可能是班表格式異常)就整個略過。 */
function pruneRemovedShifts_(unit, month, sheetShifts) {
  if (!sheetShifts.length) return { removed: [], note: '班表上讀不到任何人名,略過清理' };
  var today = today_(), present = {};
  sheetShifts.forEach(function (x) { present[x.date + '|' + x.name] = true; });
  var sh = sheet_(SHEETS.SHIFT), n = sh.getLastRow();
  if (n < 2) return { removed: [] };
  var vals = sh.getRange(2, 1, n - 1, 6).getValues(), cand = [];
  vals.forEach(function (r, i) {
    var d = dstr_(r[0]);
    if (d > today && d.slice(0, 7) === month && String(r[3]).indexOf(unit) === 0 && !present[d + '|' + String(r[2]).trim()]) cand.push({ row: i + 2, date: d, id: String(r[1]), name: String(r[2]) });
  });
  if (!cand.length) return { removed: [] };
  if (cand.length > 30) return { removed: [], note: '要清理的有 ' + cand.length + ' 筆,超過 30 筆,可能是班表格式異常,已略過' };
  for (var i = cand.length - 1; i >= 0; i--) sh.deleteRow(cand[i].row); // 由後往前刪,列號才不會跑掉
  var ss = sheet_(SHEETS.SIGNUP), m = ss.getLastRow();
  if (m > 1) {
    var rows = ss.getRange(2, 1, m - 1, 7).getValues();
    for (var k = rows.length - 1; k >= 0; k--) {
      var rd = dstr_(rows[k][0]), rid = String(rows[k][1]);
      if (rows[k][3] === unit && cand.some(function (c) { return c.date === rd && c.id === rid; })) ss.deleteRow(k + 2);
    }
  }
  return { removed: cand.map(function (c) { return c.date + ' ' + c.name; }) };
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
  return { name: sheetName, rows: rows, path: target };
}

/** 解析班表分頁:每週一組「日期」列,下面是各班別的人名列,最後是需求/預計支援人數列。
 *  名額 = 缺額 = 需求人數 − 已排人數(同一天同一類班別,最少 0);缺額 0 的班不開放報名。 */
function parseScheduleRows_(rows, tz, shiftOf) {
  var shifts = [], demand = {}, assigned = {}, gap = {}, avail = {}, dates = null;
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
    var kindOf = function (ch) { return ch === '晚' || ch === '夜' ? '夜' : '早'; };
    var gm = label.match(/^缺額[:：]?\s*(早|中|晚|夜)/) || (/^目前尚餘缺額/.test(label) ? [label, '早'] : null); // 鷹世服的列名是「目前尚餘缺額」(只有早班)          // 「缺額」列 = 可出勤人數 − 需求人數,負數代表還缺幾個人
    if (gm) {
      for (var c3 = 1; c3 < r.length; c3++) {
        var d3 = dates[c3 - 1], g = r[c3];
        if (d3 && typeof g === 'number') gap[d3 + '|' + kindOf(gm[1])] = g;
      }
      return;
    }
    var am = label.match(/^(早|中|晚|夜)班可出勤人數/);
    if (am) {
      for (var c4 = 1; c4 < r.length; c4++) {
        var d4 = dates[c4 - 1];
        if (d4 && typeof r[c4] === 'number') avail[d4 + '|' + kindOf(am[1])] = r[c4];
      }
      return;
    }
    var dm = label.match(/^(早|中|晚|夜)班(需求|預計)支援人數/);
    if (dm) {
      for (var c2 = 1; c2 < r.length; c2++) {
        var d2 = dates[c2 - 1], n = Number(r[c2]);
        if (d2 && n > 0) demand[d2 + '|' + kindOf(dm[1])] = { date: d2, kind: kindOf(dm[1]), need: Math.floor(n) };
      }
    }
  });
  // 尚餘缺額:優先用班表的「缺額」列(負數 = 少幾個人);沒有「缺額」列的單位(鷹世服)用「可出勤人數 − 需求人數」,公式相同
  var slots = [], seenKey = {};
  Object.keys(gap).concat(Object.keys(demand)).forEach(function (k) {
    if (seenKey[k]) return; seenKey[k] = true;
    var dd = demand[k], short;
    if (k in gap) short = -gap[k];
    else if (dd && k in avail) short = dd.need - avail[k];
    else return;
    if (short <= 0) return;
    var pr = k.split('|');
    slots.push({ date: pr[0], kind: pr[1], need: dd ? dd.need : null, assigned: avail[k] != null ? avail[k] : null, quota: Math.floor(short), shift: shiftOf(pr[1] === '夜' ? '夜 20:00-5:00' : '早 9:00-18:00') });
  });
  return { shifts: shifts, slots: slots };
}

/* ============ 報名寫回雲端硬碟班表 .xlsx ============ */

function xlsxColLetter_(c) { return String.fromCharCode(65 + c); } // c: 0=A … 25=Z(班表只用到 A–I)
function xlsxColIndex_(letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64); return n - 1; }

function xlsxMergedRanges_(xml) {
  var out = [], re = /<mergeCell\s+ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g, m;
  while ((m = re.exec(xml))) out.push({ c1: xlsxColIndex_(m[1]), r1: Number(m[2]), c2: xlsxColIndex_(m[3]), r2: Number(m[4]) });
  return out;
}

function xlsxCellRe_(ref) { return new RegExp('<c\\b([^>]*?\\sr="' + ref + '"[^>]*?)(?:/>|>([\\s\\S]*?)</c>)'); }

/** 取得儲存格原始 XML(找不到回傳 '') */
/** 公式儲存格的快取值 ± delta(Drive 上的 .xlsx 不會自己重算,不更新的話「可出勤人數」「缺額」會是舊的) */
function xlsxBumpCache_(xml, ref, delta) {
  var re = xlsxCellRe_(ref), m = xml.match(re);
  if (!m || !m[2]) return xml;
  var vm = m[2].match(/<v>(-?\d+(?:\.\d+)?)<\/v>/);
  if (!vm) return xml;
  return xml.replace(re, function (all) { return all.replace(/<v>-?\d+(?:\.\d+)?<\/v>/, '<v>' + (Number(vm[1]) + delta) + '</v>'); });
}

function xlsxGetCell_(xml, ref) { var m = xml.match(xlsxCellRe_(ref)); return m ? m[0] : ''; }

/** 把儲存格改成文字(text 為空字串代表清空),保留原本的樣式 s;儲存格不存在就依欄位順序插入該列。 */
function xlsxSetCell_(xml, ref, text) {
  var esc = String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  function build(style) {
    var st = style ? ' s="' + style + '"' : '';
    return text === '' ? '<c r="' + ref + '"' + st + '/>' : '<c r="' + ref + '"' + st + ' t="inlineStr"><is><t xml:space="preserve">' + esc + '</t></is></c>';
  }
  var re = xlsxCellRe_(ref), m = xml.match(re);
  if (m) {
    var sm = m[1].match(/\ss="(\d+)"/);
    return xml.replace(re, function () { return build(sm ? sm[1] : ''); });
  }
  var col = ref.match(/^[A-Z]+/)[0], rn = ref.slice(col.length), ci = xlsxColIndex_(col);
  var rowRe = new RegExp('<row\\b([^>]*?\\sr="' + rn + '"[^>]*?)(/>|>([\\s\\S]*?)</row>)');
  var rm = xml.match(rowRe);
  if (!rm) throw new Error('班表檔找不到第 ' + rn + ' 列');
  var inner = rm[3] || '', ins = inner.length, cre = /<c\b[^>]*?\sr="([A-Z]+)\d+"/g, cm;
  while ((cm = cre.exec(inner))) { if (xlsxColIndex_(cm[1]) > ci) { ins = cm.index; break; } }
  var newInner = inner.slice(0, ins) + build('') + inner.slice(ins);
  var attrs = rm[1].replace(/\/$/, '');
  return xml.replace(rowRe, function () { return '<row ' + attrs.replace(/^\s+/, '') + '>' + newInner + '</row>'; });
}

/** 在班表列(二維陣列)裡找要寫入/清除的儲存格。
 *  新增:同一週區塊、該日期那一欄、同類班別(早/夜)列中「第一個空白格」;該日期任何班別已有同名就丟錯。
 *  清除(remove):找到該名字所在格。回傳 {ref} 或 null(清除時找不到)。 */
function planScheduleCell_(rows, merged, date, kind, name, remove, isFree) {
  var col = 0, hit = null, dup = false, lastMatch = null;
  function ds(v) {
    if (typeof v === 'number' && v > 30000 && v < 80000) { var d = new Date(Math.floor(v - 25569) * 86400000); return d.toISOString().slice(0, 10); }
    var m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? m[0] : '';
  }
  function isMerged(c, rn) { return merged.some(function (g) { return c >= g.c1 && c <= g.c2 && rn >= g.r1 && rn <= g.r2 && !(c === g.c1 && rn === g.r1); }); }
  rows.forEach(function (r, i) {
    var label = String(r[0] == null ? '' : r[0]).trim();
    if (label === '日期') { var idx = r.slice(1).map(ds).indexOf(date); col = idx < 0 ? 0 : idx + 1; return; }
    if (!col) return;
    var m = label.match(/^(早|中|晚|夜)\s*\d{1,2}:\d{2}-\d{1,2}:\d{2}$/);
    if (!m) return;
    var k = m[1] === '夜' || m[1] === '晚' ? '夜' : '早';
    var cell = r[col], txt = String(cell == null ? '' : cell).trim(), rn = i + 1, ref = xlsxColLetter_(col) + rn;
    if (txt && txt.indexOf(name) === 0) { dup = true; if (k === kind) lastMatch = ref; }
    if (!remove && !hit && k === kind && !txt && !isMerged(col, rn) && (!isFree || isFree(ref))) hit = ref;
  });
  if (remove) return lastMatch ? { ref: lastMatch } : null;
  if (dup) throw new Error('班表上 ' + date + ' 已經有「' + name + '」,無法重複報名');
  if (!hit) throw new Error('班表 ' + date + ' 的' + (kind === '夜' ? '夜班' : '早班') + '找不到空白格,無法寫入(請聯絡人事確認班表)');
  return { ref: hit };
}

/** 把報名者的名字寫進(或移出)雲端硬碟班表 .xlsx 該日期、該班別的第一個空白格。
 *  寫入前一定先備份原檔到「班表備份」資料夾;找不到空白格就丟錯,不會覆蓋任何人。
 *  呼叫端(doPost)已持有 Script Lock,所以與定時同步不會同時執行。 */
function writeBackSchedule_(unit, date, shift, name, remove) {
  var cfg = SYNC_UNITS[unit];
  if (!cfg) return;
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var it = folder.getFilesByName(cfg.file);
  if (!it.hasNext()) throw new Error('雲端硬碟找不到「' + cfg.file + '」,無法寫入班表');
  var file = it.next();
  var want = date.slice(0, 7).replace('-', '');
  var sheet = readXlsxSheet_(file.getBlob(), function (n) { return /^班表/.test(n) && n.replace(/\D/g, '') === want; }, 9);
  if (!sheet) throw new Error('班表檔裡沒有 ' + date.slice(0, 7) + ' 的分頁,無法寫入');
  var parts = Utilities.unzip(file.getBlob().setContentType('application/zip'));
  var pi = -1;
  parts.forEach(function (b, i) { if (b.getName() === sheet.path) pi = i; });
  if (pi < 0) throw new Error('班表檔結構異常,找不到分頁檔');
  var xml = parts[pi].getDataAsString('UTF-8');
  var kind = /夜/.test(shift) ? '夜' : '早';
  var plan = planScheduleCell_(sheet.rows, xlsxMergedRanges_(xml), date, kind, name, remove, function (ref) { return !/<f[\s>]/.test(xlsxGetCell_(xml, ref)); });
  if (!plan) return; // 取消時班表上已找不到這個名字,不用處理
  var newXml = xlsxSetCell_(xml, plan.ref, remove ? '' : name);
  // 同步更新該區塊下方「可出勤人數」「缺額」兩列的快取值,讓班表和網頁的尚餘缺額維持一致
  var cl = plan.ref.match(/^([A-Z]+)(\d+)$/), delta = remove ? -1 : 1, hasAvail = false, hasGap = false;
  for (var ri = Number(cl[2]); ri < sheet.rows.length; ri++) {
    var lb = String(sheet.rows[ri][0] == null ? '' : sheet.rows[ri][0]).trim();
    if (lb === '日期') break;
    if (!hasAvail && /可出勤人數/.test(lb)) { newXml = xlsxBumpCache_(newXml, cl[1] + (ri + 1), delta); hasAvail = true; }
    else if (!hasGap && /缺額/.test(lb)) { newXml = xlsxBumpCache_(newXml, cl[1] + (ri + 1), delta); hasGap = true; } // 鷹世服的列名是「目前尚餘缺額」,寶雅是「缺額」
    if (hasAvail && hasGap) break;
  }

  // 備份原檔
  var bf = folder.getFoldersByName('班表備份');
  var bdir = bf.hasNext() ? bf.next() : folder.createFolder('班表備份');
  file.makeCopy('班表備份_' + cfg.file.replace(/\.xlsx$/, '') + '_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmmss') + '.xlsx', bdir);

  parts[pi] = Utilities.newBlob(newXml, 'application/xml', sheet.path);
  var mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var out = Utilities.zip(parts, file.getName()).setContentType(mime);
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + file.getId() + '?uploadType=media', {
    method: 'patch', contentType: mime, payload: out.getBytes(),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('寫入班表失敗(' + res.getResponseCode() + '):' + res.getContentText().slice(0, 200));
  return plan.ref;
}

/* ============ 新人註冊(上傳證件 → 管理員審核 → 建立帳號) ============ */

var REG_FOLDER = '人事註冊證件';
var REG_FILES = { front: { col: 8, label: '身分證正面' }, back: { col: 9, label: '身分證反面' }, bank: { col: 10, label: '存摺封面' } };

/** 證件放在擁有者「我的雲端硬碟」根目錄的私人資料夾,不分享、不公開 */
function regFolder_() { // 放在「排班班表」資料夾底下(「公司」資料夾的上一層);舊的放在雲端硬碟根目錄的會自動搬過來
  var parent = DriveApp.getFolderById(DRIVE_FOLDER_ID).getParents().next();
  var it = parent.getFoldersByName(REG_FOLDER);
  if (it.hasNext()) return it.next();
  var old = DriveApp.getRootFolder().getFoldersByName(REG_FOLDER);
  if (old.hasNext()) { var f = old.next(); f.moveTo(parent); return f; }
  return parent.createFolder(REG_FOLDER);
}

/** 試算表會把電話存成數字而吃掉開頭的 0,讀出來時補回 */
function phoneFix_(v) {
  var t = String(v == null ? '' : v).replace(/\D/g, '');
  return /^9\d{8}$/.test(t) ? '0' + t : String(v == null ? '' : v);
}
/** 一次性:把已存的註冊申請與對應員工的電話補回開頭的 0,並改存成文字 */
function regFixPhones_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('REG_PHONE_FIXED')) return;
  var g = regRows_(), esh = sheet_(SHEETS.EMP), en = esh.getLastRow();
  var eids = en > 1 ? esh.getRange(2, 1, en - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  g.rows.forEach(function (r, i) {
    var fixed = phoneFix_(r[4]);
    if (fixed !== String(r[4])) g.sh.getRange(i + 2, 5).setNumberFormat('@').setValue(fixed);
    var ei = r[15] ? eids.indexOf(String(r[15])) : -1;
    if (ei >= 0) esh.getRange(ei + 2, 11).setNumberFormat('@').setValue(fixed);
  });
  props.setProperty('REG_PHONE_FIXED', '1');
}

/** 證件依「公司 → 姓名」分資料夾,檔名:公司-姓名-證件名稱 */
function regSub_(parent, name) {
  name = String(name || '未填').replace(/[\\\/]/g, '_');
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function regPersonFolder_(unit, name) { return regSub_(regSub_(regFolder_(), unit), name); }
function regFileName_(unit, name, key, ext) { return unit + '-' + name + '-' + REG_FILES[key].label + (ext ? '.' + ext : ''); }
/** 把已上傳的舊證件整理成新的資料夾與檔名(只做一次,管理員開啟註冊審核時自動執行) */
function regMigrate_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('REG_MIGRATED_V2')) return;
  var g = regRows_();
  g.rows.forEach(function (r) {
    var name = String(r[1]), unit = String(r[6]), dest = null;
    Object.keys(REG_FILES).forEach(function (k) {
      var id = r[REG_FILES[k].col - 1];
      if (!id) return;
      try {
        var f = DriveApp.getFileById(id);
        if (!dest) dest = regPersonFolder_(unit, name);
        var m = f.getName().match(/\.([A-Za-z0-9]+)$/);
        f.setName(regFileName_(unit, name, k, m ? m[1] : ''));
        f.moveTo(dest);
      } catch (e) { Logger.log('整理證件失敗 ' + id + ':' + e.message); }
    });
  });
  props.setProperty('REG_MIGRATED_V2', '1');
}

function validTwId_(s) {
  if (!/^[A-Z][12]\d{8}$/.test(s)) return /^[A-Z][A-D89]\d{8}$/.test(s); // 居留證新舊格式只驗格式
  var map = { A: 10, B: 11, C: 12, D: 13, E: 14, F: 15, G: 16, H: 17, I: 34, J: 18, K: 19, L: 20, M: 21, N: 22, O: 35, P: 23, Q: 24, R: 25, S: 26, T: 27, U: 28, V: 29, W: 32, X: 30, Y: 31, Z: 33 };
  var v = map[s.charAt(0)], sum = Math.floor(v / 10) + (v % 10) * 9;
  for (var i = 1; i <= 8; i++) sum += Number(s.charAt(i)) * (9 - i);
  return (sum + Number(s.charAt(9))) % 10 === 0;
}

function rocPwd_(iso) { // 1990-05-18 → 0790518(試算表可能把日期存成 Date,先轉成文字)
  iso = dstr_(iso);
  var m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return ('00' + (Number(m[1]) - 1911)).slice(-3) + m[2] + m[3];
}

function checkRate_(key, max, ttl) {
  var cache = CacheService.getScriptCache(), n = Number(cache.get(key) || 0);
  if (n >= max) throw new Error('操作太頻繁,請稍後再試');
  cache.put(key, String(n + 1), ttl);
}

function fileExt_(bytes, mime) {
  var h = bytes.slice(0, 4).map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  if (h.indexOf('ffd8ff') === 0) return { ext: 'jpg', mime: 'image/jpeg' };
  return null;
}

/** 公開:新人送出申請。p = {name, idNo, birthday(yyyy-mm-dd), phone, email, unit, consent, files:{front,back,bank:{data(base64)}}} */
/** 班表檔裡的「黑名單」分頁(寶雅:黑名單-寶雅;鷹世服:黑名單),回傳 {姓名: true, 信箱: true};讀不到就回傳 null */
var BLACKLIST_SHEET = { '寶雅': '黑名單-寶雅', '鷹世服': '黑名單' };
function blacklistSet_(unit) {
  var cfg = SYNC_UNITS[unit], want = BLACKLIST_SHEET[unit];
  if (!cfg || !want) return {};
  var it = DriveApp.getFolderById(DRIVE_FOLDER_ID).getFilesByName(cfg.file);
  if (!it.hasNext()) return null;
  var f = it.next(), key = 'bl2_' + unit + '_' + f.getLastUpdated().getTime(), cache = CacheService.getScriptCache();
  var hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  var sh = readXlsxSheet_(f.getBlob(), function (n) { return n.trim() === want; }, 4);
  var set = {};
  if (sh) sh.rows.forEach(function (r, i) {
    if (i === 0) return; // 標題列
    var nm = String(r[1] || '').replace(/[\s\u3000]/g, '');
    if (nm) set[nm] = true;
    var em = String(r[2] || '').trim().toLowerCase();
    if (/@/.test(em)) set[em] = true;
  });
  try { cache.put(key, JSON.stringify(set), 21600); } catch (e) { /* 太大就不快取 */ }
  return set;
}
function isBlacklisted_(unit, name, email) {
  var set;
  try { set = blacklistSet_(unit); } catch (e) { Logger.log('黑名單讀取失敗:' + e.message); return false; }
  if (!set) return false;
  return !!(set[String(name).replace(/[\s\u3000]/g, '')] || set[String(email).trim().toLowerCase()]);
}

function registerSubmit_(p) {
  checkRate_('reg_all', 40, 3600); // 全站每小時最多 40 件,避免被灌爆雲端硬碟
  var name = String(p.name || '').trim(), idNo = String(p.idNo || '').trim().toUpperCase();
  var birthday = String(p.birthday || '').trim(), phone = String(p.phone || '').replace(/[-\s]/g, ''), email = String(p.email || '').trim();
  var unit = String(p.unit || '').trim();
  if (!p.consent) throw new Error('請先勾選同意個人資料蒐集告知事項');
  if (!/^[^\d\s]{2,20}$/.test(name)) throw new Error('請輸入正確的姓名');
  if (!validTwId_(idNo)) throw new Error('身分證字號格式不正確');
  var bd = birthday.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!bd || isNaN(new Date(birthday).getTime()) || birthday > today_() || Number(bd[1]) < 1920) throw new Error('生日格式不正確');
  if (!/^09\d{8}$/.test(phone)) throw new Error('手機號碼請輸入 09 開頭共 10 碼');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email 格式不正確');
  if (!SYNC_UNITS[unit]) throw new Error('請選擇申請單位');
  if (isBlacklisted_(unit, name, email)) throw new Error('很抱歉,無法受理您的申請。如有疑問請洽人資。');

  var sh = sheet_(SHEETS.REG), n = sh.getLastRow();
  var rows = n > 1 ? sh.getRange(2, 1, n - 1, 16).getValues() : [];
  rows.forEach(function (r) {
    if (String(r[2]) === idNo && (r[10] === '待審核' || r[10] === '已通過')) throw new Error(r[10] === '待審核' ? '您已送出申請,正在審核中,請用「查詢進度」確認' : '您的申請已通過,請直接登入');
  });
  var emps = readObjs_(SHEETS.EMP);
  if (emps.some(function (e) { return String(e['姓名']).trim() === name && e['密碼雜湊'] === hash_(rocPwd_(birthday)); })) throw new Error('系統已有相同姓名與生日的帳號,請直接登入或聯絡人資');

  var files = p.files || {}, saved = {}, appId = 'R' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMddHHmmss') + Math.floor(100 + Math.random() * 900);
  var created = [], folder = regPersonFolder_(unit, name);
  try {
    Object.keys(REG_FILES).forEach(function (k) {
      var f = files[k];
      if (!f || !f.data) throw new Error('請上傳「' + REG_FILES[k].label + '」');
      if (String(f.data).length > 7 * 1024 * 1024) throw new Error(REG_FILES[k].label + '檔案太大(上限約 5MB)');
      var bytes = Utilities.base64Decode(String(f.data));
      var t = fileExt_(bytes);
      if (!t) throw new Error(REG_FILES[k].label + '只接受 JPG 圖檔');
      var file = folder.createFile(Utilities.newBlob(bytes, t.mime, regFileName_(unit, name, k, t.ext)));
      created.push(file);
      saved[k] = file.getId();
    });
  } catch (err) {
    created.forEach(function (f) { try { f.setTrashed(true); } catch (x) { /* ignore */ } });
    throw err;
  }
  sh.appendRow([appId, name, idNo, birthday, phone, email, unit, saved.front, saved.back, saved.bank, '待審核', now_(), '', '', '', '']);
  sh.getRange(sh.getLastRow(), 4).setNumberFormat('@').setValue(birthday); // 生日存成文字,避免被轉成日期
  sh.getRange(sh.getLastRow(), 5).setNumberFormat('@').setValue(phone); // 電話存成文字,保留開頭的 0
  notifyAdminReg_(name, unit, appId);
  return { appId: appId };
}

/** 公開:用身分證字號 + 生日查詢審核進度(連續查錯會被限制) */
function registerStatus_(p) {
  var idNo = String(p.idNo || '').trim().toUpperCase(), birthday = String(p.birthday || '').trim();
  checkRate_('regq_' + idNo, 8, 900);
  var sh = sheet_(SHEETS.REG), n = sh.getLastRow(), hit = null;
  (n > 1 ? sh.getRange(2, 1, n - 1, 16).getValues() : []).forEach(function (r) { if (String(r[2]) === idNo && dstr_(r[3]) === birthday) hit = r; });
  if (!hit) throw new Error('查無申請資料,請確認身分證字號與生日');
  return { appId: hit[0], status: hit[10], time: hit[11], note: hit[10] === '已退件' ? hit[14] : '' };
}

function regRows_() {
  var sh = sheet_(SHEETS.REG), n = sh.getLastRow();
  return { sh: sh, rows: n > 1 ? sh.getRange(2, 1, n - 1, 16).getValues() : [] };
}

function regFind_(appId) {
  var g = regRows_(), i;
  for (i = 0; i < g.rows.length; i++) if (g.rows[i][0] === appId) return { sh: g.sh, row: i + 2, r: g.rows[i] };
  throw new Error('找不到申請: ' + appId);
}

/** 管理員:申請清單。p = {status}(空白 = 全部) */
/** 有新申請就寄信給部署者(自己)。收件人取自執行身分,程式碼裡不放信箱;信中不放身分證等個資 */
function notifyAdminReg_(name, unit, appId) {
  try {
    var to = Session.getEffectiveUser().getEmail();
    if (!to) return;
    MailApp.sendEmail(to, '【人事系統】新註冊申請待審核:' + name + '(' + unit + ')',
      '有一筆新的註冊申請等您審核。\n\n姓名:' + name + '\n單位:' + unit + '\n申請編號:' + appId + '\n時間:' + now_() +
      '\n\n請登入人事系統 → 註冊審核 處理。\n(此信由系統自動寄出)');
  } catch (e) { Logger.log('通知信寄送失敗:' + e.message); } // 寄信失敗不影響申請本身
}
/** 第一次使用:在編輯器執行此函式並按「允許」,之後新申請才會寄信 */
function authorizeMail() {
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), '【人事系統】通知信測試', '收到這封信代表新申請通知已啟用。');
}
function registerList_(u, p) {
  requireAdmin_(u);
  try { regMigrate_(); } catch (e) { Logger.log('證件整理失敗:' + e.message); }
  try { regFixPhones_(); } catch (e) { Logger.log('電話修正失敗:' + e.message); }
  var st = String(p.status || '');
  return regRows_().rows.filter(function (r) { return !st || r[10] === st; }).map(function (r) {
    return { appId: r[0], name: r[1], idNo: r[2], birthday: dstr_(r[3]), phone: phoneFix_(r[4]), email: r[5], unit: r[6], status: r[10], time: r[11], reviewer: r[12], reviewTime: r[13], note: r[14], empId: r[15],
      hasFiles: !!(r[7] || r[8] || r[9]) };
  }).reverse();
}

/** 管理員:取得證件內容供預覽。p = {appId, key: front|back|bank} */
function registerFile_(u, p) {
  requireAdmin_(u);
  var f = REG_FILES[p.key];
  if (!f) throw new Error('證件類型錯誤');
  var rec = regFind_(String(p.appId)), id = rec.r[f.col - 1];
  if (!id) throw new Error('這份證件已刪除');
  var blob = DriveApp.getFileById(id).getBlob();
  return { mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()), label: f.label };
}

/** 管理員輸入員編後通過,帳號自動建立(帳號 = 姓名,初始密碼 = 民國生日 7 碼)。員編只能由管理員輸入。p = {appId, empId} */
function registerApprove_(u, p) {
  requireAdmin_(u);
  var empId = String(p.empId || '').trim();
  if (empId && !/^[A-Za-z0-9_-]{1,20}$/.test(empId)) throw new Error('員編只能是英數字、底線、連字號(最多 20 碼)');
  var rec = regFind_(String(p.appId)), r = rec.r;
  if (r[10] !== '待審核') throw new Error('這件申請已處理過');
  if (!empId) { // 沒填員編:先給暫編 TMP001…,之後可在「員工資料」改成正式員編
    var used = {};
    readObjs_(SHEETS.EMP).forEach(function (e) { used[String(e['員編'])] = true; });
    for (var k = 1; k < 1000; k++) { empId = 'TMP' + ('00' + k).slice(-3); if (!used[empId]) break; }
  }
  if (findEmp_(empId)) throw new Error('員編 ' + empId + ' 已存在');
  var pwd = rocPwd_(r[3]);
  sheet_(SHEETS.EMP).appendRow([empId, r[1], r[6], '臨時PT', 'employee', '', today_(), hash_(pwd), '在職', r[5], phoneFix_(r[4])]);
  sheet_(SHEETS.EMP).getRange(sheet_(SHEETS.EMP).getLastRow(), 11).setNumberFormat('@').setValue(phoneFix_(r[4]));
  var hh = hash_(pwd); // 同名在職、且密碼也完全相同(生日相同)才無法用姓名分辨,才需要提醒
  var same = readObjs_(SHEETS.EMP).filter(function (e) { return String(e['姓名']).trim() === r[1] && e['狀態'] === '在職' && e['密碼雜湊'] === hh; }).length;
  rec.sh.getRange(rec.row, 11).setValue('已通過');
  rec.sh.getRange(rec.row, 13, 1, 4).setValues([[u['姓名'], now_(), '', empId]]);
  return { empId: empId, account: r[1], password: pwd, sameName: same };
}

/** 管理員:退件(要寫原因)。p = {appId, reason} */
function registerReject_(u, p) {
  requireAdmin_(u);
  var reason = String(p.reason || '').trim();
  if (!reason) throw new Error('請填退件原因');
  var rec = regFind_(String(p.appId));
  if (rec.r[10] !== '待審核') throw new Error('這件申請已處理過');
  rec.sh.getRange(rec.row, 11).setValue('已退件');
  rec.sh.getRange(rec.row, 13, 1, 3).setValues([[u['姓名'], now_(), reason]]);
  return true;
}

/** 管理員:刪除已處理申請的證件檔(移到雲端硬碟垃圾桶),個資不要留太久。p = {appId} */
function registerPurge_(u, p) {
  requireAdmin_(u);
  var rec = regFind_(String(p.appId));
  if (rec.r[10] === '待審核') throw new Error('尚未審核的申請不能刪除證件');
  [8, 9, 10].forEach(function (col) {
    var id = rec.r[col - 1];
    if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* 已不存在 */ } }
  });
  rec.sh.getRange(rec.row, 8, 1, 3).setValues([['', '', '']]);
  return true;
}

/* ============ 共用工具 ============ */

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh && (name === SHEETS.LOGIN || name === SHEETS.SLOT || name === SHEETS.SIGNUP || name === SHEETS.REG || name === SHEETS.CANCEL || name === SHEETS.PRIORITY || name === SHEETS.MFA)) { // 這幾張表第一次使用時自動建立,不需要重新執行 setup()
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
    if (name === SHEETS.SLOT || name === SHEETS.SIGNUP) sh.getRange('A:A').setNumberFormat('@');
    if (name === SHEETS.REG) sh.getRange('A:F').setNumberFormat('@');
    if (name === SHEETS.CANCEL) sh.getRange('A:C').setNumberFormat('@'); // 日期維持文字,避免被轉成日期格式
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
