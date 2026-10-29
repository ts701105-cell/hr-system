/* API 層:有設定 API_URL 就呼叫 Apps Script;否則使用瀏覽器內建示範資料 */
(function () {
  var cfg = window.HR_CONFIG || {};
  var DEMO = !cfg.API_URL;
  var TOKEN_KEY = 'hr_token';

  function getToken() { try { return sessionStorage.getItem(TOKEN_KEY) || ''; } catch (e) { return ''; } }
  function setToken(t) { try { t ? sessionStorage.setItem(TOKEN_KEY, t) : sessionStorage.removeItem(TOKEN_KEY); } catch (e) { /* ignore */ } }

  async function remote(action, params) {
    var res = await fetch(cfg.API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // 避免 CORS 預檢
      body: JSON.stringify({ action: action, token: getToken(), params: params || {} })
    });
    var j = await res.json();
    if (!j.ok) {
      var err = new Error(j.error || '發生錯誤');
      err.code = j.code;
      throw err;
    }
    return j.data;
  }

  /* ---------- 示範模式 ---------- */
  var KEY = 'hr_demo_db_v1';
  function today() { return fmtDate(new Date()); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function fmtDate(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function nowStr() { var d = new Date(); return fmtDate(d) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds()); }

  function seed() {
    var year = new Date().getFullYear();
    return {
      emps: [
        { id: 'A001', name: '王管理', dept: '人資部', title: '人事管理師', role: 'admin', manager: '', hireDate: year - 3 + '-04-01', status: '在職', email: '', phone: '', password: 'admin1234' },
        { id: 'M001', name: '陳主管', dept: '門市部', title: '店長', role: 'manager', manager: '', hireDate: year - 2 + '-06-15', status: '在職', email: '', phone: '', password: 'demo1234' },
        { id: 'E001', name: '林小明', dept: '門市部', title: '門市人員', role: 'employee', manager: 'M001', hireDate: year - 1 + '-09-01', status: '在職', email: '', phone: '', password: 'demo1234' },
        { id: 'E002', name: '張小美', dept: '門市部', title: '門市人員', role: 'employee', manager: 'M001', hireDate: year + '-01-10', status: '在職', email: '', phone: '', password: 'demo1234' }
      ],
      leaveTypes: [['特休', 10], ['事假', 14], ['病假', 30], ['家庭照顧假', 7], ['婚假', 8], ['喪假', 8], ['公假', null], ['補休', null]].map(function (a) { return { name: a[0], quota: a[1] }; }),
      leaves: [], punches: [],
      locs: [{ id: 'L001', name: '總公司(示範)', lat: 25.033, lng: 121.5654, radius: 100, active: true }],
      shiftTypes: [{ name: '早班', start: '08:00', end: '17:00' }, { name: '中班', start: '12:00', end: '21:00' }, { name: '晚班', start: '14:00', end: '23:00' }, { name: '休假', start: '', end: '' }],
      shifts: []
    };
  }
  function load() { try { var s = localStorage.getItem(KEY); if (s) return JSON.parse(s); } catch (e) { /* ignore */ } var d = seed(); save(d); return d; }
  function save(d) { try { localStorage.setItem(KEY, JSON.stringify(d)); } catch (e) { /* ignore */ } }
  function dist(a, b, c, d) {
    var R = 6371000, r = Math.PI / 180, dl = (c - a) * r, dn = (d - b) * r;
    var x = Math.sin(dl / 2) * Math.sin(dl / 2) + Math.cos(a * r) * Math.cos(c * r) * Math.sin(dn / 2) * Math.sin(dn / 2);
    return 2 * R * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x));
  }
  function pub(e) { return { id: e.id, name: e.name, dept: e.dept, title: e.title, role: e.role, manager: e.manager }; }
  function me(db) {
    var id = getToken().replace('demo:', '');
    var e = db.emps.filter(function (x) { return x.id === id && x.status === '在職'; })[0];
    if (!e) { var err = new Error('登入已逾時,請重新登入'); err.code = 'AUTH'; throw err; }
    return e;
  }
  function scope(db, u) { return db.emps.filter(function (e) { return u.role === 'admin' || e.manager === u.id; }).map(function (e) { return e.id; }); }
  function needMgr(u) { if (u.role === 'employee') throw new Error('需要主管權限'); }
  function needAdmin(u) { if (u.role !== 'admin') throw new Error('需要管理員權限'); }
  function balance(db, u) {
    var y = String(new Date().getFullYear());
    return db.leaveTypes.map(function (t) {
      var used = db.leaves.filter(function (l) { return l.empId === u.id && l.type === t.name && (l.status === '核准' || l.status === '待簽核') && l.start.slice(0, 4) === y; })
        .reduce(function (s, l) { return s + l.days; }, 0);
      return { type: t.name, quota: t.quota, used: used, remain: t.quota === null ? null : t.quota - used };
    });
  }

  var demoRoutes = {
    'me': function (db, u) { return pub(u); },
    'changePassword': function (db, u, p) {
      if (String(p.newPassword || '').length < 8) throw new Error('新密碼至少 8 碼');
      if (u.password !== p.oldPassword) throw new Error('舊密碼不正確');
      u.password = p.newPassword; return true;
    },
    'employees.list': function (db, u) { needMgr(u); var sc = scope(db, u); return db.emps.filter(function (e) { return u.role === 'admin' || sc.indexOf(e.id) >= 0; }).map(function (e) { var o = pub(e); o.hireDate = e.hireDate; o.status = e.status; o.email = e.email; o.phone = e.phone; return o; }); },
    'employees.save': function (db, u, p) {
      needAdmin(u);
      if (!p.id || !p.name) throw new Error('員編與姓名必填');
      var e = db.emps.filter(function (x) { return x.id === p.id; })[0];
      if (e) { ['name', 'dept', 'title', 'role', 'manager', 'hireDate', 'status', 'email', 'phone'].forEach(function (k) { if (p[k] !== undefined) e[k] = p[k]; }); }
      else db.emps.push({ id: p.id, name: p.name, dept: p.dept || '', title: p.title || '', role: p.role, manager: p.manager || '', hireDate: p.hireDate || '', status: p.status || '在職', email: p.email || '', phone: p.phone || '', password: p.password || '12345678' });
      return true;
    },
    'employees.resetPassword': function (db, u, p) { needAdmin(u); if (String(p.password || '').length < 8) throw new Error('密碼至少 8 碼'); var e = db.emps.filter(function (x) { return x.id === p.id; })[0]; if (!e) throw new Error('找不到員工'); e.password = p.password; return true; },
    'employees.import': function (db, u, p) {
      needAdmin(u);
      var rows = p.rows || [];
      if (!rows.length) throw new Error('沒有可匯入的資料');
      var RM = { admin: 'admin', '管理員': 'admin', '人資': 'admin', manager: 'manager', '主管': 'manager', employee: 'employee', '員工': 'employee', '': 'employee' };
      var errors = [], seen = {}, clean = [];
      rows.forEach(function (r, i) {
        var line = r.line || i + 2, id = String(r.id == null ? '' : r.id).trim(), name = String(r.name == null ? '' : r.name).trim();
        var role = RM[String(r.role == null ? '' : r.role).trim().toLowerCase()], status = String(r.status || '在職').trim(), hire = String(r.hireDate || '').trim(), pwd = String(r.password == null ? '' : r.password), err = [];
        if (!id) err.push('員編必填'); else if (!/^[A-Za-z0-9_-]{1,20}$/.test(id)) err.push('員編只能是英數字、底線、連字號(最多 20 碼)'); else if (seen[id]) err.push('員編與第 ' + seen[id] + ' 列重複');
        if (!name) err.push('姓名必填');
        if (!role) err.push('角色必須是 員工/主管/管理員');
        if (status !== '在職' && status !== '離職') err.push('狀態必須是 在職 或 離職');
        if (hire && !(/^\d{4}-\d{2}-\d{2}$/.test(hire) && new Date(hire + 'T00:00:00Z').toISOString().slice(0, 10) === hire)) err.push('到職日需為有效日期,格式 YYYY-MM-DD');
        if (pwd && pwd.length < 8) err.push('初始密碼至少 8 碼');
        if (id && id === u.id && (role !== 'admin' || status !== '在職')) err.push('不能降級或停用自己');
        if (id) seen[id] = line;
        if (err.length) errors.push({ line: line, id: id, message: err.join(';') });
        clean.push({ line: line, id: id, name: name, dept: String(r.dept || '').trim(), title: String(r.title || '').trim(), role: role, manager: String(r.manager || '').trim(), hireDate: hire, status: status, email: String(r.email || '').trim(), phone: String(r.phone || '').trim(), password: pwd });
      });
      var roleOf = {}; db.emps.forEach(function (e) { roleOf[e.id] = e.role; }); clean.forEach(function (c) { if (c.id) roleOf[c.id] = c.role || roleOf[c.id]; });
      clean.forEach(function (c) {
        if (!c.manager) return;
        if (c.manager === c.id) errors.push({ line: c.line, id: c.id, message: '主管不能是自己' });
        else if (!(c.manager in roleOf)) errors.push({ line: c.line, id: c.id, message: '找不到主管員編 ' + c.manager });
        else if (roleOf[c.manager] === 'employee') errors.push({ line: c.line, id: c.id, message: '主管員編 ' + c.manager + ' 的角色不是主管/管理員' });
      });
      if (errors.length) return { applied: false, errors: errors, created: [], updated: 0 };
      var created = [], updated = 0;
      clean.forEach(function (c) {
        var e = db.emps.filter(function (x) { return x.id === c.id; })[0];
        if (e) { ['name', 'dept', 'title', 'role', 'manager', 'hireDate', 'status', 'email', 'phone'].forEach(function (k) { e[k] = c[k]; }); if (c.password) e.password = c.password; updated++; }
        else { var pw = c.password || Math.random().toString(36).slice(2, 12); db.emps.push({ id: c.id, name: c.name, dept: c.dept, title: c.title, role: c.role, manager: c.manager, hireDate: c.hireDate, status: c.status, email: c.email, phone: c.phone, password: pw }); created.push({ id: c.id, name: c.name, password: pw }); }
      });
      return { applied: true, errors: [], created: created, updated: updated };
    },
    'locations.geocode': function () { return []; }, // 示範模式沒有後端,前端會改用 OpenStreetMap 查詢
    'leave.types': function (db) { return db.leaveTypes; },
    'leave.balance': function (db, u) { return balance(db, u); },
    'leave.apply': function (db, u, p) {
      var days = Number(p.days);
      if (!p.type || !p.start || !p.end) throw new Error('請填寫假別與日期');
      if (p.end < p.start) throw new Error('結束日期不可早於開始日期');
      if (!(days > 0) || days * 2 !== Math.round(days * 2)) throw new Error('天數必須大於 0,且以 0.5 天為單位');
      var b = balance(db, u).filter(function (x) { return x.type === p.type; })[0];
      if (!b) throw new Error('無此假別');
      if (b.remain !== null && days > b.remain) throw new Error('「' + p.type + '」剩餘 ' + b.remain + ' 天,不足申請 ' + days + ' 天');
      var id = 'LV' + Date.now();
      db.leaves.push({ id: id, empId: u.id, name: u.name, type: p.type, start: p.start, end: p.end, days: days, reason: p.reason || '', status: '待簽核', reviewer: '', comment: '', appliedAt: nowStr(), reviewedAt: '' });
      return id;
    },
    'leave.my': function (db, u) { return db.leaves.filter(function (l) { return l.empId === u.id; }).reverse(); },
    'leave.cancel': function (db, u, p) {
      var l = db.leaves.filter(function (x) { return x.id === p.id; })[0];
      if (!l || l.empId !== u.id) throw new Error('只能撤回自己的假單');
      if (l.status !== '待簽核') throw new Error('僅「待簽核」假單可撤回');
      l.status = '撤回'; return true;
    },
    'leave.pending': function (db, u) { needMgr(u); var sc = scope(db, u); return db.leaves.filter(function (l) { return l.status === '待簽核' && sc.indexOf(l.empId) >= 0 && l.empId !== u.id; }); },
    'leave.all': function (db, u, p) { needMgr(u); var sc = scope(db, u); return db.leaves.filter(function (l) { return sc.indexOf(l.empId) >= 0 && (!p.month || l.start.slice(0, 7) === p.month); }).reverse(); },
    'leave.review': function (db, u, p) {
      needMgr(u);
      if (p.decision !== '核准' && p.decision !== '駁回') throw new Error('簽核結果不正確');
      var l = db.leaves.filter(function (x) { return x.id === p.id; })[0];
      if (!l) throw new Error('找不到假單');
      if (l.status !== '待簽核') throw new Error('此假單已處理');
      if (l.empId === u.id) throw new Error('不可簽核自己的假單');
      if (scope(db, u).indexOf(l.empId) < 0) throw new Error('非您的簽核範圍');
      l.status = p.decision; l.reviewer = u.name; l.comment = p.comment || ''; l.reviewedAt = nowStr(); return true;
    },
    'punch.do': function (db, u, p) {
      var lat = Number(p.lat), lng = Number(p.lng);
      if (!isFinite(lat) || !isFinite(lng)) throw new Error('定位資料無效');
      var best = null;
      db.locs.forEach(function (l) { if (!l.active) return; var d = dist(lat, lng, l.lat, l.lng); if (!best || d < best.d) best = { d: d, l: l }; });
      if (!best) throw new Error('尚未設定任何打卡地點,請聯絡管理員');
      var ok = best.d <= best.l.radius, d = Math.round(best.d);
      db.punches.push({ time: nowStr(), date: today(), empId: u.id, name: u.name, type: p.type, distance: d, place: best.l.name, result: ok ? '成功' : '超出範圍' });
      save(db);
      if (!ok) throw new Error('距離「' + best.l.name + '」' + d + ' 公尺,超出打卡範圍(' + best.l.radius + ' 公尺)');
      return { time: nowStr(), place: best.l.name, distance: d };
    },
    'punch.my': function (db, u, p) { var m = p.month || today().slice(0, 7); return db.punches.filter(function (x) { return x.empId === u.id && x.date.slice(0, 7) === m; }).reverse(); },
    'punch.report': function (db, u, p) { needMgr(u); var m = p.month || today().slice(0, 7), sc = scope(db, u); return db.punches.filter(function (x) { return sc.indexOf(x.empId) >= 0 && x.date.slice(0, 7) === m; }).reverse(); },
    'locations.list': function (db, u) { return u.role === 'admin' ? db.locs : db.locs.filter(function (l) { return l.active; }); },
    'locations.save': function (db, u, p) {
      needAdmin(u);
      var lat = Number(p.lat), lng = Number(p.lng), r = Number(p.radius);
      if (!p.name) throw new Error('請輸入地點名稱');
      if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) throw new Error('經緯度無效');
      if (!(r >= 10 && r <= 5000)) throw new Error('半徑需介於 10~5000 公尺');
      var rec = { id: p.id || 'L' + Date.now(), name: p.name, lat: lat, lng: lng, radius: r, active: p.active !== false };
      var i = db.locs.map(function (x) { return x.id; }).indexOf(rec.id);
      if (i >= 0) db.locs[i] = rec; else db.locs.push(rec);
      return rec.id;
    },
    'locations.delete': function (db, u, p) { needAdmin(u); db.locs = db.locs.filter(function (l) { return l.id !== p.id; }); return true; },
    'shifts.types': function (db) { return db.shiftTypes; },
    'shifts.list': function (db, u, p) {
      var m = p.month || today().slice(0, 7), sc = u.role === 'employee' ? [u.id] : scope(db, u).concat([u.id]);
      return db.shifts.filter(function (s) { return s.date.slice(0, 7) === m && sc.indexOf(s.empId) >= 0; });
    },
    'shifts.save': function (db, u, p) {
      needMgr(u);
      var sc = scope(db, u);
      (p.items || []).forEach(function (it) {
        if (sc.indexOf(it.empId) < 0) throw new Error('無權限排班: ' + it.empId);
        db.shifts = db.shifts.filter(function (s) { return !(s.date === it.date && s.empId === it.empId); });
        if (!it.shift) return;
        var t = db.shiftTypes.filter(function (x) { return x.name === it.shift; })[0];
        if (!t) throw new Error('無此班別: ' + it.shift);
        var e = db.emps.filter(function (x) { return x.id === it.empId; })[0];
        db.shifts.push({ date: it.date, empId: it.empId, name: e ? e.name : '', shift: t.name, start: t.start, end: t.end });
      });
      return (p.items || []).length;
    }
  };

  async function demo(action, params) {
    var db = load();
    if (action === 'login') {
      var e = db.emps.filter(function (x) { return x.id === String(params.id || '').trim() && x.status === '在職'; })[0];
      if (!e || e.password !== params.password) throw new Error('員編或密碼錯誤');
      setToken('demo:' + e.id);
      return { token: 'demo:' + e.id, user: pub(e), mustChange: false };
    }
    var u = me(db);
    var h = demoRoutes[action];
    if (!h) throw new Error('未知的操作: ' + action);
    var out = h(db, u, params || {});
    save(db);
    return out;
  }

  window.API = {
    demo: DEMO,
    call: function (action, params) { return DEMO ? demo(action, params) : remote(action, params); },
    login: async function (id, password) {
      var r = await (DEMO ? demo('login', { id: id, password: password }) : remote('login', { id: id, password: password }));
      setToken(r.token);
      return r;
    },
    logout: function () { setToken(''); },
    hasToken: function () { return !!getToken(); }
  };
})();
