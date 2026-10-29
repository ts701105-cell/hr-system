(function () {
  'use strict';
  var cfg = window.HR_CONFIG || {};
  var app = document.getElementById('app');
  var user = null;
  var clockTimer = null;

  /* ---------- 工具 ---------- */
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function fmtDate(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function curMonth() { var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1); }
  function daysIn(month) { var p = month.split('-'); return new Date(+p[0], +p[1], 0).getDate(); }
  function toast(msg, isErr) {
    var t = document.getElementById('toast');
    t.textContent = msg; t.className = 'show' + (isErr ? ' err' : '');
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.className = ''; }, 3200);
  }
  async function call(action, params) {
    try { return await API.call(action, params); }
    catch (e) {
      if (e.code === 'AUTH') { API.logout(); user = null; renderLogin(); }
      throw e;
    }
  }
  async function guard(fn, btn) {
    if (btn) btn.disabled = true;
    try { return await fn(); }
    catch (e) { toast(e.message || String(e), true); }
    finally { if (btn) btn.disabled = false; }
  }
  function badge(status) {
    var cls = { '核准': 'ok', '成功': 'ok', '待簽核': 'warn', '駁回': 'bad', '超出範圍': 'bad', '撤回': '' }[status] || '';
    return '<span class="badge ' + cls + '">' + esc(status) + '</span>';
  }
  function table(head, rows, empty) {
    if (!rows.length) return '<div class="empty">' + esc(empty || '沒有資料') + '</div>';
    return '<div class="tablewrap"><table><thead><tr>' + head.map(function (h) { return '<th>' + esc(h) + '</th>'; }).join('') +
      '</tr></thead><tbody>' + rows.join('') + '</tbody></table></div>';
  }
  var dlg = document.getElementById('dlg');
  function openDialog(html) {
    dlg.innerHTML = html;
    if (!dlg.open) dlg.showModal();
    var c = dlg.querySelector('[data-close]');
    if (c) c.onclick = function () { dlg.close(); };
    return dlg;
  }
  function field(id, label, val, type, extra) {
    return '<div><label for="' + id + '">' + esc(label) + '</label><input id="' + id + '" type="' + (type || 'text') + '" value="' + esc(val) + '" ' + (extra || '') + '></div>';
  }
  function $(sel, root) { return (root || document).querySelector(sel); }
  function val(id) { return document.getElementById(id).value; }

  /* ---------- 登入 ---------- */
  function renderLogin() {
    clearInterval(clockTimer);
    app.innerHTML = '<div class="login card"><h1>' + esc(cfg.COMPANY_NAME || '人事管理系統') + (API.demo ? '<span class="demo-tag">示範模式</span>' : '') + '</h1>' +
      '<form id="lf" class="form" style="grid-template-columns:1fr">' + field('li', '員編', '', 'text', 'autocomplete="username" required') +
      field('lp', '密碼', '', 'password', 'autocomplete="current-password" required') +
      '<button class="primary" type="submit">登入</button></form>' +
      (API.demo ? '<p class="hint">示範帳號:管理員 A001 / admin1234;主管 M001 / demo1234;員工 E001 / demo1234。資料只存在這個瀏覽器。</p>' : '') + '</div>';
    $('#lf').onsubmit = function (e) {
      e.preventDefault();
      guard(async function () {
        var r = await API.login(val('li').trim(), val('lp'));
        user = r.user;
        location.hash = '#/punch';
        renderShell();
        if (r.mustChange) { toast('請先修改預設密碼', true); location.hash = '#/password'; }
      }, $('#lf button'));
    };
  }

  /* ---------- 外框與路由 ---------- */
  var ROUTES = [
    { path: 'punch', label: '打卡', roles: ['employee', 'manager', 'admin'], view: viewPunch },
    { path: 'leave', label: '請假', roles: ['employee', 'manager', 'admin'], view: viewLeave },
    { path: 'schedule', label: '排班', roles: ['employee', 'manager', 'admin'], view: viewSchedule },
    { path: 'approvals', label: '假單簽核', roles: ['manager', 'admin'], group: '主管', view: viewApprovals },
    { path: 'attendance', label: '出勤報表', roles: ['manager', 'admin'], group: '主管', view: viewAttendance },
    { path: 'employees', label: '員工資料', roles: ['admin'], group: '管理', view: viewEmployees },
    { path: 'locations', label: '打卡地點', roles: ['admin'], group: '管理', view: viewLocations },
    { path: 'password', label: '修改密碼', roles: ['employee', 'manager', 'admin'], group: '帳號', view: viewPassword }
  ];

  function renderShell() {
    if (!user) return renderLogin();
    var roleName = { admin: '管理員', manager: '主管', employee: '員工' }[user.role];
    app.innerHTML = '<div class="topbar"><div class="brand">' + esc(cfg.COMPANY_NAME || '人事管理系統') + (API.demo ? '<span class="demo-tag">示範</span>' : '') + '</div>' +
      '<div class="row"><span class="who">' + esc(user.name) + '(' + esc(user.id) + ' · ' + roleName + ')</span><button class="sm" id="lo">登出</button></div></div>' +
      '<div class="layout"><nav class="nav" id="nav"></nav><main id="main"></main></div>';
    $('#lo').onclick = function () { API.logout(); user = null; location.hash = ''; renderLogin(); };
    route();
  }

  function route() {
    if (!user) return;
    clearInterval(clockTimer);
    var path = (location.hash.replace(/^#\/?/, '') || 'punch');
    var allowed = ROUTES.filter(function (r) { return r.roles.indexOf(user.role) >= 0; });
    var cur = allowed.filter(function (r) { return r.path === path; })[0] || allowed[0];
    var html = '', lastGroup = null;
    allowed.forEach(function (r) {
      if (r.group && r.group !== lastGroup) { html += '<div class="sep">' + esc(r.group) + '</div>'; }
      lastGroup = r.group || lastGroup;
      html += '<a href="#/' + r.path + '" class="' + (r === cur ? 'active' : '') + '">' + esc(r.label) + '</a>';
    });
    $('#nav').innerHTML = html;
    var main = $('#main');
    main.innerHTML = '<div class="empty">載入中…</div>';
    Promise.resolve(cur.view(main)).catch(function (e) { toast(e.message || String(e), true); });
  }
  window.addEventListener('hashchange', route);

  /* ---------- 打卡 ---------- */
  function haversine(a, b, c, d) {
    var R = 6371000, r = Math.PI / 180, dl = (c - a) * r, dn = (d - b) * r;
    var x = Math.sin(dl / 2) * Math.sin(dl / 2) + Math.cos(a * r) * Math.cos(c * r) * Math.sin(dn / 2) * Math.sin(dn / 2);
    return 2 * R * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x));
  }
  function getPosition() {
    return new Promise(function (resolve, reject) {
      if (!navigator.geolocation) return reject(new Error('此瀏覽器不支援定位'));
      navigator.geolocation.getCurrentPosition(resolve, function (err) {
        reject(new Error(err.code === 1 ? '定位權限被拒絕,請在瀏覽器設定中允許定位' : '無法取得定位,請到空曠處重試'));
      }, { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 });
    });
  }

  async function viewPunch(main) {
    var today = fmtDate(new Date());
    var res = await Promise.all([call('locations.list'), call('punch.my', { month: curMonth() }), call('shifts.list', { month: curMonth() })]);
    var locs = res[0].filter(function (l) { return l.active !== false; });
    var records = res[1].filter(function (r) { return r.date === today; });
    var myShift = res[2].filter(function (s) { return s.date === today && s.empId === user.id; })[0];
    var pos = null;

    main.innerHTML = '<h2>打卡</h2><div class="card"><div class="clock" id="clk"></div>' +
      '<div class="muted">' + esc(today) + (myShift ? ' · 今日班別:' + esc(myShift.shift) + (myShift.start ? '(' + esc(myShift.start) + '–' + esc(myShift.end) + ')' : '') : ' · 今日未排班') + '</div>' +
      '<div class="geo" id="geo">尚未定位</div>' +
      '<div class="row" style="margin-top:8px"><button id="loc">取得目前位置</button></div>' +
      '<div class="punch-btns"><button class="primary" id="in" disabled>上班打卡</button><button class="ok" id="out" disabled>下班打卡</button></div>' +
      '<p class="hint">需在設定的打卡地點範圍內才能打卡。可打卡地點:' + (locs.length ? locs.map(function (l) { return esc(l.name) + '(' + l.radius + ' 公尺)'; }).join('、') : '尚未設定') + '</p></div>' +
      '<div class="card"><h3>今日打卡紀錄</h3><div id="today"></div></div>';

    function tick() { $('#clk').textContent = new Date().toLocaleTimeString('zh-TW', { hour12: false }); }
    tick(); clockTimer = setInterval(tick, 1000);

    function showToday() {
      $('#today').innerHTML = table(['時間', '類型', '地點', '距離', '結果'], records.map(function (r) {
        return '<tr><td>' + esc(r.time.slice(11)) + '</td><td>' + esc(r.type) + '</td><td>' + esc(r.place) + '</td><td>' + esc(r.distance) + ' m</td><td>' + badge(r.result) + '</td></tr>';
      }), '今天還沒有打卡紀錄');
    }
    showToday();

    async function locate() {
      var g = $('#geo'); g.textContent = '定位中…';
      try {
        var p = await getPosition();
        pos = { lat: p.coords.latitude, lng: p.coords.longitude, acc: p.coords.accuracy };
        var near = null;
        locs.forEach(function (l) { var d = haversine(pos.lat, pos.lng, l.lat, l.lng); if (!near || d < near.d) near = { d: d, l: l }; });
        if (near) {
          var inRange = near.d <= near.l.radius;
          g.innerHTML = '定位精準度約 ' + Math.round(pos.acc) + ' 公尺。距離「' + esc(near.l.name) + '」' + Math.round(near.d) + ' 公尺 — ' +
            (inRange ? '<span class="badge ok">在範圍內</span>' : '<span class="badge bad">超出範圍</span>');
        } else g.textContent = '尚未設定打卡地點';
        $('#in').disabled = $('#out').disabled = false;
      } catch (e) { g.textContent = e.message; $('#in').disabled = $('#out').disabled = true; }
    }
    $('#loc').onclick = locate;
    locate();

    async function punch(type, btn) {
      await guard(async function () {
        var p = await getPosition(); // 打卡當下重新定位
        var r = await call('punch.do', { type: type, lat: p.coords.latitude, lng: p.coords.longitude });
        records.unshift({ time: r.time, type: type, place: r.place, distance: r.distance, result: '成功' });
        showToday(); toast(type + '打卡成功(' + r.place + ')');
      }, btn);
      // 失敗也會被記錄,重新抓取
      try { records = (await call('punch.my', { month: curMonth() })).filter(function (r) { return r.date === today; }); showToday(); } catch (e) { /* ignore */ }
    }
    $('#in').onclick = function () { punch('上班', this); };
    $('#out').onclick = function () { punch('下班', this); };
  }

  /* ---------- 請假 ---------- */
  async function viewLeave(main) {
    var res = await Promise.all([call('leave.balance'), call('leave.my')]);
    var bal = res[0], list = res[1];
    var types = bal.map(function (b) { return b.type; });

    main.innerHTML = '<h2>請假</h2><div class="card"><h3>假別餘額(' + new Date().getFullYear() + ' 年)</h3><div class="grid">' +
      bal.map(function (b) {
        return '<div class="stat"><div class="n">' + (b.remain === null ? '不限' : b.remain) + '</div><div class="l">' + esc(b.type) + '(已用 ' + b.used + (b.quota === null ? '' : ' / ' + b.quota) + ')</div></div>';
      }).join('') + '</div></div>' +
      '<div class="card"><h3>申請請假</h3><form id="lf" class="form">' +
      '<div><label for="lt">假別</label><select id="lt">' + types.map(function (t) { return '<option>' + esc(t) + '</option>'; }).join('') + '</select></div>' +
      field('ls', '開始日期', fmtDate(new Date()), 'date', 'required') + field('le', '結束日期', fmtDate(new Date()), 'date', 'required') +
      field('ld', '天數(0.5 為單位)', '1', 'number', 'min="0.5" step="0.5" required') +
      '<div class="full"><label for="lr">事由</label><textarea id="lr" rows="2"></textarea></div>' +
      '<div class="full"><button class="primary" type="submit">送出申請</button></div></form></div>' +
      '<div class="card"><h3>我的假單</h3><div id="ll"></div></div>';

    function autoDays() {
      var a = new Date(val('ls')), b = new Date(val('le'));
      if (!isNaN(a) && !isNaN(b) && b >= a) document.getElementById('ld').value = Math.round((b - a) / 86400000) + 1;
    }
    $('#ls').onchange = autoDays; $('#le').onchange = autoDays;

    function drawList() {
      $('#ll').innerHTML = table(['單號', '假別', '期間', '天數', '事由', '狀態', '簽核', ''], list.map(function (l) {
        return '<tr><td>' + esc(l.id) + '</td><td>' + esc(l.type) + '</td><td>' + esc(l.start) + ' ~ ' + esc(l.end) + '</td><td>' + esc(l.days) + '</td><td class="wrap">' + esc(l.reason) +
          '</td><td>' + badge(l.status) + '</td><td class="wrap">' + esc(l.reviewer) + (l.comment ? ':' + esc(l.comment) : '') + '</td><td>' +
          (l.status === '待簽核' ? '<button class="sm danger" data-cancel="' + esc(l.id) + '">撤回</button>' : '') + '</td></tr>';
      }), '還沒有請假紀錄');
      document.querySelectorAll('[data-cancel]').forEach(function (b) {
        b.onclick = function () {
          if (!confirm('確定撤回這張假單?')) return;
          guard(async function () { await call('leave.cancel', { id: b.dataset.cancel }); toast('已撤回'); viewLeave(main); }, b);
        };
      });
    }
    drawList();

    $('#lf').onsubmit = function (e) {
      e.preventDefault();
      guard(async function () {
        await call('leave.apply', { type: val('lt'), start: val('ls'), end: val('le'), days: Number(val('ld')), reason: val('lr') });
        toast('已送出,等待主管簽核'); viewLeave(main);
      }, $('#lf button'));
    };
  }

  /* ---------- 排班 ---------- */
  async function viewSchedule(main) {
    var month = (viewSchedule.month = viewSchedule.month || curMonth());
    var canEdit = user.role !== 'employee';
    var res = await Promise.all([call('shifts.list', { month: month }), call('shifts.types'), canEdit ? call('employees.list') : Promise.resolve([])]);
    var shifts = res[0], types = res[1];
    var emps = res[2].filter(function (e) { return e.status === '在職' && (user.role === 'admin' || e.id !== user.id); });
    var n = daysIn(month), WD = ['日', '一', '二', '三', '四', '五', '六'];
    var map = {}; shifts.forEach(function (s) { map[s.date + '|' + s.empId] = s.shift; });
    var dirty = {};

    var head = '<div class="row"><h2 style="margin:0">排班</h2><div class="spacer"></div><button class="sm" id="pm">◀</button><input type="month" id="mon" value="' + esc(month) + '" style="width:auto"><button class="sm" id="nm">▶</button>' +
      (canEdit ? '<button class="primary" id="sv" disabled>儲存排班</button>' : '') + '</div>';

    var body = '';
    if (canEdit) {
      var th = '<th class="name">員工</th>'; for (var d = 1; d <= n; d++) { var dt = new Date(month + '-' + pad(d)); th += '<th class="' + (dt.getDay() % 6 === 0 ? 'we' : '') + '">' + d + '<br>' + WD[dt.getDay()] + '</th>'; }
      var rows = emps.map(function (e) {
        var tds = '<td class="name">' + esc(e.name) + '</td>';
        for (var d2 = 1; d2 <= n; d2++) {
          var ds = month + '-' + pad(d2), cur = map[ds + '|' + e.id] || '', we = new Date(ds).getDay() % 6 === 0;
          tds += '<td class="' + (we ? 'we' : '') + '"><select data-d="' + ds + '" data-e="' + esc(e.id) + '"><option value=""></option>' +
            types.map(function (t) { return '<option' + (t.name === cur ? ' selected' : '') + '>' + esc(t.name) + '</option>'; }).join('') + '</select></td>';
        }
        return '<tr>' + tds + '</tr>';
      });
      body = '<div class="card tablewrap"><table class="sched"><thead><tr>' + th + '</tr></thead><tbody>' + (rows.join('') || '<tr><td>沒有可排班的員工</td></tr>') + '</tbody></table></div>' +
        '<p class="hint">班別:' + types.map(function (t) { return esc(t.name) + (t.start ? '(' + esc(t.start) + '–' + esc(t.end) + ')' : ''); }).join('、') + '。選空白代表清除排班。</p>';
    } else {
      var mine = shifts.filter(function (s) { return s.empId === user.id; }).sort(function (a, b) { return a.date < b.date ? -1 : 1; });
      body = '<div class="card">' + table(['日期', '星期', '班別', '上班', '下班'], mine.map(function (s) {
        return '<tr><td>' + esc(s.date) + '</td><td>' + WD[new Date(s.date).getDay()] + '</td><td>' + esc(s.shift) + '</td><td>' + esc(s.start) + '</td><td>' + esc(s.end) + '</td></tr>';
      }), '本月尚無排班') + '</div>';
    }
    main.innerHTML = head + body;

    function go(m) { viewSchedule.month = m; viewSchedule(main).catch(function (e) { toast(e.message, true); }); }
    $('#mon').onchange = function () { if (this.value) go(this.value); };
    $('#pm').onclick = function () { var p = month.split('-'); go(fmtDate(new Date(+p[0], +p[1] - 2, 1)).slice(0, 7)); };
    $('#nm').onclick = function () { var p = month.split('-'); go(fmtDate(new Date(+p[0], +p[1], 1)).slice(0, 7)); };
    if (canEdit) {
      document.querySelectorAll('.sched select').forEach(function (s) {
        s.onchange = function () {
          var k = s.dataset.d + '|' + s.dataset.e;
          if ((map[k] || '') === s.value) delete dirty[k]; else dirty[k] = { date: s.dataset.d, empId: s.dataset.e, shift: s.value };
          $('#sv').disabled = !Object.keys(dirty).length;
        };
      });
      $('#sv').onclick = function () {
        guard(async function () {
          var items = Object.keys(dirty).map(function (k) { return dirty[k]; });
          await call('shifts.save', { items: items });
          toast('已儲存 ' + items.length + ' 筆排班'); viewSchedule(main);
        }, this);
      };
    }
  }

  /* ---------- 假單簽核 ---------- */
  async function viewApprovals(main) {
    var res = await Promise.all([call('leave.pending'), call('leave.all', { month: '' })]);
    var pending = res[0], history = res[1].filter(function (l) { return l.status !== '待簽核'; }).slice(0, 50);
    main.innerHTML = '<h2>假單簽核</h2><div class="card"><h3>待簽核(' + pending.length + ')</h3>' +
      table(['單號', '員工', '假別', '期間', '天數', '事由', ''], pending.map(function (l) {
        return '<tr><td>' + esc(l.id) + '</td><td>' + esc(l.name) + '(' + esc(l.empId) + ')</td><td>' + esc(l.type) + '</td><td>' + esc(l.start) + ' ~ ' + esc(l.end) + '</td><td>' + esc(l.days) +
          '</td><td class="wrap">' + esc(l.reason) + '</td><td><button class="sm ok" data-ok="' + esc(l.id) + '">核准</button> <button class="sm danger" data-no="' + esc(l.id) + '">駁回</button></td></tr>';
      }), '沒有待簽核的假單') + '</div>' +
      '<div class="card"><h3>近期已處理</h3>' + table(['單號', '員工', '假別', '期間', '天數', '狀態', '簽核人 / 意見'], history.map(function (l) {
        return '<tr><td>' + esc(l.id) + '</td><td>' + esc(l.name) + '</td><td>' + esc(l.type) + '</td><td>' + esc(l.start) + ' ~ ' + esc(l.end) + '</td><td>' + esc(l.days) + '</td><td>' + badge(l.status) + '</td><td class="wrap">' + esc(l.reviewer) + (l.comment ? ':' + esc(l.comment) : '') + '</td></tr>';
      }), '尚無紀錄') + '</div>';

    function review(id, decision, btn) {
      var comment = prompt(decision + '意見(可留空):', '');
      if (comment === null) return;
      guard(async function () { await call('leave.review', { id: id, decision: decision, comment: comment }); toast('已' + decision); viewApprovals(main); }, btn);
    }
    document.querySelectorAll('[data-ok]').forEach(function (b) { b.onclick = function () { review(b.dataset.ok, '核准', b); }; });
    document.querySelectorAll('[data-no]').forEach(function (b) { b.onclick = function () { review(b.dataset.no, '駁回', b); }; });
  }

  /* ---------- 出勤報表 ---------- */
  async function viewAttendance(main) {
    var month = (viewAttendance.month = viewAttendance.month || curMonth());
    var res = await Promise.all([call('punch.report', { month: month }), call('shifts.list', { month: month })]);
    var rows = res[0];
    var sum = {};
    rows.forEach(function (r) {
      var s = sum[r.empId] || (sum[r.empId] = { name: r.name, days: {}, bad: 0 });
      if (r.result === '成功') s.days[r.date] = true; else s.bad++;
    });
    main.innerHTML = '<div class="row"><h2 style="margin:0">出勤報表</h2><div class="spacer"></div><input type="month" id="mon" value="' + esc(month) + '" style="width:auto"></div>' +
      '<div class="card"><h3>月彙總</h3>' + table(['員編', '姓名', '出勤天數', '超出範圍嘗試'], Object.keys(sum).map(function (id) {
        return '<tr><td>' + esc(id) + '</td><td>' + esc(sum[id].name) + '</td><td>' + Object.keys(sum[id].days).length + '</td><td>' + (sum[id].bad ? '<span class="badge bad">' + sum[id].bad + '</span>' : 0) + '</td></tr>';
      }), '本月沒有打卡紀錄') + '</div>' +
      '<div class="card"><h3>打卡明細</h3>' + table(['時間', '員編', '姓名', '類型', '地點', '距離', '結果'], rows.map(function (r) {
        return '<tr><td>' + esc(r.time) + '</td><td>' + esc(r.empId) + '</td><td>' + esc(r.name) + '</td><td>' + esc(r.type) + '</td><td>' + esc(r.place) + '</td><td>' + esc(r.distance) + ' m</td><td>' + badge(r.result) + '</td></tr>';
      }), '沒有資料') + '</div>';
    $('#mon').onchange = function () { if (this.value) { viewAttendance.month = this.value; viewAttendance(main); } };
  }

  /* ---------- 員工資料 ---------- */
  async function viewEmployees(main) {
    var list = await call('employees.list');
    var roleName = { admin: '管理員', manager: '主管', employee: '員工' };
    main.innerHTML = '<div class="row"><h2 style="margin:0">員工資料</h2><div class="spacer"></div><button class="primary" id="add">新增員工</button></div><div class="card">' +
      table(['員編', '姓名', '部門', '職稱', '角色', '主管', '到職日', '狀態', ''], list.map(function (e) {
        return '<tr><td>' + esc(e.id) + '</td><td>' + esc(e.name) + '</td><td>' + esc(e.dept) + '</td><td>' + esc(e.title) + '</td><td>' + roleName[e.role] + '</td><td>' + esc(e.manager) + '</td><td>' + esc(e.hireDate) + '</td><td>' + badge(e.status) +
          '</td><td><button class="sm" data-edit="' + esc(e.id) + '">編輯</button> <button class="sm" data-pw="' + esc(e.id) + '">重設密碼</button></td></tr>';
      })) + '</div>';

    function form(e) {
      var isNew = !e; e = e || { role: 'employee', status: '在職' };
      var d = openDialog('<h3>' + (isNew ? '新增員工' : '編輯員工') + '</h3><form id="ef" class="form">' +
        field('f_id', '員編', e.id || '', 'text', (isNew ? '' : 'readonly ') + 'required') + field('f_name', '姓名', e.name || '', 'text', 'required') +
        field('f_dept', '部門', e.dept || '') + field('f_title', '職稱', e.title || '') +
        '<div><label for="f_role">角色</label><select id="f_role">' + ['employee', 'manager', 'admin'].map(function (r) { return '<option value="' + r + '"' + (e.role === r ? ' selected' : '') + '>' + roleName[r] + '</option>'; }).join('') + '</select></div>' +
        '<div><label for="f_mgr">直屬主管</label><select id="f_mgr"><option value="">(無)</option>' + list.filter(function (x) { return x.role !== 'employee' && x.id !== e.id; }).map(function (x) { return '<option value="' + esc(x.id) + '"' + (e.manager === x.id ? ' selected' : '') + '>' + esc(x.name) + '(' + esc(x.id) + ')</option>'; }).join('') + '</select></div>' +
        field('f_hire', '到職日', e.hireDate || '', 'date') +
        '<div><label for="f_st">狀態</label><select id="f_st"><option' + (e.status === '在職' ? ' selected' : '') + '>在職</option><option' + (e.status === '離職' ? ' selected' : '') + '>離職</option></select></div>' +
        field('f_mail', 'Email', e.email || '', 'email') + field('f_tel', '電話', e.phone || '') +
        (isNew ? field('f_pw', '初始密碼(至少 8 碼)', '', 'text', 'minlength="8" required') : '') +
        '<div class="full row"><button class="primary" type="submit">儲存</button><button type="button" data-close>取消</button></div></form>');
      $('#ef', d).onsubmit = function (ev) {
        ev.preventDefault();
        guard(async function () {
          await call('employees.save', { id: val('f_id').trim(), name: val('f_name').trim(), dept: val('f_dept'), title: val('f_title'), role: val('f_role'), manager: val('f_mgr'), hireDate: val('f_hire'), status: val('f_st'), email: val('f_mail'), phone: val('f_tel'), password: isNew ? val('f_pw') : undefined });
          d.close(); toast('已儲存'); viewEmployees(main);
        }, $('#ef button', d));
      };
    }
    $('#add').onclick = function () { form(null); };
    document.querySelectorAll('[data-edit]').forEach(function (b) { b.onclick = function () { form(list.filter(function (x) { return x.id === b.dataset.edit; })[0]); }; });
    document.querySelectorAll('[data-pw]').forEach(function (b) {
      b.onclick = function () {
        var pw = prompt('請輸入 ' + b.dataset.pw + ' 的新密碼(至少 8 碼):');
        if (!pw) return;
        guard(async function () { await call('employees.resetPassword', { id: b.dataset.pw, password: pw }); toast('密碼已重設'); }, b);
      };
    });
  }

  /* ---------- 打卡地點 ---------- */
  async function viewLocations(main) {
    var list = await call('locations.list');
    main.innerHTML = '<div class="row"><h2 style="margin:0">打卡地點(GPS 定點)</h2><div class="spacer"></div><button class="primary" id="add">新增地點</button></div><div class="card">' +
      table(['名稱', '緯度', '經度', '半徑(公尺)', '啟用', ''], list.map(function (l) {
        return '<tr><td>' + esc(l.name) + '</td><td>' + esc(l.lat) + '</td><td>' + esc(l.lng) + '</td><td>' + esc(l.radius) + '</td><td>' + (l.active ? badge('成功').replace('成功', '啟用') : badge('撤回').replace('撤回', '停用')) +
          '</td><td><button class="sm" data-edit="' + esc(l.id) + '">編輯</button> <button class="sm danger" data-del="' + esc(l.id) + '">刪除</button></td></tr>';
      }), '尚未設定地點') + '</div>' +
      '<p class="hint">提示:在 Google 地圖上對地點按右鍵,第一行數字即為「緯度, 經度」。也可在該地點用手機開啟本頁,按「使用目前位置」。</p>';

    function form(l) {
      var isNew = !l; l = l || { lat: '', lng: '', radius: 100, active: true };
      var d = openDialog('<h3>' + (isNew ? '新增地點' : '編輯地點') + '</h3><form id="lf" class="form">' +
        '<div class="full">' + field('l_name', '名稱', l.name || '', 'text', 'required') + '</div>' +
        field('l_lat', '緯度', l.lat, 'number', 'step="any" required') + field('l_lng', '經度', l.lng, 'number', 'step="any" required') +
        field('l_rad', '允許半徑(公尺)', l.radius, 'number', 'min="10" max="5000" required') +
        '<div><label for="l_act">狀態</label><select id="l_act"><option value="1"' + (l.active ? ' selected' : '') + '>啟用</option><option value="0"' + (!l.active ? ' selected' : '') + '>停用</option></select></div>' +
        '<div class="full"><button type="button" id="l_here">使用目前位置</button></div>' +
        '<div class="full row"><button class="primary" type="submit">儲存</button><button type="button" data-close>取消</button></div></form>');
      $('#l_here', d).onclick = function () {
        guard(async function () { var p = await getPosition(); document.getElementById('l_lat').value = p.coords.latitude.toFixed(6); document.getElementById('l_lng').value = p.coords.longitude.toFixed(6); toast('已帶入目前位置(精準度約 ' + Math.round(p.coords.accuracy) + ' 公尺)'); }, this);
      };
      $('#lf', d).onsubmit = function (ev) {
        ev.preventDefault();
        guard(async function () {
          await call('locations.save', { id: l.id, name: val('l_name').trim(), lat: Number(val('l_lat')), lng: Number(val('l_lng')), radius: Number(val('l_rad')), active: val('l_act') === '1' });
          d.close(); toast('已儲存'); viewLocations(main);
        }, $('#lf button[type=submit]', d));
      };
    }
    $('#add').onclick = function () { form(null); };
    document.querySelectorAll('[data-edit]').forEach(function (b) { b.onclick = function () { form(list.filter(function (x) { return x.id === b.dataset.edit; })[0]); }; });
    document.querySelectorAll('[data-del]').forEach(function (b) {
      b.onclick = function () {
        if (!confirm('確定刪除這個打卡地點?')) return;
        guard(async function () { await call('locations.delete', { id: b.dataset.del }); toast('已刪除'); viewLocations(main); }, b);
      };
    });
  }

  /* ---------- 修改密碼 ---------- */
  function viewPassword(main) {
    main.innerHTML = '<h2>修改密碼</h2><div class="card" style="max-width:420px"><form id="pf" class="form" style="grid-template-columns:1fr">' +
      field('p0', '目前密碼', '', 'password', 'autocomplete="current-password" required') + field('p1', '新密碼(至少 8 碼)', '', 'password', 'autocomplete="new-password" minlength="8" required') +
      field('p2', '再輸入一次新密碼', '', 'password', 'autocomplete="new-password" required') + '<button class="primary" type="submit">更新密碼</button></form></div>';
    $('#pf').onsubmit = function (e) {
      e.preventDefault();
      if (val('p1') !== val('p2')) return toast('兩次輸入的新密碼不一致', true);
      guard(async function () { await call('changePassword', { oldPassword: val('p0'), newPassword: val('p1') }); toast('密碼已更新'); $('#pf').reset(); }, $('#pf button'));
    };
  }

  /* ---------- 啟動 ---------- */
  document.title = cfg.COMPANY_NAME || '人事管理系統';
  if (API.hasToken()) {
    call('me').then(function (u) { user = u; renderShell(); }).catch(function () { API.logout(); renderLogin(); });
  } else renderLogin();
})();
