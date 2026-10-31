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


  /* ---------- 地點搜尋(自動抓座標) ---------- */
  async function geocode(query) {
    var list = [];
    try { list = await call('locations.geocode', { query: query }); }
    catch (e) { if (e.code === 'AUTH') throw e; list = []; } // 後端查詢失敗就改用備援
    if (list && list.length) return list;
    var wait = 1100 - (Date.now() - (geocode.last || 0)); // OSM 規定每秒最多 1 次
    if (wait > 0) await new Promise(function (r) { setTimeout(r, wait); });
    geocode.last = Date.now();
    var url = 'https://nominatim.openstreetmap.org/search?format=jsonv2&limit=5&countrycodes=tw&accept-language=zh-TW&q=' + encodeURIComponent(query);
    var res = await fetch(url);
    if (!res.ok) throw new Error('地點搜尋服務暫時無法使用,請手動輸入座標');
    var j = await res.json();
    return j.map(function (x) { return { name: x.display_name, lat: Number(x.lat), lng: Number(x.lon), source: 'osm' }; });
  }

  /* ---------- Excel 工具 ---------- */
  var xlsxPromise = null;
  function loadXLSX() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxPromise) xlsxPromise = new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = 'vendor/xlsx.mini.min.js';
      s.onload = function () { resolve(window.XLSX); };
      s.onerror = function () { xlsxPromise = null; reject(new Error('無法載入 Excel 元件')); };
      document.head.appendChild(s);
    });
    return xlsxPromise;
  }
  var libPromises = {};
  function loadLib(src, name) {
    if (window[name]) return Promise.resolve(window[name]);
    if (!libPromises[src]) libPromises[src] = new Promise(function (resolve, reject) {
      var sc = document.createElement('script');
      sc.src = src;
      sc.onload = function () { resolve(window[name]); };
      sc.onerror = function () { delete libPromises[src]; reject(new Error('無法載入 QR 元件')); };
      document.head.appendChild(sc);
    });
    return libPromises[src];
  }

  /* ---------- QR 打卡:連結 / 掃描內容解析 ---------- */
  var pendingQr = null; // 掃描(或點開連結)取得、尚未使用的 QR
  function parseQrText(text) {
    text = String(text || '').trim();
    var i = text.indexOf('#/qr?');
    if (i < 0) return null;
    var sp = new URLSearchParams(text.slice(i + 5));
    var loc = sp.get('l'), tok = sp.get('t');
    if (!loc || !tok || loc.length > 40 || !/^[0-9a-f]{4,64}$/i.test(tok)) return null;
    return { loc: loc, token: tok };
  }
  function captureQr() {
    if (!/^#\/qr\?/.test(location.hash)) return false;
    var q = parseQrText(location.hash);
    history.replaceState(null, '', location.pathname + location.search + '#/punch');
    if (q) pendingQr = q;
    return !!q;
  }

  var EMP_COLS = [['員編', 'id'], ['姓名', 'name'], ['部門', 'dept'], ['職稱', 'title'], ['角色', 'role'], ['主管員編', 'manager'], ['到職日', 'hireDate'], ['狀態', 'status'], ['Email', 'email'], ['電話', 'phone'], ['初始密碼', 'password']];
  function normDate(v) {
    v = String(v == null ? '' : v).trim();
    var m = v.match(/^(\d{4})[\/.\-年](\d{1,2})[\/.\-月](\d{1,2})日?$/);
    return m ? m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]) : v;
  }
  async function downloadEmpTemplate() {
    var X = await loadXLSX();
    var wb = X.utils.book_new();
    var rows = [EMP_COLS.map(function (c) { return c[0]; }),
      ['E101', '王小明', '門市部', '門市人員', '員工', 'M001', '2025-03-01', '在職', 'ming@example.com', '0912345678', ''],
      ['M101', '李主管', '門市部', '店長', '主管', '', '2024-01-15', '在職', '', '', '']];
    var ws = X.utils.aoa_to_sheet(rows);
    ws['!cols'] = [10, 10, 12, 12, 8, 10, 12, 8, 24, 14, 12].map(function (w) { return { wch: w }; });
    X.utils.book_append_sheet(wb, ws, '員工');
    var help = X.utils.aoa_to_sheet([
      ['欄位', '說明'], ['員編', '必填。英數字、底線、連字號,最多 20 碼。已存在的員編會更新資料'], ['姓名', '必填'],
      ['角色', '員工 / 主管 / 管理員(留空視為員工)'], ['主管員編', '選填。必須是已存在或同檔案內、角色為主管/管理員的員編'],
      ['到職日', '格式 YYYY-MM-DD(也接受 2025/3/1)'], ['狀態', '在職 / 離職(留空視為在職)'],
      ['初始密碼', '選填,至少 8 碼。新員工留空會自動產生隨機密碼;既有員工留空則不改密碼']]);
    help['!cols'] = [{ wch: 12 }, { wch: 70 }];
    X.utils.book_append_sheet(wb, help, '填寫說明');
    X.writeFile(wb, '員工匯入範本.xlsx');
  }
  async function readEmpFile(file) {
    var X = await loadXLSX();
    var wb = X.read(await file.arrayBuffer(), { type: 'array', cellDates: true, dateNF: 'yyyy-mm-dd' });
    var ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) throw new Error('檔案中沒有工作表');
    var aoa = X.utils.sheet_to_json(ws, { header: 1, raw: false, dateNF: 'yyyy-mm-dd', defval: '' });
    if (!aoa.length) throw new Error('檔案是空的');
    var head = aoa[0].map(function (h) { return String(h).trim(); });
    var alias = { '主管': '主管員編', '密碼': '初始密碼', 'E-mail': 'Email', '信箱': 'Email', '手機': '電話' };
    head = head.map(function (h) { return alias[h] || h; });
    if (head.indexOf('員編') < 0 || head.indexOf('姓名') < 0) throw new Error('找不到「員編」「姓名」欄位,請使用範本格式');
    var out = [];
    for (var i = 1; i < aoa.length; i++) {
      var row = aoa[i];
      if (!row.some(function (c) { return String(c).trim() !== ''; })) continue;
      var o = { line: i + 1 };
      EMP_COLS.forEach(function (c) { var k = head.indexOf(c[0]); o[c[1]] = k >= 0 ? String(row[k] == null ? '' : row[k]).trim() : ''; });
      o.hireDate = normDate(o.hireDate);
      out.push(o);
    }
    if (!out.length) throw new Error('沒有找到資料列');
    if (out.length > 500) throw new Error('單次最多匯入 500 筆,目前有 ' + out.length + ' 筆');
    return out;
  }


  async function downloadLocTemplate() {
    var X = await loadXLSX();
    var wb = X.utils.book_new();
    var ws = X.utils.aoa_to_sheet([['名稱', '地址(選填,用來搜尋座標)', '緯度(選填)', '經度(選填)', '允許半徑(公尺)', '狀態'],
      ['寶雅 高雄巨蛋店', '高雄市左營區博愛二路777號', '', '', 100, '啟用'],
      ['總公司', '', 25.033964, 121.564468, 150, '啟用']]);
    ws['!cols'] = [20, 34, 12, 12, 14, 8].map(function (w) { return { wch: w }; });
    X.utils.book_append_sheet(wb, ws, '打卡地點');
    var help = X.utils.aoa_to_sheet([['欄位', '說明'], ['名稱', '必填。名稱已存在的地點會被更新,不會重複新增'],
      ['地址', '緯度、經度留空時,系統用「地址」搜尋座標;地址也留空就用「名稱」搜尋。建議填完整地址以提高準確度'],
      ['緯度 / 經度', '選填。兩者都填就直接使用,不再搜尋'], ['允許半徑', '10~5000 公尺,留空視為 100'], ['狀態', '啟用 / 停用,留空視為啟用']]);
    help['!cols'] = [{ wch: 14 }, { wch: 80 }];
    X.utils.book_append_sheet(wb, help, '填寫說明');
    X.writeFile(wb, '打卡地點匯入範本.xlsx');
  }
  async function readLocFile(file) {
    var X = await loadXLSX();
    var wb = X.read(await file.arrayBuffer(), { type: 'array' });
    var ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) throw new Error('檔案中沒有工作表');
    var aoa = X.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    if (!aoa.length) throw new Error('檔案是空的');
    var head = aoa[0].map(function (h) { return String(h).trim(); });
    function col(prefix) { for (var i = 0; i < head.length; i++) if (head[i].indexOf(prefix) === 0) return i; return -1; }
    var c = { name: col('名稱'), addr: col('地址'), lat: col('緯度'), lng: col('經度'), rad: col('允許半徑') >= 0 ? col('允許半徑') : col('半徑'), st: col('狀態') };
    if (c.name < 0) throw new Error('找不到「名稱」欄位,請使用範本格式');
    var out = [];
    for (var i = 1; i < aoa.length; i++) {
      var row = aoa[i];
      if (!row.some(function (x) { return String(x).trim() !== ''; })) continue;
      function g(k) { return c[k] >= 0 ? String(row[c[k]] == null ? '' : row[c[k]]).trim() : ''; }
      var st = g('st');
      out.push({ line: i + 1, name: g('name'), query: g('addr'), lat: g('lat'), lng: g('lng'), radius: g('rad'), active: st !== '停用' });
    }
    if (!out.length) throw new Error('沒有找到資料列');
    if (out.length > 300) throw new Error('單次最多匯入 300 筆,目前有 ' + out.length + ' 筆');
    return out;
  }


  /* ---------- 取得本機對外 IP(Apps Script 看不到來源 IP,由瀏覽器查詢後一併回報) ---------- */
  var ipPromise = null;
  function getIP() {
    if (!ipPromise) {
      var tryUrl = function (url) {
        var ctl = new AbortController(), t = setTimeout(function () { ctl.abort(); }, 4000);
        return fetch(url, { signal: ctl.signal, cache: 'no-store' }).then(function (r) { return r.json(); }).then(function (j) { clearTimeout(t); return String(j.ip || ''); });
      };
      ipPromise = tryUrl('https://api64.ipify.org?format=json').catch(function () { return tryUrl('https://api.ipify.org?format=json'); }).catch(function () { return ''; })
        .then(function (ip) { return /^[0-9a-fA-F:.]{3,45}$/.test(ip) ? ip : ''; });
    }
    return ipPromise;
  }
  function shortUA(ua) {
    ua = String(ua || '');
    var os = /Android/i.test(ua) ? 'Android' : /iPhone|iPad|iOS/i.test(ua) ? 'iOS' : /Windows/i.test(ua) ? 'Windows' : /Mac OS/i.test(ua) ? 'macOS' : /Linux/i.test(ua) ? 'Linux' : '';
    var br = /Edg\//.test(ua) ? 'Edge' : /OPR\//.test(ua) ? 'Opera' : /Firefox\//.test(ua) ? 'Firefox' : /Chrome\//.test(ua) ? 'Chrome' : /Safari\//.test(ua) ? 'Safari' : '';
    return (os + ' ' + br).trim() || (ua ? ua.slice(0, 30) : '');
  }

  /* ---------- 登入 ---------- */
  function renderLogin() {
    clearInterval(clockTimer);
    app.innerHTML = '<div class="login card"><h1>' + esc(cfg.COMPANY_NAME || '人事管理系統') + (API.demo ? '<span class="demo-tag">示範模式</span>' : '') + '</h1>' +
      '<form id="lf" class="form" style="grid-template-columns:1fr">' + field('li', '員編', '', 'text', 'autocomplete="username" required') +
      field('lp', '密碼', '', 'password', 'autocomplete="current-password" required') +
      '<button class="primary" type="submit">登入</button></form>' +
      '<p class="hint" id="myip">正在取得您的 IP…</p><p class="hint">登入時系統會記錄時間與 IP 位址,供安全稽核使用。</p>' +
      (API.demo ? '<p class="hint">示範帳號:管理員 A001 / admin1234;主管 M001 / demo1234;員工 E001 / demo1234。資料只存在這個瀏覽器。</p>' : '') + '</div>';
    getIP().then(function (ip) { var el = document.getElementById('myip'); if (el) el.textContent = ip ? '您目前的 IP:' + ip : '無法取得 IP(仍可登入,紀錄會標示為「未知」)'; });
    $('#lf').onsubmit = function (e) {
      e.preventDefault();
      guard(async function () {
        var ip = await getIP();
        var r = await API.login(val('li').trim(), val('lp'), { ip: ip, ua: navigator.userAgent });
        user = r.user;
        location.hash = '#/punch';
        renderShell();
        var msg = '登入 IP:' + (ip || '未知') + (r.lastLogin ? '。上次登入:' + r.lastLogin.time + '(' + r.lastLogin.ip + ')' : '。這是您第一次登入');
        if (r.mustChange) { toast('請先修改預設密碼', true); location.hash = '#/password'; } else toast(msg);
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
    { path: 'loginlog', label: '登入紀錄', roles: ['admin'], group: '管理', view: viewLoginLog },
    { path: 'password', label: '修改密碼', roles: ['employee', 'manager', 'admin'], group: '帳號', view: viewPassword }
  ];

  function renderShell() {
    if (!user) return renderLogin();
    var roleName = { admin: '管理員', manager: '主管', employee: '員工' }[user.role];
    app.innerHTML = '<div class="topbar"><div class="brand">' + esc(cfg.COMPANY_NAME || '人事管理系統') + (API.demo ? '<span class="demo-tag">示範</span>' : '') + '</div>' +
      '<div class="row"><span class="who">' + esc(user.name) + '(' + esc(user.id) + ' · ' + roleName + ')</span><span class="who" id="ipb"></span><button class="sm" id="lo">登出</button></div></div>' +
      '<div class="layout"><nav class="nav" id="nav"></nav><main id="main"></main></div>';
    $('#lo').onclick = function () { API.logout(); user = null; location.hash = ''; renderLogin(); };
    getIP().then(function (ip) { var el = document.getElementById('ipb'); if (el && ip) el.textContent = 'IP ' + ip; });
    route();
  }

  function route() {
    if (!user) return;
    clearInterval(clockTimer);
    captureQr();
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

  /* 員工端:開相機掃描 QR(也可直接用手機內建相機掃,會開啟本系統並帶入) */
  async function scanQr(onFound) {
    var stream = null, raf = 0, done = false;
    function stop() { done = true; cancelAnimationFrame(raf); if (stream) stream.getTracks().forEach(function (t) { t.stop(); }); stream = null; }
    var d = openDialog('<h3>掃描打卡 QR</h3><video id="qv" playsinline muted style="width:100%;border-radius:8px;background:#000"></video><canvas id="qc" class="hidden"></canvas>' +
      '<p class="hint" id="qmsg">正在開啟相機…</p><div class="row" style="margin-top:8px"><button data-close>取消</button></div>');
    d.onclose = stop;
    var c = d.querySelector('[data-close]'); if (c) c.onclick = function () { d.close(); };
    try {
      if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) throw new Error('nocam');
      var jsq = await loadLib('vendor/jsQR.js', 'jsQR');
      stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' }, audio: false });
      if (done) return stop();
      var v = $('#qv'); v.srcObject = stream; await v.play();
      $('#qmsg').textContent = '請將鏡頭對準現場顯示的 QR';
      var cv = $('#qc'), cx = cv.getContext('2d', { willReadFrequently: true });
      (function loop() {
        if (done) return;
        if (v.videoWidth) {
          var sc = Math.min(1, 640 / v.videoWidth);
          cv.width = Math.round(v.videoWidth * sc); cv.height = Math.round(v.videoHeight * sc);
          cx.drawImage(v, 0, 0, cv.width, cv.height);
          var img = cx.getImageData(0, 0, cv.width, cv.height);
          var code = jsq(img.data, img.width, img.height, { inversionAttempts: 'dontInvert' });
          if (code && code.data) {
            var q = parseQrText(code.data);
            if (q) { stop(); d.close(); onFound(q); return; }
            $('#qmsg').textContent = '這不是本系統的打卡 QR,請重試';
          }
        }
        raf = requestAnimationFrame(loop);
      })();
    } catch (e) {
      stop();
      var m = $('#qmsg');
      if (m) m.textContent = e.name === 'NotAllowedError' ? '相機權限被拒絕,請在瀏覽器設定中允許,或改用手機內建相機掃描 QR。' : '無法開啟相機,請改用手機內建相機掃描 QR,掃描後會自動開啟本系統。';
    }
  }

  /* 管理員端:顯示每個地點固定的 QR,可列印張貼 */
  async function showLocQr(loc) {
    var d = openDialog('<h3>' + esc(loc.name) + ' 打卡 QR</h3><div id="qrbox" style="max-width:320px;margin:0 auto;min-height:200px"></div>' +
      '<p class="hint" id="qrmsg">產生中…</p><p class="hint">這是固定的 QR,列印後張貼在打卡地點即可。注意:有人拍照轉傳,別人也能掃;系統在有定位時仍會檢查是否在範圍內。</p>' +
      '<div class="row" style="margin-top:8px"><button class="primary" id="qrprint" disabled>列印</button><button data-close>關閉</button></div>');
    var c = d.querySelector('[data-close]'); if (c) c.onclick = function () { d.close(); };
    try {
      var r = await call('qr.token', { id: loc.id });
      var gen = await loadLib('vendor/qrcode.js', 'qrcode');
      var url = location.origin + location.pathname + '#/qr?l=' + encodeURIComponent(loc.id) + '&t=' + r.token;
      var q = gen(0, 'M'); q.addData(url); q.make();
      var svg = q.createSvgTag({ cellSize: 4, margin: 2, scalable: true });
      $('#qrbox').innerHTML = svg;
      $('#qrmsg').textContent = '用手機相機掃描即可開啟打卡。';
      var pb = $('#qrprint'); pb.disabled = false;
      pb.onclick = function () {
        var w = window.open('', '_blank');
        if (!w) { toast('瀏覽器擋住了列印視窗,請允許彈出視窗', true); return; }
        w.document.write('<!doctype html><meta charset="utf-8"><title>' + esc(loc.name) + ' 打卡 QR</title><body style="font-family:sans-serif;text-align:center;margin:40px"><h1>' + esc(loc.name) + '</h1><div style="width:360px;margin:0 auto">' + svg + '</div><h2>掃描 QR 打卡</h2><p>用手機相機掃描,登入後選擇上班或下班</p></body>');
        w.document.close(); w.focus(); w.print();
      };
    } catch (e) { $('#qrmsg').textContent = e.code === 'AUTH' ? '登入已逾時,請重新登入。' : e.message; }
  }

  async function viewPunch(main) {
    var today = fmtDate(new Date());
    var res = await Promise.all([call('locations.list'), call('punch.my', { month: curMonth() }), call('shifts.list', { month: curMonth() })]);
    var locs = res[0].filter(function (l) { return l.active !== false; });
    var records = res[1].filter(function (r) { return r.date === today; });
    var myShift = res[2].filter(function (s) { return s.date === today && s.empId === user.id; })[0];
    var pos = null;
    var qr = pendingQr; pendingQr = null; // 取用從連結帶進來的 QR
    function locName(id) { var l = locs.filter(function (x) { return x.id === id; })[0]; return l ? l.name : '未知地點'; }

    main.innerHTML = '<h2>打卡</h2><div class="card"><div class="clock" id="clk"></div>' +
      '<div class="muted">' + esc(today) + (myShift ? ' · 今日班別:' + esc(myShift.shift) + (myShift.start ? '(' + esc(myShift.start) + '–' + esc(myShift.end) + ')' : '') : ' · 今日未排班') + '</div>' +
      '<div class="geo" id="geo">尚未定位</div>' +
      '<div class="geo" id="qrst"></div>' +
      '<div class="row" style="margin-top:8px"><button id="loc">取得目前位置</button><button id="scan">掃 QR 打卡</button></div>' +
      '<div class="punch-btns"><button class="primary" id="in" disabled>上班打卡</button><button class="ok" id="out" disabled>下班打卡</button></div>' +
      '<p class="hint">需在設定的打卡地點範圍內才能打卡;或掃描現場張貼的 QR。可打卡地點:' + (locs.length ? locs.map(function (l) { return esc(l.name) + '(' + l.radius + ' 公尺)'; }).join('、') : '尚未設定') + '</p></div>' +
      '<div class="card"><h3>今日打卡紀錄</h3><div id="today"></div></div>';

    function tick() { $('#clk').textContent = new Date().toLocaleTimeString('zh-TW', { hour12: false }); }
    tick(); clockTimer = setInterval(tick, 1000);

    function showToday() {
      $('#today').innerHTML = table(['時間', '類型', '地點', '距離', '結果'], records.map(function (r) {
        return '<tr><td>' + esc(r.time.slice(11)) + '</td><td>' + esc(r.type) + '</td><td>' + esc(r.place) + '</td><td>' + (r.distance === '' || r.distance == null ? '—' : esc(r.distance) + ' m') + '</td><td>' + badge(r.result) + '</td></tr>';
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
      } catch (e) { pos = null; g.textContent = e.message; }
      refreshBtns();
    }
    function refreshBtns() { $('#in').disabled = $('#out').disabled = !(pos || qr); }
    function showQr() {
      var el = $('#qrst');
      if (!qr) { el.innerHTML = ''; return; }
      el.innerHTML = '已掃描 QR:<b>' + esc(locName(qr.loc)) + '</b> <button class="sm" id="qrx">取消</button>';
      $('#qrx').onclick = function () { qr = null; showQr(); refreshBtns(); };
    }
    $('#loc').onclick = locate;
    $('#scan').onclick = function () {
      scanQr(function (q) { qr = q; showQr(); refreshBtns(); toast('已掃描「' + locName(q.loc) + '」,請按上班或下班打卡'); });
    };
    showQr(); refreshBtns();
    locate();

    async function punch(type, btn) {
      await guard(async function () {
        var r;
        if (qr) {
          var used = qr; qr = null; showQr(); refreshBtns(); // QR 只用一次
          var q = await Promise.race([getPosition().catch(function () { return null; }), new Promise(function (res) { setTimeout(function () { res(null); }, 4000); })]);
          var args = { type: type, qr: used };
          if (q) { args.lat = q.coords.latitude; args.lng = q.coords.longitude; }
          r = await call('punch.do', args);
        } else {
          var p = await getPosition(); // 打卡當下重新定位
          r = await call('punch.do', { type: type, lat: p.coords.latitude, lng: p.coords.longitude });
        }
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
        return '<tr><td>' + esc(r.time) + '</td><td>' + esc(r.empId) + '</td><td>' + esc(r.name) + '</td><td>' + esc(r.type) + '</td><td>' + esc(r.place) + '</td><td>' + (r.distance === '' || r.distance == null ? '—' : esc(r.distance) + ' m') + '</td><td>' + badge(r.result) + '</td></tr>';
      }), '沒有資料') + '</div>';
    $('#mon').onchange = function () { if (this.value) { viewAttendance.month = this.value; viewAttendance(main); } };
  }

  /* ---------- 員工資料 ---------- */
  async function viewEmployees(main) {
    var list = await call('employees.list');
    var roleName = { admin: '管理員', manager: '主管', employee: '員工' };
    main.innerHTML = '<div class="row"><h2 style="margin:0">員工資料</h2><div class="spacer"></div><button id="tpl">下載 Excel 範本</button><button id="imp">Excel 批次匯入</button><button class="primary" id="add">新增員工</button></div><input type="file" id="impf" accept=".xlsx,.xls,.csv" class="hidden"><div class="card">' +
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

    $('#tpl').onclick = function () { guard(downloadEmpTemplate, this); };
    $('#imp').onclick = function () { $('#impf').value = ''; $('#impf').click(); };
    $('#impf').onchange = function () {
      var f = this.files[0];
      if (!f) return;
      guard(async function () {
        var rows = await readEmpFile(f);
        showImport(rows, f.name);
      });
    };
    function showImport(rows, fname) {
      var d = openDialog('<h3>批次匯入員工</h3><p>檔案:' + esc(fname) + ',共 <b>' + rows.length + '</b> 筆。員編已存在者會更新,其餘新增。</p>' +
        '<div class="tablewrap" style="max-height:240px;overflow:auto">' + table(['列', '員編', '姓名', '部門', '角色', '主管', '到職日', '狀態'], rows.slice(0, 20).map(function (r) {
          return '<tr><td>' + r.line + '</td><td>' + esc(r.id) + '</td><td>' + esc(r.name) + '</td><td>' + esc(r.dept) + '</td><td>' + esc(r.role) + '</td><td>' + esc(r.manager) + '</td><td>' + esc(r.hireDate) + '</td><td>' + esc(r.status) + '</td></tr>';
        })) + '</div>' + (rows.length > 20 ? '<p class="hint">僅預覽前 20 筆</p>' : '') +
        '<div id="impres"></div><div class="row" style="margin-top:12px"><button class="primary" id="impgo">開始匯入</button><button type="button" data-close>取消</button></div>');
      $('#impgo', d).onclick = function () {
        var btn = this;
        guard(async function () {
          var r = await call('employees.import', { rows: rows });
          var box = $('#impres', d);
          if (!r.applied) {
            box.innerHTML = '<p><span class="badge bad">未匯入</span> 有 ' + r.errors.length + ' 個問題,請修正 Excel 後重新匯入(全部通過才會寫入):</p>' +
              table(['列', '員編', '問題'], r.errors.map(function (e) { return '<tr><td>' + e.line + '</td><td>' + esc(e.id) + '</td><td class="wrap">' + esc(e.message) + '</td></tr>'; }));
            return;
          }
          btn.classList.add('hidden');
          box.innerHTML = '<p><span class="badge ok">匯入完成</span> 新增 ' + r.created.length + ' 位、更新 ' + r.updated + ' 位。</p>' +
            (r.created.length ? '<p class="hint">以下是新員工的登入密碼,只會顯示這一次,請下載後妥善發放,並提醒員工登入後修改。</p>' +
              table(['員編', '姓名', '初始密碼'], r.created.map(function (c) { return '<tr><td>' + esc(c.id) + '</td><td>' + esc(c.name) + '</td><td><code>' + esc(c.password) + '</code></td></tr>'; })) +
              '<div class="row" style="margin-top:8px"><button id="pwdl">下載帳號密碼清單</button></div>' : '');
          var dl = $('#pwdl', d);
          if (dl) dl.onclick = function () {
            guard(async function () {
              var X = await loadXLSX();
              var wb = X.utils.book_new();
              X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['員編', '姓名', '初始密碼']].concat(r.created.map(function (c) { return [c.id, c.name, c.password]; }))), '帳號密碼');
              X.writeFile(wb, '新進員工帳號密碼.xlsx');
            });
          };
          viewEmployees(main);
        }, btn);
      };
      d.addEventListener('close', function () { viewEmployees(main); }, { once: true });
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
    main.innerHTML = '<div class="row"><h2 style="margin:0">打卡地點(GPS 定點)</h2><div class="spacer"></div><button id="ltpl">下載 Excel 範本</button><button id="limp">Excel 批次匯入</button><button class="primary" id="add">新增地點</button></div><input type="file" id="limpf" accept=".xlsx,.xls,.csv" class="hidden"><div class="card">' +
      table(['名稱', '緯度', '經度', '半徑(公尺)', '啟用', ''], list.map(function (l) {
        return '<tr><td>' + esc(l.name) + '</td><td>' + esc(l.lat) + '</td><td>' + esc(l.lng) + '</td><td>' + esc(l.radius) + '</td><td>' + (l.active ? badge('成功').replace('成功', '啟用') : badge('撤回').replace('撤回', '停用')) +
          '</td><td>' + (l.active ? '<button class="sm" data-qr="' + esc(l.id) + '">顯示 QR</button> ' : '') + '<button class="sm" data-edit="' + esc(l.id) + '">編輯</button> <button class="sm danger" data-del="' + esc(l.id) + '">刪除</button></td></tr>';
      }), '尚未設定地點') + '</div>' +
      '<p class="hint">提示:新增地點時輸入名稱或地址再按「搜尋座標」即可自動取得座標;也可在現場按「使用目前位置」。儲存前請用地圖連結確認位置正確。</p>';

    function form(l) {
      var isNew = !l; l = l || { lat: '', lng: '', radius: 100, active: true };
      var d = openDialog('<h3>' + (isNew ? '新增地點' : '編輯地點') + '</h3><form id="lf" class="form">' +
        '<div class="full">' + field('l_name', '地點名稱', l.name || '', 'text', 'required placeholder="例如:寶雅 高雄巨蛋店"') + '</div>' +
        '<div class="full"><label for="l_q">地址 / 搜尋關鍵字(選填,留空則用名稱搜尋)</label><div class="row"><input id="l_q" style="flex:1" placeholder="例如:高雄市左營區博愛二路777號"><button type="button" id="l_find">搜尋座標</button></div></div>' +
        '<div class="full hidden" id="l_res"></div>' +
        field('l_lat', '緯度', l.lat, 'number', 'step="any"') + field('l_lng', '經度', l.lng, 'number', 'step="any"') +
        '<div class="full muted" id="l_chk"></div>' +
        field('l_rad', '允許半徑(公尺)', l.radius, 'number', 'min="10" max="5000" required') +
        '<div><label for="l_act">狀態</label><select id="l_act"><option value="1"' + (l.active ? ' selected' : '') + '>啟用</option><option value="0"' + (!l.active ? ' selected' : '') + '>停用</option></select></div>' +
        '<div class="full"><button type="button" id="l_here">使用目前位置</button></div>' +
        '<div class="full row"><button class="primary" type="submit">儲存</button><button type="button" data-close>取消</button></div></form>');

      function showCoord() {
        var la = val('l_lat'), ln = val('l_lng');
        $('#l_chk', d).innerHTML = la && ln ? '座標:' + esc(la) + ', ' + esc(ln) + ' — <a target="_blank" rel="noopener" href="https://www.google.com/maps?q=' + encodeURIComponent(la + ',' + ln) + '">在地圖上確認位置</a>' : '';
      }
      function pick(c) {
        document.getElementById('l_lat').value = Number(c.lat).toFixed(6);
        document.getElementById('l_lng').value = Number(c.lng).toFixed(6);
        $('#l_res', d).classList.add('hidden');
        showCoord();
        toast('已帶入座標,請確認地圖位置後儲存');
      }
      async function search() {
        var q = (val('l_q') || val('l_name')).trim();
        if (!q) return toast('請先輸入地點名稱或地址', true);
        var box = $('#l_res', d);
        box.classList.remove('hidden'); box.innerHTML = '<span class="muted">搜尋中…</span>';
        var list = [];
        try { list = await geocode(q); } catch (e) { box.classList.add('hidden'); throw e; }
        if (!list.length) { box.innerHTML = '<span class="badge bad">找不到</span> 請改用更完整的地址,或手動輸入座標'; return false; }
        if (list.length === 1) { pick(list[0]); return true; }
        box.innerHTML = '<div class="muted">找到多個結果,請選擇正確的位置:</div>' + list.map(function (c, i) {
          return '<div style="margin-top:6px"><button type="button" class="sm" data-pick="' + i + '">選擇</button> ' + esc(c.name) + '</div>';
        }).join('');
        box.querySelectorAll('[data-pick]').forEach(function (b) { b.onclick = function () { pick(list[+b.dataset.pick]); }; });
        return false;
      }
      showCoord();
      $('#l_lat', d).oninput = $('#l_lng', d).oninput = showCoord;
      $('#l_find', d).onclick = function () { guard(search, this); };
      $('#l_here', d).onclick = function () {
        guard(async function () { var p = await getPosition(); pick({ lat: p.coords.latitude, lng: p.coords.longitude }); toast('已帶入目前位置(精準度約 ' + Math.round(p.coords.accuracy) + ' 公尺)'); }, this);
      };
      $('#lf', d).onsubmit = function (ev) {
        ev.preventDefault();
        guard(async function () {
          if (!val('l_lat') || !val('l_lng')) { // 沒填座標:先自動搜尋,讓使用者確認後再儲存
            var ok = await search();
            if (!ok) toast('請確認座標後再按儲存', true);
            return;
          }
          await call('locations.save', { id: l.id, name: val('l_name').trim(), lat: Number(val('l_lat')), lng: Number(val('l_lng')), radius: Number(val('l_rad')), active: val('l_act') === '1' });
          d.close(); toast('已儲存'); viewLocations(main);
        }, $('#lf button[type=submit]', d));
      };
    }

    $('#ltpl').onclick = function () { guard(downloadLocTemplate, this); };
    $('#limp').onclick = function () { $('#limpf').value = ''; $('#limpf').click(); };
    $('#limpf').onchange = function () {
      var f = this.files[0];
      if (!f) return;
      guard(async function () { var rows = await readLocFile(f); await resolveAndPreview(rows, f.name); });
    };
    async function resolveAndPreview(rows, fname) {
      var d = openDialog('<h3>批次匯入打卡地點</h3><p>檔案:' + esc(fname) + ',共 <b>' + rows.length + '</b> 筆。</p><div id="lprog" class="muted">準備中…</div><div id="lbody"></div>' +
        '<div class="row" style="margin-top:12px"><button class="primary hidden" id="lgo">匯入</button><button type="button" data-close>關閉</button></div>');
      var cancelled = false;
      d.addEventListener('close', function () { cancelled = true; viewLocations(main); }, { once: true });
      var seenName = {};
      rows.forEach(function (r) { // 先做基本檢查,有問題的列不再搜尋
        var rad = r.radius === '' ? 100 : Number(r.radius);
        if (!r.name) r.pre = '名稱必填';
        else if (seenName[r.name]) r.pre = '名稱與第 ' + seenName[r.name] + ' 列重複';
        else if (!(rad >= 10 && rad <= 5000)) r.pre = '半徑需介於 10~5000 公尺';
        else if ((r.lat !== '' || r.lng !== '') && !(isFinite(Number(r.lat)) && isFinite(Number(r.lng)) && r.lat !== '' && r.lng !== '' && Math.abs(Number(r.lat)) <= 90 && Math.abs(Number(r.lng)) <= 180)) r.pre = '經緯度無效';
        if (r.name && !seenName[r.name]) seenName[r.name] = r.line;
      });
      var need = rows.filter(function (r) { return !r.pre && !(r.lat !== '' && r.lng !== ''); }).length, done = 0;
      for (var i = 0; i < rows.length && !cancelled; i++) {
        var r = rows[i];
        if (r.pre) { r.status = 'bad'; r.note = r.pre; continue; }
        if (r.lat !== '' && r.lng !== '') { r.status = 'ok'; r.note = '使用檔案中的座標'; continue; }
        var q = r.query || r.name;
        $('#lprog', d).textContent = '搜尋座標 ' + (++done) + ' / ' + need + ':' + q;
        if (!q) { r.status = 'bad'; r.note = '沒有名稱或地址可搜尋'; continue; }
        try {
          var list = await geocode(q);
          if (!list.length) { r.status = 'bad'; r.note = '找不到座標,請改填完整地址或直接填經緯度'; }
          else { r.lat = Number(list[0].lat).toFixed(6); r.lng = Number(list[0].lng).toFixed(6); r.status = list.length > 1 ? 'warn' : 'ok'; r.note = (list.length > 1 ? '有 ' + list.length + ' 個結果,已採用第一筆:' : '') + list[0].name; }
        } catch (e) { r.status = 'bad'; r.note = e.message; }
      }
      if (cancelled) return;
      var okRows = rows.filter(function (r) { return r.status !== 'bad'; });
      $('#lprog', d).textContent = '';
      $('#lbody', d).innerHTML = '<p>可匯入 <b>' + okRows.length + '</b> 筆;' + (rows.length - okRows.length) + ' 筆有問題將略過。<span class="hint">請點「地圖」確認自動搜尋到的位置是否正確(黃色為有多個結果)。</span></p>' +
        '<div class="tablewrap" style="max-height:300px;overflow:auto">' + table(['列', '名稱', '座標', '半徑', '結果', ''], rows.map(function (r) {
          var cls = r.status === 'ok' ? 'ok' : r.status === 'warn' ? 'warn' : 'bad';
          return '<tr><td>' + r.line + '</td><td>' + esc(r.name) + '</td><td>' + (r.status === 'bad' ? '' : esc(r.lat) + ', ' + esc(r.lng)) + '</td><td>' + esc(r.radius || 100) + '</td><td class="wrap"><span class="badge ' + cls + '">' + (r.status === 'ok' ? '正常' : r.status === 'warn' ? '請確認' : '略過') + '</span> ' + esc(r.note) + '</td><td>' +
            (r.status === 'bad' ? '' : '<a target="_blank" rel="noopener" href="https://www.google.com/maps?q=' + encodeURIComponent(r.lat + ',' + r.lng) + '">地圖</a>') + '</td></tr>';
        })) + '</div>';
      var go = $('#lgo', d);
      if (!okRows.length) return;
      go.textContent = '匯入 ' + okRows.length + ' 筆'; go.classList.remove('hidden');
      go.onclick = function () {
        guard(async function () {
          var res = await call('locations.import', { rows: okRows.map(function (r) { return { line: r.line, name: r.name, lat: Number(r.lat), lng: Number(r.lng), radius: r.radius, active: r.active }; }) });
          if (!res.applied) {
            $('#lbody', d).innerHTML = '<p><span class="badge bad">未匯入</span> 後端驗證有問題:</p>' + table(['列', '名稱', '問題'], res.errors.map(function (e) { return '<tr><td>' + e.line + '</td><td>' + esc(e.name) + '</td><td class="wrap">' + esc(e.message) + '</td></tr>'; }));
            return;
          }
          go.classList.add('hidden');
          $('#lbody', d).innerHTML = '<p><span class="badge ok">匯入完成</span> 新增 ' + res.created + ' 個、更新 ' + res.updated + ' 個,略過 ' + (rows.length - okRows.length) + ' 個。</p>';
        }, go);
      };
    }
    $('#add').onclick = function () { form(null); };
    document.querySelectorAll('[data-edit]').forEach(function (b) { b.onclick = function () { form(list.filter(function (x) { return x.id === b.dataset.edit; })[0]); }; });
    document.querySelectorAll('[data-qr]').forEach(function (b) { b.onclick = function () { showLocQr(list.filter(function (x) { return x.id === b.dataset.qr; })[0]); }; });
    document.querySelectorAll('[data-del]').forEach(function (b) {
      b.onclick = function () {
        if (!confirm('確定刪除這個打卡地點?')) return;
        guard(async function () { await call('locations.delete', { id: b.dataset.del }); toast('已刪除'); viewLocations(main); }, b);
      };
    });
  }


  /* ---------- 登入紀錄 ---------- */
  function loginRows(list) {
    return table(['時間', '員編', '姓名', '結果', 'IP', '裝置'], list.map(function (l) {
      var cls = l.result === '成功' ? 'ok' : 'bad';
      return '<tr><td>' + esc(l.time) + '</td><td>' + esc(l.empId) + '</td><td>' + esc(l.name) + '</td><td><span class="badge ' + cls + '">' + esc(l.result) + '</span></td><td>' + esc(l.ip) + '</td><td title="' + esc(l.ua) + '">' + esc(shortUA(l.ua)) + '</td></tr>';
    }), '沒有紀錄');
  }
  async function viewLoginLog(main) {
    var month = (viewLoginLog.month = viewLoginLog.month || curMonth()), emp = viewLoginLog.emp || '';
    var list = await call('login.log', { month: month, empId: emp });
    var fail = list.filter(function (l) { return l.result !== '成功'; }).length;
    main.innerHTML = '<div class="row"><h2 style="margin:0">登入紀錄</h2><div class="spacer"></div><input type="month" id="mon" value="' + esc(month) + '" style="width:auto"><input id="emp" placeholder="員編(可空白)" value="' + esc(emp) + '" style="width:130px"><button id="q">查詢</button></div>' +
      '<div class="card"><p class="muted">共 ' + list.length + ' 筆(最多顯示 500 筆),其中失敗 / 鎖定 ' + fail + ' 筆。IP 由瀏覽器回報,僅供參考。</p>' + loginRows(list) + '</div>';
    $('#q').onclick = function () { viewLoginLog.month = val('mon') || curMonth(); viewLoginLog.emp = val('emp').trim(); viewLoginLog(main); };
  }

  /* ---------- 修改密碼 ---------- */
  async function viewPassword(main) {
    main.innerHTML = '<h2>修改密碼</h2><div class="card" style="max-width:420px"><form id="pf" class="form" style="grid-template-columns:1fr">' +
      field('p0', '目前密碼', '', 'password', 'autocomplete="current-password" required') + field('p1', '新密碼(至少 8 碼)', '', 'password', 'autocomplete="new-password" minlength="8" required') +
      field('p2', '再輸入一次新密碼', '', 'password', 'autocomplete="new-password" required') + '<button class="primary" type="submit">更新密碼</button></form></div>' +
      '<div class="card"><h3>我的近期登入紀錄</h3><p class="hint">若看到不是您本人的登入(陌生 IP 或時間),請立即修改密碼並通知人資。</p><div id="mylog" class="muted">載入中…</div></div>';
    $('#pf').onsubmit = function (e) {
      e.preventDefault();
      if (val('p1') !== val('p2')) return toast('兩次輸入的新密碼不一致', true);
      guard(async function () { await call('changePassword', { oldPassword: val('p0'), newPassword: val('p1') }); toast('密碼已更新'); $('#pf').reset(); }, $('#pf button'));
    };
    try { $('#mylog').outerHTML = '<div id="mylog">' + loginRows(await call('login.my')) + '</div>'; } catch (e) { $('#mylog').textContent = e.message; }
  }

  /* ---------- 啟動 ---------- */
  document.title = cfg.COMPANY_NAME || '人事管理系統';
  captureQr();
  if (API.hasToken()) {
    call('me').then(function (u) { user = u; renderShell(); }).catch(function () { API.logout(); renderLogin(); });
  } else renderLogin();
})();
