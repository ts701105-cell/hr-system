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
    var cls = { '核准': 'ok', '成功': 'ok', '已錄取': 'ok', '已通過': 'ok', '已報名': 'warn', '待簽核': 'warn', '待審核': 'warn', '已退件': 'bad', '駁回': 'bad', '超出範圍': 'bad', '撤回': '' }[status] || '';
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
      ['初始密碼', '選填,至少 6 碼(可用民國生日 7 碼,例如 0730518)。新員工留空會自動產生隨機密碼;既有員工留空則不改密碼']]);
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
    app.innerHTML = '<div class="login card"><h1>' + esc(cfg.COMPANY_NAME || '東慧公司岡山區人事管理系統') + (API.demo ? '<span class="demo-tag">示範模式</span>' : '') + '</h1>' +
      '<form id="lf" class="form" style="grid-template-columns:1fr">' + field('li', '帳號(姓名)', '', 'text', 'autocomplete="username" required') +
      field('lp', '密碼', '', 'password', 'autocomplete="current-password" required') +
      '<button class="primary" type="submit">登入</button></form>' +
      '<p class="hint"><b>帳號</b>:您的姓名。<b>初始密碼</b>:民國出生年月日共 7 碼(年 3 碼 + 月 2 碼 + 日 2 碼),例如民國 73 年 5 月 18 日請輸入 <b>0730518</b>。如忘記密碼請聯絡人資重設。</p>' +
      '<p class="hint"><b>新人?</b> <a href="#" id="goreg">新人註冊(上傳證件)</a> ｜ <a href="#" id="gostat">查詢審核進度</a></p>' +
      '<p class="hint" id="myip">正在取得您的 IP…</p><p class="hint">登入時系統會記錄時間與 IP 位址,供安全稽核使用。</p>' +
      (API.demo ? '<p class="hint">示範帳號:管理員 A001 / admin1234;主管 M001 / demo1234;員工 E001 / demo1234。資料只存在這個瀏覽器。</p>' : '') + '</div>';
    getIP().then(function (ip) { var el = document.getElementById('myip'); if (el) el.textContent = ip ? '您目前的 IP:' + ip : '無法取得 IP(仍可登入,紀錄會標示為「未知」)'; });
    $('#goreg').onclick = function (e) { e.preventDefault(); renderRegister(); };
    $('#gostat').onclick = function (e) { e.preventDefault(); renderRegStatus(); };
    $('#lf').onsubmit = function (e) {
      e.preventDefault();
      guard(async function () {
        var ip = await getIP();
        var r = await API.login(val('li').trim(), val('lp'), { ip: ip, ua: navigator.userAgent });
        if (r.choose) { // 同名同密碼(同一個人在多個單位):選要進哪個單位
          r = await new Promise(function (resolve, reject) {
            var d = openDialog('<h3>請選擇要登入的單位</h3><p class="muted">您在系統裡有多個單位的資料,請選擇這次要使用哪一個。</p>' +
              r.options.map(function (o, i) { return '<div style="margin:8px 0"><button class="primary" data-pk="' + i + '" style="width:100%">' + esc(o.dept || '未設定單位') + '(員編 ' + esc(o.empId) + ')</button></div>'; }).join('') +
              '<div class="row" style="margin-top:12px"><button type="button" data-close>取消</button></div>');
            var opts = r.options, ticket = r.ticket, done = false;
            d.querySelectorAll('[data-pk]').forEach(function (b) {
              b.onclick = function () {
                done = true;
                API.loginPick(ticket, opts[+b.dataset.pk].empId).then(function (x) { d.close(); resolve(x); }, function (e) { d.close(); reject(e); });
              };
            });
            d.addEventListener('close', function () { if (!done) reject(new Error('已取消登入')); }, { once: true });
          });
        }
        if (r.mfa) r = await askMfa(r.ticket); // 有啟用二階段認證:再輸入手機驗證碼
        user = r.user;
        location.hash = '#/punch';
        renderShell();
        var msg = '登入 IP:' + (ip || '未知') + (r.lastLogin ? '。上次登入:' + r.lastLogin.time + '(' + r.lastLogin.ip + ')' : '。這是您第一次登入');
        if (r.mustChange) { toast('請先修改預設密碼', true); location.hash = '#/password'; } else toast(msg);
      }, $('#lf button'));
    };
  }

  /* ---------- 新人註冊(免登入) ---------- */
  function readFileB64(file, label) {
    return new Promise(function (resolve, reject) {
      if (!file) return reject(new Error('請上傳「' + label + '」'));
      if (file.type === 'application/pdf') {
        if (file.size > 4 * 1024 * 1024) return reject(new Error(label + ' PDF 太大(上限 4MB)'));
        var fr = new FileReader();
        fr.onload = function () { resolve(String(fr.result).split(',')[1]); };
        fr.onerror = function () { reject(new Error(label + ' 讀取失敗')); };
        return fr.readAsDataURL(file);
      }
      if (!/^image\//.test(file.type)) return reject(new Error(label + ' 只接受圖片或 PDF'));
      var img = new Image(), url = URL.createObjectURL(file);
      img.onload = function () { // 縮圖壓成 JPEG,手機原圖動輒 5MB 以上
        var max = 1800, k = Math.min(1, max / Math.max(img.width, img.height));
        var cv = document.createElement('canvas'); cv.width = Math.round(img.width * k); cv.height = Math.round(img.height * k);
        cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
        URL.revokeObjectURL(url);
        resolve(cv.toDataURL('image/jpeg', 0.85).split(',')[1]);
      };
      img.onerror = function () { URL.revokeObjectURL(url); reject(new Error(label + ' 無法讀取,請改用 JPG / PNG / PDF')); };
      img.src = url;
    });
  }

  function renderRegister() {
    var units = ['寶雅', '鷹世服'];
    app.innerHTML = '<div class="login card" style="max-width:520px"><h1>新人註冊</h1>' +
      '<p class="hint">送出後由人資審核,通過才會開通帳號。審核通過後帳號是您的<b>姓名</b>、初始密碼是<b>民國出生年月日 7 碼</b>。</p>' +
      '<form id="rf" class="form" style="grid-template-columns:1fr">' +
      field('rn', '姓名', '', 'text', 'required maxlength="20"') + field('rid', '身分證字號', '', 'text', 'required maxlength="10" autocomplete="off" style="text-transform:uppercase"') +
      field('rb', '出生日期', '', 'date', 'required') + field('rp', '手機號碼', '', 'tel', 'required maxlength="10" placeholder="09xxxxxxxx"') +
      field('re', 'Email', '', 'email', 'required') +
      '<div><label for="ru">申請單位</label><select id="ru" required><option value="">請選擇</option>' + units.map(function (u) { return '<option>' + u + '</option>'; }).join('') + '</select></div>' +
      '<div><label for="f1">身分證正面(圖片或 PDF)</label><input id="f1" type="file" accept="image/*,application/pdf" required></div>' +
      '<div><label for="f2">身分證反面(圖片或 PDF)</label><input id="f2" type="file" accept="image/*,application/pdf" required></div>' +
      '<div><label for="f3">郵局或國泰世華銀行存摺封面(圖片或 PDF)</label><input id="f3" type="file" accept="image/*,application/pdf" required></div>' +
      '<label style="display:flex;gap:8px;align-items:flex-start"><input id="rc" type="checkbox" style="width:auto;margin-top:4px"><span class="hint" style="margin:0">我已閱讀並同意:所填資料與證件僅供東慧公司辦理人員報到、勞健保加退保及薪資轉帳使用,由人資專人保管,不會公開;審核完成後證件檔案會刪除。</span></label>' +
      '<button class="primary" type="submit">送出申請</button></form>' +
      '<p class="hint"><a href="#" id="rback">← 回登入</a></p></div>';
    $('#rback').onclick = function (e) { e.preventDefault(); renderLogin(); };
    $('#rf').onsubmit = function (e) {
      e.preventDefault();
      guard(async function () {
        var files = {};
        files.front = { data: await readFileB64($('#f1').files[0], '身分證正面') };
        files.back = { data: await readFileB64($('#f2').files[0], '身分證反面') };
        files.bank = { data: await readFileB64($('#f3').files[0], '存摺封面') };
        var r = await API.call('register.submit', { name: val('rn').trim(), idNo: val('rid').trim(), birthday: val('rb'), phone: val('rp').trim(), email: val('re').trim(), unit: val('ru'), consent: $('#rc').checked, files: files });
        app.innerHTML = '<div class="login card"><h1>已送出申請</h1><p>申請編號:<b>' + esc(r.appId) + '</b></p><p class="hint">人資審核通過後,即可用「姓名」+「民國出生年月日 7 碼」登入。可隨時在登入頁用身分證字號與生日「查詢審核進度」。</p><button class="primary" id="rdone">回登入頁</button></div>';
        $('#rdone').onclick = renderLogin;
      }, $('#rf button'));
    };
  }

  function renderRegStatus() {
    app.innerHTML = '<div class="login card"><h1>查詢審核進度</h1><form id="sf" class="form" style="grid-template-columns:1fr">' +
      field('si', '身分證字號', '', 'text', 'required maxlength="10" style="text-transform:uppercase"') + field('sb', '出生日期', '', 'date', 'required') +
      '<button class="primary" type="submit">查詢</button></form><div id="sres"></div><p class="hint"><a href="#" id="sback">← 回登入</a></p></div>';
    $('#sback').onclick = function (e) { e.preventDefault(); renderLogin(); };
    $('#sf').onsubmit = function (e) {
      e.preventDefault();
      guard(async function () {
        var r = await API.call('register.status', { idNo: val('si').trim(), birthday: val('sb') });
        $('#sres').innerHTML = '<p>申請編號:' + esc(r.appId) + '(' + esc(r.time) + ')</p><p>狀態:' + badge(r.status) + '</p>' +
          (r.status === '已退件' ? '<p class="hint">退件原因:' + esc(r.note) + '。如需重新申請,請修正後再送出。</p>' : '') +
          (r.status === '已通過' ? '<p class="hint">已開通!帳號是您的姓名,初始密碼是民國出生年月日 7 碼。</p>' : '');
      }, $('#sf button'));
    };
  }

  /* ---------- 取消紀錄(管理員) ---------- */
  async function viewCancelLog(main) {
    var month = (viewCancelLog.month = viewCancelLog.month || curMonth());
    var list = await call('cancel.log', { month: month });
    main.innerHTML = '<div class="row"><h2 style="margin:0">取消紀錄</h2><div class="spacer"></div><input type="month" id="mon" value="' + esc(month) + '" style="width:auto"></div>' +
      '<p class="hint">員工自行取消班別時填寫的原因。月份依「班的日期」篩選。</p><div class="card">' +
      table(['取消時間', '班的日期', '員工', '單位', '班別', '原因'], list.map(function (r) {
        return '<tr><td>' + esc(r.time) + '</td><td>' + esc(r.date) + '</td><td>' + esc(r.name) + '(' + esc(r.empId) + ')</td><td>' + esc(r.unit) + '</td><td>' + esc(r.shift) + '</td><td>' + esc(r.reason) + '</td></tr>';
      }), '這個月沒有取消紀錄') + '</div>';
    $('#mon').onchange = function () { if (this.value) { viewCancelLog.month = this.value; viewCancelLog(main).catch(function (e) { toast(e.message, true); }); } };
  }

  /* ---------- 註冊審核(管理員) ---------- */
  async function viewRegister(main) {
    var st = viewRegister.st == null ? '待審核' : viewRegister.st;
    var list = await call('register.list', { status: st });
    main.innerHTML = '<div class="row"><h2 style="margin:0">註冊審核</h2><div class="spacer"></div><select id="rst" style="width:auto">' +
      ['待審核', '已通過', '已退件', ''].map(function (x) { return '<option value="' + x + '"' + (x === st ? ' selected' : '') + '>' + (x || '全部') + '</option>'; }).join('') + '</select></div>' +
      '<p class="hint">新人在登入頁送出申請後會出現在這裡。請核對證件與填寫資料,可填入員編(不填會先給暫編),系統會自動建立帳號(帳號=姓名、密碼=民國生日 7 碼)。</p>' +
      '<div class="card">' + table(['申請時間', '姓名', '單位', '電話', '狀態', ''], list.map(function (r, i) {
        return '<tr><td>' + esc(r.time) + '</td><td>' + esc(r.name) + '</td><td>' + esc(r.unit) + '</td><td>' + esc(r.phone) + '</td><td>' + badge(r.status) + '</td><td><button class="sm primary" data-v="' + i + '">檢視</button></td></tr>';
      }), '沒有申請') + '</div>';
    $('#rst').onchange = function () { viewRegister.st = this.value; viewRegister(main).catch(function (e) { toast(e.message, true); }); };
    document.querySelectorAll('[data-v]').forEach(function (b) {
      b.onclick = function () { regDetail(list[+b.dataset.v], function () { viewRegister(main).catch(function (e) { toast(e.message, true); }); }); };
    });
  }

  function rocDate(iso) {
    var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? '民國 ' + (+m[1] - 1911) + ' 年 ' + (+m[2]) + ' 月 ' + (+m[3]) + ' 日' : String(iso || '');
  }
  function regDetail(r, done) {
    var pend = r.status === '待審核';
    openDialog('<h3>' + esc(r.name) + ' · ' + esc(r.unit) + ' ' + badge(r.status) + '</h3>' +
      '<p class="muted">申請編號 ' + esc(r.appId) + '(' + esc(r.time) + ')<br>身分證字號:<b>' + esc(r.idNo) + '</b>　生日:' + esc(rocDate(r.birthday)) + '<br>電話:' + esc(r.phone) + '　Email:' + esc(r.email) +
      (r.reviewer ? '<br>審核:' + esc(r.reviewer) + '(' + esc(r.reviewTime) + ')' + (r.empId ? ',員編 ' + esc(r.empId) : '') + (r.note ? ',原因:' + esc(r.note) : '') : '') + '</p>' +
      (r.hasFiles ? '<div class="row"><button class="sm" data-k="front">身分證正面</button><button class="sm" data-k="back">身分證反面</button><button class="sm" data-k="bank">存摺封面</button></div><div id="rprev" style="margin:10px 0;text-align:center"></div>' : '<p class="hint">證件檔案已刪除。</p>') +
      (pend ? '<div class="form"><div><label>員編(選填,由管理員輸入;不填會先給暫編 TMP001…)</label><input id="aemp" placeholder="例如 164958"></div><div><label>退件原因(退件時必填)</label><input id="arsn"></div></div>' +
        '<div class="row"><button class="primary" id="aok">通過並建立帳號</button><button class="danger" id="ano">退件</button><button data-close>關閉</button></div>'
        : '<div class="row">' + (r.hasFiles ? '<button class="danger" id="apurge">刪除證件檔案</button>' : '') + '<button data-close>關閉</button></div>'));
    document.querySelectorAll('[data-k]').forEach(function (b) {
      b.onclick = function () {
        guard(async function () {
          $('#rprev').textContent = '載入中…';
          var f = await call('register.file', { appId: r.appId, key: b.dataset.k });
          var src = 'data:' + f.mime + ';base64,' + f.data;
          $('#rprev').innerHTML = f.mime === 'application/pdf' ? '<iframe src="' + src + '" style="width:100%;height:480px;border:1px solid var(--line)"></iframe>' : '<img src="' + src + '" alt="' + esc(f.label) + '" style="max-width:100%;max-height:480px">';
        }, b);
      };
    });
    if (pend) {
      $('#aok').onclick = function () {
        guard(async function () {
          var x = await call('register.approve', { appId: r.appId, empId: val('aemp').trim() });
          dlg.close();
          alert('已通過並建立帳號\n帳號:' + x.account + '(員編 ' + x.empId + ')\n初始密碼:' + x.password + '(民國生日)');
          done();
        }, $('#aok'));
      };
      $('#ano').onclick = function () {
        guard(async function () { await call('register.reject', { appId: r.appId, reason: val('arsn').trim() }); dlg.close(); toast('已退件'); done(); }, $('#ano'));
      };
    } else if (r.hasFiles) {
      $('#apurge').onclick = function () {
        if (!confirm('確定刪除這份申請的身分證與存摺檔案?(移到雲端硬碟垃圾桶)')) return;
        guard(async function () { await call('register.purge', { appId: r.appId }); dlg.close(); toast('已刪除證件檔案'); done(); }, $('#apurge'));
      };
    }
  }

  /* ---------- 外框與路由 ---------- */
  var ROUTES = [
    { path: 'punch', label: '打卡', roles: ['employee', 'manager', 'admin'], view: viewPunch },
    { path: 'leave', label: '請假', roles: ['employee', 'manager', 'admin'], view: viewLeave },
    { path: 'schedule', label: '排班狀態', roles: ['employee', 'manager', 'admin'], view: viewSchedule },
    { path: 'signup', label: '排班報名', roles: ['employee', 'manager', 'admin'], view: viewSignup },
    { path: 'approvals', label: '假單簽核', roles: ['manager', 'admin'], group: '主管', view: viewApprovals },
    { path: 'attendance', label: '出勤報表', roles: ['manager', 'admin'], group: '主管', view: viewAttendance },
    { path: 'employees', label: '員工資料', roles: ['admin'], group: '管理', view: viewEmployees },
    { path: 'register', label: '註冊審核', roles: ['admin'], group: '管理', view: viewRegister },
    { path: 'locations', label: '打卡地點', roles: ['admin'], group: '管理', view: viewLocations },
    { path: 'cancellog', label: '取消紀錄', roles: ['admin'], group: '管理', view: viewCancelLog },
    { path: 'loginlog', label: '登入紀錄', roles: ['admin'], group: '管理', view: viewLoginLog },
    { path: 'password', label: '帳號安全', roles: ['employee', 'manager', 'admin'], group: '帳號', view: viewPassword }
  ];

  function renderShell() {
    if (!user) return renderLogin();
    var roleName = { admin: '管理員', manager: '主管', employee: '員工' }[user.role];
    app.innerHTML = '<div class="topbar"><div class="brand">' + esc(cfg.COMPANY_NAME || '東慧公司岡山區人事管理系統') + (API.demo ? '<span class="demo-tag">示範</span>' : '') + '</div>' +
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
    var allEmps = res[2].filter(function (e) { return e.status === '在職' && (user.role === 'admin' || e.id !== user.id); });
    /* 派駐單位 = 員工的「部門」欄(例如 寶雅、鷹世服);可依單位篩選員工,並只顯示該單位可用的班別 */
    var units = [];
    allEmps.forEach(function (e) { if (e.dept && units.indexOf(e.dept) < 0) units.push(e.dept); });
    var unit = units.indexOf(viewSchedule.unit) >= 0 ? viewSchedule.unit : '';
    var emps = unit ? allEmps.filter(function (e) { return e.dept === unit; }) : allEmps;
    if (unit) types = types.filter(function (t) { return !t.unit || t.unit === unit; });
    var n = daysIn(month), WD = ['日', '一', '二', '三', '四', '五', '六'];
    var map = {}; shifts.forEach(function (s) { map[s.date + '|' + s.empId] = s.shift; });
    var dirty = {};

    var unitSel = (canEdit && units.length) ? '<select id="unit" style="width:auto"><option value="">全部單位</option>' +
      units.map(function (u) { return '<option' + (u === unit ? ' selected' : '') + '>' + esc(u) + '</option>'; }).join('') + '</select>' : '';
    var head = '<div class="row"><h2 style="margin:0">排班狀態</h2><div class="spacer"></div>' + unitSel + '<button class="sm" id="pm">◀</button><input type="month" id="mon" value="' + esc(month) + '" style="width:auto"><button class="sm" id="nm">▶</button>' +
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
      body = '<p class="hint">上班前 30 分鐘以前都可以按「取消」,取消時必須填寫原因,取消後名額會還給別人;超過時限請聯絡主管。</p><div class="card">' + table(['日期', '星期', '班別', '上班', '下班', ''], mine.map(function (s, i) {
        var stm = String(s.start || '').match(/(\d{1,2}):(\d{2})/);
        var open = stm && new Date(s.date + 'T' + ('0' + stm[1]).slice(-2) + ':' + stm[2] + ':00+08:00').getTime() - 30 * 60000 > Date.now();
        var btn = open ? '<button class="sm danger" data-x="' + i + '">取消</button>' : '';
        return '<tr><td>' + esc(s.date) + '</td><td>' + WD[new Date(s.date).getDay()] + '</td><td>' + esc(s.shift) + '</td><td>' + esc(s.start) + '</td><td>' + esc(s.end) + '</td><td>' + btn + '</td></tr>';
      }), '本月尚無排班') + '</div>';
      setTimeout(function () {
        document.querySelectorAll('[data-x]').forEach(function (b) {
          b.onclick = function () {
            var s = mine[+b.dataset.x];
            var reason = prompt('取消 ' + s.date + ' 的「' + s.shift + '」,班表上的名字會一併移除。\n請填寫取消原因(必填):');
            if (reason == null) return;
            if (reason.trim().length < 2) return toast('請填寫取消原因', true);
            guard(async function () { await call('shifts.cancel', { date: s.date, reason: reason.trim() }); toast('已取消'); viewSchedule(main).catch(function (e) { toast(e.message, true); }); }, b);
          };
        });
      }, 0);
    }
    main.innerHTML = head + body;

    function go(m) { viewSchedule.month = m; viewSchedule(main).catch(function (e) { toast(e.message, true); }); }
    $('#mon').onchange = function () { if (this.value) go(this.value); };
    if ($('#unit')) $('#unit').onchange = function () { viewSchedule.unit = this.value; go(month); };
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

  /* ---------- 排班報名 ---------- */
  async function viewSignup(main) {
    var month = (viewSignup.month = viewSignup.month || curMonth());
    var mgr = user.role !== 'employee';
    var WD = ['日', '一', '二', '三', '四', '五', '六'];
    function dlabel(d) { return esc(d.slice(5).replace('-', '/')) + '(' + WD[new Date(d).getDay()] + ')'; }
    function tlabel(s) { return esc(s.shift) + (s.start ? '(' + esc(s.start) + '–' + esc(s.end) + ')' : ''); }
    var bar = '<div class="row"><h2 style="margin:0">排班報名</h2><div class="spacer"></div><button class="sm" id="pm">◀</button><input type="month" id="mon" value="' + esc(month) + '" style="width:auto"><button class="sm" id="nm">▶</button></div>';
    function go(m) { viewSignup.month = m; viewSignup(main).catch(function (e) { toast(e.message, true); }); }
    function bindBar() {
      $('#mon').onchange = function () { if (this.value) go(this.value); };
      $('#pm').onclick = function () { var p = month.split('-'); go(fmtDate(new Date(+p[0], +p[1] - 2, 1)).slice(0, 7)); };
      $('#nm').onclick = function () { var p = month.split('-'); go(fmtDate(new Date(+p[0], +p[1], 1)).slice(0, 7)); };
    }

    if (!mgr) {
      var slots = await call('signup.open', { month: month });
      var hint = user.dept ? '您的所屬單位:<b>' + esc(user.dept) + '</b>。同一天只能報名一個班;寶雅依開放時間報名(優先名單較早開放)。' : '您尚未設定所屬單位,請聯絡人資。';
      main.innerHTML = bar + '<p class="hint">' + hint + '</p><div class="card">' + table(['日期', '班別', '單位', '尚餘缺額', '狀態', ''], slots.map(function (s, i) {
        var full = s.quota <= 0;
        var st = s.mine === '已錄取' ? badge('已錄取') : s.mine ? badge('已報名') : full ? '<span class="badge bad">額滿</span>' : '';
        var btn = s.mine ? '<button class="sm danger" data-c="' + i + '">取消報名</button>' : (!s.mine && !full ? (s.open === false ? '<span class="muted">' + esc(s.openAt) + ' 開放</span>' : '<button class="sm primary" data-a="' + i + '">報名</button>') : '');
        return '<tr><td>' + dlabel(s.date) + '</td><td>' + tlabel(s) + '</td><td>' + esc(s.unit) + '</td><td>' + s.quota + '</td><td>' + st + '</td><td>' + btn + '</td></tr>';
      }), '這個月目前沒有開放報名的班') + '</div>';
      bindBar();
      function act(sel, action, msg) {
        document.querySelectorAll(sel).forEach(function (b) {
          b.onclick = function () {
            var s = slots[+(b.dataset.a || b.dataset.c)];
            var reason = '';
            if (action === 'signup.cancel') {
              reason = prompt('取消 ' + s.date + ' 的「' + s.shift + '」,班表上的名字會一併移除(上班前 30 分鐘前才能取消)。\n請填寫取消原因(必填):');
              if (!reason || reason.trim().length < 2) { if (reason !== null) toast('請填寫取消原因', true); return; }
              reason = reason.trim();
            }
            guard(async function () {
              await call(action, { date: s.date, unit: s.unit, shift: s.shift, reason: reason });
              toast(msg); viewSignup(main);
            }, b);
          };
        });
      }
      act('[data-a]', 'signup.apply', '報名成功');
      act('[data-c]', 'signup.cancel', '已取消報名');
      return;
    }

    var res = await Promise.all([call('signup.open', { month: month }), call('signup.list', { month: month }), call('shifts.types'), call('employees.list')]);
    var qslots = res[0], list = res[1], types = res[2].filter(function (t) { return t.start; }), emps = res[3];
    var units = [];
    emps.forEach(function (e) { if (e.dept && units.indexOf(e.dept) < 0) units.push(e.dept); });
    qslots.forEach(function (s) { if (units.indexOf(s.unit) < 0) units.push(s.unit); });
    var dirty = {};

    var syncBar = user.role === 'admin' ? '<div class="row" style="gap:8px;margin-bottom:8px"><button id="sync">從雲端硬碟同步 ' + esc(month) + ' 班表</button>' +
      '<span class="hint">讀取寶雅、鷹世服班表,更新「排班狀態」與「名額」(這個按鈕只讀取班表,不會改動檔案)</span></div>' : '';
    var quotaCard = '<div class="card"><h3>名額設定</h3>' + syncBar + table(['日期', '班別', '單位', '已報名', '尚餘缺額'], qslots.map(function (s, i) {
      return '<tr><td>' + dlabel(s.date) + '</td><td>' + tlabel(s) + '</td><td>' + esc(s.unit) + '</td><td>' + s.count + '</td><td><input type="number" min="0" max="500" value="' + s.quota + '" data-q="' + i + '" style="width:80px"></td></tr>';
    }), '這個月還沒有設定名額') + '<div class="row" style="margin-top:8px"><div class="spacer"></div><button class="primary" id="qsv" disabled>儲存名額(填 0 = 移除)</button></div>' +
      '<h3 style="margin-top:16px">新增名額</h3><div class="row" style="flex-wrap:wrap;gap:8px;align-items:flex-end">' +
      '<div><label for="n_d">日期</label><input id="n_d" type="date" value="' + esc(month) + '-01"></div>' +
      '<div><label for="n_u">單位</label><input id="n_u" list="n_ul" value="' + esc(units[0] || '') + '"><datalist id="n_ul">' + units.map(function (u) { return '<option value="' + esc(u) + '">'; }).join('') + '</datalist></div>' +
      '<div><label for="n_s">班別</label><select id="n_s">' + types.map(function (t) { return '<option value="' + esc(t.name) + '">' + esc(t.name) + '(' + esc(t.start) + '–' + esc(t.end) + ')</option>'; }).join('') + '</select></div>' +
      '<div><label for="n_q">需求人數</label><input id="n_q" type="number" min="1" max="500" value="5" style="width:90px"></div>' +
      '<button class="primary" id="nadd">新增</button></div></div>';

    var listCard = '<div class="card"><h3>報名名單(報名即錄取)</h3>' +
      table(['日期', '班別', '單位', '員工', '狀態', '報名時間'], list.map(function (r) {
        return '<tr><td>' + dlabel(r.date) + '</td><td>' + esc(r.shift) + '</td><td>' + esc(r.unit) + '</td><td>' + esc(r.name) + '(' + esc(r.empId) + ')</td><td>' + badge(r.status) + '</td><td>' + esc(r.time) + '</td></tr>';
      }), '這個月還沒有人報名') + '</div>';
    main.innerHTML = bar + quotaCard + listCard;
    bindBar();

    if ($('#sync')) $('#sync').onclick = function () {
      guard(async function () {
        var lines = [], bad = false;
        for (var unit of ['寶雅', '鷹世服']) {
          toast('正在同步 ' + unit + '…(約需十幾秒)');
          try {
            var r = await call('sync.drive', { month: month, unit: unit });
            lines.push('<li><b>' + esc(unit) + '</b>:' + (r.note ? esc(r.note) : '排班 ' + r.shifts + ' 筆、名額 ' + r.slots + ' 筆(分頁 ' + esc(r.sheet) + ')') +
              (r.unmatched && r.unmatched.length ? '<br><span style="color:#b45309">找不到對應員工、未寫入:' + esc(r.unmatched.join('、')) + '</span>' : '') + '</li>');
          } catch (e) { bad = true; lines.push('<li><b>' + esc(unit) + '</b>:<span style="color:#b91c1c">' + esc(e.message) + '</span></li>'); }
        }
        openDialog('<h3>雲端硬碟同步結果</h3><ul>' + lines.join('') + '</ul><p class="hint">「找不到對應員工」請先把人加進員工資料(姓名要和班表一致)再同步一次。</p><div class="row"><div class="spacer"></div><button data-close>關閉</button></div>');
        if (!bad) viewSignup(main);
      }, this);
    };
    document.querySelectorAll('[data-q]').forEach(function (inp) {
      inp.oninput = function () {
        var s = qslots[+inp.dataset.q], k = s.date + '|' + s.unit + '|' + s.shift;
        if (Number(inp.value) === s.quota) delete dirty[k]; else dirty[k] = { date: s.date, unit: s.unit, shift: s.shift, quota: Number(inp.value) || 0 };
        $('#qsv').disabled = !Object.keys(dirty).length;
      };
    });
    $('#qsv').onclick = function () {
      guard(async function () {
        var items = Object.keys(dirty).map(function (k) { return dirty[k]; });
        await call('signup.quota.save', { items: items });
        toast('已儲存名額'); viewSignup(main);
      }, this);
    };
    $('#nadd').onclick = function () {
      guard(async function () {
        if (!val('n_d') || !val('n_u').trim() || !val('n_s')) throw new Error('請填日期、單位與班別');
        await call('signup.quota.save', { items: [{ date: val('n_d'), unit: val('n_u').trim(), shift: val('n_s'), quota: Number(val('n_q')) }] });
        toast('已新增名額');
        var m = val('n_d').slice(0, 7); if (m !== month) viewSignup.month = m;
        viewSignup(main);
      }, this);
    };
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
  /* ---------- 二階段認證 ---------- */
  function askMfa(ticket) {
    return new Promise(function (resolve, reject) {
      var d = openDialog('<h3>二階段認證</h3><p class="muted">請輸入手機驗證器 App 上顯示的 6 位數驗證碼。手機不在身邊時,可以改輸入一組備用碼(每組只能用一次)。</p>' +
        '<form id="mfaf" class="form" style="grid-template-columns:1fr">' + field('mfac', '驗證碼或備用碼', '', 'text', 'autocomplete="one-time-code" autocapitalize="characters" required') +
        '<div class="row"><button class="primary" type="submit">驗證</button><button type="button" data-close>取消</button></div></form>');
      var done = false;
      setTimeout(function () { var i = document.getElementById('mfac'); if (i) i.focus(); }, 50);
      $('#mfaf').onsubmit = function (e) {
        e.preventDefault();
        var b = $('#mfaf .primary'); b.disabled = true;
        API.loginMfa(ticket, val('mfac').trim()).then(function (x) { done = true; d.close(); resolve(x); }, function (err) { b.disabled = false; toast(err.message, true); });
      };
      d.addEventListener('close', function () { if (!done) reject(new Error('已取消登入')); }, { once: true });
    });
  }

  async function renderMfa() {
    var box = document.getElementById('mfabox');
    if (!box) return;
    var st;
    try { st = await call('mfa.status'); } catch (e) { box.innerHTML = '<h3>二階段認證</h3><p class="muted">' + esc(e.message) + '</p>'; return; }
    box.innerHTML = '<h3>二階段認證' + (st.enabled ? ' <span class="badge ok">已啟用</span>' : ' <span class="badge">未啟用</span>') + '</h3>' +
      '<p class="hint">登入時除了密碼,還要輸入手機驗證器 App(Google Authenticator、Microsoft Authenticator 等)上每 30 秒更換一次的 6 位數驗證碼,密碼被別人知道也無法登入。</p>' +
      (st.enabled ? '<p>剩餘備用碼:<b>' + st.backupLeft + '</b> 組。</p><button class="danger" id="mfaoff">停用二階段認證</button>' : '<button class="primary" id="mfaon">啟用二階段認證</button>');
    var on = document.getElementById('mfaon'), off = document.getElementById('mfaoff');
    if (on) on.onclick = function () {
      guard(async function () {
        var r = await call('mfa.setup');
        var gen = await loadLib('vendor/qrcode.js', 'qrcode');
        var q = gen(0, 'M'); q.addData(r.uri); q.make();
        var d = openDialog('<h3>啟用二階段認證</h3><ol class="hint" style="padding-left:20px"><li>手機安裝驗證器 App(Google Authenticator 或 Microsoft Authenticator)。</li><li>在 App 按「新增」→「掃描 QR code」,掃下面的圖。</li><li>輸入 App 顯示的 6 位數驗證碼,按「確認啟用」。</li></ol>' +
          '<div style="max-width:240px;margin:8px auto">' + q.createSvgTag({ cellSize: 4, margin: 2, scalable: true }) + '</div>' +
          '<p class="hint">掃不了 QR 的話,在 App 選「輸入設定金鑰」,帳號隨便填,金鑰:<br><b style="letter-spacing:1px;word-break:break-all">' + esc(r.secret) + '</b></p>' +
          '<form id="mfae" class="form" style="grid-template-columns:1fr">' + field('mfacode', '6 位數驗證碼', '', 'text', 'inputmode="numeric" maxlength="6" autocomplete="off" required') +
          '<div class="row"><button class="primary" type="submit">確認啟用</button><button type="button" data-close>取消</button></div></form>');
        $('#mfae').onsubmit = function (e) {
          e.preventDefault();
          guard(async function () {
            var x = await call('mfa.enable', { code: val('mfacode').trim() });
            showBackupCodes(x.codes);
            renderMfa();
          }, $('#mfae .primary'));
        };
      }, on);
    };
    if (off) off.onclick = function () {
      var c = prompt('要停用二階段認證,請輸入目前的 6 位數驗證碼(或一組備用碼):');
      if (!c) return;
      guard(async function () { await call('mfa.disable', { code: c.trim() }); toast('已停用二階段認證'); renderMfa(); }, off);
    };
  }

  function showBackupCodes(codes) {
    var text = codes.join('\n');
    var d = openDialog('<h3>備用碼(只顯示這一次)</h3><p class="hint"><b>請現在抄下或列印,收在安全的地方。</b>手機遺失或無法使用驗證器時,每組備用碼可以代替驗證碼登入一次。關掉這個視窗後就看不到了。</p>' +
      '<pre id="bkc" style="font-size:20px;line-height:1.8;background:var(--bg,#f4f6fa);padding:12px;border-radius:8px;text-align:center">' + codes.map(esc).join('\n') + '</pre>' +
      '<div class="row"><button id="bkcopy">複製</button><button id="bkprint">列印</button><button class="primary" id="bkdone">我已經保存好了</button></div>');
    $('#bkcopy').onclick = function () { (navigator.clipboard ? navigator.clipboard.writeText(text) : Promise.reject()).then(function () { toast('已複製'); }, function () { toast('無法自動複製,請手動抄下', true); }); };
    $('#bkprint').onclick = function () {
      var w = window.open('', '_blank');
      if (!w) return toast('瀏覽器擋住了列印視窗,請手動抄下', true);
      w.document.write('<pre style="font-size:22px;line-height:2">東慧人事管理系統 備用碼\n\n' + codes.map(esc).join('\n') + '</pre>'); w.document.close(); w.print();
    };
    $('#bkdone').onclick = function () { d.close(); toast('二階段認證已啟用'); };
  }

  async function viewPassword(main) {
    var locked = user.role === 'employee'; // 員工不能自行改密碼,由人資重設
    main.innerHTML = '<h2>' + (locked ? '帳號安全' : '修改密碼') + '</h2>' + (locked
      ? '<div class="card"><p class="hint" style="margin:0">員工帳號的密碼由人資統一管理,無法自行修改。如需變更或忘記密碼,請聯絡人資重設。</p></div>'
      : '<div class="card" style="max-width:420px"><form id="pf" class="form" style="grid-template-columns:1fr">' +
      field('p0', '目前密碼', '', 'password', 'autocomplete="current-password" required') + field('p1', '新密碼(至少 8 碼)', '', 'password', 'autocomplete="new-password" minlength="8" required') +
      field('p2', '再輸入一次新密碼', '', 'password', 'autocomplete="new-password" required') + '<button class="primary" type="submit">更新密碼</button></form></div>') +
      '<div class="card"><h3>我的近期登入紀錄</h3><p class="hint">若看到不是您本人的登入(陌生 IP 或時間),請' + (locked ? '' : '立即修改密碼並') + '通知人資。</p><div id="mylog" class="muted">載入中…</div></div>';
    if (user.role === 'admin') { $('#mylog').parentNode.insertAdjacentHTML('beforebegin', '<div class="card" id="mfabox">載入中…</div>'); renderMfa(); }
    if (!locked) $('#pf').onsubmit = function (e) {
      e.preventDefault();
      if (val('p1') !== val('p2')) return toast('兩次輸入的新密碼不一致', true);
      guard(async function () { await call('changePassword', { oldPassword: val('p0'), newPassword: val('p1') }); toast('密碼已更新'); $('#pf').reset(); }, $('#pf button'));
    };
    try { $('#mylog').outerHTML = '<div id="mylog">' + loginRows(await call('login.my')) + '</div>'; } catch (e) { $('#mylog').textContent = e.message; }
  }

  /* ---------- 啟動 ---------- */
  document.title = cfg.COMPANY_NAME || '東慧公司岡山區人事管理系統';
  captureQr();
  if (API.hasToken()) {
    call('me').then(function (u) { user = u; renderShell(); }).catch(function () { API.logout(); renderLogin(); });
  } else renderLogin();
})();
