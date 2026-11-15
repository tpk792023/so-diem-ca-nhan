/* Sổ điểm cá nhân – nút dấu trang chạy trên vnEdu.
   KHÔNG BAO GIỜ tự lưu điểm: chỉ đọc sổ, tải file xuất, và điền vào ô (tô màu) để thầy tự kiểm tra rồi bấm Lưu. */
(function () {
  'use strict';
  var VER = '2026-11-15a';
  if (window.__SDCN && window.__SDCN.ver === VER) { window.__SDCN.show(); return; }
  if (window.__SDCN) window.__SDCN.destroy();

  var me = document.currentScript;
  var APP = (me && me.src ? me.src.replace(/vnedu\.js.*$/, '') : 'https://tpk792023.github.io/so-diem-ca-nhan/');
  var APP_ORIGIN = new URL(APP).origin;
  var JQ = window.jQuery || window.$;

  /* ---------- tiện ích ---------- */
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function norm(s) { return String(s || '').normalize('NFC').toLowerCase().replace(/\s+/g, ' ').trim(); }
  function has(v) { return v !== '' && v != null; }
  function num(v) {
    if (!has(v)) return '';
    var s = String(v).trim(); if (!s) return '';
    if (/^c\s*đ$/i.test(s)) return 'CĐ'; if (/^đ$/i.test(s)) return 'Đ';
    var n = parseFloat(s.replace(',', '.')); return isNaN(n) ? s : Math.round(n * 100) / 100;
  }
  function same(a, b) { return String(num(a)) === String(num(b)); }
  function colName(k) { return k.indexOf('tx') === 0 ? 'TX' + k.slice(2) : k.toUpperCase(); }
  function fullName(s) { return (s.ho + ' ' + s.ten).trim(); }
  function guessYear() {
    var c = [window.app_nam_hoc, window.nam_hoc, window.NAM_HOC, window.namHoc, window.g_nam_hoc];
    for (var i = 0; i < c.length; i++) { var n = parseInt(c[i], 10); if (n > 2000 && n < 2100) return n; }
    var inp = document.querySelector('input[name=nam_hoc],input[name=app_nam_hoc]');
    if (inp && parseInt(inp.value, 10) > 2000) return parseInt(inp.value, 10);
    var m = (document.body.innerText || '').match(/Năm học[:\s]*(\d{4})\s*-\s*\d{4}/i);
    if (m) return +m[1];
    var d = new Date(); return d.getMonth() >= 7 ? d.getFullYear() : d.getFullYear() - 1;
  }
  function guessHK() { var m = new Date().getMonth(); return (m >= 7 || m === 0) ? 1 : 2; }

  function ajax(opts) {
    return new Promise(function (res, rej) {
      if (!JQ || !JQ.ajax) { rej(new Error('Không thấy jQuery của vnEdu – hãy bấm nút khi đang ở trang vnEdu đã đăng nhập.')); return; }
      JQ.ajax(Object.assign({ cache: false, dataType: 'text' }, opts))
        .done(function (d) { res(d); })
        .fail(function (x, s, e) { rej(new Error((s || 'lỗi') + ' ' + (e || x.status || ''))); });
    });
  }
  function parseLoose(t) {
    if (typeof t !== 'string') return t;
    try { return JSON.parse(t); } catch (e) { }
    try { return JSON.parse(t.replace(/([{,]\s*)([A-Za-z_]\w*)\s*:/g, '$1"$2":').replace(/'/g, '"')); } catch (e) { }
    return null;
  }
  function findArr(o, d) {
    d = d || 0; if (!o || d > 4) return null;
    if (Array.isArray(o)) return o;
    var ks = ['data', 'rows', 'items', 'records', 'result', 'list', 'root'];
    for (var i = 0; i < ks.length; i++) if (Array.isArray(o[ks[i]])) return o[ks[i]];
    for (var k in o) if (o[k] && typeof o[k] === 'object') { var r = findArr(o[k], d + 1); if (r) return r; }
    return null;
  }
  function items(resp, idKeys) {
    var a = findArr(parseLoose(resp)) || [];
    return a.map(function (x) {
      if (Array.isArray(x)) return { id: x[0], name: String(x[1] || x[0]) };
      var id = null, name = null, k;
      for (var i = 0; i < idKeys.length; i++) if (x[idKeys[i]] != null) { id = x[idKeys[i]]; break; }
      if (id == null) id = x.id != null ? x.id : x.value;
      for (k in x) if (/^ten|name$|^text$/i.test(k) && typeof x[k] === 'string' && x[k].trim()) { name = x[k]; break; }
      return { id: id, name: name || String(id), raw: x };
    }).filter(function (x) { return x.id != null && x.id !== ''; });
  }
  async function callList(call, data, idKeys) {
    var url = '?call=' + call, r;
    try { r = items(await ajax({ url: url, type: 'GET', data: data }), idKeys); if (r.length) return r; } catch (e) { }
    return items(await ajax({ url: url, type: 'POST', data: data }), idKeys);
  }
  function getBuf(url) {
    return new Promise(function (res, rej) {
      var x = new XMLHttpRequest(); x.open('GET', url, true); x.responseType = 'arraybuffer';
      x.onload = function () { x.status === 200 ? res(x.response) : rej(new Error('HTTP ' + x.status)); };
      x.onerror = function () { rej(new Error('mạng')); }; x.send();
    });
  }
  function looksExcel(buf) {
    var u = new Uint8Array(buf.slice(0, 8));
    if (u[0] === 0xD0 && u[1] === 0xCF) return true;
    if (u[0] === 0x50 && u[1] === 0x4B) return true;
    var t = new TextDecoder().decode(buf.slice(0, 600));
    return /<\?xml|<Workbook|<table/i.test(t) && !/CSRF/i.test(t);
  }

  /* ---------- giao diện ---------- */
  var css = '#sdcn{position:fixed;top:12px;right:12px;width:min(480px,96vw);max-height:92vh;display:flex;flex-direction:column;background:#fff;color:#1d2623;border:2px solid #1f6f5c;border-radius:14px;box-shadow:0 12px 40px rgba(0,0,0,.3);z-index:2147483000;font:14px/1.4 Arial,Helvetica,sans-serif}'
    + '#sdcn *{box-sizing:border-box;font-family:inherit}'
    + '#sdcn .hd{display:flex;align-items:center;gap:6px;background:#1f6f5c;color:#fff;padding:8px 10px;border-radius:11px 11px 0 0;cursor:move;user-select:none}'
    + '#sdcn .hd b{flex:1;font-size:15px}'
    + '#sdcn .hd button{background:rgba(255,255,255,.2);border:0;color:#fff;border-radius:6px;width:28px;height:26px;cursor:pointer;font-size:14px}'
    + '#sdcn .bd{overflow:auto;padding:10px 12px}'
    + '#sdcn .tb{display:flex;gap:6px;margin-bottom:10px}'
    + '#sdcn .tb button{flex:1;padding:8px 6px;border:1px solid #cfd8d4;background:#f3f5f4;border-radius:8px;cursor:pointer;font-weight:bold;color:#1d2623}'
    + '#sdcn .tb button.on{background:#1f6f5c;color:#fff;border-color:#1f6f5c}'
    + '#sdcn .btn{padding:7px 12px;border:1px solid #cfd8d4;background:#fff;border-radius:8px;cursor:pointer;margin:3px 4px 3px 0;color:#1d2623;font-size:14px}'
    + '#sdcn .btn.pri{background:#1f6f5c;color:#fff;border-color:#1f6f5c}'
    + '#sdcn .btn.acc{background:#d9822b;color:#fff;border-color:#d9822b}'
    + '#sdcn .btn:disabled{opacity:.5;cursor:default}'
    + '#sdcn .mut{color:#66736e;font-size:13px}'
    + '#sdcn .box{border:1px solid #dde3e0;border-radius:10px;padding:8px 10px;margin:8px 0;background:#fafcfb}'
    + '#sdcn .warn{background:#fff1df;border-color:#f0c48f}'
    + '#sdcn .ok{background:#e3f1ec;border-color:#9fd2c1}'
    + '#sdcn .err{background:#fdeceb;border-color:#f1aaa3}'
    + '#sdcn table{border-collapse:collapse;width:100%;font-size:13px}'
    + '#sdcn th,#sdcn td{border-bottom:1px solid #e6ebe9;padding:4px 5px;text-align:left;vertical-align:top}'
    + '#sdcn th{background:#f3f5f4;position:sticky;top:0}'
    + '#sdcn td.n{text-align:center;white-space:nowrap}'
    + '#sdcn tr.cf td{background:#fff1df}'
    + '#sdcn input[type=number],#sdcn input[type=text],#sdcn select,#sdcn textarea{border:1px solid #cfd8d4;border-radius:6px;padding:4px 6px;font-size:14px;background:#fff;color:#1d2623}'
    + '#sdcn textarea{width:100%;height:60px;font:11px monospace}'
    + '#sdcn label{margin-right:8px;white-space:nowrap}'
    + '#sdcn .scroll{max-height:42vh;overflow:auto;border:1px solid #e6ebe9;border-radius:8px}'
    + '#sdcn.min .bd{display:none}';
  var st = document.createElement('style'); st.id = 'sdcn-css'; st.textContent = css; document.head.appendChild(st);
  var P = document.createElement('div'); P.id = 'sdcn';
  P.innerHTML = '<div class="hd"><b>📒 Sổ điểm cá nhân</b><button data-a="min" title="Thu nhỏ">–</button><button data-a="close" title="Đóng">✕</button></div>'
    + '<div class="bd"><div class="tb"><button data-t="fill" class="on">✍ Điền điểm vào sổ đang mở</button><button data-t="load">📥 Lấy lớp & điểm</button></div><div id="sdcn-body"></div>'
    + '<p class="mut" style="margin:10px 0 0">Nút này <b>không tự lưu</b> điểm. Thầy kiểm tra ô tô vàng rồi tự bấm <b>Lưu</b> của vnEdu. <a href="#" data-a="diag" style="color:#66736e">Thông tin kỹ thuật</a> · bản ' + VER + '</p></div>';
  document.body.appendChild(P);
  var B = P.querySelector('#sdcn-body');
  function q(s) { return P.querySelector(s); }

  // kéo khung
  (function () {
    var hd = P.querySelector('.hd'), sx, sy, ox, oy, on = false;
    hd.addEventListener('mousedown', function (e) { if (e.target.tagName === 'BUTTON') return; on = true; sx = e.clientX; sy = e.clientY; var r = P.getBoundingClientRect(); ox = r.left; oy = r.top; e.preventDefault(); });
    document.addEventListener('mousemove', function (e) { if (!on) return; P.style.left = (ox + e.clientX - sx) + 'px'; P.style.top = (oy + e.clientY - sy) + 'px'; P.style.right = 'auto'; });
    document.addEventListener('mouseup', function () { on = false; });
  })();

  var tab = 'fill';
  P.addEventListener('click', function (e) {
    var b = e.target.closest('[data-a],[data-t]'); if (!b) return;
    if (b.dataset.t) { tab = b.dataset.t; P.querySelectorAll('.tb button').forEach(function (x) { x.classList.toggle('on', x === b); }); tab === 'fill' ? renderFill() : renderLoad(); return; }
    e.preventDefault();
    var a = b.dataset.a;
    if (a === 'min') P.classList.toggle('min');
    else if (a === 'close') { P.style.display = 'none'; }
    else if (a === 'diag') diag();
  });

  /* ---------- liên lạc với trang Sổ điểm cá nhân ---------- */
  var appWin = null, waiters = [];
  window.addEventListener('message', function (e) {
    if (e.origin !== APP_ORIGIN || !e.data || typeof e.data !== 'object') return;
    waiters.slice().forEach(function (w) { if (w.type === e.data.type) { waiters.splice(waiters.indexOf(w), 1); w.res(e.data); } });
  });
  function waitMsg(type, ms) {
    return new Promise(function (res, rej) {
      var w = { type: type, res: res }; waiters.push(w);
      setTimeout(function () { var i = waiters.indexOf(w); if (i >= 0) { waiters.splice(i, 1); rej(new Error('Không nhận được trả lời từ trang Sổ điểm cá nhân (trình duyệt có chặn cửa sổ bật lên?)')); } }, ms || 20000);
    });
  }
  function openApp() {
    appWin = window.open(APP + '?k=' + Date.now() + '#vnedu', 'sdcn_app');
    if (!appWin) throw new Error('Trình duyệt chặn mở cửa sổ. Cho phép cửa sổ bật lên (popup) cho trang vnEdu rồi bấm lại.');
    return waitMsg('sdcn-ready', 25000);
  }

  /* ---------- ĐỌC SỔ ĐIỂM ĐANG MỞ ---------- */
  function visible(el) { return !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length); }
  function readGrid() {
    var all = Array.prototype.slice.call(document.querySelectorAll('span.input_diem'));
    var spans = all.filter(visible); if (!spans.length) spans = all;
    if (!spans.length) return null;
    // b của từng loại cột: lấy từ tiêu đề th[cn] nếu có, không thì b=2 (TX), 3 (GK), 4 (CK)
    var bTx = '2', bGk = '3', bCk = '4';
    document.querySelectorAll('th[cn],td[cn]').forEach(function (th) {
      var cn = norm(th.getAttribute('cn')), b = th.getAttribute('b'); if (!b || !visible(th)) return;
      if (/tx/.test(cn)) bTx = b; else if (/gk/.test(cn)) bGk = b; else if (/ck/.test(cn)) bCk = b;
    });
    var rowsMap = new Map();
    spans.forEach(function (sp) {
      var b = sp.getAttribute('b'), c = sp.getAttribute('c'), k = null;
      if (b === bTx) k = 'tx' + (c || '1'); else if (b === bGk) k = 'gk'; else if (b === bCk) k = 'ck';
      if (!k) return;
      var tr = sp.closest('tr'); if (!tr) return;
      if (!rowsMap.has(tr)) rowsMap.set(tr, { tr: tr, cells: {} });
      rowsMap.get(tr).cells[k] = { b: b, c: c };
    });
    var rows = Array.from(rowsMap.values());
    rows.forEach(function (r) {
      var t = r.tr.innerText || r.tr.textContent || '';
      var m = t.match(/(^|\D)(\d{9,12})(?!\d)/); r.ma = m ? m[2] : '';
      var best = '';
      Array.prototype.forEach.call(r.tr.children, function (td) {
        if (td.querySelector('span.input_diem')) return;
        var s = (td.innerText || td.textContent || '').replace(/\s+/g, ' ').trim();
        if (/\d/.test(s) || s.length < 3) return;
        if (/[a-zà-ỹđ]{1,}\s+[a-zà-ỹđ]/i.test(s) && s.length > best.length) best = s;
      });
      r.name = best;
    });
    // lưới tách hai bảng (cột tên cố định): ghép theo thứ tự
    if (rows.length && rows.every(function (r) { return !r.name && !r.ma; })) {
      var tbody = rows[0].tr.parentNode, cand = null;
      document.querySelectorAll('tbody').forEach(function (tb) {
        if (tb === tbody || cand) return;
        if (tb.children.length === tbody.children.length && /[a-zà-ỹ]+\s+[a-zà-ỹ]+/i.test(tb.innerText || '')) cand = tb;
      });
      if (cand) rows.forEach(function (r) {
        var i = Array.prototype.indexOf.call(tbody.children, r.tr), tr2 = cand.children[i]; if (!tr2) return;
        var t = tr2.innerText || ''; var m = t.match(/(^|\D)(\d{9,12})(?!\d)/); r.ma = m ? m[2] : '';
        var best = ''; Array.prototype.forEach.call(tr2.children, function (td) { var s = (td.innerText || '').replace(/\s+/g, ' ').trim(); if (!/\d/.test(s) && s.length > best.length) best = s; }); r.name = best;
      });
    }
    var txt = document.body.innerText || '';
    var role = (txt.match(/Quyền hạn:\s*([^\n]+)/i) || [])[1] || '';
    return { rows: rows, role: role.trim(), dec: spans.some(function (s) { return /\d,\d/.test(s.textContent); }) ? ',' : '.' };
  }
  function spanOf(row, k) {
    var c = row.cells[k]; if (!c) return null;
    return row.tr.querySelector('span.input_diem[b="' + c.b + '"]' + (c.c != null ? '[c="' + c.c + '"]' : ''));
  }
  function cellVal(row, k) { var s = spanOf(row, k); return s ? num((s.textContent || '').trim()) : ''; }

  function matchClass(cls, grid) {
    var byMa = new Map(), byName = new Map();
    cls.st.forEach(function (s, i) { if (s.ma) byMa.set(String(s.ma), s); var n = norm(fullName(s)); if (!byName.has(n)) byName.set(n, []); byName.get(n).push(s); });
    var used = new Set(), pairs = [], lost = [];
    grid.rows.forEach(function (r) {
      var s = r.ma && byMa.get(r.ma);
      if (!s && r.name) { var l = (byName.get(norm(r.name)) || []).filter(function (x) { return !used.has(x); }); s = l[0]; }
      if (s && !used.has(s)) { used.add(s); pairs.push({ row: r, s: s }); } else lost.push(r);
    });
    var miss = cls.st.filter(function (s) { return !used.has(s); });
    return { pairs: pairs, lostRows: lost, miss: miss, score: pairs.length };
  }

  var DATA = null, PLAN = null, FILLED = [];
  function renderFill(msg) {
    var h = '';
    if (msg) h += '<div class="box ' + (msg.cls || '') + '">' + msg.html + '</div>';
    if (!DATA) {
      h += '<p><b>Bước 1.</b> Mở <b>Sổ điểm</b> của lớp – môn cần đưa điểm trên vnEdu.</p>'
        + '<p><b>Bước 2.</b> Lấy điểm đã nhập:</p>'
        + '<button class="btn pri" id="sd-fromapp">Lấy từ trang Sổ điểm cá nhân trên máy này</button>'
        + '<div class="box"><div class="mut">Hoặc chọn file điểm gửi từ điện thoại (so-diem-ca-nhan-….json):</div><input type="file" id="sd-file" accept=".json,.txt"></div>'
        + '<details><summary class="mut">Hoặc dán mã SDCN1.</summary><textarea id="sd-code"></textarea><button class="btn" id="sd-usecode">Dùng mã</button></details>';
    } else {
      h += '<div class="box ok">Đã có điểm của <b>' + DATA.classes.length + '</b> lớp. <a href="#" id="sd-reset">Lấy lại</a></div>';
      h += '<button class="btn pri" id="sd-scan">🔍 So sánh với sổ điểm đang mở</button>';
      if (PLAN) h += planHtml();
    }
    B.innerHTML = h;
    bindFill();
  }
  function bindFill() {
    var b;
    if (b = q('#sd-fromapp')) b.onclick = async function () {
      b.disabled = true; b.textContent = 'Đang mở trang Sổ điểm cá nhân…';
      try { await openApp(); appWin.postMessage({ type: 'sdcn-want' }, APP_ORIGIN); var d = await waitMsg('sdcn-data', 15000); setData(d.data); try { window.focus(); } catch (e) { } }
      catch (e) { renderFill({ cls: 'err', html: esc(e.message) }); }
    };
    if (b = q('#sd-file')) b.onchange = async function (e) {
      var f = e.target.files[0]; if (!f) return;
      try { var t = await f.text(); setData(t.trim().indexOf('SDCN1.') === 0 ? await decodeCode(t) : JSON.parse(t)); }
      catch (err) { renderFill({ cls: 'err', html: 'Không đọc được file: ' + esc(err.message) }); }
    };
    if (b = q('#sd-usecode')) b.onclick = async function () {
      try { setData(await decodeCode(q('#sd-code').value)); } catch (err) { renderFill({ cls: 'err', html: 'Mã không đúng: ' + esc(err.message) }); }
    };
    if (b = q('#sd-reset')) b.onclick = function (e) { e.preventDefault(); DATA = null; PLAN = null; renderFill(); };
    if (b = q('#sd-scan')) b.onclick = function () { scan(); };
    if (b = q('#sd-cls')) b.onchange = function () { scan(b.value); };
    if (b = q('#sd-go')) b.onclick = function () { doFill(); };
    if (b = q('#sd-unmark')) b.onclick = function () { unmark(); };
    if (b = q('#sd-all')) b.onchange = function () { P.querySelectorAll('.sd-ck').forEach(function (x) { x.checked = b.checked; }); updGo(); };
    P.querySelectorAll('.sd-ck').forEach(function (x) { x.onchange = updGo; });
  }
  function setData(d) {
    if (!d || d.app !== 'so-diem-ca-nhan' || !Array.isArray(d.classes)) throw new Error('Không phải dữ liệu của Sổ điểm cá nhân');
    DATA = d; PLAN = null; scan();
  }
  async function decodeCode(code) {
    var b = code.trim().replace(/^SDCN1\./, '').replace(/\s+/g, '').replace(/-/g, '+').replace(/_/g, '/');
    var bin = atob(b), u = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    var ds = new Blob([u]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return JSON.parse(await new Response(ds).text());
  }
  function scan(forceId) {
    var g = readGrid();
    if (!g || !g.rows.length) { PLAN = null; renderFill({ cls: 'warn', html: 'Chưa thấy ô điểm nào trên trang. Hãy mở <b>Sổ điểm</b> của lớp – môn (màn hình có các ô TX, GK, CK) rồi bấm <b>So sánh</b> lại.' }); return; }
    var ranked = DATA.classes.map(function (c) { return { c: c, m: matchClass(c, g) }; }).sort(function (a, b) { return b.m.score - a.m.score; });
    var pick = forceId ? ranked.filter(function (x) { return x.c.id === forceId; })[0] : ranked[0];
    if (!pick || !pick.m.score) { PLAN = { g: g, ranked: ranked, pick: null, items: [] }; renderFill({ cls: 'warn', html: 'Không lớp nào trong dữ liệu khớp tên học sinh với sổ đang mở (' + g.rows.length + ' dòng).' }); return; }
    var c = pick.c, items = [];
    pick.m.pairs.forEach(function (p) {
      Object.keys(p.row.cells).forEach(function (k) {
        var want = p.s.d[k]; if (!has(want)) return;
        var cur = cellVal(p.row, k); if (same(cur, want)) return;
        var orig = p.s.o[k];
        var kind = !has(cur) ? 'mới' : (same(cur, orig) ? 'sửa' : 'khác');
        items.push({ row: p.row, s: p.s, k: k, cur: cur, want: want, kind: kind });
      });
    });
    PLAN = { g: g, ranked: ranked, pick: pick, items: items };
    renderFill();
  }
  function fmt(v) { return has(v) ? String(v).replace('.', ',') : '∅'; }
  function planHtml() {
    var p = PLAN, h = '';
    if (p.ranked.length) {
      h += '<div class="box">Lớp trong dữ liệu: <select id="sd-cls">' + p.ranked.map(function (x) {
        return '<option value="' + esc(x.c.id) + '"' + (p.pick && x.c.id === p.pick.c.id ? ' selected' : '') + '>' + esc(x.c.lop + ' – ' + x.c.mon + ' – HK' + x.c.hk) + ' (' + x.m.score + ' em khớp)</option>';
      }).join('') + '</select>';
      if (p.pick) {
        h += '<div class="mut">Sổ đang mở: ' + p.g.rows.length + ' dòng; khớp ' + p.pick.m.score + ' em.' + (p.g.role ? ' Quyền hạn: ' + esc(p.g.role) : '') + '</div>';
        h += '<div class="mut">⚠ Kiểm tra đúng <b>lớp, môn, học kỳ HK' + p.pick.c.hk + '</b> trên vnEdu trước khi điền.</div>';
        if (p.pick.m.miss.length) h += '<div class="mut">Không thấy trong sổ: ' + esc(p.pick.m.miss.slice(0, 6).map(fullName).join(', ')) + (p.pick.m.miss.length > 6 ? '…' : '') + '</div>';
      }
      h += '</div>';
    }
    if (!p.pick) return h;
    if (!p.items.length) return h + '<div class="box ok">✔ Sổ điểm đang mở đã có đủ điểm như Sổ điểm cá nhân. Không có gì để điền.</div>' + (FILLED.length ? '<button class="btn" id="sd-unmark">Bỏ tô màu</button>' : '');
    var nCf = p.items.filter(function (x) { return x.kind === 'khác'; }).length;
    h += '<p><b>' + p.items.length + '</b> ô sẽ điền' + (nCf ? ', trong đó <b>' + nCf + '</b> ô (tô cam) trên vnEdu đã có điểm khác lúc thầy lấy về – mặc định <b>không</b> điền, thầy tự chọn.' : '.') + '</p>';
    h += '<div class="scroll"><table><thead><tr><th><input type="checkbox" id="sd-all" title="Chọn tất cả"></th><th>Học sinh</th><th>Cột</th><th>Trong sổ</th><th>Điền</th></tr></thead><tbody>'
      + p.items.map(function (x, i) {
        return '<tr class="' + (x.kind === 'khác' ? 'cf' : '') + '"><td class="n"><input type="checkbox" class="sd-ck" data-i="' + i + '"' + (x.kind === 'khác' ? '' : ' checked') + '></td><td>' + esc(fullName(x.s)) + '</td><td class="n">' + colName(x.k) + '</td><td class="n">' + esc(fmt(x.cur)) + '</td><td class="n"><b>' + esc(fmt(x.want)) + '</b></td></tr>';
      }).join('') + '</tbody></table></div>';
    h += '<button class="btn acc" id="sd-go" style="margin-top:8px">✍ Điền vào sổ (chưa lưu)</button>' + (FILLED.length ? '<button class="btn" id="sd-unmark">Bỏ tô màu</button>' : '');
    return h;
  }
  function updGo() { var n = P.querySelectorAll('.sd-ck:checked').length, b = q('#sd-go'); if (b) { b.disabled = !n; b.textContent = '✍ Điền ' + n + ' ô vào sổ (chưa lưu)'; } }

  /* ---------- điền một ô: làm như người gõ ---------- */
  function fire(el, type, opt) {
    var E = /^(mouse|click|dblclick)/.test(type) ? MouseEvent : (/^key/.test(type) ? KeyboardEvent : (/^(focus|blur)/.test(type) ? FocusEvent : Event));
    el.dispatchEvent(new E(type, Object.assign({ bubbles: type !== 'focus' && type !== 'blur', cancelable: true, view: window }, opt || {})));
  }
  function editorNear(sp, td) {
    var a = document.activeElement;
    if (a && /^(INPUT|TEXTAREA)$/.test(a.tagName) && !P.contains(a) && a.type !== 'checkbox') return a;
    var c = (sp && sp.querySelector('input,textarea')) || (td && td.querySelector('input:not([type=checkbox]),textarea'));
    return c && visible(c) ? c : null;
  }
  function setNative(inp, v) {
    var proto = inp.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    var d = Object.getOwnPropertyDescriptor(proto, 'value'); d && d.set ? d.set.call(inp, v) : (inp.value = v);
  }
  async function fillOne(x, dec) {
    var sp = spanOf(x.row, x.k); if (!sp) return 'fail';
    var td = sp.closest('td') || sp.parentNode;
    var txt = typeof x.want === 'number' ? String(x.want).replace('.', dec) : String(x.want);
    ['mouseover', 'mousedown', 'mouseup', 'click'].forEach(function (t) { fire(sp, t); });
    await sleep(40);
    var inp = editorNear(sp, td);
    if (!inp) { fire(sp, 'dblclick'); await sleep(40); inp = editorNear(sp, td); }
    if (inp) {
      try { inp.focus(); } catch (e) { }
      setNative(inp, txt);
      fire(inp, 'input'); fire(inp, 'keyup', { key: txt.slice(-1) }); fire(inp, 'change');
      try { inp.blur(); } catch (e) { } fire(inp, 'blur'); fire(inp, 'focusout');
      await sleep(40);
      var sp2 = spanOf(x.row, x.k);
      if (sp2 && same((sp2.textContent || '').trim(), x.want)) { mark(sp2, x, 'gõ'); return 'ok'; }
      if (inp.isConnected && same(inp.value, x.want)) { mark(inp, x, 'gõ'); return 'ok'; }
    }
    // không thấy ô gõ: ghi thẳng chữ vào ô (cần thầy kiểm tra kỹ sau khi Lưu)
    sp = spanOf(x.row, x.k); if (!sp) return 'fail';
    sp.textContent = txt; fire(sp, 'change'); fire(sp, 'input');
    mark(sp, x, 'ghi thẳng'); return 'direct';
  }
  function mark(el, x, how) {
    var td = el.closest('td') || el;
    if (!td.dataset.sdcnBg) td.dataset.sdcnBg = td.style.background || ' ';
    td.style.background = x.kind === 'khác' ? '#ffd8a8' : '#fff59d';
    td.style.outline = '2px solid #d9822b';
    td.title = 'Sổ điểm cá nhân (' + how + '): ' + fmt(x.cur) + ' → ' + fmt(x.want) + ' – CHƯA LƯU';
    FILLED.push(td);
  }
  function unmark() {
    FILLED.forEach(function (td) { td.style.background = td.dataset.sdcnBg === ' ' ? '' : td.dataset.sdcnBg; td.style.outline = ''; td.removeAttribute('title'); delete td.dataset.sdcnBg; });
    FILLED = []; renderFill();
  }
  async function doFill() {
    var sel = Array.prototype.map.call(P.querySelectorAll('.sd-ck:checked'), function (c) { return PLAN.items[+c.dataset.i]; });
    if (!sel.length) return;
    var go = q('#sd-go'); go.disabled = true;
    var r = { ok: 0, direct: 0, fail: 0 }, bad = [];
    for (var i = 0; i < sel.length; i++) {
      go.textContent = 'Đang điền ' + (i + 1) + '/' + sel.length + '…';
      var k;
      try { k = await fillOne(sel[i], PLAN.g.dec); } catch (e) { k = 'fail'; }
      r[k]++; if (k === 'fail') bad.push(fullName(sel[i].s) + ' ' + colName(sel[i].k));
    }
    try { if (document.activeElement && !P.contains(document.activeElement)) document.activeElement.blur(); } catch (e) { }
    var h = '✔ Đã điền <b>' + (r.ok + r.direct) + '</b> ô và tô vàng. <b>Chưa lưu.</b> Thầy kiểm tra từng ô rồi bấm <b>Lưu</b> của vnEdu.';
    if (r.direct) h += '<br>⚠ ' + r.direct + ' ô được ghi thẳng chữ (không thấy ô gõ của vnEdu). Sau khi Lưu, hãy tải lại sổ để chắc vnEdu đã nhận; nếu không nhận, dùng cách <b>Nhập từ Excel</b> (file tải ở trang Sổ điểm cá nhân).';
    if (r.fail) h += '<br>✖ Không điền được: ' + esc(bad.join(', '));
    var keep = FILLED.slice();
    scan(PLAN.pick.c.id); FILLED = keep;
    renderFill({ cls: r.fail || r.direct ? 'warn' : 'ok', html: h });
  }

  /* ---------- LẤY LỚP & ĐIỂM CÁC MÔN TÔI DẠY ---------- */
  var FOUND = [];
  function renderLoad(msg) {
    var h = (msg ? '<div class="box ' + (msg.cls || '') + '">' + msg.html + '</div>' : '')
      + '<p class="mut">Tìm các sổ điểm thầy là <b>giáo viên bộ môn</b>, tải file xuất của vnEdu và gửi sang trang Sổ điểm cá nhân (danh sách học sinh + điểm hiện có).</p>'
      + '<div class="box">Năm học <input type="number" id="sd-nam" style="width:80px" value="' + guessYear() + '">–<span id="sd-nam2"></span>'
      + ' &nbsp; Học kỳ <select id="sd-hk"><option value="1">1</option><option value="2">2</option></select><br>'
      + 'Khối: ' + [6, 7, 8, 9].map(function (k) { return '<label><input type="checkbox" class="sd-khoi" value="' + k + '" checked> ' + k + '</label>'; }).join('')
      + '<br>Chỉ môn có chữ: <input type="text" id="sd-mon" placeholder="vd: khoa học (bỏ trống = mọi môn)" style="width:220px"></div>'
      + '<button class="btn pri" id="sd-find">🔍 Tìm sổ điểm của tôi</button><div id="sd-prog" class="mut"></div><div id="sd-res"></div>';
    B.innerHTML = h;
    q('#sd-hk').value = String(guessHK());
    var upd = function () { q('#sd-nam2').textContent = (+q('#sd-nam').value + 1); }; upd(); q('#sd-nam').oninput = upd;
    q('#sd-find').onclick = find;
    if (FOUND.length) showFound();
  }
  function prog(t) { var p = q('#sd-prog'); if (p) p.textContent = t; }
  async function pool(list, n, fn) { var i = 0; async function w() { while (i < list.length) { var k = i++; await fn(list[k], k); } } var ws = []; for (var j = 0; j < n; j++) ws.push(w()); await Promise.all(ws); }
  async function find() {
    var nam = +q('#sd-nam').value, hk = +q('#sd-hk').value, monF = norm(q('#sd-mon').value);
    var khois = Array.prototype.map.call(P.querySelectorAll('.sd-khoi:checked'), function (x) { return +x.value; });
    var btn = q('#sd-find'); btn.disabled = true; FOUND = []; q('#sd-res').innerHTML = '';
    try {
      var jobs = [];
      for (var i = 0; i < khois.length; i++) {
        prog('Đọc danh sách lớp khối ' + khois[i] + '…');
        var lops = await callList('edu.khoi_hoc.getLopByKhoix', { nam_hoc: nam, khoi: khois[i] }, ['lop_id', 'id']);
        for (var j = 0; j < lops.length; j++) {
          prog('Đọc môn học lớp ' + lops[j].name + '…');
          var mons = [];
          try { mons = await callList('edu.so_diem.getMon', { lop_id: lops[j].id, nam_hoc: nam, hoc_ky: hk }, ['mon_hoc_id', 'id']); } catch (e) { }
          mons.forEach(function (m) { if (!monF || norm(m.name).indexOf(monF) >= 0) jobs.push({ khoi: khois[i], lop: lops[j], mon: m }); });
        }
      }
      if (!jobs.length) throw new Error('Không đọc được lớp/môn nào. Kiểm tra năm học, khối; hoặc bấm “Thông tin kỹ thuật” và gửi cho người làm phần mềm.');
      var done = 0;
      await pool(jobs, 4, async function (jb) {
        try {
          var html = await ajax({
            url: '?load=edu.so_diem.nhap', type: 'POST',
            data: { winid: 'sdcn' + Math.floor(Math.random() * 1e6), nam_hoc: nam, iLopId: jb.lop.id, iKhoi: jb.khoi, iMonHocId: jb.mon.id, iHocKyId: hk, thang: 0, dot_id: 0 }
          });
          var t = String(html).replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ');
          var role = (t.match(/Quyền hạn:\s*(.{0,60})/i) || [])[1] || '';
          jb.role = role; jb.mine = /giáo viên bộ môn/i.test(role);
          jb.nHs = (String(html).match(/class="[^"]*input_diem/g) || []).length;
        } catch (e) { jb.err = e.message; }
        done++; prog('Kiểm tra sổ điểm ' + done + '/' + jobs.length + '…');
      });
      jobs.forEach(function (j) { j.nam = nam; j.hk = hk; });
      FOUND = jobs.filter(function (j) { return j.mine; });
      var others = jobs.filter(function (j) { return !j.mine; });
      if (!FOUND.length) FOUND = others.map(function (j) { j.unsure = true; return j; });
      else FOUND = FOUND.concat(others.map(function (j) { j.unsure = true; return j; }));
      prog('Đã kiểm tra ' + jobs.length + ' sổ điểm.');
      showFound();
    } catch (e) { prog(''); q('#sd-res').innerHTML = '<div class="box err">' + esc(e.message) + '</div>'; }
    btn.disabled = false;
  }
  function showFound() {
    var mine = FOUND.filter(function (j) { return !j.unsure; }).length;
    var h = '<p><b>' + mine + '</b> sổ điểm thầy là giáo viên bộ môn' + (mine ? ' (đã chọn sẵn).' : '. Thầy tự chọn sổ cần lấy.') + '</p><div class="scroll"><table><thead><tr><th></th><th>Lớp</th><th>Môn</th><th>Quyền hạn</th></tr></thead><tbody>'
      + FOUND.map(function (j, i) { return '<tr><td class="n"><input type="checkbox" class="sd-pick" data-i="' + i + '"' + (j.unsure ? '' : ' checked') + '></td><td>' + esc(j.lop.name) + '</td><td>' + esc(j.mon.name) + '</td><td class="mut">' + esc(j.err ? 'lỗi: ' + j.err : (j.role || '?')) + '</td></tr>'; }).join('')
      + '</tbody></table></div><button class="btn acc" id="sd-send">📤 Gửi sang Sổ điểm cá nhân</button><button class="btn" id="sd-dl">⬇ Tải file .xls về máy</button>';
    q('#sd-res').innerHTML = h;
    q('#sd-send').onclick = function () { getFiles(true); };
    q('#sd-dl').onclick = function () { getFiles(false); };
  }
  async function getFiles(send) {
    var picks = Array.prototype.map.call(P.querySelectorAll('.sd-pick:checked'), function (c) { return FOUND[+c.dataset.i]; });
    if (!picks.length) { prog('Chưa chọn sổ nào.'); return; }
    var files = [], bad = [];
    var tok = window.myToken || window.my_token || '';
    for (var i = 0; i < picks.length; i++) {
      var j = picks[i];
      prog('Tải file ' + (i + 1) + '/' + picks.length + ': ' + j.lop.name + ' – ' + j.mon.name + '…');
      var url = '?call=edu.so_diem.export&action=xml&nam_hoc=' + j.nam + '&iLopId=' + encodeURIComponent(j.lop.id) + '&iKhoi=' + j.khoi
        + '&iMonHocId=' + encodeURIComponent(j.mon.id) + '&iHocKyId=' + j.hk + '&thang=0&dot_id=0&my_token=' + encodeURIComponent(tok) + '&app_nam_hoc=' + j.nam;
      try {
        var buf = await getBuf(url);
        if (!looksExcel(buf)) throw new Error('vnEdu không trả file Excel');
        files.push({ name: 'so_diem_' + j.lop.name + '_' + j.mon.name + '.xls', buf: buf, meta: { lopId: j.lop.id, khoi: j.khoi, monId: j.mon.id } });
      } catch (e) { bad.push(j.lop.name + ' – ' + j.mon.name + ' (' + e.message + ')'); }
    }
    if (!files.length) { prog(''); q('#sd-res').insertAdjacentHTML('afterbegin', '<div class="box err">Không tải được file nào: ' + esc(bad.join('; ')) + '</div>'); return; }
    if (!send) {
      files.forEach(function (f, i) { setTimeout(function () { var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([f.buf], { type: 'application/vnd.ms-excel' })); a.download = f.name; document.body.appendChild(a); a.click(); a.remove(); }, i * 400); });
      prog('Đã tải ' + files.length + ' file. Nạp các file này ở mục "Lớp & file" của trang Sổ điểm cá nhân.' + (bad.length ? ' Lỗi: ' + bad.join('; ') : ''));
      return;
    }
    try {
      prog('Đang mở trang Sổ điểm cá nhân…');
      await openApp();
      appWin.postMessage({ type: 'sdcn-xls', files: files }, APP_ORIGIN);
      var r = await waitMsg('sdcn-got', 30000);
      prog('');
      q('#sd-res').insertAdjacentHTML('afterbegin', '<div class="box ok">' + esc(r.msg) + (bad.length ? '<br>Lỗi: ' + esc(bad.join('; ')) : '') + '<br>Muốn nhập trên điện thoại: ở trang Sổ điểm cá nhân bấm <b>📤 Gửi file điểm</b>, mở file đó trên điện thoại.</div>');
    } catch (e) { prog(''); q('#sd-res').insertAdjacentHTML('afterbegin', '<div class="box err">' + esc(e.message) + '</div>'); }
  }

  /* ---------- thông tin kỹ thuật ---------- */
  function diag() {
    var sp = document.querySelector('span.input_diem'), tr = sp && sp.closest('tr');
    var ths = Array.prototype.slice.call(document.querySelectorAll('th[cn],td[cn]'), 0, 12).map(function (t) { return t.outerHTML.slice(0, 160); });
    var info = {
      ver: VER, host: location.host, path: location.pathname, jquery: JQ && JQ.fn && JQ.fn.jquery, ext: window.Ext && Ext.version,
      nam: [window.app_nam_hoc, window.nam_hoc, window.NAM_HOC].join('|'), token: !!window.myToken,
      spans: document.querySelectorAll('span.input_diem').length,
      span: sp ? sp.outerHTML.slice(0, 300) : null, th: ths,
      row: tr ? tr.outerHTML.replace(/\s+/g, ' ').slice(0, 1500) : null
    };
    var t = JSON.stringify(info, null, 1);
    B.innerHTML = '<p class="mut">Thông tin này giúp sửa nút khi vnEdu thay đổi (có thể chứa tên 1 học sinh). Chép và gửi cho người làm phần mềm nếu cần.</p><textarea style="height:200px">' + esc(t) + '</textarea><button class="btn" id="sd-back">Quay lại</button>';
    q('#sd-back').onclick = function () { tab === 'fill' ? renderFill() : renderLoad(); };
  }

  window.__SDCN = {
    ver: VER,
    show: function () { P.style.display = ''; P.classList.remove('min'); },
    destroy: function () { P.remove(); st.remove(); }
  };
  renderFill();
})();
