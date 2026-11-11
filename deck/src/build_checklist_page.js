// 手機可勾選版行李核對清單：讀 checklist_athlete.json → 單一 HTML（Artifact 頁面內容，不含 html/head/body）
const fs = require('fs');
const path = require('path');
const SRC = process.argv[2] || path.join(__dirname, '..', 'checklist', 'checklist_athlete.json');
const OUT = process.argv[3] || path.join(__dirname, '..', 'checklist', 'fxt022-packing-checklist.html');
const D = JSON.parse(fs.readFileSync(SRC, 'utf8'));
const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const hash = s => { let h = 0x811c9dc5; for (const ch of Buffer.from(s, 'utf8')) { h ^= ch; h = Math.imul(h, 0x01000193) >>> 0; } return h.toString(36); };
const seen = new Set();
const idOf = (sec, it) => { let id = `${sec.id}-${hash(it.t)}`; while (seen.has(id)) id += 'x'; seen.add(id); return id; };
const total = D.sections.reduce((n, s) => n + s.items.length, 0);
const row = (sec, it) => { const id = idOf(sec, it); return `<li><label class="row" for="c-${id}"><input type="checkbox" id="c-${id}" data-id="${id}"><span class="txt"><span class="t">${esc(it.t)}${it.q ? ` <span class="qty">${esc(it.q)}</span>` : ''}${it.s ? ' <span class="sug">建議</span>' : ''}</span>${it.n ? `<span class="n">${esc(it.n)}</span>` : ''}</span></label></li>`; };
const section = s => `<details class="sec${s.id === 'last' ? ' hot' : ''}" open data-sec="${esc(s.id)}"><summary><span class="st">${esc(s.title)}</span><span class="sc" data-count></span></summary><ul>${s.items.map(it => row(s, it)).join('')}</ul></details>`;
const html = `<meta name="claude:viewport-fit-sides" content="cover">
<title>大頭 022 行李核對清單</title>
<style>
/* 版面：頂端固定一條進度列（號碼布 022＋已勾數），下面是一欄到底的可點清單；不用卡片，分類用粗標題與細線分隔 */
:root {
  --bg: #F5F6F8; --surface: #FFFFFF; --ink: #0B1F3A; --muted: #5B6673; --line: #DCE1E8;
  --accent: #D7262F; --ok: #1F8A4C; --tag-bg: #FBE7C3; --tag-ink: #6E4300; --bib-bg: #FFFFFF; --bib-ink: #0B1F3A;
  --font: -apple-system, BlinkMacSystemFont, "PingFang TC", "Noto Sans TC", "Microsoft JhengHei", "Segoe UI", sans-serif;
}
@media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) {
  --bg: #0B131F; --surface: #121D2C; --ink: #E9EEF4; --muted: #9BA8B7; --line: #233146;
  --accent: #FF5D62; --ok: #4CC583; --tag-bg: #3B2C0E; --tag-ink: #F3CB7C; --bib-bg: #E9EEF4; --bib-ink: #0B131F; color-scheme: dark } }
:root[data-theme="dark"] {
  --bg: #0B131F; --surface: #121D2C; --ink: #E9EEF4; --muted: #9BA8B7; --line: #233146;
  --accent: #FF5D62; --ok: #4CC583; --tag-bg: #3B2C0E; --tag-ink: #F3CB7C; --bib-bg: #E9EEF4; --bib-ink: #0B131F; color-scheme: dark }
* { box-sizing: border-box; }
body { background: var(--bg); color: var(--ink); font-family: var(--font); font-size: 16px; line-height: 1.4; -webkit-text-size-adjust: 100%; }
body::before { content: ""; position: fixed; inset: 0 0 auto; height: env(safe-area-inset-top, 0px); z-index: 10; background: var(--surface); }
.top { position: sticky; top: env(safe-area-inset-top, 0px); z-index: 5; background: var(--surface); border-bottom: 1px solid var(--line); }
.wrap { max-width: 640px; margin: 0 auto; padding-left: max(16px, env(safe-area-inset-left, 0px)); padding-right: max(16px, env(safe-area-inset-right, 0px)); }
.top .wrap { padding-block: 10px 8px; display: grid; gap: 8px; }
.head { display: flex; align-items: center; gap: 12px; }
.bib { flex: none; background: var(--bib-bg); color: var(--bib-ink); border: 2px solid var(--ink); border-radius: 6px; padding: 2px 8px; font-weight: 800; font-size: 20px; letter-spacing: 1px; font-variant-numeric: tabular-nums; }
.ttl { flex: 1; min-width: 0; }
h1 { margin: 0; font-size: 17px; font-weight: 800; text-wrap: balance; }
.ttl p { margin: 0; font-size: 12.5px; color: var(--muted); }
.prog { flex: none; text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.prog b { font-size: 26px; font-weight: 800; }
.prog span { font-size: 13px; color: var(--muted); }
.bar { height: 6px; border-radius: 3px; background: var(--line); overflow: hidden; }
.bar i { display: block; height: 100%; width: 0; background: var(--ok); border-radius: 3px; transition: width .25s ease; }
.tools { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
.seg { display: inline-flex; border: 1px solid var(--line); border-radius: 999px; overflow: hidden; }
.seg button { appearance: none; border: 0; background: transparent; color: var(--muted); font: inherit; font-size: 14px; padding: 7px 14px; min-height: 36px; cursor: pointer; }
.seg button[aria-pressed="true"] { background: var(--ink); color: var(--surface); font-weight: 700; }
#status { margin-left: auto; font-size: 12px; color: var(--muted); text-align: right; min-width: 0; }
main.wrap { padding-block: 6px 28px; }
.lead { margin: 10px 0 2px; font-size: 13px; color: var(--muted); }
.sec { border-bottom: 1px solid var(--line); }
.sec summary { list-style: none; display: flex; align-items: center; justify-content: space-between; gap: 10px; padding: 16px 0 10px; cursor: pointer; }
.sec summary::-webkit-details-marker { display: none; }
.st { font-size: 16px; font-weight: 800; min-width: 0; }
.sec.hot .st { color: var(--accent); }
.sc { flex: none; font-size: 13px; color: var(--muted); font-variant-numeric: tabular-nums; white-space: nowrap; }
.sec.complete .sc { color: var(--ok); font-weight: 700; }
.sec:not([open]) summary { padding-bottom: 16px; }
ul { list-style: none; margin: 0; padding: 0 0 6px; }
li { border-top: 1px solid var(--line); }
.row { display: flex; align-items: flex-start; gap: 12px; padding: 11px 0; min-height: 48px; cursor: pointer; -webkit-tap-highlight-color: transparent; }
.row input { appearance: none; -webkit-appearance: none; flex: none; width: 26px; height: 26px; margin: 0; border: 2px solid var(--ink); border-radius: 7px; background: var(--surface); position: relative; cursor: pointer; }
.row input:checked { background: var(--ok); border-color: var(--ok); }
.row input:checked::after { content: ""; position: absolute; left: 7px; top: 2px; width: 7px; height: 13px; border: solid var(--surface); border-width: 0 3px 3px 0; transform: rotate(45deg); }
.row input:focus-visible, .seg button:focus-visible, .sec summary:focus-visible, .reset button:focus-visible { outline: 3px solid var(--accent); outline-offset: 2px; }
.txt { flex: 1; min-width: 0; display: grid; gap: 2px; padding-top: 2px; }
.t { font-weight: 650; overflow-wrap: anywhere; }
.n { font-size: 13.5px; color: var(--muted); overflow-wrap: anywhere; }
li.done .t { color: var(--muted); font-weight: 500; }
.qty { display: inline-block; padding: 0 7px; border-radius: 999px; background: var(--tag-bg); color: var(--tag-ink); font-size: 12.5px; font-weight: 600; white-space: nowrap; }
.sug { display: inline-block; padding: 0 6px; border: 1px solid var(--line); border-radius: 5px; color: var(--muted); font-size: 12px; font-weight: 500; white-space: nowrap; }
.only-todo li.done { display: none; }
.only-todo .sec.complete ul::after { content: "這一類都勾完了"; display: block; padding: 4px 0 10px; font-size: 13.5px; color: var(--ok); }
.alldone { margin: 14px 0 0; padding: 12px 14px; border-radius: 10px; background: var(--surface); border: 1px solid var(--ok); color: var(--ok); font-weight: 700; }
.reset { margin: 22px 0 0; display: flex; align-items: center; gap: 10px; flex-wrap: wrap; font-size: 14px; color: var(--muted); }
.reset button { appearance: none; font: inherit; font-size: 14px; min-height: 40px; padding: 8px 14px; border-radius: 8px; border: 1px solid var(--line); background: var(--surface); color: var(--ink); cursor: pointer; }
.reset button.danger { border-color: var(--accent); color: var(--accent); font-weight: 700; }
.foot { margin: 18px 0 0; font-size: 12px; color: var(--muted); }
@media (prefers-reduced-motion: reduce) { .bar i { transition: none; } }
</style>
<header class="top"><div class="wrap">
  <div class="head">
    <div class="bib" aria-label="號碼 022">022</div>
    <div class="ttl"><h1>大頭 個人行李核對</h1><p>2026 FXT VI・11/13 比賽・選手本人</p></div>
    <div class="prog" aria-live="polite"><b id="n">0</b><span> / ${total}</span></div>
  </div>
  <div class="bar" aria-hidden="true"><i id="bar"></i></div>
  <div class="tools">
    <div class="seg" role="group" aria-label="顯示範圍"><button type="button" id="f-all" aria-pressed="true">全部</button><button type="button" id="f-todo" aria-pressed="false">只看沒勾的</button></div>
    <span id="status" role="status"></span>
  </div>
</div></header>
<main class="wrap">
  <p class="lead">點一下就打勾，再點一下取消。標「建議」的是計畫沒寫到、建議一併確認的項目。點分類標題可以收合。</p>
  <p class="alldone" id="alldone" hidden>全部勾完了。出發前再看一眼第一類「最容易忘」。</p>
  ${D.sections.map(section).join('\n  ')}
  <div class="reset" id="reset">
    <button type="button" id="r-ask">清除全部勾選</button>
    <span id="r-confirm" hidden>確定要全部清除嗎？ <button type="button" class="danger" id="r-yes">確定清除</button> <button type="button" id="r-no">取消</button></span>
  </div>
  <p class="foot">依 2026 FXT 選手手冊與團隊補給計畫（${esc(D.version || '')}）整理；以 11/12 賽前說明會公布內容為最終依據。</p>
</main>
<script>
(function () {
  var KEY = 'fxt022-packing-v1';
  var boxes = Array.prototype.slice.call(document.querySelectorAll('input[type=checkbox][data-id]'));
  var valid = {}; boxes.forEach(function (b) { valid[b.dataset.id] = true; });
  var total = boxes.length;
  var done = {};            // id -> true
  var unsynced = false;     // 本機有尚未確認寫入雲端的變更（會存進 localStorage）
  var dirty = false, writing = false, ref = null, timer = null, retried = false;
  var statusEl = document.getElementById('status');
  var MSG = { connecting: '連線中，勾選先存在這支手機', synced: '已儲存', saving: '儲存中…', local: '勾選存在這支手機的瀏覽器',
    waiting: '還沒連上，勾選先存在這支手機', error: '暫時存不到雲端，勾選先存在這支手機', readonly: '沒有修改權限，勾選只存在這支手機' };
  function setStatus(k) { statusEl.textContent = MSG[k] || ''; }

  function loadLocal() { try { var o = JSON.parse(localStorage.getItem(KEY) || '{}'); (o.done || []).forEach(function (id) { if (valid[id]) done[id] = true; }); unsynced = !!o.unsynced; } catch (e) {} }
  function saveLocal() { try { localStorage.setItem(KEY, JSON.stringify({ done: Object.keys(done), unsynced: unsynced })); } catch (e) {} }

  function paint() {
    var n = 0;
    boxes.forEach(function (b) { var on = !!done[b.dataset.id]; b.checked = on; b.closest('li').classList.toggle('done', on); if (on) n++; });
    document.getElementById('n').textContent = n;
    document.getElementById('bar').style.width = (total ? Math.round(n / total * 100) : 0) + '%';
    document.getElementById('alldone').hidden = n !== total;
    Array.prototype.forEach.call(document.querySelectorAll('.sec'), function (sec) {
      var bs = sec.querySelectorAll('input[data-id]'), k = 0;
      Array.prototype.forEach.call(bs, function (b) { if (b.checked) k++; });
      sec.querySelector('[data-count]').textContent = k === bs.length ? '完成 ' + k + '/' + bs.length : k + '/' + bs.length;
      sec.classList.toggle('complete', k === bs.length);
    });
  }

  function changed() { dirty = true; unsynced = true; saveLocal(); paint(); clearTimeout(timer); timer = setTimeout(flush, 600); }

  function flush() {
    if (!ref || writing || !dirty) return;
    writing = true; dirty = false; setStatus('saving');
    var body = { done: {} }; Object.keys(done).forEach(function (id) { if (valid[id]) body.done[id] = true; });
    ref.set(body).then(function () {
      writing = false; retried = false;
      if (dirty) { flush(); return; }
      unsynced = false; saveLocal(); setStatus('synced');
    }).catch(function (e) {
      writing = false; dirty = true;
      var code = e && e.code;
      if (code === 'invalid_argument' || code === 'not_granted' || code === 'revoked') { ref = null; setStatus('readonly'); return; }
      setStatus('error');
      if (!retried) { retried = true; setTimeout(flush, 1200 + Math.random() * 1500); }
    });
  }

  function connect() {
    var c = window.claude;
    if (!c || typeof c.use !== 'function') { setStatus('local'); return; }
    setStatus('connecting');
    c.use('db').then(function (db) {
      if (!db) { setStatus('local'); return; }
      ref = db.doc('checklist/athlete');
      var wait = setTimeout(function () { setStatus('waiting'); }, 10000);
      ref.onSnapshot(function (snap) {
        if (!snap.metadata || !snap.metadata.fromCache) clearTimeout(wait);
        if (unsynced || dirty || writing) { if (!writing) { dirty = true; flush(); } return; }   // 本機變更優先，先寫上去
        if (snap.exists) {
          var d = snap.data() || {}, r = d.done || {}, next = {};
          Object.keys(r).forEach(function (id) { if (r[id] && valid[id]) next[id] = true; });
          done = next; saveLocal(); paint();
        }
        setStatus('synced');
      }, function () { clearTimeout(wait); ref = null; setStatus('error'); });
    }).catch(function () { setStatus('local'); });
  }

  boxes.forEach(function (b) { b.addEventListener('change', function () { if (b.checked) done[b.dataset.id] = true; else delete done[b.dataset.id]; changed(); }); });
  var fAll = document.getElementById('f-all'), fTodo = document.getElementById('f-todo');
  function setFilter(todo) { document.body.classList.toggle('only-todo', todo); fAll.setAttribute('aria-pressed', String(!todo)); fTodo.setAttribute('aria-pressed', String(todo)); try { localStorage.setItem(KEY + '-todo', todo ? '1' : ''); } catch (e) {} }
  fAll.addEventListener('click', function () { setFilter(false); });
  fTodo.addEventListener('click', function () { setFilter(true); });
  var ask = document.getElementById('r-ask'), conf = document.getElementById('r-confirm');
  ask.addEventListener('click', function () { ask.hidden = true; conf.hidden = false; });
  document.getElementById('r-no').addEventListener('click', function () { ask.hidden = false; conf.hidden = true; });
  document.getElementById('r-yes').addEventListener('click', function () { done = {}; ask.hidden = false; conf.hidden = true; changed(); });

  loadLocal(); paint();
  try { if (localStorage.getItem(KEY + '-todo')) setFilter(true); } catch (e) {}
  connect();
})();
</script>
`;
fs.mkdirSync(path.dirname(OUT), { recursive: true });
fs.writeFileSync(OUT, html);
console.log('written', OUT, 'items', total, 'bytes', Buffer.byteLength(html));
