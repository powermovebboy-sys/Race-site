// 行李核對清單：讀 deck/checklist/checklist.json → HTML → A4 直式 PDF（Playwright/Chromium）
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const SRC = process.argv[2] || path.join(__dirname, '..', 'checklist', 'checklist_athlete.json');
const OUT = process.argv[3] || path.join(__dirname, '..', '2026_FXT_行李核對清單.pdf');
const D = JSON.parse(fs.readFileSync(SRC, 'utf8'));
const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const total = D.sections.reduce((n, s) => n + s.items.length, 0);
const item = it => `<li class="${it.s ? 'sug' : ''}"><span class="box"></span><span class="txt"><b>${it.s ? '＊' : ''}${esc(it.t)}</b>${it.q ? `<span class="qty">${esc(it.q)}</span>` : ''}${it.new ? '<span class="new">新</span>' : ''}${it.n ? `<span class="note">${esc(it.n)}</span>` : ''}</span></li>`;
const section = s => `<section class="sec ${s.id === 'last' ? 'hot' : ''}"><h2>${esc(s.title)}<span class="cnt">${s.items.length}</span></h2>${s.note ? `<p class="secnote">${esc(s.note)}</p>` : ''}<ul>${s.items.map(item).join('')}</ul></section>`;
const html = `<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><style>
@page { size: A4; margin: 12mm 11mm 13mm 11mm; }
* { box-sizing: border-box; }
html { font-family: "Microsoft JhengHei", "PingFang TC", "Noto Sans CJK TC", "WenQuanYi Zen Hei", sans-serif; color: #0B0F14; }
body { margin: 0; font-size: ${D.fontPt || 9.6}pt; line-height: 1.32; }
header { border-bottom: 2.5pt solid #D7262F; padding-bottom: 5pt; margin-bottom: 8pt; display: flex; align-items: flex-end; justify-content: space-between; gap: 10pt; }
h1 { font-size: 19pt; margin: 0; letter-spacing: .5pt; }
.sub { font-size: 9pt; color: #5B6673; margin-top: 2pt; }
.legend { font-size: 8pt; color: #5B6673; text-align: right; white-space: nowrap; }
.cols { column-count: 2; column-gap: 9mm; }
.sec { break-inside: avoid; margin: 0 0 8pt 0; border: .75pt solid #C9D3DE; border-radius: 5pt; overflow: hidden; }
.sec h2 { margin: 0; font-size: 1.06em; background: #0B1F3A; color: #fff; padding: 3.5pt 7pt; display: flex; justify-content: space-between; align-items: center; gap: 6pt; }
.sec.hot h2 { background: #D7262F; }
.cnt { font-size: .75em; font-weight: 400; opacity: .8; white-space: nowrap; }
.cnt::after { content: " 項"; }
.secnote { margin: 3pt 7pt 0; font-size: .83em; color: #5B6673; }
ul { list-style: none; margin: 0; padding: 3pt 7pt 4pt; }
li { display: flex; gap: 6pt; align-items: flex-start; padding: 2.2pt 0; border-bottom: .5pt dotted #D9DEE5; break-inside: avoid; }
li:last-child { border-bottom: 0; }
.box { flex: 0 0 1em; height: 1em; border: 1pt solid #0B0F14; border-radius: 1.5pt; margin-top: 1.4pt; background: #fff; }
.txt { flex: 1; }
.txt b { font-weight: 700; }
.qty { display: inline-block; margin-left: 5pt; padding: 0 4pt; border-radius: 7pt; background: #F3E3C2; font-size: .83em; white-space: nowrap; }
.note { display: block; font-size: .83em; color: #5B6673; }
.new { display: inline-block; margin-left: 5pt; padding: 0 3.5pt; border-radius: 3pt; background: #D7262F; color: #fff; font-size: .72em; font-weight: 700; }
li.sug .txt b { font-weight: 400; }
li.sug .box { border-style: dashed; }
footer { margin-top: 4pt; font-size: 7.5pt; color: #7A8794; }
</style></head><body>
<header><div><h1>${esc(D.title)}</h1><div class="sub">${esc(D.subtitle)}・共 ${total} 項</div></div><div class="legend">${esc(D.legend).split('　').join('<br>')}</div></header>
<div class="cols">${D.sections.map(section).join('')}</div>
<footer>依 2026 FXT 選手手冊與團隊補給計畫（${esc(D.version || '')}）整理。數量以補給明細為準；以 11/12 賽前說明會公布內容為最終依據。</footer>
</body></html>`;
(async () => {
  const htmlPath = SRC.replace(/\.json$/, '.html');
  fs.writeFileSync(htmlPath, html);
  const browser = await chromium.launch();
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'load' });
  await page.pdf({ path: OUT, format: 'A4', printBackground: true, displayHeaderFooter: true, headerTemplate: '<span></span>',
    footerTemplate: '<div style="width:100%;font-size:7px;color:#7A8794;text-align:center;font-family:sans-serif;"><span class="pageNumber"></span> / <span class="totalPages"></span></div>',
    margin: { top: '12mm', bottom: '13mm', left: '11mm', right: '11mm' } });
  await browser.close();
  console.log('written', OUT, 'items', total);
})().catch(e => { console.error(e); process.exit(1); });
