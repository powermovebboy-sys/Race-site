// 教練（002）比賽日簡報 — 簡單、自然、大字。讀 deck/coach/copy.json 產生 ../coach/2026_FXT_002_比賽日.pptx
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');
const COPY = process.argv[2] || path.join(__dirname, '..', 'coach', 'copy.json');
const OUT = process.argv[3] || path.join(__dirname, '..', 'coach', '2026_FXT_002_比賽日.pptx');
const PHOTOS = path.join(__dirname, '..', 'coach', 'photos');
const ICONS = path.join(__dirname, 'assets');
const D = JSON.parse(fs.readFileSync(COPY, 'utf8'));

const FONT = 'Microsoft JhengHei';
const PAPER = 'F7F4EE', INK = '1F2A24', MUTED = '8A8F85', GREEN = '2F5D50', LEAF = 'DCE8E1', ORANGE = 'C9742B', SAND = 'EFE7D8', WHITE = 'FFFFFF', LINE = 'D8D2C6';
const W = 13.333, H = 7.5;
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = D.deck_title || '2026 FXT 002';
const T = (s, text, o) => s.addText(text, Object.assign({ fontFace: FONT, color: INK, margin: 2 }, o));
const rect = (s, x, y, w, h, fill, o = {}) => s.addShape(pres.shapes.RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 } }, o));
const rrect = (s, x, y, w, h, fill, line, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.1, fill: { color: fill }, line: line ? { color: line, width: 1 } : { color: fill, width: 0 } }, o));
const photoPath = f => f && fs.existsSync(path.join(PHOTOS, f)) ? path.join(PHOTOS, f) : null;
const icon = (name, color = 'white') => { const p = path.join(ICONS, `icon_${name}_${color}.png`); return fs.existsSync(p) ? p : null; };
const photo = (s, f, x, y, w, h) => {
  const p = photoPath(f); if (!p) return false;
  rrect(s, x - 0.06, y - 0.06, w + 0.12, h + 0.12, WHITE, LINE);
  s.addImage({ path: p, x, y, w, h, sizing: { type: 'cover', w, h } });
  return true;
};
const base = (s) => { rect(s, 0, 0, W, H, PAPER); rect(s, 0, H - 0.08, W, 0.08, GREEN); };
const footer = (s, n) => {
  T(s, `002 謝漢霖 ・ 2026 FXT VI ・ 11/13`, { x: 0.5, y: 7.02, w: 6, h: 0.3, fontSize: 9, color: MUTED });
  T(s, String(n), { x: W - 1.0, y: 7.02, w: 0.5, h: 0.3, fontSize: 9, color: MUTED, align: 'right' });
};
const ICON_BY_PLACE = [[/早餐|起床/, 'food'], [/游泳|下水|T1|開賽/, 'swim'], [/上岸|車衣/, 'bike'], [/7-11|門市/, 'store'], [/管制|草坪頭/, 'ban'], [/夫妻樹|塔塔加/, 'mountain'], [/冰沙|豐丘/, 'cup'], [/水里|午餐/, 'bread'], [/前導|埔里/, 'car'], [/T2|賽德克/, 'shoe'], [/路跑|跑步|42/, 'run']];
const pickIcon = (sl) => { const key = `${sl.place || ''}${sl.title || ''}`; for (const [re, n] of ICON_BY_PLACE) if (re.test(key)) return n; return 'pin'; };

let n = 0;
const steps = D.slides.filter(s => s.kind === 'step');
D.slides.forEach(sl => {
  n += 1;
  const s = pres.addSlide();
  base(s);
  if (sl.kind === 'cover') {
    rect(s, 0, 0, W, H, GREEN);
    T(s, '002', { x: 0.8, y: 0.9, w: 4, h: 1.6, fontSize: 96, bold: true, color: WHITE });
    T(s, '謝漢霖', { x: 0.85, y: 2.45, w: 5, h: 0.7, fontSize: 30, bold: true, color: WHITE });
    T(s, sl.title || D.deck_title, { x: 0.85, y: 3.15, w: 6.6, h: 0.9, fontSize: 36, bold: true, color: WHITE });
    T(s, D.deck_subtitle || '', { x: 0.85, y: 4.05, w: 6.6, h: 0.4, fontSize: 14, color: 'DCE8E1' });
    T(s, (sl.bullets || []).join('　・　'), { x: 0.85, y: 4.5, w: 6.6, h: 0.4, fontSize: 14, color: 'DCE8E1' });
    if (sl.crew && sl.crew.length) T(s, sl.crew.map(c => '• ' + c).join('\n'), { x: 0.85, y: 5.05, w: 6.6, h: 0.65, fontSize: 12.5, color: 'DCE8E1', valign: 'top' });
    if (sl.reminder) T(s, '提醒　' + sl.reminder, { x: 0.85, y: 5.72, w: 6.6, h: 0.35, fontSize: 11, color: 'F2C14E' });
    if (sl.mood) T(s, sl.mood, { x: 0.85, y: 6.12, w: 6.4, h: 0.5, fontSize: 20, bold: true, color: 'F2C14E' });
    T(s, '2026 Formosa Xtreme Triathlon VI ・ 11/13（五）・ 日月潭 → 合歡主峰', { x: 0.85, y: 6.7, w: 7, h: 0.3, fontSize: 10, color: 'B9CFC5' });
    if (photoPath(sl.photo || 'coach.jpg')) s.addImage({ path: photoPath(sl.photo || 'coach.jpg'), x: 7.9, y: 0.0, w: 5.433, h: H, sizing: { type: 'cover', w: 5.433, h: H, pos: 'top' } });
    return;
  }
  if (sl.kind === 'overview') {
    T(s, sl.title, { x: 0.5, y: 0.4, w: 12.3, h: 0.7, fontSize: 28, bold: true });
    // timeline of step slides
    const items = steps.map(st => ({ time: st.time, place: st.short || (st.place || st.title).split('｜')[0] }));
    const x0 = 0.8, x1 = W - 0.8, ly = 3.0;
    rect(s, x0, ly - 0.03, x1 - x0, 0.06, GREEN);
    items.forEach((it, i) => {
      const x = x0 + (x1 - x0) * i / Math.max(1, items.length - 1);
      s.addShape(pres.shapes.OVAL, { x: x - 0.12, y: ly - 0.12, w: 0.24, h: 0.24, fill: { color: i % 2 ? ORANGE : GREEN }, line: { color: PAPER, width: 1.5 } });
      const up = i % 2 === 0;
      T(s, it.time, { x: x - 0.8, y: up ? ly - 1.0 : ly + 0.25, w: 1.6, h: 0.42, fontSize: 17, bold: true, color: i % 2 ? ORANGE : GREEN, align: 'center' });
      T(s, it.place, { x: x - 0.9, y: up ? ly - 1.5 : ly + 0.65, w: 1.8, h: 0.45, fontSize: 12, color: INK, align: 'center', valign: up ? 'bottom' : 'top' });
    });
    T(s, (sl.bullets || []).map(b => '• ' + b).join('\n'), { x: 0.8, y: 4.6, w: 7.5, h: 1.6, fontSize: 15, color: INK, valign: 'top', paraSpaceAfter: 6 });
    if (sl.crew && sl.crew.length) { rrect(s, 8.6, 4.6, 4.2, 1.6, LEAF, null); T(s, '團隊', { x: 8.8, y: 4.68, w: 2, h: 0.3, fontSize: 11, bold: true, color: GREEN }); T(s, sl.crew.map(b => '• ' + b).join('\n'), { x: 8.8, y: 4.98, w: 3.9, h: 1.15, fontSize: 12, color: INK, valign: 'top' }); }
    if (sl.reminder) { rrect(s, 0.8, 6.3, 12.0, 0.5, SAND, null); T(s, '提醒　' + sl.reminder, { x: 0.95, y: 6.3, w: 11.7, h: 0.5, fontSize: 11.5, color: ORANGE, bold: true, valign: 'middle' }); }
    else if (sl.mood) T(s, sl.mood, { x: 0.8, y: 6.35, w: 8, h: 0.5, fontSize: 18, bold: true, color: GREEN });
    footer(s, n); return;
  }
  if (sl.kind === 'table') {
    T(s, sl.title, { x: 0.5, y: 0.4, w: 12.3, h: 0.7, fontSize: 28, bold: true });
    const rows = (sl.table && sl.table.length) ? sl.table : [['檢查點', '關門', '我的時間', '餘裕']];
    const tr = rows.map((r, i) => r.map(c => ({ text: c, options: { fontFace: FONT, fontSize: i === 0 ? 13 : 15, bold: i === 0, color: i === 0 ? WHITE : INK, fill: { color: i === 0 ? GREEN : (i % 2 ? WHITE : 'FBF9F4') }, valign: 'middle', margin: 5 } })));
    s.addTable(tr, { x: 0.5, y: 1.3, w: 8.1, colW: [2.8, 1.1, 2.1, 2.1], rowH: 0.5, border: { type: 'solid', pt: 0.5, color: LINE } });
    rrect(s, 8.9, 1.3, 3.93, 3.4, LEAF, null);
    T(s, sl.bullets.concat(sl.crew || []).map(b => '• ' + b).join('\n'), { x: 9.1, y: 1.45, w: 3.6, h: 3.1, fontSize: 13, color: INK, valign: 'top', paraSpaceAfter: 6 });
    if (sl.reminder) { rrect(s, 8.9, 4.9, 3.93, 1.2, SAND, null); T(s, sl.reminder, { x: 9.1, y: 4.98, w: 3.6, h: 1.05, fontSize: 12, color: ORANGE, bold: true, valign: 'top' }); }
    if (sl.mood) T(s, sl.mood, { x: 0.5, y: 6.35, w: 8, h: 0.5, fontSize: 18, bold: true, color: GREEN });
    footer(s, n); return;
  }
  if (sl.kind === 'closing') {
    rect(s, 0, 0, W, H, GREEN);
    T(s, sl.title, { x: 0.8, y: 2.2, w: 11.7, h: 1.4, fontSize: 48, bold: true, color: WHITE, align: 'center' });
    T(s, (sl.bullets || []).join('　・　'), { x: 0.8, y: 3.7, w: 11.7, h: 0.6, fontSize: 18, color: 'DCE8E1', align: 'center' });
    if (sl.mood) T(s, sl.mood, { x: 0.8, y: 4.6, w: 11.7, h: 0.6, fontSize: 22, bold: true, color: 'F2C14E', align: 'center' });
    if (sl.crew && sl.crew.length) T(s, sl.crew.join('　・　'), { x: 0.8, y: 5.45, w: 11.7, h: 0.4, fontSize: 13, color: 'DCE8E1', align: 'center' });
    if (sl.reminder) T(s, '提醒　' + sl.reminder, { x: 0.8, y: 5.9, w: 11.7, h: 0.4, fontSize: 11.5, color: 'F2C14E', align: 'center' });
    T(s, '#See you at the top.', { x: 0.8, y: 6.5, w: 11.7, h: 0.4, fontSize: 12, color: 'B9CFC5', align: 'center' });
    return;
  }
  // ---- step slide ----
  const idx = steps.indexOf(sl) + 1;
  T(s, sl.time || '', { x: 0.5, y: 0.35, w: 4.6, h: 1.2, fontSize: 60, bold: true, color: GREEN });
  T(s, `${idx} / ${steps.length}`, { x: 11.3, y: 0.45, w: 1.55, h: 0.3, fontSize: 10, color: MUTED, align: 'right' });
  if (sl.place) { rrect(s, 0.55, 1.6, Math.min(6.0, 0.5 + sl.place.length * 0.24), 0.42, LEAF, null); T(s, sl.place, { x: 0.55, y: 1.6, w: Math.min(6.0, 0.5 + sl.place.length * 0.24), h: 0.42, fontSize: 13, bold: true, color: GREEN, align: 'center', valign: 'middle' }); }
  T(s, sl.title, { x: 0.5, y: 2.15, w: 6.4, h: 0.75, fontSize: 24, bold: true, color: INK, valign: 'middle' });
  const bl = (sl.bullets || []).slice(0, 3);
  bl.forEach((b, i) => {
    const y = 3.05 + i * 0.62;
    s.addShape(pres.shapes.OVAL, { x: 0.58, y: y + 0.17, w: 0.16, h: 0.16, fill: { color: ORANGE }, line: { color: ORANGE, width: 0 } });
    T(s, b, { x: 0.9, y, w: 6.0, h: 0.5, fontSize: 18, color: INK, valign: 'middle' });
  });
  let cy = 3.05 + bl.length * 0.62 + 0.15;
  if (sl.crew && sl.crew.length) {
    const h = 0.42 + sl.crew.length * 0.36;
    rrect(s, 0.55, cy, 6.3, h, LEAF, null);
    const ic = icon('vest', 'teal'); if (ic) s.addImage({ path: ic, x: 0.72, y: cy + 0.12, w: 0.3, h: 0.3 });
    T(s, '團隊', { x: 1.1, y: cy + 0.08, w: 2, h: 0.34, fontSize: 11.5, bold: true, color: GREEN });
    T(s, sl.crew.map(c => '• ' + c).join('\n'), { x: 1.1, y: cy + 0.4, w: 5.6, h: sl.crew.length * 0.36, fontSize: 13, color: INK, valign: 'top' });
    cy += h + 0.12;
  }
  if (sl.reminder) {
    const lines = Math.ceil((sl.reminder.length + 3) / 34);
    const h = 0.22 + lines * 0.24;
    const ry = Math.min(cy, 6.9 - h);
    rrect(s, 0.55, ry, 6.3, h, SAND, null);
    T(s, '提醒　' + sl.reminder, { x: 0.7, y: ry, w: 6.05, h, fontSize: 11, color: ORANGE, bold: true, valign: 'middle' });
  }
  if (sl.mood) T(s, sl.mood, { x: 7.2, y: 6.45, w: 5.6, h: 0.45, fontSize: 16, bold: true, color: GREEN, align: 'right' });
  // photo(s)
  const px = 7.2, py = 0.5, pw = 5.63, ph = 5.8;
  if (sl.photo === 'breakfast') {
    const files = ['b_banana.jpg', 'b_guava.jpg', 'b_coffee.jpg', 'b_riceball.jpg', 'b_biscuit.jpg', 'b_noodle.jpg'];
    const gw = (pw - 0.2) / 2, gh = (ph - 0.4) / 3;
    files.forEach((f, i) => photo(s, f, px + (i % 2) * (gw + 0.2), py + Math.floor(i / 2) * (gh + 0.2), gw, gh));
  } else if (!photo(s, sl.photo, px, py, pw, ph)) {
    rrect(s, px, py, pw, ph, LEAF, null);
    const ic = icon(pickIcon(sl), 'teal'); if (ic) s.addImage({ path: ic, x: px + pw / 2 - 1.0, y: py + ph / 2 - 1.5, w: 2.0, h: 2.0 });
    T(s, sl.place || sl.title, { x: px + 0.3, y: py + ph / 2 + 0.7, w: pw - 0.6, h: 0.6, fontSize: 20, bold: true, color: GREEN, align: 'center' });
    T(s, (sl.bullets || [])[0] || '', { x: px + 0.3, y: py + ph / 2 + 1.3, w: pw - 0.6, h: 0.5, fontSize: 14, color: MUTED, align: 'center' });
  }
  footer(s, n);
});
pres.writeFile({ fileName: OUT }).then(() => console.log('written', OUT));
