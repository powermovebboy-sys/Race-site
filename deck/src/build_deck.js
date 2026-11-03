// 2026 FXT 補給作戰計畫 — deck builder (pptxgenjs structured deck)
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const { applyTheme } = require(path.join(__dirname, 'vendor', 'apply_theme.js'));

const OUT = process.argv[2] || path.join(__dirname, 'FXT_2026_補給作戰計畫.pptx');
const ASSETS = path.join(__dirname, 'assets');
const PHOTOS = path.join(__dirname, 'photos');
const icon = (name, color = 'white') => path.join(ASSETS, `icon_${name}_${color}.png`);
const photoPath = f => path.join(PHOTOS, f);
const PHOTO_MAP = fs.existsSync(path.join(__dirname, 'photo_map.json')) ? JSON.parse(fs.readFileSync(path.join(__dirname, 'photo_map.json'), 'utf8')) : {};
const crewPhoto = f => (f && fs.existsSync(path.join(PHOTOS, 'crew', f))) ? path.join(PHOTOS, 'crew', f) : null;

const THEME = {
  name: 'FXT 2026 Crew',
  headFontFace: 'Microsoft JhengHei',
  bodyFontFace: 'Microsoft JhengHei',
  colors: {
    dk1: '0B0F14', lt1: 'FFFFFF', dk2: '3C4858', lt2: 'F3F5F7',
    accent1: 'D7262F', // FXT red
    accent2: '0E7C86', // teal / bike
    accent3: 'F26B1F', // orange / run
    accent4: '2F80ED', // blue / swim
    accent5: 'E0A800', // gold / cut-off & finish
    accent6: '2E9E5B', // green / pacer allowed
    hlink: '0E7C86', folHlink: '7A8794',
  },
};
const HEX = THEME.colors;
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = '2026 FXT 補給作戰計畫';
pres.author = 'FXT 補給團隊';
const C = pres.SchemeColor;
const W = 13.333, H = 7.5;
const MUTED = '7A8794', BORDER = 'E3E7EC', NAVY = '0B1F3A', INK = '0B0F14', RED = 'D7262F', TEAL = '0E7C86', ORANGE = 'F26B1F', BLUE = '2F80ED', GOLD = 'E0A800', GREEN = '2E9E5B', CARD = 'F3F5F7', WHITE = 'FFFFFF';

// ---------- layouts ----------
pres.defineSlideMaster({
  title: 'CONTENT', background: { color: WHITE },
  objects: [
    { text: { text: '2026 FXT 補給作戰計畫 ｜ 團隊內部使用 ｜ 以賽前說明會公布內容為最終依據', options: { x: 0.5, y: 7.08, w: 9, h: 0.3, fontSize: 9, color: MUTED, margin: 0 } } },
    { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.32, w: 12.3, h: 0.75, fontSize: 28, bold: true, color: INK, margin: 0, valign: 'middle', align: 'left' }, text: '' } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.6, h: 0.3, fontSize: 9, color: MUTED, align: 'right' },
});
pres.defineSlideMaster({
  title: 'DARK', background: { path: path.join(ASSETS, 'bg_section.png') },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.32, w: 12.3, h: 0.75, fontSize: 28, bold: true, color: WHITE, margin: 0, valign: 'middle', align: 'left' }, text: '' } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.6, h: 0.3, fontSize: 9, color: 'A9B4C2', align: 'right' },
});
pres.defineSlideMaster({
  title: 'NIGHT', background: { path: path.join(ASSETS, 'bg_night.png') },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.32, w: 12.3, h: 0.75, fontSize: 28, bold: true, color: WHITE, margin: 0, valign: 'middle', align: 'left' }, text: '' } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.6, h: 0.3, fontSize: 9, color: 'A9B4C2', align: 'right' },
});
pres.defineSlideMaster({ title: 'BLANK', background: { color: INK }, objects: [] });

// ---------- helpers ----------
let section = '開場';
function slide(master, title, opts = {}) {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  if (title) s.addText(title, { placeholder: 'title' });
  if (opts.notes) s.addNotes(opts.notes);
  return s;
}
function T(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: INK, valign: 'top' }, o)); }
function rect(s, x, y, w, h, fill, o = {}) { s.addShape(pres.shapes.RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 } }, o)); }
function card(s, x, y, w, h, o = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: o.fill || CARD }, line: { color: o.line || (o.fill || CARD), width: o.lineW || 0 }, shadow: o.shadow ? { type: 'outer', color: '000000', blur: 6, offset: 2, angle: 90, opacity: 0.18 } : undefined });
}
function cardHeader(s, x, y, w, h, color) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color }, line: { color, width: 0 } });
  s.addShape(pres.shapes.RECTANGLE, { x, y: y + h / 2, w, h: h / 2, fill: { color }, line: { color, width: 0 } });
}
function iconCircle(s, x, y, d, name, circle, iconColor = 'white') {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: circle }, line: { color: circle, width: 0 } });
  const p = d * 0.24; s.addImage({ path: icon(name, iconColor), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}
function pill(s, x, y, w, h, text, fill, color = WHITE, fontSize = 11) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: fill }, line: { color: fill, width: 0 } });
  s.addText(text, { x, y, w, h, fontSize, bold: true, color, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
}
function photo(s, file, x, y, w, h, o = {}) { s.addImage(Object.assign({ path: photoPath(file), x, y, w, h, sizing: { type: 'cover', w, h } }, o)); }
function bullets(items, o = {}) {
  return items.map((t, i) => ({ text: t, options: Object.assign({ bullet: { indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: o.gap == null ? 4 : o.gap }, o.run || {}) }));
}
// Icon + header + body card
function infoCard(s, x, y, w, h, { iconName, color, title, lines, fontSize = 12, titleSize = 15, fill = CARD, dark = false }) {
  card(s, x, y, w, h, { fill });
  iconCircle(s, x + 0.18, y + 0.18, 0.5, iconName, color);
  T(s, title, { x: x + 0.8, y: y + 0.18, w: w - 0.95, h: 0.5, fontSize: titleSize, bold: true, color: dark ? WHITE : INK, valign: 'middle' });
  if (lines && lines.length) T(s, bullets(lines), { x: x + 0.2, y: y + 0.82, w: w - 0.4, h: h - 0.95, fontSize, color: dark ? 'DCE3EA' : HEX.dk2, paraSpaceAfter: 3 });
}
function tbl(s, rows, { x, y, w, colW, fontSize = 11, headFill = NAVY, rowH, align, bold0 = false, border = BORDER }) {
  const data = rows.map((r, ri) => r.map((cell, ci) => {
    const isHead = ri === 0;
    const base = { fontSize, color: isHead ? WHITE : INK, bold: isHead || (bold0 && ci === 0), fill: { color: isHead ? headFill : (ri % 2 === 0 ? 'FAFBFC' : WHITE) }, valign: 'middle', align: (align && align[ci]) || 'left', margin: [3, 8, 3, 8], border: { type: 'solid', color: border, pt: 0.5 } };
    if (typeof cell === 'object' && cell !== null && cell.text !== undefined) return { text: cell.text, options: Object.assign(base, cell.options || {}) };
    return { text: String(cell), options: base };
  }));
  s.addTable(data, { x, y, w, colW, rowH, autoPage: false });
}
function stat(s, x, y, w, big, label, color = WHITE, labelColor = 'C9D3DE', bigSize = 34) {
  T(s, big, { x, y, w, h: 0.6, fontSize: bigSize, bold: true, color, align: 'center', valign: 'middle' });
  T(s, label, { x, y: y + 0.62, w, h: 0.3, fontSize: 11, color: labelColor, align: 'center' });
}
function divider(num, title, sub, photoFile) {
  const s = pres.addSlide({ masterName: 'BLANK', sectionTitle: section });
  photo(s, photoFile, 0, 0, W, H);
  rect(s, 0, 0, W, H, '000000', { fill: { color: '000000', transparency: 45 }, line: { color: '000000', width: 0 } });
  T(s, num, { x: 0.8, y: 2.3, w: 4, h: 0.6, fontSize: 20, bold: true, color: GOLD, charSpacing: 4 });
  T(s, title, { x: 0.8, y: 2.9, w: 11, h: 1.1, fontSize: 44, bold: true, color: WHITE });
  T(s, sub, { x: 0.8, y: 4.05, w: 11, h: 0.6, fontSize: 18, color: 'DCE3EA' });
  return s;
}

// ================= SLIDES =================
pres.addSection({ title: '開場' });

// 1. COVER
{
  const s = pres.addSlide({ masterName: 'BLANK', sectionTitle: section });
  photo(s, PHOTO_MAP.cover || 'cover_summit.jpg', 0, 0, W, 5.6);
  rect(s, 0, 0, W, 5.6, '000000', { fill: { color: '000000', transparency: PHOTO_MAP.cover ? 45 : 35 }, line: { color: '000000', width: 0 } });
  s.addImage({ path: photoPath('logo_fxt_xtri.jpg'), x: 0.6, y: 0.45, w: 3.0, h: 0.64 });
  T(s, 'FORMOSA XTREME TRIATHLON VI ・ 2026.11.13（五）', { x: 0.6, y: 2.55, w: 12, h: 0.45, fontSize: 16, bold: true, color: GOLD, charSpacing: 2 });
  T(s, '2026 FXT 補給作戰計畫', { x: 0.6, y: 3.0, w: 12, h: 1.1, fontSize: 52, bold: true, color: WHITE });
  T(s, '日月潭 伊達邵碼頭 → 塔塔加夫妻樹 → 合歡主峰 3,417 m ｜ 補給團隊任務說明與執行依據', { x: 0.6, y: 4.15, w: 12, h: 0.5, fontSize: 18, color: 'E6ECF2' });
  rect(s, 0, 5.6, W, 1.9, INK);
  const sx = 0.6, sw = 2.43;
  stat(s, sx + sw * 0, 5.95, sw, '3.8 km', '游泳・日月潭', BLUE);
  stat(s, sx + sw * 1, 5.95, sw, '196.7 km', '自行車・爬升 3,600 m+', TEAL);
  stat(s, sx + sw * 2, 5.95, sw, '42 km', '跑步・爬升 3,000 m+', ORANGE);
  stat(s, sx + sw * 3, 5.95, sw, '3,417 m', '終點・合歡主峰', GOLD);
  stat(s, sx + sw * 4, 5.95, sw, '17 hr', '06:00 出發 → 23:00 登山口關門', RED);
  s.addNotes('封面。照片取自 2025 FXT 選手手冊（大會官方影像）。數據來源：2026 FXT 選手手冊。');
}

// 2. 通往夢想的最後一哩路（主題＋開場影片）
{
  const s = pres.addSlide({ masterName: 'DARK', sectionTitle: section });
  s.addText('通往夢想的最後一哩路', { placeholder: 'title' });
  T(s, '「如果沒有天賦，\n那就一直重複。」', { x: 0.5, y: 1.45, w: 5.7, h: 1.9, fontSize: 34, bold: true, color: WHITE, valign: 'top' });
  T(s, '大頭 ・ 號碼 022 ・ Formosa Xtreme Triathlon VI', { x: 0.5, y: 3.45, w: 5.7, h: 0.4, fontSize: 14, bold: true, color: GOLD });
  T(s, bullets(['從 2020 年 FXT 誕生開始，這條路走了六年', '2024 停辦一年；2025 起賽道改寫：日月潭 → 塔塔加 → 合歡主峰', '2026.11.13，第六屆。登山口到主峰的最後 1.8 公里，我們一起走完']), { x: 0.5, y: 4.0, w: 5.7, h: 1.7, fontSize: 13.5, color: 'DCE3EA', paraSpaceAfter: 6 });
  T(s, '這不是終點，是通往夢想的最後一哩路。', { x: 0.5, y: 5.9, w: 5.7, h: 0.5, fontSize: 15, italic: true, color: 'C9D3DE' });
  // online video (YouTube) — plays in PowerPoint with internet
  s.addMedia({ type: 'online', link: 'https://www.youtube.com/embed/yXOe-3TqXQ8', x: 6.6, y: 1.35, w: 6.23, h: 3.5 });
  T(s, '開場影片（線上播放，簡報時需連網）', { x: 6.6, y: 4.95, w: 6.23, h: 0.3, fontSize: 10, color: 'A9B4C2' });
  card(s, 6.6, 5.4, 6.23, 1.45, { fill: '14263F', line: '2A4160', lineW: 1 });
  iconCircle(s, 6.8, 5.6, 0.5, 'mountain', GOLD, 'navy');
  T(s, '今天要講的事', { x: 7.45, y: 5.55, w: 5.2, h: 0.3, fontSize: 12, bold: true, color: GOLD });
  T(s, '這條路長什麼樣子 → 規則紅線 → 誰在哪台車、做什麼 → 每一站給什麼 → 夜間高山 → 住宿與回程', { x: 7.45, y: 5.88, w: 5.2, h: 0.9, fontSize: 11.5, color: 'E6ECF2' });
  s.addNotes('開場影片連結由選手提供：https://youtu.be/yXOe-3TqXQ8 。主題句「如果沒有天賦，那就一直重複」為選手指定。');
}

// 3. 六屆的路：2020 → 2026
{
  const s = pres.addSlide({ masterName: 'CONTENT', sectionTitle: section });
  s.addText('六屆的路：2020 → 2026', { placeholder: 'title' });
  const eds = [
    ['2020', 'I', '11.27 花蓮', '第一次下水，秀姑巒溪 → 合歡主峰', NAVY],
    ['2021', 'II', '11 月 花蓮', '再來一次', NAVY],
    ['2022', 'III', '11.25 花蓮', '第三次站上起點', NAVY],
    ['2023', 'IV', '11 月 花蓮', '太魯閣的霧與鐘聲', NAVY],
    ['2025', 'V', '11.14 日月潭', '新賽道・242 km', TEAL],
    ['2026', 'VI', '11.13 日月潭', '最後一哩路', RED],
  ];
  const w = 1.95, gap = 0.12, x0 = 0.5;
  eds.forEach((e, i) => {
    const x = x0 + i * (w + gap), y = 1.3, h = 4.55;
    card(s, x, y, w, h, { fill: WHITE, line: e[4], lineW: 1.5 });
    cardHeader(s, x, y, w, 0.9, e[4]);
    T(s, e[0], { x: x + 0.15, y: y + 0.08, w: w - 0.3, h: 0.5, fontSize: 22, bold: true, color: WHITE });
    T(s, 'FXT ' + e[1], { x: x + 0.15, y: y + 0.55, w: w - 0.3, h: 0.3, fontSize: 11, color: WHITE });
    const yp = crewPhoto((PHOTO_MAP.years || {})[e[0]]);
    if (yp) {
      s.addImage({ path: yp, x: x + 0.15, y: y + 1.05, w: w - 0.3, h: 2.2, sizing: { type: 'cover', w: w - 0.3, h: 2.2 } });
    } else {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.15, y: y + 1.05, w: w - 0.3, h: 2.2, rectRadius: 0.1, fill: { color: CARD }, line: { color: 'B8C2CE', width: 1, dashType: 'dash' } });
      s.addImage({ path: icon('camera', 'gray'), x: x + w / 2 - 0.25, y: y + 1.75, w: 0.5, h: 0.5 });
      T(s, '大頭的照片', { x: x + 0.15, y: y + 2.35, w: w - 0.3, h: 0.3, fontSize: 10, color: MUTED, align: 'center' });
    }
    T(s, e[2], { x: x + 0.15, y: y + 3.4, w: w - 0.3, h: 0.3, fontSize: 12, bold: true, color: INK });
    T(s, e[3], { x: x + 0.15, y: y + 3.72, w: w - 0.3, h: 0.7, fontSize: 11, color: HEX.dk2 });
  });
  card(s, 0.5, 6.0, 12.33, 0.9, { fill: CARD });
  iconCircle(s, 0.7, 6.2, 0.5, 'clock', MUTED);
  T(s, (PHOTO_MAP.years ? '2024 停辦一年。六年、五場、同一個目標——這條路的每一年都有人在旁邊。' : '2024 停辦一年。六年、五場、同一個目標。照片請放進 Drive「FXT2026 簡報素材」資料夾，標上年份，我會一張張放進對應的格子。'), { x: 1.35, y: 6.0, w: 11.3, h: 0.9, fontSize: 12.5, color: INK, valign: 'middle' });
  s.addNotes('歷屆日期來源：2020 FXT Race Manual（11/27 花蓮）、2022 選手手冊（11/25）、2025 選手手冊（11/14）、2026 選手手冊（11/13）；2021、2023 僅知 11 月。照片格待選手提供照片後替換。');
}

// 3b. 回憶牆（有照片才出現）
if (PHOTO_MAP.hero && crewPhoto(PHOTO_MAP.hero.file)) {
  const s = pres.addSlide({ masterName: 'BLANK', sectionTitle: section });
  photo(s, path.relative(PHOTOS, crewPhoto(PHOTO_MAP.hero.file)), 0, 0, W, H);
  rect(s, 0, 4.6, W, 2.9, '000000', { fill: { color: '000000', transparency: 35 }, line: { color: '000000', width: 0 } });
  T(s, PHOTO_MAP.hero.line || '如果沒有天賦，那就一直重複。', { x: 0.8, y: 4.9, w: 11.7, h: 1.0, fontSize: 36, bold: true, color: WHITE });
  T(s, PHOTO_MAP.hero.caption || '', { x: 0.8, y: 5.95, w: 11.7, h: 0.5, fontSize: 14, color: 'E6ECF2' });
}
if (PHOTO_MAP.wall && PHOTO_MAP.wall.length) {
  const s = slide('CONTENT', PHOTO_MAP.wallTitle || '六年，一直重複');
  const items = PHOTO_MAP.wall.filter(w => crewPhoto(w.file)).slice(0, 8);
  const cols = items.length <= 4 ? items.length : 4, rows = Math.ceil(items.length / cols);
  const gw = (12.33 - (cols - 1) * 0.15) / cols, gh = (5.55 - (rows - 1) * 0.15) / rows;
  items.forEach((it, i) => {
    const x = 0.5 + (i % cols) * (gw + 0.15), y = 1.3 + Math.floor(i / cols) * (gh + 0.15);
    s.addImage({ path: crewPhoto(it.file), x, y, w: gw, h: gh, sizing: { type: 'cover', w: gw, h: gh } });
    if (it.label) { rect(s, x, y + gh - 0.38, gw, 0.38, '000000', { fill: { color: '000000', transparency: 40 }, line: { color: '000000', width: 0 } }); T(s, it.label, { x: x + 0.1, y: y + gh - 0.38, w: gw - 0.2, h: 0.38, fontSize: 10.5, bold: true, color: WHITE, valign: 'middle' }); }
  });
}

// 4. 什麼是 FXT
{
  const s = slide('CONTENT', '這是一場什麼樣的比賽？');
  photo(s, 'hb-020-059.jpg', 0.5, 1.3, 5.6, 3.1);
  card(s, 0.5, 4.55, 5.6, 2.3, { fill: NAVY });
  T(s, '「勇敢，是台灣最美的精神，也是這場賽事的靈魂。」', { x: 0.75, y: 4.75, w: 5.1, h: 0.9, fontSize: 16, bold: true, color: WHITE, italic: true });
  T(s, '— 賽事 CEO Jovi Lo，2026 選手手冊歡迎信', { x: 0.75, y: 5.65, w: 5.1, h: 0.35, fontSize: 11, color: 'C9D3DE' });
  T(s, '#See you at the top.', { x: 0.75, y: 6.15, w: 5.1, h: 0.45, fontSize: 20, bold: true, color: GOLD });
  const cx = 6.45, cw = 6.4;
  const facts = [
    ['flag', RED, '第六屆・2020 年誕生的台灣極限鐵人', ['2020 首屆於花蓮起跑，2024 停辦一年，2025 以「日月潭→合歡山」全新賽道回歸', '2026.11.13（五）06:00 伊達邵碼頭下水，約 80 位選手']],
    ['star', GOLD, 'XTRI World Tour 世界巡迴賽成員', ['與挪威 Norseman 同一系列；本屆 4 個 Norseman 名額', '所有完賽者獲 200 X-Points']],
    ['users', TEAL, '自我補給賽：沒有大會補給站', ['補給團隊至少 2 人＋1 車，全程只能靠我們；歷屆完賽率約三到五成（媒體估）', '只有穿大會反光背心的「陪跑員＋補給員」可以接觸選手']],
    ['mountain', ORANGE, '點對點、四季一日', ['日月潭 748 m → 塔塔加 2,478 m → 合歡主峰 3,417 m', '湖泊、玉山山腳到高山夜間低溫，補給與保暖都要分段規劃']],
  ];
  facts.forEach((f, i) => {
    const y = 1.3 + i * 1.42;
    card(s, cx, y, cw, 1.3);
    iconCircle(s, cx + 0.18, y + 0.2, 0.5, f[0], f[1]);
    T(s, f[2], { x: cx + 0.82, y: y + 0.15, w: cw - 1.0, h: 0.4, fontSize: 15, bold: true, valign: 'middle' });
    T(s, bullets(f[3]), { x: cx + 0.82, y: y + 0.58, w: cw - 1.0, h: 0.7, fontSize: 12, color: HEX.dk2, paraSpaceAfter: 2 });
  });
  s.addNotes('資料來源：2026 FXT 選手手冊（歡迎信、Information of Supporter）、XTRI World Tour 歡迎信（2024 停辦、2026 回歸）。2020 首屆花蓮起跑：FXT 2020 Race Manual（花蓮靜浦報到、秀姑巒溪游泳）。2025 新賽道：2025 FXT 選手手冊（日月潭→合歡主峰）。「第六屆」為團隊說法（2020–2023、2025、2026）。參賽人數約 80 人取自 2026/11/13 交通管制新聞。');
}

// ===== Section 1 =====
section = '一、認識賽事與規則';
pres.addSection({ title: section });
divider('PART 1', '認識賽事與規則', '路線、關門時間、官方規定——補給團隊的共同語言', 'hb-008-022.jpg');

// 4. 路線總覽
{
  const s = slide('CONTENT', '全程路線總覽：從湖心到雲端');
  s.addImage({ path: path.join(ASSETS, 'course_map.png'), x: 0.5, y: 1.25, w: 6.6, h: 5.5 });
  const segs = [
    ['swim', BLUE, '游泳 3.8 km', '日月潭伊達邵碼頭・水溫約 20–24°C', ['06:00 出發，07:40 關門（1 小時 40 分）', '只有穿背心、戴手環的陪跑員可進 T1']],
    ['bike', TEAL, '自行車 196.7 km', '日月潭 → 塔塔加夫妻樹折返 → 日月潭 → 埔里 → T2', ['748 m 爬到 2,478 m，總爬升 3,600 m 以上', '夫妻樹 86 km 12:00 關門；T2 15:50 關門']],
    ['run', ORANGE, '跑步 42 km', 'T2 → 霧社 → 台14甲 西進武嶺 → 合歡主峰步道', ['T2 約 714 m（手冊寫埔里 450 m）爬升到 3,417 m，爬升 3,000 m 以上', '翠峰 26 km 開放陪跑，鳶峰 33 km 強制陪跑']],
  ];
  segs.forEach((g, i) => {
    const y = 1.25 + i * 1.85, x = 7.4, w = 5.45, h = 1.7;
    card(s, x, y, w, h);
    iconCircle(s, x + 0.18, y + 0.2, 0.55, g[0], g[1]);
    T(s, g[2], { x: x + 0.9, y: y + 0.14, w: w - 1.05, h: 0.4, fontSize: 17, bold: true, color: g[1], valign: 'middle' });
    T(s, g[3], { x: x + 0.9, y: y + 0.54, w: w - 1.05, h: 0.32, fontSize: 11, color: MUTED });
    T(s, bullets(g[4]), { x: x + 0.9, y: y + 0.9, w: w - 1.05, h: 0.75, fontSize: 12, color: HEX.dk2, paraSpaceAfter: 2 });
  });
  T(s, '總距離約 242 km・總爬升 6,000 m 以上。地圖：自行車為 2025 大會 GPX 實際軌跡，跑步段為示意。', { x: 0.5, y: 6.78, w: 12.3, h: 0.28, fontSize: 10, color: MUTED });
}

// 5. 高度剖面
{
  const s = slide('CONTENT', '高度剖面：補給車該在哪裡、不能在哪裡');
  s.addImage({ path: path.join(ASSETS, 'course_profile.png'), x: 0.5, y: 1.2, w: 12.33, h: 4.74 });
  const notes = [
    ['ban', RED, '補給車禁行區', '草坪頭管制哨 → 塔塔加 → 草坪頭（約 56–116 km）。所有補給、裝備交接在管制哨「之前」完成；不得在管制哨停車或補給。'],
    ['clock', GOLD, '兩個自行車關門點', '夫妻樹 86 km 12:00、T2 15:50。11:30 起補給車可依工作人員指示上塔塔加接回被關門的選手。'],
    ['run', ORANGE, '跑步段四個關門點', '清境 18:50 → 翠峰 20:20（開放陪跑）→ 鳶峰 21:40（強制陪跑）→ 登山口 23:00，之後步道登頂。'],
  ];
  notes.forEach((n, i) => {
    const x = 0.5 + i * 4.15, y = 6.0, w = 4.0, h = 0.98;
    card(s, x, y, w, h);
    iconCircle(s, x + 0.12, y + 0.12, 0.42, n[0], n[1]);
    T(s, n[2], { x: x + 0.64, y: y + 0.1, w: w - 0.75, h: 0.3, fontSize: 12, bold: true, color: n[1] });
    T(s, n[3], { x: x + 0.64, y: y + 0.4, w: w - 0.75, h: 0.56, fontSize: 9.5, color: HEX.dk2 });
  });
}

// 6. 關門時間表
{
  const s = slide('CONTENT', '官方關門時間表（2026 選手手冊）');
  const rows = [
    ['時間', '賽段／檢查點', '地點', '陪跑規定', '補給團隊要做的事'],
    ['05:00–05:45', 'T1 轉換區開放', '伊達邵碼頭', '—', '陪跑員穿背心＋手環進 T1，協助穿防寒衣、檢查車燈與 GPS'],
    ['06:00', '游泳出發', '伊達邵碼頭', '—', '上岸補給就位（威德、蜂蜜水＋Race On＋冰塊）'],
    ['07:40', '游泳關門（T1）', '伊達邵碼頭', '—', '選手離開後收拾 T1 裝備、垃圾；主車出發往信義'],
    ['12:00', '自行車 86 km 關門', '塔塔加夫妻樹', '—', '黑色特殊需求袋在此；補給車不得進入禁行區，11:30 後可上山接人'],
    ['15:50', '自行車關門（T2）', 'T2：全家仁愛賽德克店停車場（2025 手冊）', '—', '跑步強制裝備＋醫療檢測 ≥2 分鐘，陪跑員不得觸碰選手'],
    ['18:50', '跑步 17.5 km 關門', '清境高空觀景步道', '—', '脫鐵服上半身換 T 恤、補給葡萄；天已黑，確認頭燈、背包掛燈、反光背心'],
    ['20:20', '跑步 26 km 關門', '翠峰停車場', { text: '開放陪跑', options: { bold: true, color: GREEN } }, '陪跑員就位；換衣斗篷→脫鐵服下半身換跑褲；陪跑員裝備自檢'],
    ['21:40', '跑步 33 km 關門', '鳶峰停車場', { text: '強制陪跑', options: { bold: true, color: RED } }, '裝備檢查；套迪卡儂長褲＋外套＋手套、吃熱食，開始走路'],
    ['23:00', '跑步 40 km 關門', '合歡主峰步道登山口', '陪跑員同行', '登頂步道約 1.8 km、爬升約 230 m；接駁車只到昆陽停車場，完賽後通知待命人員'],
  ];
  tbl(s, rows, { x: 0.5, y: 1.25, w: 12.33, colW: [1.25, 2.1, 2.1, 1.3, 5.58], fontSize: 11, headFill: NAVY, rowH: 0.47 });
  T(s, '賽前 11/12：13:00–15:00 報到（日月潭力麗溫德姆）・15:30 中文說明會・17:30 祈福典禮（日月潭青年活動中心）。賽後 11/14 09:00：頒獎／Party（清境愛立家民宿）。', { x: 0.5, y: 6.1, w: 12.33, h: 0.5, fontSize: 11, color: HEX.dk2 });
  T(s, '手冊提醒：大會可依天候與安全評估調整路線與關門時間，一切以賽前說明會公布為準。', { x: 0.5, y: 6.6, w: 12.33, h: 0.35, fontSize: 10.5, bold: true, color: RED });
  s.addNotes('全部時間與地點與 2026 FXT 選手手冊 Schedule 一致，並與原簡報「賽程與關門時間」相符。T2 在 2026 手冊僅提供 Google 地圖連結（https://maps.app.goo.gl/q7BjotezpH8ctXj76）；2025 手冊寫「全家仁愛賽德克店停車場」，原簡報寫「全家仁愛霧社店」，請賽前確認。');
}

// 7. 比賽日時間軸
{
  const s = slide('CONTENT', '比賽日一眼看完：05:00 → 隔日 01:00');
  const x0 = 0.7, x1 = 12.65, t0 = 3, t1 = 26; // hours (26 = 02:00 next day)
  const X = h => x0 + ((h - t0) / (t1 - t0)) * (x1 - x0);
  const by = 3.15, bh = 0.5;
  // night / day band
  rect(s, X(3), by, X(6.17) - X(3), bh, NAVY);
  rect(s, X(6.17), by, X(17.17) - X(6.17), bh, 'FFF1C2');
  rect(s, X(17.17), by, X(26) - X(17.17), bh, NAVY);
  s.addImage({ path: icon('sun', 'gold'), x: X(6.17) + 0.1, y: by + 0.1, w: 0.3, h: 0.3 });
  T(s, '日出 06:10', { x: X(6.17) + 0.45, y: by, w: 1.3, h: bh, fontSize: 10, bold: true, color: NAVY, valign: 'middle' });
  s.addImage({ path: icon('moon', 'white'), x: X(17.17) + 0.1, y: by + 0.1, w: 0.3, h: 0.3 });
  T(s, '日落 17:10', { x: X(17.17) + 0.45, y: by, w: 1.3, h: bh, fontSize: 10, bold: true, color: WHITE, valign: 'middle' });
  // segments bar
  const sy = by + bh + 0.12, sh = 0.42;
  const seg = (a, b, color, label) => { rect(s, X(a), sy, X(b) - X(a), sh, color); T(s, label, { x: X(a), y: sy, w: X(b) - X(a), h: sh, fontSize: 10, bold: true, color: WHITE, align: 'center', valign: 'middle' }); };
  seg(6, 7.67, BLUE, '游泳');
  seg(7.67, 15.83, TEAL, '自行車 196.7 km');
  seg(15.83, 23, ORANGE, '跑步 40 km 公路');
  seg(23, 24.5, GOLD, '登頂');
  rect(s, X(24.5), sy, X(26) - X(24.5), sh, 'D9DEE5'); T(s, '手機待命至 01:00', { x: X(24.5), y: sy, w: X(26) - X(24.5), h: sh, fontSize: 9, bold: true, color: HEX.dk2, align: 'center', valign: 'middle' });
  // hour ticks
  for (let h = 4; h <= 26; h += 2) {
    const lab = `${String(h % 24).padStart(2, '0')}:00`;
    rect(s, X(h) - 0.005, sy + sh, 0.01, 0.12, MUTED);
    T(s, lab, { x: X(h) - 0.4, y: sy + sh + 0.14, w: 0.8, h: 0.25, fontSize: 9, color: MUTED, align: 'center' });
  }
  // events: above (cut-offs) and below (crew actions)
  const above = [
    [5, 'T1 開放', '05:00'], [6, '出發', '06:00'], [7.67, '游泳關門', '07:40'], [12, '夫妻樹關門', '12:00'], [15.83, 'T2 關門', '15:50'],
    [18.83, '清境關門', '18:50'], [20.33, '翠峰 開放陪跑', '20:20'], [21.67, '鳶峰 強制陪跑', '21:40'], [23, '登山口關門', '23:00'],
  ];
  above.forEach((e, i) => {
    const lvl = i % 2; const ty = by - 0.95 - lvl * 0.75;
    rect(s, X(e[0]) - 0.008, ty + 0.55, 0.016, by - (ty + 0.55), 'B8C2CE');
    s.addShape(pres.shapes.OVAL, { x: X(e[0]) - 0.07, y: by - 0.07, w: 0.14, h: 0.14, fill: { color: RED }, line: { color: WHITE, width: 1 } });
    T(s, [{ text: e[2], options: { bold: true, color: RED, breakLine: true } }, { text: e[1], options: { color: INK } }], { x: X(e[0]) - 0.75, y: ty, w: 1.5, h: 0.55, fontSize: 10, align: 'center', valign: 'bottom' });
  });
  const below = [
    [4.1, '起床 03:30\n泡麵＋吐司', MUTED], [7.9, '主車出發\n往信義 7-11', TEAL], [9.8, '草坪頭前完成補給\n（紅袋 56 km）', RED], [11.5, '11:30 起可上\n塔塔加接人', GOLD],
    [13.5, '和社→豐丘\n→水里坡→埔里高工', TEAL], [16.3, 'T2 檢錄\n換裝 ≥2 分鐘', ORANGE], [19.5, '翠峰：斗篷換跑褲\n陪跑員上場', GREEN], [22.3, '昆陽停車場\n等候、保暖', NAVY], [24.2, '接駁：登山口\n→ 昆陽', GOLD],
  ];
  below.forEach((e, i) => {
    const lvl = i % 2; const ty = sy + sh + 0.5 + lvl * 0.72;
    rect(s, X(e[0]) - 0.008, sy + sh + 0.42, 0.016, ty - (sy + sh + 0.42), 'D9DEE5');
    card(s, X(e[0]) - 0.78, ty, 1.56, 0.62, { fill: WHITE, line: e[2], lineW: 1 });
    T(s, e[1], { x: X(e[0]) - 0.78, y: ty, w: 1.56, h: 0.62, fontSize: 9, color: INK, align: 'center', valign: 'middle', bold: true });
  });
  T(s, '上方：官方關門點。下方：補給團隊關鍵動作（建議，依車長指示調整）。日出日落為 11/13 南投山區概略值。', { x: 0.5, y: 6.78, w: 12.3, h: 0.28, fontSize: 9.5, color: MUTED });
}

// 8. 賽前一天 11/12
{
  const s = slide('CONTENT', '賽前一天 11/12（四）：報到、說明會、祈福');
  const steps = [
    ['id', BLUE, '13:00–15:00', '報到／檢錄／健康檢測', '日月潭力麗溫德姆溫泉酒店', ['選手「必須與陪跑員一同報到」，大會只認定當天報到的陪跑員並發手環', '帶照片證件（身分證／健保卡／護照）、簽切結書、醫護檢查、裝備＋車檢', '建議備妥意外與責任保險證明放在補給車']],
    ['megaphone', TEAL, '15:30–16:15', '賽前說明會（中文）', '力麗溫德姆（16:20 英文場）', ['強制出席；路線／規則若有變更以說明會為準', '車長與陪跑員務必到場，當晚回報全隊']],
    ['star', GOLD, '17:30–18:00', '開賽祈福典禮', '日月潭青年活動中心', ['建議：典禮後 18:30 團隊賽前會議（分工、通訊、特殊需求袋打包）']],
  ];
  steps.forEach((st, i) => {
    const y = 1.3 + i * 1.78, x = 0.5, w = 7.3, h = 1.62;
    card(s, x, y, w, h);
    iconCircle(s, x + 0.18, y + 0.2, 0.55, st[0], st[1]);
    T(s, st[2], { x: x + 0.9, y: y + 0.14, w: 1.6, h: 0.4, fontSize: 15, bold: true, color: st[1], valign: 'middle' });
    T(s, st[3], { x: x + 2.5, y: y + 0.14, w: 3.0, h: 0.4, fontSize: 15, bold: true, valign: 'middle' });
    T(s, st[4], { x: x + 0.9, y: y + 0.52, w: w - 1.0, h: 0.3, fontSize: 10.5, color: MUTED });
    T(s, bullets(st[5]), { x: x + 0.9, y: y + 0.84, w: w - 1.05, h: h - 0.9, fontSize: 11, color: HEX.dk2, paraSpaceAfter: 2 });
  });
  // right: material bag checklist
  const x = 8.1, w = 4.73;
  card(s, x, 1.3, w, 5.18, { fill: NAVY });
  iconCircle(s, x + 0.2, 1.5, 0.5, 'bag', GOLD, 'navy');
  T(s, '報到領取物資袋，當場清點 9 項', { x: x + 0.85, y: 1.5, w: w - 1.0, h: 0.5, fontSize: 15, bold: true, color: WHITE, valign: 'middle' });
  const items = ['選手手環／陪跑員手環（報到當天就戴上）', '選手號碼布', '自行車車貼＆安全帽貼', '大會泳帽', '紋身貼紙', '補給車車貼（貼後擋風玻璃左上角）', '反光背心 3 件：選手 1＋陪跑員／補給員 2', 'GPS 追蹤器（說明會教安裝於車上）', '特殊需求袋（黑袋 塔塔加／紅袋 草坪頭）'];
  items.forEach((it, i) => {
    const yy = 2.2 + i * 0.44;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: yy + 0.06, w: 0.26, h: 0.26, rectRadius: 0.05, fill: { color: '1E3A5F' }, line: { color: GOLD, width: 1 } });
    T(s, it, { x: x + 0.65, y: yy, w: w - 0.85, h: 0.4, fontSize: 11.5, color: 'E6ECF2', valign: 'middle' });
  });
  T(s, '缺任何一項立即向主辦單位反映。', { x: x + 0.25, y: 6.08, w: w - 0.5, h: 0.35, fontSize: 11, bold: true, color: GOLD });
  s.addNotes('來源：2026 FXT 選手手冊 Schedule、Registration Information、報到＆檢錄須知。');
}

// 9. 規則紅線
{
  const s = slide('CONTENT', '補給團隊規則紅線：違規是罰「選手」的時間');
  const rules = [
    ['vest', RED, '只有穿背心的人能碰選手', ['陪跑員（背心＋手環）與 1 位指定補給員（背心）才能接觸選手、遞補給', '其他親友、加油團、隨車人員一律不得碰觸或遞任何東西']],
    ['car', NAVY, '每位選手只有 1 輛官方補給車', ['車貼貼在後擋風玻璃左上角；所有補給只能由這輛車進行', '親友、加油團車輛不得進入賽道補給或跟隨選手']],
    ['parking', TEAL, '靜態補給四要件', ['停在白線內、車停道路外、人在車外、在「車頭」交接', '不可從車窗遞、不可行進中補給；白線外補給警告超過 2 次取消資格']],
    ['ban', ORANGE, '不跟車、不配速、不尾隨', ['補給車、攝影車、加油團都不得跟在選手前後（Non-drafting）', '陪跑員不得騎自行車或機車跟隨；補給完就離開定點']],
    ['route', GOLD, '草坪頭 → 塔塔加禁行', ['補給與裝備交接在草坪頭管制哨之前完成；管制哨不得停車或補給', '11:30 起可依工作人員指示上山接回被關門的選手']],
    ['warn', '8B1E2D', '罰則', ['輕微違規 5 分鐘；第二次輕微或第一次重大 15 分鐘；第三次直接 DSQ', '重大違規：闖紅燈、亂丟垃圾；搭車或抄捷徑直接取消資格']],
  ];
  rules.forEach((r, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 4.15, y = 1.3 + row * 2.45, w = 4.0, h = 2.3;
    infoCard(s, x, y, w, h, { iconName: r[0], color: r[1], title: r[2], lines: r[3], fontSize: 12.5, titleSize: 14.5 });
  });
  card(s, 0.5, 6.2, 12.33, 0.72, { fill: NAVY });
  iconCircle(s, 0.7, 6.31, 0.5, 'phone', GOLD, 'navy');
  T(s, '不確定能不能做的時候：先不要碰選手、先不要停車，打給車長。罰時在 T1、T2 執行，跑步段違規直接加在總成績。', { x: 1.35, y: 6.2, w: 11.3, h: 0.72, fontSize: 13, bold: true, color: WHITE, valign: 'middle' });
  s.addNotes('來源：2026 FXT 選手手冊「陪跑員相關資訊」「補給車輛」「補給車與加油團規定」「塔塔加折返點關門與補給車規定」「詳細處罰規則」。');
}

// 10. 特殊需求袋 & 塔塔加
{
  const s = slide('CONTENT', '特殊需求袋與塔塔加折返：禁行區裡唯一的補給方式');
  // left: two bags
  const bag = (x, y, color, title, sub, items, foot) => {
    card(s, x, y, 3.95, 5.6, { fill: WHITE, line: color, lineW: 1.5 });
    cardHeader(s, x, y, 3.95, 0.95, color);
    iconCircle(s, x + 0.2, y + 0.2, 0.55, 'bag', WHITE, color === RED ? 'red' : 'navy');
    T(s, title, { x: x + 0.9, y: y + 0.12, w: 3.0, h: 0.42, fontSize: 16, bold: true, color: WHITE, valign: 'middle' });
    T(s, sub, { x: x + 0.9, y: y + 0.54, w: 3.0, h: 0.3, fontSize: 10.5, color: 'FFFFFF' });
    T(s, bullets(items), { x: x + 0.25, y: y + 1.15, w: 3.5, h: 3.4, fontSize: 12.5, color: INK, paraSpaceAfter: 5 });
    T(s, foot, { x: x + 0.25, y: y + 4.75, w: 3.5, h: 0.75, fontSize: 10, color: MUTED, italic: true, valign: 'bottom' });
  };
  bag(0.5, 1.3, RED, '紅袋・草坪頭管制哨', '56 km・上山途中・禁行區起點', ['建議放：原「草坪頭」補給品——梅子可樂＋Race On（黃壺）2、可頌 1', '補給車不得在管制哨補給或停車，所以要靠紅袋', '冰塊會融化：改用保冷杯／冷凍過的飲料'], '原表草坪頭補給品由車上遞交，與 2026 規定衝突，建議改為紅袋或在管制哨前車頭補給（待確認）');
  bag(4.65, 1.3, INK, '黑袋・塔塔加夫妻樹', '86 km 折返點・12:00 關門', ['葡萄汁＋冰塊（冰霸杯）', '瓶裝礦泉水 1', '吐司 2–3 片', '蜂蜜水＋Race On 1（白壺）', '水 1（白壺）', '建議加：防風外套＋手套（下滑 30 km 很冷）'], '2025 車隊清單（FXT_050 補給清單）的夫妻樹特殊補給袋：防風外套＋手套、保溫瓶溫水、乳清蜂蜜、鹽膠囊、果膠');
  // right: rules
  const x = 8.8, w = 4.03;
  const rr = [
    ['clock', GOLD, '何時交？', '比賽當天早上出發前繳交給大會；大會送達指定地點，不負遺失損壞責任'],
    ['box', TEAL, '能放什麼？', '只能放個人補給或裝備，禁止危險、易碎物品'],
    ['home', NAVY, '怎麼拿回？', '賽後隔日（11/14）送到清境愛立家民宿領取'],
    ['ban', RED, '禁行區提醒', '草坪頭管制哨 ⇄ 塔塔加全段補給車禁行；11:30 起可依指示上山接關門選手，低速、不干擾比賽'],
  ];
  rr.forEach((r, i) => {
    const y = 1.3 + i * 1.08;
    card(s, x, y, w, 0.98);
    iconCircle(s, x + 0.14, y + 0.14, 0.42, r[0], r[1]);
    T(s, r[2], { x: x + 0.66, y: y + 0.1, w: w - 0.8, h: 0.3, fontSize: 12, bold: true, color: r[1] });
    T(s, r[3], { x: x + 0.66, y: y + 0.4, w: w - 0.8, h: 0.56, fontSize: 10, color: HEX.dk2 });
  });
  if (PHOTO_MAP.supply && crewPhoto(PHOTO_MAP.supply.car)) s.addImage({ path: crewPhoto(PHOTO_MAP.supply.car), x, y: 5.7, w, h: 1.2, sizing: { type: 'cover', w, h: 1.2 } }); else photo(s, 'hb-012-034.jpg', x, 5.7, w, 1.2);
  s.addNotes('來源：2026 FXT 選手手冊「特殊需求袋」「塔塔加折返點關門與補給車規定」；袋內品項來自原簡報「6. 夫妻樹特殊需求袋」。');
}

// ===== Section 2 =====
section = '二、車隊編組與角色';
pres.addSection({ title: section });
divider('PART 2', '車隊編組與角色', '三台補給車加一台親友專車、十二位隊友——誰在哪裡、做什麼', 'hb-002-002.jpg');

// 12. 車隊編組
{
  const s = slide('CONTENT', '車隊編組：四台車的任務');
  const cars = [
    ['主補給車・RAV4', RED, 'car', ['大顆頭（駕駛）', '小麥', '金剛', '聖恩'], '官方補給車・貼車貼', ['依計劃進行補給，隨時注意選手狀況', '所有接觸選手的補給由穿背心者執行', 'T1 收裝備 → 信義 → 草坪頭前 → 豐丘 → 水里坡 → 埔里高工 → T2 → 跑步段各站']],
    ['第二補給車・GLC 250', TEAL, 'van', ['傑哥（駕駛）', '佳惠', '珮淳', '攝影師'], '支援・聽主車指令', ['接受主補給車指令並確實執行', '和社 7-11：冰葡萄汁＋吐司放定點自取（不接觸、不聊天），收選手脫下的外套', '攝影師隨車定點拍攝，定點完成即離開']],
    ['攝影器材車・CRV', NAVY, 'camera', ['小朱（駕駛）', '阿徹'], '影像紀錄', ['配合攝影師拍攝需求移動器材與人員', '定點拍攝，不得尾隨或擋到選手與用路人', '夜間段在翠峰／鳶峰／昆陽安全處待命']],
    ['親友專車・MODEL Y', GOLD, 'megaphone', ['俊儀（駕駛）'], '11/13 中午南下・直奔 T2', ['在 T2 見證歷史，之後於安全地點加油', '不進賽道、不攔車拍照、不遞補給', '可支援昆陽停車場保暖物資與接人']],
  ];
  cars.forEach((c, i) => {
    const x = 0.5 + i * 3.12, y = 1.3, w = 2.97, h = 4.1;
    card(s, x, y, w, h, { fill: WHITE, line: c[1], lineW: 1.5 });
    cardHeader(s, x, y, w, 1.0, c[1]);
    iconCircle(s, x + 0.18, y + 0.2, 0.6, c[2], WHITE, c[1] === GOLD ? 'gold' : (c[1] === TEAL ? 'teal' : (c[1] === NAVY ? 'navy' : 'red')));
    T(s, c[0], { x: x + 0.9, y: y + 0.12, w: w - 1.0, h: 0.5, fontSize: 14, bold: true, color: WHITE, valign: 'middle' });
    T(s, c[4], { x: x + 0.9, y: y + 0.6, w: w - 1.0, h: 0.3, fontSize: 10, color: WHITE });
    c[3].forEach((n, j) => pill(s, x + 0.2 + (j % 2) * 1.32, y + 1.2 + Math.floor(j / 2) * 0.42, 1.22, 0.34, n, CARD, INK, 12));
    T(s, bullets(c[5]), { x: x + 0.2, y: y + 2.15, w: w - 0.4, h: h - 2.3, fontSize: 12, color: HEX.dk2, paraSpaceAfter: 5 });
  });
  card(s, 0.5, 5.6, 12.33, 0.62, { fill: CARD });
  T(s, '兩件背心', { x: 0.7, y: 5.6, w: 2.6, h: 0.62, fontSize: 13, bold: true, color: INK, valign: 'middle' });
  pill(s, 3.4, 5.74, 1.3, 0.34, '陪跑員（手環）', GREEN, WHITE, 10.5);
  T(s, '小麥（太太，11/12 與大頭一同報到）', { x: 4.8, y: 5.6, w: 2.0, h: 0.62, fontSize: 12, bold: true, color: INK, valign: 'middle' });
  pill(s, 6.9, 5.74, 1.3, 0.34, '補給手', ORANGE, WHITE, 10.5);
  T(s, '金剛', { x: 8.3, y: 5.6, w: 1.8, h: 0.62, fontSize: 14, bold: true, color: INK, valign: 'middle' });
  T(s, '其餘成員不碰選手、不遞補給', { x: 10.1, y: 5.6, w: 2.6, h: 0.62, fontSize: 11, bold: true, color: RED, valign: 'middle' });
  card(s, 0.5, 6.35, 12.33, 0.55, { fill: 'FFF4F4', line: RED, lineW: 1 });
  T(s, '手冊規定「每位選手僅限一輛官方認定補給車，所有補給只能由該車進行」。第二車在和社的補給方式請在會議中確認並向大會釐清。', { x: 0.65, y: 6.35, w: 12.1, h: 0.55, fontSize: 11, bold: true, color: RED, valign: 'middle' });
  s.addNotes('成員分配沿用原簡報「補給車成員與分工」。任務描述依原簡報各補給點備註與 2026 手冊規則整理，供會議討論。');
}

// 12b. 我們的人
{
  const s = slide('CONTENT', '我們的人：十二位隊友、一位選手');
  const people = [
    ['大頭', '選手・022', INK], ['小麥', '太太兼陪跑員（背心＋手環）・RAV4', GREEN], ['大顆頭', 'RAV4 駕駛・車長', RED], ['金剛', '補給手（背心）・RAV4', ORANGE], ['聖恩', 'RAV4・補給支援', RED],
    ['傑哥', 'GLC 250 駕駛・車長', TEAL], ['佳惠', 'GLC 250・支援', TEAL], ['珮淳', 'GLC 250・支援', TEAL], ['攝影師', 'GLC 250・影像紀錄', TEAL],
    ['小朱', 'CRV 駕駛・器材', NAVY], ['阿徹', 'CRV・攝影組', NAVY], ['俊儀', 'MODEL Y 駕駛・親友團', GOLD],
  ];
  const cols = 7, tw = 1.72, gapx = 0.047, th = 2.6;
  people.forEach((pp, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const rowCount = row === 0 ? Math.min(cols, people.length) : people.length - cols;
    const rowW = rowCount * tw + (rowCount - 1) * gapx;
    const x = 0.5 + (12.33 - rowW) / 2 + col * (tw + gapx), y = 1.3 + row * (th + 0.2);
    card(s, x, y, tw, th, { fill: CARD });
    const d = 1.15, cx = x + (tw - d) / 2, cy = y + 0.2;
    const pf = crewPhoto((PHOTO_MAP.crew || {})[pp[0]]);
    if (pf) {
      s.addImage({ path: pf, x: cx, y: cy, w: d, h: d, sizing: { type: 'cover', w: d, h: d }, rounding: true });
    } else {
      s.addShape(pres.shapes.OVAL, { x: cx, y: cy, w: d, h: d, fill: { color: pp[2] }, line: { color: WHITE, width: 2 } });
      T(s, pp[0].slice(0, 1), { x: cx, y: cy, w: d, h: d, fontSize: 30, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    }
    T(s, pp[0], { x: x + 0.08, y: y + 1.45, w: tw - 0.16, h: 0.4, fontSize: 15, bold: true, align: 'center', valign: 'middle' });
    T(s, pp[1], { x: x + 0.1, y: y + 1.85, w: tw - 0.2, h: 0.65, fontSize: 10, color: HEX.dk2, align: 'center' });
  });
  T(s, '頭像來自各自提供的照片；攝影師為示意圖像。', { x: 0.5, y: 6.82, w: 12.3, h: 0.26, fontSize: 9.5, color: MUTED });
}

// 13. 角色與職責
{
  const s = slide('CONTENT', '角色與職責：每個人上車前都知道自己的三件事');
  const roles = [
    ['crown', RED, '車長／駕駛', '每車 1 人', ['遵守交規與速限，停車一律停道路外、白線內', '決定何時離開定點、下一站路線與停靠點', '與其他車保持聯繫，回報選手通過時間']],
    ['vest', ORANGE, '補給手（背心）', '金剛', ['唯一能遞補給、碰選手的補給員', '選手到站前備妥品項，在車頭交接，動作要快', '記錄吃了什麼、喝多少，交棒給陪跑員']],
    ['run', GREEN, '陪跑員', '小麥（太太）', ['穿背心＋戴手環；與選手一同報到；翠峰 26 km 可陪，鳶峰 33 km 起強制', '自備強制裝備：手機、頭燈、急救毯、水袋背包、背包掛燈、急救包', '掌握配速與關門時間，適時提醒選手']],
    ['timer', GOLD, '計時／導航官', '每車 1 人', ['對照關門表與低標時間，提前 15 分鐘預警', '追蹤 GPS 位置，預估到站時間給下一車', '記錄補給與通過時間（賽後檢討）']],
    ['camera', NAVY, '攝影師＋器材車', 'GLC／CRV', ['攝影師隨第二車、器材與阿徹在 CRV；定點拍攝，不跟車、不站在賽道上', '優先畫面：下水、信義／草坪頭前、T2、翠峰夜跑、登頂（夫妻樹為禁行區）', '隨時能接聽電話']],
    ['tools', TEAL, '後勤／物資', '每車 1 人', ['保冰桶、冰塊、熱水、行動電源、垃圾袋', 'T1 收裝備、T2 收腳踏車，特殊需求袋早上繳交', '夜間保暖物資：外套、毛毯、熱飲']],
  ];
  roles.forEach((r, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 4.15, y = 1.3 + row * 2.45, w = 4.0, h = 2.3;
    card(s, x, y, w, h);
    iconCircle(s, x + 0.18, y + 0.18, 0.5, r[0], r[1]);
    T(s, r[2], { x: x + 0.8, y: y + 0.14, w: w - 2.0, h: 0.5, fontSize: 15, bold: true, valign: 'middle' });
    pill(s, x + w - 1.25, y + 0.24, 1.05, 0.3, r[3], r[1], WHITE, 9.5);
    T(s, bullets(r[4]), { x: x + 0.2, y: y + 0.82, w: w - 0.4, h: h - 0.95, fontSize: 12, color: HEX.dk2, paraSpaceAfter: 4 });
  });
  card(s, 0.5, 6.25, 12.33, 0.65, { fill: NAVY });
  iconCircle(s, 0.7, 6.33, 0.5, 'vest', GOLD, 'navy');
  T(s, '背心只有 2 件：補給手金剛、陪跑員小麥（手環，11/12 與大頭一同報到）。其餘人員不碰選手、不遞補給。', { x: 1.35, y: 6.25, w: 11.3, h: 0.65, fontSize: 12.5, bold: true, color: WHITE, valign: 'middle' });
}

// 14. 車輛任務板
{
  const s = slide('CONTENT', '比賽日車輛任務板（會議中填寫、賽前定稿）');
  const M = '■', B = '□';
  const rows = [
    ['時間（參考）', '站點', 'RAV4 主車', 'GLC 第二車', 'CRV 器材車', 'MODEL Y', '補給手 金剛／陪跑員 小麥'],
    ['05:00–07:40', 'T1 伊達邵碼頭', M, B, M, '—', '陪跑員進 T1'],
    ['約 08:30', '4 信義 7-11（約 33 km）', M, B, B, '—', '補給手'],
    ['約 09:40', '5 草坪頭管制哨前（56 km）', M, B, B, '—', '補給手／紅袋'],
    ['約 11:30', '6 塔塔加夫妻樹（86 km）', '禁行・黑袋', '禁行', '禁行', '禁行', '—'],
    ['約 12:30', '7 和社 7-11（約 125 km）', B, '■ 定點自取', B, '—', '（不接觸）'],
    ['約 13:00', '8 豐丘葡萄冰沙（約 142 km）', M, B, B, '—', '補給手'],
    ['約 13:20', '9 水里坡（約 149 km）', M, B, B, '—', '補給手（機動）'],
    ['約 14:40', '10 埔里高工（約 185 km）', M, B, B, '南下中', '補給手'],
    ['≤15:50', 'T2 全家仁愛賽德克店停車場', M, M, M, '■ 見證歷史', '補給手（檢錄時不碰）'],
    ['≤18:50', '清境高空步道（17.5 km）', M, B, B, B, '補給手：T 恤＋葡萄'],
    ['≤20:20', '翠峰停車場（26 km）', M, B, M, B, '陪跑員上場・斗篷換跑褲'],
    ['≤21:40', '鳶峰停車場（33 km）', M, B, B, B, '陪跑員（強制）・長褲外套手套'],
    ['≤23:00', '昆陽停車場／登山口（40 km）', M, M, M, M, '陪跑員登頂'],
  ];
  tbl(s, rows, { x: 0.5, y: 1.25, w: 12.33, colW: [1.25, 2.9, 1.35, 1.1, 1.4, 1.4, 2.93], fontSize: 10.5, headFill: NAVY, rowH: 0.37, align: [null, null, 'center', 'center', 'center', 'center', null] });
  T(s, '■ 已安排　□ 待分配　— 不在場　「時間（參考）」為依關門時間回推的保守估計，請依選手目標配速與 GPS 實況調整。地標里程（約）依 2025 大會 GPX 推估。第二車一律聽主車指令。', { x: 0.5, y: 6.5, w: 12.3, h: 0.5, fontSize: 10, color: HEX.dk2 });
}

// ===== Section 3 =====
section = '三、賽段執行手冊';
pres.addSection({ title: section });
divider('PART 3', '賽段執行手冊', '游泳、自行車、跑步——每一站給什麼、怎麼給、誰來給', 'hb-022-065.jpg');

// 16. 游泳與 T1
{
  const s = slide('CONTENT', '游泳 3.8 km 與 T1：清晨的第一場補給');
  photo(s, 'hb-007-019.jpg', 0.5, 1.3, 5.2, 3.0);
  card(s, 0.5, 4.45, 5.2, 2.35, { fill: NAVY });
  T(s, 'T1 規則（2026 手冊）', { x: 0.7, y: 4.6, w: 4.8, h: 0.35, fontSize: 13, bold: true, color: GOLD });
  T(s, bullets(['只有選手與「穿背心＋戴手環」的陪跑員能進 T1', '陪跑員可協助穿脫防寒衣、上岸後給保暖物與補給', '選手離開後，陪跑員收拾裝備、個人物品與垃圾', '進 T1 要備齊：坐管貼、安全帽＋帽貼、號碼布、GPS（裝車上）、前後燈（裝車上）、紋身貼、泳帽']), { x: 0.7, y: 4.98, w: 4.85, h: 1.85, fontSize: 11, color: 'E6ECF2', paraSpaceAfter: 2 });
  const x = 6.0, w = 6.83;
  const steps = [
    ['bed', MUTED, '約 03:30 起床（2025 車隊時間）', '日清泡麵＋吐司'],
    ['sunrise', GOLD, '05:00–05:45 T1 開放', '裝車、檢查前燈（白光 ≥200 流明常亮）與後燈（紅光）、GPS 追蹤器'],
    ['swim', BLUE, '下水前（06:00 出發）', '威德 1；小瓶水沖蛙鏡。水溫約 20–24°C；防寒衣為強制裝備（可加防寒帽、襪套），手套不可'],
    ['bolt', TEAL, '上岸（目標 07:30 前，關門 07:40）', '威德 1；蜂蜜水＋Race On＋冰塊 1（保冰）。遞上保暖物，協助脫防寒衣'],
    ['car', RED, '選手出發後', '收拾 T1 裝備與垃圾 → 主車立刻出發往信義 7-11（選手約 1 小時後到）'],
  ];
  steps.forEach((st, i) => {
    const y = 1.3 + i * 1.12;
    card(s, x, y, w, 1.02);
    iconCircle(s, x + 0.16, y + 0.16, 0.46, st[0], st[1]);
    T(s, st[2], { x: x + 0.75, y: y + 0.1, w: w - 0.9, h: 0.34, fontSize: 13, bold: true, color: st[1] === MUTED ? INK : st[1] });
    T(s, st[3], { x: x + 0.75, y: y + 0.45, w: w - 0.9, h: 0.55, fontSize: 11, color: HEX.dk2 });
  });
  s.addNotes('補給品項來自原簡報「1–3. 起床與 T1 補給」。規則來自 2026 手冊「轉換區 1 開放須知」「強制性游泳／自行車裝備」。');
}

// 17. 自行車段總覽
{
  const s = slide('CONTENT', '自行車 196.7 km：十個補給點、一個禁行區');
  // outbound row
  const chip = (x, y, w, num, name, km, color, note) => {
    card(s, x, y, w, 1.25, { fill: WHITE, line: color, lineW: 1.5 });
    s.addShape(pres.shapes.OVAL, { x: x + 0.12, y: y + 0.12, w: 0.42, h: 0.42, fill: { color }, line: { color, width: 0 } });
    T(s, num, { x: x + 0.12, y: y + 0.12, w: 0.42, h: 0.42, fontSize: 12, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    T(s, name, { x: x + 0.62, y: y + 0.1, w: w - 0.7, h: 0.3, fontSize: 12, bold: true, valign: 'middle' });
    T(s, km, { x: x + 0.62, y: y + 0.4, w: w - 0.7, h: 0.25, fontSize: 9.5, color: MUTED });
    T(s, note, { x: x + 0.12, y: y + 0.68, w: w - 0.24, h: 0.55, fontSize: 9.5, color: HEX.dk2 });
  };
  T(s, '去程（日月潭 → 塔塔加）', { x: 0.5, y: 1.25, w: 6, h: 0.3, fontSize: 13, bold: true, color: TEAL });
  const out = [
    ['T1', '伊達邵碼頭', '0 km・07:40 關門', TEAL, '上岸補給後主車出發'],
    ['4', '7-11 信義門市', '約 33 km', TEAL, '蜂蜜水＋Race On＋冰塊 1（保冰）'],
    ['5', '草坪頭管制哨前', '56 km・紅袋', RED, '梅子可樂＋Race On＋冰塊（黃）2、可頌 1；管制哨不得停車'],
    ['—', '觀山', '約 76 km', 'A9B4C2', '禁行區內，無補給'],
    ['6', '塔塔加夫妻樹', '86 km・12:00 關門', INK, '黑色特殊需求袋；11:30 後可上山接關門選手'],
  ];
  out.forEach((o, i) => chip(0.5 + i * 2.5, 1.58, 2.38, o[0], o[1], o[2], o[3], o[4]));
  rect(s, 5.5, 2.95, 7.33, 0.34, 'FDECEC');
  T(s, '補給車禁行區：草坪頭管制哨 ⇄ 塔塔加（補給、交接一律在管制哨「之前」完成）', { x: 5.6, y: 2.95, w: 7.2, h: 0.34, fontSize: 10.5, bold: true, color: RED, valign: 'middle' });
  T(s, '回程（塔塔加 → 日月潭 → 埔里 → T2）', { x: 0.5, y: 3.42, w: 6, h: 0.3, fontSize: 13, bold: true, color: TEAL });
  const back = [
    ['7', '7-11 和社門市', '約 125 km', TEAL, '原表「合社」。第二車停對面順向；選手停車脫外套。吐司 2 片、葡萄汁，放車頭處'],
    ['8', '豐丘葡萄冰沙', '約 142 km', TEAL, '選手上廁所。冰水（黑）1、蜂蜜水＋Race On＋冰塊（黑）1、葡萄汁＋冰塊＋吐司'],
    ['9', '水里坡', '約 149 km', ORANGE, '冰水至少換 5 隻（機動、跳蛙式定點）、冰頸圈一開始提供'],
    ['10', '埔里高工', '約 185 km', TEAL, '紅茶＋冰塊（紅）2'],
    ['T2', '全家仁愛賽德克店', '196.7 km・15:50 關門', RED, '南豐村、人止關下方；可樂、蠻牛、香蕉、毛巾、濕紙巾；跑步裝備檢錄'],
  ];
  back.forEach((o, i) => chip(0.5 + i * 2.5, 3.75, 2.38, o[0], o[1], o[2], o[3], o[4]));
  // rules strip
  const tips = [
    ['parking', '每站：停道路外、白線內，在車頭交接，人在車外', TEAL],
    ['ban', '不跟車、不配速、不開在選手前後；補給完就走', RED],
    ['headlamp', '車燈全程開啟：前白光 ≥200 流明常亮、後紅光', GOLD],
    ['tools', '不可換車或換輪組（重大機械故障需大會同意）', NAVY],
  ];
  tips.forEach((t, i) => {
    const x = 0.5 + i * 3.1, y = 5.3, w = 2.98;
    card(s, x, y, w, 1.35);
    iconCircle(s, x + 0.15, y + 0.15, 0.44, t[0], t[2]);
    T(s, t[1], { x: x + 0.15, y: y + 0.7, w: w - 0.3, h: 0.75, fontSize: 11, color: HEX.dk2 });
  });
  T(s, '「約」里程依 2025 大會 GPX 推估；以車隊實際路線與 GPS 為準。', { x: 0.5, y: 6.82, w: 12.3, h: 0.26, fontSize: 9.5, color: MUTED });
}

// 18. 補給站 1–5 詳表
{
  const s = slide('CONTENT', '補給明細 ①：起床、T1、信義、草坪頭');
  const rows = [
    ['站點', '品項', '數量', '顏色／容器', '備註'],
    ['1 起床', '日清泡麵＋吐司', '', '', '2025 車隊 03:30 起床'],
    ['2 T1 下水前', '威德', '1', '', ''],
    ['2 T1 下水前', '小瓶水', '', '', '沖蛙鏡用'],
    ['3 T1 上岸', '威德', '1', '', ''],
    ['3 T1 上岸', '蜂蜜水＋Race On＋冰塊', '1', '', '保冰'],
    ['4 7-11 信義門市', '蜂蜜水＋Race On＋冰塊', '1', '', '保冰'],
    ['5 草坪頭', '梅子可樂＋Race On＋冰塊', '2', '黃', '2026：管制哨不得補給→紅袋或管制哨前車頭補給'],
    ['5 草坪頭', '可頌', '1', '', '同上'],
  ];
  tbl(s, rows, { x: 0.5, y: 1.25, w: 7.9, colW: [1.7, 2.3, 0.7, 1.1, 2.1], fontSize: 11, headFill: TEAL, rowH: 0.46, bold0: true });
  T(s, '表格中數量或容器留白，表示原表未註明。', { x: 0.5, y: 5.5, w: 7.9, h: 0.3, fontSize: 10, color: MUTED });
  card(s, 0.5, 5.9, 7.9, 0.95, { fill: 'EAF5F6' });
  iconCircle(s, 0.65, 6.12, 0.5, 'bike', TEAL);
  T(s, '騎車段原則（原表）', { x: 1.3, y: 5.97, w: 6.9, h: 0.3, fontSize: 12, bold: true, color: TEAL });
  T(s, '以主要補給車補給為主，原則上選手不停車；合社與豐丘依各站安排停靠。所有接觸選手的補給由穿背心者在車頭執行。', { x: 1.3, y: 6.27, w: 6.9, h: 0.55, fontSize: 10.5, color: HEX.dk2 });
  // right: visual of bottle color code
  const x = 8.7, w = 4.13;
  card(s, x, 1.25, w, 5.6, { fill: NAVY });
  if (PHOTO_MAP.supply && crewPhoto(PHOTO_MAP.supply.chair)) { s.addImage({ path: crewPhoto(PHOTO_MAP.supply.chair), x: x + 0.2, y: 5.0, w: w - 0.4, h: 1.7, sizing: { type: 'cover', w: w - 0.4, h: 1.7 } }); }
  T(s, '水壺顏色一眼辨認', { x: x + 0.2, y: 1.4, w: w - 0.4, h: 0.4, fontSize: 15, bold: true, color: WHITE });
  const bottles = [['FFFFFF', '白', '蜂蜜水＋Race On／水（黑袋）'], ['F5B700', '黃', '梅子可樂＋Race On（草坪頭）'], ['1F2937', '黑', '冰水／蜂蜜水＋Race On（豐丘）'], ['D7262F', '紅', '紅茶＋冰塊（埔里高工）']];
  const hasChair = !!(PHOTO_MAP.supply && crewPhoto(PHOTO_MAP.supply.chair));
  bottles.forEach((b, i) => {
    const y = 1.9 + i * (hasChair ? 0.62 : 0.78);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.3, y: y + 0.05, w: 0.34, h: hasChair ? 0.48 : 0.6, rectRadius: 0.1, fill: { color: b[0] }, line: { color: 'C9D3DE', width: 1 } });
    T(s, b[1], { x: x + 0.8, y: y + 0.0, w: 0.5, h: 0.3, fontSize: 14, bold: true, color: WHITE });
    T(s, b[2], { x: x + 0.8, y: y + 0.28, w: w - 1.05, h: 0.3, fontSize: 10, color: 'C9D3DE' });
  });
  if (!(PHOTO_MAP.supply && crewPhoto(PHOTO_MAP.supply.chair))) {
  T(s, '補給手口訣', { x: x + 0.2, y: 5.15, w: w - 0.4, h: 0.3, fontSize: 12, bold: true, color: GOLD });
  T(s, '到站前 5 分鐘備好 → 車頭站定 → 遞、收、記 → 收垃圾 → 回報下一車', { x: x + 0.2, y: 5.45, w: w - 0.4, h: 1.2, fontSize: 11.5, color: 'E6ECF2' });
  } else { T(s, '口訣：到站前 5 分鐘備好 → 車頭站定 → 遞、收、記 → 收垃圾 → 回報下一車', { x: x + 0.2, y: 4.42, w: w - 0.4, h: 0.55, fontSize: 10, color: GOLD, bold: true }); }
  s.addNotes('資料來源：原簡報「1–3. 起床與 T1 補給」「4–5. 信義門市與草坪頭」。草坪頭備註依 2026 手冊規定新增。');
}

// 19. 補給站 7–10 詳表
{
  const s = slide('CONTENT', '補給明細 ②：和社、豐丘、水里坡、埔里高工');
  const rows = [
    ['站點', '品項', '數量', '顏色', '使用方式／備註'],
    ['7 7-11 和社門市', '吐司', '2 片', '', '現場吃，放車頭處'],
    ['7 7-11 和社門市', '葡萄汁', '', '', '現場喝，放車頭處'],
    ['8 豐丘葡萄冰沙', '冰水', '1', '黑', '先準備水里坡要用的冰水壺'],
    ['8 豐丘葡萄冰沙', '蜂蜜水＋Race On＋冰塊', '1', '黑', ''],
    ['8 豐丘葡萄冰沙', '葡萄汁＋冰塊＋吐司', '', '', '現場吃'],
    ['9 水里坡', '冰水', '至少換 5 隻', '', '機動補給（跳蛙式定點，不尾隨）'],
    ['9 水里坡', '冰頸圈', '', '', '一開始提供'],
    ['10 埔里高工', '紅茶＋冰塊', '2（選手筆記 1）', '紅', '數量以會議定稿為準'],
  ];
  tbl(s, rows, { x: 0.5, y: 1.25, w: 7.9, colW: [1.75, 2.15, 1.2, 0.8, 2.0], fontSize: 11, headFill: TEAL, rowH: 0.46, bold0: true });
  card(s, 0.5, 5.6, 7.9, 1.32, { fill: 'FFF4F4' });
  iconCircle(s, 0.65, 5.8, 0.5, 'warn', RED);
  T(s, '最容易被開罰的四件事', { x: 1.3, y: 5.66, w: 6.9, h: 0.3, fontSize: 12, bold: true, color: RED });
  T(s, bullets(['車停在白線外或擋到車道；從車窗遞補給', '沒穿背心的人幫忙遞水、扶車、脫外套', '補給完繼續開在選手後面「看一下」', '垃圾留在路邊（亂丟垃圾是重大違規）']), { x: 1.3, y: 5.96, w: 6.9, h: 0.9, fontSize: 10.5, color: HEX.dk2, paraSpaceAfter: 0 });
  T(s, bullets(['車停在白線外或擋到車道；從車窗遞補給', '沒穿背心的人幫忙遞水、扶車、脫外套']), { x: 1.3, y: 5.96, w: 3.4, h: 0.9, fontSize: 10.5, color: HEX.dk2, paraSpaceAfter: 0 });
  const x = 8.7, w = 4.13;
  const notes = [
    ['store', TEAL, '和社 7-11（原表「合社」）', '台21 約 101K。第二補給車停對面（順向）；選手停下脫外套——只有穿背心者可接觸選手。選手 10/25 筆記未列此站，請確認。'],
    ['toilet', NAVY, '豐丘', '選手會停下上廁所；利用這段時間把水里坡的冰水壺備好。選手 10/25 筆記只列冰水＋蜂蜜水，葡萄汁＋吐司請確認。'],
    ['ice', ORANGE, '水里坡', '水里 → 日月潭的長爬坡、正午最熱。原表「補給車跟隨選手勿離太遠」改為：提前到下一個安全停靠點等候，不並行、不尾隨。'],
    ['cup', RED, '埔里高工', '紅茶＋冰塊 2（紅壺）。之後約 12 km 到 T2，提醒選手準備檢錄。'],
  ];
  notes.forEach((n, i) => {
    const y = 1.25 + i * 1.42;
    card(s, x, y, w, 1.32);
    iconCircle(s, x + 0.15, y + 0.15, 0.44, n[0], n[1]);
    T(s, n[2], { x: x + 0.7, y: y + 0.12, w: w - 0.85, h: 0.3, fontSize: 13, bold: true, color: n[1] });
    T(s, n[3], { x: x + 0.2, y: y + 0.62, w: w - 0.4, h: 0.7, fontSize: 10.5, color: HEX.dk2 });
  });
  s.addNotes('資料來源：原簡報「7. 7-11 合社門市」「8. 豐丘葡萄冰沙」「9–10. 水里坡與埔里高工」。選手 2026/10/25 補給筆記：豐丘 冰水x1、蜂蜜水+race on+冰塊x1；水里坡 冰水機動；埔里高工 紅茶+冰塊x1；未列和社站——與原表的差異已標示。');
}

// 20. T2
{
  const s = slide('CONTENT', 'T2 檢錄（15:50 關門）：兩分鐘，不能碰');
  photo(s, 'hb-022-065.jpg', 0.5, 1.3, 4.6, 1.6);
  card(s, 0.5, 3.05, 4.6, 3.85, { fill: NAVY });
  T(s, 'T2 流程', { x: 0.7, y: 3.2, w: 4.2, h: 0.35, fontSize: 14, bold: true, color: GOLD });
  const flow = ['選手進站，補給車停道路外', '跑步強制裝備檢查＋醫療檢測，每人至少停留 2 分鐘', '檢測期間陪跑員、補給員不得協助或觸碰選手', '檢測後換裝、補給、穿反光背心（跑步段全程）', '選手出發 → 腳踏車上車頂架，收拾裝備'];
  flow.forEach((f, i) => {
    const y = 3.62 + i * 0.6;
    s.addShape(pres.shapes.OVAL, { x: 0.75, y: y + 0.05, w: 0.34, h: 0.34, fill: { color: GOLD }, line: { color: GOLD, width: 0 } });
    T(s, String(i + 1), { x: 0.75, y: y + 0.05, w: 0.34, h: 0.34, fontSize: 11, bold: true, color: NAVY, align: 'center', valign: 'middle' });
    T(s, f, { x: 1.22, y: y, w: 3.75, h: 0.48, fontSize: 11, color: 'E6ECF2', valign: 'middle' });
  });
  const x = 5.4, w = 7.43;
  T(s, '跑步強制裝備（T2 檢錄，缺一不可）', { x, y: 1.3, w, h: 0.35, fontSize: 14, bold: true, color: RED });
  const gear = [['phone', '手機（滿電）'], ['headlamp', '頭燈'], ['medkit', '急救毯'], ['gps', 'GPS 追蹤器'], ['backpack', '輕量水袋背包'], ['flashlight', '背包掛燈'], ['vest', '大會反光背心']];
  gear.forEach((g, i) => {
    const gx = x + (i % 4) * 1.86, gy = 1.72 + Math.floor(i / 4) * 1.12;
    card(s, gx, gy, 1.74, 1.02, { fill: WHITE, line: RED, lineW: 1.2 });
    iconCircle(s, gx + 0.12, gy + 0.12, 0.44, g[0], RED);
    T(s, g[1], { x: gx + 0.6, y: gy + 0.1, w: 1.12, h: 0.85, fontSize: 11, bold: true, valign: 'middle' });
  });
  T(s, '團隊準備（原表＋建議）', { x, y: 4.05, w, h: 0.35, fontSize: 14, bold: true, color: TEAL });
  const team = [['shoe', '跑鞋'], ['cap', '跑帽'], ['cup', '可樂・蠻牛・香蕉'], ['towel', '毛巾、濕紙巾'], ['jacket', '保暖衣物（夜）'], ['battery', '頭燈備用電池'], ['bicycle', '車頂架收腳踏車'], ['pin', 'T2：全家仁愛賽德克店']];
  team.forEach((g, i) => {
    const gx = x + (i % 4) * 1.86, gy = 4.47 + Math.floor(i / 4) * 1.12;
    card(s, gx, gy, 1.74, 1.02);
    iconCircle(s, gx + 0.12, gy + 0.12, 0.44, g[0], TEAL);
    T(s, g[1], { x: gx + 0.6, y: gy + 0.1, w: 1.12, h: 0.85, fontSize: 10.5, bold: true, valign: 'middle' });
  });
  T(s, '原表檢錄清單缺「手機、GPS 追蹤器、背包掛燈」，已依 2026 手冊補齊。', { x, y: 6.75, w, h: 0.3, fontSize: 10, color: RED, bold: true });
  s.addNotes('來源：2026 手冊「選手裝備及物資」「Mandatory Run item」；團隊品項來自原簡報「T2 補給與檢錄」。');
}

// 21. 跑步段總覽
{
  const s = slide('CONTENT', '跑步 42 km：西進武嶺，一路向上');
  s.addImage({ path: path.join(ASSETS, 'run_profile.png'), x: 0.5, y: 1.25, w: 7.8, h: 5.2 });
  const rows = [
    ['道路', '地點／路標', '公里數'],
    ['台14', 'T2 全家仁愛賽德克店（原表：霧社店）', '71'],
    ['台14', '霧社', '80'],
    ['台14甲', '霧社', '−0.5'],
    ['台14甲', '清境高空步道', '9.5'],
    ['台14甲', '清境小 7', '11.5'],
    ['台14甲', '梅峰', '14.5'],
    ['台14甲', '翠峰／翠峰路標', '17.5／18'],
    ['台14甲', '鳶峰', '24.2'],
    ['台14甲', '昆陽', '29.3'],
    ['台14甲', '合歡山主峰登山口', '30.6'],
  ];
  tbl(s, rows, { x: 8.6, y: 1.25, w: 4.23, colW: [0.9, 2.2, 1.13], fontSize: 10.5, headFill: ORANGE, rowH: 0.38, align: [null, null, 'right'] });
  T(s, '路標速查：賽事累積里程 ≈ 台14甲里程＋9.5 km（T2 → 霧社約 9 km）。', { x: 8.6, y: 5.5, w: 4.23, h: 0.6, fontSize: 10, color: HEX.dk2 });
  T(s, '剖面海拔為路標概略值（示意）；跑步規則：靠右、禁用登山杖、公路上禁用耳機與手機、紅燈必停。', { x: 0.5, y: 6.55, w: 12.3, h: 0.45, fontSize: 10, color: MUTED });
}

// 22. 跑步前段
{
  const s = slide('CONTENT', '跑步前段：T2 → 清境 → 翠峰（入夜）');
  const rows = [
    ['路段', '距離 km', '可跑 km', '低標時間', '關門時間'],
    ['T2 → 清境高空步道', '19（官方 17.5）', '', '', '18:50'],
    ['清境高空步道 → 翠峰', '8', '3', '1 小時 10 分', '20:20'],
  ];
  tbl(s, rows, { x: 0.5, y: 1.25, w: 7.6, colW: [2.4, 1.8, 0.9, 1.3, 1.2], fontSize: 12, headFill: ORANGE, rowH: 0.5, bold0: true, align: [null, 'center', 'center', 'center', 'center'] });
  const pts = [
    ['sun', GOLD, '17:10 日落在清境之前', '日落約 17:10，離開 T2 一個多小時就天黑。頭燈、背包掛燈、反光背心從 T2 起全程；補給車備手電筒與反光裝備。'],
    ['apple', ORANGE, '清境高空步道（17.5 km，18:50）', '選手脫鐵服上半身、套 T 恤；補給葡萄。停車場有限，提前卡位、停道路外。'],
    ['jacket', TEAL, '翠峰（26 km，20:20）', '套換衣斗篷，脫鐵服下半身換跑褲；26 km 起開放陪跑——陪跑員穿背心、戴手環、裝備齊全在此上場。手冊：車輛停入停車場內。'],
    ['timer', RED, '低標時間怎麼用', '清境→翠峰 8 km 低標 1 小時 10 分：若 19:10 前未離開清境，就會壓到翠峰關門。計時官提前 15 分鐘預警。'],
  ];
  pts.forEach((p, i) => {
    const y = 2.95 + i * 0.98;
    card(s, 0.5, y, 7.6, 0.9);
    iconCircle(s, 0.65, y + 0.12, 0.44, p[0], p[1]);
    T(s, p[2], { x: 1.25, y: y + 0.08, w: 6.7, h: 0.3, fontSize: 12.5, bold: true, color: p[1] });
    T(s, p[3], { x: 1.25, y: y + 0.4, w: 6.7, h: 0.5, fontSize: 10.5, color: HEX.dk2 });
  });
  if (PHOTO_MAP.supply && crewPhoto(PHOTO_MAP.supply.grapes)) s.addImage({ path: crewPhoto(PHOTO_MAP.supply.grapes), x: 8.4, y: 1.25, w: 4.43, h: 1.35, sizing: { type: 'cover', w: 4.43, h: 1.35 } }); else photo(s, 'hb-014-040.jpg', 8.4, 1.25, 4.43, 1.35);
  card(s, 8.4, 2.72, 4.43, 2.3, { fill: WHITE, line: ORANGE, lineW: 1.5 });
  T(s, '跑步段換裝與補給（選手指示）', { x: 8.6, y: 2.8, w: 4.1, h: 0.32, fontSize: 12.5, bold: true, color: ORANGE });
  const chg = [['shirt', ORANGE, '清境高空 17.5 km', '脫鐵服上半身 → 套 T 恤'], ['tent', TEAL, '翠峰 26 km', '套換衣斗篷，鐵服下半身 → 跑褲'], ['cold', BLUE, '鳶峰 33 km', '迪卡儂長褲＋外套＋手套，開始走路']];
  chg.forEach((c, i) => {
    const y = 3.17 + i * 0.5;
    iconCircle(s, 8.6, y, 0.38, c[0], c[1]);
    T(s, c[2], { x: 9.08, y: y - 0.02, w: 1.55, h: 0.42, fontSize: 10.5, bold: true, color: c[1], valign: 'middle' });
    T(s, c[3], { x: 10.55, y: y - 0.02, w: 2.2, h: 0.42, fontSize: 10.5, color: INK, valign: 'middle' });
  });
  iconCircle(s, 8.6, 4.62, 0.3, 'apple', GREEN);
  T(s, '跑步賽段補給：葡萄（每站都備）', { x: 9.0, y: 4.56, w: 3.7, h: 0.4, fontSize: 11, bold: true, color: GREEN, valign: 'middle' });
  card(s, 8.4, 5.14, 4.43, 1.76, { fill: NAVY });
  T(s, '陪跑員強制裝備（翠峰前自檢）', { x: 8.6, y: 5.22, w: 4.1, h: 0.3, fontSize: 12, bold: true, color: GOLD });
  T(s, bullets(['手機（充滿電）', '頭燈', '急救毯']), { x: 8.6, y: 5.55, w: 1.9, h: 0.9, fontSize: 10.5, color: 'E6ECF2', paraSpaceAfter: 1 });
  T(s, bullets(['輕量水袋背包', '背包掛燈', '個人急救包']), { x: 10.6, y: 5.55, w: 2.1, h: 0.9, fontSize: 10.5, color: 'E6ECF2', paraSpaceAfter: 1 });
  T(s, '手冊英文版另列 GPS 追蹤器；建議加帶防風防水外套、手套、帽子、保溫瓶熱飲', { x: 8.6, y: 6.42, w: 4.1, h: 0.45, fontSize: 9.5, color: 'C9D3DE' });
  s.addNotes('距離與低標時間來自原簡報「跑步前段」。清境累計里程賽程表列 17.5 km、跑段表列 19 km（待確認）。');
}

// 23. 跑步後段
{
  const s = slide('CONTENT', '跑步後段：翠峰 → 鳶峰 → 登山口 → 合歡主峰');
  const rows = [
    ['路段', '距離 km', '可跑 km', '低標時間', '關門時間'],
    ['翠峰 → 鳶峰', '6.7', '2.3', '1 小時', '21:40'],
    ['鳶峰 → 合歡山主峰登山口', '6.4', '1.2', '1 小時', '23:00'],
    ['登山口 → 合歡主峰（步道）', '約 1.8', '—', '—', 'FINISH'],
  ];
  tbl(s, rows, { x: 0.5, y: 1.25, w: 7.6, colW: [2.9, 1.3, 1.0, 1.2, 1.2], fontSize: 12, headFill: ORANGE, rowH: 0.5, bold0: true, align: [null, 'center', 'center', 'center', 'center'] });
  const pts = [
    ['food', ORANGE, '鳶峰（33 km，21:40）', '套迪卡儂長褲＋外套＋手套、吃熱食（備「吃不下」方案：薑茶、熱湯、可樂），之後以走路為主。33 km 起強制陪跑，工作人員檢查裝備。'],
    ['parking', NAVY, '昆陽停車場（約 38 km）', '建議作為車隊最後據點（接駁車只行駛登山口 ⇄ 昆陽）。小型車約 47 格、收費；鳶峰以上為暗空公園，停車後 3 分鐘內關大燈。'],
    ['hiking', GOLD, '登山口（40 km，23:00）→ 主峰', '步道約 1.8 km、爬升約 230 m、約 1 小時。黑暗、強風、接近 0°C 的可能，陪跑員全程同行。登山口無停車位：車停武嶺或昆陽。'],
    ['van', TEAL, '完賽後', '走回登山口時告知待命工作人員，由其指派接駁車到昆陽；車隊在昆陽接人後下山至清境住宿。'],
  ];
  pts.forEach((p, i) => {
    const y = 3.45 + i * 0.87;
    card(s, 0.5, y, 7.6, 0.8);
    iconCircle(s, 0.65, y + 0.12, 0.44, p[0], p[1]);
    T(s, p[2], { x: 1.25, y: y + 0.08, w: 6.7, h: 0.3, fontSize: 12.5, bold: true, color: p[1] });
    T(s, p[3], { x: 1.25, y: y + 0.38, w: 6.7, h: 0.42, fontSize: 10.5, color: HEX.dk2 });
  });
  photo(s, 'cover_summit.jpg', 8.4, 1.25, 4.43, 2.6);
  card(s, 8.4, 4.0, 4.43, 2.85, { fill: NAVY });
  T(s, '登頂前後的三個提醒', { x: 8.6, y: 4.15, w: 4.1, h: 0.35, fontSize: 13, bold: true, color: GOLD });
  T(s, bullets(['23:00 檢查點名稱：賽程表「合歡主峰登山口」＝路標表「合歡山登山口」', '罰時在跑步段採總成績加時；申訴須完賽 60 分鐘內書面＋NT$3,000', '手機保持可聯繫到週六（11/14）凌晨 01:00']), { x: 8.6, y: 4.55, w: 4.1, h: 2.2, fontSize: 11, color: 'E6ECF2', paraSpaceAfter: 3 });
  s.addNotes('距離與低標時間來自原簡報「跑步後段」。接駁、申訴、聯繫時間來自 2026 手冊。步道長度與爬升為概略值。');
}

// ===== Section 4 =====
section = '四、後勤、安全與待確認';
pres.addSection({ title: section });
divider('PART 4', '後勤、安全與待確認', '夜間高山、交通住宿、物資清單——以及我們還不確定的事', 'hb-004-010.jpg');

// 25. 夜間高山安全
{
  const s = slide('NIGHT', '夜間高山安全：3,000 公尺以上的最後四小時');
  const cards = [
    ['cold', '2F80ED', '低溫與風寒', ['合歡山 11 月月均約 5°C、最低可到 −2°C；2021 武嶺曾 3°C 體感 −3°C（賽前查中央氣象署預報）', '手冊「登頂保暖提醒」：選手與陪跑員自備防風、防水、保暖衣物（不列入檢查但強烈建議）', '補給車備毛毯、暖暖包、熱飲、乾衣']],
    ['heart', RED, '失溫與高山症徵兆', ['發抖、口齒不清、動作遲緩 → 立即保暖、補熱量', '頭痛、噁心、暈眩、走不直 → 停止上升、通知大會醫護', '醫護與裁判有權隨時中止選手比賽，請配合']],
    ['headlamp', GOLD, '照明與能見度', ['頭燈＋背包掛燈＋反光背心（選手與陪跑員）', '每人備用電池；補給車備備用頭燈', '濃霧常見，停車開警示燈、人員穿反光背心；鳶峰以上暗空公園，停妥 3 分鐘內關大燈']],
    ['phone', GREEN, '通訊與待命', ['所有陪跑員／補給員手機保持可聯繫到週六 01:00', '山區訊號不穩：約定「到站回報」與失聯處置', '緊急狀況（建議流程）：先打 119，再通知大會與車長']],
  ];
  cards.forEach((c, i) => {
    const x = 0.5 + i * 3.12, y = 1.3, w = 2.97, h = 3.45;
    card(s, x, y, w, h, { fill: '14263F', line: '2A4160', lineW: 1 });
    iconCircle(s, x + 0.2, y + 0.2, 0.56, c[0], c[1]);
    T(s, c[2], { x: x + 0.2, y: y + 0.88, w: w - 0.4, h: 0.4, fontSize: 15, bold: true, color: WHITE });
    T(s, bullets(c[3]), { x: x + 0.2, y: y + 1.3, w: w - 0.4, h: h - 1.4, fontSize: 11, color: 'DCE3EA', paraSpaceAfter: 4 });
  });
  const st = [['17:10', '日落（清境之前）', GOLD], ['3,085 m', '昆陽停車場・最後據點', '8FB8FF'], ['3,417 m', '合歡主峰・終點', WHITE], ['01:00', '手機待命至週六凌晨', GREEN]];
  st.forEach((v, i) => stat(s, 0.5 + i * 3.12, 4.9, 2.97, v[0], v[1], v[2], 'C9D3DE', 30));
  card(s, 0.5, 5.85, 12.33, 1.05, { fill: GOLD });
  T(s, '昆陽停車場是最後據點：23:00 前全車隊到齊，分工「等人、保暖、熱飲、接駁聯絡」。完賽後選手走回登山口 → 接駁到昆陽 → 車隊下山到清境住宿；11/14 09:00 愛立家民宿頒獎。', { x: 0.75, y: 5.85, w: 11.9, h: 1.05, fontSize: 12.5, bold: true, color: NAVY, valign: 'middle' });
  s.addNotes('來源：2026 手冊「登頂保暖提醒」「醫療相關」「大會接駁車」「陪跑員相關資訊」。氣溫為 11 月合歡山夜間常見範圍，請以賽前中央氣象署預報為準。');
}

// 26. 後勤
{
  const s = slide('CONTENT', '後勤：出發、配車、住宿與回程');
  const rows = [
    ['日期', '車輛', '成員（駕駛）', '行程／任務'],
    ['11/11（三）\n06:00 出發', '主補給車 RAV4', '大頭（駕駛）、小麥', '新潤翡麗出發 → 埔里 → 日月潭；提前兩天適應場地與心情'],
    ['11/12（四）\n06:00 出發', '第二補給車 GLC 250', '傑哥（駕駛）、佳惠、珮淳', '和我一起愛運動工作室出發 → 日月潭溫德姆飯店'],
    ['11/12（四）\n06:00 出發', '攝影器材車 CRV', '金剛（駕駛）、小朱、聖恩、大顆頭', '同上；載自行車'],
    ['11/12（四）\n06:00 出發', '攝影師座車 HELIX', '阿徹（駕駛）、攝影師', '同上'],
    ['11/13（五）', '親友專車 MODEL Y', '俊儀（駕駛）', '中午南下，直奔 T2 見證歷史'],
    ['11/13 賽事', 'RAV4／GLC 250／CRV', '大顆頭／傑哥／小朱 駕駛', '賽事分配見第 14 頁；攝影師改搭 GLC'],
    ['11/14（六）', '攝影器材車 CRV', '金剛（駕駛）不載人', '頒獎前離開，載自行車與冰箱回工作室'],
    ['11/14（六）', '親友專車 MODEL Y', '俊儀（駕駛）、小朱、大顆頭、聖恩', '頒獎典禮後返北'],
    ['11/14（六）', '主補給車 RAV4', '大頭（駕駛）、小麥', '頒獎典禮後返北'],
    ['11/14（六）', '第二補給車 GLC 250', '傑哥（駕駛）、佳惠、珮淳', '頒獎典禮後返北；阿徹與攝影師開 HELIX 自行返北'],
  ];
  tbl(s, rows, { x: 0.5, y: 1.25, w: 8.2, colW: [1.35, 1.75, 2.35, 2.75], fontSize: 10.5, headFill: NAVY, rowH: 0.47, bold0: true });
  const x = 9.0, w = 3.83;
  const info = [
    ['hotel', BLUE, '賽前住宿（11/12）', '日月潭力麗溫德姆溫泉酒店＋湧泉民宿（攝影師）。分房見下一頁。'],
    ['bed', TEAL, '賽後住宿（11/13 夜）', '清境 白雲渡假山莊。分房見下一頁。終點距台北約 270 km。'],
    ['car', ORANGE, '車程參考', '塔塔加 → 霧社約 3 小時以上（含山路）；霧社 → 昆陽約 1.5 小時。提早出發，避開管制。'],
    ['calendar', GOLD, '交通管制與施工', '11/13 06:00–24:00 台14 58.5–80K、台14甲 0–31.5K、台21 68–141K 志工疏導（非全線封閉）。台21 132.3K 和社二號明隧道災修：每整點放行約 10 分鐘（12–13 時不管制），賽日是否暫停待確認。'],
  ];
  info.forEach((n, i) => {
    const y = 1.25 + i * 1.42;
    card(s, x, y, w, 1.32);
    iconCircle(s, x + 0.15, y + 0.15, 0.44, n[0], n[1]);
    T(s, n[2], { x: x + 0.7, y: y + 0.12, w: w - 0.85, h: 0.3, fontSize: 12.5, bold: true, color: n[1] });
    T(s, n[3], { x: x + 0.2, y: y + 0.62, w: w - 0.4, h: 0.7, fontSize: 10, color: HEX.dk2 });
  });
  s.addNotes('車輛與成員分配依選手 11/2 提供、11/3 更新之分配表；住宿與距離來自 2026 手冊「交通與住宿資訊」；交通管制來自公路局公告與新聞。車程為估計值。');
}

// 27b. 住宿分房
{
  const s = slide('CONTENT', '住宿分房：11/12 日月潭、11/13 清境');
  const roomCard = (x, y, w, h, title, sub, rooms, color) => {
    card(s, x, y, w, h, { fill: WHITE, line: color, lineW: 1.5 });
    cardHeader(s, x, y, w, 0.95, color);
    iconCircle(s, x + 0.2, y + 0.2, 0.55, 'hotel', WHITE, color === TEAL ? 'teal' : 'navy');
    T(s, title, { x: x + 0.9, y: y + 0.12, w: w - 1.05, h: 0.45, fontSize: 16, bold: true, color: WHITE, valign: 'middle' });
    T(s, sub, { x: x + 0.9, y: y + 0.56, w: w - 1.05, h: 0.3, fontSize: 10.5, color: WHITE });
    rooms.forEach((r, i) => {
      const ry = y + 1.15 + i * 0.86;
      pill(s, x + 0.25, ry + 0.1, 1.45, 0.34, r[0], color, WHITE, 10.5);
      const names = r[1];
      names.forEach((nm, j) => pill(s, x + 1.8 + (j % 4) * 1.02, ry + (j < 4 ? 0.1 : 0.5), 0.95, 0.32, nm, CARD, INK, 11));
    });
  };
  roomCard(0.5, 1.3, 6.05, 5.1, '11/12（四）日月潭', '力麗溫德姆溫泉酒店・報到、說明會、T1 皆在附近', [
    ['雙人房', ['大頭', '小麥']],
    ['雙人房', ['佳惠', '珮淳']],
    ['三人房', ['小朱', '阿徹', '傑哥']],
    ['三人房', ['金剛', '大顆頭', '聖恩']],
    ['湧泉民宿 單人', ['攝影師']],
  ], NAVY);
  roomCard(6.78, 1.3, 6.05, 5.1, '11/13（五）清境', '白雲渡假山莊・完賽下山後直接入住，11/14 09:00 愛立家頒獎', [
    ['雙人房', ['大頭', '小麥']],
    ['雙人房', ['佳惠', '珮淳']],
    ['雙人房', ['阿徹', '攝影師']],
    ['六人房', ['俊儀', '傑哥', '大顆頭', '聖恩', '金剛', '小朱']],
  ], TEAL);
  T(s, '分房依選手 11/3 更新版。11/13 深夜從昆陽下山約 1 小時到清境，先把鑰匙與房號發給每台車的車長。', { x: 0.5, y: 6.55, w: 12.3, h: 0.45, fontSize: 10.5, color: HEX.dk2 });
  s.addNotes('住宿分房：選手 2026/11/3 更新版。11/12 溫德姆四間＋湧泉民宿一間；11/13 白雲渡假山莊四間（三雙人＋一六人）。');
}

// 28. 裝備與物資清單
{
  const s = slide('CONTENT', '裝備與物資總清單（出發前逐項打勾）');
  const cols = [
    ['選手強制裝備', RED, 'list', [
      '游泳：防寒衣（可加防寒帽、襪套；禁手套）、蛙鏡、大會泳帽',
      '自行車：三鐵車或公路車（擇一）、前燈白光 ≥200 流明、後燈紅光、安全帽、車貼、GPS',
      '跑步：手機、頭燈、急救毯、GPS 追蹤器、輕量水袋背包、背包掛燈、反光背心',
      '號碼布、紋身貼、選手手環、保險證明（放補給車）',
      '跑步換裝（自備）：跑鞋、跑帽、T 恤、跑褲、迪卡儂長褲、外套、手套、換衣斗篷',
    ]],
    ['陪跑員／補給員', GREEN, 'vest', [
      '陪跑員：手機、頭燈、急救毯、輕量水袋背包、背包掛燈、個人急救包',
      '反光背心 2 件、陪跑員手環（報到當天戴）',
      '保暖：防風防水外套、手套、帽子、保溫瓶',
      '備用電池、行動電源、哨子、雨衣',
    ]],
    ['補給車物資', TEAL, 'box', [
      '飲品：威德、蜂蜜水、Race On、梅子可樂、紅茶、葡萄汁、可樂、蠻牛、水、冰塊、冰頸圈',
      '食物：泡麵、吐司、可頌、香蕉、葡萄（跑步段主要補給）、熱食（鳶峰）',
      '水壺：白、黃、黑、紅各色；冰霸杯、保溫瓶、保冰桶',
      '毛巾、濕紙巾、垃圾袋、暖暖包、毛毯、乾衣、換洗褲、車貼',
      '急救包、車充、行動電源、備用頭燈 ×2＋電池、手電筒、警示燈、內胎／CO2／工具（破胎備案）',
    ]],
  ];
  cols.forEach((c, i) => {
    const x = 0.5 + i * 4.15, y = 1.3, w = 4.0, h = 5.4;
    card(s, x, y, w, h, { fill: WHITE, line: c[1], lineW: 1.5 });
    cardHeader(s, x, y, w, 0.8, c[1]);
    iconCircle(s, x + 0.18, y + 0.14, 0.52, c[2], WHITE, c[1] === GREEN ? 'green' : (c[1] === TEAL ? 'teal' : 'red'));
    T(s, c[0], { x: x + 0.85, y: y + 0.12, w: w - 1.0, h: 0.56, fontSize: 16, bold: true, color: WHITE, valign: 'middle' });
    c[3].forEach((it, j) => {
      const yy = y + 0.98 + j * 0.86;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.22, y: yy + 0.05, w: 0.26, h: 0.26, rectRadius: 0.05, fill: { color: WHITE }, line: { color: c[1], width: 1.2 } });
      T(s, it, { x: x + 0.6, y: yy, w: w - 0.8, h: 0.85, fontSize: 11, color: INK, valign: 'top' });
    });
  });
  T(s, '依 2026 手冊強制裝備＋原簡報補給品項＋2025 車隊清單（FXT_050 補給清單）整理；數量請在會議中填入。', { x: 0.5, y: 6.78, w: 12.3, h: 0.25, fontSize: 9.5, color: MUTED });
}

// 28. 待確認事項
{
  const s = slide('CONTENT', '待確認事項：會議前請車長逐項回覆');
  const items = [
    ['清境檢查點里程', '賽程表 17.5 km、跑段距離表 T2→清境 19 km。以哪個算低標時間？'],
    ['T2 位置與店名', '原表「全家仁愛霧社店」；2025 手冊寫「全家仁愛賽德克店停車場」（南豐村、人止關下方），與 GPX 終點（約 714 m）吻合；2026 手冊只給地圖連結。請賽前導航實測。'],
    ['最後檢查點名稱', '賽程表「合歡主峰登山口」＝跑段／路標表「合歡山登山口」。'],
    ['草坪頭補給方式', '2026 規定管制哨不得停車／補給。改紅色特殊需求袋，或在管制哨前的車頭補給？由誰執行？'],
    ['第二車在和社的補給', '手冊：每位選手僅限一輛官方補給車進行補給。第二車定點是否合規？需向大會確認。'],
    ['陪跑員報到（已定）', '陪跑員小麥 11/12 13:00–15:00 須與大頭一同報到、健康檢測、戴手環；補給手金剛領背心。'],
    ['特殊需求袋打包與繳交', '黑袋（塔塔加）與紅袋（草坪頭）內容定稿；比賽早上出發前由誰繳交？'],
    ['夜間值班與通訊', '手機待命到 11/14 01:00；各車「到站回報」窗口與失聯處置。'],
    ['回程地標里程', '和社、豐丘、水里坡、埔里高工的「約」里程為 GPX 推估，請以車隊實際路線為準。'],
    ['終點關門時間', '手冊只列到登山口 23:00；交管公告活動至 24:00。主峰是否另有關門？'],
    ['台21 132.3K 明隧道管制', '和社二號明隧道災修每整點放行約 10 分鐘（至 2027/2）。賽日是否暫停？補給車行程以整點為節點。'],
    ['清晨與終點停車', '伊達邵 05:00 補給車停哪裡（纜車停車場 09:00 才開）？登山口無停車位，停武嶺或昆陽。'],
  ];
  items.forEach((it, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 4.15, y = 1.25 + row * 1.37, w = 4.0, h = 1.27;
    card(s, x, y, w, h, { fill: WHITE, line: BORDER, lineW: 1 });
    s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: y + 0.12, w: 0.4, h: 0.4, fill: { color: GOLD }, line: { color: GOLD, width: 0 } });
    T(s, String(i + 1).padStart(2, '0'), { x: x + 0.15, y: y + 0.12, w: 0.4, h: 0.4, fontSize: 11, bold: true, color: NAVY, align: 'center', valign: 'middle' });
    T(s, it[0], { x: x + 0.65, y: y + 0.1, w: w - 0.8, h: 0.44, fontSize: 12.5, bold: true, valign: 'middle' });
    T(s, it[1], { x: x + 0.18, y: y + 0.56, w: w - 0.36, h: h - 0.6, fontSize: 9.5, color: HEX.dk2 });
  });
  T(s, '前三項為原簡報已列之待確認事項；其餘為對照 2026 選手手冊與公開資料後新增。', { x: 0.5, y: 6.82, w: 12.3, h: 0.26, fontSize: 9.5, color: MUTED });
}

// 29. 結尾
{
  const s = pres.addSlide({ masterName: 'BLANK', sectionTitle: section });
  photo(s, 'hb-001-000.jpg', 0, 0, W, H);
  rect(s, 0, 0, W, H, '000000', { fill: { color: '000000', transparency: 40 }, line: { color: '000000', width: 0 } });
  T(s, '如果沒有天賦，那就一直重複。', { x: 0.8, y: 1.0, w: 11.7, h: 1.1, fontSize: 44, bold: true, color: WHITE });
  T(s, '#See you at the top. ・ 合歡主峰 3,417 m ・ 登山口 23:00 前通過，一起走完最後 1.8 km', { x: 0.8, y: 2.1, w: 11.7, h: 0.45, fontSize: 16, color: GOLD, bold: true });
  T(s, '大頭 022 ｜ 小麥（太太・陪跑員）・大顆頭・金剛・聖恩 ｜ 傑哥・佳惠・珮淳・攝影師 ｜ 小朱・阿徹 ｜ 俊儀', { x: 0.8, y: 2.6, w: 11.7, h: 0.5, fontSize: 14, color: 'E6ECF2' });
  const three = [['vest', '只有背心能碰選手'], ['parking', '停白線內・車外・車頭'], ['clock', '關門前 15 分鐘預警']];
  three.forEach((t, i) => {
    const x = 0.8 + i * 4.0;
    card(s, x, 3.4, 3.8, 1.1, { fill: '000000' });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 3.4, w: 3.8, h: 1.1, rectRadius: 0.12, fill: { color: '000000', transparency: 35 }, line: { color: GOLD, width: 1 } });
    iconCircle(s, x + 0.2, 3.65, 0.6, t[0], GOLD, 'navy');
    T(s, t[1], { x: x + 0.95, y: 3.4, w: 2.75, h: 1.1, fontSize: 16, bold: true, color: WHITE, valign: 'middle' });
  });
  card(s, 0.8, 4.85, 11.73, 2.1, { fill: '000000' });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.85, w: 11.73, h: 2.1, rectRadius: 0.12, fill: { color: '000000', transparency: 35 }, line: { color: '3C4858', width: 1 } });
  T(s, '緊急聯絡（會議中填寫）', { x: 1.0, y: 4.95, w: 6, h: 0.35, fontSize: 13, bold: true, color: GOLD });
  const contacts = ['大會賽事辦公室：＿＿＿＿＿＿＿', '大會醫護／救護：＿＿＿＿＿＿＿', 'RAV4 大顆頭：＿＿＿＿＿＿＿', 'GLC 傑哥：＿＿＿＿＿＿＿', 'CRV 小朱：＿＿＿＿＿＿＿', 'MODEL Y 俊儀：＿＿＿＿＿＿＿', '小麥（陪跑）／金剛（補給）：＿＿＿', '緊急救援：119'];
  contacts.forEach((c, i) => T(s, c, { x: 1.0 + (i % 4) * 2.9, y: 5.4 + Math.floor(i / 4) * 0.6, w: 2.8, h: 0.5, fontSize: 11.5, color: 'E6ECF2', valign: 'middle' }));
  T(s, '資料來源：2026 FXT 選手手冊（Google Drive 2026/10/29 更新版）、原「FXT 賽事補給計畫」簡報、2025 FXT 選手手冊影像、2025 大會 GPX、2025 車隊清單。', { x: 0.8, y: 7.05, w: 11.7, h: 0.3, fontSize: 9, color: 'A9B4C2' });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log('written', OUT);
})().catch(e => { console.error(e); process.exit(1); });
