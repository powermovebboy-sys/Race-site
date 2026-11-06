// 手寫列印版（A4 橫式、白底）：第 1 頁 給大頭的一句話、第 2 頁 緊急聯絡與關門速查
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');
const OUT = process.argv[2] || path.join(__dirname, '..', '2026_FXT_手寫列印版.pptx');
const PHOTOS = path.join(__dirname, 'photos');
const PHOTO_MAP = fs.existsSync(path.join(__dirname, 'photo_map.json')) ? JSON.parse(fs.readFileSync(path.join(__dirname, 'photo_map.json'), 'utf8')) : {};
const crewPhoto = f => (f && fs.existsSync(path.join(PHOTOS, 'crew', f))) ? path.join(PHOTOS, 'crew', f) : null;
const FONT = 'Microsoft JhengHei';
const INK = '0B0F14', MUTED = '7A8794', RED = 'D7262F', LINE = 'B8C2CE', NAVY = '0B1F3A', WHITE = 'FFFFFF', LIGHT = 'F3F5F7';

const pres = new pptxgen();
pres.defineLayout({ name: 'A4L', width: 11.69, height: 8.27 });
pres.layout = 'A4L';
pres.title = '2026 FXT 手寫列印版';
const T = (s, text, o) => s.addText(text, Object.assign({ fontFace: FONT, color: INK, margin: 2 }, o));
const box = (s, x, y, w, h, o = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.08, fill: { color: WHITE }, line: { color: LINE, width: 1 } }, o));
const foot = (s, t) => T(s, t, { x: 0.5, y: 7.75, w: 10.69, h: 0.3, fontSize: 9, color: MUTED });

// ---- 第 1 頁：給大頭的一句話 ----
{
  const s = pres.addSlide();
  T(s, '最後一哩路：給大頭的一句話', { x: 0.5, y: 0.35, w: 8, h: 0.6, fontSize: 24, bold: true });
  T(s, '2026 FXT VI・11/13 登山口 → 合歡主峰 1.8 km，一句一句說給他聽。今天先寫，先不說。', { x: 0.5, y: 0.92, w: 10.69, h: 0.3, fontSize: 11, color: MUTED });
  const team = [['小麥', '太太・陪跑員'], ['大顆頭', 'RAV4 車長'], ['金剛', '補給手'], ['聖恩', 'RAV4'], ['傑哥', 'GLC 車長'], ['佳惠', 'GLC'], ['珮淳', 'GLC'], ['攝影師', '影像紀錄'], ['小朱', 'CRV 車長'], ['阿徹', 'CRV・HELIX'], ['俊儀', 'MODEL X'], ['教練', '謝漢霖']];
  const cols = 4, cw = 2.5, gx = 0.23, ch = 2.0, gy = 0.15, x0 = 0.5, y0 = 1.35;
  team.forEach((m, i) => {
    const x = x0 + (i % cols) * (cw + gx), y = y0 + Math.floor(i / cols) * (ch + gy);
    box(s, x, y, cw, ch, { fill: { color: LIGHT }, line: { color: LINE, width: 0.75 } });
    const pf = crewPhoto((PHOTO_MAP.crew || {})[m[0]]), d = 0.6;
    if (pf) s.addImage({ path: pf, x: x + 0.15, y: y + 0.14, w: d, h: d, sizing: { type: 'cover', w: d, h: d }, rounding: true });
    T(s, m[0], { x: x + 0.85, y: y + 0.12, w: cw - 1.0, h: 0.36, fontSize: 14, bold: true, valign: 'middle' });
    T(s, m[1], { x: x + 0.85, y: y + 0.46, w: cw - 1.0, h: 0.26, fontSize: 9.5, color: MUTED });
    box(s, x + 0.15, y + 0.85, cw - 0.3, ch - 1.0);
    // 兩條淡淡的書寫底線
    [0.5, 0.85].forEach(dy => s.addShape(pres.shapes.LINE, { x: x + 0.3, y: y + 0.85 + dy, w: cw - 0.6, h: 0, line: { color: 'E1E6EB', width: 0.75, dashType: 'dash' } }));
  });
  foot(s, '寫完交給小麥帶上山。這一頁不放進會議簡報，比賽當天最後再說。');
}

// ---- 第 2 頁：緊急聯絡與關門速查 ----
{
  const s = pres.addSlide();
  T(s, '緊急聯絡與關門速查（每車一張）', { x: 0.5, y: 0.35, w: 8, h: 0.6, fontSize: 24, bold: true });
  T(s, '2026 FXT VI・11/13（五）・大頭 022・以 11/12 賽前說明會公布內容為最終依據', { x: 0.5, y: 0.92, w: 10.69, h: 0.3, fontSize: 11, color: MUTED });
  // 左：聯絡表
  T(s, '緊急聯絡（會議中手寫填入）', { x: 0.5, y: 1.35, w: 5.6, h: 0.35, fontSize: 13, bold: true, color: RED });
  const contacts = ['大會賽事辦公室', '大會醫護／救護', 'RAV4 主車 大顆頭', 'GLC 第二車 傑哥', 'CRV 器材車 小朱', 'HELIX 阿徹', 'MODEL X 俊儀', '陪跑員 小麥', '補給手 金剛', '力麗溫德姆（11/12 住宿）', '清境愛立家民宿（11/13 住宿）', '緊急救援 119・警察 110'];
  contacts.forEach((c, i) => {
    const y = 1.78 + i * 0.47;
    T(s, c, { x: 0.5, y, w: 2.3, h: 0.4, fontSize: 11, bold: true, valign: 'middle' });
    box(s, 2.85, y + 0.02, 3.25, 0.36, { rectRadius: 0.04 });
  });
  // 右：關門速查
  const rx = 6.5, rw = 4.69;
  T(s, '官方關門速查（2026 手冊）', { x: rx, y: 1.35, w: rw, h: 0.35, fontSize: 13, bold: true, color: NAVY });
  const rows = [['時間', '地點', '團隊動作'],
    ['05:00–05:45', 'T1 伊達邵碼頭開放', '陪跑員進 T1；車停溫德姆'],
    ['06:00', '游泳出發', '上岸補給就位'],
    ['07:40', '游泳關門', '收 T1 → 主車往信義'],
    ['12:00', '夫妻樹 86 km', '黑袋；11:30 後可上山接人'],
    ['15:50', 'T2 全家仁愛賽德克店', '檢錄 ≥2 分鐘；不碰選手'],
    ['18:50', '清境高空步道 19 km', 'T 恤＋葡萄'],
    ['20:20', '翠峰 26 km', '陪跑員上場；斗篷換跑褲'],
    ['21:40', '鳶峰 33 km', '強制陪跑；長褲外套手套'],
    ['23:00', '主峰登山口 40 km', '最後關門；接駁只到昆陽']];
  s.addTable(rows.map((r, i) => r.map(c => ({ text: c, options: { fontFace: FONT, fontSize: 9.5, bold: i === 0, color: i === 0 ? WHITE : INK, fill: { color: i === 0 ? NAVY : WHITE }, valign: 'middle', margin: 3 } }))), { x: rx, y: 1.78, w: rw, colW: [1.0, 1.75, 1.94], rowH: 0.42, border: { type: 'solid', pt: 0.5, color: LINE } });
  box(s, rx, 6.2, rw, 1.3, { fill: { color: LIGHT }, line: { color: LINE, width: 0.75 } });
  T(s, '三件事', { x: rx + 0.15, y: 6.26, w: 2, h: 0.3, fontSize: 11, bold: true, color: RED });
  T(s, '1. 只有背心能碰選手（補給手金剛、陪跑員小麥）\n2. 停白線內・人在車外・在車頭交接；不跟車、不配速\n3. 關門前 15 分鐘預警；到站回報下一車', { x: rx + 0.15, y: 6.55, w: rw - 0.3, h: 0.9, fontSize: 10, valign: 'top' });
  foot(s, '來源：2026 FXT 選手手冊、團隊補給計畫（11/6 版）。');
}

pres.writeFile({ fileName: OUT }).then(() => console.log('written', OUT));
