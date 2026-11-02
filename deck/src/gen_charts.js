// Generates the course elevation profile (two panels: bike / run) and schematic route map as PNGs (SVG -> sharp).
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

const OUT = path.join(__dirname, 'assets');
const DATA = path.join(__dirname, 'data', 'bike_profile.csv');
const FONT = "'WenQuanYi Zen Hei','Microsoft JhengHei','Noto Sans CJK TC',sans-serif";

const C = {
  ink: '#0B0F14', ink2: '#3C4858', muted: '#7A8794', grid: '#E3E7EC', surface: '#FFFFFF',
  bike: '#0E7C86', run: '#F26B1F', red: '#D7262F', gold: '#E0A800', navy: '#0B1F3A',
};

const rows = fs.readFileSync(DATA, 'utf8').trim().split('\n').slice(1).map(l => l.split(',').map(Number));
const bike = rows.map(r => ({ km: r[0], ele: r[1], lat: r[2], lon: r[3] }));
const BIKE_END = bike[bike.length - 1].km; // ~198 km in GPX (official 196.7)
const runPts = [
  [0, 714], [4, 900], [9, 1148], [13, 1450], [17.5, 1750], [20, 1800], [23, 2100], [26, 2309],
  [29.5, 2520], [33, 2756], [36, 2950], [38, 3085], [40, 3275], [42, 3417],
];
function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;'); }
function textW(s, fs) { let w = 0; for (const ch of s) w += (ch.charCodeAt(0) < 256 ? 0.56 : 1.0) * fs; return w; }
function eleAt(km) { let b = bike[0]; for (const p of bike) if (Math.abs(p.km - km) < Math.abs(b.km - km)) b = p; return b.ele; }
function runEleAt(km) {
  for (let i = 1; i < runPts.length; i++) if (km <= runPts[i][0]) { const [k0, e0] = runPts[i - 1], [k1, e1] = runPts[i]; return e0 + (e1 - e0) * (km - k0) / (k1 - k0); }
  return runPts[runPts.length - 1][1];
}

// greedy label level assignment to avoid horizontal overlap at the same level
function assignLevels(marks, fs, nLevels) {
  const sorted = [...marks].sort((a, b) => a.X - b.X);
  const lastRight = new Array(nLevels).fill(-1e9);
  sorted.forEach(mk => {
    if (mk.level != null) { const w0 = Math.max(textW(mk.label, fs), textW(mk.sub || '', fs * 0.8)) + 24; const l0 = mk.anchor === 'end' ? mk.X - w0 : (mk.anchor === 'start' ? mk.X : mk.X - w0 / 2); lastRight[mk.level] = Math.max(lastRight[mk.level], l0 + w0); return; }
    const w = Math.max(textW(mk.label, fs), textW(mk.sub || '', fs * 0.8)) + 24;
    const left = mk.anchor === 'end' ? mk.X - w : (mk.anchor === 'start' ? mk.X : mk.X - w / 2);
    let lvl = 0;
    for (let l = 0; l < nLevels; l++) { if (left > lastRight[l]) { lvl = l; break; } lvl = l; if (l === nLevels - 1) { lvl = lastRight.indexOf(Math.min(...lastRight)); } }
    mk.level = lvl; lastRight[lvl] = left + w;
  });
}

function profileSvg() {
  const W = 2600, H = 1120;
  const top = 380, bottom = 110, gap = 90, left = 120, right = 40;
  const ph = H - top - bottom;
  const panelW = W - left - right - gap;
  const bw = Math.round(panelW * 0.66), rw = panelW - bw;
  const bx0 = left, rx0 = left + bw + gap;
  const y = ele => top + ph - (ele / 3600) * ph;
  const xb = km => bx0 + (km / BIKE_END) * bw;
  const xr = km => rx0 + (km / 42) * rw;

  let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT}">`;
  s += `<rect width="${W}" height="${H}" fill="${C.surface}"/>`;
  s += `<defs><pattern id="hatch" width="18" height="18" patternUnits="userSpaceOnUse" patternTransform="rotate(45)"><rect width="18" height="18" fill="${C.red}" fill-opacity="0.05"/><line x1="0" y1="0" x2="0" y2="18" stroke="${C.red}" stroke-opacity="0.30" stroke-width="3"/></pattern>
  <linearGradient id="gb" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${C.bike}" stop-opacity="0.35"/><stop offset="1" stop-color="${C.bike}" stop-opacity="0.04"/></linearGradient>
  <linearGradient id="gr" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${C.run}" stop-opacity="0.40"/><stop offset="1" stop-color="${C.run}" stop-opacity="0.05"/></linearGradient></defs>`;
  // grid (both panels)
  for (let e = 500; e <= 3500; e += 500) {
    s += `<line x1="${bx0}" x2="${bx0 + bw}" y1="${y(e)}" y2="${y(e)}" stroke="${C.grid}" stroke-width="2"/>`;
    s += `<line x1="${rx0}" x2="${rx0 + rw}" y1="${y(e)}" y2="${y(e)}" stroke="${C.grid}" stroke-width="2"/>`;
    s += `<text x="${bx0 - 14}" y="${y(e) + 9}" text-anchor="end" font-size="26" fill="${C.muted}">${e.toLocaleString()} m</text>`;
  }
  // panel titles
  s += `<text x="${bx0}" y="52" font-size="40" font-weight="700" fill="${C.bike}">自行車 196.7 km</text><text x="${bx0 + 330}" y="52" font-size="28" fill="${C.ink2}">日月潭 → 塔塔加夫妻樹折返 → 日月潭 → 埔里 → T2・爬升 3,600 m+</text>`;
  s += `<text x="${rx0}" y="52" font-size="40" font-weight="700" fill="${C.run}">跑步 42 km</text><text x="${rx0 + 220}" y="52" font-size="28" fill="${C.ink2}">T2 → 台14甲 → 合歡主峰・爬升 3,000 m+</text>`;
  // band
  const z0 = xb(56), z1 = xb(2 * 86.2 - 56);
  s += `<rect x="${z0}" y="${top}" width="${z1 - z0}" height="${ph}" fill="url(#hatch)"/>`;
  s += `<rect x="${(z0 + z1) / 2 - 250}" y="${y(0) - 64}" width="500" height="44" rx="10" fill="#FFFFFF" fill-opacity="0.9"/>`;
  s += `<text x="${(z0 + z1) / 2}" y="${y(0) - 32}" text-anchor="middle" font-size="27" font-weight="700" fill="${C.red}">補給車禁行區：草坪頭 ⇄ 塔塔加（56–116 km）</text>`;
  // bike area + line
  let d = `M${xb(0)},${y(0)} `; bike.forEach(p => { d += `L${xb(p.km).toFixed(1)},${y(p.ele).toFixed(1)} `; }); d += `L${xb(BIKE_END)},${y(0)} Z`;
  s += `<path d="${d}" fill="url(#gb)"/>`;
  let dl = ''; bike.forEach((p, i) => { dl += (i ? 'L' : 'M') + `${xb(p.km).toFixed(1)},${y(p.ele).toFixed(1)} `; });
  s += `<path d="${dl}" fill="none" stroke="${C.bike}" stroke-width="4" stroke-linejoin="round"/>`;
  // run area + line
  let dr = `M${xr(0)},${y(0)} `; runPts.forEach(p => { dr += `L${xr(p[0]).toFixed(1)},${y(p[1]).toFixed(1)} `; }); dr += `L${xr(42)},${y(0)} Z`;
  s += `<path d="${dr}" fill="url(#gr)"/>`;
  let drl = ''; runPts.forEach((p, i) => { drl += (i ? 'L' : 'M') + `${xr(p[0]).toFixed(1)},${y(p[1]).toFixed(1)} `; });
  s += `<path d="${drl}" fill="none" stroke="${C.run}" stroke-width="4" stroke-linejoin="round"/>`;
  // baselines + ticks
  s += `<line x1="${bx0}" x2="${bx0 + bw}" y1="${y(0)}" y2="${y(0)}" stroke="${C.ink2}" stroke-width="2"/>`;
  s += `<line x1="${rx0}" x2="${rx0 + rw}" y1="${y(0)}" y2="${y(0)}" stroke="${C.ink2}" stroke-width="2"/>`;
  for (let k = 0; k <= 180; k += 20) s += `<line x1="${xb(k)}" x2="${xb(k)}" y1="${y(0)}" y2="${y(0) + 10}" stroke="${C.muted}" stroke-width="2"/><text x="${xb(k)}" y="${y(0) + 44}" text-anchor="middle" font-size="26" fill="${C.muted}">${k}</text>`;
  for (let k = 0; k <= 40; k += 10) s += `<line x1="${xr(k)}" x2="${xr(k)}" y1="${y(0)}" y2="${y(0) + 10}" stroke="${C.muted}" stroke-width="2"/><text x="${xr(k)}" y="${y(0) + 44}" text-anchor="middle" font-size="26" fill="${C.muted}">${k}</text>`;
  s += `<text x="${bx0 + bw}" y="${y(0) + 86}" text-anchor="end" font-size="24" fill="${C.muted}">自行車累積 km（2025 大會 GPX 實測 198 km；回程地標里程為估計值）</text>`;
  s += `<text x="${rx0 + rw}" y="${y(0) + 86}" text-anchor="end" font-size="24" fill="${C.muted}">跑步累積 km（依路標海拔繪製之示意剖面）</text>`;

  const bikeMarks = [
    { km: 0, label: '伊達邵 T1', sub: '06:00 出發', anchor: 'start' },
    { km: 33.5, label: '信義 7-11', sub: '約 33 km' },
    { km: 56, label: '草坪頭管制哨', sub: '56 km・紅袋' , cutoff: true },
    { km: 76, label: '觀山', sub: '約 76 km' },
    { km: 86.2, label: '塔塔加夫妻樹', sub: '86 km・黑袋・12:00 關門', cutoff: true },
    { km: 125, label: '和社 7-11', sub: '約 125 km' },
    { km: 142, label: '豐丘', sub: '約 142 km' },
    { km: 149, label: '水里坡', sub: '約 149 km' },
    { km: 165, label: '日月潭', sub: '約 165 km' },
    { km: 185, label: '埔里高工', sub: '約 185 km' },
    { km: BIKE_END, label: 'T2 檢錄', sub: '15:50 關門', cutoff: true, anchor: 'end' },
  ].map(mk => ({ ...mk, X: xb(mk.km), Y: y(eleAt(mk.km)), color: mk.cutoff ? C.red : C.bike, anchor: mk.anchor || 'middle' }));
  const runMarks = [
    { km: 0, label: 'T2', sub: '714 m', anchor: 'start', level: 0 },
    { km: 17.5, label: '清境高空步道', sub: '17.5 km・18:50', cutoff: true, level: 0 },
    { km: 26, label: '翠峰', sub: '26 km・20:20・陪跑', cutoff: true, level: 3 },
    { km: 33, label: '鳶峰', sub: '33 km・21:40・強制', cutoff: true, level: 2 },
    { km: 40, label: '登山口', sub: '40 km・23:00', cutoff: true, anchor: 'end', level: 1 },
    { km: 42, label: '合歡主峰', sub: 'FINISH 3,417 m', finish: true, anchor: 'end', level: 0 },
  ].map(mk => ({ ...mk, X: xr(mk.km), Y: y(runEleAt(mk.km)), color: mk.finish ? C.gold : (mk.cutoff ? C.red : C.run), anchor: mk.anchor || 'middle' }));
  assignLevels(bikeMarks, 28, 3); assignLevels(runMarks, 28, 4);
  const drawMarks = marks => marks.forEach(mk => {
    const ly = 120 + mk.level * 62; // label block top
    s += `<line x1="${mk.X}" x2="${mk.X}" y1="${mk.Y}" y2="${ly + 46}" stroke="${mk.color}" stroke-width="2" stroke-opacity="0.55"/>`;
    s += `<circle cx="${mk.X}" cy="${mk.Y}" r="9" fill="${C.surface}" stroke="${mk.color}" stroke-width="4"/>`;
    s += `<text x="${mk.X}" y="${ly + 10}" text-anchor="${mk.anchor}" font-size="28" font-weight="700" fill="${C.ink}">${esc(mk.label)}</text>`;
    s += `<text x="${mk.X}" y="${ly + 38}" text-anchor="${mk.anchor}" font-size="22" font-weight="${mk.cutoff || mk.finish ? 700 : 400}" fill="${mk.cutoff || mk.finish ? mk.color : C.ink2}">${esc(mk.sub)}</text>`;
  });
  drawMarks(bikeMarks); drawMarks(runMarks);
  s += `</svg>`;
  return s;
}

function mapSvg() {
  const W = 1800, H = 1500;
  const run = [[24.0039, 121.0869], [24.0230, 121.1330], [24.0560, 121.1610], [24.0900, 121.1740], [24.1030, 121.1980], [24.1190, 121.2330], [24.1330, 121.2680], [24.1380, 121.2730], [24.1416, 121.2714]];
  const lats = bike.map(p => p.lat).concat(run.map(r => r[0]));
  const lons = bike.map(p => p.lon).concat(run.map(r => r[1]));
  const lat0 = Math.min(...lats) - 0.04, lat1 = Math.max(...lats) + 0.06;
  const lon0 = Math.min(...lons) - 0.10, lon1 = Math.max(...lons) + 0.12;
  const cosl = Math.cos(((lat0 + lat1) / 2) * Math.PI / 180);
  const spanX = (lon1 - lon0) * cosl, spanY = (lat1 - lat0);
  const pad = 50;
  const scale = Math.min((W - 2 * pad) / spanX, (H - 2 * pad) / spanY);
  const px = lon => pad + (lon - lon0) * cosl * scale + ((W - 2 * pad) - spanX * scale) / 2;
  const py = lat => pad + (lat1 - lat) * scale + ((H - 2 * pad) - spanY * scale) / 2;
  let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT}">`;
  s += `<rect width="${W}" height="${H}" fill="#F6F8FA"/>`;
  for (let la = 23.5; la <= 24.3; la += 0.1) s += `<line x1="0" x2="${W}" y1="${py(la)}" y2="${py(la)}" stroke="#E6EAEF" stroke-width="2"/>`;
  for (let lo = 120.7; lo <= 121.5; lo += 0.1) s += `<line y1="0" y2="${H}" x1="${px(lo)}" x2="${px(lo)}" stroke="#E6EAEF" stroke-width="2"/>`;
  let d = ''; bike.forEach((p, i) => { d += (i ? 'L' : 'M') + `${px(p.lon).toFixed(1)},${py(p.lat).toFixed(1)} `; });
  s += `<path d="${d}" fill="none" stroke="#FFFFFF" stroke-width="16" stroke-linejoin="round" stroke-linecap="round"/><path d="${d}" fill="none" stroke="${C.bike}" stroke-width="8" stroke-linejoin="round" stroke-linecap="round"/>`;
  let dr = ''; run.forEach((p, i) => { dr += (i ? 'L' : 'M') + `${px(p[1]).toFixed(1)},${py(p[0]).toFixed(1)} `; });
  s += `<path d="${dr}" fill="none" stroke="#FFFFFF" stroke-width="16" stroke-linejoin="round" stroke-linecap="round"/><path d="${dr}" fill="none" stroke="${C.run}" stroke-width="8" stroke-linejoin="round" stroke-linecap="round" stroke-dasharray="22 14"/>`;
  const label = (lat, lon, text, sub, color, dx, dy, anchor, big) => {
    s += `<circle cx="${px(lon)}" cy="${py(lat)}" r="${big ? 16 : 12}" fill="${color}" stroke="#FFFFFF" stroke-width="5"/>`;
    const halo = 'paint-order="stroke" stroke="#F6F8FA" stroke-width="10" stroke-linejoin="round"';
    s += `<text x="${px(lon) + dx}" y="${py(lat) + dy}" text-anchor="${anchor}" font-size="${big ? 44 : 36}" font-weight="700" fill="${C.ink}" ${halo}>${esc(text)}</text>`;
    if (sub) s += `<text x="${px(lon) + dx}" y="${py(lat) + dy + 38}" text-anchor="${anchor}" font-size="28" fill="${C.ink2}" ${halo}>${esc(sub)}</text>`;
  };
  const near = km => { let b = bike[0]; for (const p of bike) if (Math.abs(p.km - km) < Math.abs(b.km - km)) b = p; return b; };
  label(23.8489, 120.9302, '日月潭 伊達邵碼頭', 'START 06:00・游泳 3.8 km', C.navy, 26, -24, 'start', true);
  label(23.5123, 120.8910, '塔塔加夫妻樹', '86 km 折返・12:00 關門・黑袋', C.red, 30, 46, 'start', true);
  const l1 = near(33.5); label(l1.lat, l1.lon, '信義 7-11', '約 33 km／回程約 139 km', C.bike, -24, 8, 'end');
  const l2 = near(56); label(l2.lat, l2.lon, '草坪頭管制哨', '56 km・紅袋・禁行區起點', C.red, -30, -10, 'end');
  const l3 = near(142); label(l3.lat, l3.lon, '豐丘', '約 142 km', C.bike, 26, 40, 'start');
  const l4 = near(149); label(l4.lat, l4.lon, '水里坡', '約 149 km', C.bike, -24, -4, 'end');
  const l5 = near(185); label(l5.lat, l5.lon, '埔里高工', '約 185 km', C.bike, -24, 44, 'end');
  label(24.0039, 121.0869, 'T2', '196.7 km・15:50 關門・檢錄', C.red, 26, 56, 'start', true);
  label(24.0560, 121.1610, '清境高空步道', '17.5 km・18:50', C.run, 30, 16, 'start');
  label(24.1030, 121.1980, '翠峰', '26 km・20:20・開放陪跑', C.run, -30, -26, 'end');
  label(24.1190, 121.2330, '鳶峰', '33 km・21:40・強制陪跑', C.run, 30, 54, 'start');
  label(24.1330, 121.2680, '昆陽・主峰登山口', '40 km・23:00', C.run, 30, -30, 'start');
  label(24.1416, 121.2714, '合歡主峰 3,417 m', 'FINISH', C.gold, -30, -40, 'end', true);
  // legend top-left
  const lx = 80, ly = 110;
  s += `<rect x="${lx - 30}" y="${ly - 50}" width="600" height="130" rx="18" fill="#FFFFFF" stroke="#E3E7EC" stroke-width="2"/>`;
  s += `<line x1="${lx}" x2="${lx + 80}" y1="${ly}" y2="${ly}" stroke="${C.bike}" stroke-width="8" stroke-linecap="round"/><text x="${lx + 100}" y="${ly + 10}" font-size="28" fill="${C.ink}">自行車 196.7 km（2025 大會 GPX 軌跡）</text>`;
  s += `<line x1="${lx}" x2="${lx + 80}" y1="${ly + 52}" y2="${ly + 52}" stroke="${C.run}" stroke-width="8" stroke-dasharray="22 14" stroke-linecap="round"/><text x="${lx + 100}" y="${ly + 62}" font-size="28" fill="${C.ink}">跑步 42 km（台14 → 台14甲，示意）</text>`;
  s += `<text x="${W - 90}" y="110" text-anchor="middle" font-size="40" font-weight="700" fill="${C.ink2}">N</text><path d="M${W - 90},120 l-14,40 l14,-12 l14,12 z" fill="${C.ink2}"/>`;
  const kmPx = (10 / 111.32) * scale;
  s += `<line x1="${W - 90 - kmPx}" x2="${W - 90}" y1="${H - 80}" y2="${H - 80}" stroke="${C.ink2}" stroke-width="6"/><text x="${W - 90 - kmPx / 2}" y="${H - 95}" text-anchor="middle" font-size="26" fill="${C.ink2}">10 km</text>`;
  s += `</svg>`;
  return s;
}


function runOnlySvg() {
  const W = 1500, H = 1080; const left = 110, right = 60, top = 330, bottom = 100;
  const pw = W - left - right, ph = H - top - bottom;
  const y = ele => top + ph - ((ele - 500) / 3100) * ph; const x = km => left + (km / 42) * pw;
  let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT}"><rect width="${W}" height="${H}" fill="${C.surface}"/>`;
  s += `<defs><linearGradient id="gr2" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="${C.run}" stop-opacity="0.40"/><stop offset="1" stop-color="${C.run}" stop-opacity="0.05"/></linearGradient></defs>`;
  for (let e = 1000; e <= 3500; e += 500) s += `<line x1="${left}" x2="${W - right}" y1="${y(e)}" y2="${y(e)}" stroke="${C.grid}" stroke-width="2"/><text x="${left - 14}" y="${y(e) + 9}" text-anchor="end" font-size="26" fill="${C.muted}">${e.toLocaleString()} m</text>`;
  let d = `M${x(0)},${y(500)} `; runPts.forEach(p => { d += `L${x(p[0]).toFixed(1)},${y(p[1]).toFixed(1)} `; }); d += `L${x(42)},${y(500)} Z`;
  s += `<path d="${d}" fill="url(#gr2)"/>`;
  let dl = ''; runPts.forEach((p, i) => { dl += (i ? 'L' : 'M') + `${x(p[0]).toFixed(1)},${y(p[1]).toFixed(1)} `; });
  s += `<path d="${dl}" fill="none" stroke="${C.run}" stroke-width="5" stroke-linejoin="round"/>`;
  s += `<line x1="${left}" x2="${W - right}" y1="${y(500)}" y2="${y(500)}" stroke="${C.ink2}" stroke-width="2"/>`;
  for (let k = 0; k <= 40; k += 5) s += `<line x1="${x(k)}" x2="${x(k)}" y1="${y(500)}" y2="${y(500) + 10}" stroke="${C.muted}" stroke-width="2"/><text x="${x(k)}" y="${y(500) + 44}" text-anchor="middle" font-size="26" fill="${C.muted}">${k}</text>`;
  s += `<text x="${W - right}" y="${y(500) + 84}" text-anchor="end" font-size="24" fill="${C.muted}">跑步累積 km（依路標海拔繪製之示意剖面）</text>`;
  const marks = [
    { km: 0, label: 'T2 全家仁愛賽德克店', sub: '約 714 m・15:50 關門', anchor: 'start', color: C.red, cutoff: true, level: 1 },
    { km: 9, label: '霧社', sub: '約 1,148 m', anchor: 'middle', color: C.run, level: 0 },
    { km: 17.5, label: '清境高空步道', sub: '18:50 關門・換 T 恤・葡萄', anchor: 'middle', color: C.red, cutoff: true, level: 2 },
    { km: 23, label: '梅峰', sub: '約 2,100 m', anchor: 'middle', color: C.run, level: 0 },
    { km: 26, label: '翠峰', sub: '20:20 關門・開放陪跑・換跑褲', anchor: 'middle', color: C.red, cutoff: true, level: 1 },
    { km: 33, label: '鳶峰', sub: '21:40 關門・強制陪跑・長褲外套', anchor: 'middle', color: C.red, cutoff: true, level: 3 },
    { km: 38, label: '昆陽', sub: '約 3,085 m・接駁車站', anchor: 'middle', color: C.run, level: 1 },
    { km: 40, label: '主峰登山口', sub: '23:00 關門', anchor: 'end', color: C.red, cutoff: true, level: 2 },
    { km: 42, label: '合歡主峰 3,417 m', sub: 'FINISH', anchor: 'end', color: C.gold, finish: true, level: 0 },
  ].map(mk => ({ ...mk, X: x(mk.km), Y: y(runEleAt(mk.km)) }));
  assignLevels(marks, 28, 4);
  marks.forEach(mk => {
    const ly = 70 + mk.level * 62;
    s += `<line x1="${mk.X}" x2="${mk.X}" y1="${mk.Y}" y2="${ly + 46}" stroke="${mk.color}" stroke-width="2" stroke-opacity="0.55"/>`;
    s += `<circle cx="${mk.X}" cy="${mk.Y}" r="9" fill="${C.surface}" stroke="${mk.color}" stroke-width="4"/>`;
    s += `<text x="${mk.X}" y="${ly + 10}" text-anchor="${mk.anchor}" font-size="28" font-weight="700" fill="${C.ink}">${esc(mk.label)}</text>`;
    s += `<text x="${mk.X}" y="${ly + 38}" text-anchor="${mk.anchor}" font-size="22" font-weight="${mk.cutoff || mk.finish ? 700 : 400}" fill="${mk.cutoff || mk.finish ? mk.color : C.ink2}">${esc(mk.sub)}</text>`;
  });
  s += `</svg>`; return s;
}

(async () => {
  await sharp(Buffer.from(runOnlySvg())).png().toFile(path.join(OUT, 'run_profile.png'));
  await sharp(Buffer.from(profileSvg())).png().toFile(path.join(OUT, 'course_profile.png'));
  await sharp(Buffer.from(mapSvg())).png().toFile(path.join(OUT, 'course_map.png'));
  console.log('charts written');
})().catch(e => { console.error(e); process.exit(1); });
