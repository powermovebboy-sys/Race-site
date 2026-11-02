// Generates local visual assets for the FXT deck: mountain backgrounds + icon PNGs.
// Run: node gen_assets.js  (from the build dir)
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const Fa = require('react-icons/fa');
const Md = require('react-icons/md');
const Gi = require('react-icons/gi');

const OUT = path.join(__dirname, 'assets');
fs.mkdirSync(OUT, { recursive: true });

// ---------- helpers ----------
function ridge(seed, baseY, amp, width, segments) {
  // deterministic pseudo-random mountain ridge polygon points
  let x = 0, pts = [];
  let s = seed;
  const rnd = () => { s = (s * 9301 + 49297) % 233280; return s / 233280; };
  const step = width / segments;
  let y = baseY - amp * rnd();
  pts.push([0, baseY - amp * 0.3]);
  for (let i = 1; i <= segments; i++) {
    x = i * step;
    const target = baseY - amp * (0.15 + 0.85 * rnd());
    // sharp peaks: alternate
    y = (i % 2 === 0) ? target : baseY - amp * (0.05 + 0.35 * rnd());
    pts.push([x, y]);
  }
  pts.push([width, baseY - amp * 0.25]);
  return pts;
}
function poly(pts, h) {
  return `M0,${h} ` + pts.map(p => `L${p[0].toFixed(1)},${p[1].toFixed(1)}`).join(' ') + ` L${pts[pts.length - 1][0]},${h} Z`;
}
function stars(seed, w, h, n) {
  let s = seed; const rnd = () => { s = (s * 9301 + 49297) % 233280; return s / 233280; };
  let out = '';
  for (let i = 0; i < n; i++) {
    const x = rnd() * w, y = rnd() * h * 0.55, r = 0.8 + rnd() * 1.8, o = 0.3 + rnd() * 0.7;
    out += `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="${r.toFixed(1)}" fill="#FFFFFF" opacity="${o.toFixed(2)}"/>`;
  }
  return out;
}

// ---------- backgrounds ----------
const W = 2560, H = 1440;

// 1) Title: night-to-dawn sky over layered mountains, sun glow low right
function bgTitle() {
  const r1 = ridge(11, H * 0.78, 300, W, 14);
  const r2 = ridge(23, H * 0.86, 230, W, 18);
  const r3 = ridge(37, H * 0.95, 160, W, 22);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
  <defs>
    <linearGradient id="sky" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#07142B"/>
      <stop offset="0.45" stop-color="#0F2A52"/>
      <stop offset="0.72" stop-color="#3B4A7A"/>
      <stop offset="0.86" stop-color="#F26B1F"/>
      <stop offset="1" stop-color="#F5B700"/>
    </linearGradient>
    <radialGradient id="glow" cx="0.78" cy="0.82" r="0.35">
      <stop offset="0" stop-color="#FFD27A" stop-opacity="0.95"/>
      <stop offset="0.35" stop-color="#F26B1F" stop-opacity="0.45"/>
      <stop offset="1" stop-color="#F26B1F" stop-opacity="0"/>
    </radialGradient>
  </defs>
  <rect width="${W}" height="${H}" fill="url(#sky)"/>
  ${stars(5, W, H, 160)}
  <rect width="${W}" height="${H}" fill="url(#glow)"/>
  <circle cx="${W * 0.78}" cy="${H * 0.80}" r="54" fill="#FFE6A8"/>
  <path d="${poly(r1, H)}" fill="#1B2F55" opacity="0.95"/>
  <path d="${poly(r2, H)}" fill="#0F1F3D"/>
  <path d="${poly(r3, H)}" fill="#07142B"/>
  </svg>`;
}

// 2) Section divider: deep navy with teal lake band and subtle ridges (calmer)
function bgSection() {
  const r1 = ridge(41, H * 0.82, 260, W, 16);
  const r2 = ridge(53, H * 0.92, 170, W, 20);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
  <defs>
    <linearGradient id="sky" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#07142B"/>
      <stop offset="0.6" stop-color="#0E2A4F"/>
      <stop offset="1" stop-color="#0E7C86"/>
    </linearGradient>
  </defs>
  <rect width="${W}" height="${H}" fill="url(#sky)"/>
  ${stars(9, W, H, 90)}
  <path d="${poly(r1, H)}" fill="#122A4E" opacity="0.9"/>
  <path d="${poly(r2, H)}" fill="#07142B"/>
  </svg>`;
}

// 3) Night summit: for the night/safety slides, cold palette with moon
function bgNight() {
  const r1 = ridge(71, H * 0.80, 320, W, 12);
  const r2 = ridge(83, H * 0.90, 200, W, 18);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
  <defs>
    <linearGradient id="sky" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0" stop-color="#030A18"/>
      <stop offset="0.7" stop-color="#0B1F3A"/>
      <stop offset="1" stop-color="#1E3A5F"/>
    </linearGradient>
    <radialGradient id="moon" cx="0.2" cy="0.22" r="0.18">
      <stop offset="0" stop-color="#DDE8FF" stop-opacity="0.35"/>
      <stop offset="1" stop-color="#DDE8FF" stop-opacity="0"/>
    </radialGradient>
  </defs>
  <rect width="${W}" height="${H}" fill="url(#sky)"/>
  ${stars(13, W, H, 220)}
  <rect width="${W}" height="${H}" fill="url(#moon)"/>
  <circle cx="${W * 0.2}" cy="${H * 0.22}" r="46" fill="#EEF3FF"/>
  <path d="${poly(r1, H)}" fill="#13284A"/>
  <path d="${poly(r2, H)}" fill="#07142B"/>
  </svg>`;
}

// 4) Light content accent: a very soft ridge strip for the bottom of light slides (used sparingly)
function bgRidgeLight() {
  const w = 2560, h = 300;
  const r = ridge(91, h * 0.95, 180, w, 20);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
  <path d="${poly(r, h)}" fill="#0E7C86" opacity="0.08"/>
  </svg>`;
}

async function svgToPng(svg, file, opts = {}) {
  const buf = Buffer.from(svg);
  let img = sharp(buf, { density: 144 });
  if (opts.resize) img = img.resize(opts.resize);
  await img.png().toFile(path.join(OUT, file));
}

// ---------- icons ----------
const ICONS = {
  swim: Fa.FaSwimmer, bike: Fa.FaBiking, run: Fa.FaRunning, mountain: Fa.FaMountain,
  car: Fa.FaCar, camera: Fa.FaCamera, flag: Fa.FaFlagCheckered, clock: Fa.FaClock,
  firstaid: Fa.FaFirstAid, headlamp: Fa.FaLightbulb, cold: Fa.FaTemperatureLow,
  moon: Fa.FaMoon, sun: Fa.FaSun, warn: Fa.FaExclamationTriangle, check: Fa.FaCheckCircle,
  pin: Fa.FaMapMarkerAlt, users: Fa.FaUsers, megaphone: Fa.FaBullhorn, bed: Fa.FaBed,
  food: Fa.FaUtensils, bolt: Fa.FaBolt, vest: Fa.FaTshirt, phone: Fa.FaPhone,
  bag: Fa.FaShoppingBag, van: Fa.FaShuttleVan, wind: Fa.FaWind, water: Fa.FaTint,
  ban: Fa.FaBan, route: Fa.FaRoute, id: Fa.FaIdBadge, box: Fa.FaBoxOpen, bottle: Fa.FaWineBottle,
  ice: Fa.FaSnowflake, bread: Fa.FaBreadSlice, coffee: Fa.FaCoffee, battery: Fa.FaBatteryFull,
  list: Fa.FaClipboardList, question: Fa.FaQuestionCircle, star: Fa.FaStar, arrow: Fa.FaArrowRight,
  medkit: Fa.FaMedkit, hotel: Fa.FaHotel, trophy: Fa.FaTrophy, bicycle: Fa.FaBicycle,
  handshake: Fa.FaHandsHelping, eye: Fa.FaEye, stop: Fa.FaHandPaper, gps: Fa.FaSatelliteDish,
  shirt: Gi.GiShirt || Fa.FaTshirt, backpack: Md.MdBackpack || Fa.FaShoppingBag,
  flashlight: Md.MdFlashlightOn || Fa.FaLightbulb, hiking: Fa.FaHiking, timer: Md.MdTimer || Fa.FaClock,
  tent: Fa.FaCampground, thermo: Fa.FaThermometerEmpty, heart: Fa.FaHeartbeat, numbers: Fa.FaHashtag,
  lake: Fa.FaWater, trees: Fa.FaTree, store: Fa.FaStore, fuel: Fa.FaGasPump, parking: Fa.FaParking,
  sunrise: Md.MdWbTwilight || Fa.FaSun, apple: Fa.FaAppleAlt, glass: Fa.FaGlassWhiskey, cup: Md.MdLocalCafe || Fa.FaCoffee,
  dollar: Fa.FaDollarSign, calendar: Fa.FaCalendarAlt, home: Fa.FaHome, map: Fa.FaMapMarkedAlt, gift: Fa.FaGift,
  tools: Fa.FaTools, lightbulb: Fa.FaLightbulb, user: Fa.FaUser, crown: Fa.FaCrown, video: Fa.FaVideo,
  drive: Fa.FaCarSide, toilet: Fa.FaRestroom, jacket: Gi.GiMonclerJacket || Fa.FaTshirt, shoe: Gi.GiRunningShoe || Fa.FaRunning,
  cap: Gi.GiBilledCap || Fa.FaHatCowboy, towel: Gi.GiTowel || Fa.FaTshirt, banana: Gi.GiBanana || Fa.FaAppleAlt,
};
const COLORS = { white: '#FFFFFF', navy: '#0B1F3A', teal: '#0E7C86', orange: '#F26B1F', gold: '#F5B700', red: '#D7263D', green: '#2A9D8F', gray: '#5B6B7F' };

async function renderIcon(name, Comp, colorName, hex) {
  if (!Comp) { console.warn('missing icon', name); return; }
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: hex, size: 512 }));
  const file = `icon_${name}_${colorName}.png`;
  await sharp(Buffer.from(svg)).resize(512, 512, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toFile(path.join(OUT, file));
}

(async () => {
  await svgToPng(bgTitle(), 'bg_title.png');
  await svgToPng(bgSection(), 'bg_section.png');
  await svgToPng(bgNight(), 'bg_night.png');
  await svgToPng(bgRidgeLight(), 'bg_ridge_light.png');
  const names = Object.keys(ICONS);
  for (const n of names) {
    for (const [cn, hex] of Object.entries(COLORS)) {
      await renderIcon(n, ICONS[n], cn, hex);
    }
  }
  console.log('assets written:', fs.readdirSync(OUT).length);
})().catch(e => { console.error(e); process.exit(1); });
