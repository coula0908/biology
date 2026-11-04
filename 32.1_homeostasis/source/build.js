// Chapter 32.1 Homeostasis — bilingual (EN + KO) presentation deck.
// Build (from this folder):  npm install  &&  npm run build
//   → writes ../32.1_Homeostasis.pptx   (or: node build.js <out.pptx>)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const gi = require("react-icons/gi");
const md = require("react-icons/md");
const tb = require("react-icons/tb");

// Usage: node build.js <full|short> [out.pptx]
const VARIANT = process.argv[2] || "main";
if (!["main", "short"].includes(VARIANT)) throw new Error(`unknown variant "${VARIANT}" (use main or short)`);
const FILE = { main: "32.1_Homeostasis.pptx", short: "32.1_Homeostasis_short.pptx" }[VARIANT];
const OUT = path.resolve(process.argv[3] || path.join(__dirname, "..", FILE));
const NOTES = require(`./notes_${VARIANT}.js`);
function NOTE(id) {
  if (!NOTES[id]) throw new Error(`no ${VARIANT} speaker notes for slide "${id}"`);
  return NOTES[id];
}
let CURRENT_SECTION = null; // set from the variant's slide list
const OUTLINE = []; // [{ id, title, kr }] for the script export

const HEX = { ink: "1D2733", navy: "12304A", light: "EEF3F6", coral: "E2553B", teal: "167D8C", amber: "F2A541", slate: "5C6B7A", ice: "8ED1DC", green: "3E8E5E", white: "FFFFFF" };
const THEME = {
  name: "Thermal Balance",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: { dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.light, accent1: HEX.coral, accent2: HEX.teal, accent3: HEX.amber, accent4: HEX.slate, accent5: HEX.ice, accent6: HEX.green, hlink: HEX.teal, folHlink: HEX.slate },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "32.1 Homeostasis: Why and How Do Animals Regulate Their Internal Environment?";
pres.subject = "General Biology — Chapter 32.1 (bilingual EN/KO)";
const C = pres.SchemeColor;
const S = pres.shapes;

// ---------------------------------------------------------------- assets
const P = (...a) => path.join(__dirname, "assets", ...a);
const PHOTO = {
  pupfish: { f: P("photos/desert_pupfish.jpg"), credit: "Photo: P. V. Loiselle, CC BY-SA 3.0, via Wikimedia Commons" },
  humming: { f: P("photos/hummingbird.jpg"), credit: "Photo: Charles J. Sharp, CC BY-SA 4.0, via Wikimedia Commons" },
  iguana: { f: P("photos/green_iguana.jpg"), credit: "Photo: Cayambe, CC BY-SA 3.0, via Wikimedia Commons" },
  bear: { f: P("photos/polar_bear.jpg"), credit: "Photo: Andreas Weith, CC BY-SA 4.0, via Wikimedia Commons" },
  bernard: { f: P("photos/claude_bernard.jpg"), credit: "Public domain, via Wikimedia Commons" },
  cannon: { f: P("photos/walter_cannon.jpg"), credit: "Bachrach (1934), CC BY 4.0, via Wikimedia Commons" },
};
const FIG = {
  ap110: { f: P("figures/openstax_ap2e_fig1-10.jpg"), credit: "Figure: OpenStax, Anatomy and Physiology 2e, Fig. 1.10 (CC BY 4.0)" },
  ap111: { f: P("figures/openstax_ap2e_fig1-11.jpg"), credit: "Figure: OpenStax, Anatomy and Physiology 2e, Fig. 1.11 (CC BY 4.0)" },
  bio3323: { f: P("figures/openstax_bio2e_fig33-23.jpg"), credit: "Figure: OpenStax, Biology 2e, Fig. 33.23 (CC BY 4.0)" },
  // open form of yeast hexokinase (left half of the PDB-101 figure; white background made transparent)
  hexokinase: { f: P("figures/hexokinase_open_pdb101.png"), credit: "Hexokinase image: David S. Goodsell & RCSB PDB, Molecule of the Month (doi:10.2210/rcsb_pdb/mom_2004_2), CC BY 4.0" },
  // npseaice_ssm_1984258.png + npseaice_am2_2012257.png side by side (NASA Earth Observatory, image of the day 79256)
  nasaArctic: { f: P("figures/nasa_arctic_sea_ice_1984_2012.png"), credit: "September sea-ice minimum · NASA Earth Observatory (public domain)" },
};

const iconCache = new Map();
async function icon(Comp, hex, px = 256) {
  const key = (Comp.name || Math.random()) + hex + px;
  if (iconCache.has(key)) return iconCache.get(key);
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: px }));
  const buf = await sharp(Buffer.from(svg)).resize(px, px).png().toBuffer();
  const d = "image/png;base64," + buf.toString("base64");
  iconCache.set(key, d);
  return d;
}

async function photo(file, wIn, hIn, { pos = "centre", radius = 0.14, bg = HEX.white, px = 1200 } = {}) {
  const W = Math.round(px), H = Math.round((px * hIn) / wIn);
  const base = sharp(file).rotate().toColorspace("srgb").resize(W, H, { fit: "cover", position: pos });
  if (radius > 0) {
    const r = Math.round((radius / wIn) * W);
    const mask = Buffer.from(`<svg width="${W}" height="${H}"><rect x="0" y="0" width="${W}" height="${H}" rx="${r}" ry="${r}" fill="#fff"/></svg>`);
    const rounded = await base.composite([{ input: mask, blend: "dest-in" }]).png().toBuffer();
    const out = await sharp(rounded).flatten({ background: "#" + bg }).jpeg({ quality: 85 }).toBuffer();
    return "image/jpeg;base64," + out.toString("base64");
  }
  const out = await base.jpeg({ quality: 85 }).toBuffer();
  return "image/jpeg;base64," + out.toString("base64");
}

async function figure(file) {
  const out = await sharp(file).toColorspace("srgb").jpeg({ quality: 92 }).toBuffer();
  const m = await sharp(out).metadata();
  return { data: "image/jpeg;base64," + out.toString("base64"), ar: m.width / m.height };
}

async function svgPng(svg, px) {
  const buf = await sharp(Buffer.from(svg)).resize(px).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------------------------------------------------------------- helpers
// Rough text width (inches) for Calibri / Malgun Gothic, used to size chips.
function tw(str, pt, bold = false) {
  let em = 0;
  for (const ch of str) {
    const c = ch.codePointAt(0);
    if ((c >= 0xac00 && c <= 0xd7a3) || (c >= 0x3130 && c <= 0x318f)) em += 0.93;
    else if (/[A-Z]/.test(ch)) em += 0.6;
    else if (/[a-z]/.test(ch)) em += 0.48;
    else if (/[0-9]/.test(ch)) em += 0.51;
    else if (ch === " ") em += 0.23;
    else em += 0.36;
  }
  return ((em * pt) / 72) * (bold ? 1.07 : 1);
}

const TXT = { valign: "top", margin: 0, isTextBox: true, color: C.text1, fontSize: 16 };
function text(s, runs, x, y, w, h, o = {}) {
  s.addText(runs, Object.assign({ x, y, w, h }, TXT, o));
}
// bullets: items are strings or arrays of {text, options}
function bl(items, { gap = 6, indent = 14, size } = {}) {
  const out = [];
  items.forEach((it, i) => {
    const runs = typeof it === "string" ? [{ text: it }] : it;
    runs.forEach((r, j) => {
      const o = Object.assign({}, r.options || {});
      if (size) o.fontSize = o.fontSize || size;
      if (j === 0) { o.bullet = { indent }; o.paraSpaceAfter = gap; }
      if (j === runs.length - 1 && i < items.length - 1) o.breakLine = true;
      out.push({ text: r.text, options: o });
    });
  });
  return out;
}
const b = (t, o = {}) => ({ text: t, options: Object.assign({ bold: true }, o) });
const t = (t, o = {}) => ({ text: t, options: o });
const ko = (k, o = {}) => ({ text: k, options: Object.assign({ color: C.accent4 }, o) });

function card(s, x, y, w, h, { fill = C.background2, line, r = 0.12, transparency, shadow = false, name } = {}) {
  const o = { x, y, w, h, rectRadius: r, fill: { color: fill }, line: line ? { color: line, width: 1.25 } : { type: "none" } };
  if (transparency !== undefined) o.fill.transparency = transparency;
  if (shadow) o.shadow = { type: "outer", color: "000000", opacity: 0.14, blur: 8, offset: 2, angle: 90 };
  if (name) o.objectName = name;
  s.addShape(S.ROUNDED_RECTANGLE, o);
}

function chip(s, x, y, en, kr, { col = C.accent2, size = 14, h = 0.38, fillT = 87, enCol, w } = {}) {
  const width = w || tw(en, size, true) + (kr ? tw("  " + kr, size) : 0) + 0.34;
  const runs = [{ text: en, options: { bold: true, color: enCol || col } }];
  if (kr) runs.push({ text: "  " + kr, options: { color: C.text1 } });
  s.addText(runs, { shape: S.ROUNDED_RECTANGLE, x, y, w: width, h, rectRadius: h / 2, fill: { color: col, transparency: fillT }, line: { color: col, width: 1 }, fontSize: size, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "term chip: " + en });
  return width;
}

function tag(s, label, { col = C.text2, txt = C.background1 } = {}) {
  const w = tw(label, 11, true) * 1.12 + 0.42;
  s.addText(label, { shape: S.ROUNDED_RECTANGLE, x: 0.6, y: 0.3, w, h: 0.3, rectRadius: 0.15, fill: { color: col }, line: { type: "none" }, fontSize: 11, bold: true, color: txt, charSpacing: 1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "section tag" });
}

async function iconCircle(s, x, y, d, Comp, { bg = C.accent2, fg = HEX.white, scale = 0.56, name } = {}) {
  s.addShape(S.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" }, objectName: name ? name + " circle" : undefined });
  const iw = d * scale;
  s.addImage({ data: await icon(Comp, fg), x: x + (d - iw) / 2, y: y + (d - iw) / 2, w: iw, h: iw, altText: name || "icon" });
}
async function ic(s, Comp, hex, x, y, d, alt) {
  s.addImage({ data: await icon(Comp, hex), x, y, w: d, h: d, altText: alt || "icon" });
}

function arrow(s, x1, y1, x2, y2, { col = C.accent4, w = 2.25, head = "triangle", dash, name } = {}) {
  const ln = { color: col, width: w };
  if (head) ln.endArrowType = head;
  if (dash) ln.dashType = dash;
  s.addShape(S.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), flipH: x2 < x1, flipV: y2 < y1, line: ln, objectName: name });
}
// polyline through points; arrowhead on the last segment only
function polyline(s, pts, o = {}) {
  for (let i = 1; i < pts.length; i++) arrow(s, pts[i - 1][0], pts[i - 1][1], pts[i][0], pts[i][1], Object.assign({}, o, { head: i === pts.length - 1 ? (o.head === undefined ? "triangle" : o.head) : null }));
}

// smooth curve through points (inches, relative to box origin) using horizontal-tangent cubics
function curve(s, x, y, w, h, pts, line) {
  const P2 = [{ x: pts[0][0], y: pts[0][1] }];
  for (let i = 1; i < pts.length; i++) {
    const [x0, y0] = pts[i - 1], [x1, y1] = pts[i], dx = x1 - x0;
    P2.push({ x: x1, y: y1, curve: { type: "cubic", x1: x0 + dx / 2, y1: y0, x2: x1 - dx / 2, y2: y1 } });
  }
  s.addShape(S.CUSTOM_GEOMETRY, { x, y, w, h, points: P2, line });
}

function credit(s, str, x, y, w, { col = C.accent4, align = "left" } = {}) {
  text(s, str, x, y, w, 0.24, { fontSize: 10, color: col, italic: true, align, objectName: "image credit" });
}

function contentSlide(_section, tagLabel, tagOpts, title, kr, master = "CONTENT") {
  const s = pres.addSlide({ masterName: master, sectionTitle: CURRENT_SECTION });
  OUTLINE.push({ title, kr });
  if (tagLabel) tag(s, tagLabel, tagOpts);
  s.addText(title, { placeholder: "title" });
  s.addText(kr, { placeholder: "kr" });
  return s;
}

// ---------------------------------------------------------------- layouts
const footer = { text: { text: "32.1 Homeostasis · 항상성", options: { x: 0.6, y: 7.02, w: 5, h: 0.28, fontSize: 10, color: C.accent4, margin: 0, valign: "middle" } } };
pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: HEX.navy },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.75, y: 2.35, w: 7.6, h: 2.3, fontSize: 38, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "kr", type: "body", x: 0.75, y: 4.8, w: 7.6, h: 0.9, fontSize: 20, color: C.accent5, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.white },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.66, w: 12.13, h: 0.64, fontSize: 36, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "kr", type: "body", x: 0.6, y: 1.3, w: 12.13, h: 0.38, fontSize: 16, color: C.accent4, align: "left", valign: "top", margin: 0 }, text: "" } },
    footer,
  ],
  slideNumber: { x: 12.23, y: 7.02, w: 0.5, h: 0.28, fontSize: 10, color: HEX.slate, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENT_DARK",
  background: { color: HEX.navy },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.66, w: 12.13, h: 0.64, fontSize: 36, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "kr", type: "body", x: 0.6, y: 1.3, w: 12.13, h: 0.38, fontSize: 16, color: C.accent5, align: "left", valign: "top", margin: 0 }, text: "" } },
    { text: { text: "32.1 Homeostasis · 항상성", options: { x: 0.6, y: 7.02, w: 5, h: 0.28, fontSize: 10, color: C.accent5, margin: 0, valign: "middle" } } },
  ],
  slideNumber: { x: 12.23, y: 7.02, w: 0.5, h: 0.28, fontSize: 10, color: HEX.ice, align: "right" },
});

// tag presets
const T_CASE = { col: C.accent3, txt: C.text1 };
const T_PART = { col: C.text2, txt: C.background1 };
const T_WOND = { col: C.accent2, txt: C.background1 };
const T_EARTH = { col: C.accent6, txt: C.background1 };
const T_REV = { col: C.accent4, txt: C.background1 };

// ================================================================ slides
const SL = {};

SL.title = async () => {
    const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: CURRENT_SECTION });
    OUTLINE.push({ title: "Homeostasis: Why and How Do Animals Regulate Their Internal Environment?", kr: "항상성: 동물은 왜 그리고 어떻게 자신의 내부 환경을 조절하는가?" });
    text(s, "GENERAL BIOLOGY  ·  CHAPTER 32", 0.75, 1.55, 7, 0.35, { fontSize: 14, bold: true, color: C.accent3, charSpacing: 2 });
    text(s, "32.1", 0.75, 1.0, 3, 0.6, { fontSize: 30, bold: true, color: C.accent1 });
    s.addText("Homeostasis: Why and How Do Animals Regulate Their Internal Environment?", { placeholder: "title" });
    // line break placed by the presenter in PowerPoint
    s.addText([{ text: "항상성: 동물은 왜 그리고 어떻게" }, { text: "자신의 내부 환경을 조절하는가?", options: { softBreakBefore: true } }], { placeholder: "kr" });
    // thermal ring motif: cold half + hot half around a thermometer
    const cx = 10.55, cy = 3.75, D = 3.6;
    s.addShape(S.BLOCK_ARC, { x: cx - D / 2, y: cy - D / 2, w: D, h: D, angleRange: [95, 265], arcThicknessRatio: 0.16, fill: { color: C.accent5 }, line: { type: "none" }, objectName: "ring cold half" });
    s.addShape(S.BLOCK_ARC, { x: cx - D / 2, y: cy - D / 2, w: D, h: D, angleRange: [275, 85], arcThicknessRatio: 0.16, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "ring hot half" });
    await ic(s, fa.FaThermometerHalf, HEX.white, cx - 0.75, cy - 0.95, 1.5, "thermometer");
    text(s, "set point", cx - 1.0, cy + 0.62, 2.0, 0.35, { fontSize: 14, color: C.background1, align: "center", italic: true });
    await ic(s, fa.FaSnowflake, HEX.ice, cx - 1.55, cy + D / 2 + 0.2, 0.4, "cold");
    text(s, "cold", cx - 1.1, cy + D / 2 + 0.2, 0.8, 0.4, { fontSize: 14, color: C.accent5, valign: "middle" });
    await ic(s, fa.FaFire, HEX.coral, cx + 0.55, cy + D / 2 + 0.2, 0.4, "hot");
    text(s, "hot", cx + 1.0, cy + D / 2 + 0.2, 0.8, 0.4, { fontSize: 14, color: C.accent1, valign: "middle" });
    s.addNotes(NOTE("title"));
};

SL.case = async () => {
    const s = contentSlide("Opening", "CASE STUDY", T_CASE, "Case Study: Overheated", "과도한 열 — 열사병(heat stroke)으로 쓰러진 대학 운동선수들");
    card(s, 0.6, 1.95, 4.75, 4.7, { fill: C.accent3, transparency: 82, r: 0.16 });
    await ic(s, fa.FaQuoteLeft, HEX.amber, 0.9, 2.2, 0.55, "quote");
    text(s, "“I’m having the race of my life. … Nothing is going to get in my way, not the 90° heat and 90% humidity, … not the fire in my feet that is spreading up my legs.”", 0.9, 2.9, 4.2, 2.3, { fontSize: 17, italic: true, color: C.text1 });
    text(s, [b("Kierann Smith", { breakLine: true }), t("medical student, Stanford triathlon team — just before she collapsed (2011 USA Triathlon Collegiate Nationals, Alabama)", { fontSize: 13, color: C.accent4 })], 0.9, 5.3, 4.2, 1.2, { fontSize: 15 });

    const rows = [
      { big: "41.1°C", f: "106°F", who: "Kierann Smith · triathlete", what: "Core body temperature at collapse → ice water + ice packs brought it just below 40°C (104°F). She survived.", col: C.accent3 },
      { big: "41.6°C", f: "107°F", who: "Ben Richards · wrestler, 20", what: "Collapsed during a 5-mile training run (2013); died 10 days later.", col: C.accent1 },
      { big: "2014", f: "cause: heat stroke", who: "Marquese Meadow · football player, freshman", what: "Disoriented after practice; died in a nearby hospital. Autopsy confirmed heat stroke as the cause of death.", col: C.accent1 },
    ];
    let y = 1.95;
    for (const r of rows) {
      card(s, 5.65, y, 7.08, 1.42, { fill: C.background2, r: 0.14 });
      text(s, r.big, 5.85, y + 0.16, 1.9, 0.72, { fontSize: 34, bold: true, color: r.col, fontFace: "+mj-lt", align: "center", valign: "middle" });
      text(s, r.f, 5.85, y + 0.88, 1.9, 0.35, { fontSize: 14, color: C.accent4, align: "center" });
      text(s, [b(r.who, { breakLine: true, color: C.text2 }), t(r.what, { fontSize: 14 })], 7.95, y + 0.17, 4.6, 1.15, { fontSize: 16, valign: "middle" });
      y += 1.64;
    }
    s.addNotes(NOTE("case"));
};

SL.roadmap = async () => {
    const s = contentSlide("Opening", "OVERVIEW", T_REV, "Today’s Roadmap", "오늘 다룰 내용 — 교재의 학습 목표(Check Your Learning) 네 가지");
    const items = [
      // kr lines: line breaks placed by the presenter in PowerPoint
      { n: "1", en: "What is homeostasis — and why do organisms maintain it?", kr: ["항상성이란 무엇이고,", "왜 유지할까?"], I: fa.FaBalanceScale, col: C.text2 },
      { n: "2", en: "How do ectotherms and endotherms differ?", kr: ["외온동물과 내온동물은", "어떻게 다를까?"], I: fa.FaThermometerHalf, col: C.accent3 },
      { n: "3", en: "Why is negative feedback crucial for homeostasis?", kr: ["음성피드백은", "왜 중요할까?"], I: fa.FaSyncAlt, col: C.accent2 },
      { n: "4", en: "What is positive feedback — and why is it rare?", kr: ["양성피드백이란", "무엇이고, 왜 드물까?"], I: tb.TbArrowBigUpLines, col: C.accent1 },
    ];
    const w = 2.86, gap = 0.23;
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + i * (w + gap), it = items[i];
      card(s, x, 2.0, w, 3.85, { fill: C.background2, r: 0.16 });
      await iconCircle(s, x + 0.3, 2.3, 0.9, it.I, { bg: it.col, name: "part " + it.n });
      text(s, "PART " + it.n, x + 1.35, 2.48, 1.4, 0.5, { fontSize: 16, bold: true, color: it.col, valign: "middle" });
      text(s, it.en, x + 0.3, 3.45, w - 0.55, 1.5, { fontSize: 17, bold: true, color: C.text1 });
      text(s, it.kr.map((line, k) => ({ text: line, options: k > 0 ? { softBreakBefore: true } : {} })), x + 0.3, 5.0, w - 0.55, 0.7, { fontSize: 15, color: C.accent4 });
    }
    card(s, 0.6, 6.12, 12.13, 0.62, { fill: C.accent3, transparency: 80, r: 0.31 });
    await ic(s, fa.FaRunning, HEX.ink, 0.88, 6.2, 0.46, "athlete");
    text(s, [b("Running thread: "), t("the heat-stroke case (열사병 사례) — what keeps body temperature constant, and what goes wrong.")], 1.5, 6.12, 11.0, 0.62, { fontSize: 15, valign: "middle" });
    s.addNotes(NOTE("roadmap"));
};

SL.constancy = async () => {
    const s = contentSlide("1. What & Why", "PART 1 · WHAT & WHY", T_PART, "From Internal Constancy to Homeostasis", "‘내부 불변성’에서 ‘항상성’으로");
    // left: two observations
    card(s, 0.6, 1.95, 6.25, 2.25, { fill: C.background2 });
    await iconCircle(s, 0.85, 2.2, 0.75, fa.FaThermometerHalf, { bg: C.accent1, name: "temperature" });
    text(s, [t("Sitting in your room, sweltering in a jungle, or shivering in a blizzard — ", { color: C.accent4 }), t("most cells of your body (heart, brain, muscles) maintain an "), b("almost constant temperature", { color: C.accent1 }), t(".")], 1.8, 2.15, 4.85, 1.95, { fontSize: 18 });
    card(s, 0.6, 4.4, 6.25, 2.3, { fill: C.background2 });
    await iconCircle(s, 0.85, 4.65, 0.75, fa.FaTint, { bg: C.accent2, name: "fluid" });
    text(s, [t("Parched desert air, a pool, the ocean, or the Great Salt Lake — ", { color: C.accent4 }), t("your cells are bathed in "), b("interstitial fluid", { color: C.accent2 }), t(" (세포간질액) with an "), b("almost constant composition"), t(".")], 1.8, 4.6, 4.85, 2.0, { fontSize: 18 });
    // right: two portraits on a timeline
    const px = [7.25, 10.1], pw = 2.6, ph = 2.75;
    const pb = await photo(PHOTO.bernard.f, pw, ph, { pos: "north", px: 700 });
    const pc = await photo(PHOTO.cannon.f, pw, ph, { pos: "north", px: 700 });
    s.addImage({ data: pb, x: px[0], y: 1.95, w: pw, h: ph, altText: "Portrait of Claude Bernard" });
    s.addImage({ data: pc, x: px[1], y: 1.95, w: pw, h: ph, altText: "Portrait of Walter B. Cannon" });
    text(s, [b("Claude Bernard", { breakLine: true, color: C.text2 }), t("French physiologist, mid-19th century", { fontSize: 13, color: C.accent4, breakLine: true }), t("described the “internal constancy” ", { fontSize: 14 }), t("(내부 불변성)", { fontSize: 14, color: C.accent4 })], px[0], 4.82, pw + 0.05, 1.5, { fontSize: 16 });
    text(s, [b("Walter Cannon", { breakLine: true, color: C.text2 }), t("1920s", { fontSize: 13, color: C.accent4, breakLine: true }), t("coined the term ", { fontSize: 14 }), b("homeostasis", { fontSize: 14, color: C.accent2 }), t(" (항상성)", { fontSize: 14, color: C.accent4 })], px[1], 4.82, pw + 0.05, 1.5, { fontSize: 16 });
    credit(s, "Portraits: Bernard — public domain; Cannon — Bachrach 1934, CC BY 4.0 (Wikimedia Commons)", 7.25, 6.45, 5.5);
    s.addNotes(NOTE("constancy"));
};

SL.what = async () => {
    const s = contentSlide("1. What & Why", "PART 1 · WHAT & WHY", T_PART, "What Is Homeostasis?", "항상성이란 무엇인가?");
    card(s, 0.6, 1.95, 12.13, 1.75, { fill: C.text2, r: 0.16 });
    text(s, [b("Homeostasis ", { color: C.accent5 }), t("(항상성)", { color: C.accent5 }), t(" — the ability of an organism to maintain its internal environment within the narrow limits that allow optimal cell functioning.", { color: C.background1 })], 0.95, 2.1, 11.4, 0.95, { fontSize: 21, valign: "middle" });
    text(s, "유기체가 세포의 적절한 기능을 위해 필요한 조건의 범위 속에서 내부 환경을 유지하려는 능력 (Walter Cannon, 1920s)", 0.95, 3.07, 11.4, 0.45, { fontSize: 15, color: C.accent5 });
    text(s, "Homeostatic mechanisms regulate …", 0.6, 3.95, 8, 0.4, { fontSize: 16, bold: true, color: C.text2 });
    const tiles = [
      { en: "Temperature", kr: "체온", I: fa.FaThermometerHalf, col: C.accent1 },
      { en: "Water & salt in body fluids", kr: "체액의 물과 염 농도", I: fa.FaTint, col: C.accent2 },
      { en: "Glucose", kr: "포도당 농도", I: gi.GiCubes, col: C.accent3 },
      { en: "pH (acid–base balance)", kr: "산-염기의 균형", I: fa.FaFlask, col: C.accent6 },
      { en: "Hormone secretion", kr: "호르몬 분비", I: fa.FaVial, col: C.text2 },
      { en: "O₂ and CO₂", kr: "산소와 이산화탄소 농도", I: fa.FaWind, col: C.accent4 },
    ];
    const w = 1.86, gap = 0.194;
    for (let i = 0; i < tiles.length; i++) {
      const x = 0.6 + i * (w + gap), it = tiles[i];
      card(s, x, 4.45, w, 2.3, { fill: C.background2 });
      await iconCircle(s, x + (w - 0.8) / 2, 4.65, 0.8, it.I, { bg: it.col, name: it.en });
      text(s, it.en, x + 0.1, 5.55, w - 0.2, 0.62, { fontSize: 15, bold: true, align: "center", valign: "middle" });
      text(s, it.kr, x + 0.1, 6.17, w - 0.2, 0.45, { fontSize: 13, color: C.accent4, align: "center", valign: "top" });
    }
    s.addNotes(NOTE("what"));
};

SL.dynamic = async () => {
    const s = contentSlide("1. What & Why", "PART 1 · WHAT & WHY", T_PART, "Homeostasis Is Dynamic, Not Static", "말뜻은 ‘같은 상태로 머무르다’이지만, 내부 환경은 끊임없이 조절된다");
    // left concept
    card(s, 0.6, 1.95, 3.9, 4.75, { fill: C.background2, r: 0.16 });
    text(s, [t("“homeostasis”", { bold: true, color: C.accent2, fontSize: 22, breakLine: true }), t("= “to stay the same”", { fontSize: 17, breakLine: true }), t("(같은 상태로 머무르다)", { fontSize: 14, color: C.accent4 })], 0.85, 2.15, 3.45, 1.35, {});
    text(s, [t("The word implies a static, unchanging state — but the internal environment actually "), b("seethes with activity", { color: C.accent1 }), t(" as the body continuously adjusts to maintain constancy amid varying internal and external conditions.")], 0.85, 3.6, 3.45, 2.0, { fontSize: 15 });
    chip(s, 0.85, 5.95, "dynamic equilibrium", "동적인 평형상태", { col: C.accent2, size: 13, h: 0.42 });
    // right flow: exercise -> demands -> responses
    await iconCircle(s, 4.95, 3.55, 1.25, fa.FaRunning, { bg: C.accent1, name: "exercise" });
    text(s, [b("Exercise", { breakLine: true }), t("운동", { color: C.accent4, fontSize: 14 })], 4.75, 4.85, 1.65, 0.65, { fontSize: 16, align: "center" });
    text(s, "challenges homeostasis", 4.65, 5.5, 1.85, 0.5, { fontSize: 13, italic: true, color: C.accent4, align: "center" });
    const dem = [
      [t("More O₂ must be supplied to sustain the energy demands of working muscles")],
      [t("The extra heat that muscles produce must be dissipated")],
    ];
    const dy = [2.35, 4.45];
    for (let i = 0; i < 2; i++) {
      card(s, 6.85, dy[i], 2.75, 1.6, { fill: C.accent1, transparency: 86 });
      text(s, dem[i], 7.05, dy[i] + 0.12, 2.4, 1.36, { fontSize: 15, valign: "middle" });
      arrow(s, 6.25, 4.18, 6.82, dy[i] + 0.8, { col: C.accent1 });
    }
    text(s, "Demands", 6.85, 1.95, 2.75, 0.35, { fontSize: 14, bold: true, color: C.accent1 });
    text(s, "The body rises to the occasion", 10.0, 1.95, 2.75, 0.35, { fontSize: 14, bold: true, color: C.accent2 });
    const res = [
      { en: "↑ Respiratory rate", kr: "호흡수 증가", I: fa.FaLungs },
      { en: "↑ Sweat-gland activity", kr: "땀샘 활동 증가", I: fa.FaTint },
      { en: "↑ Blood flow to the skin", kr: "피부로의 혈류 증가", I: fa.FaHeartbeat },
    ];
    for (let i = 0; i < 3; i++) {
      const y = 2.35 + i * 1.42;
      card(s, 10.0, y, 2.73, 1.22, { fill: C.accent2, transparency: 86 });
      await iconCircle(s, 10.15, y + 0.3, 0.62, res[i].I, { bg: C.accent2, name: res[i].en });
      text(s, [b(res[i].en, { breakLine: true }), t(res[i].kr, { color: C.accent4, fontSize: 13 })], 10.88, y + 0.1, 1.8, 1.02, { fontSize: 14, valign: "middle" });
    }
    arrow(s, 9.65, 3.15, 9.95, 3.15, { col: C.accent2 });
    arrow(s, 9.65, 5.25, 9.95, 5.25, { col: C.accent2 });
    text(s, "Interacting systems throughout the body help maintain homeostasis.", 6.85, 6.42, 5.9, 0.35, { fontSize: 13, italic: true, color: C.accent4 });
    s.addNotes(NOTE("dynamic"));
};

SL.enzymes = async () => {
    const s = contentSlide("1. What & Why", "PART 1 · WHAT & WHY", T_PART, "Homeostasis Allows Enzymes to Function", "항상성은 효소가 기능하도록 한다 — 세포가 주변 환경에 민감한 이유");
    // Left: structure
    card(s, 0.6, 1.95, 6.35, 4.0, { fill: C.background2, r: 0.16 });
    text(s, [b("① Structure ", { color: C.accent2 }), t("(구조)", { color: C.accent4 })], 0.85, 2.1, 5.9, 0.42, { fontSize: 19 });
    // explicit line breaks keep each Korean gloss on the line of its English term (widths measured)
    text(s, [t("Almost every reaction in a cell is catalyzed by a specific", { breakLine: true }), b("enzyme"), t(" (효소). Its function depends on its precise", { breakLine: true }), b("three-dimensional structure"), t(", maintained in part by", { breakLine: true }), b("hydrogen bonds", { color: C.accent2 }), t(" (수소결합).")], 0.85, 2.55, 5.9, 1.15, { fontSize: 15 });
    const broken = await svgPng(ENZ_BROKEN, 700);
    s.addImage({ path: FIG.hexokinase.f, x: 0.95, y: 3.75, w: 2.45, h: 1.75, altText: "Hexokinase enzyme (open form) with an arrow to its active-site cleft" });
    s.addImage({ data: broken, x: 4.15, y: 3.75, w: 2.45, h: 1.75, altText: "Unfolded enzyme with disrupted hydrogen bonds" });
    arrow(s, 3.5, 4.62, 4.05, 4.62, { col: C.accent1, w: 2.5 });
    text(s, "works (hexokinase)", 0.95, 5.5, 2.45, 0.3, { fontSize: 13, bold: true, color: C.accent2, align: "center" });
    text(s, "bonds disrupted", 4.15, 5.5, 2.45, 0.3, { fontSize: 13, bold: true, color: C.accent1, align: "center" });
    // stressor chips (bottom row, left)
    const st = [["too salty", "염분 과다"], ["too acidic", "강산성"], ["too basic", "강염기성"], ["too hot", "고온"]];
    let cx = 0.6;
    for (const [en, k] of st) cx += chip(s, cx, 6.2, en, k, { col: C.accent1, size: 12, h: 0.42 }) + 0.1;
    // Right: rate
    card(s, 7.2, 1.95, 5.53, 4.0, { fill: C.background2, r: 0.16 });
    text(s, [b("② Rate ", { color: C.accent1 }), t("(반응 속도)", { color: C.accent4 })], 7.45, 2.1, 5.1, 0.42, { fontSize: 19 });
    text(s, [t("The rate at which enzymes catalyze reactions is "), b("temperature-dependent"), t(".")], 7.45, 2.55, 5.1, 0.75, { fontSize: 15 });
    // temperature bar: cold -> hot
    const bx = 7.6, by = 3.55, bw = 4.75;
    s.addShape(S.ROUNDED_RECTANGLE, { x: bx, y: by, w: bw / 2, h: 0.32, rectRadius: 0.16, fill: { color: C.accent5 }, line: { type: "none" } });
    s.addShape(S.ROUNDED_RECTANGLE, { x: bx + bw / 2, y: by, w: bw / 2, h: 0.32, rectRadius: 0.16, fill: { color: C.accent1 }, line: { type: "none" } });
    s.addShape(S.RECTANGLE, { x: bx + bw / 2 - 0.2, y: by, w: 0.4, h: 0.32, fill: { color: C.accent3 }, line: { type: "none" } });
    await ic(s, fa.FaSnowflake, HEX.teal, bx - 0.05, by + 0.42, 0.38, "low temperature");
    await ic(s, fa.FaFire, HEX.coral, bx + bw - 0.33, by + 0.42, 0.38, "high temperature");
    text(s, [b("Low temperature", { color: C.accent2, breakLine: true }), t("slows molecular motion → fewer and slower interactions between enzymes and their molecules")], 7.45, 4.4, 2.5, 1.5, { fontSize: 14 });
    text(s, [b("Rising temperature", { color: C.accent1, breakLine: true }), t("reactions speed up — sometimes to a detrimental level")], 10.2, 4.4, 2.4, 1.5, { fontSize: 14 });
    // conclusion (bottom row, right)
    const ccx = cx + 0.1;
    card(s, ccx, 6.2, 12.73 - ccx, 0.42, { fill: C.text2, r: 0.21 });
    text(s, [t("→ keep ", { color: C.background1 }), b("salt, pH & temperature", { color: C.accent3 }), t(" in narrow ranges", { color: C.background1 })], ccx + 0.1, 6.2, 12.73 - ccx - 0.2, 0.42, { fontSize: 13, valign: "middle", align: "center" });
    credit(s, FIG.hexokinase.credit, 0.6, 6.68, 12.13);
    s.addNotes(NOTE("enzymes"));
};

SL.case2 = async () => {
    const s = contentSlide("1. What & Why", "CASE STUDY · CONTINUED", T_CASE, "Case Study: Why Didn’t They Recover?", "과도한 열 — 체온은 몇 시간 만에 정상이 되었는데, 왜 회복하지 못했을까?");
    text(s, [t("Meadow and Richards were treated in hospital for over a week, and their body temperatures were brought down to normal "), b("within a few hours"), t(". Yet both died.")], 0.6, 1.92, 12.13, 0.7, { fontSize: 16 });
    // flow
    card(s, 0.6, 2.85, 2.7, 2.6, { fill: C.accent1, r: 0.16 });
    await ic(s, fa.FaThermometerFull, HEX.white, 1.6, 3.05, 0.7, "heat stroke");
    text(s, [b("Heat stroke", { breakLine: true, color: C.background1 }), t("열사병", { color: C.background1, fontSize: 14, breakLine: true }), t("→ metabolic chaos", { color: C.background1, fontSize: 14, breakLine: true }), t("(신진대사의 이상)", { color: C.background1, fontSize: 13 })], 0.75, 3.85, 2.4, 1.5, { fontSize: 18, align: "center" });
    // branch A: muscle -> kidney
    arrow(s, 3.35, 3.6, 3.95, 3.2, { col: C.accent1 });
    arrow(s, 3.35, 4.7, 3.95, 5.1, { col: C.accent1 });
    card(s, 4.0, 2.55, 4.3, 1.45, { fill: C.background2 });
    await iconCircle(s, 4.15, 2.85, 0.85, gi.GiMuscleUp, { bg: C.accent4, name: "muscle" });
    text(s, [b("Muscle cells die", { breakLine: true }), t("and release substances into the bloodstream")], 5.15, 2.62, 3.05, 1.3, { fontSize: 15, valign: "middle" });
    arrow(s, 8.35, 3.27, 8.95, 3.27, { col: C.accent1 });
    card(s, 9.0, 2.55, 3.73, 1.45, { fill: C.accent1, transparency: 85 });
    await iconCircle(s, 9.15, 2.85, 0.85, gi.GiKidneys, { bg: C.accent1, name: "kidneys" });
    text(s, [b("Kidney failure", { breakLine: true }), t("신부전", { color: C.accent4, fontSize: 14 })], 10.15, 2.62, 2.5, 1.3, { fontSize: 16, valign: "middle" });
    // branch B: temp -> liver
    card(s, 4.0, 4.4, 4.3, 1.45, { fill: C.background2 });
    await iconCircle(s, 4.15, 4.7, 0.85, fa.FaThermometerFull, { bg: C.accent4, name: "elevated temperature" });
    text(s, [b("Elevated body temperature", { breakLine: true }), t("also damages the liver")], 5.15, 4.47, 3.05, 1.3, { fontSize: 15, valign: "middle" });
    arrow(s, 8.35, 5.12, 8.95, 5.12, { col: C.accent1 });
    card(s, 9.0, 4.4, 3.73, 1.45, { fill: C.accent1, transparency: 85 });
    await iconCircle(s, 9.15, 4.7, 0.85, gi.GiLiver, { bg: C.accent1, name: "liver" });
    text(s, [b("Liver damage", { breakLine: true }), t("간 손상 — the liver is crucial for removing toxic substances from the blood", { color: C.accent4, fontSize: 13 })], 10.15, 4.47, 2.5, 1.3, { fontSize: 16, valign: "middle" });
    // bottom question
    card(s, 0.6, 6.08, 12.13, 0.68, { fill: C.text2, r: 0.34 });
    text(s, [t("Human metabolism functions only within a narrow range of temperatures. ", { color: C.background1 }), b("Is this true for all animals?", { color: C.accent3 }), t("  → Part 2", { color: C.accent5 })], 0.9, 6.08, 11.6, 0.68, { fontSize: 16, valign: "middle" });
    s.addNotes(NOTE("case2"));
};

SL.warmcold = async () => {
    const s = contentSlide("2. Body Temperature", "PART 2 · BODY TEMPERATURE", T_PART, "“Warm-Blooded” or “Cold-Blooded”?", "온혈동물일까, 냉혈동물일까? — 이 용어들은 오해를 부르기 쉽다");
    text(s, [t("Mammals and birds are often called "), b("“warm-blooded”", { color: C.accent1 }), t(" (온혈); other reptiles, amphibians, fish, and invertebrates", { breakLine: true }), b("“cold-blooded”", { color: C.accent2 }), t(" (냉혈). But these terms are often "), b("misleading"), t(":")], 0.6, 1.92, 12.13, 0.75, { fontSize: 16 });
    const cards = [
      { p: PHOTO.pupfish, x: 0.6, title: "Desert pupfish", kr: "사막 pupfish — ‘냉혈’ 물고기", stat: "> 37.8°C", statF: "over 100°F", desc: "when the summer sun heats the springs and small ponds where it lives — so pupfish are sometimes quite “warm-blooded.”", col: C.accent1, pos: "centre" },
      { p: PHOTO.humming, x: 6.77, title: "Hummingbird", kr: "벌새 — ‘온혈’ 조류", stat: "40.5 → 13°C", statF: "105°F by day → 55°F at night", desc: "as high as 40.5°C while foraging during the day, but may cool down to 13°C during the night.", col: C.accent2, pos: "centre" },
    ];
    for (const c of cards) {
      const w = 5.96;
      card(s, c.x, 2.8, w, 3.92, { fill: C.background2, r: 0.16 });
      text(s, [b(c.title, { color: C.text2, fontSize: 19 }), t("   " + c.kr, { color: C.accent4, fontSize: 14 })], c.x + 0.25, 2.9, w - 0.5, 0.5, { valign: "middle" });
      const img = await photo(c.p.f, 3.0, 2.0, { pos: c.pos, bg: HEX.light, px: 900 });
      s.addImage({ data: img, x: c.x + 0.25, y: 3.48, w: 3.0, h: 2.0, altText: c.title });
      text(s, c.stat, c.x + 3.4, 3.62, 2.45, 0.75, { fontSize: c.stat.length > 8 ? 27 : 32, bold: true, color: c.col, fontFace: "+mj-lt", valign: "middle" });
      text(s, c.statF, c.x + 3.4, 4.38, 2.4, 0.6, { fontSize: 13, color: C.accent4 });
      text(s, c.desc, c.x + 0.25, 5.6, w - 0.5, 1.0, { fontSize: 15 });
      credit(s, c.p.credit, c.x, 6.74, w);
    }
    s.addNotes(NOTE("warmcold"));
};

SL.ectoendo = async () => {
    const s = contentSlide("2. Body Temperature", "PART 2 · BODY TEMPERATURE", T_PART, "Ectotherms vs. Endotherms", "외온동물과 내온동물 — 열을 주로 어디서 얻는가에 따른 분류");
    text(s, "To avoid confusion, scientists classify animals by their major source of body warmth.", 0.6, 1.92, 12.13, 0.4, { fontSize: 16, color: C.text1 });
    const cols = [
      { x: 0.6, col: C.accent2, colHex: HEX.teal, title: "ECTOTHERMS", kr: "외온동물", greek: "Greek: “outside heat” (외부열)", I: fa.FaSun, src: [t("Derive most of their heat", { breakLine: true }), t("from the "), b("environment"), t(".")], ex: [["Reptiles (except birds)", gi.GiGecko], ["Amphibians", gi.GiFrog], ["Most fish", fa.FaFish], ["Most invertebrates", gi.GiButterfly]], note: "Simplest cases — an earthworm in its burrow or a fish in a stream: body temperature = environment temperature." },
      { x: 6.77, col: C.accent1, colHex: HEX.coral, title: "ENDOTHERMS", kr: "내온동물", greek: "Greek: “inside heat” (내부열)", I: fa.FaFire, src: [t("Produce most of their heat", { breakLine: true }), t("by "), b("metabolic reactions"), t(" (대사반응).")], ex: [["Birds", gi.GiHummingbird], ["Mammals", fa.FaPaw]], note: "Birds and mammals are the most common endotherms; some other animals occasionally generate significant heat metabolically." },
    ];
    for (const c of cols) {
      const w = 5.96;
      card(s, c.x, 2.5, w, 4.25, { fill: c.col, transparency: 88, r: 0.16 });
      await iconCircle(s, c.x + 0.3, 2.72, 1.0, c.I, { bg: c.col, name: c.title });
      text(s, [b(c.title, { color: c.col, fontSize: 24, breakLine: true }), t(c.kr + "  ·  " + c.greek, { fontSize: 14, color: C.accent4 })], c.x + 1.5, 2.72, 4.3, 1.0, { valign: "middle" });
      text(s, c.src, c.x + 0.3, 3.9, w - 0.6, 0.75, { fontSize: 16 });
      let ex = c.x + 0.55;
      for (const [name, I] of c.ex) {
        await iconCircle(s, ex, 4.75, 0.55, I, { bg: C.background1, fg: c.colHex, scale: 0.7, name });
        text(s, name, ex + 0.275 - 0.66, 5.33, 1.32, 0.55, { fontSize: 12, align: "center", color: C.text1 });
        ex += 1.38;
      }
      text(s, c.note, c.x + 0.3, 5.95, w - 0.6, 0.75, { fontSize: 14, italic: true, color: C.text1 });
    }
    s.addNotes(NOTE("ectoendo"));
};

SL.ectocycle = async () => {
    const s = contentSlide("2. Body Temperature", "PART 2 · BODY TEMPERATURE", T_PART, "Ectotherms: Cooling Down, Warming Up", "외온동물의 체온은 시간·일·주 단위의 외부 기온 변화에 따라 크게 달라진다");
    text(s, [t("Many ectotherms — "), b("butterflies, bees, and most lizards"), t(" — cool down and become inactive at night, "), b("conserving energy", { color: C.accent2 }), t(", then warm up and become active during the day.")], 0.6, 1.92, 12.13, 0.75, { fontSize: 16 });
    const st = [
      { I: fa.FaMoon, bg: C.text2, h: "NIGHT", kr: "밤", body: [b("Butterflies and bees"), t(" become so cool that they cannot fly."), t("", { breakLine: true }), b("Lizards"), t(" are often too sluggish to hunt or escape predators.")], fill: C.text2 },
      { I: md.MdWbTwilight, bg: C.accent3, h: "DAYBREAK", kr: "새벽", body: [b("Bees shiver"), t(" and "), b("butterflies beat their wings"), t(" to generate metabolic heat (대사열)."), t("", { breakLine: true }), b("Lizards"), t(" seek a warm, sunlit stone.")], fill: C.accent3 },
      { I: fa.FaSun, bg: C.accent1, h: "DAY", kr: "낮", body: [t("Warmed up, they "), b("become active during the day"), t(" — the heat they gained lets them "), b("resume their active lifestyles"), t(".")], fill: C.accent1 },
    ];
    const w = 3.75, gap = 0.44;
    for (let i = 0; i < 3; i++) {
      const x = 0.6 + i * (w + gap), it = st[i];
      card(s, x, 2.9, w, 3.85, { fill: it.fill, transparency: 87, r: 0.16 });
      await iconCircle(s, x + 0.3, 3.12, 0.95, it.I, { bg: it.bg, name: it.h });
      text(s, [b(it.h, { fontSize: 20, color: C.text2, breakLine: true }), t(it.kr, { fontSize: 14, color: C.accent4 })], x + 1.4, 3.15, 2.2, 0.9, { valign: "middle" });
      text(s, it.body, x + 0.3, 4.3, w - 0.55, 2.35, { fontSize: 18 });
      if (i < 2) arrow(s, x + w + 0.06, 4.8, x + w + gap - 0.06, 4.8, { col: C.accent4, w: 2.5 });
    }
    await ic(s, gi.GiButterfly, HEX.slate, 3.62, 3.15, 0.5, "butterfly");
    await ic(s, gi.GiBee, HEX.slate, 7.8, 3.15, 0.5, "bee");
    await ic(s, gi.GiGecko, HEX.slate, 12.0, 3.15, 0.5, "lizard");
    s.addNotes(NOTE("ectocycle"));
};

SL.ectoreg = async () => {
    const s = contentSlide("2. Body Temperature", "PART 2 · BODY TEMPERATURE", T_PART, "Ectotherms Can Still Regulate Temperature", "외온동물도 행동으로, 또는 매우 일정한 환경 덕분에 체온을 안정적으로 유지한다");
    const img = await photo(PHOTO.iguana.f, 4.6, 3.95, { pos: "centre", px: 1000 });
    s.addImage({ data: img, x: 0.6, y: 1.95, w: 4.6, h: 3.95, altText: "Green iguana basking on a rock" });
    text(s, [b("Basking in the sun", { color: C.text2 }), t(" — a behavioral mechanism", { breakLine: true }), t("햇볕 쬐기: 외온동물이 체온을 조절하는 행동 메커니즘", { color: C.accent4, fontSize: 12 })], 0.6, 5.98, 4.6, 0.5, { fontSize: 13 });
    credit(s, PHOTO.iguana.credit, 0.6, 6.5, 4.6);
    const rx = 5.55, rw = 7.18;
    // card A behaviour
    card(s, rx, 1.95, rw, 1.35, { fill: C.background2 });
    await iconCircle(s, rx + 0.25, 2.22, 0.8, fa.FaSun, { bg: C.accent3, name: "behavior" });
    text(s, [b("Behavior  ", { color: C.text2 }), ko("행동", { fontSize: 14, breakLine: true }), t("Insects and lizards often "), b("warm up by basking"), t(" in the sun and "), b("cool off by moving to a shady spot"), t(".")], rx + 1.25, 2.02, rw - 1.45, 1.2, { fontSize: 15, valign: "middle" });
    // card B pupfish with range bar
    card(s, rx, 3.45, rw, 1.75, { fill: C.background2 });
    await iconCircle(s, rx + 0.25, 3.7, 0.8, fa.FaFish, { bg: C.accent2, name: "pupfish" });
    text(s, [b("Desert pupfish  ", { color: C.text2 }), t("tolerates water of "), b("2.2–45°C"), t(" (36–113°F), but can "), b("breed only within a narrow range"), t(" → swims to different areas of its pool as water temperature changes.")], rx + 1.25, 3.52, rw - 1.45, 1.0, { fontSize: 15 });
    const bx = rx + 1.25, by = 4.68, bw = rw - 1.65;
    s.addShape(S.ROUNDED_RECTANGLE, { x: bx, y: by, w: bw, h: 0.2, rectRadius: 0.1, fill: { color: C.accent2, transparency: 65 }, line: { type: "none" }, objectName: "tolerance range" });
    s.addShape(S.ROUNDED_RECTANGLE, { x: bx + bw * 0.52, y: by - 0.05, w: bw * 0.14, h: 0.3, rectRadius: 0.1, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "breeding range (schematic)" });
    text(s, "2.2°C", bx - 0.05, by + 0.22, 0.8, 0.26, { fontSize: 11, color: C.accent4 });
    text(s, "45°C", bx + bw - 0.6, by + 0.22, 0.65, 0.26, { fontSize: 11, color: C.accent4, align: "right" });
    text(s, "breeding (narrow)", bx + bw * 0.4, by + 0.22, 1.8, 0.26, { fontSize: 11, color: C.accent1, bold: true, align: "center" });
    // card C deep sea
    card(s, rx, 5.35, rw, 1.4, { fill: C.background2 });
    await iconCircle(s, rx + 0.25, 5.65, 0.8, md.MdWaves, { bg: C.text2, name: "deep sea" });
    text(s, [b("Very constant environment  ", { color: C.text2 }), ko("매우 일정한 환경", { fontSize: 14, breakLine: true }), t("In the deep ocean the temperature stays at about "), b("0–3°C"), t(" (32–37.5°F), so ectothermic deep-sea fish experience little variation in body temperature throughout their lives.")], rx + 1.25, 5.4, rw - 1.45, 1.3, { fontSize: 15, valign: "middle" });
    s.addNotes(NOTE("ectoreg"));
};

SL.endocost = async () => {
    const s = contentSlide("2. Body Temperature", "PART 2 · BODY TEMPERATURE", T_PART, "Endotherms: Benefits and Costs", "내온동물 — 따뜻한 체온을 유지하는 데 따르는 이점과 비용");
    card(s, 0.6, 1.95, 4.1, 4.8, { fill: C.text2, r: 0.16 });
    // content block vertically centred in the dark card (1.95–6.75)
    text(s, "Most endotherms keep a fairly constant body temperature of about", 0.9, 2.8, 3.55, 0.65, { fontSize: 16, color: C.background1 });
    text(s, "35–41°C", 0.9, 3.45, 3.55, 0.95, { fontSize: 48, bold: true, color: C.accent3, fontFace: "+mj-lt", valign: "middle" });
    text(s, "(95–106°F)", 0.9, 4.4, 3.55, 0.4, { fontSize: 15, color: C.accent5 });
    text(s, [t("A few, such as ", { color: C.background1 }), b("hummingbirds", { color: C.accent5 }), t(", allow their body temperature to fall at night.", { color: C.background1 })], 0.9, 5.05, 3.55, 0.9, { fontSize: 15 });
    // balance scale
    const fx = 8.7;
    s.addShape(S.ISOSCELES_TRIANGLE, { x: fx - 0.45, y: 5.9, w: 0.9, h: 0.75, fill: { color: C.accent4 }, line: { type: "none" }, objectName: "fulcrum" });
    s.addShape(S.LINE, { x: 5.4, y: 5.92, w: 6.6, h: 0, line: { color: C.accent4, width: 4 } });
    card(s, 5.1, 1.95, 4.2, 3.75, { fill: C.accent2, transparency: 86, r: 0.16 });
    text(s, [b("BENEFITS", { color: C.accent2, fontSize: 20 }), t("  이점", { color: C.accent4, fontSize: 15 })], 5.35, 2.1, 3.8, 0.5, {});
    text(s, "A warm body usually can …", 5.35, 2.65, 3.8, 0.4, { fontSize: 15, italic: true, color: C.accent4 });
    const ben = [[fa.FaEye, "sense its environment better"], [fa.FaBolt, "respond more quickly"], [fa.FaRunning, "move faster than a cold body"]];
    for (let i = 0; i < 3; i++) {
      await iconCircle(s, 5.4, 3.18 + i * 0.8, 0.58, ben[i][0], { bg: C.accent2, name: ben[i][1] });
      text(s, ben[i][1], 6.15, 3.18 + i * 0.8, 3.05, 0.58, { fontSize: 16, valign: "middle" });
    }
    card(s, 9.55, 2.9, 3.18, 2.8, { fill: C.accent1, transparency: 86, r: 0.16 });
    text(s, [b("COST", { color: C.accent1, fontSize: 20 }), t("  비용", { color: C.accent4, fontSize: 15 })], 9.8, 3.05, 2.8, 0.5, {});
    await iconCircle(s, 9.8, 3.7, 0.75, fa.FaFire, { bg: C.accent1, name: "energy cost" });
    text(s, [t("the "), b("energy"), t(" required to maintain a high body temperature")], 9.8, 4.55, 2.75, 1.05, { fontSize: 16 });
    arrow(s, 7.2, 5.72, 7.2, 5.9, { col: C.accent4, w: 2, head: null });
    arrow(s, 11.14, 5.72, 11.14, 5.9, { col: C.accent4, w: 2, head: null });
    s.addNotes(NOTE("endocost"));
};

SL.twotypes = async () => {
    const s = contentSlide("3. Negative Feedback", "PART 3 · NEGATIVE FEEDBACK", T_PART, "Two Types of Feedback Systems", "피드백 시스템은 내부 조건을 조절한다");
    text(s, [t("How does an animal sense the conditions within its body and adjust them when necessary? Its internal environment is maintained by "), b("feedback systems", { color: C.text2 }), t(" (피드백 시스템).")], 0.6, 1.92, 12.13, 0.75, { fontSize: 16 });
    const cs = [
      { x: 0.6, col: C.accent2, title: "Negative feedback", kr: "음성피드백", def: [t("counteracts the effects of changes in the internal environment → "), b("principally responsible for maintaining homeostasis")], pts: [[0, 1.0], [0.75, 0.25], [1.5, 1.45], [2.25, 0.65], [3.0, 1.2], [3.75, 0.9], [4.5, 1.02], [5.0, 1.0]], label: "returns toward the set point" },
      { x: 6.77, col: C.accent1, title: "Positive feedback", kr: "양성피드백", def: [t("creates cycles in which "), b("changes amplify themselves")], pts: null, label: "change keeps growing" },
    ];
    for (const c of cs) {
      const w = 5.96;
      card(s, c.x, 2.85, w, 3.9, { fill: c.col, transparency: 89, r: 0.16 });
      text(s, [b(c.title, { color: c.col, fontSize: 22 }), t("  " + c.kr, { color: C.accent4, fontSize: 16 })], c.x + 0.3, 3.0, w - 0.6, 0.5, {});
      text(s, c.def, c.x + 0.3, 3.55, w - 0.6, 0.9, { fontSize: 16 });
      // mini schematic graph
      const gx = c.x + 0.45, gy = 4.75, gw = 5.0, gh = 1.6;
      s.addShape(S.LINE, { x: gx, y: gy + gh, w: gw, h: 0, line: { color: C.accent4, width: 1 } });
      s.addShape(S.LINE, { x: gx, y: gy, w: 0, h: gh, line: { color: C.accent4, width: 1 } });
      s.addShape(S.LINE, { x: gx, y: gy + 1.0, w: gw, h: 0, line: { color: C.accent4, width: 1.25, dashType: "dash" } });
      text(s, "set point", gx + gw - 1.0, gy + 1.02, 1.0, 0.28, { fontSize: 11, color: C.accent4, italic: true, align: "right" });
      if (c.pts) curve(s, gx, gy, gw, gh, c.pts, { color: C.accent2, width: 3 });
      else s.addShape(S.CUSTOM_GEOMETRY, { x: gx, y: gy, w: gw, h: gh, points: [{ x: 0, y: 1.0 }, { x: 5.0, y: 0.02, curve: { type: "cubic", x1: 2.8, y1: 1.0, x2: 4.4, y2: 0.75 } }], line: { color: C.accent1, width: 3 } });
      text(s, c.label, gx, gy - 0.32, 3.2, 0.28, { fontSize: 12, color: c.col, bold: true });
      text(s, "time →", gx + gw - 0.8, gy + gh + 0.02, 0.8, 0.25, { fontSize: 11, color: C.accent4, align: "right" });
      text(s, "schematic graph (개념도)", gx, gy + gh + 0.02, 2.4, 0.25, { fontSize: 10, italic: true, color: C.accent4 });
    }
    s.addNotes(NOTE("twotypes"));
};

SL.negfb = async () => {
    const s = contentSlide("3. Negative Feedback", "PART 3 · NEGATIVE FEEDBACK", T_PART, "Negative Feedback Reverses Change", "음성피드백은 변화의 효과를 역전시킨다 — 항상성을 조절하는 가장 중요한 메커니즘");
    text(s, [b("Negative feedback", { color: C.accent2 }), t(" (음성피드백): a change causes responses that "), b("counteract the change"), t(" → the system returns to its original condition.")], 0.6, 1.95, 6.4, 1.05, { fontSize: 17 });
    text(s, "Found in living and nonliving systems:", 0.6, 3.05, 6.4, 0.35, { fontSize: 15, bold: true, color: C.text2 });
    text(s, bl([
      [b("Body: "), t("body temperature; blood levels of glucose, hormones, water, salts, and oxygen; even standing upright")],
      [b("Machines: "), t("house temperature, filling a toilet tank, cruise control")],
    ], { gap: 4 }), 0.6, 3.42, 6.4, 1.3, { fontSize: 15 });
    text(s, "All negative feedback systems contain three principal components:", 0.6, 4.75, 6.4, 0.35, { fontSize: 15, bold: true, color: C.text2 });
    const comp = [
      { en: "Sensor", kr: "감지기", d: "detects the current condition", I: md.MdSensors },
      { en: "Control center", kr: "조절 센터", d: "compares it to the set point (설정점)", I: fa.FaBrain },
      { en: "Effector", kr: "효과기", d: "produces an output that restores the condition", I: fa.FaCogs },
    ];
    for (let i = 0; i < 3; i++) {
      const y = 5.17 + i * 0.55;
      await iconCircle(s, 0.6, y, 0.46, comp[i].I, { bg: C.accent2, name: comp[i].en });
      text(s, [b(comp[i].en, { color: C.accent2 }), t(" " + comp[i].kr, { color: C.accent4 }), t("  " + comp[i].d)], 1.2, y, 5.85, 0.46, { fontSize: 14, valign: "middle" });
    }
    const f = await figure(FIG.ap110.f);
    const fw = 5.4, fh = fw / f.ar;
    card(s, 7.2, 1.95, 5.53, 4.8, { fill: C.background1, line: C.background2, r: 0.12 });
    s.addImage({ data: f.data, x: 7.2 + (5.53 - fw) / 2, y: 2.05, w: fw, h: fh, altText: "OpenStax A&P Figure 1.10: negative feedback loop and body temperature regulation" });
    text(s, "(a) 음성피드백 고리 · (b) 체온 조절 — 체온이 37°C를 넘을 때", 7.35, 2.1 + fh, 5.3, 0.3, { fontSize: 12, color: C.accent4, align: "center" });
    credit(s, FIG.ap110.credit, 7.35, 2.42 + fh, 5.3, { align: "center" });
    s.addNotes(NOTE("negfb"));
};

SL.fig322 = async () => {
    const s = contentSlide("3. Negative Feedback", "PART 3 · NEGATIVE FEEDBACK", T_PART, "Negative Feedback in Action (Fig. 32-2)", "음성피드백은 항상성을 유지한다 — 집의 온도와 체온 조절");
    const Lx = 1.55, Lw = 3.05, Cx = 5.0, Cw = 3.35, Rx = 8.75, Rw = 3.98;
    const mid = Cx + Cw / 2;
    // column headers
    // header positions as moved by the presenter in PowerPoint
    await ic(s, fa.FaHome, HEX.slate, 2.072, 1.86, 0.36, "house");
    text(s, [b("House", { color: C.text2 }), t("  집", { color: C.accent4 })], 2.518, 1.86, 2.4, 0.36, { fontSize: 15, valign: "middle" });
    await ic(s, fa.FaChild, HEX.slate, 10.065, 1.86, 0.36, "body");
    text(s, [b("Body", { color: C.text2 }), t("  몸", { color: C.accent4 })], 10.475, 1.86, 2.4, 0.36, { fontSize: 15, valign: "middle" });
    await ic(s, fa.FaSnowflake, HEX.teal, mid - 1.15, 1.86, 0.34, "cold");
    text(s, [b("stimulus: cold", { color: C.accent2 }), t("  자극: 추위", { color: C.accent4 })], mid - 0.75, 1.86, 2.6, 0.34, { fontSize: 14, valign: "middle" });
    arrow(s, mid, 2.22, mid, 2.42, { col: C.accent2, w: 2.5 });
    // condition ellipse
    s.addText([b("CONDITION", { color: C.text2 }), t("  조건", { color: C.accent4, breakLine: true }), t("room or body temperature", { fontSize: 12, color: C.text1 })], { shape: S.OVAL, x: Cx + 0.25, y: 2.42, w: Cw - 0.5, h: 0.62, fill: { color: C.accent5, transparency: 45 }, line: { color: C.accent2, width: 1.25 }, fontSize: 14, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "condition" });
    const rows = [
      { y: 3.42, en: "SENSOR", kr: "감지기", L: [fa.FaThermometerHalf, "thermometer", "온도계"], R: [fa.FaHandPaper, "temperature receptors in skin", "피부의 온도 수용체"] },
      { y: 4.42, en: "CONTROL CENTER", kr: "설정점을 가진 조절 센터", en2: "with SET POINT", L: [md.MdThermostat, "thermostat", "온도조절기"], R: [fa.FaBrain, "hypothalamus", "시상하부"] },
      { y: 5.52, en: "EFFECTOR", kr: "효과기", L: [gi.GiFurnace, "furnace output: heat", "난방 장치: 열"], R: [gi.GiMuscleUp, "muscle output: shivering → heat", "근육: 떨림 → 열"] },
    ];
    const rh = 0.6;
    for (const r of rows) {
      const runs = [b(r.en, { color: C.background1 })];
      if (r.en2) runs.push(b(" " + r.en2, { color: C.background1 }));
      runs[runs.length - 1].options.breakLine = true;
      runs.push(t(r.kr, { color: C.accent5, fontSize: 12 }));
      s.addText(runs, { shape: S.ROUNDED_RECTANGLE, x: Cx, y: r.y, w: Cw, h: rh, rectRadius: 0.1, fill: { color: C.text2 }, line: { type: "none" }, fontSize: 14, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: r.en });
      for (const [side, X, W] of [["L", Lx, Lw], ["R", Rx, Rw]]) {
        const [I, en, k] = r[side];
        card(s, X, r.y, W, rh, { fill: C.background2, r: 0.1 });
        await iconCircle(s, X + 0.1, r.y + 0.08, 0.44, I, { bg: side === "L" ? C.accent4 : C.accent2, name: en });
        text(s, [b(en, { breakLine: true }), t(k, { color: C.accent4, fontSize: 12 })], X + 0.65, r.y + 0.02, W - 0.72, rh - 0.04, { fontSize: 13, valign: "middle" });
        // dotted connector to the center box
        if (side === "L") arrow(s, X + W + 0.03, r.y + rh / 2, Cx - 0.03, r.y + rh / 2, { col: C.accent4, w: 1, head: null, dash: "sysDot" });
        else arrow(s, Cx + Cw + 0.03, r.y + rh / 2, X - 0.03, r.y + rh / 2, { col: C.accent4, w: 1, head: null, dash: "sysDot" });
      }
    }
    // down arrows + labels
    const lbl = [
      { y1: 3.04, y2: 3.42, t1: "measured by", t2: "측정" },
      { y1: 4.02, y2: 4.42, t1: "sends data to", t2: "결과 전송" },
      { y1: 5.02, y2: 5.52, t1: "if sensor and set point differ, activates", t2: "감지기가 측정한 온도와 설정점이 다르면 작동", two: true },
    ];
    for (const l of lbl) {
      arrow(s, mid, l.y1, mid, l.y2, { col: C.accent2, w: 2.5 });
      const runs = l.two ? [t(l.t1, { italic: true, color: C.text1, breakLine: true }), t(l.t2, { color: C.accent4 })] : [t(l.t1, { italic: true, color: C.text1 }), t("  " + l.t2, { color: C.accent4 })];
      text(s, runs, mid + 0.15, l.y1 + 0.02, 4.6, l.y2 - l.y1 - 0.04, { fontSize: 12, valign: "middle" });
    }
    // return path: effector -> bottom -> left -> up -> condition
    const by = 6.36, lx = 0.95, cy = 2.73;
    polyline(s, [[mid, 5.52 + rh], [mid, by], [lx, by], [lx, cy], [Cx + 0.22, cy]], { col: C.accent2, w: 2.75 });
    text(s, [b("Effector changes the conditions to match set point", { color: C.accent2 }), t("  효과기는 설정점을 맞추기 위해 조건을 변경", { color: C.accent4 })], 1.15, by + 0.06, 8.6, 0.34, { fontSize: 12, valign: "middle" });
    s.addNotes(NOTE("fig322"));
};

SL.hypo = async () => {
    const s = contentSlide("3. Negative Feedback", "PART 3 · NEGATIVE FEEDBACK", T_PART, "Hypothalamus: The Body’s Thermostat", "시상하부의 체온 조절 — 체온이 설정점(정상 체온)보다 낮을 때와 높을 때");
    text(s, [b("Hypothalamus", { color: C.text2 }), t(" (시상하부) — a region deep in the brain; the key homeostatic regulator in vertebrates. In birds and mammals it coordinates "), b("nervous signals, hormone release, and behaviors"), t(". Nerve endings in the skin and elsewhere act as temperature sensors.")], 0.6, 1.88, 12.13, 0.82, { fontSize: 15 });
    const cw = 3.62;
    // cold panel
    card(s, 0.6, 2.85, cw, 3.55, { fill: C.accent5, transparency: 60, r: 0.16 });
    text(s, [b("Below set point", { color: C.text2, fontSize: 17, breakLine: true }), t("generate & conserve heat", { color: C.accent2, bold: true })], 0.8, 2.95, cw - 0.4, 0.75, { fontSize: 14 });
    text(s, bl([
      [b("Shivering"), t(" (떨림) — small, rapid muscular contractions")],
      [t("Moving to a warmer place")],
      [t("Vessels to "), b("nonvital areas"), t(" (face, hands, feet, skin) "), b("constrict"), t(" → less heat loss; warm blood diverted to the body core")],
    ], { gap: 4, indent: 12 }), 0.8, 3.72, cw - 0.35, 2.6, { fontSize: 14 });
    // hot panel
    card(s, 13.33 - 0.6 - cw, 2.85, cw, 3.55, { fill: C.accent1, transparency: 85, r: 0.16 });
    const hx = 13.33 - 0.6 - cw + 0.2;
    text(s, [b("Above set point", { color: C.text2, fontSize: 17, breakLine: true }), t("reduce body temperature", { color: C.accent1, bold: true })], hx, 2.95, cw - 0.4, 0.75, { fontSize: 14 });
    text(s, bl([
      [t("More blood "), b("diverted to the skin"), t(" → heat radiated to the surrounding air")],
      [b("Sweat glands"), t(" (땀샘): watery secretion cools the body as it evaporates")],
      [t("Fatigue & discomfort → rest, seek shade or cooling water")],
    ], { gap: 4, indent: 12 }), hx, 3.72, cw - 0.35, 2.6, { fontSize: 14 });
    // figure in the middle
    const f = await figure(FIG.bio3323.f);
    const fw = 4.4, fh = fw / f.ar, fx = (13.333 - fw) / 2;
    s.addImage({ data: f.data, x: fx, y: 2.85 + (3.55 - fh) / 2 - 0.12, w: fw, h: fh, altText: "OpenStax Biology 2e Figure 33.23: body temperature falls vs rises" });
    credit(s, FIG.bio3323.credit, fx - 0.1, 6.1, fw + 0.2, { align: "center" });
    card(s, 0.6, 6.52, 12.13, 0.38, { fill: C.text2, r: 0.19 });
    text(s, [t("When normal body temperature is restored, the sensors signal the hypothalamus → it ", { color: C.background1 }), b("switches off", { color: C.accent3 }), t(" those actions (negative feedback).", { color: C.background1 })], 0.85, 6.52, 11.7, 0.38, { fontSize: 13, valign: "middle" });
    s.addNotes(NOTE("hypo"));
};

SL.water = async () => {
    const s = contentSlide("3. Negative Feedback", "HAVE YOU EVER WONDERED …", T_WOND, "Can You Drink Too Much Water?", "물도 너무 많이 마시면 해로울까?");
    card(s, 0.6, 1.95, 4.0, 4.8, { fill: C.accent2, r: 0.16 });
    await ic(s, fa.FaTint, HEX.white, 1.0, 2.25, 0.8, "water drop");
    text(s, "13%", 1.0, 3.15, 3.3, 1.1, { fontSize: 66, bold: true, color: C.background1, fontFace: "+mj-lt" });
    text(s, "of Boston Marathon runners tested at the finish line had at least mild hyponatremia — a few had dangerously low blood Na⁺.", 1.0, 4.3, 3.3, 1.9, { fontSize: 15, color: C.background1 });
    text(s, [t("Too much of almost anything — "), b("even water"), t(" — can be harmful. “Hydration” has become a buzzword that sometimes encourages people to drink considerably more than they need.")], 4.95, 1.95, 7.78, 0.95, { fontSize: 16 });
    card(s, 4.95, 3.0, 7.78, 1.5, { fill: C.background2 });
    text(s, [b("Zyrees Oliver", { color: C.text2 }), t(", 17, a Georgia high-school football player, reportedly drank "), b("2 gallons of water + 2 gallons of Gatorade"), t(" during practice to relieve muscle cramps — he collapsed at home and died.")], 5.2, 3.08, 7.3, 1.35, { fontSize: 15, valign: "middle" });
    chip(s, 4.95, 4.68, "hyponatremia", "저나트륨혈증 — “too little sodium”", { col: C.accent1, size: 15, h: 0.44 });
    text(s, [t("Excess water dilutes the sodium ions (Na⁺, 나트륨이온) in the blood. Na⁺ is essential for "), b("neuronal signaling and muscle contraction"), t(", so elaborate homeostatic mechanisms keep it within a very narrow range.")], 4.95, 5.25, 7.78, 0.9, { fontSize: 15 });
    text(s, [b("Symptoms: ", { color: C.accent1 }), t("nausea and vomiting, muscle cramps, convulsions, unconsciousness, coma, occasionally death")], 4.95, 6.2, 7.78, 0.55, { fontSize: 14 });
    s.addNotes(NOTE("water"));
};

SL.vicious = async () => {
    const s = contentSlide("4. Positive Feedback", "CASE STUDY · CONTINUED", T_CASE, "Case Study: When Feedback Fails", "열사병 — 음성피드백 메커니즘이 제대로 작동하지 않는다");
    // left: broken negative feedback
    card(s, 0.6, 1.95, 5.2, 4.8, { fill: C.background2, r: 0.16 });
    text(s, [b("Negative feedback breaks down", { color: C.text2, breakLine: true }), t("음성피드백의 오작동", { color: C.accent4, fontSize: 14 })], 0.85, 2.1, 4.75, 0.8, { fontSize: 18 });
    const fails = [
      { I: fa.FaBed, h: "Exhaustion should make you rest …", d: "but heat stroke causes mental confusion, which disrupts this feedback — athletes ignore the body’s warning signs." },
      { I: fa.FaTint, h: "Sweating may stop", d: "skyrocketing body temperature makes the brain’s temperature control systems malfunction." },
    ];
    for (let i = 0; i < 2; i++) {
      const y = 3.05 + i * 1.75;
      await iconCircle(s, 0.85, y + 0.05, 0.7, fails[i].I, { bg: C.accent4, name: fails[i].h });
      text(s, [b(fails[i].h, { breakLine: true }), t(fails[i].d, { fontSize: 14 })], 1.7, y, 3.95, 1.6, { fontSize: 15 });
    }
    // right: vicious cycle (4 nodes)
    const cx = 9.45, cy = 4.15, rx = 2.05, ry = 1.55, nw = 2.35, nh = 0.9;
    const nodes = [
      { x: cx, y: cy - ry, en: "Body temperature ↑", kr: "체온 상승" },
      { x: cx + rx, y: cy, en: "Enzyme activity ↑", kr: "효소 활성도 증가" },
      { x: cx, y: cy + ry, en: "Metabolic rate ↑", kr: "대사율 상승" },
      { x: cx - rx, y: cy, en: "Still more heat", kr: "더 많은 열 발생" },
    ];
    for (const n of nodes) {
      s.addText([b(n.en, { color: C.background1, breakLine: true }), t(n.kr, { color: C.background1, fontSize: 13 })], { shape: S.ROUNDED_RECTANGLE, x: n.x - nw / 2, y: n.y - nh / 2, w: nw, h: nh, rectRadius: 0.14, fill: { color: C.accent1 }, line: { type: "none" }, fontSize: 15, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "cycle node: " + n.en });
    }
    const A = (x1, y1, x2, y2) => arrow(s, x1, y1, x2, y2, { col: C.accent1, w: 3 });
    A(cx + nw / 2 + 0.05, cy - ry + 0.1, cx + rx - 0.25, cy - nh / 2 - 0.05);
    A(cx + rx - 0.25, cy + nh / 2 + 0.05, cx + nw / 2 + 0.05, cy + ry - 0.1);
    A(cx - nw / 2 - 0.05, cy + ry - 0.1, cx - rx + 0.25, cy + nh / 2 + 0.05);
    A(cx - rx + 0.25, cy - nh / 2 - 0.05, cx - nw / 2 - 0.05, cy - ry + 0.1);
    text(s, [b("vicious cycle", { color: C.accent1, breakLine: true }), t("악순환", { color: C.accent4, fontSize: 13, breakLine: true }), t("= positive feedback", { fontSize: 13, italic: true })], cx - 1.0, cy - 0.45, 2.0, 0.9, { fontSize: 16, align: "center", valign: "middle" });
    card(s, 6.2, 6.3, 6.53, 0.46, { fill: C.text2, r: 0.23 });
    text(s, [t("Not terminated by ", { color: C.background1 }), b("immediate cooling", { color: C.accent5 }), t(" → may lead to death", { color: C.background1 })], 6.3, 6.3, 6.33, 0.46, { fontSize: 14, align: "center", valign: "middle" });
    s.addNotes(NOTE("vicious"));
};

SL.posfb = async () => {
    const s = contentSlide("4. Positive Feedback", "PART 4 · POSITIVE FEEDBACK", T_PART, "Positive Feedback Amplifies Change", "양성피드백은 변화의 효과를 촉진시킨다");
    card(s, 0.6, 1.95, 12.13, 0.9, { fill: C.accent1, transparency: 86, r: 0.16 });
    text(s, [b("Positive feedback", { color: C.accent1 }), t(" (양성피드백): a change produces a response that "), b("amplifies that change"), t(".")], 0.9, 1.95, 11.6, 0.9, { fontSize: 19, valign: "middle" });
    // haywire thermostat staircase
    text(s, [b("Imagine a thermostat gone haywire", { color: C.text2, breakLine: true }), t("고장 난 온도조절기를 떠올려 보자", { color: C.accent4, fontSize: 14 })], 0.6, 3.1, 6.2, 0.75, { fontSize: 17 });
    const steps = ["You nudge up the setting", "Room reaches the set point", "Instead of turning off: “more heat!”", "… and so on"];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + i * 1.55, h = 0.9 + i * 0.55, y = 6.7 - h;
      s.addShape(S.ROUNDED_RECTANGLE, { x, y, w: 1.4, h, rectRadius: 0.1, fill: { color: C.accent1, transparency: 70 - i * 18 }, line: { type: "none" }, objectName: "step " + (i + 1) });
      text(s, steps[i], x + 0.08, y + 0.08, 1.24, h - 0.16, { fontSize: 13, bold: i === 2, color: i >= 2 ? C.background1 : C.text1, align: "center", valign: "top" });
    }
    await ic(s, fa.FaFire, HEX.coral, 5.73, 3.66, 0.44, "more heat");
    text(s, "a very uncomfortable situation", 2.3, 3.66, 3.35, 0.44, { fontSize: 13, italic: true, color: C.accent4, align: "right", valign: "middle" });
    // right: why rare
    card(s, 7.3, 3.1, 5.43, 3.6, { fill: C.text2, r: 0.16 });
    text(s, [b("Why is it rare?", { color: C.accent3, fontSize: 20, breakLine: true }), t("왜 드물까?", { color: C.accent5, fontSize: 14 })], 7.6, 3.25, 4.9, 0.8, {});
    text(s, bl([
      [t("Its effects could "), b("spiral out of control", { color: C.accent3 }), t(" → relatively rare in biological systems.")],
      [t("Under normal physiological conditions, positive feedback events are "), b("eventually halted by negative feedback", { color: C.accent5 }), t(".")],
      [t("It also occurs in nonliving systems (→ Earth Watch).")],
    ], { gap: 8, indent: 14 }), 7.6, 4.1, 4.9, 2.5, { fontSize: 15, color: C.background1 });
    s.addNotes(NOTE("posfb"));
};

SL.birth = async () => {
    const s = contentSlide("4. Positive Feedback", "PART 4 · POSITIVE FEEDBACK", T_PART, "Childbirth: A Positive Feedback Loop", "출산 — 사람과 다른 포유류에서 양성피드백의 대표적인 사례");
    const f = await figure(FIG.ap111.f);
    const fh = 4.45, fw = fh * f.ar;
    s.addImage({ data: f.data, x: 0.6, y: 1.95, w: fw, h: fh, altText: "OpenStax A&P Figure 1.11: positive feedback loop of childbirth" });
    credit(s, FIG.ap111.credit, 0.6, 6.45, fw);
    const x0 = 0.6 + fw + 0.35, w0 = 13.33 - 0.6 - x0;
    const steps = [
      [t("Early labor contractions push the baby’s head against the "), b("cervix"), t(" (자궁경부) → it stretches and begins to open")],
      [t("Nerve cells in the cervix signal the "), b("hypothalamus"), t(" (시상하부)")],
      [t("The hypothalamus triggers release of "), b("oxytocin", { color: C.accent1 }), t(" (옥시토신)")],
      [t("Oxytocin → stronger "), b("uterine contractions"), t(" (자궁수축) → cervix stretched even further → more oxytocin …")],
    ];
    for (let i = 0; i < 4; i++) {
      const y = 1.95 + i * 0.86;
      s.addText(String(i + 1), { shape: S.OVAL, x: x0, y: y + 0.12, w: 0.46, h: 0.46, fill: { color: C.accent1 }, line: { type: "none" }, fontSize: 15, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
      text(s, steps[i], x0 + 0.6, y, w0 - 0.6, 0.8, { fontSize: 14, valign: "middle" });
    }
    card(s, x0, 5.45, w0, 1.3, { fill: C.accent2, transparency: 86 });
    text(s, [b("How it ends: ", { color: C.accent2 }), t("the contractions cause the "), b("birth of the baby"), t(", relieving pressure on the cervix — "), b("negative feedback", { color: C.accent2 }), t(" ends the cycle.")], x0 + 0.2, 5.5, w0 - 0.35, 0.72, { fontSize: 14, valign: "middle" });
    text(s, [b("Also: let-down reflex", { color: C.text2 }), t(" (하강반사) — milk release during breastfeeding; stops when the full baby stops suckling.")], x0 + 0.2, 6.2, w0 - 0.35, 0.5, { fontSize: 12, color: C.text1, valign: "middle" });
    s.addNotes(NOTE("birth"));
};

SL.arctic1 = async () => {
    const s = contentSlide("4. Positive Feedback", "EARTH WATCH", T_EARTH, "Positive Feedback in the Arctic", "북극의 양성피드백 — 줄어드는 북극의 얼음");
    // NASA satellite maps of the September minimum (1984 vs 2012) on top; polar bear photo underneath
    s.addImage({ path: FIG.nasaArctic.f, x: 0.65, y: 1.95, w: 5.5, h: 1.8, altText: "NASA satellite maps of September Arctic sea-ice minimum, 1984 (left) and 2012 (right)" });
    text(s, "September 1984", 0.78, 2.03, 2.4, 0.3, { fontSize: 13, bold: true, color: HEX.white });
    text(s, "September 2012", 3.58, 2.03, 2.4, 0.3, { fontSize: 13, bold: true, color: HEX.white });
    text(s, FIG.nasaArctic.credit, 0.6, 3.8, 5.6, 0.26, { fontSize: 11, italic: true, color: C.accent4 });
    const img = await photo(PHOTO.bear.f, 5.6, 2.15, { pos: "centre", px: 1200 });
    s.addImage({ data: img, x: 0.6, y: 4.15, w: 5.6, h: 2.15, altText: "Polar bear on Arctic drift ice" });
    text(s, "Polar bears depend on Arctic ice to hunt and rest.", 0.6, 6.33, 5.6, 0.3, { fontSize: 13, italic: true });
    credit(s, PHOTO.bear.credit, 0.6, 6.62, 5.6);
    const x0 = 6.55, w0 = 6.18;
    text(s, [t("Climate scientists are convinced that global climate change has triggered positive feedback in the Arctic — the result may be the "), b("planetary equivalent of hyperthermia", { color: C.accent1 }), t(".")], x0, 1.95, w0, 1.05, { fontSize: 15 });
    text(s, [t("The North Pole (≈400 miles from the nearest land) was formerly covered with ice "), b("2–5 m"), t(" (6–15 ft) thick.")], x0, 3.0, w0, 0.7, { fontSize: 15 });
    const stats = [
      { big: "2×", d: "faster Arctic warming than the global average", col: C.accent1 },
      { big: "50%", d: "thinner ice cap in the past 30 years", col: C.accent2 },
      { big: "35%", d: "smaller in area in the past 30 years", col: C.accent2 },
      { big: "10%", d: "average shrinkage per decade (currently)", col: C.text2 },
    ];
    for (let i = 0; i < 4; i++) {
      const x = x0 + (i % 2) * 3.14, y = 3.85 + Math.floor(i / 2) * 1.47;
      card(s, x, y, 3.04, 1.35, { fill: C.background2 });
      text(s, stats[i].big, x + 0.15, y + 0.1, 1.35, 1.15, { fontSize: 36, bold: true, color: stats[i].col, fontFace: "+mj-lt", valign: "middle", align: "center" });
      text(s, stats[i].d, x + 1.55, y + 0.1, 1.4, 1.15, { fontSize: 13, valign: "middle" });
    }
    s.addNotes(NOTE("arctic1"));
};

SL.arctic2 = async () => {
    const s = contentSlide("4. Positive Feedback", "EARTH WATCH", T_EARTH, "Ice Reflects, Water Absorbs", "얼음은 햇빛을 반사하고, 바닷물은 흡수한다 — 전형적인 양성피드백");
    text(s, "Share of incoming sunlight reflected back into space:", 0.6, 1.92, 5.6, 0.4, { fontSize: 15, bold: true, color: C.text2 });
    const refl = [
      { name: "Snow cover on ice", kr: "얼음을 덮은 눈", val: "as much as 80%", frac: 0.8, col: C.accent5 },
      { name: "Sea ice", kr: "해빙", val: "50–70%", frac: 0.7, frac0: 0.5, col: C.accent2 },
      { name: "Seawater", kr: "바닷물", val: "only 5–10%", frac: 0.1, frac0: 0.05, col: C.text2 },
    ];
    for (let i = 0; i < 3; i++) {
      const y = 2.45 + i * 1.2, r = refl[i];
      text(s, [b(r.name, { breakLine: true }), t(r.kr, { color: C.accent4, fontSize: 13 })], 0.6, y, 2.0, 0.85, { fontSize: 15, valign: "middle" });
      const bx = 2.7, bw = 2.6;
      s.addShape(S.RECTANGLE, { x: bx, y: y + 0.25, w: bw, h: 0.36, fill: { color: C.background2 }, line: { type: "none" } });
      if (r.frac0) {
        s.addShape(S.RECTANGLE, { x: bx, y: y + 0.25, w: bw * r.frac0, h: 0.36, fill: { color: r.col }, line: { type: "none" } });
        s.addShape(S.RECTANGLE, { x: bx + bw * r.frac0, y: y + 0.25, w: bw * (r.frac - r.frac0), h: 0.36, fill: { color: r.col, transparency: 55 }, line: { type: "none" } });
      } else s.addShape(S.RECTANGLE, { x: bx, y: y + 0.25, w: bw * r.frac, h: 0.36, fill: { color: r.col }, line: { type: "none" } });
      text(s, r.val, bx, y + 0.63, bw, 0.3, { fontSize: 13, bold: true, color: C.text1 });
    }
    text(s, "→ nearly all solar energy hitting seawater is absorbed, warming the Arctic water.", 0.6, 6.05, 5.4, 0.7, { fontSize: 14, italic: true });
    // loop
    const cx = 9.45, cy = 4.2, rx = 2.0, ry = 1.5, nw = 2.4, nh = 0.92;
    const nodes = [
      { x: cx, y: cy - ry, en: "Arctic ice melts", kr: "북극의 얼음이 녹음", col: C.accent2 },
      { x: cx + rx, y: cy, en: "More water exposed", kr: "더 많은 물이 드러남", col: C.text2 },
      { x: cx, y: cy + ry, en: "Water absorbs more sunlight", kr: "더 많은 햇빛을 흡수", col: C.text2 },
      { x: cx - rx, y: cy, en: "Water becomes warmer", kr: "물의 온도 상승", col: C.accent1 },
    ];
    for (const n of nodes) s.addText([b(n.en, { color: C.background1, breakLine: true }), t(n.kr, { color: C.background1, fontSize: 12 })], { shape: S.ROUNDED_RECTANGLE, x: n.x - nw / 2, y: n.y - nh / 2, w: nw, h: nh, rectRadius: 0.14, fill: { color: n.col }, line: { type: "none" }, fontSize: 14, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "loop node: " + n.en });
    const A = (x1, y1, x2, y2) => arrow(s, x1, y1, x2, y2, { col: C.accent1, w: 3 });
    A(cx + nw / 2 + 0.05, cy - ry + 0.1, cx + rx - 0.25, cy - nh / 2 - 0.05);
    A(cx + rx - 0.25, cy + nh / 2 + 0.05, cx + nw / 2 + 0.05, cy + ry - 0.1);
    A(cx - nw / 2 - 0.05, cy + ry - 0.1, cx - rx + 0.25, cy + nh / 2 + 0.05);
    A(cx - rx + 0.25, cy - nh / 2 - 0.05, cx - nw / 2 - 0.05, cy - ry + 0.1);
    await ic(s, fa.FaSun, HEX.amber, cx - 0.33, cy - 0.62, 0.66, "sun");
    text(s, [b("positive", { color: C.accent1, breakLine: true }), b("feedback", { color: C.accent1 })], cx - 0.9, cy + 0.07, 1.8, 0.6, { fontSize: 14, align: "center" });
    text(s, "“… and so on — a classic example of positive feedback, with unpredictable global consequences.”", 6.25, 6.32, 6.48, 0.45, { fontSize: 13, italic: true, color: C.accent4, align: "center" });
    s.addNotes(NOTE("arctic2"));
};

SL.revisited = async () => {
    const s = contentSlide("Wrap-up", "CASE STUDY · REVISITED", T_CASE, "Case Study Revisited: Who Is at Risk?", "열사병은 운동선수만의 문제가 아니다");
    text(s, [t("When the air temperature becomes extremely high — especially with "), b("high humidity"), t(" — people can die of heat stroke while sitting in their living rooms.")], 0.6, 1.92, 12.13, 0.72, { fontSize: 16 });
    const grp = [
      { I: md.MdElderly, en: "The elderly", kr: "노인", d: "far more likely than athletes to die of heat stroke, during summer heat waves", col: C.accent1 },
      { I: fa.FaChild, en: "Children", kr: "어린이", d: "temperature homeostasis not fully developed; produce more heat and sweat less than adults", col: C.accent1 },
      { I: md.MdAgriculture, en: "Farm workers", kr: "농업 노동자", d: "toiling under the hot sun — at considerable risk", col: C.accent3 },
    ];
    for (let i = 0; i < 3; i++) {
      const y = 2.85 + i * 1.3;
      card(s, 0.6, y, 6.3, 1.15, { fill: C.background2 });
      await iconCircle(s, 0.8, y + 0.2, 0.75, grp[i].I, { bg: grp[i].col, name: grp[i].en });
      text(s, [b(grp[i].en), t("  " + grp[i].kr, { color: C.accent4, fontSize: 13, breakLine: true }), t(grp[i].d, { fontSize: 14 })], 1.75, y + 0.05, 5.0, 1.05, { fontSize: 16, valign: "middle" });
    }
    // car stat
    card(s, 7.2, 2.85, 5.53, 2.5, { fill: C.text2, r: 0.16 });
    await ic(s, fa.FaCar, HEX.ice, 7.5, 3.05, 0.6, "car");
    text(s, [b("Closed car", { color: C.background1 }), t("  밀폐된 차 안", { color: C.accent5, fontSize: 13 })], 8.25, 3.05, 4.3, 0.6, { fontSize: 17, valign: "middle" });
    text(s, "32°C", 7.5, 3.8, 1.75, 0.8, { fontSize: 34, bold: true, color: C.accent5, fontFace: "+mj-lt", align: "center", valign: "middle" });
    arrow(s, 9.3, 4.2, 10.15, 4.2, { col: C.accent3, w: 3 });
    text(s, "60°C", 10.2, 3.8, 1.9, 0.8, { fontSize: 34, bold: true, color: C.accent1, fontFace: "+mj-lt", align: "center", valign: "middle" });
    text(s, "outside (90°F)  →  inside in < 1 hour (140°F)", 7.5, 4.62, 5.0, 0.35, { fontSize: 13, color: C.background1, align: "center" });
    text(s, "About 35–40 U.S. children die each year when left alone in cars.", 7.5, 4.95, 5.0, 0.35, { fontSize: 13, color: C.accent3, align: "center", bold: true });
    card(s, 7.2, 5.5, 5.53, 1.25, { fill: C.accent3, transparency: 80 });
    text(s, [b("Heat waves: ", { color: C.text2 }), t("climate models predict more intense and frequent heat waves. The 2003 European heat wave caused 35,000–52,000 “excess deaths.”")], 7.4, 5.55, 5.15, 1.15, { fontSize: 13, valign: "middle" });
    s.addNotes(NOTE("revisited"));
};

SL.keyconcepts = async () => {
    const s = contentSlide("Wrap-up", null, null, "Key Concepts of 32.1", "핵심개념의 요약", "CONTENT_DARK");
    const kc = [
      { I: fa.FaBalanceScale, col: C.background1, h: "Homeostasis", kr: "항상성", d: "a dynamic equilibrium: internal conditions (temperature, salt, oxygen, glucose, pH, water) are kept within a range in which proteins can function and energy can be made available." },
      { I: fa.FaThermometerHalf, col: C.accent3, h: "Ectotherms vs. endotherms", kr: "외온·내온동물", d: "Ectotherms get most body warmth from the environment; endotherms get most heat from metabolic activities and regulate body temperature within a narrow range." },
      { I: fa.FaSyncAlt, col: C.accent5, h: "Negative feedback", kr: "음성피드백", d: "a change triggers a response that counteracts the change and restores the original conditions — the main way homeostasis is maintained." },
      { I: tb.TbArrowBigUpLines, col: C.accent1, h: "Positive feedback", kr: "양성피드백", d: "a change initiates events that intensify the change — rare, and self-limiting (자기 제한적), e.g., childbirth." },
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + (i % 2) * 6.17, y = 1.98 + Math.floor(i / 2) * 2.45;
      card(s, x, y, 5.96, 2.25, { fill: C.background1, transparency: 92, r: 0.16 });
      await iconCircle(s, x + 0.3, y + 0.3, 0.8, kc[i].I, { bg: kc[i].col, fg: HEX.navy, name: kc[i].h });
      text(s, [b(kc[i].h, { color: kc[i].col, fontSize: 19 }), t("  " + kc[i].kr, { color: C.accent5, fontSize: 14 })], x + 1.3, y + 0.3, 4.5, 0.8, { valign: "middle" });
      text(s, kc[i].d, x + 0.3, y + 1.15, 5.4, 1.05, { fontSize: 14, color: C.background1 });
    }
    s.addNotes(NOTE("keyconcepts"));
};

SL.refs = async () => {
    const s = contentSlide("Wrap-up", "REFERENCES", T_REV, "References & Image Credits", "참고 자료 및 이미지 출처");
    text(s, bl([
      [b("Main textbook: "), t("Audesirk, Audesirk & Byers, "), t("Biology: Life on Earth", { italic: true }), t(", Ch. 32 “Homeostasis and the Organization of the Animal Body” — 32.1 (English ed. pp. 604–609, Chapter Review pp. 617–618; Korean ed. pp. 602–607, 615–616).")],
      [b("Figures: "), t("OpenStax, "), t("Anatomy and Physiology 2e", { italic: true }), t(", Figs. 1.10 & 1.11; OpenStax, "), t("Biology 2e", { italic: true }), t(", Fig. 33.23 — openstax.org, licensed CC BY 4.0. Hexokinase structure (slide 7): David S. Goodsell & RCSB PDB, Molecule of the Month “The Glycolytic Enzymes” (doi:10.2210/rcsb_pdb/mom_2004_2), CC BY 4.0. Arctic sea-ice maps (slide 22): NASA Earth Observatory images by Jesse Allen (public domain).")],
      [b("Photos (Wikimedia Commons): "), t("desert pupfish — P. V. Loiselle (CC BY-SA 3.0); cinnamon hummingbird — Charles J. Sharp (CC BY-SA 4.0); green iguana — Cayambe (CC BY-SA 3.0); polar bear — Andreas Weith (CC BY-SA 4.0); Claude Bernard — public domain; Walter B. Cannon — Bachrach 1934 (CC BY 4.0).")],
      [b("Diagrams: "), t("negative feedback loop (after textbook Fig. 32-2), heat-stroke vicious cycle, Arctic ice loop, and schematic graphs were drawn for this presentation from the textbook text. Icons: react-icons (Font Awesome, Game-icons, Material, Tabler).")],
    ], { gap: 10, indent: 14 }), 0.6, 1.95, 12.13, 4.8, { fontSize: 15 });
    s.addNotes(NOTE("refs"));
};

// ================================================================ short-version slides
// Merged slides that keep only the essential 32.1 textbook content.

SL.caseShort = async () => {
  const s = contentSlide(null, "CASE STUDY", T_CASE, "Case Study: Overheated", "과도한 열 — 열사병으로 쓰러진 대학 운동선수들");
  const rows = [
    { big: "41.1°C", who: "Kierann Smith · triathlete", what: "Collapsed in a 2011 collegiate triathlon. Ice water and ice packs brought her below 40°C — she survived.", col: C.accent3 },
    { big: "41.6°C", who: "Ben Richards · wrestler", what: "Collapsed during a 5-mile training run (2013) and died 10 days later.", col: C.accent1 },
    { big: "2014", who: "Marquese Meadow · football player", what: "Disoriented after practice; died in hospital. Autopsy: heat stroke.", col: C.accent1 },
  ];
  let y = 1.95;
  for (const r of rows) {
    card(s, 0.6, y, 6.25, 1.42, { fill: C.background2, r: 0.14 });
    text(s, r.big, 0.72, y, 1.8, 1.42, { fontSize: 30, bold: true, color: r.col, fontFace: "+mj-lt", align: "center", valign: "middle" });
    text(s, [b(r.who, { breakLine: true, color: C.text2 }), t(r.what, { fontSize: 14 })], 2.62, y + 0.08, 4.08, 1.26, { fontSize: 16, valign: "middle" });
    y += 1.55;
  }
  card(s, 7.1, 1.95, 5.63, 4.52, { fill: C.text2, r: 0.16 });
  await ic(s, fa.FaQuestionCircle, HEX.amber, 7.38, 2.17, 0.46, "questions");
  text(s, [b("Questions for 32.1", { color: C.accent3 }), t("  이 절에서 답할 질문", { color: C.accent5, fontSize: 14 })], 7.98, 2.15, 4.6, 0.5, { fontSize: 18, valign: "middle" });
  const qs = [
    ["What is homeostasis?", "항상성이란 무엇인가?"],
    ["What control mechanisms normally keep body temperature within narrow limits?", "어떤 조절작용이 체온을 좁은 범위 안에서 유지하는가?"],
    ["Why must body temperature be closely regulated?", "왜 체온은 면밀히 조절되어야 하는가?"],
    ["What happens to the body during heat stroke?", "열사병 동안 우리 몸은 어떻게 변하는가?"],
  ];
  qs.forEach((q, i) => {
    const yy = 2.85 + i * 0.9;
    s.addText(String(i + 1), { shape: S.OVAL, x: 7.38, y: yy + 0.04, w: 0.42, h: 0.42, fill: { color: C.accent3 }, line: { type: "none" }, fontSize: 14, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    text(s, [t(q[0], { color: C.background1, breakLine: true }), t(q[1], { color: C.accent5, fontSize: 13 })], 7.98, yy, 4.6, 0.86, { fontSize: 15 });
  });
  s.addNotes(NOTE("caseShort"));
};

SL.whatShort = async () => {
  const s = contentSlide(null, "PART 1 · WHAT & WHY", T_PART, "What Is Homeostasis?", "항상성이란 무엇인가?");
  card(s, 0.6, 1.95, 12.13, 1.45, { fill: C.text2, r: 0.16 });
  text(s, [b("Homeostasis ", { color: C.accent5 }), t("(항상성)", { color: C.accent5 }), t(" — the ability of an organism to maintain its internal environment within the narrow limits that allow optimal cell functioning.", { color: C.background1 })], 0.95, 2.03, 11.4, 0.88, { fontSize: 19, valign: "middle" });
  text(s, "유기체가 세포의 적절한 기능을 위해 필요한 조건의 범위 속에서 내부 환경을 유지하려는 능력", 0.95, 2.93, 11.4, 0.36, { fontSize: 14, color: C.accent5 });
  // who named it
  const pw = 2.0, ph = 1.85, py = 3.65, px = [0.6, 2.85];
  s.addImage({ data: await photo(PHOTO.bernard.f, pw, ph, { pos: "north", px: 600 }), x: px[0], y: py, w: pw, h: ph, altText: "Portrait of Claude Bernard" });
  s.addImage({ data: await photo(PHOTO.cannon.f, pw, ph, { pos: "north", px: 600 }), x: px[1], y: py, w: pw, h: ph, altText: "Portrait of Walter B. Cannon" });
  text(s, [b("Claude Bernard", { color: C.text2, breakLine: true }), t("mid-19th century — “internal constancy” (내부 불변성)", { fontSize: 12 })], px[0], py + ph + 0.06, pw + 0.12, 0.85, { fontSize: 14 });
  text(s, [b("Walter Cannon", { color: C.text2, breakLine: true }), t("1920s — coined the term “homeostasis”", { fontSize: 12 })], px[1], py + ph + 0.06, pw + 0.12, 0.85, { fontSize: 14 });
  credit(s, "Bernard: public domain · Cannon: Bachrach 1934, CC BY 4.0", 0.6, 6.5, 4.4);
  // what it regulates
  const rx = 5.3, rw = 7.43;
  text(s, "Homeostatic mechanisms regulate …", rx, 3.6, rw, 0.38, { fontSize: 15, bold: true, color: C.text2, valign: "middle" });
  const tiles = [
    ["Temperature", "체온", fa.FaThermometerHalf, C.accent1],
    ["Water & salt", "물과 염 농도", fa.FaTint, C.accent2],
    ["Glucose", "포도당 농도", gi.GiCubes, C.accent3],
    ["pH", "산-염기 균형", fa.FaFlask, C.accent6],
    ["Hormones", "호르몬 분비", fa.FaVial, C.text2],
    ["O₂ & CO₂", "산소·이산화탄소", fa.FaWind, C.accent4],
  ];
  const tW = (rw - 0.3) / 3, tH = 0.82;
  for (let i = 0; i < 6; i++) {
    const x = rx + (i % 3) * (tW + 0.15), y = 4.05 + Math.floor(i / 3) * (tH + 0.13);
    card(s, x, y, tW, tH, { fill: C.background2, r: 0.12 });
    await iconCircle(s, x + 0.14, y + 0.14, 0.54, tiles[i][2], { bg: tiles[i][3], name: tiles[i][0] });
    text(s, [b(tiles[i][0], { breakLine: true }), t(tiles[i][1], { color: C.accent4, fontSize: 12 })], x + 0.8, y + 0.05, tW - 0.86, tH - 0.1, { fontSize: 14, valign: "middle" });
  }
  card(s, rx, 6.0, rw, 0.75, { fill: C.accent2, transparency: 86, r: 0.14 });
  text(s, [b("Dynamic, not static", { color: C.accent2 }), t(" (동적인 평형상태) — e.g., exercise → ↑ respiratory rate, ↑ sweat-gland activity, ↑ blood flow to the skin")], rx + 0.2, 6.0, rw - 0.35, 0.75, { fontSize: 14, valign: "middle" });
  s.addNotes(NOTE("whatShort"));
};

SL.ectoendoShort = async () => {
  const s = contentSlide(null, "PART 2 · BODY TEMPERATURE", T_PART, "Ectotherms vs. Endotherms", "외온동물과 내온동물 — 열을 주로 어디서 얻는가에 따른 분류");
  const cols = [
    { x: 0.6, col: C.accent2, title: "ECTOTHERMS", kr: "외온동물 · Greek “outside heat”", I: fa.FaSun, src: [b("Heat source: "), t("mostly the environment (환경)")], ex: "reptiles (except birds), amphibians, most fish and invertebrates" },
    { x: 6.77, col: C.accent1, title: "ENDOTHERMS", kr: "내온동물 · Greek “inside heat”", I: fa.FaFire, src: [b("Heat source: "), t("mostly metabolic reactions (대사반응)")], ex: "birds and mammals" },
  ];
  for (const c of cols) {
    const w = 5.96, y = 1.95, h = 2.5;
    card(s, c.x, y, w, h, { fill: c.col, transparency: 88, r: 0.16 });
    await iconCircle(s, c.x + 0.3, y + 0.25, 0.85, c.I, { bg: c.col, name: c.title });
    text(s, [b(c.title, { color: c.col, fontSize: 22, breakLine: true }), t(c.kr, { fontSize: 13, color: C.accent4 })], c.x + 1.35, y + 0.22, w - 1.6, 0.9, { valign: "middle" });
    text(s, c.src, c.x + 0.3, y + 1.27, w - 0.6, 0.55, { fontSize: 15 });
    text(s, [b("e.g., "), t(c.ex)], c.x + 0.3, y + 1.84, w - 0.6, 0.55, { fontSize: 15 });
  }
  text(s, [b("“Warm-blooded” and “cold-blooded” are misleading", { color: C.text2 }), t("   온혈·냉혈이라는 말은 오해의 소지가 있다", { color: C.accent4, fontSize: 13 })], 0.6, 4.6, 12.13, 0.38, { fontSize: 15, valign: "middle" });
  const ex = [
    { p: PHOTO.pupfish, x: 0.6, name: "Desert pupfish", kind: "“cold-blooded” fish", d: "over 37.8°C when the summer sun heats the springs and ponds it lives in", col: C.accent2 },
    { p: PHOTO.humming, x: 6.77, name: "Hummingbird", kind: "“warm-blooded” bird", d: "up to 40.5°C while foraging by day, but may cool to 13°C at night", col: C.accent1 },
  ];
  for (const e of ex) {
    const w = 5.96, y = 5.08, h = 1.67;
    card(s, e.x, y, w, h, { fill: C.background2, r: 0.14 });
    s.addImage({ data: await photo(e.p.f, 2.05, 1.37, { bg: HEX.light, px: 700 }), x: e.x + 0.15, y: y + 0.15, w: 2.05, h: 1.37, altText: e.name });
    text(s, [b(e.name, { color: C.text2 }), t(" — " + e.kind, { color: e.col, bold: true, breakLine: true }), t(e.d, { fontSize: 14 })], e.x + 2.38, y + 0.12, w - 2.55, 1.15, { fontSize: 15 });
    credit(s, e.p.credit.replace("Photo: ", "").replace(", via Wikimedia Commons", " · Wikimedia Commons"), e.x + 2.38, y + 1.33, w - 2.55);
  }
  s.addNotes(NOTE("ectoendoShort"));
};

SL.tempregShort = async () => {
  const s = contentSlide(null, "PART 2 · BODY TEMPERATURE", T_PART, "How Animals Regulate Body Temperature", "동물들은 체온 조절 방식이 다르다");
  const y0 = 1.95, h = 4.8;
  // ectotherms
  const lx = 0.6, lw = 7.0;
  card(s, lx, y0, lw, h, { fill: C.accent2, transparency: 88, r: 0.16 });
  text(s, [b("Ectotherms", { color: C.accent2, fontSize: 20 }), t("  외온동물 — body temperature varies with the environment", { color: C.accent4, fontSize: 13 })], lx + 0.25, y0 + 0.12, lw - 0.5, 0.5, { valign: "middle" });
  const pw = 2.35, ph = 2.0;
  s.addImage({ data: await photo(PHOTO.iguana.f, pw, ph, { px: 800, bg: "E4F1F3" }), x: lx + 0.25, y: y0 + 0.75, w: pw, h: ph, altText: "Green iguana basking on a rock" });
  credit(s, "Cayambe, CC BY-SA 3.0 · Wikimedia", lx + 0.25, y0 + 0.75 + ph + 0.03, pw);
  text(s, bl([
    [t("Many (butterflies, bees, most lizards) cool down and become "), b("inactive at night"), t(", conserving energy")],
    [t("At daybreak bees "), b("shiver"), t(", butterflies "), b("beat their wings"), t(", and lizards seek a "), b("warm, sunlit stone")],
  ], { gap: 6, indent: 14 }), lx + 2.85, y0 + 0.75, lw - 3.1, 2.3, { fontSize: 14 });
  text(s, bl([
    [b("Behavior: "), t("bask in the sun to warm up, move to a shady spot to cool off")],
    [b("Desert pupfish: "), t("breeds only in a narrow temperature range → swims to different areas of its pool")],
    [b("Deep-sea fish: "), t("the water stays at about 0–3°C, so body temperature barely changes")],
  ], { gap: 6, indent: 14 }), lx + 0.25, y0 + 3.1, lw - 0.5, 1.62, { fontSize: 14 });
  // endotherms
  const rx = 7.85, rw = 4.88;
  card(s, rx, y0, rw, h, { fill: C.accent1, transparency: 88, r: 0.16 });
  text(s, [b("Endotherms", { color: C.accent1, fontSize: 20 }), t("  내온동물", { color: C.accent4, fontSize: 13 })], rx + 0.25, y0 + 0.12, rw - 0.5, 0.5, { valign: "middle" });
  text(s, "35–41°C", rx + 0.25, y0 + 0.68, rw - 0.5, 0.75, { fontSize: 36, bold: true, color: C.accent1, fontFace: "+mj-lt", valign: "middle" });
  text(s, "a fairly constant body temperature (a few, like hummingbirds, let it fall at night)", rx + 0.25, y0 + 1.45, rw - 0.5, 0.62, { fontSize: 13 });
  text(s, [b("Benefits", { color: C.accent2 }), t("  이점 — a warm body usually can …", { color: C.accent4, fontSize: 12 })], rx + 0.25, y0 + 2.15, rw - 0.5, 0.36, { fontSize: 15, valign: "middle" });
  const ben = [[fa.FaEye, "sense its environment better"], [fa.FaBolt, "respond more quickly"], [fa.FaRunning, "move faster than a cold body"]];
  for (let i = 0; i < 3; i++) {
    const yy = y0 + 2.58 + i * 0.5;
    await iconCircle(s, rx + 0.25, yy, 0.4, ben[i][0], { bg: C.accent2, name: ben[i][1] });
    text(s, ben[i][1], rx + 0.8, yy, rw - 1.0, 0.4, { fontSize: 14, valign: "middle" });
  }
  text(s, [b("Cost", { color: C.accent1 }), t("  비용 — ", { color: C.accent4, fontSize: 12 }), t("the energy required to maintain a high body temperature")], rx + 0.25, y0 + 4.1, rw - 0.5, 0.62, { fontSize: 14 });
  s.addNotes(NOTE("tempregShort"));
};

SL.negfbShort = async () => {
  const s = contentSlide(null, "PART 3 · NEGATIVE FEEDBACK", T_PART, "Negative Feedback Reverses Change", "음성피드백은 변화의 효과를 역전시킨다");
  text(s, [t("Internal conditions are regulated by "), b("feedback systems", { color: C.text2 }), t(" (피드백 시스템). There are two types:")], 0.6, 1.9, 12.13, 0.4, { fontSize: 16, valign: "middle" });
  const cs = [
    { x: 0.6, col: C.accent2, title: "Negative feedback", kr: "음성피드백", def: [t("A change triggers responses that "), b("counteract"), t(" it, returning the system to its original condition — the "), b("main mechanism of homeostasis"), t(".")], neg: true },
    { x: 6.77, col: C.accent1, title: "Positive feedback", kr: "양성피드백", def: [t("Creates cycles in which changes "), b("amplify themselves"), t(". Relatively rare in living things.")], neg: false },
  ];
  for (const c of cs) {
    const w = 5.96, y = 2.42, h = 2.0;
    card(s, c.x, y, w, h, { fill: c.col, transparency: 89, r: 0.16 });
    text(s, [b(c.title, { color: c.col, fontSize: 19 }), t("  " + c.kr, { color: C.accent4, fontSize: 14 })], c.x + 0.25, y + 0.12, 3.4, 0.45, { valign: "middle" });
    text(s, c.def, c.x + 0.25, y + 0.62, 3.25, 1.3, { fontSize: 14 });
    const gx = c.x + 3.62, gy = y + 0.35, gw = 2.12, gh = 1.35;
    s.addShape(S.LINE, { x: gx, y: gy + gh, w: gw, h: 0, line: { color: C.accent4, width: 1 } });
    s.addShape(S.LINE, { x: gx, y: gy, w: 0, h: gh, line: { color: C.accent4, width: 1 } });
    s.addShape(S.LINE, { x: gx, y: gy + 0.85, w: gw, h: 0, line: { color: C.accent4, width: 1, dashType: "dash" } });
    text(s, "set point", gx + gw - 0.9, gy + 0.87, 0.9, 0.22, { fontSize: 10, color: C.accent4, italic: true, align: "right" });
    if (c.neg) curve(s, gx, gy, gw, gh, [[0, 0.85], [0.35, 0.18], [0.7, 1.22], [1.05, 0.58], [1.4, 1.0], [1.75, 0.8], [2.12, 0.85]], { color: C.accent2, width: 2.5 });
    else s.addShape(S.CUSTOM_GEOMETRY, { x: gx, y: gy, w: gw, h: gh, points: [{ x: 0, y: 0.85 }, { x: 2.12, y: 0.02, curve: { type: "cubic", x1: 1.2, y1: 0.85, x2: 1.9, y2: 0.6 } }], line: { color: C.accent1, width: 2.5 } });
  }
  text(s, [b("All negative feedback systems have three components", { color: C.text2 }), t("   음성피드백의 세 가지 구성요소", { color: C.accent4, fontSize: 13 })], 0.6, 4.6, 12.13, 0.38, { fontSize: 15, valign: "middle" });
  // descriptions are pre-broken into lines so no word (or Korean term) is split by automatic wrapping
  const comp = [
    { en: "Sensor", kr: "감지기", d: ["detects the current condition"], I: md.MdSensors },
    { en: "Control center", kr: "조절 센터", d: ["compares it to the", "set point (설정점)"], I: fa.FaBrain },
    { en: "Effector", kr: "효과기", d: ["produces an output that", "restores the condition"], I: fa.FaCogs },
  ];
  const cw = 3.75, gap = 0.44;
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * (cw + gap), y = 5.08;
    card(s, x, y, cw, 1.12, { fill: C.background2, r: 0.12 });
    await iconCircle(s, x + 0.15, y + 0.24, 0.64, comp[i].I, { bg: C.accent2, name: comp[i].en });
    const desc = comp[i].d.map((line, k) => t(line, { fontSize: 13, breakLine: k < comp[i].d.length - 1 }));
    text(s, [b(comp[i].en, { color: C.accent2 }), t("  " + comp[i].kr, { color: C.accent4, fontSize: 12, breakLine: true }), ...desc], x + 0.92, y + 0.06, cw - 1.0, 1.0, { fontSize: 15, valign: "middle" });
    if (i < 2) arrow(s, x + cw + 0.06, y + 0.56, x + cw + gap - 0.06, y + 0.56, { col: C.accent2, w: 2.25 });
  }
  text(s, [b("Examples: "), t("body temperature, blood glucose and salts, standing upright · a home thermostat, a toilet tank, cruise control")], 0.6, 6.32, 12.13, 0.42, { fontSize: 13, valign: "middle" });
  s.addNotes(NOTE("negfbShort"));
};

SL.arcticShort = async () => {
  const s = contentSlide(null, "EARTH WATCH", T_EARTH, "Positive Feedback in the Arctic", "북극의 양성피드백 — 얼음은 햇빛을 반사하고, 바닷물은 흡수한다");
  const pw = 4.9, ph = 3.0;
  s.addImage({ data: await photo(PHOTO.bear.f, pw, ph, { px: 1100 }), x: 0.6, y: 1.95, w: pw, h: ph, altText: "Polar bear on Arctic drift ice" });
  text(s, "Polar bears depend on Arctic ice to hunt and rest.", 0.6, 4.99, pw, 0.28, { fontSize: 12, italic: true });
  credit(s, PHOTO.bear.credit, 0.6, 5.26, pw);
  const stats = [["2×", ["faster warming than", "the global average"], C.accent1], ["50%", ["thinner ice cap", "in 30 years"], C.accent2], ["35%", ["smaller ice area", "in 30 years"], C.accent2]];
  const sw = (pw - 0.3) / 3;
  stats.forEach(([big, d, col], i) => {
    const x = 0.6 + i * (sw + 0.15), y = 5.6;
    card(s, x, y, sw, 1.15, { fill: C.background2, r: 0.12 });
    text(s, big, x, y + 0.05, sw, 0.58, { fontSize: 26, bold: true, color: col, fontFace: "+mj-lt", align: "center", valign: "middle" });
    text(s, [t(d[0], { breakLine: true }), t(d[1])], x + 0.04, y + 0.63, sw - 0.08, 0.48, { fontSize: 11, align: "center" });
  });
  const rx = 5.85, rw = 6.88;
  text(s, [t("Sea ice reflects "), b("50–70%"), t(" of sunlight (snow cover as much as "), b("80%"), t("); seawater reflects only "), b("5–10%", { color: C.accent1 }), t(" and absorbs the rest.")], rx, 1.92, rw, 0.75, { fontSize: 15 });
  const cx = rx + rw / 2, cy = 4.6, rX = 2.15, rY = 1.4, nw = 2.45, nh = 0.85;
  const nodes = [
    { x: cx, y: cy - rY, en: "Arctic ice melts", kr: "북극의 얼음이 녹음", col: C.accent2 },
    { x: cx + rX, y: cy, en: "More water exposed", kr: "더 많은 물이 드러남", col: C.text2 },
    { x: cx, y: cy + rY, en: "Water absorbs more sunlight", kr: "더 많은 햇빛을 흡수", col: C.text2 },
    { x: cx - rX, y: cy, en: "Water becomes warmer", kr: "물의 온도 상승", col: C.accent1 },
  ];
  for (const n of nodes) s.addText([b(n.en, { color: C.background1, breakLine: true }), t(n.kr, { color: C.background1, fontSize: 12 })], { shape: S.ROUNDED_RECTANGLE, x: n.x - nw / 2, y: n.y - nh / 2, w: nw, h: nh, rectRadius: 0.14, fill: { color: n.col }, line: { type: "none" }, fontSize: 14, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "loop node: " + n.en });
  const A = (x1, y1, x2, y2) => arrow(s, x1, y1, x2, y2, { col: C.accent1, w: 3 });
  A(cx + nw / 2 + 0.05, cy - rY + 0.1, cx + rX - 0.25, cy - nh / 2 - 0.05);
  A(cx + rX - 0.25, cy + nh / 2 + 0.05, cx + nw / 2 + 0.05, cy + rY - 0.1);
  A(cx - nw / 2 - 0.05, cy + rY - 0.1, cx - rX + 0.25, cy + nh / 2 + 0.05);
  A(cx - rX + 0.25, cy - nh / 2 - 0.05, cx - nw / 2 - 0.05, cy - rY + 0.1);
  await ic(s, fa.FaSun, HEX.amber, cx - 0.3, cy - 0.58, 0.6, "sun");
  text(s, [b("positive", { color: C.accent1, breakLine: true }), b("feedback", { color: C.accent1 })], cx - 0.9, cy + 0.06, 1.8, 0.55, { fontSize: 13, align: "center" });
  text(s, "… and so on — a classic example of positive feedback", rx, 6.47, rw, 0.28, { fontSize: 12, italic: true, color: C.accent4, align: "center" });
  s.addNotes(NOTE("arcticShort"));
};

SL.refsShort = async () => {
  const s = contentSlide(null, "REFERENCES", T_REV, "References & Image Credits", "참고 자료 및 이미지 출처");
  text(s, bl([
    [b("Main textbook: "), t("Audesirk, Audesirk & Byers, "), t("Biology: Life on Earth", { italic: true }), t(", Ch. 32 “Homeostasis and the Organization of the Animal Body,” section 32.1 (English and Korean editions).")],
    [b("Figures: "), t("OpenStax, "), t("Anatomy and Physiology 2e", { italic: true }), t(", Fig. 1.11; OpenStax, "), t("Biology 2e", { italic: true }), t(", Fig. 33.23 — openstax.org, CC BY 4.0. Hexokinase structure: David S. Goodsell & RCSB PDB, Molecule of the Month “The Glycolytic Enzymes” (doi:10.2210/rcsb_pdb/mom_2004_2), CC BY 4.0.")],
    [b("Photos (Wikimedia Commons): "), t("desert pupfish — P. V. Loiselle (CC BY-SA 3.0); cinnamon hummingbird — Charles J. Sharp (CC BY-SA 4.0); green iguana — Cayambe (CC BY-SA 3.0); polar bear — Andreas Weith (CC BY-SA 4.0); Claude Bernard — public domain; Walter B. Cannon — Bachrach 1934 (CC BY 4.0).")],
    [b("Diagrams: "), t("the negative feedback loop (after textbook Fig. 32-2), the heat-stroke cycle, the Arctic loop, and the schematic graphs were drawn for this presentation from the textbook text.")],
  ], { gap: 10, indent: 14 }), 0.6, 1.95, 12.13, 4.8, { fontSize: 15 });
  s.addNotes(NOTE("refsShort"));
};

// ================================================================ variants
// "§Name" starts a section; every other entry is a slide id from SL.
const VARIANTS = {
  main: ["§Opening", "title", "case", "roadmap", "§1. What & Why", "constancy", "what", "dynamic", "enzymes", "case2",
    "§2. Body Temperature", "warmcold", "ectoendo", "ectocycle", "ectoreg", "endocost",
    "§3. Negative Feedback", "twotypes", "negfb", "fig322", "hypo", "water",
    "§4. Positive Feedback", "vicious", "posfb", "birth", "arctic1", "arctic2",
    "§Wrap-up", "revisited", "keyconcepts", "refs"],
  short: ["§Opening", "title", "caseShort", "§1. What & Why", "whatShort", "enzymes",
    "§2. Body Temperature", "ectoendoShort", "tempregShort",
    "§3. Negative Feedback", "negfbShort", "fig322", "hypo",
    "§4. Positive Feedback", "vicious", "posfb", "birth", "arcticShort",
    "§Wrap-up", "keyconcepts", "refsShort"],
};

async function build() {
  for (const item of VARIANTS[VARIANT]) {
    if (item.startsWith("§")) {
      CURRENT_SECTION = item.slice(1);
      pres.addSection({ title: CURRENT_SECTION });
      continue;
    }
    if (!SL[item]) throw new Error(`unknown slide id "${item}"`);
    const before = OUTLINE.length;
    await SL[item]();
    OUTLINE[before].id = item;
    OUTLINE[before].section = CURRENT_SECTION;
  }
}

// ---------------------------------------------------------------- illustrations (SVG)
const ENZ_BROKEN = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 420 300" width="420" height="300">
  <path d="M24 170 C60 70 108 262 158 160 C204 64 248 70 290 160 C328 240 372 240 398 130" fill="none" stroke="#${HEX.coral}" stroke-width="16" stroke-linecap="round"/>
  <g stroke="#9AA6B2" stroke-width="5" stroke-dasharray="8 7" stroke-linecap="round">
    <line x1="80" y1="70" x2="92" y2="104"/><line x1="214" y1="214" x2="226" y2="248"/><line x1="330" y1="70" x2="342" y2="100"/>
  </g>
  <polygon points="300,40 362,16 362,70" fill="#${HEX.amber}" opacity="0.55" stroke="#${HEX.amber}" stroke-width="10" stroke-linejoin="round"/>
</svg>`;

(async () => {
  await build();
  await pres.writeFile({ fileName: OUT });
  const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
  const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
  // Theme: pptxgenjs writes Office's stock palette, so write this deck's colors and name into the theme,
  // and give East-Asian text a theme font so Korean runs resolve to Malgun Gothic (맑은 고딕).
  const part = "ppt/theme/theme1.xml";
  const SLOTS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
  const scheme = `<a:clrScheme name="${THEME.name}">` + SLOTS.map((k) => `<a:${k}><a:srgbClr val="${THEME.colors[k]}"/></a:${k}>`).join("") + "</a:clrScheme>";
  let xml = await zip.file(part).async("string");
  xml = xml
    .replace(/<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/, () => scheme)
    .replace(/(<a:(?:theme|fontScheme)\b[^>]*?\bname=")[^"]*"/g, (_, head) => `${head}${THEME.name}"`)
    .replace(/<a:ea typeface=""\/>/g, '<a:ea typeface="맑은 고딕"/>');
  if (!xml.includes(scheme)) throw new Error("theme color scheme not found in " + part);
  zip.file(part, xml);
  // pptxgenjs 4.0.1 writes a <a:pPr> before every run of a multi-run paragraph (schema-invalid, and the
  // later buNone blocks cancel bullets). Keep only the leading pPr of each paragraph, and set eaLnBrk="0"
  // so Korean wraps at word boundaries instead of mid-word.
  const PPR = /<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g;
  let fixedParas = 0;
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/(slides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(n))) {
    let sx = await zip.file(name).async("string");
    sx = sx.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, inner) => {
      const lead = inner.match(/^\s*(<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>))/);
      const first = lead ? lead[1] : null;
      let rest = first ? inner.slice(inner.indexOf(first) + first.length) : inner;
      const before = rest.length;
      rest = rest.replace(PPR, "");
      if (rest.length !== before) fixedParas++;
      let ppr = first || "<a:pPr/>";
      if (!/eaLnBrk=/.test(ppr)) ppr = ppr.replace(/^<a:pPr/, '<a:pPr eaLnBrk="0"');
      return "<a:p>" + ppr + rest + "</a:p>";
    });
    // Keep a Korean gloss on the same line as the word before it ("enzyme (효소)"): the space before
    // "(" + Hangul becomes a no-break space, both inside one run and across a run boundary.
    const HANGUL_PAREN = /\((?=[가-힣])/.source;
    sx = sx.replace(/(<a:t(?:\s[^>]*)?>)([^<]*)(<\/a:t>)/g, (m, open, body, close) => open + body.replace(new RegExp(" " + HANGUL_PAREN, "g"), " (") + close);
    sx = sx.replace(new RegExp(" (</a:t></a:r><a:r>(?:(?!</a:r>)[\\s\\S])*?<a:t(?:\\s[^>]*)?>)" + HANGUL_PAREN, "g"), " $1(");
    if (name === "ppt/slides/slide1.xml") {
      const KR_BOX = '<a:off x="685800" y="4389120"/><a:ext cx="6949440" cy="822960"/>';
      if (!sx.includes(KR_BOX)) throw new Error("title slide subtitle box not found");
      sx = sx.replace(KR_BOX, '<a:off x="685800" y="4389120"/><a:ext cx="7315200" cy="822960"/>');
    }
    zip.file(name, sx);
  }
  console.log("paragraphs repaired:", fixedParas);
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT);

  // Export the short version's read-aloud script as a Markdown file next to the deck.
  {
    const syllables = (str) => (str.match(/[가-힣]/g) || []).length + Math.ceil((str.match(/[A-Za-z]/g) || []).length / 3) + (str.match(/[0-9]/g) || []).length;
    const total = OUTLINE.reduce((n, o) => n + syllables(NOTE(o.id)), 0);
    const minutes = Math.round(total / 300); // ~300 syllables per minute of presenting
    const md = [
      `# 32.1 항상성 — 발표 대본${VARIANT === "short" ? " (짧은 버전)" : ""}`,
      "",
      `슬라이드 ${OUTLINE.length}장 · 그대로 읽으면 약 ${minutes}분. PowerPoint 슬라이드 노트에도 같은 대본이 들어 있습니다.`,
      "",
      ...OUTLINE.flatMap((o, i) => [`## ${i + 1}. ${o.title}`, "", `*${o.kr}*`, "", NOTE(o.id).split("\n").join("\n\n"), ""]),
    ].join("\n");
    const mdPath = OUT.replace(/\.pptx$/i, "_script.md");
    fs.writeFileSync(mdPath, md);
    console.log("wrote", mdPath, `(about ${minutes} min to read aloud)`);

    const docxPath = OUT.replace(/\.pptx$/i, "_script.docx");
    fs.writeFileSync(docxPath, await scriptDocx(minutes));
    console.log("wrote", docxPath);
  }
})().catch((e) => { console.error(e); process.exit(1); });

// ---------------------------------------------------------------- Word script
// The same read-aloud script as a Word file: one page per part (easy to hand each presenter their part),
// slide number + title + Korean subtitle above each slide's script.
async function scriptDocx(minutes) {
  const D = require("docx");
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Header, Footer, PageNumber, BorderStyle } = D;
  const FONT = "맑은 고딕";
  const LANG = { value: "en-US", eastAsia: "ko-KR" }; // Latin text breaks as English, Hangul as Korean
  const PART = {
    "Opening": "도입",
    "1. What & Why": "Part 1 · 항상성이란 무엇이고, 왜 유지할까?",
    "2. Body Temperature": "Part 2 · 외온동물과 내온동물은 어떻게 다를까?",
    "3. Negative Feedback": "Part 3 · 음성피드백은 왜 중요할까?",
    "4. Positive Feedback": "Part 4 · 양성피드백이란 무엇이고, 왜 드물까?",
    "Wrap-up": "마무리",
  };
  const muted = (text, size = 20) => new TextRun({ text, size, color: HEX.slate });

  const body = [
    new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun(`32.1 항상성 — 발표 대본${VARIANT === "short" ? " (짧은 버전)" : ""}`)] }),
    new Paragraph({ spacing: { after: 60 }, children: [muted("Homeostasis: Why and How Do Animals Regulate Their Internal Environment?", 22)] }),
    new Paragraph({
      spacing: { after: 240 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "C9D3DC", space: 8 } },
      children: [muted(`슬라이드 ${OUTLINE.length}장 · 그대로 읽으면 약 ${minutes}분 · PowerPoint 슬라이드 노트에도 같은 대본이 들어 있습니다.`)],
    }),
  ];
  let lastPart = null;
  OUTLINE.forEach((o, i) => {
    if (o.section !== lastPart) {
      body.push(new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: lastPart !== null, children: [new TextRun(PART[o.section] || o.section)] }));
      lastPart = o.section;
    }
    body.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: `슬라이드 ${i + 1}`, color: HEX.teal }), new TextRun(`   ${o.title}`)] }));
    body.push(new Paragraph({ keepNext: true, spacing: { after: 120 }, children: [muted(o.kr)] }));
    for (const line of NOTE(o.id).split("\n")) body.push(new Paragraph({ children: [new TextRun(line)] }));
  });

  const doc = new Document({
    title: `32.1 항상성 — 발표 대본${VARIANT === "short" ? " (짧은 버전)" : ""}`,
    description: "General Biology 32.1 Homeostasis — read-aloud presentation script",
    styles: {
      default: {
        document: { run: { font: FONT, size: 24, language: LANG }, paragraph: { spacing: { line: 360, after: 160 } } },
        title: { run: { font: FONT, size: 36, bold: true, color: HEX.navy }, paragraph: { spacing: { before: 0, after: 80 }, keepNext: true } },
        heading1: { run: { font: FONT, size: 30, bold: true, color: HEX.navy }, paragraph: { spacing: { before: 120, after: 160 }, keepNext: true, outlineLevel: 0 } },
        heading2: { run: { font: FONT, size: 25, bold: true, color: HEX.navy }, paragraph: { spacing: { before: 320, after: 40 }, keepNext: true, outlineLevel: 1 } },
      },
    },
    sections: [{
      properties: { page: { margin: { top: 1247, bottom: 1247, left: 1304, right: 1304 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 }, children: [muted("32.1 항상성 · 발표 대본", 16)] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 0 }, children: [new TextRun({ children: [PageNumber.CURRENT, " / ", PageNumber.TOTAL_PAGES], size: 16, color: HEX.slate })] })] }) },
      children: body,
    }],
  });
  return Packer.toBuffer(doc);
}
