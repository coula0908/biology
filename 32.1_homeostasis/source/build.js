// Chapter 32.1 Homeostasis — presentation decks with Korean read-aloud speaker notes:
// English slides (main) and the same slides in Korean (ko).
// Build (from this folder):  npm install  &&  npm run build
//   → writes ../32.1_Homeostasis.pptx with its script (_script.docx, _script.md) and ../32.1_Homeostasis_Korean.pptx
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

// Usage: node build.js <main|ko> [out.pptx]
//   main → English slides, plus the read-aloud script as Word and Markdown files
//   ko   → the same 26 slides with Korean slide text (same layout, same speaker notes)
// Every visible slide string is written L(english, korean).
// (The 15-slide short version was retired from the build; its last files are kept as they were.)
const VARIANT = process.argv[2] || "main";
if (!["main", "ko"].includes(VARIANT)) throw new Error(`unknown variant "${VARIANT}" (use main or ko)`);
const KO = VARIANT === "ko";
const L = (en, ko) => (KO ? ko : en);
const FILE = { main: "32.1_Homeostasis.pptx", ko: "32.1_Homeostasis_Korean.pptx" }[VARIANT];
const OUT = path.resolve(process.argv[3] || path.join(__dirname, "..", FILE));
const NOTES = require("./notes_main.js"); // both decks share the Korean read-aloud script
function NOTE(id) {
  if (!NOTES[id]) throw new Error(`no speaker notes for slide "${id}"`);
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
pres.title = L("32.1 Homeostasis: Why and How Do Animals Regulate Their Internal Environment?", "32.1 항상성: 동물은 왜 그리고 어떻게 자신의 내부 환경을 조절하는가?");
pres.subject = L("General Biology — Chapter 32.1", "일반생물학 — 32장 1절");
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
// Rough text width (inches) for Calibri / Malgun Gothic, used to size chips and tags.
// Hangul counts a full em so Korean chips and tags never wrap.
function tw(str, pt, bold = false) {
  let em = 0;
  for (const ch of str) {
    const c = ch.codePointAt(0);
    if ((c >= 0xac00 && c <= 0xd7a3) || (c >= 0x3130 && c <= 0x318f)) em += 1.0;
    else if (/[A-Z]/.test(ch)) em += 0.6;
    else if (/[a-z]/.test(ch)) em += 0.48;
    else if (/[0-9]/.test(ch)) em += 0.51;
    else if (ch === " ") em += 0.23;
    else em += 0.36;
  }
  return ((em * pt) / 72) * (bold ? 1.07 : 1);
}

const TXT = { valign: "top", margin: 0, isTextBox: true, color: C.text1, fontSize: 16 };
const HANGUL = /[\uac00-\ud7a3]/;
function text(s, runs, x, y, w, h, o = {}) {
  if (KO) {
    // Malgun Gothic has no italic, and slanted Hangul reads poorly: Korean text is set upright.
    const all = typeof runs === "string" ? runs : runs.map((q) => q.text).join("");
    if (o.italic && HANGUL.test(all)) o = Object.assign({}, o, { italic: false });
    if (typeof runs !== "string") runs = runs.map((q) => (q.options && q.options.italic && HANGUL.test(q.text) ? { text: q.text, options: Object.assign({}, q.options, { italic: false }) } : q));
  }
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
  const w = tw(label, 12, true) * 1.12 + 0.44;
  s.addText(label, { shape: S.ROUNDED_RECTANGLE, x: 0.6, y: 0.28, w, h: 0.32, rectRadius: 0.16, fill: { color: col }, line: { type: "none" }, fontSize: 12, bold: true, color: txt, charSpacing: 1, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "section tag" });
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

// kr: Korean subtitle. It is not drawn on the slide (English-only slides, larger text);
// it heads the slide's part of the Word/Markdown script.
function contentSlide(_section, tagLabel, tagOpts, title, kr, master = "CONTENT") {
  const s = pres.addSlide({ masterName: master, sectionTitle: CURRENT_SECTION });
  OUTLINE.push({ title, kr });
  if (tagLabel) tag(s, tagLabel, tagOpts);
  s.addText(title, { placeholder: "title" });
  return s;
}

// ---------------------------------------------------------------- layouts
const footer = { text: { text: L("32.1 Homeostasis · 항상성", "32.1 항상성 · Homeostasis"), options: { x: 0.6, y: 7.02, w: 5, h: 0.28, fontSize: 10, color: C.accent4, margin: 0, valign: "middle" } } };
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
    footer,
  ],
  slideNumber: { x: 12.23, y: 7.02, w: 0.5, h: 0.28, fontSize: 10, color: HEX.slate, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENT_DARK",
  background: { color: HEX.navy },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.66, w: 12.13, h: 0.64, fontSize: 36, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: L("32.1 Homeostasis · 항상성", "32.1 항상성 · Homeostasis"), options: { x: 0.6, y: 7.02, w: 5, h: 0.28, fontSize: 10, color: C.accent5, margin: 0, valign: "middle" } } },
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
    text(s, L("GENERAL BIOLOGY  ·  CHAPTER 32", "일반생물학  ·  32장"), 0.75, 1.55, 7, 0.35, { fontSize: 14, bold: true, color: C.accent3, charSpacing: 2 });
    text(s, "32.1", 0.75, 1.0, 3, 0.6, { fontSize: 30, bold: true, color: C.accent1 });
    if (KO) {
      // Korean title in three lines; the English title takes the subtitle line
      s.addText([{ text: "항상성:", options: { fontSize: 36 } }, { text: "동물은 왜 그리고 어떻게", options: { fontSize: 36, softBreakBefore: true } }, { text: "자신의 내부 환경을 조절하는가?", options: { fontSize: 36, softBreakBefore: true } }], { placeholder: "title" });
      s.addText("Homeostasis: Why and How Do Animals Regulate Their Internal Environment?", { placeholder: "kr" });
    } else {
      s.addText("Homeostasis: Why and How Do Animals Regulate Their Internal Environment?", { placeholder: "title" });
      // line break placed by the presenter in PowerPoint
      s.addText([{ text: "항상성: 동물은 왜 그리고 어떻게" }, { text: "자신의 내부 환경을 조절하는가?", options: { softBreakBefore: true } }], { placeholder: "kr" });
    }
    // thermal ring motif: cold half + hot half around a thermometer
    const cx = 10.55, cy = 3.75, D = 3.6;
    s.addShape(S.BLOCK_ARC, { x: cx - D / 2, y: cy - D / 2, w: D, h: D, angleRange: [95, 265], arcThicknessRatio: 0.16, fill: { color: C.accent5 }, line: { type: "none" }, objectName: "ring cold half" });
    s.addShape(S.BLOCK_ARC, { x: cx - D / 2, y: cy - D / 2, w: D, h: D, angleRange: [275, 85], arcThicknessRatio: 0.16, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "ring hot half" });
    await ic(s, fa.FaThermometerHalf, HEX.white, cx - 0.75, cy - 0.95, 1.5, "thermometer");
    text(s, L("set point", "설정점"), cx - 1.0, cy + 0.62, 2.0, 0.35, { fontSize: 14, color: C.background1, align: "center", italic: true });
    await ic(s, fa.FaSnowflake, HEX.ice, cx - 1.55, cy + D / 2 + 0.2, 0.4, "cold");
    text(s, L("cold", "추위"), cx - 1.1, cy + D / 2 + 0.2, 0.8, 0.4, { fontSize: 14, color: C.accent5, valign: "middle" });
    await ic(s, fa.FaFire, HEX.coral, cx + 0.55, cy + D / 2 + 0.2, 0.4, "hot");
    text(s, L("hot", "더위"), cx + 1.0, cy + D / 2 + 0.2, 0.8, 0.4, { fontSize: 14, color: C.accent1, valign: "middle" });
    s.addNotes(NOTE("title"));
};

// Content slides, large-type layout, content area y 1.55 – 6.85. The Korean subtitle passed to
// contentSlide() heads the slide in the Word/Markdown script; it is not drawn on the slide.
// Sizes: body 18 pt+, secondary text and diagram labels 15–16 pt+, credits 10–12 pt.
// Korean lines are taller than English ones (Malgun Gothic about 1.34 × the font size, Calibri 1.22 ×),
// so a few boxes take Korean geometry through L().

SL.case = async () => {
  const s = contentSlide(null, "CASE STUDY", T_CASE, L("Case Study: Overheated", "과도한 열"), "과도한 열 — 열사병(heat stroke)으로 쓰러진 대학 운동선수들");
  card(s, 0.6, 1.55, 4.75, 5.3, { fill: C.accent3, transparency: 82, r: 0.16 });
  await ic(s, fa.FaQuoteLeft, HEX.amber, 0.9, 1.8, 0.55, "quote");
  text(s, L("“I’m having the race of my life. … Nothing is going to get in my way, not the 90° heat and 90% humidity, … not the fire in my feet that is spreading up my legs.”", "“내 인생 최고의 레이스를 펼치고 있다. … 아무것도 나를 막을 수 없다. 32°C의 더위도, 90%의 습도도, … 다리로 번져 오는 발의 열기도.”"), 0.9, 2.5, 4.2, 2.5, { fontSize: 20, italic: true });
  text(s, [b("Kierann Smith", { breakLine: true, fontSize: 18 }), t(L("medical student, Stanford triathlon team — just before she collapsed (2011 USA Triathlon Collegiate Nationals, Alabama)", "스탠퍼드대 철인 3종 팀의 의대생 — 쓰러지기 직전의 생각 (2011년 미국 대학 철인 3종 대회, 앨라배마)"), { fontSize: 16, color: C.accent4 })], 0.9, 5.1, 4.2, 1.6, {});
  const rows = [
    { big: "41.1°C", f: "106°F", who: L("Kierann Smith · triathlete", "Kierann Smith · 철인 3종 선수"), what: L("Core body temperature at collapse. Ice water and ice packs brought it just below 40°C (104°F). She survived.", "쓰러졌을 때의 심부 체온. 얼음물과 얼음팩으로 40°C(104°F) 바로 아래까지 낮췄고, 살아남았다."), col: C.accent3 },
    { big: "41.6°C", f: "107°F", who: L("Ben Richards · wrestler, 20", "Ben Richards · 레슬링 선수, 20세"), what: L("Collapsed during a 5-mile training run (2013); died 10 days later.", "8 km(5마일) 달리기 훈련 중 쓰러졌고(2013), 열흘 뒤 숨졌다."), col: C.accent1 },
    { big: "2014", f: L("heat stroke", "열사병"), who: L("Marquese Meadow · football player", "Marquese Meadow · 미식축구 선수"), what: L("Disoriented after practice; died in a nearby hospital. Autopsy confirmed heat stroke as the cause of death.", "연습 후 방향 감각을 잃었고 근처 병원에서 숨졌다. 부검 결과 사인은 열사병이었다."), col: C.accent1 },
  ];
  let y = 1.55;
  for (const r of rows) {
    card(s, 5.65, y, 7.08, 1.6, { fill: C.background2, r: 0.14 });
    text(s, r.big, 5.8, y + 0.2, 1.95, 0.75, { fontSize: 36, bold: true, color: r.col, fontFace: "+mj-lt", align: "center", valign: "middle" });
    text(s, r.f, 5.8, y + 0.98, 1.95, 0.4, { fontSize: 16, color: C.accent4, align: "center" });
    text(s, [b(r.who, { breakLine: true, color: C.text2, fontSize: 18 }), t(r.what, { fontSize: 16 })], 7.95, y + 0.1, 4.62, 1.4, { valign: "middle" });
    y += 1.85;
  }
  s.addNotes(NOTE("case"));
};

SL.roadmap = async () => {
  const s = contentSlide(null, L("OVERVIEW", "개요"), T_REV, L("Today’s Roadmap", "오늘 다룰 내용"), "오늘 다룰 내용 — 교재의 학습 목표(Check Your Learning) 네 가지");
  const items = [
    { n: "1", en: L("What is homeostasis — and why do organisms maintain it?", "항상성이란 무엇이고, 왜 유지할까?"), I: fa.FaBalanceScale, col: C.text2 },
    { n: "2", en: L("How do ectotherms and endotherms differ?", "외온동물과 내온동물은 어떻게 다를까?"), I: fa.FaThermometerHalf, col: C.accent3 },
    { n: "3", en: L("Why is negative feedback crucial for homeostasis?", "음성피드백은 항상성에 왜 중요할까?"), I: fa.FaSyncAlt, col: C.accent2 },
    { n: "4", en: L("What is positive feedback — and why is it rare?", "양성피드백은 무엇이고, 왜 드물까?"), I: tb.TbArrowBigUpLines, col: C.accent1 },
  ];
  const w = 2.86, gap = 0.23;
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + i * (w + gap), it = items[i];
    card(s, x, 1.55, w, 4.3, { fill: C.background2, r: 0.16 });
    await iconCircle(s, x + 0.3, 1.85, 0.95, it.I, { bg: it.col, name: "part " + it.n });
    text(s, "PART " + it.n, x + 1.4, 1.85, 1.35, 0.95, { fontSize: 20, bold: true, color: it.col, valign: "middle" });
    text(s, it.en, x + 0.3, 3.1, w - 0.55, 2.5, { fontSize: 24, bold: true });
  }
  card(s, 0.6, 6.1, 12.13, 0.72, { fill: C.accent3, transparency: 80, r: 0.36 });
  await ic(s, fa.FaRunning, HEX.ink, 0.9, 6.2, 0.52, "athlete");
  text(s, L([b("Running thread: "), t("heat stroke — what keeps body temperature constant, and what goes wrong")], [b("이어지는 사례: "), t("열사병 — 무엇이 체온을 일정하게 지키고, 무엇이 잘못되는가")]), 1.6, 6.1, 10.95, 0.72, { fontSize: 18, valign: "middle" });
  s.addNotes(NOTE("roadmap"));
};

SL.constancy = async () => {
  const s = contentSlide(null, L("PART 1 · WHAT & WHY", "PART 1 · 항상성이란"), T_PART, L("From Internal Constancy to Homeostasis", "내부 불변성에서 항상성으로"), "‘내부 불변성’에서 ‘항상성’으로");
  card(s, 0.6, 1.55, 6.25, 2.55, { fill: C.background2 });
  await iconCircle(s, 0.85, 1.85, 0.8, fa.FaThermometerHalf, { bg: C.accent1, name: "temperature" });
  text(s, L([t("Sitting in your room, sweltering in a jungle, or shivering in a blizzard — ", { color: C.accent4 }), t("most cells of your body (heart, brain, muscles) maintain an "), b("almost constant temperature", { color: C.accent1 }), t(".")], [t("방에 앉아 있든, 무더운 정글에 있든, 눈보라 속에서 떨고 있든 — ", { color: C.accent4 }), t("심장, 뇌, 근육 같은 몸속 세포 대부분은 "), b("거의 일정한 온도", { color: C.accent1 }), t("를 유지한다.")]), 1.85, 1.75, 4.8, 2.2, { fontSize: 20 });
  card(s, 0.6, 4.3, 6.25, 2.55, { fill: C.background2 });
  await iconCircle(s, 0.85, 4.6, 0.8, fa.FaTint, { bg: C.accent2, name: "fluid" });
  text(s, L([t("Parched desert air, a pool, the ocean, or the Great Salt Lake — ", { color: C.accent4 }), t("your cells are bathed in "), b("interstitial fluid", { color: C.accent2 }), t(" with an "), b("almost constant composition"), t(".")], [t("바싹 마른 사막 공기, 수영장, 바다, 그레이트솔트레이크 — 어디에 있든 ", { color: C.accent4 }), t("세포는 "), b("세포간질액", { color: C.accent2 }), t("에 잠겨 있고, 그 조성도 "), b("거의 일정"), t("하다.")]), 1.85, 4.5, 4.8, 2.2, { fontSize: 20 });
  const px = [7.25, 10.1], pw = 2.6, ph = 2.7;
  s.addImage({ data: await photo(PHOTO.bernard.f, pw, ph, { pos: "north", px: 700 }), x: px[0], y: 1.55, w: pw, h: ph, altText: "Portrait of Claude Bernard" });
  s.addImage({ data: await photo(PHOTO.cannon.f, pw, ph, { pos: "north", px: 700 }), x: px[1], y: 1.55, w: pw, h: ph, altText: "Portrait of Walter B. Cannon" });
  text(s, L([b("Claude Bernard", { breakLine: true, color: C.text2, fontSize: 20 }), t("French physiologist, mid-19th century", { fontSize: 16, color: C.accent4, breakLine: true }), t("first described the body’s “internal constancy”", { fontSize: 16 })], [b("클로드 베르나르", { breakLine: true, color: C.text2, fontSize: 20 }), t("프랑스 생리학자, 19세기 중반", { fontSize: 16, color: C.accent4, breakLine: true }), t("몸의 ‘내부 불변성’을 처음 설명했다", { fontSize: 16 })]), px[0], 4.4, pw + 0.1, 1.95, {});
  text(s, L([b("Walter Cannon", { breakLine: true, color: C.text2, fontSize: 20 }), t("1920s", { fontSize: 16, color: C.accent4, breakLine: true }), t("coined the term ", { fontSize: 16 }), b("homeostasis", { fontSize: 16, color: C.accent2 })], [b("월터 캐넌", { breakLine: true, color: C.text2, fontSize: 20 }), t("1920년대", { fontSize: 16, color: C.accent4, breakLine: true }), b("항상성", { fontSize: 16, color: C.accent2 }), t("이라는 용어를 만들었다", { fontSize: 16 })]), px[1], 4.4, pw + 0.1, 1.95, {});
  credit(s, "Portraits: Bernard — public domain; Cannon — Bachrach 1934, CC BY 4.0 (Wikimedia Commons)", 7.25, 6.6, 5.5);
  s.addNotes(NOTE("constancy"));
};

SL.what = async () => {
  const s = contentSlide(null, L("PART 1 · WHAT & WHY", "PART 1 · 항상성이란"), T_PART, L("What Is Homeostasis?", "항상성이란 무엇인가?"), "항상성이란 무엇인가?");
  card(s, 0.6, 1.55, 12.13, 1.55, { fill: C.text2, r: 0.16 });
  text(s, L([b("Homeostasis", { color: C.accent5 }), t(" — the ability of an organism to maintain its internal environment within the narrow limits that allow optimal cell functioning.", { color: C.background1 })], [b("항상성", { color: C.accent5 }), t(" — 유기체가 세포의 적절한 기능을 위해 필요한 조건의 범위 속에서 내부 환경을 유지하려는 능력", { color: C.background1 })]), 0.95, 1.6, 11.4, 1.45, { fontSize: 24, valign: "middle" });
  text(s, L("Homeostatic mechanisms regulate …", "항상성 메커니즘이 조절하는 것 …"), 0.6, 3.35, 8, 0.45, { fontSize: 20, bold: true, color: C.text2, valign: "middle" });
  const tiles = [
    { en: L("Temperature", "체온"), I: fa.FaThermometerHalf, col: C.accent1 },
    { en: L("Water & salt in body fluids", "체액의 물과 염분"), I: fa.FaTint, col: C.accent2 },
    { en: L("Glucose", "포도당"), I: gi.GiCubes, col: C.accent3 },
    { en: L("pH (acid–base balance)", "pH (산-염기 균형)"), I: fa.FaFlask, col: C.accent6 },
    { en: L("Hormone secretion", "호르몬 분비"), I: fa.FaVial, col: C.text2 },
    { en: L("O₂ and CO₂", "산소와 이산화탄소"), I: fa.FaWind, col: C.accent4 },
  ];
  const w = 1.86, gap = 0.194;
  for (let i = 0; i < tiles.length; i++) {
    const x = 0.6 + i * (w + gap), it = tiles[i];
    card(s, x, 3.95, w, 2.9, { fill: C.background2 });
    await iconCircle(s, x + (w - 0.9) / 2, 4.5, 0.9, it.I, { bg: it.col, name: it.en });
    text(s, it.en, x + 0.1, 5.55, w - 0.2, 1.2, { fontSize: 20, bold: true, align: "center" });
  }
  s.addNotes(NOTE("what"));
};

SL.dynamic = async () => {
  const s = contentSlide(null, L("PART 1 · WHAT & WHY", "PART 1 · 항상성이란"), T_PART, L("Homeostasis Is Dynamic, Not Static", "항상성은 정적이 아니라 동적이다"), "말뜻은 ‘같은 상태로 머무르다’이지만, 내부 환경은 끊임없이 조절된다");
  card(s, 0.6, 1.55, 3.9, 5.3, { fill: C.background2, r: 0.16 });
  text(s, [t("“homeostasis”", { bold: true, color: C.accent2, fontSize: 26, breakLine: true }), t(L("= “to stay the same”", "= ‘같은 상태로 머무르다’"), { fontSize: 20 })], 0.85, 1.75, 3.45, 1.05, {});
  text(s, L([t("The word implies a static, unchanging state — but the internal environment actually "), b("seethes with activity", { color: C.accent1 }), t(" as the body continuously adjusts to maintain constancy amid varying internal and external conditions.")], [t("이 말은 변화 없는 정적인 상태처럼 들리지만, 실제 내부 환경은 "), b("끊임없이 바쁘게 움직인다", { color: C.accent1 }), t(". 몸 안팎의 조건이 계속 바뀌는 가운데, 몸이 일정함을 유지하려고 쉬지 않고 조절하기 때문이다.")]), 0.85, 2.95, 3.45, 3.0, { fontSize: 18 });
  chip(s, 0.85, 6.1, L("dynamic equilibrium", "동적인 평형상태"), null, { col: C.accent2, size: 18, h: 0.55 });
  await iconCircle(s, 4.97, 3.4, 1.25, fa.FaRunning, { bg: C.accent1, name: "exercise" });
  text(s, L("Exercise", "운동"), 4.6, 4.75, 2.0, 0.42, { fontSize: 20, bold: true, align: "center" });
  text(s, L("challenges homeostasis", "항상성에 대한 도전"), 4.55, 5.2, 2.1, 0.7, { fontSize: 16, italic: true, color: C.accent4, align: "center" });
  text(s, L("Demands", "요구"), 6.75, 1.55, 2.6, 0.7, { fontSize: 20, bold: true, color: C.accent1, valign: "bottom" });
  text(s, L("The body rises to the occasion", "몸의 대응"), 9.6, 1.55, 3.13, 0.7, { fontSize: 20, bold: true, color: C.accent2, valign: "bottom" });
  const dem = L(["More O₂ must be supplied to sustain the energy demands of working muscles", "The extra heat that muscles produce must be dissipated"], ["일하는 근육의 에너지 요구를 채우려면 산소를 더 많이 공급해야 한다", "근육이 만든 여분의 열을 몸 밖으로 내보내야 한다"]);
  const dy = [2.4, 4.7];
  for (let i = 0; i < 2; i++) {
    card(s, 6.75, dy[i], 2.6, 2.1, { fill: C.accent1, transparency: 86 });
    text(s, dem[i], 6.95, dy[i] + 0.15, 2.25, 1.8, { fontSize: 18, valign: "middle" });
    arrow(s, 6.25, 4.03, 6.72, dy[i] + 1.05, { col: C.accent1 });
  }
  const res = [[L("↑ Respiratory rate", "↑ 호흡수"), fa.FaLungs], [L("↑ Sweat-gland activity", "↑ 땀샘 활동"), fa.FaTint], [L("↑ Blood flow to the skin", "↑ 피부로 가는 혈류"), fa.FaHeartbeat]];
  for (let i = 0; i < 3; i++) {
    const y = 2.4 + i * 1.5;
    card(s, 9.6, y, 3.13, 1.35, { fill: C.accent2, transparency: 86 });
    await iconCircle(s, 9.75, y + 0.33, 0.68, res[i][1], { bg: C.accent2, name: res[i][0] });
    text(s, res[i][0], 10.55, y + 0.1, 2.08, 1.15, { fontSize: 18, bold: true, valign: "middle" });
  }
  arrow(s, 9.38, 3.45, 9.57, 3.45, { col: C.accent2 });
  arrow(s, 9.38, 5.75, 9.57, 5.75, { col: C.accent2 });
  s.addNotes(NOTE("dynamic"));
};

SL.enzymes = async () => {
  const s = contentSlide(null, L("PART 1 · WHAT & WHY", "PART 1 · 항상성이란"), T_PART, L("Homeostasis Allows Enzymes to Function", "항상성은 효소가 기능하도록 한다"), "항상성은 효소가 기능하도록 한다 — 세포가 주변 환경에 민감한 이유");
  card(s, 0.6, 1.55, 6.35, 4.35, { fill: C.background2, r: 0.16 });
  text(s, [b(L("① Structure", "① 구조"), { color: C.accent2 })], 0.85, 1.7, 5.9, 0.45, { fontSize: 22 });
  // English: explicit line breaks keep "three-dimensional" from splitting at its hyphen (widths measured at 18 pt).
  // Korean wraps on its own; its four lines need more height, so the drawings sit a little lower.
  text(s, L([t("Almost every reaction in a cell is catalyzed by a specific", { breakLine: true }), b("enzyme"), t(". Its function depends on its precise", { breakLine: true }), b("three-dimensional structure"), t(", maintained in part by", { breakLine: true }), b("hydrogen bonds", { color: C.accent2 }), t(".")], [t("세포의 거의 모든 반응은 특정 "), b("효소"), t("가 촉매한다. 효소의 기능은 정확한 "), b("3차원 구조"), t("에 달려 있고, 이 구조의 일부를 "), b("수소결합", { color: C.accent2 }), t("이 유지한다.")]), 0.85, 2.2, 5.9, L(1.3, 1.45), { fontSize: 18 });
  const folded = await svgPng(ENZ_FOLDED, 700), broken = await svgPng(ENZ_BROKEN, 700);
  s.addImage({ data: folded, x: 0.95, y: L(3.55, 3.7), w: 2.45, h: 1.75, altText: "Folded enzyme with substrate in its active site" });
  s.addImage({ data: broken, x: 4.15, y: L(3.55, 3.7), w: 2.45, h: 1.75, altText: "Unfolded enzyme with disrupted hydrogen bonds" });
  arrow(s, 3.5, L(4.42, 4.57), 4.05, L(4.42, 4.57), { col: C.accent1, w: 2.5 });
  text(s, L("works", "정상 작동"), 0.8, L(5.35, 5.47), 2.75, 0.4, { fontSize: 16, bold: true, color: C.accent2, align: "center" });
  text(s, L("bonds disrupted", "결합이 깨짐"), 4.0, L(5.35, 5.47), 2.75, 0.4, { fontSize: 16, bold: true, color: C.accent1, align: "center" });
  card(s, 7.2, 1.55, 5.53, 4.35, { fill: C.background2, r: 0.16 });
  text(s, [b(L("② Rate", "② 반응 속도"), { color: C.accent1 })], 7.45, 1.7, 5.1, 0.45, { fontSize: 22 });
  text(s, L([t("The rate at which enzymes catalyze reactions is "), b("temperature-dependent"), t(".")], [t("효소가 반응을 촉매하는 속도는 "), b("온도에 따라 달라진다"), t(".")]), 7.45, 2.2, 5.1, 0.7, { fontSize: 18 });
  const bx = 7.6, by = 3.05, bw = 4.75;
  s.addShape(S.ROUNDED_RECTANGLE, { x: bx, y: by, w: bw / 2, h: 0.32, rectRadius: 0.16, fill: { color: C.accent5 }, line: { type: "none" } });
  s.addShape(S.ROUNDED_RECTANGLE, { x: bx + bw / 2, y: by, w: bw / 2, h: 0.32, rectRadius: 0.16, fill: { color: C.accent1 }, line: { type: "none" } });
  s.addShape(S.RECTANGLE, { x: bx + bw / 2 - 0.2, y: by, w: 0.4, h: 0.32, fill: { color: C.accent3 }, line: { type: "none" } });
  await ic(s, fa.FaSnowflake, HEX.teal, bx - 0.05, by + 0.42, 0.38, "low temperature");
  await ic(s, fa.FaFire, HEX.coral, bx + bw - 0.33, by + 0.42, 0.38, "high temperature");
  text(s, L([b("Low temperature", { color: C.accent2, breakLine: true }), t("slows molecular motion → fewer and slower interactions between enzymes and their molecules")], [b("낮은 온도", { color: C.accent2, breakLine: true }), t("분자 운동이 느려진다 → 효소와 분자의 상호작용이 줄고 느려진다")]), 7.45, 3.95, 2.55, 1.9, { fontSize: 16 });
  text(s, L([b("Rising temperature", { color: C.accent1, breakLine: true }), t("reactions speed up — sometimes to a detrimental level")], [b("온도 상승", { color: C.accent1, breakLine: true }), t("반응이 빨라진다 — 때로는 해로운 수준까지")]), 10.2, 3.95, 2.4, 1.9, { fontSize: 16 });
  let cx = 0.6;
  for (const en of L(["too salty", "too acidic", "too basic", "too hot"], ["염분 과다", "강산성", "강염기성", "고온"])) cx += chip(s, cx, 6.1, en, null, { col: C.accent1, size: 16, h: 0.5 }) + 0.12;
  const ccx = cx + 0.1;
  card(s, ccx, 6.1, 12.73 - ccx, 0.5, { fill: C.text2, r: 0.25 });
  text(s, L([t("→ keep ", { color: C.background1 }), b("salt, pH & temperature", { color: C.accent3 }), t(" in narrow ranges", { color: C.background1 })], [t("→ ", { color: C.background1 }), b("염 농도, pH, 온도", { color: C.accent3 }), t("를 좁은 범위로 유지", { color: C.background1 })]), ccx + 0.1, 6.1, 12.73 - ccx - 0.2, 0.5, { fontSize: 16, valign: "middle", align: "center" });
  s.addNotes(NOTE("enzymes"));
};

SL.case2 = async () => {
  const s = contentSlide(null, L("CASE STUDY · CONTINUED", "CASE STUDY · 계속"), T_CASE, L("Case Study: Why Didn’t They Recover?", "왜 회복하지 못했을까?"), "과도한 열 — 체온은 몇 시간 만에 정상이 되었는데, 왜 회복하지 못했을까?");
  text(s, L([t("Meadow and Richards were treated in hospital for over a week, and their body temperatures were brought down to normal "), b("within a few hours"), t(". Yet both died.")], [t("Meadow와 Richards는 일주일 넘게 병원에서 치료를 받았고, 체온은 "), b("몇 시간 만에"), t(" 정상으로 내려갔다. 그런데도 둘 다 숨졌다.")]), 0.6, 1.55, 12.13, 0.75, { fontSize: 18 });
  card(s, 0.6, 2.55, 2.7, 3.15, { fill: C.accent1, r: 0.16 });
  await ic(s, fa.FaThermometerFull, HEX.white, 1.55, 2.8, 0.8, "heat stroke");
  text(s, [b(L("Heat stroke", "열사병"), { breakLine: true, color: C.background1, fontSize: 24 }), t(L("→ metabolic chaos", "→ 신진대사의 이상"), { color: C.background1, fontSize: 18 })], 0.75, 3.8, 2.4, 1.6, { align: "center" });
  arrow(s, 3.35, 3.6, 3.95, 3.15, { col: C.accent1 });
  arrow(s, 3.35, 4.7, 3.95, 5.1, { col: C.accent1 });
  card(s, 4.0, 2.45, 4.3, 1.4, { fill: C.background2 });
  await iconCircle(s, 4.15, 2.72, 0.85, gi.GiMuscleUp, { bg: C.accent4, name: "muscle" });
  text(s, [b(L("Muscle cells die", "근육세포가 죽으면서"), { breakLine: true }), t(L("and release substances into the bloodstream", "여러 물질을 혈액으로 방출한다"))], 5.15, 2.5, 3.05, 1.3, { fontSize: 18, valign: "middle" });
  arrow(s, 8.35, 3.15, 8.95, 3.15, { col: C.accent1 });
  card(s, 9.0, 2.45, 3.73, 1.4, { fill: C.accent1, transparency: 85 });
  await iconCircle(s, 9.15, 2.72, 0.85, gi.GiKidneys, { bg: C.accent1, name: "kidneys" });
  text(s, [b(L("Kidney failure", "신부전"))], 10.15, 2.5, 2.5, 1.3, { fontSize: 20, valign: "middle" });
  card(s, 4.0, 4.4, 4.3, 1.4, { fill: C.background2 });
  await iconCircle(s, 4.15, 4.67, 0.85, fa.FaThermometerFull, { bg: C.accent4, name: "elevated temperature" });
  text(s, [b(L("Elevated body temperature", "높아진 체온은"), { breakLine: true }), t(L("also damages the liver", "간도 손상시킨다"))], 5.15, 4.45, 3.05, 1.3, { fontSize: 18, valign: "middle" });
  arrow(s, 8.35, 5.1, 8.95, 5.1, { col: C.accent1 });
  card(s, 9.0, 4.4, 3.73, 1.4, { fill: C.accent1, transparency: 85 });
  await iconCircle(s, 9.15, 4.67, 0.85, gi.GiLiver, { bg: C.accent1, name: "liver" });
  text(s, [b(L("Liver damage", "간 손상"), { breakLine: true, fontSize: 20 }), t(L("the liver is crucial for removing toxic substances from the blood", "간은 혈액에서 독성 물질을 없애는 데 중요하다"), { color: C.accent4, fontSize: 16 })], 10.15, 4.42, 2.5, 1.36, { valign: "middle" });
  card(s, 0.6, 6.05, 12.13, 0.78, { fill: C.text2, r: 0.3 });
  text(s, [t(L("Human metabolism functions only within a narrow range of temperatures. ", "사람의 대사는 좁은 온도 범위에서만 제대로 작동한다. "), { color: C.background1 }), b(L("Is this true for all animals?", "모든 동물이 그럴까?"), { color: C.accent3 }), t("  → Part 2", { color: C.accent5 })], 0.9, 6.05, 11.6, 0.78, { fontSize: 18, valign: "middle" });
  s.addNotes(NOTE("case2"));
};

SL.warmcold = async () => {
  const s = contentSlide(null, L("PART 2 · BODY TEMPERATURE", "PART 2 · 체온"), T_PART, L("“Warm-Blooded” or “Cold-Blooded”?", "온혈동물일까, 냉혈동물일까?"), "온혈동물일까, 냉혈동물일까? — 이 용어들은 오해를 부르기 쉽다");
  text(s, L([t("Mammals and birds are often called "), b("“warm-blooded”", { color: C.accent1 }), t("; other reptiles, amphibians, fish, and invertebrates", { breakLine: true }), b("“cold-blooded”", { color: C.accent2 }), t(". But these terms are often "), b("misleading"), t(":")], [t("흔히 포유류와 조류를 "), b("‘온혈동물’", { color: C.accent1 }), t(", 나머지 파충류·양서류·어류·무척추동물을 "), b("‘냉혈동물’", { color: C.accent2 }), t("이라고 부른다. 하지만 이 용어들은 "), b("오해를 부르기 쉽다"), t(".")]), 0.6, 1.55, 12.13, 0.8, { fontSize: 18 });
  const cards = [
    { p: PHOTO.pupfish, x: 0.6, title: L("Desert pupfish", "사막 pupfish"), kind: L("“cold-blooded” fish", "‘냉혈’ 물고기"), stat: "> 37.8°C", statF: L("over 100°F", "100°F 이상"), desc: L("when the summer sun heats the springs and small ponds where it lives — so pupfish are sometimes quite “warm-blooded.”", "한여름 햇볕에 서식지인 샘과 작은 연못이 데워질 때 — 그래서 pupfish는 때로 꽤 ‘온혈’인 셈이다."), col: C.accent1 },
    { p: PHOTO.humming, x: 6.77, title: L("Hummingbird", "벌새"), kind: L("“warm-blooded” bird", "‘온혈’ 조류"), stat: "40.5 → 13°C", statF: L("105°F by day → 55°F at night", "낮 105°F → 밤 55°F"), desc: L("as high as 40.5°C while foraging during the day, but may cool down to 13°C during the night.", "낮에 먹이를 찾을 때는 40.5°C까지 오르지만, 밤에는 13°C까지 내려가기도 한다."), col: C.accent2 },
  ];
  for (const c of cards) {
    const w = 5.96;
    card(s, c.x, 2.55, w, 4.3, { fill: C.background2, r: 0.16 });
    text(s, [b(c.title, { color: C.text2, fontSize: 22 }), t("   " + c.kind, { color: c.col, fontSize: 18, bold: true })], c.x + 0.25, 2.65, w - 0.5, 0.55, { valign: "middle" });
    s.addImage({ data: await photo(c.p.f, 3.0, 2.0, { bg: HEX.light, px: 900 }), x: c.x + 0.25, y: 3.3, w: 3.0, h: 2.0, altText: c.title });
    text(s, c.stat, c.x + 3.4, 3.4, 2.45, 0.8, { fontSize: c.stat.length > 8 ? 30 : 34, bold: true, color: c.col, fontFace: "+mj-lt", valign: "middle" });
    text(s, c.statF, c.x + 3.4, 4.2, 2.4, 0.65, { fontSize: 16, color: C.accent4 });
    text(s, c.desc, c.x + 0.25, 5.45, w - 0.5, 1.0, { fontSize: 18 });
    credit(s, c.p.credit, c.x + 0.25, 6.5, w - 0.5);
  }
  s.addNotes(NOTE("warmcold"));
};

SL.ectoendo = async () => {
  const s = contentSlide(null, L("PART 2 · BODY TEMPERATURE", "PART 2 · 체온"), T_PART, L("Ectotherms vs. Endotherms", "외온동물과 내온동물"), "외온동물과 내온동물 — 열을 주로 어디서 얻는가에 따른 분류");
  text(s, L("To avoid confusion, scientists classify animals by their major source of body warmth.", "과학자들은 혼동을 피하려고, 몸을 데우는 열을 주로 어디서 얻는지에 따라 동물을 분류한다."), 0.6, 1.55, 12.13, 0.45, { fontSize: 18 });
  const cols = [
    { x: 0.6, col: C.accent2, colHex: HEX.teal, title: L("ECTOTHERMS", "외온동물"), greek: L("Greek: “outside heat”", "ectotherm · 그리스어로 ‘외부의 열’"), I: fa.FaSun, src: L([t("Derive most of their heat", { breakLine: true }), t("from the "), b("environment"), t(".")], [t("열의 대부분을 "), b("환경"), t("에서 얻는다.")]), ex: [[L("Reptiles (except birds)", "파충류 (조류 제외)"), gi.GiGecko], [L("Amphibians", "양서류"), gi.GiFrog], [L("Most fish", "대부분의 어류"), fa.FaFish], [L("Most invertebrates", "대부분의 무척추동물"), gi.GiButterfly]], note: L("Simplest cases — an earthworm in its burrow or a fish in a stream: body temperature = environment temperature.", "가장 단순한 경우 — 굴 속의 지렁이나 개울 속 물고기: 체온 = 주변 환경의 온도.") },
    { x: 6.77, col: C.accent1, colHex: HEX.coral, title: L("ENDOTHERMS", "내온동물"), greek: L("Greek: “inside heat”", "endotherm · 그리스어로 ‘내부의 열’"), I: fa.FaFire, src: L([t("Produce most of their heat", { breakLine: true }), t("by "), b("metabolic reactions"), t(".")], [t("열의 대부분을 "), b("대사반응"), t("으로 만든다.")]), ex: [[L("Birds", "조류"), gi.GiHummingbird], [L("Mammals", "포유류"), fa.FaPaw]], note: L("Birds and mammals are the most common endotherms; some other animals occasionally generate significant heat metabolically.", "조류와 포유류가 가장 흔한 내온동물이다. 다른 동물도 가끔 대사작용으로 상당한 열을 낸다.") },
  ];
  for (const c of cols) {
    const w = 5.96;
    card(s, c.x, 2.2, w, 4.65, { fill: c.col, transparency: 88, r: 0.16 });
    await iconCircle(s, c.x + 0.3, 2.42, 1.0, c.I, { bg: c.col, name: c.title });
    text(s, [b(c.title, { color: c.col, fontSize: 26, breakLine: true }), t(c.greek, { fontSize: 18, color: C.accent4 })], c.x + 1.5, 2.42, 4.3, 1.0, { valign: "middle" });
    text(s, c.src, c.x + 0.3, 3.62, w - 0.6, 0.8, { fontSize: 20 });
    let ex = c.x + 0.55;
    for (const [name, I] of c.ex) {
      await iconCircle(s, ex, 4.6, 0.6, I, { bg: C.background1, fg: c.colHex, scale: 0.7, name });
      text(s, name, ex + 0.3 - 0.68, 5.25, 1.36, 0.6, { fontSize: 16, align: "center" });
      ex += 1.4;
    }
    text(s, c.note, c.x + 0.3, 5.92, w - 0.6, 0.88, { fontSize: 16, italic: true });
  }
  s.addNotes(NOTE("ectoendo"));
};

SL.ectocycle = async () => {
  const s = contentSlide(null, L("PART 2 · BODY TEMPERATURE", "PART 2 · 체온"), T_PART, L("Ectotherms: Cooling Down, Warming Up", "외온동물: 밤에 식고, 낮에 데워진다"), "외온동물의 체온은 시간·일·주 단위의 외부 기온 변화에 따라 크게 달라진다");
  text(s, L([t("Many ectotherms — "), b("butterflies, bees, and most lizards"), t(" — cool down and become inactive at night, "), b("conserving energy", { color: C.accent2 }), t(", then warm up and become active during the day.")], [b("나비, 벌, 대부분의 도마뱀"), t(" 같은 많은 외온동물은 밤에 몸이 식어 활동을 멈추고 "), b("에너지를 아낀다", { color: C.accent2 }), t(". 그리고 낮에 다시 데워져 활동한다.")]), 0.6, 1.55, 12.13, 0.85, { fontSize: 18 });
  const st = [
    { I: fa.FaMoon, bg: C.text2, h: L("NIGHT", "밤"), body: L([b("Butterflies and bees"), t(" become so cool that they cannot fly."), t("", { breakLine: true }), b("Lizards"), t(" are often too sluggish to hunt or escape predators.")], [b("나비와 벌"), t("은 날지 못할 만큼 몸이 차가워진다."), t("", { breakLine: true }), b("도마뱀"), t("은 너무 굼떠서 사냥하거나 포식자를 피하기 어렵다.")]), fill: C.text2, A: gi.GiButterfly },
    { I: md.MdWbTwilight, bg: C.accent3, h: L("DAYBREAK", "새벽"), body: L([b("Bees shiver"), t(" and "), b("butterflies beat their wings"), t(" to generate metabolic heat."), t("", { breakLine: true }), b("Lizards"), t(" seek a warm, sunlit stone.")], [b("벌은 몸을 떨고"), t(", "), b("나비는 날갯짓"), t("을 해서 대사열을 만든다."), t("", { breakLine: true }), b("도마뱀"), t("은 햇볕에 데워진 바위를 찾는다.")]), fill: C.accent3, A: gi.GiBee },
    { I: fa.FaSun, bg: C.accent1, h: L("DAY", "낮"), body: L([t("Warmed up, they "), b("become active during the day"), t(" — the heat they gained lets them "), b("resume their active lifestyles"), t(".")], [t("몸이 데워지면 "), b("낮 동안 다시 활발해진다"), t(" — 얻은 열 덕분에 "), b("활동적인 생활"), t("을 이어 간다.")]), fill: C.accent1, A: gi.GiGecko },
  ];
  const w = 3.75, gap = 0.44;
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * (w + gap), it = st[i];
    card(s, x, 2.6, w, 4.25, { fill: it.fill, transparency: 87, r: 0.16 });
    await iconCircle(s, x + 0.3, 2.85, 0.95, it.I, { bg: it.bg, name: it.h });
    text(s, [b(it.h, { color: C.text2 })], x + 1.4, 2.85, 1.8, 0.95, { fontSize: 24, valign: "middle" });
    await ic(s, it.A, HEX.slate, x + w - 0.8, 2.95, 0.55, "animal");
    text(s, it.body, x + 0.3, 4.0, w - 0.55, 2.7, { fontSize: 20 });
    if (i < 2) arrow(s, x + w + 0.06, 4.7, x + w + gap - 0.06, 4.7, { col: C.accent4, w: 2.5 });
  }
  s.addNotes(NOTE("ectocycle"));
};

SL.ectoreg = async () => {
  const s = contentSlide(null, L("PART 2 · BODY TEMPERATURE", "PART 2 · 체온"), T_PART, L("Ectotherms Can Still Regulate Temperature", "외온동물도 체온을 조절한다"), "외온동물도 행동으로, 또는 매우 일정한 환경 덕분에 체온을 안정적으로 유지한다");
  s.addImage({ data: await photo(PHOTO.iguana.f, 4.6, 3.95, { pos: "centre", px: 1000 }), x: 0.6, y: 1.55, w: 4.6, h: 3.95, altText: "Green iguana basking on a rock" });
  text(s, L([b("Basking in the sun", { color: C.text2 }), t(" — a behavioral mechanism ectotherms use to regulate body temperature")], [b("햇볕 쬐기", { color: C.text2 }), t(" — 외온동물이 체온을 조절하는 행동 메커니즘")]), 0.6, 5.6, 4.6, 0.75, { fontSize: 16 });
  credit(s, PHOTO.iguana.credit, 0.6, 6.5, 4.6);
  const rx = 5.5, rw = 7.23;
  card(s, rx, 1.55, rw, 1.2, { fill: C.background2 });
  await iconCircle(s, rx + 0.25, 1.75, 0.8, fa.FaSun, { bg: C.accent3, name: "behavior" });
  text(s, L([b("Behavior  ", { color: C.text2 }), t("Insects and lizards often "), b("warm up by basking"), t(" in the sun and "), b("cool off by moving to a shady spot"), t(".")], [b("행동  ", { color: C.text2 }), t("곤충과 도마뱀은 흔히 "), b("햇볕을 쬐어 몸을 데우고"), t(", "), b("그늘로 옮겨 몸을 식힌다"), t(".")]), rx + 1.25, 1.58, rw - 1.45, 1.14, { fontSize: 18, valign: "middle" });
  card(s, rx, 2.9, rw, 2.05, { fill: C.background2 });
  await iconCircle(s, rx + 0.25, 3.2, 0.8, fa.FaFish, { bg: C.accent2, name: "pupfish" });
  text(s, L([b("Desert pupfish  ", { color: C.text2 }), t("tolerates water of "), b("2.2–45°C"), t(" (36–113°F), but can "), b("breed only within a narrow range"), t(" → swims to different areas of its pool as water temperature changes.")], [b("사막 pupfish  ", { color: C.text2 }), t("수온 "), b("2.2–45°C"), t("를 견디지만 "), b("번식은 좁은 범위에서만"), t(" 가능하다 → 수온이 바뀌면 연못의 다른 곳으로 헤엄쳐 간다.")]), rx + 1.25, 2.98, rw - 1.45, 1.25, { fontSize: 18 });
  const bx = rx + 1.25, by = 4.32, bw = rw - 1.65;
  s.addShape(S.ROUNDED_RECTANGLE, { x: bx, y: by, w: bw, h: 0.2, rectRadius: 0.1, fill: { color: C.accent2, transparency: 65 }, line: { type: "none" }, objectName: "tolerance range" });
  s.addShape(S.ROUNDED_RECTANGLE, { x: bx + bw * 0.52, y: by - 0.05, w: bw * 0.14, h: 0.3, rectRadius: 0.1, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "breeding range (schematic)" });
  text(s, "2.2°C", bx - 0.05, by + 0.26, 0.9, 0.3, { fontSize: 14, color: C.accent4 });
  text(s, "45°C", bx + bw - 0.75, by + 0.26, 0.8, 0.3, { fontSize: 14, color: C.accent4, align: "right" });
  text(s, L("breeding (narrow)", "번식 (좁은 범위)"), bx + bw * 0.38, by + 0.26, 2.2, 0.3, { fontSize: 14, color: C.accent1, bold: true, align: "center" });
  card(s, rx, 5.1, rw, 1.75, { fill: C.background2 });
  await iconCircle(s, rx + 0.25, 5.575, 0.8, md.MdWaves, { bg: C.text2, name: "deep sea" });
  text(s, L([b("Very constant environment", { color: C.text2, breakLine: true }), t("In the deep ocean the temperature stays at about "), b("0–3°C"), t(" (32–37.5°F), so ectothermic deep-sea fish experience little variation in body temperature throughout their lives.")], [b("매우 일정한 환경", { color: C.text2, breakLine: true }), t("깊은 바다는 수온이 약 "), b("0–3°C"), t("로 매우 일정해서, 외온동물인 심해어는 평생 체온 변화를 거의 겪지 않는다.")]), rx + 1.25, 5.13, rw - 1.45, 1.69, { fontSize: 18, valign: "middle" });
  s.addNotes(NOTE("ectoreg"));
};

SL.endocost = async () => {
  const s = contentSlide(null, L("PART 2 · BODY TEMPERATURE", "PART 2 · 체온"), T_PART, L("Endotherms: Benefits and Costs", "내온동물: 이점과 비용"), "내온동물 — 따뜻한 체온을 유지하는 데 따르는 이점과 비용");
  card(s, 0.6, 1.55, 4.1, 5.3, { fill: C.text2, r: 0.16 });
  text(s, L("Most endotherms keep a fairly constant body temperature of about", "대부분의 내온동물이 꽤 일정하게 유지하는 체온"), 0.9, 2.25, 3.55, 1.05, { fontSize: 20, color: C.background1 });
  text(s, "35–41°C", 0.9, 3.35, 3.55, 0.95, { fontSize: 48, bold: true, color: C.accent3, fontFace: "+mj-lt", valign: "middle" });
  text(s, "(95–106°F)", 0.9, 4.3, 3.55, 0.4, { fontSize: 18, color: C.accent5 });
  text(s, L([t("A few, such as ", { color: C.background1 }), b("hummingbirds", { color: C.accent5 }), t(", allow their body temperature to fall at night.", { color: C.background1 })], [b("벌새", { color: C.accent5 }), t("처럼 밤에 체온이 떨어지도록 두는 내온동물도 일부 있다.", { color: C.background1 })]), 0.9, 4.9, 3.55, 1.6, { fontSize: 20 });
  card(s, 5.1, 1.55, 4.2, 4.1, { fill: C.accent2, transparency: 86, r: 0.16 });
  text(s, [b(L("BENEFITS", "이점"), { color: C.accent2 })], 5.35, 1.7, 3.8, 0.45, { fontSize: 22 });
  text(s, L("A warm body usually can …", "따뜻한 몸은 대체로 …"), 5.35, 2.2, 3.8, 0.4, { fontSize: 18, italic: true, color: C.accent4 });
  const ben = [[fa.FaEye, L("sense its environment better", "환경을 더 잘 감지한다")], [fa.FaBolt, L("respond more quickly", "더 빨리 반응한다")], [fa.FaRunning, L("move faster than a cold body", "차가운 몸보다 더 빨리 움직인다")]];
  for (let i = 0; i < 3; i++) {
    await iconCircle(s, 5.4, 2.8 + i * 0.9, 0.62, ben[i][0], { bg: C.accent2, name: ben[i][1] });
    text(s, ben[i][1], 6.17, 2.78 + i * 0.9, 3.05, 0.66, { fontSize: 18, valign: "middle" });
  }
  card(s, 9.55, 2.4, 3.18, 3.25, { fill: C.accent1, transparency: 86, r: 0.16 });
  text(s, [b(L("COST", "비용"), { color: C.accent1 })], 9.8, 2.55, 2.8, 0.45, { fontSize: 22 });
  await iconCircle(s, 9.8, 3.1, 0.75, fa.FaFire, { bg: C.accent1, name: "energy cost" });
  text(s, L([t("the "), b("energy"), t(" required to maintain a high body temperature")], [t("높은 체온을 유지하는 데 드는 "), b("에너지")]), 9.8, 3.95, 2.75, 1.6, { fontSize: 18 });
  s.addShape(S.ISOSCELES_TRIANGLE, { x: 8.25, y: 5.92, w: 0.9, h: 0.75, fill: { color: C.accent4 }, line: { type: "none" }, objectName: "fulcrum" });
  s.addShape(S.LINE, { x: 5.4, y: 5.94, w: 6.6, h: 0, line: { color: C.accent4, width: 4 } });
  arrow(s, 7.2, 5.65, 7.2, 5.94, { col: C.accent4, w: 2, head: null });
  arrow(s, 11.14, 5.65, 11.14, 5.94, { col: C.accent4, w: 2, head: null });
  s.addNotes(NOTE("endocost"));
};

SL.twotypes = async () => {
  const s = contentSlide(null, L("PART 3 · NEGATIVE FEEDBACK", "PART 3 · 음성피드백"), T_PART, L("Two Types of Feedback Systems", "피드백 시스템의 두 가지 형태"), "피드백 시스템은 내부 조건을 조절한다");
  text(s, L([t("How does an animal sense the conditions within its body and adjust them when necessary? Its internal environment is maintained by "), b("feedback systems", { color: C.text2 }), t(".")], [t("동물은 몸속 상태를 어떻게 감지하고, 필요할 때 어떻게 조절할까? 내부 환경은 "), b("피드백 시스템", { color: C.text2 }), t("이라는 메커니즘으로 유지된다.")]), 0.6, 1.55, 12.13, 0.85, { fontSize: 18 });
  const cs = [
    { x: 0.6, col: C.accent2, title: L("Negative feedback", "음성피드백"), def: L([t("counteracts the effects of changes in the internal environment → "), b("principally responsible for maintaining homeostasis")], [t("내부 환경에 생긴 변화의 효과에 반대로 작용한다 → "), b("주로 항상성 유지에 이용된다")]), neg: true, label: L("returns toward the set point", "설정점으로 되돌아온다") },
    { x: 6.77, col: C.accent1, title: L("Positive feedback", "양성피드백"), def: L([t("creates cycles in which "), b("changes amplify themselves")], [b("변화가 스스로를 증폭시키는"), t(" 사이클을 만든다")]), neg: false, label: L("change keeps growing", "변화가 점점 커진다") },
  ];
  for (const c of cs) {
    const w = 5.96;
    card(s, c.x, 2.6, w, 4.25, { fill: c.col, transparency: 89, r: 0.16 });
    text(s, [b(c.title, { color: c.col, fontSize: 26 })], c.x + 0.3, 2.72, w - 0.6, 0.55, { valign: "middle" });
    text(s, c.def, c.x + 0.3, 3.32, w - 0.6, 1.0, { fontSize: 18 });
    const gx = c.x + 0.5, gy = 4.85, gw = 4.95, gh = 1.45;
    s.addShape(S.LINE, { x: gx, y: gy + gh, w: gw, h: 0, line: { color: C.accent4, width: 1 } });
    s.addShape(S.LINE, { x: gx, y: gy, w: 0, h: gh, line: { color: C.accent4, width: 1 } });
    s.addShape(S.LINE, { x: gx, y: gy + 0.9, w: gw, h: 0, line: { color: C.accent4, width: 1.25, dashType: "dash" } });
    text(s, L("set point", "설정점"), gx + gw - 1.1, gy + 0.93, 1.1, 0.28, { fontSize: 14, color: C.accent4, italic: true, align: "right" });
    if (c.neg) curve(s, gx, gy, gw, gh, [[0, 0.9], [0.75, 0.2], [1.5, 1.35], [2.25, 0.55], [3.0, 1.1], [3.75, 0.82], [4.5, 0.92], [4.95, 0.9]], { color: C.accent2, width: 3 });
    else s.addShape(S.CUSTOM_GEOMETRY, { x: gx, y: gy, w: gw, h: gh, points: [{ x: 0, y: 0.9 }, { x: 4.95, y: 0.02, curve: { type: "cubic", x1: 2.8, y1: 0.9, x2: 4.35, y2: 0.65 } }], line: { color: C.accent1, width: 3 } });
    text(s, c.label, gx, gy - 0.36, 3.6, 0.32, { fontSize: 15, color: c.col, bold: true });
    text(s, L("time →", "시간 →"), gx + gw - 0.9, gy + gh + 0.03, 0.9, 0.28, { fontSize: 14, color: C.accent4, align: "right" });
    text(s, L("schematic graph", "개념도"), gx, gy + gh + 0.03, 2.2, 0.28, { fontSize: 12, italic: true, color: C.accent4 });
  }
  s.addNotes(NOTE("twotypes"));
};

SL.negfb = async () => {
  const s = contentSlide(null, L("PART 3 · NEGATIVE FEEDBACK", "PART 3 · 음성피드백"), T_PART, L("Negative Feedback Reverses Change", "음성피드백은 변화의 효과를 역전시킨다"), "음성피드백은 변화의 효과를 역전시킨다 — 항상성을 조절하는 가장 중요한 메커니즘");
  const lw = 6.85;
  text(s, L([b("Negative feedback", { color: C.accent2 }), t(": a change causes responses that "), b("counteract the change"), t(" → the system returns to its original condition.")], [b("음성피드백", { color: C.accent2 }), t(": 변화가 그 변화와 "), b("반대 방향의 반응"), t("을 일으켜 → 시스템이 처음 상태로 돌아간다.")]), 0.6, 1.55, lw, 0.95, { fontSize: 18 });
  text(s, L("Found in living and nonliving systems:", "생물계와 무생물계 모두에서 볼 수 있다:"), 0.6, 2.6, lw, 0.4, { fontSize: 18, bold: true, color: C.text2 });
  text(s, bl(L([
    [b("Body: "), t("body temperature; blood levels of glucose, hormones, water, salts, and oxygen; even standing upright")],
    [b("Machines: "), t("house temperature, filling a toilet tank, cruise control")],
  ], [
    [b("몸: "), t("체온, 혈액 속 포도당·호르몬·물·염분·산소의 수준, 똑바로 서 있기까지")],
    [b("기계: "), t("집 안 온도 유지, 변기 탱크 채우기, 크루즈 컨트롤")],
  ]), { gap: 4 }), 0.6, 3.02, lw, 1.38, { fontSize: 18 });
  text(s, L("All negative feedback systems contain three principal components:", "모든 음성피드백 시스템의 세 가지 구성요소:"), 0.6, 4.5, lw, 0.65, { fontSize: 18, bold: true, color: C.text2 });
  const comp = [
    { en: L("Sensor", "감지기"), d: L("detects the current condition", "현재 상태를 감지한다"), I: md.MdSensors },
    { en: L("Control center", "조절 센터"), d: L("compares it to the set point", "설정점과 비교한다"), I: fa.FaBrain },
    { en: L("Effector", "효과기"), d: L("produces an output that restores the condition", "상태를 되돌리는 출력을 만든다"), I: fa.FaCogs },
  ];
  for (let i = 0; i < 3; i++) {
    const y = 5.25 + i * 0.53;
    await iconCircle(s, 0.6, y + 0.02, 0.46, comp[i].I, { bg: C.accent2, name: comp[i].en });
    text(s, [b(comp[i].en, { color: C.accent2 }), t(" — " + comp[i].d)], 1.2, y, lw - 0.6, 0.5, { fontSize: 16, valign: "middle" });
  }
  const f = await figure(FIG.ap110.f);
  const cx0 = 7.65, cw0 = 5.08, fw = 4.9, fh = fw / f.ar;
  card(s, cx0, 1.55, cw0, 5.3, { fill: C.background1, line: C.background2, r: 0.12 });
  // the figure labels its own panels (a) and (b), so it needs no caption; centered above the credit line
  s.addImage({ data: f.data, x: cx0 + (cw0 - fw) / 2, y: 1.55 + (4.85 - fh) / 2, w: fw, h: fh, altText: "OpenStax A&P Figure 1.10: negative feedback loop and body temperature regulation" });
  credit(s, FIG.ap110.credit, cx0 + 0.15, 6.5, cw0 - 0.3, { align: "center" });
  s.addNotes(NOTE("negfb"));
};

SL.fig322 = async () => {
  const s = contentSlide(null, L("PART 3 · NEGATIVE FEEDBACK", "PART 3 · 음성피드백"), T_PART, L("Negative Feedback in Action (Fig. 32-2)", "음성피드백의 작동 (그림 32-2)"), "음성피드백은 항상성을 유지한다 — 집의 온도와 체온 조절");
  const Lx = 1.25, Lw = 3.35, Cx = 4.85, Cw = 3.6, Rx = 8.7, Rw = 4.03;
  const mid = Cx + Cw / 2;
  await ic(s, fa.FaHome, HEX.slate, Lx + Lw / 2 - 0.7, 1.55, 0.42, "house");
  text(s, [b(L("House", "집"), { color: C.text2 })], Lx + Lw / 2 - 0.18, 1.55, 1.6, 0.42, { fontSize: 20, valign: "middle" });
  await ic(s, fa.FaChild, HEX.slate, Rx + Rw / 2 - 0.62, 1.55, 0.42, "body");
  text(s, [b(L("Body", "몸"), { color: C.text2 })], Rx + Rw / 2 - 0.1, 1.55, 1.6, 0.42, { fontSize: 20, valign: "middle" });
  await ic(s, fa.FaSnowflake, HEX.teal, mid - 1.12, 1.57, 0.38, "cold");
  text(s, [b(L("stimulus: cold", "자극: 추위"), { color: C.accent2 })], mid - 0.66, 1.55, 2.3, 0.42, { fontSize: 18, valign: "middle" });
  s.addText([b(L("CONDITION", "조건"), { color: C.text2, breakLine: true }), t(L("room or body temperature", "실내 온도 또는 체온"), { fontSize: 15, color: C.text1 })], { shape: S.OVAL, x: Cx + 0.15, y: 2.08, w: Cw - 0.3, h: 0.78, fill: { color: C.accent5, transparency: 45 }, line: { color: C.accent2, width: 1.25 }, fontSize: 17, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "condition" });
  arrow(s, mid, 1.97, mid, 2.08, { col: C.accent2, w: 2.5 });
  const rh = 0.72;
  const rows = [
    { y: 3.3, en: L("SENSOR", "감지기"), L: [fa.FaThermometerHalf, L("thermometer", "온도계")], R: [fa.FaHandPaper, L("temperature receptors in skin", "피부의 온도 수용체")] },
    { y: 4.5, en: L("CONTROL CENTER with SET POINT", "설정점을 가진 조절 센터"), L: [md.MdThermostat, L("thermostat", "온도조절기")], R: [fa.FaBrain, L("hypothalamus", "시상하부")] },
    { y: 5.7, en: L("EFFECTOR", "효과기"), L: [gi.GiFurnace, L("furnace output: heat", "난방 장치 출력: 열")], R: [gi.GiMuscleUp, L("muscle output: shivering → heat", "근육 출력: 떨림 → 열")] },
  ];
  for (const r of rows) {
    s.addText([b(r.en, { color: C.background1 })], { shape: S.ROUNDED_RECTANGLE, x: Cx, y: r.y, w: Cw, h: rh, rectRadius: 0.1, fill: { color: C.text2 }, line: { type: "none" }, fontSize: 17, align: "center", valign: "middle", margin: 0.05, isTextBox: true, objectName: r.en });
    for (const [side, X, W] of [["L", Lx, Lw], ["R", Rx, Rw]]) {
      const [I, label] = r[side];
      card(s, X, r.y, W, rh, { fill: C.background2, r: 0.1 });
      await iconCircle(s, X + 0.1, r.y + 0.11, 0.5, I, { bg: side === "L" ? C.accent4 : C.accent2, name: label });
      text(s, [b(label)], X + 0.72, r.y + 0.02, W - 0.8, rh - 0.04, { fontSize: 17, valign: "middle" });
      if (side === "L") arrow(s, X + W + 0.03, r.y + rh / 2, Cx - 0.03, r.y + rh / 2, { col: C.accent4, w: 1, head: null, dash: "sysDot" });
      else arrow(s, Cx + Cw + 0.03, r.y + rh / 2, X - 0.03, r.y + rh / 2, { col: C.accent4, w: 1, head: null, dash: "sysDot" });
    }
  }
  const lbl = [
    { y1: 2.86, y2: 3.3, t1: L("measured by", "측정") },
    { y1: 4.02, y2: 4.5, t1: L("sends data to", "결과 전송") },
    { y1: 5.22, y2: 5.7, t1: L("if the measured temperature differs from the set point, activates", "감지기가 측정한 온도와 설정점이 다르면 작동") },
  ];
  for (const l of lbl) {
    arrow(s, mid, l.y1, mid, l.y2, { col: C.accent2, w: 2.5 });
    text(s, [t(l.t1, { italic: true })], mid + 0.15, l.y1 + 0.01, 12.73 - mid - 0.15, l.y2 - l.y1 - 0.02, { fontSize: 15, valign: "middle" });
  }
  const by = 6.55, lx = 0.85, cy = 2.47;
  polyline(s, [[mid, 5.7 + rh], [mid, by], [lx, by], [lx, cy], [Cx + 0.17, cy]], { col: C.accent2, w: 2.75 });
  text(s, [b(L("Effector changes the conditions to match set point", "효과기는 설정점을 맞추기 위해 조건을 변경"), { color: C.accent2 })], 1.05, by + 0.03, 8.0, 0.3, { fontSize: 16, valign: "middle" });
  s.addNotes(NOTE("fig322"));
};

SL.hypo = async () => {
  const s = contentSlide(null, L("PART 3 · NEGATIVE FEEDBACK", "PART 3 · 음성피드백"), T_PART, L("Hypothalamus: The Body’s Thermostat", "시상하부: 몸의 온도조절기"), "시상하부의 체온 조절 — 체온이 설정점(정상 체온)보다 낮을 때와 높을 때");
  text(s, L([b("Hypothalamus", { color: C.text2 }), t(" — a region deep in the brain; the key homeostatic regulator in vertebrates. In birds and mammals it coordinates "), b("nervous signals, hormone release, and behaviors"), t(". Nerve endings in the skin and elsewhere act as temperature sensors.")], [b("시상하부", { color: C.text2 }), t(" — 뇌 깊숙한 곳에 있는, 척추동물의 가장 중요한 항상성 조절자. 조류와 포유류에서 "), b("신경 신호, 호르몬 방출, 행동"), t("을 조정한다. 피부와 몸 곳곳의 신경말단이 온도 감지기 역할을 한다.")]), 0.6, 1.55, 12.13, 1.0, { fontSize: 18 });
  const pw = 3.95, py = 2.72, ph = 3.38;
  card(s, 0.6, py, pw, ph, { fill: C.accent5, transparency: 60, r: 0.16 });
  text(s, [b(L("Below set point", "설정점보다 낮을 때"), { color: C.text2, fontSize: 20, breakLine: true }), t(L("generate & conserve heat", "열을 만들고 지킨다"), { color: C.accent2, bold: true })], 0.8, py + 0.1, pw - 0.4, 0.75, { fontSize: 16 });
  text(s, bl(L([
    [b("Shivering"), t(" — small, rapid muscular contractions")],
    [t("Moving to a warmer place")],
    [t("Vessels to "), b("nonvital areas"), t(" (face, hands, feet, skin) "), b("constrict"), t(" → less heat loss; warm blood diverted to the body core")],
  ], [
    [b("떨림"), t(" — 작고 빠른 근육 수축")],
    [t("따뜻한 곳으로 이동")],
    [b("덜 중요한 부위"), t("(얼굴, 손발, 피부)의 혈관이 "), b("수축"), t(" → 열 손실 감소, 따뜻한 혈액은 몸 중심부로")],
  ]), { gap: 4, indent: 12 }), 0.8, py + 0.9, pw - 0.35, ph - 1.0, { fontSize: 16 });
  const hx = 13.333 - 0.6 - pw;
  card(s, hx, py, pw, ph, { fill: C.accent1, transparency: 85, r: 0.16 });
  text(s, [b(L("Above set point", "설정점보다 높을 때"), { color: C.text2, fontSize: 20, breakLine: true }), t(L("reduce body temperature", "체온을 낮춘다"), { color: C.accent1, bold: true })], hx + 0.2, py + 0.1, pw - 0.4, 0.75, { fontSize: 16 });
  text(s, bl(L([
    [t("More blood "), b("diverted to the skin"), t(" → heat radiated to the surrounding air")],
    [b("Sweat glands"), t(": watery secretion cools the body as it evaporates")],
    [t("Fatigue & discomfort → rest, seek shade or cooling water")],
  ], [
    [t("피부로 가는 혈액 "), b("증가"), t(" → 열을 주변 공기로 방출")],
    [b("땀샘"), t(": 땀이 증발하면서 몸을 식힌다")],
    [t("피로감과 불편함 → 쉬거나 그늘·시원한 물을 찾는다")],
  ]), { gap: 4, indent: 12 }), hx + 0.2, py + 0.9, pw - 0.35, ph - 1.0, { fontSize: 16 });
  const f = await figure(FIG.bio3323.f);
  const fw = 3.95, fh = fw / f.ar, fx = (13.333 - fw) / 2;
  s.addImage({ data: f.data, x: fx, y: py + 0.05, w: fw, h: fh, altText: "OpenStax Biology 2e Figure 33.23: body temperature falls vs rises" });
  credit(s, FIG.bio3323.credit, fx - 0.1, py + fh + 0.12, fw + 0.2, { align: "center" });
  card(s, 0.6, 6.22, 12.13, 0.62, { fill: C.text2, r: 0.31 });
  text(s, L([t("When normal body temperature is restored, the hypothalamus ", { color: C.background1 }), b("switches off", { color: C.accent3 }), t(" these responses (negative feedback).", { color: C.background1 })], [t("체온이 정상으로 돌아오면 시상하부는 이런 반응을 ", { color: C.background1 }), b("멈춘다", { color: C.accent3 }), t(" (음성피드백).", { color: C.background1 })]), 0.85, 6.22, 11.7, 0.62, { fontSize: 17, valign: "middle" });
  s.addNotes(NOTE("hypo"));
};

SL.water = async () => {
  const s = contentSlide(null, "HAVE YOU EVER WONDERED …", T_WOND, L("Can You Drink Too Much Water?", "물도 너무 많이 마시면 해로울까?"), "물도 너무 많이 마시면 해로울까?");
  card(s, 0.6, 1.55, 4.0, 5.3, { fill: C.accent2, r: 0.16 });
  await ic(s, fa.FaTint, HEX.white, 1.0, 1.85, 0.8, "water drop");
  text(s, "13%", 1.0, 2.75, 3.3, 1.1, { fontSize: 66, bold: true, color: C.background1, fontFace: "+mj-lt" });
  text(s, L("of Boston Marathon runners tested at the finish line had at least mild hyponatremia — a few had dangerously low blood Na⁺.", "보스턴 마라톤 결승선에서 검사한 참가자 가운데 적어도 가벼운 저나트륨혈증을 보인 비율 — 몇몇은 혈중 Na⁺가 위험할 만큼 낮았다."), 1.0, 3.95, 3.3, 2.7, { fontSize: 18, color: C.background1 });
  const x0 = 4.95, w0 = 7.78;
  text(s, L([t("Too much of almost anything — "), b("even water"), t(" — can be harmful. “Hydration” has become a buzzword that sometimes encourages people to drink considerably more than they need.")], [t("대부분의 것이 그렇듯 "), b("물도"), t(" 지나치면 해롭다. ‘수분 보충’이 유행어가 되면서 필요보다 훨씬 많이 마시게 되기도 한다.")]), x0, 1.55, w0, 0.95, { fontSize: 18 });
  card(s, x0, 2.6, w0, 1.3, { fill: C.background2 });
  text(s, L([b("Zyrees Oliver", { color: C.text2 }), t(", 17, a Georgia high-school football player, reportedly drank "), b("2 gallons of water + 2 gallons of Gatorade"), t(" during practice to relieve muscle cramps — he collapsed at home and died.")], [b("Zyrees Oliver", { color: C.text2 }), t("(17세, 조지아주 고등학교 미식축구 선수)는 연습 중 근육 경련을 풀려고 "), b("물 2갤런과 게토레이 2갤런"), t("을 마셨다 — 집에서 쓰러져 숨졌다.")]), x0 + 0.25, 2.65, w0 - 0.5, 1.2, { fontSize: 18, valign: "middle" });
  chip(s, x0, 4.05, L("hyponatremia", "저나트륨혈증"), L("(literally, “too little sodium”)", "(말 그대로 ‘너무 적은 나트륨’)"), { col: C.accent1, size: 18, h: 0.52 });
  text(s, L([t("Excess water dilutes the sodium ions (Na⁺) in the blood. Na⁺ is essential for "), b("neuronal signaling and muscle contraction"), t(", so elaborate homeostatic mechanisms keep it within a very narrow range.")], [t("너무 많은 물이 혈액의 나트륨이온(Na⁺)을 묽게 만든다. Na⁺는 "), b("신경 신호 전달과 근육 수축"), t("에 꼭 필요해서, 정교한 항상성 메커니즘이 아주 좁은 범위로 유지한다.")]), x0, 4.72, w0, 1.25, { fontSize: 18 });
  text(s, L([b("Symptoms: ", { color: C.accent1 }), t("nausea and vomiting, muscle cramps, convulsions, unconsciousness, coma, occasionally death")], [b("증상: ", { color: C.accent1 }), t("메스꺼움과 구토, 근육 경련, 발작, 의식 상실, 혼수, 때로는 사망")]), x0, 6.05, w0, 0.78, { fontSize: 18 });
  s.addNotes(NOTE("water"));
};

SL.vicious = async () => {
  const s = contentSlide(null, L("CASE STUDY · CONTINUED", "CASE STUDY · 계속"), T_CASE, L("Case Study: When Feedback Fails", "피드백이 실패할 때"), "열사병 — 음성피드백 메커니즘이 제대로 작동하지 않는다");
  card(s, 0.6, 1.55, 5.2, 5.3, { fill: C.background2, r: 0.16 });
  text(s, [b(L("Negative feedback breaks down", "음성피드백이 무너진다"), { color: C.text2 })], 0.85, 1.75, 4.75, 0.5, { fontSize: 22 });
  const fails = [
    { I: fa.FaBed, h: L("Exhaustion should make you rest …", "지치면 쉬어야 하지만 …"), d: L("but heat stroke causes mental confusion, which disrupts this feedback — athletes ignore the body’s warning signs.", "열사병으로 정신이 혼란해지면 이 피드백이 방해받아, 선수들은 몸의 경고 신호를 무시한다.") },
    { I: fa.FaTint, h: L("Sweating may stop", "땀이 멈출 수도 있다"), d: L("skyrocketing body temperature makes the brain’s temperature control systems malfunction.", "체온이 치솟으면 뇌의 체온 조절 시스템이 오작동한다.") },
  ];
  for (let i = 0; i < 2; i++) {
    const y = 2.55 + i * 2.05;
    await iconCircle(s, 0.85, y + 0.05, 0.75, fails[i].I, { bg: C.accent4, name: fails[i].h });
    text(s, [b(fails[i].h, { breakLine: true, fontSize: 18 }), t(fails[i].d, { fontSize: 16 })], 1.8, y, 3.85, 1.9, {});
  }
  // side nodes pushed out so the centre label fits between them (widest node text 2.18 in, label 1.72 in at these sizes)
  const cx = 9.35, cy = 4.1, rx = 2.15, ry = 1.6, nw = 2.35, nh = 0.85;
  const nodes = [
    { x: cx, y: cy - ry, en: L("Body temperature ↑", "체온 상승") },
    { x: cx + rx, y: cy, en: L("Enzyme activity ↑", "효소 활성 증가") },
    { x: cx, y: cy + ry, en: L("Metabolic rate ↑", "대사율 상승") },
    { x: cx - rx, y: cy, en: L("Still more heat", "더 많은 열 발생") },
  ];
  for (const n of nodes) {
    s.addText([b(n.en, { color: C.background1 })], { shape: S.ROUNDED_RECTANGLE, x: n.x - nw / 2, y: n.y - nh / 2, w: nw, h: nh, rectRadius: 0.14, fill: { color: C.accent1 }, line: { type: "none" }, fontSize: 18, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "cycle node: " + n.en });
  }
  const A = (x1, y1, x2, y2) => arrow(s, x1, y1, x2, y2, { col: C.accent1, w: 3 });
  A(cx + nw / 2 + 0.05, cy - ry + 0.1, cx + rx - 0.25, cy - nh / 2 - 0.05);
  A(cx + rx - 0.25, cy + nh / 2 + 0.05, cx + nw / 2 + 0.05, cy + ry - 0.1);
  A(cx - nw / 2 - 0.05, cy + ry - 0.1, cx - rx + 0.25, cy + nh / 2 + 0.05);
  A(cx - rx + 0.25, cy - nh / 2 - 0.05, cx - nw / 2 - 0.05, cy - ry + 0.1);
  text(s, [b(L("vicious cycle", "악순환"), { color: C.accent1, breakLine: true, fontSize: 20 }), t(L("= positive feedback", "= 양성피드백"), { fontSize: 16, italic: true })], cx - 0.95, cy - 0.45, 1.9, 0.9, { align: "center", valign: "middle" });
  card(s, 6.2, 6.3, 6.53, 0.52, { fill: C.text2, r: 0.26 });
  text(s, L([t("Not terminated by ", { color: C.background1 }), b("immediate cooling", { color: C.accent5 }), t(" → may lead to death", { color: C.background1 })], [b("즉시 몸을 식혀", { color: C.accent5 }), t(" 끊지 않으면 → 사망에 이를 수 있다", { color: C.background1 })]), 6.3, 6.3, 6.33, 0.52, { fontSize: 16, align: "center", valign: "middle" });
  s.addNotes(NOTE("vicious"));
};

SL.posfb = async () => {
  const s = contentSlide(null, L("PART 4 · POSITIVE FEEDBACK", "PART 4 · 양성피드백"), T_PART, L("Positive Feedback Amplifies Change", "양성피드백은 변화의 효과를 촉진시킨다"), "양성피드백은 변화의 효과를 촉진시킨다");
  card(s, 0.6, 1.55, 12.13, 0.9, { fill: C.accent1, transparency: 86, r: 0.16 });
  text(s, L([b("Positive feedback", { color: C.accent1 }), t(": a change produces a response that "), b("amplifies that change"), t(".")], [b("양성피드백", { color: C.accent1 }), t(": 변화가 "), b("그 변화를 더 키우는"), t(" 반응을 일으킨다.")]), 0.9, 1.55, 11.6, 0.9, { fontSize: 22, valign: "middle" });
  text(s, [b(L("Imagine a thermostat gone haywire", "고장 난 온도조절기를 떠올려 보자"), { color: C.text2 })], 0.6, 2.65, 6.3, 0.45, { fontSize: 20 });
  const steps = L(["You nudge up the setting", "Room reaches the set point", "Instead of turning off: “more heat!”", "… and so on"], ["설정을 살짝 올린다", "방이 설정 온도에 도달한다", "꺼지는 대신 ‘열을 더!’", "… 계속 반복"]);
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + i * 1.55, h = 1.3 + i * 0.62, y = 6.8 - h;
    s.addShape(S.ROUNDED_RECTANGLE, { x, y, w: 1.4, h, rectRadius: 0.1, fill: { color: C.accent1, transparency: 70 - i * 18 }, line: { type: "none" }, objectName: "step " + (i + 1) });
    text(s, steps[i], x + 0.08, y + 0.08, 1.24, h - 0.16, { fontSize: 16, bold: i === 2, color: i >= 2 ? C.background1 : C.text1, align: "center" });
  }
  await ic(s, fa.FaFire, HEX.coral, 5.73, 3.12, 0.45, "more heat");
  text(s, L("a very uncomfortable situation", "아주 불편한 상황"), 1.9, 3.12, 3.75, 0.45, { fontSize: 16, italic: true, color: C.accent4, align: "right", valign: "middle" });
  card(s, 7.1, 2.65, 5.63, 4.15, { fill: C.text2, r: 0.16 });
  text(s, [b(L("Why is it rare?", "왜 드물까?"), { color: C.accent3 })], 7.4, 2.8, 5.0, 0.5, { fontSize: 24 });
  text(s, bl(L([
    [t("Its effects could "), b("spiral out of control", { color: C.accent3 }), t(" → relatively rare in biological systems.")],
    [t("Under normal physiological conditions, positive feedback events are "), b("eventually halted by negative feedback", { color: C.accent5 }), t(".")],
    [t("It also occurs in nonliving systems (→ Earth Watch).")],
  ], [
    [t("효과가 "), b("통제 불능", { color: C.accent3 }), t("이 될 수 있어 → 생물에서는 비교적 드물다.")],
    [t("정상적인 생리 조건에서는 양성피드백도 "), b("결국 음성피드백에 의해 멈춘다", { color: C.accent5 }), t(".")],
    [t("무생물 시스템에서도 일어난다 (→ Earth Watch).")],
  ]), { gap: 8, indent: 14 }), 7.4, 3.45, 5.05, 3.2, { fontSize: 18, color: C.background1 });
  s.addNotes(NOTE("posfb"));
};

SL.birth = async () => {
  const s = contentSlide(null, L("PART 4 · POSITIVE FEEDBACK", "PART 4 · 양성피드백"), T_PART, L("Childbirth: A Positive Feedback Loop", "출산: 양성피드백 고리"), "출산 — 사람과 다른 포유류에서 양성피드백의 대표적인 사례");
  const f = await figure(FIG.ap111.f);
  const fh = 4.5, fw = fh * f.ar;
  s.addImage({ data: f.data, x: 0.6, y: 1.6, w: fw, h: fh, altText: "OpenStax A&P Figure 1.11: positive feedback loop of childbirth" });
  credit(s, FIG.ap111.credit, 0.6, 6.25, fw);
  const x0 = 0.6 + fw + 0.35, w0 = 12.73 - x0;
  const steps = L([
    [t("Early labor contractions push the baby’s head against the "), b("cervix"), t(" → it stretches and begins to open")],
    [t("Nerve cells in the cervix signal the "), b("hypothalamus")],
    [t("The hypothalamus triggers release of "), b("oxytocin", { color: C.accent1 })],
    [t("Oxytocin → stronger "), b("uterine contractions"), t(" → cervix stretched even further → more oxytocin …")],
  ], [
    [t("진통 초기의 수축이 아기의 머리를 "), b("자궁경부"), t(" 쪽으로 민다 → 자궁경부가 늘어나며 열리기 시작한다")],
    [t("자궁경부의 신경세포가 "), b("시상하부"), t("에 신호를 보낸다")],
    [t("시상하부가 "), b("옥시토신", { color: C.accent1 }), t(" 분비를 일으킨다")],
    [t("옥시토신 → 더 강한 "), b("자궁수축"), t(" → 자궁경부가 더 늘어남 → 옥시토신 더 분비 …")],
  ]);
  for (let i = 0; i < 4; i++) {
    const y = 1.55 + i * 0.82;
    s.addText(String(i + 1), { shape: S.OVAL, x: x0, y: y + 0.13, w: 0.48, h: 0.48, fill: { color: C.accent1 }, line: { type: "none" }, fontSize: 17, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    text(s, steps[i], x0 + 0.62, y, w0 - 0.62, 0.75, { fontSize: 18, valign: "middle" });
  }
  card(s, x0, 4.95, w0, 1.9, { fill: C.accent2, transparency: 86 });
  text(s, L([b("How it ends: ", { color: C.accent2 }), t("the contractions cause the "), b("birth of the baby"), t(", relieving pressure on the cervix — "), b("negative feedback", { color: C.accent2 }), t(" ends the cycle.")], [b("끝나는 방식: ", { color: C.accent2 }), b("아기가 태어나면"), t(" 자궁경부를 누르던 압력이 사라진다 — "), b("음성피드백", { color: C.accent2 }), t("이 고리를 끝낸다.")]), x0 + 0.2, 4.98, w0 - 0.35, 0.95, { fontSize: 18, valign: "middle" });
  text(s, L([b("Also: let-down reflex", { color: C.text2 }), t(" — milk release during breastfeeding; stops when the full baby stops suckling.")], [b("하강반사", { color: C.text2 }), t("도 양성피드백 — 수유 때 젖이 나오는 반사로, 아기가 배불러 젖을 그만 빨면 멈춘다.")]), x0 + 0.2, 5.92, w0 - 0.35, 0.85, { fontSize: 16, valign: "middle" });
  s.addNotes(NOTE("birth"));
};

SL.arctic1 = async () => {
  const s = contentSlide(null, "EARTH WATCH", T_EARTH, L("Positive Feedback in the Arctic", "북극의 양성피드백"), "북극의 양성피드백 — 줄어드는 북극의 얼음");
  // NASA satellite maps of the September minimum (1984 vs 2012) on top; polar bear photo underneath
  s.addImage({ path: FIG.nasaArctic.f, x: 0.65, y: 1.55, w: 5.5, h: 1.8, altText: "NASA satellite maps of September Arctic sea-ice minimum, 1984 (left) and 2012 (right)" });
  text(s, L("September 1984", "1984년 9월"), 0.78, 1.6, 2.5, 0.34, { fontSize: 16, bold: true, color: HEX.white });
  text(s, L("September 2012", "2012년 9월"), 3.58, 1.6, 2.5, 0.34, { fontSize: 16, bold: true, color: HEX.white });
  text(s, L(FIG.nasaArctic.credit, "9월 해빙 최소 면적 · NASA Earth Observatory (퍼블릭 도메인)"), 0.6, 3.4, 5.6, 0.28, { fontSize: 12, italic: true, color: C.accent4 });
  // crop from the top: the bear's head is near the top edge of the photo, so a centred crop cuts it off
  s.addImage({ data: await photo(PHOTO.bear.f, 5.6, 2.3, { pos: "north", px: 1200 }), x: 0.6, y: 3.78, w: 5.6, h: 2.3, altText: "Polar bear on Arctic drift ice" });
  text(s, L("Polar bears depend on Arctic ice to hunt and rest.", "북극곰은 사냥하고 쉬기 위해 북극의 얼음에 의존한다."), 0.6, 6.13, 5.6, 0.36, { fontSize: 16, italic: true });
  credit(s, PHOTO.bear.credit, 0.6, 6.55, 5.6);
  const x0 = 6.55, w0 = 6.18;
  text(s, L([t("Climate scientists are convinced that global climate change has triggered positive feedback in the Arctic — the result may be the "), b("planetary equivalent of hyperthermia", { color: C.accent1 }), t(".")], [t("기후 과학자들은 지구 기후 변화가 북극에서 양성피드백을 일으켰다고 확신한다 — 그 결과는 "), b("지구 규모의 고체온증", { color: C.accent1 }), t("이 될 수 있다.")]), x0, 1.55, w0, 1.25, { fontSize: 18 });
  text(s, L([t("The North Pole (≈400 miles from the nearest land) was formerly covered with ice "), b("2–5 m"), t(" (6–15 ft) thick.")], [t("가장 가까운 육지에서 약 640 km 떨어진 북극점은 예전에 "), b("2–5 m"), t(" 두께의 얼음으로 덮여 있었다.")]), x0, 2.85, w0, 0.95, { fontSize: 18 });
  const stats = [
    { big: "2×", d: L("faster Arctic warming than the global average", "지구 평균보다 빠른 북극의 온난화"), col: C.accent1 },
    { big: "50%", d: L("thinner ice cap in the past 30 years", "지난 30년간 얇아진 얼음 두께"), col: C.accent2 },
    { big: "35%", d: L("smaller in area in the past 30 years", "지난 30년간 줄어든 얼음 면적"), col: C.accent2 },
    { big: "10%", d: L("average shrinkage per decade (currently)", "현재 10년마다 평균 감소율"), col: C.text2 },
  ];
  for (let i = 0; i < 4; i++) {
    const x = x0 + (i % 2) * 3.14, y = 3.95 + Math.floor(i / 2) * 1.5;
    card(s, x, y, 3.04, 1.38, { fill: C.background2 });
    text(s, stats[i].big, x + 0.2, y + 0.08, 2.64, 0.6, { fontSize: 34, bold: true, color: stats[i].col, fontFace: "+mj-lt", valign: "middle" });
    text(s, stats[i].d, x + 0.2, y + 0.68, 2.7, 0.66, { fontSize: 16 });
  }
  s.addNotes(NOTE("arctic1"));
};

SL.arctic2 = async () => {
  const s = contentSlide(null, "EARTH WATCH", T_EARTH, L("Ice Reflects, Water Absorbs", "얼음은 반사하고, 바닷물은 흡수한다"), "얼음은 햇빛을 반사하고, 바닷물은 흡수한다 — 전형적인 양성피드백");
  text(s, L("Share of incoming sunlight reflected back into space:", "들어오는 햇빛 중 우주로 반사되는 비율:"), 0.6, 1.55, 5.75, 0.75, { fontSize: 18, bold: true, color: C.text2 });
  const refl = [
    { name: L("Snow cover on ice", "얼음을 덮은 눈"), val: L("as much as 80%", "최대 80%"), frac: 0.8, col: C.accent5 },
    { name: L("Sea ice", "해빙"), val: "50–70%", frac: 0.7, frac0: 0.5, col: C.accent2 },
    { name: L("Seawater", "바닷물"), val: L("only 5–10%", "5–10%뿐"), frac: 0.1, frac0: 0.05, col: C.text2 },
  ];
  for (let i = 0; i < 3; i++) {
    const y = 2.45 + i * 1.15, r = refl[i];
    text(s, [b(r.name)], 0.6, y, 2.1, 0.85, { fontSize: 18, valign: "middle" });
    const bx = 2.75, bw = 2.9;
    s.addShape(S.RECTANGLE, { x: bx, y: y + 0.08, w: bw, h: 0.4, fill: { color: C.background2 }, line: { type: "none" } });
    if (r.frac0) {
      s.addShape(S.RECTANGLE, { x: bx, y: y + 0.08, w: bw * r.frac0, h: 0.4, fill: { color: r.col }, line: { type: "none" } });
      s.addShape(S.RECTANGLE, { x: bx + bw * r.frac0, y: y + 0.08, w: bw * (r.frac - r.frac0), h: 0.4, fill: { color: r.col, transparency: 55 }, line: { type: "none" } });
    } else s.addShape(S.RECTANGLE, { x: bx, y: y + 0.08, w: bw * r.frac, h: 0.4, fill: { color: r.col }, line: { type: "none" } });
    text(s, r.val, bx, y + 0.52, bw, 0.34, { fontSize: 16, bold: true });
  }
  text(s, L("→ nearly all solar energy hitting seawater is absorbed, warming the Arctic water.", "→ 바닷물에 닿는 햇빛 에너지는 거의 다 흡수되어 북극의 바닷물을 데운다."), 0.6, 5.9, 5.6, 0.9, { fontSize: 16, italic: true });
  // side nodes pushed out so the centre label fits between them
  const cx = 9.35, cy = 4.15, rX = 2.1, rY = 1.55, nw = 2.4, nh = 0.9;
  const nodes = [
    { x: cx, y: cy - rY, en: L("Arctic ice melts", "북극의 얼음이 녹는다"), col: C.accent2 },
    { x: cx + rX, y: cy, en: L("More water exposed", "더 많은 물이 드러난다"), col: C.text2 },
    { x: cx, y: cy + rY, en: L("Water absorbs more sunlight", "물이 햇빛을 더 흡수한다"), col: C.text2 },
    { x: cx - rX, y: cy, en: L("Water becomes warmer", "물이 더 따뜻해진다"), col: C.accent1 },
  ];
  for (const n of nodes) s.addText([b(n.en, { color: C.background1 })], { shape: S.ROUNDED_RECTANGLE, x: n.x - nw / 2, y: n.y - nh / 2, w: nw, h: nh, rectRadius: 0.14, fill: { color: n.col }, line: { type: "none" }, fontSize: 18, align: "center", valign: "middle", margin: 0.05, isTextBox: true, objectName: "loop node: " + n.en });
  const A = (x1, y1, x2, y2) => arrow(s, x1, y1, x2, y2, { col: C.accent1, w: 3 });
  A(cx + nw / 2 + 0.05, cy - rY + 0.1, cx + rX - 0.25, cy - nh / 2 - 0.05);
  A(cx + rX - 0.25, cy + nh / 2 + 0.05, cx + nw / 2 + 0.05, cy + rY - 0.1);
  A(cx - nw / 2 - 0.05, cy + rY - 0.1, cx - rX + 0.25, cy + nh / 2 + 0.05);
  A(cx - rX + 0.25, cy - nh / 2 - 0.05, cx - nw / 2 - 0.05, cy - rY + 0.1);
  await ic(s, fa.FaSun, HEX.amber, cx - 0.33, cy - 0.66, 0.66, "sun");
  text(s, [b(L("positive", "양성"), { color: C.accent1, breakLine: true }), b(L("feedback", "피드백"), { color: C.accent1 })], cx - 0.8, cy + 0.02, 1.6, 0.66, { fontSize: 18, align: "center" });
  text(s, L("“… and so on — a classic example of positive feedback, with unpredictable global consequences.”", "“… 이런 일이 계속된다 — 예측할 수 없는 전 지구적 결과를 낳을 수 있는 전형적인 양성피드백이다.”"), 6.25, L(6.28, 6.2), 6.48, L(0.56, 0.64), { fontSize: 16, italic: true, color: C.accent4, align: "center" });
  s.addNotes(NOTE("arctic2"));
};

SL.revisited = async () => {
  const s = contentSlide(null, L("CASE STUDY · REVISITED", "CASE STUDY · 다시 보기"), T_CASE, L("Case Study Revisited: Who Is at Risk?", "누가 위험한가?"), "열사병은 운동선수만의 문제가 아니다");
  text(s, L([t("When the air temperature becomes extremely high — especially with "), b("high humidity"), t(" — people can die of heat stroke while sitting in their living rooms.")], [t("기온이 아주 높고 특히 "), b("습도까지 높으면"), t(", 거실에 앉아 있다가도 열사병으로 숨질 수 있다.")]), 0.6, 1.55, 12.13, 0.75, { fontSize: 18 });
  const grp = [
    { I: md.MdElderly, en: L("The elderly", "노인"), d: L("far more likely than athletes to die of heat stroke, during summer heat waves", "여름 폭염 때 운동선수보다 열사병으로 숨질 가능성이 훨씬 높다"), col: C.accent1 },
    { I: fa.FaChild, en: L("Children", "어린이"), d: L("temperature homeostasis not fully developed; produce more heat and sweat less than adults", "체온 항상성이 아직 덜 발달했고, 어른보다 열은 더 내고 땀은 덜 흘린다"), col: C.accent1 },
    { I: md.MdAgriculture, en: L("Farm workers", "농업 노동자"), d: L("toiling under the hot sun — at considerable risk", "뜨거운 햇볕 아래에서 일해 위험이 크다"), col: C.accent3 },
  ];
  for (let i = 0; i < 3; i++) {
    const y = 2.45 + i * 1.47;
    card(s, 0.6, y, 6.3, 1.3, { fill: C.background2 });
    await iconCircle(s, 0.8, y + 0.27, 0.76, grp[i].I, { bg: grp[i].col, name: grp[i].en });
    text(s, [b(grp[i].en, { breakLine: true, fontSize: 20 }), t(grp[i].d, { fontSize: 16 })], 1.8, y + 0.06, 4.95, 1.18, { valign: "middle" });
  }
  card(s, 7.2, 2.45, 5.53, 2.8, { fill: C.text2, r: 0.16 });
  await ic(s, fa.FaCar, HEX.ice, 7.5, 2.62, 0.6, "car");
  text(s, [b(L("Closed car", "닫힌 차 안"), { color: C.background1 })], 8.25, 2.62, 4.3, 0.6, { fontSize: 20, valign: "middle" });
  text(s, "32°C", 7.5, 3.35, 1.85, 0.8, { fontSize: 36, bold: true, color: C.accent5, fontFace: "+mj-lt", align: "center", valign: "middle" });
  arrow(s, 9.4, 3.75, 10.15, 3.75, { col: C.accent3, w: 3 });
  text(s, "60°C", 10.2, 3.35, 1.9, 0.8, { fontSize: 36, bold: true, color: C.accent1, fontFace: "+mj-lt", align: "center", valign: "middle" });
  text(s, L("outside (90°F)  →  inside in < 1 hour (140°F)", "바깥 (90°F)  →  1시간 안에 차 안 (140°F)"), 7.45, 4.18, 5.05, 0.36, { fontSize: 16, color: C.background1, align: "center" });
  text(s, L("About 35–40 U.S. children die each year when left alone in cars.", "미국에서 해마다 35–40명의 어린이가 차 안에 혼자 남겨졌다가 숨진다."), 7.45, 4.56, 5.05, 0.62, { fontSize: 16, color: C.accent3, align: "center", bold: true });
  card(s, 7.2, 5.4, 5.53, 1.45, { fill: C.accent3, transparency: 80 });
  text(s, L([b("Heat waves: ", { color: C.text2 }), t("climate models predict more intense and frequent heat waves. The 2003 European heat wave caused 35,000–52,000 “excess deaths.”")], [b("폭염: ", { color: C.text2 }), t("기후 모델은 폭염이 더 강하고 잦아질 것으로 예측한다. 2003년 유럽 폭염 때는 3만 5천–5만 2천 명의 ‘초과 사망’이 발생했다.")]), 7.4, 5.45, 5.15, 1.35, { fontSize: 18, valign: "middle" });
  s.addNotes(NOTE("revisited"));
};

SL.keyconcepts = async () => {
  const s = contentSlide(null, null, null, L("Key Concepts of 32.1", "핵심개념의 요약"), "핵심개념의 요약", "CONTENT_DARK");
  const kc = [
    { I: fa.FaBalanceScale, col: C.background1, h: L("Homeostasis", "항상성"), d: L("a dynamic equilibrium: internal conditions (temperature, salt, oxygen, glucose, pH, water) are kept within a range in which proteins can function and energy can be made available.", "단백질이 기능하고 에너지를 쓸 수 있도록 체온, 염, 산소, 포도당, pH, 수분을 일정 범위로 유지하는 동적인 평형상태.") },
    { I: fa.FaThermometerHalf, col: C.accent3, h: L("Ectotherms vs. endotherms", "외온동물과 내온동물"), d: L("Ectotherms get most body warmth from the environment; endotherms get most heat from metabolic activities and regulate body temperature within a narrow range.", "외온동물은 몸을 데우는 열을 주로 환경에서 얻고, 내온동물은 대사 활동으로 열을 얻어 체온을 좁은 범위로 조절한다.") },
    { I: fa.FaSyncAlt, col: C.accent5, h: L("Negative feedback", "음성피드백"), d: L("a change triggers a response that counteracts the change and restores the original conditions — the main way homeostasis is maintained.", "변화가 그 변화를 거스르는 반응을 일으켜 원래 상태로 되돌린다 — 항상성을 유지하는 주된 방법.") },
    { I: tb.TbArrowBigUpLines, col: C.accent1, h: L("Positive feedback", "양성피드백"), d: L("a change initiates events that intensify the change — rare, and self-limiting, e.g., childbirth.", "변화가 그 변화를 더 키우는 일련의 사건을 일으킨다 — 드물고, 자기 제한적이다. 예: 출산.") },
  ];
  for (let i = 0; i < 4; i++) {
    const x = 0.6 + (i % 2) * 6.17, y = 1.55 + Math.floor(i / 2) * 2.75;
    card(s, x, y, 5.96, 2.55, { fill: C.background1, transparency: 92, r: 0.16 });
    await iconCircle(s, x + 0.3, y + 0.28, 0.8, kc[i].I, { bg: kc[i].col, fg: HEX.navy, name: kc[i].h });
    text(s, [b(kc[i].h, { color: kc[i].col, fontSize: 24 })], x + 1.3, y + 0.28, 4.45, 0.8, { valign: "middle" });
    text(s, kc[i].d, x + 0.3, y + 1.2, 5.4, 1.3, { fontSize: 18, color: C.background1 });
  }
  s.addNotes(NOTE("keyconcepts"));
};

SL.refs = async () => {
  const s = contentSlide(null, L("REFERENCES", "참고 자료"), T_REV, L("References & Image Credits", "참고 자료 및 이미지 출처"), "참고 자료 및 이미지 출처");
  text(s, bl(L([
    [b("Main textbook: "), t("Audesirk, Audesirk & Byers, "), t("Biology: Life on Earth", { italic: true }), t(", Ch. 32 “Homeostasis and the Organization of the Animal Body” — 32.1 (English ed. pp. 604–609, Chapter Review pp. 617–618; Korean ed. pp. 602–607, 615–616).")],
    [b("Figures: "), t("OpenStax, "), t("Anatomy and Physiology 2e", { italic: true }), t(", Figs. 1.10 & 1.11; OpenStax, "), t("Biology 2e", { italic: true }), t(", Fig. 33.23 — openstax.org, licensed CC BY 4.0. Arctic sea-ice maps (slide 22): NASA Earth Observatory images by Jesse Allen (public domain).")],
    [b("Photos (Wikimedia Commons): "), t("desert pupfish — P. V. Loiselle (CC BY-SA 3.0); cinnamon hummingbird — Charles J. Sharp (CC BY-SA 4.0); green iguana — Cayambe (CC BY-SA 3.0); polar bear — Andreas Weith (CC BY-SA 4.0); Claude Bernard — public domain; Walter B. Cannon — Bachrach 1934 (CC BY 4.0).")],
    [b("Diagrams: "), t("negative feedback loop (after textbook Fig. 32-2), heat-stroke vicious cycle, Arctic ice loop, and schematic graphs were drawn for this presentation from the textbook text. Icons: react-icons (Font Awesome, Game-icons, Material, Tabler).")],
  ], [
    [b("주 교재: "), t("Audesirk, Audesirk & Byers, "), t("Biology: Life on Earth", { italic: true }), t(", 32장 “Homeostasis and the Organization of the Animal Body” — 32.1 (영문판 pp. 604–609, 617–618; 한국어판 pp. 602–607, 615–616).")],
    [b("그림: "), t("OpenStax, "), t("Anatomy and Physiology 2e", { italic: true }), t(", 그림 1.10, 1.11; OpenStax, "), t("Biology 2e", { italic: true }), t(", 그림 33.23 — openstax.org, CC BY 4.0. 북극 해빙 지도(22번 슬라이드): NASA Earth Observatory, Jesse Allen (퍼블릭 도메인).")],
    [b("사진 (Wikimedia Commons): "), t("사막 pupfish — P. V. Loiselle (CC BY-SA 3.0); 벌새 — Charles J. Sharp (CC BY-SA 4.0); 이구아나 — Cayambe (CC BY-SA 3.0); 북극곰 — Andreas Weith (CC BY-SA 4.0); 클로드 베르나르 — 퍼블릭 도메인; 월터 캐넌 — Bachrach 1934 (CC BY 4.0).")],
    [b("도식: "), t("음성피드백 고리(교재 그림 32-2 재구성), 열사병 악순환, 북극 얼음 고리, 개념 그래프는 교재 본문을 바탕으로 이 발표를 위해 새로 그렸다. 아이콘: react-icons (Font Awesome, Game-icons, Material, Tabler).")],
  ]), { gap: 10, indent: 14 }), 0.6, 1.55, 12.13, 5.3, { fontSize: 16 });
  s.addNotes(NOTE("refs"));
};

// ================================================================ variants
// "§Name" starts a section; every other entry is a slide id from SL. Both decks use this order.
const ORDER = ["§Opening", "title", "case", "roadmap", "§1. What & Why", "constancy", "what", "dynamic", "enzymes", "case2",
    "§2. Body Temperature", "warmcold", "ectoendo", "ectocycle", "ectoreg", "endocost",
    "§3. Negative Feedback", "twotypes", "negfb", "fig322", "hypo", "water",
    "§4. Positive Feedback", "vicious", "posfb", "birth", "arctic1", "arctic2",
    "§Wrap-up", "revisited", "keyconcepts", "refs"];
const SECTION_KO = { "Opening": "도입", "1. What & Why": "1. 항상성이란", "2. Body Temperature": "2. 체온", "3. Negative Feedback": "3. 음성피드백", "4. Positive Feedback": "4. 양성피드백", "Wrap-up": "마무리" };

async function build() {
  for (const item of ORDER) {
    if (item.startsWith("§")) {
      CURRENT_SECTION = L(item.slice(1), SECTION_KO[item.slice(1)]);
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
const ENZ_FOLDED = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 420 300" width="420" height="300">
  <path d="M175 28 C255 22 318 62 322 112 L240 150 L322 188 C318 240 255 278 175 272 C92 268 48 215 48 150 C48 85 92 34 175 28 Z" fill="#${HEX.teal}"/>
  <path d="M92 95 C140 52 228 62 262 104 C222 126 140 112 104 146 C76 172 120 196 182 186 C236 178 252 210 232 240" fill="none" stroke="#A9DCE2" stroke-width="11" stroke-linecap="round"/>
  <g stroke="#ffffff" stroke-width="5" stroke-dasharray="8 7" stroke-linecap="round">
    <line x1="140" y1="72" x2="140" y2="112"/><line x1="198" y1="70" x2="198" y2="110"/>
    <line x1="120" y1="140" x2="128" y2="184"/><line x1="178" y1="128" x2="182" y2="178"/>
  </g>
  <polygon points="252,150 336,110 336,190" fill="#${HEX.amber}" stroke="#${HEX.amber}" stroke-width="12" stroke-linejoin="round"/>
</svg>`;
const ENZ_BROKEN =`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 420 300" width="420" height="300">
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
    // Korean deck: tag the slide text as Korean so PowerPoint does not spell-check it as English.
    if (KO) sx = sx.replace(/lang="en-US"/g, 'lang="ko-KR"');
    zip.file(name, sx);
  }
  console.log("paragraphs repaired:", fixedParas);
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT);

  // Export the read-aloud script (speaker notes) as Markdown and Word files next to the English deck.
  // The Korean deck has the same slides in the same order, so it shares that script.
  if (!KO) {
    const syllables = (str) => (str.match(/[가-힣]/g) || []).length + Math.ceil((str.match(/[A-Za-z]/g) || []).length / 3) + (str.match(/[0-9]/g) || []).length;
    const total = OUTLINE.reduce((n, o) => n + syllables(NOTE(o.id)), 0);
    const minutes = Math.round(total / 300); // ~300 syllables per minute of presenting
    const md = [
      "# 32.1 항상성 — 발표 대본",
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
    new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("32.1 항상성 — 발표 대본")] }),
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
    title: "32.1 항상성 — 발표 대본",
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
