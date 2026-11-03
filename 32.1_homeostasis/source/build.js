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

const OUT = path.resolve(process.argv[2] || path.join(__dirname, "..", "32.1_Homeostasis.pptx"));

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

function contentSlide(section, tagLabel, tagOpts, title, kr, master = "CONTENT") {
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
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
async function build() {
  // ---------- 1. Title
  pres.addSection({ title: "Opening" });
  {
    const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Opening" });
    text(s, "GENERAL BIOLOGY  ·  CHAPTER 32", 0.75, 1.55, 7, 0.35, { fontSize: 14, bold: true, color: C.accent3, charSpacing: 2 });
    text(s, "32.1", 0.75, 1.0, 3, 0.6, { fontSize: 30, bold: true, color: C.accent1 });
    s.addText("Homeostasis: Why and How Do Animals Regulate Their Internal Environment?", { placeholder: "title" });
    s.addText("항상성: 동물은 왜 그리고 어떻게 자신의 내부 환경을 조절하는가?", { placeholder: "kr" });
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
    s.addNotes(`[발표 시작 · 약 1분]
안녕하세요. 저희 조는 32장 1절 "항상성: 동물은 왜 그리고 어떻게 자신의 내부 환경을 조절하는가?"를 발표하겠습니다.
오른쪽 그림처럼 차가운 쪽(파란색)과 뜨거운 쪽(빨간색) 사이에서 '설정점(set point)'을 유지하는 것이 오늘의 핵심 주제입니다.
발표는 열사병 사례로 시작해서, (1) 항상성의 정의와 이유, (2) 외온동물과 내온동물, (3) 음성피드백, (4) 양성피드백 순서로 진행하고 마지막에 복습 문제를 함께 풀어 보겠습니다.`);
  }

  // ---------- 2. Case study
  {
    const s = contentSlide("Opening", "CASE STUDY", T_CASE, "Case Study: Overheated", "과도한 열 — 열사병(heat stroke)으로 쓰러진 대학 운동선수들");
    card(s, 0.6, 1.95, 4.75, 4.7, { fill: C.accent3, transparency: 82, r: 0.16 });
    await ic(s, fa.FaQuoteLeft, HEX.amber, 0.9, 2.2, 0.55, "quote");
    text(s, "“I’m having the race of my life. … Nothing is going to get in my way, not the 90° heat and 90% humidity, … not the fire in my feet that is spreading up my legs.”", 0.9, 2.9, 4.2, 2.3, { fontSize: 17, italic: true, color: C.text1 });
    text(s, [b("Kierann Smith", { breakLine: true }), t("medical student, Stanford triathlon team — just before she collapsed (2011 USA Triathlon Collegiate Nationals, Alabama)", { fontSize: 13, color: C.accent4 })], 0.9, 5.3, 4.2, 1.2, { fontSize: 15 });

    const rows = [
      { big: "41.1°C", f: "106°F", who: "Kierann Smith · triathlete", what: "Core body temperature at collapse → ice water + ice packs brought it just below 40°C (104°F). She survived.", col: C.accent3 },
      { big: "41.6°C", f: "107°F", who: "Ben Richards · wrestler, 20", what: "Collapsed during a 5-mile training run (2013); died 10 days later.", col: C.accent1 },
      { big: "2014", f: "cause: heat stroke", who: "Marquese Meadow · football, freshman", what: "Disoriented after practice; died in a nearby hospital. Autopsy confirmed heat stroke as the cause of death.", col: C.accent1 },
    ];
    let y = 1.95;
    for (const r of rows) {
      card(s, 5.65, y, 7.08, 1.42, { fill: C.background2, r: 0.14 });
      text(s, r.big, 5.85, y + 0.16, 1.9, 0.72, { fontSize: 34, bold: true, color: r.col, fontFace: "+mj-lt", align: "center", valign: "middle" });
      text(s, r.f, 5.85, y + 0.88, 1.9, 0.35, { fontSize: 14, color: C.accent4, align: "center" });
      text(s, [b(r.who, { breakLine: true, color: C.text2 }), t(r.what, { fontSize: 14 })], 7.95, y + 0.17, 4.6, 1.15, { fontSize: 16, valign: "middle" });
      y += 1.64;
    }
    s.addNotes(`[Case Study · 약 1분 30초]
교재 32장은 '과도한 열(Overheated)'이라는 사례로 시작합니다.
스탠퍼드 철인 3종 선수이자 의대생인 Kierann Smith는 2011년 대학생 철인 3종 경기에서 "32℃의 온도와 90%의 습도도 나를 막을 수 없다"고 생각한 직후 쓰러졌고, 체온이 41.1℃까지 올랐습니다. 얼음물과 얼음팩으로 체온을 40℃ 아래로 떨어뜨렸고, 그녀는 살아남았습니다.
하지만 레슬링 선수 Ben Richards는 2013년 5마일 달리기 훈련 중 쓰러져 병원 도착 당시 체온이 41.6℃였고 10일 후 사망했습니다. 2014년 축구선수 Marquese Meadow도 연습 후 방향 감각을 잃고 사망했는데, 검사 결과 열사병으로 확인되었습니다.
(교재 원문은 화씨: 106°F, 107°F, 104°F)`);
  }

  // ---------- 3. Roadmap
  {
    const s = contentSlide("Opening", "OVERVIEW", T_REV, "Today’s Roadmap", "오늘 다룰 내용 — 교재의 Check Your Learning 4가지 질문");
    const items = [
      { n: "1", en: "What is homeostasis — and why do organisms maintain it?", kr: "항상성의 정의와 유지 이유", I: fa.FaBalanceScale, col: C.text2 },
      { n: "2", en: "How do ectotherms and endotherms differ?", kr: "외온동물 vs. 내온동물", I: fa.FaThermometerHalf, col: C.accent3 },
      { n: "3", en: "Why is negative feedback crucial for homeostasis?", kr: "음성피드백이 중요한 이유", I: fa.FaSyncAlt, col: C.accent2 },
      { n: "4", en: "What is positive feedback — and why is it rare?", kr: "양성피드백과 그것이 드문 이유", I: tb.TbArrowBigUpLines, col: C.accent1 },
    ];
    const w = 2.86, gap = 0.23;
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + i * (w + gap), it = items[i];
      card(s, x, 2.0, w, 3.85, { fill: C.background2, r: 0.16 });
      await iconCircle(s, x + 0.3, 2.3, 0.9, it.I, { bg: it.col, name: "part " + it.n });
      text(s, "PART " + it.n, x + 1.35, 2.48, 1.4, 0.5, { fontSize: 16, bold: true, color: it.col, valign: "middle" });
      text(s, it.en, x + 0.3, 3.45, w - 0.55, 1.5, { fontSize: 17, bold: true, color: C.text1 });
      text(s, it.kr, x + 0.3, 5.0, w - 0.55, 0.7, { fontSize: 15, color: C.accent4 });
    }
    card(s, 0.6, 6.12, 12.13, 0.62, { fill: C.accent3, transparency: 80, r: 0.31 });
    await ic(s, fa.FaRunning, HEX.ink, 0.88, 6.2, 0.46, "athlete");
    text(s, [b("Running thread: "), t("the heat-stroke case (열사병 사례) — what keeps body temperature constant, and what goes wrong.")], 1.5, 6.12, 11.0, 0.62, { fontSize: 15, valign: "middle" });
    s.addNotes(`[목차 · 약 1분]
오늘 발표는 교재 32.1절 끝의 "Check Your Learning" 네 가지 질문에 답할 수 있도록 구성했습니다.
1) 항상성을 정의하고 유기체가 항상성을 유지하는 이유, 2) 냉혈(외온)동물과 온혈(내온)동물의 차이와 예, 3) 항상성 유지에 음성피드백이 중요한 이유, 4) 양성피드백의 정의와 그것이 생명체에서 드문 이유입니다.
그리고 열사병 사례가 발표 전체에 걸쳐 계속 등장합니다.`);
  }

  // ---------- 4. Internal constancy
  pres.addSection({ title: "1. What & Why" });
  {
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
    s.addNotes(`[약 1분 30초]
여러분이 방에 앉아 있든, 정글에서 더위에 고생하든, 눈보라 속에서 떨고 있든 심장·뇌·근육 같은 신체 대부분의 세포는 거의 일정한 온도를 유지합니다.
또 사막에 있든 수영장, 바다, Great Salt Lake에서 수영을 하든 세포들은 거의 일정한 조성을 가진 '세포간질액(interstitial fluid)'에 잠겨 있습니다.
이런 동물체의 '내부 불변성'은 19세기 중반 프랑스 생리학자 Claude Bernard가 처음 기술했고, 1920년대에 Walter Cannon이 '항상성(homeostasis)'이라는 용어를 만들었습니다.`);
  }

  // ---------- 5. What is homeostasis
  {
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
    s.addNotes(`[약 1분 30초]
Cannon은 항상성을 "유기체가 세포의 적절한 기능을 위해 필요한 조건의 범위 속에서 내부 환경을 유지하려는 능력"이라고 설명했습니다.
항상성 메커니즘은 체온과 체액의 물·염 농도를 유지하는 것 외에도 포도당 농도, pH(산-염기의 균형), 호르몬 분비, 산소와 이산화탄소의 농도 등 체내의 다른 환경을 조절합니다.`);
  }

  // ---------- 6. Dynamic, not static
  {
    const s = contentSlide("1. What & Why", "PART 1 · WHAT & WHY", T_PART, "Homeostasis Is Dynamic, Not Static", "‘동일하게 유지하려는 것’이지만, 내부 환경은 끊임없이 조절된다");
    // left concept
    card(s, 0.6, 1.95, 3.9, 4.75, { fill: C.background2, r: 0.16 });
    text(s, [t("“homeostasis”", { bold: true, color: C.accent2, fontSize: 22, breakLine: true }), t("= “to stay the same”", { fontSize: 17, breakLine: true }), t("(동일하게 유지하려는 것)", { fontSize: 14, color: C.accent4 })], 0.85, 2.15, 3.45, 1.35, {});
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
    s.addNotes(`[약 1분 30초]
'항상성'이라는 단어는 "동일하게 유지하려는 것"을 의미해서 정적인 불변의 상태처럼 들리지만, 실제 내부 환경은 변화하는 내부·외부 조건에 맞서 일관성을 유지하기 위해 끊임없이 조절되고 있습니다. 교재 요약에서는 이를 '동적인 평형상태(dynamic equilibrium)'라고 표현합니다.
예를 들어 운동은 항상성 메커니즘에 도전합니다. 일하는 근육의 에너지 요구를 유지하려면 더 많은 산소가 공급되어야 하고, 근육이 만든 여분의 열은 사라져야 합니다.
그래서 몸은 호흡수, 땀샘 활동, 피부로의 혈류를 증가시킵니다. 이 단원 전체에서 몸 전체의 상호작용 시스템들이 항상성 유지를 돕는 예를 계속 보게 됩니다.
※ 한국어판은 이 문장을 "호흡수, 땀샘 활동, 그리고 피부로의 혈류 증가에 따라 몸의 온도는 상승한다"로 옮겼지만, 영어 원문 "The body rises to the occasion"은 '몸이 (그 요구에) 대응한다'는 뜻입니다. 땀샘 활동과 피부 혈류 증가는 오히려 열을 내보내는 반응이므로 원문 의미로 설명하세요.`);
  }

  // ---------- 7. Enzymes
  {
    const s = contentSlide("1. What & Why", "PART 1 · WHAT & WHY", T_PART, "Homeostasis Allows Enzymes to Function", "항상성은 효소가 기능하도록 한다 — 세포가 주변 환경에 민감한 이유");
    // Left: structure
    card(s, 0.6, 1.95, 6.35, 4.0, { fill: C.background2, r: 0.16 });
    text(s, [b("① Structure ", { color: C.accent2 }), t("(구조)", { color: C.accent4 })], 0.85, 2.1, 5.9, 0.42, { fontSize: 19 });
    text(s, [t("Almost every reaction in a cell is catalyzed by a specific "), b("enzyme"), t(" (효소). Its function depends on its precise "), b("three-dimensional structure"), t(", maintained in part by "), b("hydrogen bonds", { color: C.accent2 }), t(" (수소결합).")], 0.85, 2.55, 5.9, 1.25, { fontSize: 15 });
    const folded = await svgPng(ENZ_FOLDED, 700), broken = await svgPng(ENZ_BROKEN, 700);
    s.addImage({ data: folded, x: 0.95, y: 3.75, w: 2.45, h: 1.75, altText: "Folded enzyme with substrate in its active site" });
    s.addImage({ data: broken, x: 4.15, y: 3.75, w: 2.45, h: 1.75, altText: "Unfolded enzyme with disrupted hydrogen bonds" });
    arrow(s, 3.5, 4.62, 4.05, 4.62, { col: C.accent1, w: 2.5 });
    text(s, "works", 0.95, 5.5, 2.45, 0.3, { fontSize: 13, bold: true, color: C.accent2, align: "center" });
    text(s, "bonds disrupted", 4.15, 5.5, 2.45, 0.3, { fontSize: 13, bold: true, color: C.accent1, align: "center" });
    // stressor chips (bottom row, left)
    const st = [["too salty", "극염"], ["too acidic", "극산성"], ["too basic", "극염기성"], ["too hot", "너무 뜨거움"]];
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
    text(s, [t("→ need a narrow range of ", { color: C.background1 }), b("salt, pH, temperature", { color: C.accent3 })], ccx + 0.1, 6.2, 12.73 - ccx - 0.2, 0.42, { fontSize: 13, valign: "middle", align: "center" });
    s.addNotes(`[약 2분]
왜 세포는 주변 환경에 이렇게 민감할까요? 생명체가 항상성을 유지하는 큰 이유는 단백질, 특히 효소 때문입니다.
① 구조: 세포 내 거의 모든 생화학 반응은 특정 효소에 의해 촉매되고, 효소의 활성은 정확한 3차 구조에 의존합니다. 이 구조는 부분적으로 수소결합으로 유지되는데(2장, 3장), 수소결합은 중요하지만 취약해서 너무 짜거나(극염), 너무 산성이거나, 너무 염기성이거나, 너무 뜨거운 환경에서 파괴될 수 있습니다. 그래서 염 농도, pH, 온도가 좁은 범위 안에 있어야 합니다.
② 속도: 온도는 또 다른 이유로도 중요합니다. 효소가 반응을 촉매하는 속도는 온도에 따라 달라집니다. 저온은 분자 운동을 느리게 해서 효소와 그 대상 분자 간 상호작용의 수와 속도를 감소시키고, 온도가 올라가면 반응이 빨라지며 때로는 해로운 수준에 이르기도 합니다.
(왼쪽 효소 그림은 이해를 돕기 위한 개념도입니다.)`);
  }

  // ---------- 8. Case continued: organ damage
  {
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
    text(s, [b("Kidney failure", { breakLine: true }), t("신부전 (kidneys: 신장)", { color: C.accent4, fontSize: 14 })], 10.15, 2.62, 2.5, 1.3, { fontSize: 16, valign: "middle" });
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
    s.addNotes(`[Case Study Continued · 약 1분 30초]
축구선수 Marquese Meadow와 레슬링선수 Ben Richards는 사망 전 일주일 넘게 병원에서 치료를 받았고, 체온은 열사병 후 몇 시간 내에 정상으로 내려갔습니다. 그런데 왜 회복하지 못했을까요?
열사병에 따른 신진대사의 이상(metabolic chaos)이 생명에 중요한 기관, 특히 간과 신장을 손상시키기 때문입니다. 근육세포들이 파괴되면서 물질들이 혈액으로 방출되어 신부전을 일으킬 수 있고, 상승한 체온은 혈액의 독소물질을 제거하는 데 중요한 간도 손상시킵니다.
Kierann Smith도 "심각한 근육 손상과 일부 간 손상을 확인했을 때 큰 타격을 받았다"고 했습니다.
인체의 대사과정은 좁은 온도 범위 내에서만 기능합니다. 그렇다면 모든 동물에게도 해당될까요? → 다음 파트로 넘어갑니다.`);
  }

  // ---------- 9. Warm or cold blooded
  pres.addSection({ title: "2. Body Temperature" });
  {
    const s = contentSlide("2. Body Temperature", "PART 2 · BODY TEMPERATURE", T_PART, "“Warm-Blooded” or “Cold-Blooded”?", "온혈 혹은 냉혈? — 이 용어들은 종종 오해의 소지가 있다");
    text(s, [t("Mammals and birds are often called "), b("“warm-blooded”", { color: C.accent1 }), t(" (온혈); other reptiles, amphibians, fish, and invertebrates "), b("“cold-blooded”", { color: C.accent2 }), t(" (냉혈). But these terms are often "), b("misleading"), t(":")], 0.6, 1.92, 12.13, 0.75, { fontSize: 16 });
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
    s.addNotes(`[약 1분 30초]
동물들 사이에서 가장 많이 연구된 항상성 메커니즘은 아마도 체온 조절일 것입니다.
흔히 포유류와 조류를 '온혈(warm-blooded)', 파충류·양서류·어류·무척추동물을 '냉혈(cold-blooded)'이라고 하지만 이 용어들은 종종 오해의 소지가 있습니다.
예를 들어 한여름 태양이 사막 pupfish가 사는 샘과 작은 연못을 뜨겁게 만들면 pupfish의 체온은 37.8℃를 넘어서, 때로는 완전한 '온혈'동물이 됩니다(그림 32-1a).
반대로 벌새는 먹이를 찾아다니는 낮 동안에는 체온이 40.5℃까지 올라가지만 밤에는 13℃까지 내려갑니다(그림 32-1b).
(사진의 벌새는 교재의 루비목벌새와 다른 종인 cinnamon hummingbird입니다.)`);
  }

  // ---------- 10. Ectotherms vs Endotherms
  {
    const s = contentSlide("2. Body Temperature", "PART 2 · BODY TEMPERATURE", T_PART, "Ectotherms vs. Endotherms", "외온동물 vs. 내온동물 — 체온 유지를 위한 주요 열원(熱源)으로 분류");
    text(s, "To avoid confusion, scientists classify animals by their major source of body warmth.", 0.6, 1.92, 12.13, 0.4, { fontSize: 16, color: C.text1 });
    const cols = [
      { x: 0.6, col: C.accent2, colHex: HEX.teal, title: "ECTOTHERMS", kr: "외온동물", greek: "Greek: “outside heat” (외부열)", I: fa.FaSun, src: "derive most of their heat from the environment", ex: [["Reptiles (except birds)", gi.GiGecko], ["Amphibians", gi.GiFrog], ["Most fish", fa.FaFish], ["Most invertebrates", gi.GiButterfly]], note: "Simplest cases — an earthworm in its burrow or a fish in a stream: body temperature = environment temperature." },
      { x: 6.77, col: C.accent1, colHex: HEX.coral, title: "ENDOTHERMS", kr: "내온동물", greek: "Greek: “inside heat” (내부열)", I: fa.FaFire, src: "produce most of their heat by metabolic reactions (대사반응)", ex: [["Birds", gi.GiHummingbird], ["Mammals", fa.FaPaw]], note: "Birds and mammals are the most common endotherms; some other animals occasionally generate significant heat metabolically." },
    ];
    for (const c of cols) {
      const w = 5.96;
      card(s, c.x, 2.5, w, 4.25, { fill: c.col, transparency: 88, r: 0.16 });
      await iconCircle(s, c.x + 0.3, 2.72, 1.0, c.I, { bg: c.col, name: c.title });
      text(s, [b(c.title, { color: c.col, fontSize: 24, breakLine: true }), t(c.kr + "  ·  " + c.greek, { fontSize: 14, color: C.accent4 })], c.x + 1.5, 2.72, 4.3, 1.0, { valign: "middle" });
      text(s, [b("Heat source: "), t(c.src)], c.x + 0.3, 3.9, w - 0.6, 0.75, { fontSize: 16 });
      let ex = c.x + 0.55;
      for (const [name, I] of c.ex) {
        await iconCircle(s, ex, 4.75, 0.55, I, { bg: C.background1, fg: c.colHex, scale: 0.7, name });
        text(s, name, ex + 0.275 - 0.66, 5.33, 1.32, 0.55, { fontSize: 12, align: "center", color: C.text1 });
        ex += 1.38;
      }
      text(s, c.note, c.x + 0.3, 5.95, w - 0.6, 0.75, { fontSize: 14, italic: true, color: C.text1 });
    }
    s.addNotes(`[약 1분 30초]
이러한 혼란을 피하기 위해 과학자들은 체온 유지를 위한 주요 열원을 토대로 동물들을 분류합니다.
외온동물(ectotherms, 그리스어로 '외부열')은 환경으로부터 대부분의 열을 얻습니다. 파충류(조류는 제외)와 양서류, 그리고 대부분의 어류와 무척추동물이 외온동물입니다. 가장 단순한 예로 굴 속의 지렁이나 개울 속의 물고기는 체온이 주변 환경의 온도와 같아집니다.
내온동물(endotherms, '내부열')은 대사반응을 통하여 대부분의 열을 생산합니다. 조류와 포유류가 가장 대표적인 내온동물이지만, 몇몇 다른 동물들도 때때로 대사작용을 통해 상당한 열을 생산하기도 합니다.`);
  }

  // ---------- 11. Ectotherms warm up / cool down
  {
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
    s.addNotes(`[약 1분 30초]
대부분의 외온동물의 체온은 시간이나 일간, 주간에 따른 외부 기온 변화에 따라 상당히 다양합니다.
나비, 벌, 그리고 대부분의 도마뱀과 같은 많은 외온동물들은 밤에는 체온을 낮추고 비활성 상태가 되어 에너지를 보존하지만, 낮 동안에는 체온을 높이며 활동합니다.
밤에 나비와 벌은 체온이 너무 낮아져 날지 못하고, 도마뱀도 종종 너무 느려 먹이를 사냥하지 못하거나 포식자로부터 도망치지 못합니다.
새벽에 벌은 스스로 몸을 떨고 나비는 날갯짓을 하면서 대사열을 만들어 내는 반면, 도마뱀은 따뜻하게 햇볕에 데워진 바위를 찾습니다.`);
  }

  // ---------- 12. Ectotherms can regulate
  {
    const s = contentSlide("2. Body Temperature", "PART 2 · BODY TEMPERATURE", T_PART, "Ectotherms Can Still Regulate Temperature", "행동(behavior) 혹은 매우 일정한 환경을 통해 체온을 안정적으로 유지");
    const img = await photo(PHOTO.iguana.f, 4.6, 3.95, { pos: "centre", px: 1000 });
    s.addImage({ data: img, x: 0.6, y: 1.95, w: 4.6, h: 3.95, altText: "Green iguana basking on a rock" });
    text(s, [b("Basking in the sun", { color: C.text2 }), t(" — a behavioral mechanism (행동메커니즘), cf. Fig. 32-1c")], 0.6, 5.98, 4.6, 0.5, { fontSize: 13 });
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
    s.addNotes(`[약 1분 30초]
가변 체온이 외온동물에서 일반적이지만, 행동을 통해서 혹은 매우 일정하게 유지되는 환경 내에서 몇몇 냉혈동물은 꽤 안정적으로 체온을 유지할 수 있습니다.
예를 들어 곤충과 도마뱀은 종종 햇볕을 쬐며 몸을 데우고(그림 32-1c의 이구아나), 그늘로 이동하여 몸을 식힙니다.
앞서 언급한 pupfish는 2.2~45℃ 범위의 수온에서 견딜 수 있지만 좁은 온도 범위 내에서만 번식할 수 있습니다. 그래서 번식기에 pupfish는 수온 변화에 따라 수영장(연못)의 서로 다른 지역을 수영함으로써 매우 정확하게 체온을 조절합니다.
심해의 경우 온도가 상당히 일정(0~3℃)하므로 대부분의 냉혈 심해어류는 일생 동안 체온 변화를 거의 겪지 않습니다.
(막대의 빨간 '번식 범위'는 위치를 나타낸 개념도이며, 교재에 구체적 수치는 없습니다.)`);
  }

  // ---------- 13. Endotherms benefits/costs
  {
    const s = contentSlide("2. Body Temperature", "PART 2 · BODY TEMPERATURE", T_PART, "Endotherms: Benefits and Costs", "내온동물 — 따뜻한 체온을 유지하는 데에는 이점과 비용 두 가지가 있다");
    card(s, 0.6, 1.95, 4.1, 4.8, { fill: C.text2, r: 0.16 });
    text(s, "Most endotherms keep a fairly constant body temperature of about", 0.9, 2.2, 3.55, 1.0, { fontSize: 16, color: C.background1 });
    text(s, "35–41°C", 0.9, 3.15, 3.55, 1.0, { fontSize: 48, bold: true, color: C.accent3, fontFace: "+mj-lt" });
    text(s, "(95–106°F)", 0.9, 4.1, 3.55, 0.4, { fontSize: 15, color: C.accent5 });
    text(s, [t("A few, such as ", { color: C.background1 }), b("hummingbirds", { color: C.accent5 }), t(", allow their body temperature to fall at night.", { color: C.background1 })], 0.9, 4.75, 3.55, 1.2, { fontSize: 15 });
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
    s.addNotes(`[약 1분 30초]
벌새와 같은 몇몇 내온동물들은 야간에 체온이 내려가기는 하지만, 대부분은 약 35℃부터 41℃ 사이의 상당히 일정한 체온을 유지합니다.
체온을 따뜻하게 유지하는 데에는 이점과 비용 두 가지가 있습니다. 주요한 이점으로는 따뜻한 몸이 일반적으로 환경을 더 잘 감지하고, 더 빨리 반응하며, 추운 몸보다 더 빨리 움직일 수 있다는 것입니다.
물론 주요한 비용은 높은 체온을 유지하기 위해 요구되는 에너지입니다.
그러면 동물들은 어떻게 체내의 상태를 인지하고 필요에 따라 조절할까요? 그들의 내부 환경은 '피드백 시스템'이라는 메커니즘에 의해 유지됩니다. → 다음 파트`);
  }

  // ---------- 14. Two types of feedback
  pres.addSection({ title: "3. Negative Feedback" });
  {
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
    }
    text(s, "Graphs are schematic (개념도).", 0.6, 6.82, 4, 0.2, { fontSize: 10, italic: true, color: C.accent4 });
    s.addNotes(`[약 1분 30초]
피드백 시스템에는 두 가지 형태가 존재합니다.
하나는 음성피드백 시스템으로, 내부 환경에 있어 변화의 효과에 반대로 작용하는 것이고 이는 주로 항상성을 유지하는 데 이용됩니다. 왼쪽 그래프처럼 값이 설정점에서 벗어나면 반대 방향의 반응이 일어나 다시 설정점으로 돌아옵니다.
다른 하나는 양성피드백 시스템으로, 변화가 스스로를 증폭시키는 사이클을 만듭니다. 오른쪽처럼 변화가 점점 커집니다.
(그래프는 개념을 보여주기 위한 개념도입니다.)`);
  }

  // ---------- 15. Negative feedback reverses change + A&P Fig 1.10
  {
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
    s.addNotes(`[약 2분]
항상성을 조절하는 가장 중요한 메커니즘은 음성피드백이며, 변화에 대한 반응을 역방향으로 전환하는 것입니다. 음성피드백의 전체적인 결과는 그 반응계를 역전시켜 초기의 조건 상태로 되돌아오게 하는 것입니다.
음성피드백은 생물계와 무생물계 모두가 가지는 공통적 특성입니다. 체온, 혈액 내의 혈당량과 호르몬·물·염분·산소의 수준, 똑바로 서 있기와 같은 거의 모든 생리적 측면을 조절합니다. 집안의 온도 유지, 변기 탱크 채우기, 크루즈 컨트롤에도 음성피드백이 적용됩니다.
모든 음성피드백 시스템은 3가지 주요 구성요소, 즉 감지기(sensor), 조절 센터(control center), 효과기(effector)를 포함합니다. 감지기는 현재의 조건을 확인하고, 조절 센터는 '설정점(set point)'이라는 원하는 상태와 현재 상태를 비교하며, 효과기는 원하는 상태로 되돌리는 출력을 만듭니다.
오른쪽 그림(OpenStax)은 (a) 일반적인 고리와 (b) 체온이 37℃를 넘었을 때 땀샘이 작동해 열 손실을 늘리는 예입니다.`);
  }

  // ---------- 16. Fig 32-2 diagram (house vs body)
  {
    const s = contentSlide("3. Negative Feedback", "PART 3 · NEGATIVE FEEDBACK", T_PART, "Negative Feedback in Action (Fig. 32-2)", "음성피드백은 항상성을 유지한다 — 집의 온도와 체온 조절");
    const Lx = 1.55, Lw = 3.05, Cx = 5.0, Cw = 3.35, Rx = 8.75, Rw = 3.98;
    const mid = Cx + Cw / 2;
    // column headers
    await ic(s, fa.FaHome, HEX.slate, Lx, 1.86, 0.36, "house");
    text(s, [b("House", { color: C.text2 }), t("  집", { color: C.accent4 })], Lx + 0.45, 1.86, 2.4, 0.36, { fontSize: 15, valign: "middle" });
    await ic(s, fa.FaChild, HEX.slate, Rx, 1.86, 0.36, "body");
    text(s, [b("Body", { color: C.text2 }), t("  몸", { color: C.accent4 })], Rx + 0.45, 1.86, 2.4, 0.36, { fontSize: 15, valign: "middle" });
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
      { y1: 5.02, y2: 5.52, t1: "if sensor and set point differ, activates", t2: "감지기와 설정점이 다르면 작동", two: true },
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
    s.addNotes(`[핵심 그림 · 약 2분 30초]
교재 그림 32-2를 재구성한 도식입니다. 가운데가 음성피드백의 단계이고, 왼쪽은 집, 오른쪽은 우리 몸의 예입니다.
매우 추운 날 집의 온도를 조절하는 음성피드백 시스템에서 감지기는 온도측정기(온도계), 조절 센터는 온도조절기, 효과기는 열발생기(난방 장치)입니다. 온도계는 방의 온도를 감지하고 그 정보를 온도조절기로 보내며, 여기서 실제 온도는 원하는 설정 온도와 비교됩니다. 방의 온도가 설정 온도보다 낮으면 온도조절기는 열발생기로 신호를 보내 열을 발생시키고, 방이 따뜻해져 설정 온도로 돌아오면 열발생기의 작동을 멈추게 합니다.
몸에서는 피부의 온도 수용체가 감지기, 시상하부가 설정점을 가진 조절 센터, 근육이 효과기로서 떨림을 통해 열을 냅니다. 그리고 바뀐 조건이 다시 감지되는 '고리(loop)'가 만들어집니다.
[교재 비판적 사고 질문] 추위에 떨고 있는 포유류가 열을 감지하는 신경말단을 파괴하는 독을 섭취했다면 어떤 현상이 일어날까? → 마지막 토론 슬라이드에서 다시 다룹니다.`);
  }

  // ---------- 17. Hypothalamus + Bio2e 33.23
  {
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
    s.addNotes(`[약 2분]
척추동물에서 중요한 항상성 조절자는 뇌의 깊숙한 부위인 시상하부(hypothalamus)에 위치하고 있습니다. 조류와 포유동물에서 시상하부는 신경신호, 호르몬 방출 및 환경에 따른 온도의 큰 변동에도 불구하고 체온을 유지하려는 행동을 조정합니다. 피부와 신체 다른 부위의 신경말단은 온도감지기 역할을 하며 이 정보를 시상하부로 전달합니다.
만약 체온이 설정점(정상 체온) 아래로 떨어지면, 시상하부는 몸을 떨거나(작고 빠른 근 수축) 따뜻한 곳으로 이동함으로써 열을 생성하거나 보존하려는 메커니즘을 활성화합니다. 신체의 중요하지 않은 부위(얼굴, 손, 발, 피부)의 혈관은 수축되어 열손실을 감소시키고 신체 중심부(뇌, 심장, 내부 기관)로 따뜻한 혈액을 보냅니다. 정상 체온으로 돌아오면 온도감지기는 시상하부로 신호를 보내 이 활동들을 멈추게 합니다.
체온이 설정점보다 올라가면 시상하부는 더 많은 혈액을 피부로 보내 열이 주변 공기로 방출되게 합니다. 땀샘은 물이 피부에서 증발함에 따라 몸을 시원하게 하는 분비물을 만들고, 피로와 불편함은 사람들이 휴식을 취하고 그늘이나 냉수를 찾도록 합니다.
가운데 그림은 OpenStax 교재의 그림으로 같은 과정을 보여 줍니다.`);
  }

  // ---------- 18. Have you ever wondered: water
  {
    const s = contentSlide("3. Negative Feedback", "HAVE YOU EVER WONDERED …", T_WOND, "Can You Drink Too Much Water?", "과량의 물을 마실 수 있는가? — 간단히 답하자면 “그렇다”");
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
    s.addNotes(`[Have You Ever Wondered · 약 1분 30초]
"과량의 물을 마실 수 있을까?" 간단히 답하자면 '그렇다'입니다. 대부분의 것들이 그렇듯 물이라 할지라도 과하면 해로울 수 있습니다.
수화(Hydration)는 사람들이 필요 이상으로 많이 마시게 부추기는 유행어가 되었고, 열사병을 걱정하는 일부 코치들은 과한 수분 섭취를 요구하기도 합니다.
조지아주 17살 고등학교 축구선수 Zyrees Oliver는 연습 중 근육 경련을 풀려고 물 2갤런과 게토레이 2갤런을 마셨고, 집에 돌아온 후 쓰러져 사망했습니다. 원인은 '너무 적은 나트륨'을 뜻하는 저나트륨혈증(hyponatremia)으로, 과량의 수분 섭취로 혈액 내 나트륨이온(Na+)이 희석된 것입니다.
보스턴 마라톤 결승선에서 테스트한 사람들 중 13%는 가벼운 저나트륨혈증을 보였고 일부는 위험할 정도로 낮았습니다.
나트륨이온은 신경 신호전달과 근육 수축에 필수적이기 때문에, 체액의 나트륨이온을 매우 정교한 범위 내에서 조절하는 항상성 메커니즘이 진화했습니다.`);
  }

  // ---------- 19. Case: when feedback fails (vicious cycle)
  pres.addSection({ title: "4. Positive Feedback" });
  {
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
    s.addNotes(`[Case Study Continued · 약 1분 30초]
열사병에 걸리면 체온을 상대적으로 일정하게 유지시키는 음성피드백 메커니즘이 제대로 작동하지 않습니다.
과열로 인한 피드백의 한 가지 형태는 피로감이며, 피로감은 일반적으로 사람을 쉬게 합니다. 그러나 열사병의 증상인 정신적 혼란은 이 음성피드백을 파괴하고 운동선수가 신체의 경고 신호를 쉽게 무시하도록 합니다.
또 열사병 환자는 급격한 체온 상승으로 인한 뇌의 온도 조절 시스템 오작동으로 땀 흘리는 것이 멈출 수 있습니다.
효소 활성도는 온도에 따라 증가하므로 대사율이 상승하여 더 많은 열을 발생시키고, 체온은 더 치솟습니다. 즉각적인 냉각으로 이를 막지 않으면 이 악순환—양성피드백의 일례—은 사망으로 이어질 수 있습니다.
그렇다면 양성피드백이란 무엇일까요? 항상 해로울까요? 얼마나 자주 자연적으로 일어날까요?`);
  }

  // ---------- 20. Positive feedback amplifies change
  {
    const s = contentSlide("4. Positive Feedback", "PART 4 · POSITIVE FEEDBACK", T_PART, "Positive Feedback Amplifies Change", "양성피드백은 변화의 효과를 촉진시킨다");
    card(s, 0.6, 1.95, 12.13, 0.9, { fill: C.accent1, transparency: 86, r: 0.16 });
    text(s, [b("Positive feedback", { color: C.accent1 }), t(" (양성피드백): a change produces a response that "), b("amplifies that change"), t(".")], 0.9, 1.95, 11.6, 0.9, { fontSize: 19, valign: "middle" });
    // haywire thermostat staircase
    text(s, [b("Imagine a thermostat gone haywire", { color: C.text2, breakLine: true }), t("고장 난 온도조절기를 상상해 보라", { color: C.accent4, fontSize: 14 })], 0.6, 3.1, 6.2, 0.75, { fontSize: 17 });
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
    text(s, [b("Why is it rare?", { color: C.accent3, fontSize: 20, breakLine: true }), t("왜 드문가?", { color: C.accent5, fontSize: 14 })], 7.6, 3.25, 4.9, 0.8, {});
    text(s, bl([
      [t("Its effects could "), b("spiral out of control", { color: C.accent3 }), t(" → relatively rare in biological systems.")],
      [t("Under normal physiological conditions, positive feedback events are "), b("eventually halted by negative feedback", { color: C.accent5 }), t(".")],
      [t("It also occurs in nonliving systems (→ Earth Watch).")],
    ], { gap: 8, indent: 14 }), 7.6, 4.1, 4.9, 2.5, { fontSize: 15, color: C.background1 });
    s.addNotes(`[약 1분 30초]
양성피드백(positive feedback)에서는 하나의 변화가 그 변화를 증폭시키는 결과를 초래합니다.
온도조절기가 고장 났다고 상상해 보세요. 온도 설정을 올렸는데, 방이 설정 온도에 도달해도 조절기가 꺼지는 대신 더 많은 열을 요구하고, 또 요구하고… 매우 불편한 상황이 됩니다.
그 효과가 통제 불능으로 될 수 있기 때문에, 양성피드백은 생물학적 시스템에서는 상대적으로 드물며, 정상적인 생리 조건하에서 양성피드백은 음성피드백에 의해 결국 중단됩니다.
양성피드백은 무생물계에서도 일어날 수 있습니다(뒤의 Earth Watch).`);
  }

  // ---------- 21. Childbirth + A&P Fig 1.11
  {
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
    s.addNotes(`[약 2분]
대표적인 사례가 사람과 다른 포유류에서의 출산입니다.
출산 초기의 수축은 태아의 머리를 자궁의 아래쪽에 위치한 자궁경부로 밀면서 자궁경부를 팽창시켜 열기 시작합니다. 자궁경부의 신경세포들은 스트레칭에 반응해 시상하부로 신호를 보내고, 시상하부는 호르몬 옥시토신의 방출을 시작합니다. 옥시토신은 더 강한 자궁수축을 촉진하여 태아의 머리가 자궁경부를 더 확장하게 하며, 더 많은 옥시토신의 방출을 촉진합니다.
자궁수축이 태아의 출산을 유도하면 자궁경부의 압력이 완화되므로 음성피드백이 이 양성피드백 주기를 종결합니다.
양성피드백의 또 다른 예는 모유 수유 기간 동안 우유를 방출하는 하강반사(let-down reflex)이며, 여기에도 옥시토신과 시상하부가 관여합니다(38장). 아기가 배가 불러 젖빨기를 멈추면(음성피드백) 이 사이클은 멈춥니다.
※ 왼쪽 OpenStax 그림에는 '뇌하수체(pituitary gland)가 옥시토신을 분비'한다고 되어 있습니다. 우리 교재는 '시상하부가 옥시토신 방출을 일으킨다'고 서술하는데, 시상하부의 신호를 받아 뇌하수체가 혈액으로 내보낸다고 이해하면 됩니다(호르몬은 38장).`);
  }

  // ---------- 22. Earth Watch facts + photo
  {
    const s = contentSlide("4. Positive Feedback", "EARTH WATCH", T_EARTH, "Positive Feedback in the Arctic", "북극의 양성피드백 — 줄어드는 북극의 얼음");
    const img = await photo(PHOTO.bear.f, 5.6, 4.35, { pos: "centre", px: 1200 });
    s.addImage({ data: img, x: 0.6, y: 1.95, w: 5.6, h: 4.35, altText: "Polar bear on Arctic drift ice" });
    text(s, "Polar bears depend on Arctic ice to hunt and rest.", 0.6, 6.33, 5.6, 0.3, { fontSize: 13, italic: true });
    credit(s, PHOTO.bear.credit, 0.6, 6.62, 5.6);
    const x0 = 6.55, w0 = 6.18;
    text(s, [t("Climate scientists are convinced that global climate change has triggered positive feedback in the Arctic — the result may be the "), b("planetary equivalent of hyperthermia", { color: C.accent1 }), t(".")], x0, 1.95, w0, 1.05, { fontSize: 15 });
    text(s, [t("The North Pole (≈400 miles from the nearest land) was formerly covered with ice "), b("2–5 m"), t(" (6–15 ft) thick.")], x0, 3.0, w0, 0.7, { fontSize: 15 });
    const stats = [
      { big: "2×", d: "Arctic warming vs. the global average", col: C.accent1 },
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
    s.addNotes(`[Earth Watch · 약 1분 30초]
양성피드백에서 변화하는 조건들은 그 변화를 더 증가시키는 반응들을 일으킵니다. 기상학자들은 세계의 기후 변화가 북극의 양성피드백을 촉발시키고, 이러한 결과가 지구 고열화(행성 규모의 고체온증)를 일으킬 수 있다고 확신합니다.
북극은 가장 가까운 대륙으로부터 약 400마일 정도 떨어져 있고, 이전에는 2에서 5미터 두께의 얼음으로 덮여 있었습니다. 그러나 지구의 온도가 올라가면서 북극은 지구 평균보다 두 배나 빠르게 온도가 상승하고 있습니다.
그 결과 지난 30년 동안 북극의 만년설은 거의 50% 더 얇아지고 면적은 35% 줄어들었으며(그림 E32-1), 현재 10년마다 평균 약 10%씩 줄어들고 있습니다.
북극곰들은 북극의 얼음에 의존하여 사냥과 휴식을 취합니다.`);
  }

  // ---------- 23. Ice-albedo loop
  {
    const s = contentSlide("4. Positive Feedback", "EARTH WATCH", T_EARTH, "Ice Reflects, Water Absorbs", "얼음은 태양빛을 반사하고, 바닷물은 흡수한다 — 전통적인 양성피드백의 일례");
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
      { x: cx + rx, y: cy, en: "More water exposed", kr: "더 많은 물이 노출", col: C.text2 },
      { x: cx, y: cy + ry, en: "Water absorbs more sunlight", kr: "더 많은 태양빛 흡수", col: C.text2 },
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
    s.addNotes(`[Earth Watch · 약 2분]
여기에서 양성피드백은 어떻게 작용할까요? 해답은 노출된 바닷물과 비교하였을 때 얼마나 많은 태양빛이 얼음에 의해 반사되는지에 있습니다.
해빙은 일반적으로 50%에서 70%의 태양빛을 반사해 우주로 돌려보내고, 얼음을 뒤덮은 눈은 태양빛의 80% 정도를 반사합니다. 반면 바닷물은 단지 5%에서 10%의 빛만을 반사하므로 거의 대부분의 태양 에너지가 흡수되어 북극의 물을 데웁니다.
북극의 얼음층이 녹으면서 더 많은 물이 노출되고, 노출된 물은 얼음에 비해 더 많은 태양빛을 흡수해 온도가 더 상승하고, 더 많은 얼음이 녹고, 또 다시 더 많은 물이 노출되고… 이것이 예측할 수 없는 전 지구적 결과를 가져오는 전통적인 양성피드백의 일례입니다.
[교재 생각해 보기 / Consider This] 이산화탄소와 다른 온실 가스들을 대기로 뿜어내면서 우리는 의도하지 않은 지구의 생명유지 시스템을 실험하고 있다. 무엇이 결국 이 사이클을 멈추게 할까? → 토론 슬라이드에서 다시 다룹니다.`);
  }

  // ---------- 24. Case revisited
  pres.addSection({ title: "Wrap-up" });
  {
    const s = contentSlide("Wrap-up", "CASE STUDY · REVISITED", T_CASE, "Case Study Revisited: Who Is at Risk?", "열사병은 운동선수만의 문제가 아니다");
    text(s, [t("When the air temperature becomes extremely high — especially with "), b("high humidity"), t(" — people can die of heat stroke while sitting in their living rooms.")], 0.6, 1.92, 12.13, 0.72, { fontSize: 16 });
    const grp = [
      { I: md.MdElderly, en: "The elderly", kr: "노인", d: "during summer heat waves — far more likely than athletes", col: C.accent1 },
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
    text(s, [b("Consider this: ", { color: C.text2 }), t("in the 2003 European heat wave, 35,000–52,000 “excess deaths” occurred. Should they count as a cost of inaction on CO₂ emissions?")], 7.4, 5.55, 5.15, 1.15, { fontSize: 13, valign: "middle" });
    s.addNotes(`[Case Study Revisited · 약 1분 30초]
열사병은 운동선수만의 문제가 아닙니다. 기온이 극도로 높고 특히 습도까지 높으면, 사람들은 거실에 앉아 있다가도 열사병으로 사망할 수 있습니다. 운동선수가 뉴스에 더 많이 나오지만, 열사병 사망은 여름 폭염 중의 노인과 어린이에게서 훨씬 더 많이 일어나고, 뜨거운 태양 아래서 일하는 농업 노동자도 상당한 위험에 처해 있습니다.
어린이는 특히 취약합니다. 아주 어린 아이들은 체온 항상성이 완전히 발달하지 않았고, 같은 조건에서 어른보다 열을 더 많이 내며, 땀은 덜 흘립니다. 미국에서는 매년 약 35~40명의 어린이가 차 안에 혼자 남겨져 열사병으로 사망합니다. 바깥이 32℃(90℉)를 넘으면 닫힌 차 안은 1시간도 안 되어 60℃(140℉)까지 치솟습니다. 어떤 볼일도 이 위험을 정당화할 수 없습니다.
[Consider This] 2003년 유럽 폭염 때 35,000~52,000명의 '초과 사망'이 있었습니다. 이산화탄소 감축 비용을 계산할 때 이런 초과 사망을 '행동하지 않은 비용'에 포함해야 할까요?`);
  }

  // ---------- 25. Key concepts (dark)
  {
    const s = contentSlide("Wrap-up", null, null, "Key Concepts of 32.1", "핵심개념의 요약", "CONTENT_DARK");
    const kc = [
      { I: fa.FaBalanceScale, col: C.background1, h: "Homeostasis", kr: "항상성", d: "a dynamic equilibrium: internal conditions (temperature, salt, oxygen, glucose, pH, water) are kept within a range in which proteins can function and energy can be made available." },
      { I: fa.FaThermometerHalf, col: C.accent3, h: "Ectotherms vs. endotherms", kr: "외온동물 vs. 내온동물", d: "Ectotherms get most body warmth from the environment; endotherms get most heat from metabolic activities and regulate body temperature within a narrow range." },
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
    s.addNotes(`[요약 · 약 1분 30초]
교재의 '핵심개념의 요약'입니다.
항상성은 동물체 내부 상태(온도, 염, 산소, 당, pH, 수분의 정도 포함)의 동적인 평형상태를 말하며, 단백질이 기능할 수 있고 에너지가 이용될 수 있는 범위 내에서 유지됩니다.
동물들은 서로 다른 온도 조절방식을 가지고 있습니다. 외온동물은 환경으로부터 대부분의 신체 따뜻함을 얻고, 내온동물은 대사 활동으로부터 열의 대부분을 얻으며 좁은 범위 내에서 체온을 조절하는 경향이 있습니다.
항상성 상태는 어떤 변화가 그 변화에 대한 반대 반응을 촉진시킴으로써 원래의 상태로 돌아가도록 하는 음성피드백을 통해서 유지됩니다. 양성피드백은 변화가 변화를 강하게 하지만, 이러한 상황들은 모두 자기 제한적입니다.`);
  }

  // ---------- 26. Quick check
  {
    const s = contentSlide("Wrap-up", "REVIEW", T_REV, "Quick Check", "개념을 통한 이해 — 선택형 문제와 빈칸 채우기 (교재 Chapter Review)");
    const qs = [
      { q: "1. Which of the following statements is FALSE?", o: ["a. Overheating to 108°F (42°C) is likely to disrupt covalent bonds in enzymes.", "b. Overly acidic conditions can disrupt protein structure.", "c. Low body temperature will reduce the rate at which enzymes catalyze reactions.", "d. Considerable metabolic activity is required to maintain a relatively constant environment."] },
      { q: "2. Which of the following statements is TRUE?", o: ["a. Ectotherms generate most of their body heat metabolically.", "b. Ectotherms include birds and mammals.", "c. Ectotherms experience variable body temperatures.", "d. Ectotherms are known as “warm-blooded” animals."] },
      { q: "3. Negative feedback systems do NOT …", o: ["a. include thermostats that maintain room temperature.", "b. include a sensor, a control center, and an effector.", "c. control body temperature.", "d. regulate uterine contractions during labor."] },
    ];
    const pos = [[0.6, 1.95, 6.0, 2.45], [6.73, 1.95, 6.0, 2.45], [0.6, 4.55, 6.0, 2.2]];
    for (let i = 0; i < 3; i++) {
      const [x, y, w, h] = pos[i];
      card(s, x, y, w, h, { fill: C.background2 });
      text(s, [b(qs[i].q, { color: C.text2, breakLine: true }), ...qs[i].o.map((o, j) => t(o, { breakLine: j < 3 }))], x + 0.25, y + 0.15, w - 0.45, h - 0.25, { fontSize: 13, paraSpaceAfter: 3 });
    }
    card(s, 6.73, 4.55, 6.0, 2.2, { fill: C.accent3, transparency: 80 });
    text(s, [b("Fill in the blank", { color: C.text2, breakLine: true }), t("The ability of the body to maintain its internal conditions within the narrow range required by cells to function is called ________. The most important general mechanism for maintaining these conditions is ________.", { breakLine: true }), t("신체가 기능하도록 좁은 범위 내에서 내부 조건을 유지하는 능력을 ____이라 부르고, 이를 유지하는 가장 중요한 메커니즘은 ____이다.", { color: C.accent4, fontSize: 12 })], 6.98, 4.7, 5.55, 1.95, { fontSize: 13, paraSpaceAfter: 4 });
    s.addNotes(`[복습 퀴즈 · 약 2분]
교재 Chapter Review의 선택형 문제 1~3번과 빈칸 채우기 1번입니다. 청중에게 잠깐 시간을 주고 손을 들어 답하게 하면 좋습니다.
(한국어판 1번 a 보기는 '42℃로 과열하면 효소의 공유결합은 파괴될 수 있다'입니다.)
정답은 다음 슬라이드에 있습니다.`);
  }

  // ---------- 27. Answers
  {
    const s = contentSlide("Wrap-up", "REVIEW", T_REV, "Quick Check: Answers", "정답과 해설 — 근거는 모두 교재 32.1 본문");
    const ans = [
      { n: "1", a: "a", why: [t("Heat disrupts the "), b("hydrogen bonds"), t(" that hold an enzyme’s 3-D shape — not covalent bonds. (b), (c), (d) are all true.")] },
      { n: "2", a: "c", why: [t("Ectotherms get most heat from the environment, so their body temperatures vary with external temperature.")] },
      { n: "3", a: "d", why: [t("Uterine contractions during labor are driven by "), b("positive feedback", { color: C.accent1 }), t(" (oxytocin), which is ended by negative feedback at birth.")] },
      { n: "✎", a: "homeostasis · negative feedback", why: [t("항상성 · 음성피드백")] },
    ];
    for (let i = 0; i < 4; i++) {
      const y = 1.95 + i * 1.2;
      card(s, 0.6, y, 12.13, 1.05, { fill: i === 3 ? C.accent3 : C.background2, transparency: i === 3 ? 80 : undefined });
      s.addText(ans[i].n, { shape: S.OVAL, x: 0.8, y: y + 0.2, w: 0.65, h: 0.65, fill: { color: C.text2 }, line: { type: "none" }, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
      text(s, ans[i].a, 1.65, y, i === 3 ? 4.6 : 1.0, 1.05, { fontSize: i === 3 ? 20 : 30, bold: true, color: C.accent6, valign: "middle", fontFace: "+mj-lt" });
      text(s, ans[i].why, i === 3 ? 6.4 : 2.75, y, i === 3 ? 6.1 : 9.75, 1.05, { fontSize: 15, valign: "middle" });
    }
    s.addNotes(`[정답 · 약 1분]
1번 정답 a: 과열로 파괴되는 것은 효소의 3차 구조를 유지하는 수소결합이지 공유결합이 아닙니다. b(과도한 산성은 단백질 구조 파괴), c(저체온은 효소 반응 속도 감소), d(일정한 환경 유지에는 상당한 대사 활동 필요)는 모두 맞는 설명입니다.
2번 정답 c: 외온동물은 환경으로부터 열을 얻으므로 체온이 외부 기온에 따라 변합니다. a·b·d는 내온동물에 대한 설명이거나 틀린 용어입니다.
3번 정답 d: 분만 시 자궁 수축은 옥시토신에 의한 양성피드백이며, 출산 후 음성피드백으로 종결됩니다.
빈칸: 항상성(homeostasis), 음성피드백(negative feedback).
(정답과 해설은 교재 32.1 본문 내용을 근거로 정리한 것입니다. 발표 전에 교재 뒤 Answers 부분과 한 번 대조해 확인하세요.)`);
  }

  // ---------- 28. Discussion (dark)
  {
    const s = contentSlide("Wrap-up", null, null, "Let’s Discuss", "생각해 볼 문제 — 교재의 Think Critically · Applying the Concepts · Consider This", "CONTENT_DARK");
    const qs = [
      { I: fa.FaBrain, col: C.background1, h: "Think Critically (Fig. 32-2)", q: "What would happen if a cold, shivering mammal ingested a poison that destroyed all of its body’s nerve endings that detect heat?", kr: "열을 감지하는 신경말단이 모두 파괴된다면?" },
      { I: fa.FaWater, col: C.accent5, h: "Applying the Concepts 1", q: "Why does life on land present particular difficulties in maintaining temperature homeostasis compared to life in water?", kr: "육지의 삶이 수중보다 체온 항상성 유지가 어려운 이유는?" },
      { I: fa.FaBaby, col: C.accent1, h: "Applying the Concepts 3", q: "Design a machine with sensors, electrical currents, and motors that illustrates the feedback relationships in labor and childbirth.", kr: "진통과 출산의 피드백을 장비에 비유해 설계한다면?" },
      { I: fa.FaGlobeAmericas, col: C.accent3, h: "Consider This (Earth Watch)", q: "By pumping CO₂ and other greenhouse gases into the atmosphere, we are inadvertently experimenting with the life-support systems of our planet. What do you think will eventually stop this cycle?", kr: "무엇이 결국 이 사이클을 멈추게 할까?" },
    ];
    for (let i = 0; i < 4; i++) {
      const x = 0.6 + (i % 2) * 6.17, y = 1.98 + Math.floor(i / 2) * 2.45;
      card(s, x, y, 5.96, 2.25, { fill: C.background1, transparency: 92, r: 0.16 });
      await iconCircle(s, x + 0.3, y + 0.28, 0.7, qs[i].I, { bg: qs[i].col, fg: HEX.navy, name: qs[i].h });
      text(s, qs[i].h, x + 1.2, y + 0.28, 4.6, 0.7, { fontSize: 17, bold: true, color: qs[i].col, valign: "middle" });
      text(s, [t(qs[i].q, { color: C.background1, breakLine: true }), t(qs[i].kr, { color: C.accent5, fontSize: 13 })], x + 0.3, y + 1.05, 5.4, 1.15, { fontSize: 14 });
    }
    s.addNotes(`[토론 · 시간이 남으면 약 3~5분]
청중과 함께 생각해 볼 질문들입니다. 아래는 교재 내용을 바탕으로 한 '예시 답안 방향'이니 발표 전에 조원들과 다듬으세요.
① 비판적 사고(그림 32-2): 감지기(열을 감지하는 신경말단)가 사라지면 체온이 회복되어도 그 정보가 시상하부로 전달되지 않습니다. 그러면 조절 센터가 설정점과 비교할 수 없어 음성피드백 고리가 끊어지고, 떨림 같은 열 생산 반응을 적절히 멈추지 못할 수 있습니다.
② 개념의 적용 1: 교재 예시처럼 심해는 0~3℃로 거의 일정해 심해어는 일생 동안 체온 변화를 거의 겪지 않습니다. 반면 육지의 공기 온도는 시간·일·계절에 따라 크게 변하므로(사막의 낮과 밤 등) 체온을 일정하게 유지하기가 더 어렵습니다.
③ 개념의 적용 3: 감지기 = 자궁경부의 신장 감지 센서, 전류 = 시상하부로 가는 신경 신호, 모터 = 옥시토신에 의해 더 강해지는 자궁 근육 수축 → 출산(압력 해소)이라는 '정지 스위치'가 음성피드백 역할.
④ Consider This: 정답이 정해진 질문은 아닙니다. 예: 얼음이 모두 녹아 더 이상 반사면이 없어질 때, 혹은 온실가스 배출 감소 같은 인간의 개입 등 의견을 나눠 봅니다.`);
  }

  // ---------- 29. References
  {
    const s = contentSlide("Wrap-up", "REFERENCES", T_REV, "References & Image Credits", "참고 자료 및 이미지 출처");
    text(s, bl([
      [b("Main textbook: "), t("Audesirk, Audesirk & Byers, "), t("Biology: Life on Earth", { italic: true }), t(", Ch. 32 “Homeostasis and the Organization of the Animal Body” — 32.1 (English ed. pp. 604–609, Chapter Review pp. 617–618; Korean ed. pp. 602–607, 615–616).")],
      [b("Figures: "), t("OpenStax, "), t("Anatomy and Physiology 2e", { italic: true }), t(", Figs. 1.10 & 1.11; OpenStax, "), t("Biology 2e", { italic: true }), t(", Fig. 33.23 — openstax.org, licensed CC BY 4.0.")],
      [b("Photos (Wikimedia Commons): "), t("desert pupfish — P. V. Loiselle (CC BY-SA 3.0); cinnamon hummingbird — Charles J. Sharp (CC BY-SA 4.0); green iguana — Cayambe (CC BY-SA 3.0); polar bear — Andreas Weith (CC BY-SA 4.0); Claude Bernard — public domain; Walter B. Cannon — Bachrach 1934 (CC BY 4.0).")],
      [b("Diagrams: "), t("negative feedback loop (after textbook Fig. 32-2), heat-stroke vicious cycle, Arctic ice loop, and schematic graphs were drawn for this presentation from the textbook text. Icons: react-icons (Font Awesome, Game-icons, Material, Tabler).")],
    ], { gap: 10, indent: 14 }), 0.6, 1.95, 12.13, 4.8, { fontSize: 15 });
    s.addNotes(`참고 자료와 이미지 출처입니다. 사진과 OpenStax 그림은 모두 재사용 가능한 라이선스(CC BY / CC BY-SA / 퍼블릭 도메인)이며, 출처 표기가 조건입니다.`);
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
    zip.file(name, sx);
  }
  console.log("paragraphs repaired:", fixedParas);
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
