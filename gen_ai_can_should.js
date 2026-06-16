"use strict";
const pptxgen = require("pptxgenjs");
const fs      = require("fs");
const path    = require("path");
const React   = require("react");
const RDS     = require("react-dom/server");
const sharp   = require("sharp");

// ── Assets ────────────────────────────────────────────────────────────────────
const SKILL_DIR = "/Users/arminas.juknevicius/.claude/skills/domo-branded-pptx/assets";
const img64 = fp => { const b = fs.readFileSync(fp); return `image/${path.extname(fp).slice(1)};base64,` + b.toString("base64"); };
const bgGradient  = img64(SKILL_DIR + "/domo_gradient_bg.png");
const logoPrimary = img64(SKILL_DIR + "/domo_logo_blue_800.png");
const logoIcon    = img64(SKILL_DIR + "/domo_icon_400.png");

// ── Colours ───────────────────────────────────────────────────────────────────
const C = {
  domoBlue:"99CCEE", white:"FFFFFF", neutral1:"F1F6FA", neutral2:"DCE4EA",
  neutral3:"B7C1CB", neutral4:"68737F", charcoal:"3F454D",
  accentOrange:"FF9922", accentPurple:"776CB0", accentGreen:"ADD4C1",
};
const F = "Open Sans";
const mkShadow = () => ({ type:"outer", color:"000000", blur:6, offset:2, angle:135, opacity:0.10 });

// ── Icons ─────────────────────────────────────────────────────────────────────
const {
  FaLock, FaLayerGroup, FaEye, FaShieldAlt, FaCoins,
  FaCheckCircle, FaRobot, FaDatabase, FaCode, FaComments,
  FaArrowRight, FaUsers, FaKey, FaClipboardList, FaFilter,
  FaBolt, FaCogs, FaChartLine, FaServer, FaUserShield,
  FaIndustry, FaExclamationTriangle,
} = require("react-icons/fa");

async function icon64(Ic, color = "#99CCEE", size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Ic, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function addFooter(s, pres, n) {
  s.addShape(pres.shapes.LINE,      { x:0.5, y:5.15, w:9.0, h:0, line:{color:C.neutral3, width:0.5} });
  s.addImage({ data:logoIcon,         x:0.35, y:5.2,  w:0.38, h:0.38 });
  s.addText("CONFIDENTIAL",          { x:7.0, y:5.2, w:2.0, h:0.3, fontFace:F, fontSize:8, color:C.neutral4, align:"right", margin:0 });
  s.addText(String(n),               { x:9.2, y:5.2, w:0.4, h:0.3, fontFace:F, fontSize:9, bold:true, color:C.charcoal, align:"center", margin:0 });
}

// hasChip: pass true when a lensChip() sits at x:8.5 — caps title+subtitle to w:7.5
// lines: number of lines the title wraps to (default 1). Each extra line adds ~0.42" to the
// title box, pushing the accent bar and subtitle down so they never overlap the title text.
// Content on the calling slide must start at: accentY + 0.10 + 0.28 + 0.04 = accentY + 0.42
//   lines=1 → accentY=0.88, content ≥ y:1.30  (existing 1-line slides use y:1.28, fine)
//   lines=2 → accentY=1.30, content ≥ y:1.72
//   lines=3 → accentY=1.72, content ≥ y:2.14
function titleBar(s, pres, title, sub, hasChip = false, lines = 1) {
  const tw = hasChip ? 7.5 : 8.8;
  const titleH = 0.58 + (lines - 1) * 0.42;
  const accentY = 0.28 + titleH + 0.02;
  s.addText(title, { x:0.6, y:0.28, w:tw, h:titleH, fontFace:F, fontSize:30, bold:true, color:C.charcoal, margin:0, lineSpacingMultiple:0.95 });
  s.addShape(pres.shapes.RECTANGLE,  { x:0.6, y:accentY, w:1.4, h:0.04, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
  if (sub) s.addText(sub,            { x:0.6, y:accentY+0.10, w:tw, h:0.28, fontFace:F, fontSize:10, color:C.neutral4, margin:0 });
}

function lensChip(s, label) {
  // Small "LENS X" pill top-right
  s.addShape("roundRect", { x:8.5, y:0.28, w:1.0, h:0.28, fill:{color:C.domoBlue}, rectRadius:0.06 });
  s.addText(label, { x:8.5, y:0.28, w:1.0, h:0.28, fontFace:F, fontSize:7.5, bold:true, color:C.charcoal, align:"center", valign:"middle", margin:0 });
}

// Circle with number (w===h always)
function numberedCircle(s, pres, n, x, y, sz, bgColor, txtColor, fontSize) {
  s.addShape(pres.shapes.OVAL, { x, y, w:sz, h:sz, fill:{color:bgColor} });
  s.addText(String(n),          { x, y, w:sz, h:sz, fontFace:F, fontSize, bold:true, color:txtColor, align:"center", valign:"middle", margin:0 });
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title  = "Just because AI Can doesn't mean it Should";
  pres.author = "Ben Schein, Chief AI & Analytics Officer, Domo";

  // Pre-render icons
  const icLock    = await icon64(FaLock);
  const icLayer   = await icon64(FaLayerGroup);
  const icEye     = await icon64(FaEye);
  const icShield  = await icon64(FaShieldAlt);
  const icCoins   = await icon64(FaCoins);
  const icCheck   = await icon64(FaCheckCircle);
  const icRobot   = await icon64(FaRobot);
  const icDB      = await icon64(FaDatabase);
  const icCode    = await icon64(FaCode);
  const icChat    = await icon64(FaComments);
  const icArrow   = await icon64(FaArrowRight);
  const icUsers   = await icon64(FaUsers);
  const icKey     = await icon64(FaKey);
  const icClip    = await icon64(FaClipboardList);
  const icFilter  = await icon64(FaFilter);
  const icBolt    = await icon64(FaBolt);
  const icCogs    = await icon64(FaCogs);
  const icChart   = await icon64(FaChartLine);
  const icServer  = await icon64(FaServer);
  const icUserSh  = await icon64(FaUserShield);
  const icInd     = await icon64(FaIndustry);

  // ── Screenshots from original deck ────────────────────────────────────────
  const SCR_DIR = SKILL_DIR + "/screenshots";
  const scr10 = img64(SCR_DIR + "/s10_domo_text_gen.png");
  const scr11 = img64(SCR_DIR + "/s11_post_summary.png");
  const scr12 = img64(SCR_DIR + "/s12_agent_general.png");
  const scr13 = img64(SCR_DIR + "/s13_agent_knowledge.png");
  const scr14 = img64(SCR_DIR + "/s14_agent_tools.png");
  const scr18 = img64(SCR_DIR + "/s18_ai_activity_log.png");

  // ── Pillar list helper (right-side nav for slides 10–14) ──────────────────
  // items: [{icon, label}], highlighted: 0-based index to call out
  function pillarList(s, pres, x, y, w, items, highlighted, ih = 0.82, gap = 0.07) {
    items.forEach((it, i) => {
      const iy = y + i * (ih + gap);
      const isDark = i === highlighted;
      const bg = isDark ? C.charcoal : C.neutral1;
      const tc = isDark ? C.white    : C.charcoal;
      const circFill = isDark ? C.neutral1 : C.charcoal;
      s.addShape(pres.shapes.RECTANGLE, { x, y:iy, w, h:ih, fill:{color:bg}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x, y:iy, w, h:0.05, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
      const cs = 0.42; const oy = iy + (ih - cs) / 2;
      s.addShape(pres.shapes.OVAL, { x:x+0.18, y:oy, w:cs, h:cs, fill:{color:circFill} });
      s.addImage({ data:it.icon, x:x+0.18+(cs-0.27)/2, y:oy+(cs-0.27)/2, w:0.27, h:0.27 });
      s.addText(it.label, { x:x+0.72, y:iy, w:w-0.84, h:ih, fontFace:F, fontSize:13, bold:true, color:tc, valign:"middle", margin:0 });
    });
  }

  // ── Shared card helper ──────────────────────────────────────────────────────
  // fullWidthBody: body text spans full card width from the left margin, starting below the icon.
  // Use for tall cards where the indented-body layout looks unbalanced (e.g. 3.5" tall panels).
  function card(s, pres, x, y, w, h, { icon, title, body, dark = false, accentColor = null, fullWidthBody = false }) {
    const bg = dark ? C.charcoal : C.neutral1;
    const tc = dark ? C.white : C.charcoal;
    const bc = dark ? C.neutral3 : C.neutral4;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill:{color:bg}, shadow:mkShadow() });
    const accent = accentColor || C.domoBlue;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h:0.06, fill:{color:accent}, line:{color:accent} });
    if (icon) {
      const cs = 0.46; const circFill = dark ? C.neutral1 : C.charcoal;
      const cx = x + 0.18; const cy = y + 0.18;
      s.addShape(pres.shapes.OVAL, { x:cx, y:cy, w:cs, h:cs, fill:{color:circFill} });
      s.addImage({ data:icon, x:cx + (cs-0.3)/2, y:cy + (cs-0.3)/2, w:0.3, h:0.3 });
    }
    const tx = icon ? x + 0.78 : x + 0.18;
    const tw = icon ? w - 0.9  : w - 0.36;
    const tY = y + 0.18;
    // body positioning: full-width starts at card left margin below the icon row; default indents under title
    const bx = (fullWidthBody && icon) ? x + 0.18 : tx;
    const bw = (fullWidthBody && icon) ? w - 0.36  : tw;
    const bY = (fullWidthBody && icon) ? y + 0.72  : tY + 0.42;
    const bH = (fullWidthBody && icon) ? h - 0.90  : h - 0.65;
    s.addText(title, { x:tx, y:tY, w:tw, h:0.38, fontFace:F, fontSize:11, bold:true, color:tc, margin:0, lineSpacingMultiple:0.95 });
    if (body) s.addText(body, { x:bx, y:bY, w:bw, h:bH, fontFace:F, fontSize:9.5, color:bc, margin:0, lineSpacingMultiple:1.15 });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 1 — Title
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { data:bgGradient };
    s.addImage({ data:logoPrimary, x:0.0, y:0.0, w:1.0, h:1.0 });
    s.addText("A PRACTITIONER'S FRAMEWORK FOR ENTERPRISE AI", {
      x:0.65, y:1.1, w:8.5, h:0.35, fontFace:F, fontSize:11, bold:true, color:C.neutral4,
      charSpacing:1.5, margin:0,
    });
    // "Just because AI" — single line, 42pt ~0.58" tall
    s.addText("Just because AI", {
      x:0.65, y:1.45, w:8.5, h:0.62, fontFace:F, fontSize:42, bold:true, color:C.charcoal, margin:0, lineSpacingMultiple:0.95,
    });
    // '"can" doesn't mean it "should"' — wraps to 2 lines at 42pt; needs h ≥ 1.15"
    s.addText([
      { text:'"can"',   options:{ bold:true, color:C.domoBlue } },
      { text:" doesn't mean it ", options:{ bold:true, color:C.charcoal } },
      { text:'"should"', options:{ bold:true, color:C.accentOrange } },
    ], { x:0.65, y:2.12, w:8.5, h:1.25, fontFace:F, fontSize:42, margin:0, lineSpacingMultiple:0.95 });
    s.addShape(pres.shapes.RECTANGLE, { x:0.65, y:3.46, w:2.0, h:0.04, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
    s.addText("A framework for deciding when, where, and how AI belongs in the work.", {
      x:0.65, y:3.58, w:7.0, h:0.38, fontFace:F, fontSize:13, color:C.neutral4, margin:0,
    });
    s.addText("Ben Schein  ·  Chief AI & Analytics Officer  ·  Domo", {
      x:0.65, y:4.62, w:7.0, h:0.35, fontFace:F, fontSize:11, color:C.charcoal, margin:0,
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 2 — Where We're Starting
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 2);
    titleBar(s, pres, "Where We're Starting", "AI is going to reshape how we work. That part isn't in question.");
    const points = [
      { icon:icRobot, title:"It's not about the model.", body:"The capability curve is moving so fast that no one wins by picking the best model." },
      { icon:icChart, title:"It's not about the demo.",  body:"Demos answer \"can.\" Production answers \"should.\" Those are very different questions." },
      { icon:icFilter,title:"It is about judgment.",     body:"Which use cases. Which guardrails. Which models. Which moments to let AI run, and which moments to keep humans in the lead." },
    ];
    // 3×(1.10+0.18)−0.18+1.10 = 4.98 — last row ends at 4.98, clear of footer
    const ch = 1.10, cy = 1.32, gap = 0.18;
    points.forEach((p, i) => {
      card(s, pres, 0.55, cy + i*(ch+gap), 8.9, ch, { icon:p.icon, title:p.title, body:p.body });
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 3 — Understanding AI's Impact
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 3);
    titleBar(s, pres, "Understanding AI's Impact", "Where the value actually shows up.");
    const cols = [
      { label:"EASY",   title:"Capability",  body:"What the model can do in a demo.",                                            icon:icBolt,  dark:false },
      { label:"HARD",   title:"Deployment",  body:"Where AI actually fits in real workflows.",                                   icon:icCogs,  dark:true  },
      { label:"HARDER", title:"Outcomes",    body:"Whether it created value: measurable, durable, governed.",                   icon:icChart, dark:true  },
    ];
    const cw = 2.85, cy = 1.28, ch = 3.30, gap = 0.2;
    cols.forEach((c, i) => {
      const cx = 0.55 + i*(cw+gap);
      const bg = c.dark ? C.charcoal : C.neutral1;
      const tc = c.dark ? C.white    : C.charcoal;
      const bc = c.dark ? C.neutral3 : C.neutral4;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:ch, fill:{color:bg}, shadow:mkShadow() });
      // Difficulty label
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:0.42, fill:{color:C.domoBlue} });
      s.addText(c.label, { x:cx, y:cy, w:cw, h:0.42, fontFace:F, fontSize:11, bold:true, color:C.charcoal, align:"center", valign:"middle", margin:0 });
      // Icon
      // circY centers the content block (icon 0.56 + gap 0.14 + title 0.44 + gap 0.06 + body 0.55 = 1.75)
      // within available height (ch 3.30 - header 0.42 = 2.88): padding = (2.88-1.75)/2 = 0.565 ≈ 0.57
      const cs = 0.56, circY = cy+0.99;
      const circFill = c.dark ? C.neutral1 : C.charcoal;
      s.addShape(pres.shapes.OVAL, { x:cx+(cw-cs)/2, y:circY, w:cs, h:cs, fill:{color:circFill} });
      s.addImage({ data:c.icon, x:cx+(cw-cs)/2+(cs-0.36)/2, y:circY+(cs-0.36)/2, w:0.36, h:0.36 });
      s.addText(c.title, { x:cx+0.18, y:circY+0.70, w:cw-0.36, h:0.44, fontFace:F, fontSize:15, bold:true, color:tc, align:"center", margin:0 });
      s.addText(c.body,  { x:cx+0.18, y:circY+1.20, w:cw-0.36, h:0.55, fontFace:F, fontSize:10, color:bc,  align:"center", margin:0, lineSpacingMultiple:1.2 });
    });
    s.addText("Capability is the easy part. Deployment is where most companies are stuck.", {
      x:0.55, y:4.72, w:8.9, h:0.28, fontFace:F, fontSize:9, italic:true, color:C.neutral4, margin:0,
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 4 — Domo Platform: Foundation → Activation → Distribution
  // Retains original layout and logos (cloud providers + AI models),
  // with Domo colour rules applied.
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.charcoal };
    // ── Logo + title ──────────────────────────────────────────────────────────
    s.addImage({ data:logoPrimary, x:0.0, y:0.0, w:0.92, h:0.92 });
    s.addText("How does Domo drive towards outcomes?", {
      x:1.02, y:0.06, w:8.6, h:0.42, fontFace:F, fontSize:15, bold:true, color:C.white, margin:0,
    });
    s.addText("Building an Intelligent Enterprise: from raw data and AI to agentic outcomes.", {
      x:1.02, y:0.5,  w:8.6, h:0.28, fontFace:F, fontSize:9.5, color:C.neutral3, margin:0,
    });
    // ── Governance banner ─────────────────────────────────────────────────────
    // y:0.94 starts below logo bottom (y:0+h:0.92) to avoid overlap
    s.addShape(pres.shapes.RECTANGLE, { x:0.35, y:0.94, w:9.3, h:0.26, fill:{color:C.domoBlue} });
    s.addText("ENTERPRISE-GRADE GOVERNANCE", {
      x:0.35, y:0.94, w:9.3, h:0.26, fontFace:F, fontSize:8, bold:true, color:C.charcoal,
      align:"center", valign:"middle", margin:0, charSpacing:1.5,
    });
    // ── Column layout ─────────────────────────────────────────────────────────
    // FOUNDATION x:0.35 w:3.45 | ACTIVATION x:3.88 w:2.15 | DISTRIBUTION x:6.11 w:2.15 | PEOPLE x:8.34 w:1.0
    const colY = 1.24, colH = 3.55, hdrH = 0.36;
    const dark1 = "2A2F36", dark2 = "1E2328";

    // ── FOUNDATION ────────────────────────────────────────────────────────────
    {
      const cx = 0.35, cw = 3.45;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:colH, fill:{color:dark1} });
      // Header
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:hdrH, fill:{color:dark2} });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:0.04, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
      s.addText("FOUNDATION", { x:cx, y:colY, w:cw, h:hdrH, fontFace:F, fontSize:9, bold:true, color:C.domoBlue, align:"center", valign:"middle", margin:0, charSpacing:1 });
      // RAW DATA & AI MODELS sub-header
      const subY = colY + hdrH;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:subY, w:cw, h:0.26, fill:{color:dark2} });
      s.addText("RAW DATA & AI MODELS", { x:cx, y:subY, w:cw, h:0.26, fontFace:F, fontSize:7.5, bold:true, color:C.neutral3, align:"center", valign:"middle", margin:0, charSpacing:0.8 });
      // Cloud logos (Snowflake, Databricks, GCP, AWS, Azure) — original strip
      const scrCloudW = 2.9; const scrCloudH = 0.32;
      s.addImage({ data:img64(SCR_DIR + "/s04_cloud_logos.png"), x:cx+(cw-scrCloudW)/2, y:subY+0.3, w:scrCloudW, h:scrCloudH });
      // AI logos (Anthropic, OpenAI, etc.) — original strip
      const scrAiW = 2.6; const scrAiH = 0.32;
      s.addImage({ data:img64(SCR_DIR + "/s04_ai_logos.png"),    x:cx+(cw-scrAiW)/2,    y:subY+0.7, w:scrAiW,    h:scrAiH });
      // Separator
      s.addShape(pres.shapes.LINE, { x:cx+0.2, y:subY+1.1, w:cw-0.4, h:0, line:{color:C.neutral4, width:0.5} });
      // Foundation text
      s.addText("Build your AI-ready data foundation", { x:cx+0.2, y:subY+1.2, w:cw-0.4, h:0.42, fontFace:F, fontSize:11, bold:true, color:C.white, margin:0, lineSpacingMultiple:0.95 });
      s.addText("Integration + Transformation. Governed in your CDMS. Runs on your cloud, your AI inference, with governance by design.", {
        x:cx+0.2, y:subY+1.65, w:cw-0.4, h:1.2, fontFace:F, fontSize:9, color:C.neutral3, margin:0, lineSpacingMultiple:1.15,
      });
    }
    // Arrow FOUNDATION → ACTIVATION — centered in 0.08" gap; image extends slightly into columns
    s.addImage({ data:icArrow, x:3.76, y:colY+colH/2-0.08, w:0.16, h:0.16 });

    // ── ACTIVATION ───────────────────────────────────────────────────────────
    {
      const cx = 3.88, cw = 2.15;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:colH, fill:{color:dark1} });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:hdrH, fill:{color:dark2} });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:0.04, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
      s.addText("ACTIVATION", { x:cx, y:colY, w:cw, h:hdrH, fontFace:F, fontSize:9, bold:true, color:C.domoBlue, align:"center", valign:"middle", margin:0, charSpacing:1 });
      // Icon — neutral1 circle so domoBlue icon is visible on dark bg
      const cs = 0.52; const circX = cx+(cw-cs)/2; const circY = colY+hdrH+0.25;
      s.addShape(pres.shapes.OVAL, { x:circX, y:circY, w:cs, h:cs, fill:{color:C.neutral1} });
      s.addImage({ data:icRobot, x:circX+(cs-0.34)/2, y:circY+(cs-0.34)/2, w:0.34, h:0.34 });
      s.addText("Use AI Agents to activate it", { x:cx+0.15, y:circY+0.62, w:cw-0.3, h:0.55, fontFace:F, fontSize:11, bold:true, color:C.white, align:"center", margin:0, lineSpacingMultiple:0.95 });
      s.addText("Analytics. Apps. Automation. Agents help your people go further, faster.", {
        x:cx+0.15, y:circY+1.25, w:cw-0.3, h:1.6, fontFace:F, fontSize:9, color:C.neutral3, align:"center", margin:0, lineSpacingMultiple:1.15,
      });
    }
    // Arrow ACTIVATION → DISTRIBUTION — centered in 0.08" gap; image extends slightly into columns
    s.addImage({ data:icArrow, x:5.99, y:colY+colH/2-0.08, w:0.16, h:0.16 });

    // ── DISTRIBUTION ──────────────────────────────────────────────────────────
    {
      const cx = 6.11, cw = 2.15;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:colH, fill:{color:dark1} });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:hdrH, fill:{color:dark2} });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:0.04, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
      s.addText("DISTRIBUTION", { x:cx, y:colY, w:cw, h:hdrH, fontFace:F, fontSize:9, bold:true, color:C.domoBlue, align:"center", valign:"middle", margin:0, charSpacing:1 });
      const cs = 0.52; const circX = cx+(cw-cs)/2; const circY = colY+hdrH+0.25;
      s.addShape(pres.shapes.OVAL, { x:circX, y:circY, w:cs, h:cs, fill:{color:C.neutral1} });
      s.addImage({ data:icUsers, x:circX+(cs-0.34)/2, y:circY+(cs-0.34)/2, w:0.34, h:0.34 });
      s.addText("Distribute to people to drive outcomes", { x:cx+0.15, y:circY+0.62, w:cw-0.3, h:0.55, fontFace:F, fontSize:11, bold:true, color:C.white, align:"center", margin:0, lineSpacingMultiple:0.95 });
      s.addText("On mobile, embedded, or MCP, delivered in the workflows your people already use.", {
        x:cx+0.15, y:circY+1.25, w:cw-0.3, h:1.6, fontFace:F, fontSize:9, color:C.neutral3, align:"center", margin:0, lineSpacingMultiple:1.15,
      });
    }

    // ── PEOPLE OUTCOMES (narrow right column) ─────────────────────────────────
    {
      const cx = 8.34, cw = 1.0;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:colH, fill:{color:dark2} });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:colY, w:cw, h:0.04, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
      s.addText("PEOPLE", { x:cx, y:colY+0.06, w:cw, h:0.22, fontFace:F, fontSize:7.5, bold:true, color:C.domoBlue, align:"center", margin:0, charSpacing:0.5 });
      s.addText("OUTCOMES", { x:cx, y:colY+0.26, w:cw, h:0.22, fontFace:F, fontSize:7.5, bold:true, color:C.domoBlue, align:"center", margin:0, charSpacing:0.5 });
      const cs = 0.44; const circX = cx+(cw-cs)/2; const circY = colY+0.62;
      s.addShape(pres.shapes.OVAL, { x:circX, y:circY, w:cs, h:cs, fill:{color:C.neutral1} });
      s.addImage({ data:icChart, x:circX+(cs-0.28)/2, y:circY+(cs-0.28)/2, w:0.28, h:0.28 });
      s.addText("Measurable outcomes for the business.", {
        x:cx+0.08, y:circY+0.55, w:cw-0.16, h:2.3, fontFace:F, fontSize:8, color:C.neutral3, align:"center", margin:0, lineSpacingMultiple:1.15,
      });
    }

    // ── Bottom bar ────────────────────────────────────────────────────────────
    s.addShape(pres.shapes.RECTANGLE, { x:0.35, y:4.78, w:9.3, h:0.28, fill:{color:dark2} });
    s.addText("RUNS ON YOUR CLOUD  |  ON YOUR AI INFERENCE  |  WITH GOVERNANCE BY DESIGN", {
      x:0.35, y:4.78, w:9.3, h:0.28, fontFace:F, fontSize:7.5, color:C.neutral4, align:"center", valign:"middle", margin:0, charSpacing:0.5,
    });

    // ── Speaker note documenting changes from original ─────────────────────
    s.addNotes("CHANGES FROM ORIGINAL (slide 4):\n" +
      "- Layout retained: FOUNDATION / ACTIVATION / DISTRIBUTION / PEOPLE OUTCOMES columns with original cloud-provider logos (Snowflake, Databricks, GCP, AWS, Azure) and AI-model logos (Anthropic, OpenAI, Vertex AI, Gemini) embedded from the source file.\n" +
      "- Colour changes: orange/branded-blue column headers replaced with Domo Blue (#99CCEE) per Domo brand guide. Dark panel backgrounds changed to charcoal (#3F454D) and dark-variant (#2A2F36/#1E2328) instead of original navy tones.\n" +
      "- Icon circles changed from domoBlue fill to neutral1 fill (brand rule: domoBlue icon on domoBlue circle is invisible on dark backgrounds).\n" +
      "- Governance banner moved to top (below title) and styled in domoBlue.\n" +
      "SOURCE: AI_Can_vs_Should_Generic.pptx, slide 4.");
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 5 — The Framework: Five Lenses
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 5);
    titleBar(s, pres, "The Framework", "Five lenses for deciding when AI should. Each one is a filter. If it fails any of them, the answer is probably \"not here, not yet.\"");
    const lenses = [
      { n:1, icon:icLock,   title:"Lock It Down",              body:"Discovery, or production? Probabilistic, or deterministic?" },
      { n:2, icon:icLayer,  title:"Context Is the Moat",       body:"The model is commodity. Your context isn't." },
      { n:3, icon:icEye,    title:"Governance & Observability", body:"Can you see what AI is doing?" },
      { n:4, icon:icShield, title:"Security & Compliance",      body:"Can the data even go to the model?" },
      { n:5, icon:icCoins,  title:"Token Economics",            body:"Frequency, model, trigger: chosen on purpose?" },
    ];
    const lh = 0.62, ly0 = 1.38, gap = 0.10;
    lenses.forEach((l, i) => {
      const ly = ly0 + i*(lh+gap);
      s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:ly, w:8.9, h:lh, fill:{color:C.neutral1}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:ly, w:0.07, h:lh, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
      const cs = 0.44; const oy = ly+(lh-cs)/2;
      numberedCircle(s, pres, l.n, 0.72, oy, cs, C.domoBlue, C.charcoal, 14);
      const cs2 = 0.44; const oy2 = ly+(lh-cs2)/2;
      s.addShape(pres.shapes.OVAL, { x:1.3, y:oy2, w:cs2, h:cs2, fill:{color:C.charcoal} });
      s.addImage({ data:l.icon, x:1.3+(cs2-0.28)/2, y:oy2+(cs2-0.28)/2, w:0.28, h:0.28 });
      s.addText(l.title, { x:1.9, y:ly, w:2.6, h:lh, fontFace:F, fontSize:12, bold:true, color:C.charcoal, valign:"middle", margin:0 });
      s.addText(l.body,  { x:4.6, y:ly, w:4.75, h:lh, fontFace:F, fontSize:10, color:C.neutral4, valign:"middle", margin:0, lineSpacingMultiple:1.1 });
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 6 — Lens 1: Probabilistic vs Deterministic
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 6);
    lensChip(s, "LENS 1");
    titleBar(s, pres, "Two Ways a System Can Answer", "Probabilistic vs. Deterministic. Both have a place. Knowing which one a task wants is half the job.", true);
    const cols = [
      { label:"PROBABILISTIC", sub:"Same question. Different answer each time.", what:"What AI is", whatBody:"Generates plausible answers from a model.", best:"Open questions. Exploration. Drafting. Synthesis. Pattern-finding in messy data.", warn:"Answers that need to match every time. Audit. Compliance. High-volume reruns.", hdr:C.domoBlue, dark:false },
      { label:"DETERMINISTIC",  sub:"Same input. Same answer. Every time.",       what:"What code is", whatBody:"Same logic runs the same way each time.",  best:"Reproducible rollups. Compliance reporting. Anything you'll re-run at scale.", warn:"Brittleness when the question changes. Needs upfront work to define the logic.",   hdr:C.charcoal, dark:true  },
    ];
    const cw = 4.3, cy = 1.28, ch = 3.55, gap = 0.3;
    cols.forEach((c, i) => {
      const cx = 0.55 + i*(cw+gap);
      const bg = c.dark ? C.charcoal : C.neutral1;
      const tc = c.dark ? C.white    : C.charcoal;
      const bc = c.dark ? C.neutral3 : C.neutral4;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:ch, fill:{color:bg}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:0.5, fill:{color:c.hdr} });
      s.addText(c.label, { x:cx, y:cy,      w:cw, h:0.28, fontFace:F, fontSize:11, bold:true, color:c.dark?C.white:C.charcoal, align:"center", valign:"middle", margin:0 });
      s.addText(c.sub,   { x:cx, y:cy+0.28, w:cw, h:0.22, fontFace:F, fontSize:8.5,           color:c.dark?C.neutral3:C.neutral4, align:"center", margin:0 });
      let yy = cy+0.62;
      s.addText("What it is:", { x:cx+0.2, y:yy, w:cw-0.4, h:0.25, fontFace:F, fontSize:10, bold:true, color:tc, margin:0 }); yy+=0.28;
      s.addText(c.whatBody,   { x:cx+0.2, y:yy, w:cw-0.4, h:0.3,  fontFace:F, fontSize:9.5, color:bc, margin:0 }); yy+=0.42;
      s.addText("Best for:",  { x:cx+0.2, y:yy, w:cw-0.4, h:0.25, fontFace:F, fontSize:10, bold:true, color:tc, margin:0 }); yy+=0.28;
      s.addText(c.best,       { x:cx+0.2, y:yy, w:cw-0.4, h:0.55, fontFace:F, fontSize:9.5, color:bc, margin:0, lineSpacingMultiple:1.1 }); yy+=0.65;
      s.addText("Watch out for:", { x:cx+0.2, y:yy, w:cw-0.4, h:0.25, fontFace:F, fontSize:10, bold:true, color:tc, margin:0 }); yy+=0.28;
      s.addText(c.warn,       { x:cx+0.2, y:yy, w:cw-0.4, h:0.6,  fontFace:F, fontSize:9.5, color:bc, margin:0, lineSpacingMultiple:1.1 });
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 7 — Lens 1 In Practice: Text → Production
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 7);
    lensChip(s, "LENS 1");
    titleBar(s, pres, "From Text Message to Production App", "What started as an ad-hoc question became a system. The shape of the work changed at every stage.", true);
    const stages = [
      { n:"1", stage:"STAGE 1", head:"The Text",       sub:"CONVERSATION",  mode:"Probabilistic", body:"CEO sends a question over text. AI helps explore: what angles matter, what does the data look like, what's the right cut." },
      { n:"2", stage:"STAGE 2", head:"Walking the Data", sub:"ITERATION",   mode:"Probabilistic", body:"Working through the data live, with AI as a thinking partner. The answer isn't final yet, but the logic is forming." },
      { n:"3", stage:"STAGE 3", head:"Hardened",        sub:"PRODUCTION",   mode:"Deterministic", body:"Python in Magic ETL + a Domo app. Same answer every day. Audit-friendly. Zero AI cost on every refresh.", dark:true },
    ];
    const sw = 2.85, sy = 1.28, sh = 3.5, gap = 0.2;
    stages.forEach((st, i) => {
      const sx = 0.55 + i*(sw+gap);
      const dark = !!st.dark;
      const bg = dark ? C.charcoal : C.neutral1;
      const tc = dark ? C.white    : C.charcoal;
      const bc = dark ? C.neutral3 : C.neutral4;
      s.addShape(pres.shapes.RECTANGLE, { x:sx, y:sy, w:sw, h:sh, fill:{color:bg}, shadow:mkShadow() });
      // Stage badge
      s.addShape(pres.shapes.RECTANGLE, { x:sx, y:sy, w:sw, h:0.5, fill:{color:C.domoBlue} });
      s.addText(st.stage, { x:sx, y:sy, w:sw, h:0.28, fontFace:F, fontSize:9, bold:true, color:C.charcoal, align:"center", valign:"middle", margin:0 });
      s.addText(st.head,  { x:sx, y:sy+0.28, w:sw, h:0.22, fontFace:F, fontSize:9, color:C.charcoal, align:"center", margin:0 });
      // Mode chip
      const modeW = sw-0.4; const modeX = sx+0.2; const modeY = sy+0.62;
      s.addShape(pres.shapes.RECTANGLE, { x:modeX, y:modeY, w:modeW, h:0.28, fill:{color: dark?"1E2328":C.neutral2} });
      s.addText(st.sub,  { x:modeX, y:modeY, w:modeW/2, h:0.28, fontFace:F, fontSize:8, color:bc, align:"left",  valign:"middle", margin:4 });
      s.addText(st.mode, { x:modeX+modeW/2, y:modeY, w:modeW/2, h:0.28, fontFace:F, fontSize:8, bold:true, color:dark?C.domoBlue:C.charcoal, align:"right", valign:"middle", margin:4 });
      s.addText(st.body, { x:sx+0.18, y:sy+1.05, w:sw-0.36, h:2.35, fontFace:F, fontSize:10, color:bc, margin:0, lineSpacingMultiple:1.2 });
      if (i < 2) s.addImage({ data:icArrow, x:sx+sw+0.02, y:sy+sh/2-0.08, w:0.16, h:0.16 });
    });
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:4.88, w:8.9, h:0.26, fill:{color:C.neutral1} });
    s.addText("Probabilistic gets you to the answer. Deterministic keeps it.", {
      x:0.55, y:4.88, w:8.9, h:0.26, fontFace:F, fontSize:9.5, bold:true, italic:true, color:C.charcoal, align:"center", valign:"middle", margin:0,
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 8 — Lens 1 In Practice: Think Outside. Run Inside.
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 8);
    lensChip(s, "LENS 1");
    titleBar(s, pres, "Think Outside. Run Inside.", "Discovery happens in conversation. Production happens in governed tools, with the right mix per task.", true);
    const cw = 4.3, cy = 1.28, ch = 3.55, gap = 0.3;
    // Discovery column
    {
      const cx = 0.55;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:ch, fill:{color:C.neutral1}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:0.46, fill:{color:C.domoBlue} });
      s.addText("DISCOVERY  ·  THE THINKING", { x:cx, y:cy, w:cw, h:0.28, fontFace:F, fontSize:9, bold:true, color:C.charcoal, align:"center", valign:"middle", margin:0 });
      s.addText("Probabilistic. Conversational. Where the answer is still forming.", { x:cx, y:cy+0.28, w:cw, h:0.18, fontFace:F, fontSize:7.5, color:C.charcoal, align:"center", margin:0 });
      const tools = [
        { label:"Claude / Gemini / Cursor  via MCP", sub:"Open protocol. AI tools connect securely to your data." },
        { label:"Domo AI Chat",                      sub:"Grounded in your data." },
      ];
      let ty = cy+0.6;
      tools.forEach(t => {
        s.addShape(pres.shapes.RECTANGLE, { x:cx+0.2, y:ty, w:cw-0.4, h:0.72, fill:{color:C.white}, shadow:mkShadow() });
        s.addShape(pres.shapes.RECTANGLE, { x:cx+0.2, y:ty, w:0.06, h:0.72, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
        s.addText(t.label, { x:cx+0.35, y:ty+0.06, w:cw-0.6, h:0.3, fontFace:F, fontSize:10, bold:true, color:C.charcoal, margin:0 });
        s.addText(t.sub,   { x:cx+0.35, y:ty+0.38, w:cw-0.6, h:0.28, fontFace:F, fontSize:8.5, color:C.neutral4, margin:0 });
        ty += 0.82;
      });
    }
    // Arrow — centered in 0.30" gap (4.85 to 5.15), ends at 5.075 well clear of Production
    s.addImage({ data:icArrow, x:4.93, y:cy+ch/2-0.08, w:0.15, h:0.15 });
    s.addShape(pres.shapes.RECTANGLE, { x:4.50, y:cy+ch/2+0.06, w:1.00, h:0.22, fill:{color:C.white}, line:{color:C.neutral3, width:0.75} });
    s.addText("hardens into", { x:4.50, y:cy+ch/2+0.06, w:1.00, h:0.22, fontFace:F, fontSize:9, italic:true, color:C.neutral4, align:"center", valign:"middle", margin:0 });
    // Production column
    {
      const cx = 5.15;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:ch, fill:{color:C.charcoal}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:0.46, fill:{color:"2A2F36"} });
      s.addText("PRODUCTION  ·  THE DOING", { x:cx, y:cy, w:cw, h:0.28, fontFace:F, fontSize:9, bold:true, color:C.domoBlue, align:"center", valign:"middle", margin:0 });
      s.addText("Governed. Auditable. Runs every refresh.", { x:cx, y:cy+0.28, w:cw, h:0.18, fontFace:F, fontSize:7.5, color:C.neutral3, align:"center", margin:0 });
      const tools = [
        { label:"Magic ETL",  sub:"Python / SQL tiles for deterministic logic. AI Text Summarization tile where AI earns its place." },
        { label:"Workflows",  sub:"Agent tasks orchestrated as part of a governed process." },
      ];
      let ty = cy+0.6;
      tools.forEach(t => {
        s.addShape(pres.shapes.RECTANGLE, { x:cx+0.2, y:ty, w:cw-0.4, h:0.82, fill:{color:"2A2F36"}, shadow:mkShadow() });
        s.addShape(pres.shapes.RECTANGLE, { x:cx+0.2, y:ty, w:0.06, h:0.82, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
        s.addText(t.label, { x:cx+0.35, y:ty+0.06, w:cw-0.6, h:0.3, fontFace:F, fontSize:10, bold:true, color:C.white, margin:0 });
        s.addText(t.sub,   { x:cx+0.35, y:ty+0.38, w:cw-0.6, h:0.38, fontFace:F, fontSize:8.5, color:C.neutral3, margin:0, lineSpacingMultiple:1.1 });
        ty += 0.92;
      });
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 9 — Lens 2: Context Is the Moat
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 9);
    lensChip(s, "LENS 2");
    titleBar(s, pres, "Context Is the Moat", "The model is commodity. Context isn't.", true);
    // Left: key statements
    const lw = 5.1, rw = 3.55;
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:1.28, w:lw, h:3.55, fill:{color:C.charcoal}, shadow:mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:1.28, w:lw, h:0.06, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
    s.addText("Every competitor will have the same models in 18 months.", {
      x:0.75, y:1.42, w:lw-0.4, h:0.5, fontFace:F, fontSize:14, bold:true, color:C.white, margin:0, lineSpacingMultiple:0.95,
    });
    s.addText("The advantage is the data, dictionary, lineage, and judgment around them.", {
      x:0.75, y:1.98, w:lw-0.4, h:0.38, fontFace:F, fontSize:10.5, color:C.neutral3, margin:0, lineSpacingMultiple:1.1,
    });
    s.addShape(pres.shapes.LINE, { x:0.75, y:2.5, w:lw-0.5, h:0, line:{color:C.neutral4, width:0.5} });
    s.addText([
      { text:"A great model on bad data is worse than a decent model on great data.", options:{ italic:true, color:C.domoBlue } }
    ], { x:0.75, y:2.62, w:lw-0.4, h:0.45, fontFace:F, fontSize:11.5, margin:0, lineSpacingMultiple:1.0 });
    s.addText("Generic AI knows about your industry. Only your AI knows about your customers, your products, your processes, your codes, your audit history.", {
      x:0.75, y:3.18, w:lw-0.4, h:0.55, fontFace:F, fontSize:10, color:C.neutral3, margin:0, lineSpacingMultiple:1.15,
    });
    // Right: evidence panel
    const rx = 0.55+lw+0.3;
    s.addShape(pres.shapes.RECTANGLE, { x:rx, y:1.28, w:rw, h:3.55, fill:{color:C.neutral1}, shadow:mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:rx, y:1.28, w:rw, h:0.38, fill:{color:C.domoBlue} });
    s.addText("THE EVIDENCE", { x:rx, y:1.28, w:rw, h:0.38, fontFace:F, fontSize:10, bold:true, color:C.charcoal, align:"center", valign:"middle", margin:0 });
    s.addText("Better context beats better models.", { x:rx+0.18, y:1.74, w:rw-0.36, h:0.35, fontFace:F, fontSize:11.5, bold:true, color:C.charcoal, margin:0 });
    s.addText("Forecasting error, four ways to combine humans and AI:", { x:rx+0.18, y:2.14, w:rw-0.36, h:0.3, fontFace:F, fontSize:9, color:C.neutral4, margin:0 });
    // Stat callout
    s.addShape(pres.shapes.RECTANGLE, { x:rx+0.18, y:2.52, w:rw-0.36, h:0.85, fill:{color:C.white}, shadow:mkShadow() });
    s.addText("2.5×", { x:rx+0.18, y:2.58, w:rw-0.36, h:0.42, fontFace:F, fontSize:30, bold:true, color:C.charcoal, align:"center", margin:0 });
    s.addText("error drop with qualitative human context", { x:rx+0.18, y:3.02, w:rw-0.36, h:0.3, fontFace:F, fontSize:8.5, color:C.neutral4, align:"center", margin:0 });
    s.addText("Tong (WSB); Ibrahim, Kim & Tong, Management Science 67(4) (2021)", { x:rx+0.18, y:3.55, w:rw-0.36, h:0.45, fontFace:F, fontSize:7.5, italic:true, color:C.neutral4, margin:0, lineSpacingMultiple:1.1 });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 10 — Lens 2: Four Steps for Agentic Success
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 10);
    lensChip(s, "LENS 2");
    titleBar(s, pres, "The Four Steps for Agentic Success", "Instructions, LLM, Tools, and Knowledge, configured together in Domo.", true);
    // Left: 4-pillar list
    const pillars = [
      { icon:icRobot, label:"LLM"          },
      { icon:icClip,  label:"Instructions" },
      { icon:icDB,    label:"Knowledge"    },
      { icon:icBolt,  label:"Tools"        },
    ];
    pillarList(s, pres, 0.55, 1.32, 3.3, pillars, -1);
    // Right: screenshot of Domo Text Generation tile (6:3 ratio → w:5.5, h:2.75)
    s.addImage({ data:scr10, x:4.1, y:1.32, w:5.5, h:2.75, sizing:{type:"contain", w:5.5, h:2.75} });
    s.addText("Domo AI Agent builder: configuring Instructions, LLM, Knowledge, and Tools in one place.", {
      x:4.1, y:4.15, w:5.5, h:0.62, fontFace:F, fontSize:8.5, italic:true, color:C.neutral4, margin:0, lineSpacingMultiple:1.1,
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 11 — Lens 2: Adjusting Instructions
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 11);
    lensChip(s, "LENS 2");
    titleBar(s, pres, "Adjusting Instructions", "Changing the system prompt changes the output. Same data, completely different voice.", true);
    // Instruction quote chip
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:1.32, w:8.9, h:0.5, fill:{color:C.charcoal}, shadow:mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:1.32, w:0.06, h:0.5, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
    s.addText([
      { text:"Instruction: ", options:{bold:true, color:C.domoBlue} },
      { text:"\"You are a sports reporter... make it sound like something an announcer might say on the Sports Center program for live TV.\"", options:{italic:true, color:C.neutral3} },
    ], { x:0.72, y:1.32, w:8.55, h:0.5, fontFace:F, fontSize:9.5, valign:"middle", margin:0 });
    // Output screenshot (7:3 ratio → w:7.0, h:2.88); reduced to leave room for caption
    s.addImage({ data:scr11, x:1.5, y:1.9, w:7.0, h:2.88, sizing:{type:"contain", w:7.0, h:2.88} });
    s.addText("Actual agent output: two LinkedIn posts from the same LinkedIn dataset, shaped by the instruction.", {
      x:1.5, y:4.83, w:7.0, h:0.15, fontFace:F, fontSize:7.5, italic:true, color:C.neutral4, margin:0,
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 12 — Lens 2: Expanding Instructions (in Domo Workflows)
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 12);
    lensChip(s, "LENS 2");
    titleBar(s, pres, "Expanding Context in Domo Workflows", "Four pillars: Instructions, LLM, Knowledge, Tools, configured inside Domo Workflows.", true, 2);
    // Content starts at y:1.72 (lines=2 titleBar: accentY=1.30, sub ends 1.68, +0.04 gap)
    s.addImage({ data:scr12, x:0.55, y:1.72, w:3.55, h:3.22, sizing:{type:"contain", w:3.55, h:3.22} });
    const pillars12 = [
      { icon:icRobot, label:"LLM"          },
      { icon:icClip,  label:"Instructions" },
      { icon:icDB,    label:"Knowledge"    },
      { icon:icBolt,  label:"Tools"        },
    ];
    pillarList(s, pres, 4.3, 1.72, 5.15, pillars12, 1, 0.73, 0.06);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 13 — Lens 2: Knowledge (structured and unstructured)
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 13);
    lensChip(s, "LENS 2");
    titleBar(s, pres, "Knowledge (structured and unstructured)", "What the model can reference: DataSets, Document Collections, Files.", true, 2);
    // Content starts at y:1.72 (lines=2 titleBar)
    s.addImage({ data:scr13, x:0.55, y:1.72, w:3.55, h:3.22, sizing:{type:"contain", w:3.55, h:3.22} });
    const pillars13 = [
      { icon:icRobot, label:"LLM"          },
      { icon:icClip,  label:"Instructions" },
      { icon:icDB,    label:"Knowledge"    },
      { icon:icBolt,  label:"Tools"        },
    ];
    pillarList(s, pres, 4.3, 1.72, 5.15, pillars13, 2, 0.73, 0.06);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 14 — Lens 2: Tools (what the model can do)
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 14);
    lensChip(s, "LENS 2");
    s.addText("Tools: What the Model Can Do", { x:0.6, y:0.28, w:7.5, h:0.58, fontFace:F, fontSize:30, bold:true, color:C.charcoal, margin:0, lineSpacingMultiple:0.95 });
    s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:0.88, w:1.4, h:0.04, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
    s.addText("APIs, datasets, and actions the agent can invoke, inside your governed environment.", { x:0.6, y:0.98, w:7.5, h:0.28, fontFace:F, fontSize:10, color:C.neutral4, margin:0 });
    // Left: Tools tab screenshot (4:4 square → 3.6×3.6)
    s.addImage({ data:scr14, x:0.55, y:1.32, w:3.6, h:3.6, sizing:{type:"contain", w:3.6, h:3.6} });
    // Right: pillar list with Tools highlighted (index 3)
    const pillars14 = [
      { icon:icRobot, label:"LLM"          },
      { icon:icClip,  label:"Instructions" },
      { icon:icDB,    label:"Knowledge"    },
      { icon:icBolt,  label:"Tools"        },
    ];
    pillarList(s, pres, 4.3, 1.32, 5.15, pillars14, 3);
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 15 — Lens 2 In Practice: Global Manufacturer
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 15);
    lensChip(s, "LENS 2");
    titleBar(s, pres, "In Practice: A Global Manufacturer", "Identifying details removed. The model is off-the-shelf. The value isn't.", true);
    const panels = [
      { title:"THE WORKLOAD",       body:"AI agents read omnichannel engagement data and summarize patterns by product line and channel for the commercial team.",                                                                 icon:icInd },
      { title:"WHAT MAKES IT WORK", body:"Their own channel taxonomy. Their own product history. Their own market segmentation. This is what an established enterprise has that no foundation-model vendor will ever ship.",     icon:icLayer },
      { title:"THE LESSON",         body:"Decades of customer context, taxonomy, history, judgment. The moat is the organisation of that context, not the model that operates on it.",                                             icon:icDB, dark:true },
    ];
    const pw = 2.85, py = 1.28, ph = 3.55, gap = 0.2;
    panels.forEach((p, i) => {
      const px = 0.55 + i*(pw+gap);
      card(s, pres, px, py, pw, ph, { icon:p.icon, title:p.title, body:p.body, dark:p.dark, fullWidthBody:true });
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 16 — Lens 3: Governance & Observability
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 16);
    lensChip(s, "LENS 3");
    titleBar(s, pres, "Governance & Observability", "Treat your bots like humans. Performance review. Audit trail. Ownership.", true);
    const dims = [
      { label:"WHO",  body:"Which users are prompting? Which roles? Which agents are acting on whose behalf?",     icon:icUsers },
      { label:"WHAT", body:"Which datasets did the model touch? Which prompts? Which responses?",                   icon:icEye },
      { label:"COST", body:"Which models ran? How often? At what tier? With what failure rate?",                    icon:icCoins },
    ];
    const dw = 2.85, dy = 1.28, dh = 2.1, gap = 0.2;
    dims.forEach((d, i) => {
      const dx = 0.55+i*(dw+gap);
      s.addShape(pres.shapes.RECTANGLE, { x:dx, y:dy, w:dw, h:dh, fill:{color:C.neutral1}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:dx, y:dy, w:dw, h:0.42, fill:{color:C.domoBlue} });
      s.addText(d.label, { x:dx, y:dy, w:dw, h:0.42, fontFace:F, fontSize:13, bold:true, color:C.charcoal, align:"center", valign:"middle", margin:0 });
      // circY centers content (icon 0.46 + gap 0.14 + body 0.80 = 1.40) in available (dh 2.10 - header 0.42 = 1.68): padding=(1.68-1.40)/2=0.14
      const cs = 0.46; const cx = dx+(dw-cs)/2; const circY = dy+0.56;
      s.addShape(pres.shapes.OVAL, { x:cx, y:circY, w:cs, h:cs, fill:{color:C.charcoal} });
      s.addImage({ data:d.icon, x:cx+(cs-0.3)/2, y:circY+(cs-0.3)/2, w:0.3, h:0.3 });
      s.addText(d.body, { x:dx+0.18, y:circY+0.60, w:dw-0.36, h:0.80, fontFace:F, fontSize:9.5, color:C.neutral4, align:"center", margin:0, lineSpacingMultiple:1.15 });
    });
    // Quote
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:3.55, w:8.9, h:1.28, fill:{color:C.charcoal}, shadow:mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:3.55, w:0.07, h:1.28, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
    s.addText([
      { text:"\"An AI-first firm runs on ", options:{color:C.neutral3} },
      { text:"a scalable decision factory",   options:{color:C.domoBlue, bold:true} },
      { text:", and like any factory, that means data pipelines, experimentation, software infrastructure, and clear, multidisciplinary governance.\"", options:{color:C.neutral3} },
    ], { x:0.78, y:3.65, w:8.5, h:0.72, fontFace:F, fontSize:10, margin:0, lineSpacingMultiple:1.2 });
    s.addText("M. Iansiti & K. Lakhani, Harvard Business School, \"Competing in the Age of AI\"", {
      x:0.78, y:4.38, w:8.5, h:0.28, fontFace:F, fontSize:8, italic:true, color:C.neutral4, margin:0,
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 17 — Lens 3: Governance Primitives
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 17);
    lensChip(s, "LENS 3");
    titleBar(s, pres, "Governance Isn't a Posture. It's Plumbing.", "The same controls that govern people should govern AI. In Domo, they are the same controls.", true, 2);
    // Content at y:1.72 (lines=2). Cards reduced to ph:1.42, gap:0.18 so 2 rows end at 4.74; quote fits at y:4.78.
    const prims = [
      { title:"Row-Level Security", body:"An AI agent sees only the rows that user is allowed to see. PDP rules that protect dashboards protect AI responses.",          icon:icKey    },
      { title:"Column Masking",     body:"PII and sensitive fields masked at the column level. AI agents inherit the mask. They cannot reveal what the user couldn't.", icon:icShield },
      { title:"OAuth / SSO",        body:"Identity flows through to every AI interaction. No service accounts. No shared credentials. Every action is attributable.",    icon:icUserSh },
      { title:"Audit Trail",        body:"Every prompt, response, dataset access, and model invocation is logged. The same audit posture extends to AI.",               icon:icClip   },
    ];
    const pw = 4.2, ph = 1.42, gap = 0.18;
    prims.forEach((p, i) => {
      const col = i%2, row = Math.floor(i/2);
      const px = 0.55+col*(pw+gap); const py = 1.72+row*(ph+gap);
      card(s, pres, px, py, pw, ph, { icon:p.icon, title:p.title, body:p.body });
    });
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:4.78, w:8.9, h:0.22, fill:{color:C.neutral1} });
    s.addText("If your governance model already works for your dashboards, it already works for your AI.", {
      x:0.75, y:4.78, w:8.5, h:0.22, fontFace:F, fontSize:10, bold:true, italic:true, color:C.charcoal, valign:"middle", margin:0,
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 18 — Lens 3: What It Looks Like (logging)
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 18);
    lensChip(s, "LENS 3");
    titleBar(s, pres, "Every Request. Every Dataset. Logged.", "Domo Stats AI Activity Log: every prompt, dataset access, model, and status. Governance built in.", true, 2);
    // Content at y:1.72 (lines=2). Screenshot h:2.90 so it ends at 4.62, chips at y:4.70 unchanged.
    s.addImage({ data:scr18, x:0.55, y:1.72, w:8.9, h:2.90, sizing:{type:"contain", w:8.9, h:2.90} });
    // Field chips below — fy:4.70+fh:0.24 = 4.94, clear of footer at 5.15
    const fields = ["Who","When","Prompt","Response","Datasets touched","Model","Status"];
    const fw = 1.18, fh = 0.24, fy = 4.70;
    const totalW = fields.length*fw + (fields.length-1)*0.08;
    const fx0 = (10-totalW)/2;
    fields.forEach((f, i) => {
      const fx = fx0 + i*(fw+0.08);
      s.addShape(pres.shapes.RECTANGLE, { x:fx, y:fy, w:fw, h:fh, fill:{color:C.neutral1} });
      s.addText(f, { x:fx, y:fy, w:fw, h:fh, fontFace:F, fontSize:7.5, bold:true, color:C.neutral4, align:"center", valign:"middle", margin:0 });
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 19 — Lens 4: Security & Compliance
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 19);
    lensChip(s, "LENS 4");
    titleBar(s, pres, "Security & Compliance", "When the data can't go to the model. In regulated work, this is the first filter, not an afterthought.", true);
    const cw = 4.3, cy = 1.28, ch = 3.55, gap = 0.3;
    // Constraint
    {
      const cx = 0.55;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:ch, fill:{color:C.neutral1}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:0.42, fill:{color:C.domoBlue} });
      s.addText("THE CONSTRAINT", { x:cx, y:cy, w:cw, h:0.42, fontFace:F, fontSize:11, bold:true, color:C.charcoal, align:"center", valign:"middle", margin:0 });
      s.addText("Regulated data, contracts, audit trails: non-negotiable.", { x:cx+0.18, y:cy+0.55, w:cw-0.36, h:0.42, fontFace:F, fontSize:11, bold:true, color:C.charcoal, margin:0, lineSpacingMultiple:0.95 });
      s.addText("A mandated, centrally-governed AI platform is a constraint working as designed. But governed-by-default doesn't mean the same model is right for every use case.", {
        x:cx+0.18, y:cy+1.1, w:cw-0.36, h:0.85, fontFace:F, fontSize:10, color:C.neutral4, margin:0, lineSpacingMultiple:1.15,
      });
    }
    // Answer
    {
      const cx = 0.55+cw+gap;
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:ch, fill:{color:C.charcoal}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:cw, h:0.42, fill:{color:"2A2F36"} });
      s.addText("THE ANSWER", { x:cx, y:cy, w:cw, h:0.42, fontFace:F, fontSize:11, bold:true, color:C.domoBlue, align:"center", valign:"middle", margin:0 });
      s.addText("Bring Your Own Model", { x:cx+0.18, y:cy+0.55, w:cw-0.36, h:0.42, fontFace:F, fontSize:14, bold:true, color:C.white, margin:0 });
      const items = [
        "Inference inside your cloud.",
        "Your governance, your audit trail.",
        "Match the model to the job, not the other way around.",
        "BYO Model = no token-credit charges on customer-provided models.",
      ];
      let iy = cy+1.08;
      items.forEach(it => {
        const rh = 0.42; const cs = 0.24; const oy = iy+(rh-cs)/2;
        s.addShape(pres.shapes.OVAL, { x:cx+0.18, y:oy, w:cs, h:cs, fill:{color:C.domoBlue} });
        s.addText("✓", { x:cx+0.18, y:oy, w:cs, h:cs, fontFace:F, fontSize:9, bold:true, color:C.charcoal, align:"center", valign:"middle", margin:0 });
        s.addText(it,  { x:cx+0.5,  y:iy, w:cw-0.65, h:rh, fontFace:F, fontSize:9.5, color:C.neutral3, valign:"middle", margin:0, lineSpacingMultiple:1.0 });
        iy += rh+0.12;
      });
    }
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 20 — Lens 4 In Practice: Financial Services Workflow
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 20);
    lensChip(s, "LENS 4");
    titleBar(s, pres, "In Practice: A Real Production AI Workflow", "A regulated financial-services firm. Identifying details removed.", true, 2);
    // Content at y:1.72 (lines=2). Steps sh:1.80 ends at 3.52; WHY block starts at 3.60.
    const steps = [
      { label:"Claims & Risk Data",   sub:"Core systems of record",       icon:icDB },
      { label:"Governed Boundary",    sub:"Inference inside the firm's cloud", icon:icShield },
      { label:"Policy-Aware AI Agent",sub:"Pattern detection + guidance match",icon:icRobot },
      { label:"Action",               sub:"Recommendations to the operations team",icon:icUsers },
    ];
    const sw = 1.95, sy = 1.72, sh = 1.80, gap = 0.32;
    steps.forEach((st, i) => {
      const sx = 0.55 + i*(sw+gap);
      const dark = i===1||i===2;
      const bg = dark ? C.charcoal : C.neutral1;
      const tc = dark ? C.white    : C.charcoal;
      const bc = dark ? C.neutral3 : C.neutral4;
      s.addShape(pres.shapes.RECTANGLE, { x:sx, y:sy, w:sw, h:sh, fill:{color:bg}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:sx, y:sy, w:sw, h:0.06, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
      const cs = 0.46; const circX = sx+(sw-cs)/2; const circY = sy+0.18;
      const circFill = dark ? C.neutral1 : C.charcoal;
      s.addShape(pres.shapes.OVAL, { x:circX, y:circY, w:cs, h:cs, fill:{color:circFill} });
      s.addImage({ data:st.icon, x:circX+(cs-0.3)/2, y:circY+(cs-0.3)/2, w:0.3, h:0.3 });
      s.addText(st.label, { x:sx+0.1, y:circY+0.55, w:sw-0.2, h:0.48, fontFace:F, fontSize:10, bold:true, color:tc, align:"center", margin:0, lineSpacingMultiple:0.95 });
      s.addText(st.sub,   { x:sx+0.1, y:circY+1.08, w:sw-0.2, h:0.45, fontFace:F, fontSize:8.5, color:bc, align:"center", margin:0, lineSpacingMultiple:1.1 });
      if (i < 3) s.addImage({ data:icArrow, x:sx+sw+0.06, y:sy+sh/2-0.08, w:0.2, h:0.2 });
    });
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:3.60, w:8.9, h:1.20, fill:{color:C.charcoal}, shadow:mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:3.60, w:0.07, h:1.20, fill:{color:C.accentOrange}, line:{color:C.accentOrange} });
    s.addText("WHY THIS PASSES LENS 4", { x:0.78, y:3.67, w:8.5, h:0.28, fontFace:F, fontSize:9, bold:true, color:C.accentOrange, margin:0 });
    s.addText("The sensitive data never crosses a trust boundary the firm doesn't already own. The AI does real, high-value work inside the same governance envelope that protects regulated data today. That's the test.", {
      x:0.78, y:3.99, w:8.5, h:0.75, fontFace:F, fontSize:10, color:C.neutral3, margin:0, lineSpacingMultiple:1.2,
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 21 — Lens 5: Token Economics (32×)
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 21);
    lensChip(s, "LENS 5");
    titleBar(s, pres, "Tokens Are the New Compute", "The price spread between the cheapest and most expensive model on Domo's published rate card is 32×.", true);
    // Big stat
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:1.28, w:3.8, h:3.5, fill:{color:C.charcoal}, shadow:mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:1.28, w:3.8, h:0.06, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
    s.addText("32×", { x:0.55, y:1.8, w:3.8, h:1.4, fontFace:F, fontSize:80, bold:true, color:C.domoBlue, align:"center", valign:"middle", margin:0 });
    s.addText("spread between the cheapest and most expensive models on the Domo rate card", {
      x:0.75, y:3.35, w:3.4, h:0.9, fontFace:F, fontSize:11, color:C.neutral3, align:"center", margin:0, lineSpacingMultiple:1.2,
    });
    // Right panel
    s.addShape(pres.shapes.RECTANGLE, { x:4.65, y:1.28, w:4.8, h:3.5, fill:{color:C.neutral1}, shadow:mkShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:4.65, y:1.28, w:4.8, h:0.06, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
    s.addText("Representative sample: 4 of 90 published model SKUs", {
      x:4.85, y:1.42, w:4.4, h:0.3, fontFace:F, fontSize:8.5, color:C.neutral4, margin:0,
    });
    const rows = [
      { model:"Small / Lightweight",    price:"$0.08",  pct:8 },
      { model:"Mid-tier",               price:"$0.50",  pct:28 },
      { model:"Performance",            price:"$1.20",  pct:60 },
      { model:"Frontier / Premium",     price:"$2.56",  pct:100 },
    ];
    let ry = 1.82;
    rows.forEach(r => {
      s.addText(r.model, { x:4.85, y:ry, w:2.1, h:0.35, fontFace:F, fontSize:9.5, color:C.charcoal, valign:"middle", margin:0 });
      s.addText(r.price, { x:6.98, y:ry, w:0.7, h:0.35, fontFace:F, fontSize:9.5, bold:true, color:C.charcoal, align:"right", valign:"middle", margin:0 });
      // bx:7.75 + maxBarW:1.60 = 9.35 — stays within 9.5" right margin
      const bw = (r.pct/100)*1.60; const bx = 7.75; const bh = 0.22; const by = ry+(0.35-bh)/2;
      s.addShape(pres.shapes.RECTANGLE, { x:bx, y:by, w:1.60, h:bh, fill:{color:C.neutral2} });
      s.addShape(pres.shapes.RECTANGLE, { x:bx, y:by, w:bw,  h:bh, fill:{color:r.pct===100?C.charcoal:C.domoBlue} });
      ry += 0.5;
    });
    s.addText("USD per 1M tokens  ·  source: Domo rate card  ·  figures illustrative", {
      x:4.85, y:4.4, w:4.4, h:0.25, fontFace:F, fontSize:7.5, italic:true, color:C.neutral4, margin:0,
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 22 — Lens 5: Three Decisions
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 22);
    lensChip(s, "LENS 5");
    titleBar(s, pres, "Three Questions Before You Ship", "Each one is a knob. Set it wrong and the use case fails the lens.", true);
    const decisions = [
      { n:1, label:"FREQUENCY", title:"How often does this need to run?",    body:"Every record? Every refresh? On demand? If the answer is \"every refresh,\" lock it down (Lens 1).", icon:icBolt },
      { n:2, label:"MODEL",     title:"Which tier is right for the job?",    body:"Routine classification rarely needs the most expensive model. Reserve premium tiers for novel reasoning and sensitive output.", icon:icRobot },
      { n:3, label:"TRIGGER",   title:"Human-initiated or automatic?",       body:"What's the failure cost if it fires when it shouldn't? The higher the cost, the more the trigger belongs in human hands.", icon:icEye },
    ];
    const dw = 2.85, dy = 1.28, dh = 3.55, gap = 0.2;
    decisions.forEach((d, i) => {
      const dx = 0.55 + i*(dw+gap);
      s.addShape(pres.shapes.RECTANGLE, { x:dx, y:dy, w:dw, h:dh, fill:{color:C.neutral1}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:dx, y:dy, w:dw, h:0.5, fill:{color:C.charcoal} });
      s.addText(d.label, { x:dx, y:dy, w:dw, h:0.28, fontFace:F, fontSize:9, bold:true, color:C.domoBlue, align:"center", valign:"middle", margin:0, charSpacing:1 });
      const cs = 0.36; const oy = dy+0.22;
      numberedCircle(s, pres, d.n, dx+(dw-cs)/2, oy, cs, C.domoBlue, C.charcoal, 13);
      // circY centers content (icon 0.52 + gap 0.13 + title 0.55 + gap 0.08 + body 0.85 = 2.13)
      // in available (dh 3.55 - header 0.50 = 3.05): padding = (3.05-2.13)/2 = 0.46
      const cs2 = 0.52; const circX = dx+(dw-cs2)/2; const circY = dy+0.96;
      s.addShape(pres.shapes.OVAL, { x:circX, y:circY, w:cs2, h:cs2, fill:{color:C.charcoal} });
      s.addImage({ data:d.icon, x:circX+(cs2-0.34)/2, y:circY+(cs2-0.34)/2, w:0.34, h:0.34 });
      s.addText(d.title, { x:dx+0.18, y:circY+0.65, w:dw-0.36, h:0.55, fontFace:F, fontSize:11, bold:true, color:C.charcoal, align:"center", margin:0, lineSpacingMultiple:0.95 });
      s.addText(d.body,  { x:dx+0.18, y:circY+1.28, w:dw-0.36, h:0.85, fontFace:F, fontSize:10, color:C.neutral4, align:"center", margin:0, lineSpacingMultiple:1.15 });
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 23 — The Framework Summary
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 23);
    titleBar(s, pres, "The Framework", "A checklist, not a rulebook. If a use case passes all five lenses, ship it. If it fails one, fix that first, or pick a different use case.");
    const items = [
      { n:1, icon:icLock,   title:"Lock It Down",              q:"Is this discovery, or is this production?" },
      { n:2, icon:icLayer,  title:"Context Is the Moat",       q:"Is the AI grounded in your context?" },
      { n:3, icon:icEye,    title:"Governance & Observability", q:"Can you see what AI is doing?" },
      { n:4, icon:icShield, title:"Security & Compliance",      q:"Can the data even go to the model?" },
      { n:5, icon:icCoins,  title:"Token Economics",            q:"Frequency, model, and trigger: chosen on purpose?" },
    ];
    const ih = 0.64, iy0 = 1.35, gap = 0.14;
    items.forEach((it, i) => {
      const iy = iy0 + i*(ih+gap);
      s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:iy, w:8.9, h:ih, fill:{color:C.neutral1}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:iy, w:0.07, h:ih, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
      const cs = 0.4; const oy = iy+(ih-cs)/2;
      numberedCircle(s, pres, it.n, 0.72, oy, cs, C.domoBlue, C.charcoal, 13);
      const cs2 = 0.4; const oy2 = iy+(ih-cs2)/2;
      s.addShape(pres.shapes.OVAL, { x:1.24, y:oy2, w:cs2, h:cs2, fill:{color:C.charcoal} });
      s.addImage({ data:it.icon, x:1.24+(cs2-0.26)/2, y:oy2+(cs2-0.26)/2, w:0.26, h:0.26 });
      s.addText(it.title, { x:1.78, y:iy, w:3.0, h:ih, fontFace:F, fontSize:12, bold:true, color:C.charcoal, valign:"middle", margin:0 });
      s.addText(it.q,     { x:4.88, y:iy, w:4.42, h:ih, fontFace:F, fontSize:10.5, color:C.neutral4, valign:"middle", margin:0 });
    });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 24 — Closing Quote
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { data:bgGradient };
    s.addImage({ data:logoPrimary, x:0.0, y:0.0, w:1.0, h:1.0 });
    s.addText('"Can" is the easy question.', {
      x:0.65, y:1.6, w:8.5, h:1.0, fontFace:F, fontSize:42, bold:true, color:C.charcoal, margin:0, lineSpacingMultiple:0.92,
    });
    s.addText([
      { text:'"Should"', options:{ bold:true, color:C.accentOrange } },
      { text:" is the job.", options:{ bold:true, color:C.charcoal } },
    ], { x:0.65, y:2.58, w:8.5, h:1.0, fontFace:F, fontSize:42, margin:0, lineSpacingMultiple:0.92 });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 25 — Thank You / Questions
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { data:bgGradient };
    s.addImage({ data:logoPrimary, x:0.0, y:0.0, w:1.0, h:1.0 });
    s.addText("Let's start the", { x:0.65, y:1.5, w:8.5, h:0.65, fontFace:F, fontSize:36, bold:true, color:C.charcoal, margin:0 });
    s.addText([
      { text:'"should AI"', options:{ bold:true, color:C.accentOrange } },
      { text:" conversation.", options:{ bold:true, color:C.charcoal } },
    ], { x:0.65, y:2.12, w:8.5, h:0.75, fontFace:F, fontSize:36, margin:0 });
    s.addShape(pres.shapes.RECTANGLE, { x:0.65, y:2.95, w:2.0, h:0.04, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
    s.addText("Thank you. Questions?", { x:0.65, y:3.12, w:8.5, h:0.42, fontFace:F, fontSize:16, color:C.neutral4, margin:0 });
    s.addText("Ben Schein  ·  Chief AI & Analytics Officer  ·  Domo", { x:0.65, y:4.6, w:8.5, h:0.35, fontFace:F, fontSize:11, color:C.charcoal, margin:0 });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 26 — Appendix Divider
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { data:bgGradient };
    s.addImage({ data:logoPrimary, x:0.0, y:0.0, w:1.0, h:1.0 });
    s.addText("APPENDIX", { x:0.65, y:1.4, w:8.5, h:0.42, fontFace:F, fontSize:13, bold:true, color:C.neutral4, charSpacing:2, margin:0 });
    s.addText("AI Solutions in the Wild", { x:0.65, y:1.88, w:8.5, h:0.9, fontFace:F, fontSize:42, bold:true, color:C.charcoal, margin:0, lineSpacingMultiple:0.92 });
    s.addShape(pres.shapes.RECTANGLE, { x:0.65, y:2.82, w:2.2, h:0.04, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
    s.addText("Five anonymized production patterns, same five lenses, real workloads.", { x:0.65, y:2.96, w:8.5, h:0.38, fontFace:F, fontSize:13, color:C.neutral4, margin:0 });
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 27 — Appendix: Five Cases
  // ═══════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color:C.white };
    addFooter(s, pres, 27);
    titleBar(s, pres, "AI Solutions in the Wild", "Five anonymized production patterns from real customers. Details removed, figures rounded.");
    const cases = [
      { n:1, sector:"Retail / Telecom",    body:"Batch classification at scale: ~33 billion tokens a month through one solution." },
      { n:2, sector:"Manufacturing",       body:"A right-sized model mix: 25 solutions across small, medium, and interactive tiers." },
      { n:3, sector:"Insurance",           body:"Tiered escalation: routine work on a small model, a premium model for hard cases." },
      { n:4, sector:"Financial Services",  body:"Org-wide adoption: 44 distinct AI solutions in production across the business." },
      { n:5, sector:"Sports & Leisure",    body:"One focused use case: a single solution on an off-the-shelf model." },
    ];
    // ch:0.62 + gap:0.12 → last row ends at 1.35+4×0.74+0.62 = 4.93, clear of footer
    const ch = 0.62, cy0 = 1.35, gap = 0.12;
    cases.forEach((c, i) => {
      const cy = cy0 + i*(ch+gap);
      s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:cy, w:8.9, h:ch, fill:{color:C.neutral1}, shadow:mkShadow() });
      s.addShape(pres.shapes.RECTANGLE, { x:0.55, y:cy, w:0.07, h:ch, fill:{color:C.domoBlue}, line:{color:C.domoBlue} });
      const cs = 0.38; const oy = cy+(ch-cs)/2;
      numberedCircle(s, pres, c.n, 0.72, oy, cs, C.domoBlue, C.charcoal, 12);
      s.addText(c.sector, { x:1.24, y:cy, w:2.5, h:ch, fontFace:F, fontSize:11.5, bold:true, color:C.charcoal, valign:"middle", margin:0 });
      s.addText(c.body,   { x:3.85, y:cy, w:5.45, h:ch, fontFace:F, fontSize:10, color:C.neutral4, valign:"middle", margin:0 });
    });
  }

  // ── Write ─────────────────────────────────────────────────────────────────
  const OUT = "/Users/arminas.juknevicius/Downloads/AI_Can_vs_Should_Domo.pptx";
  await pres.writeFile({ fileName: OUT });
  console.log("Written:", OUT);
})();
