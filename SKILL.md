---
name: domo-branded-pptx
description: "Create Domo-branded PowerPoint presentations with official Domo styling, colors, gradient backgrounds, and logo placement. Use this skill whenever the user asks to create a presentation, deck, or slides for Domo, about Domo products, or wants Domo branding applied. Also triggers when the user uploads a .pptx and asks to rebrand it in Domo style, or provides a brief/document/notes to turn into a Domo-branded deck. Keywords: Domo slides, Domo deck, Domo presentation, Domo branding, Domo template, Domo PowerPoint, branded slides. Even if the user just says 'make me a deck' or 'create a presentation' and they've used this skill before or Domo context is present, use this skill."
---

# Domo Branded PowerPoint Skill

Creates professional, on-brand Domo presentations from scratch, from a brief, or by rebranding existing decks.

## Prerequisites

**Always read the base pptx skill first:**
```
view /mnt/skills/public/pptx/SKILL.md
```
Then read:
```
view /mnt/skills/public/pptx/pptxgenjs.md
```
These contain the core PptxGenJS API, QA process, and image conversion steps. This skill **overrides all design/color/layout decisions** from the base skill with Domo branding below.

## CRITICAL: Slide Layout — Read This First

**Always use `LAYOUT_16x9`** — this is the only layout whose coordinate system matches every position reference in this skill.

```javascript
const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9";   // 10" wide × 5.625" tall (16:9 widescreen)
// NEVER use "LAYOUT_WIDE" (13.33"×7.5") — content will be crammed into 75% of the slide
// NEVER use "LAYOUT_4x3" (10"×7.5") — footer positions won't match
```

**Why this matters:** PptxGenJS's `LAYOUT_WIDE` is 13.33" × 7.5". Every coordinate in this skill (footer at y=5.15, page number at x=9.2, content at w=9.0, etc.) assumes a **10"-wide, 5.625"-tall** canvas. Using any other layout makes all content compress into one corner with empty space filling the rest of the slide.

**Coordinate bounds to always respect:**

| Zone | Y range | Notes |
|------|---------|-------|
| Title | y=0.2 – 0.9 | h≤0.7 for single-line, h≤0.9 for two-line |
| Content | y=1.0 – 5.0 | must END before 5.0 — never let content touch the footer |
| Footer separator | y=5.15 | thin line across full width |
| Footer elements | y=5.18 – 5.55 | logo icon, CONFIDENTIAL label, page number |
| Slide edge | y=5.625 | hard bottom — nothing beyond this |
| Left margin | x=0.5 | consistent across all slides |
| Right margin | x+w ≤ 9.5 | content never past 9.5" |

**Verify content bounds before generating.** For every block at (x, y, w, h): confirm y+h ≤ 5.0 (or ≤ 5.55 for footer elements only). Platform notes, caveat boxes, and vision banners placed at y≥4.8 are the most common overflow points — calculate them explicitly.

## Brand Assets

Assets are bundled in the `assets/` directory relative to this skill file. Load them at the top of your generation script:

```javascript
// Get the skill's asset directory path
const SKILL_DIR = "<path-to-this-skill>/assets";

// Load as base64 for embedding in slides
function img64(filePath) {
  const buf = require("fs").readFileSync(filePath);
  const ext = require("path").extname(filePath).slice(1);
  return `image/${ext};base64,` + buf.toString("base64");
}

const bgGradient = img64(SKILL_DIR + "/domo_gradient_bg.png");  // Soft Domo-blue gradient
const logoPrimary = img64(SKILL_DIR + "/domo_logo_blue_800.png"); // PRIMARY logo: Domo-blue box + white DOMO wordmark (preferred)
const logoIcon   = img64(SKILL_DIR + "/domo_icon_400.png");      // Small DOMO square badge
```

Replace `<path-to-this-skill>` with the actual resolved path to this skill's directory.

## Logo Usage (Official Brand Guide)

There are two logo lockups:
- **Primary** — the DOMO wordmark inside a box. Preferred forms: **Domo-blue box + white text** (use whenever possible); also black box / white box with Domo-blue text. `domo_logo_blue_800.png` is this primary blue-box lockup.
- **Secondary / text lockup** — the bare DOMO wordmark (blue, black, or white). Use **sparingly**, only where vertical space is limited but horizontal extension is allowed (booths, billboards, video bumpers).

**Dos:**
- Use the **primary** logo whenever possible; prefer the blue box + white text.
- Use simple, neutral backgrounds with **high contrast** — the primary logo reads best on very light or very dark backgrounds.
- If the Domo-blue logo doesn't work on a background, use the **black or white** version.
- Give the logo **plenty of breathing space**; keep text well clear of it.
- When placing over an image, ensure enough contrast to clearly read the DOMO letters.

**Don'ts:**
- Don't modify the logo in any way — no stretching, recoloring, gradients, or patterns.
- Don't use any shade of blue that is similar to but not Domo Blue.
- Don't place text too close to the logo.
- Don't cut, trim, or block the logo with other images or icons.

## Domo Color Palette

Source of truth: the **official Domo Brand Guide**. The palette is *anchored by Domo Blue* and complemented by gray neutrals and a set of accents. The accents are to be used **sparingly and must never overpower Domo Blue**.

```javascript
const C = {
  // ── PRIMARY ──────────────────────────────────────────────
  // Domo Blue — Pantone 291 C · CMYK 37,6,1,0 · RGB 152,204,238
  // The signature colour. Should be the overall takeaway on every deck.
  domoBlue:  "99CCEE",  // fills/accents/logo box only — NOT a text colour (too light on white)

  // ── NEUTRALS (supportive gray hues) ──────────────────────
  white:     "FFFFFF",  // content-slide backgrounds, cards, logo text
  neutral1:  "F1F6FA",  // lightest gray — panel/section/card backgrounds  (alias: appBg, lightCard)
  neutral2:  "DCE4EA",  // light border — subtle dividers, hover fills      (alias: lineLight)
  neutral3:  "B7C1CB",  // mid border — separators, table/swatch outlines   (alias: lineGray)
  neutral4:  "68737F",  // secondary text / captions / footer               (alias: bodyGray, mutedGray)
  charcoal:  "3F454D",  // darkest neutral — primary heading/body text, dark panels & headers

  // ── ACCENTS (use sparingly — never overpower Domo Blue) ───
  accentOrange:  "FF9922",
  accentPurple:  "776CB0",
  accentMagenta: "C179BD",
  accentPink:    "ECA4DD",
  accentGreen:   "ADD4C1",

  // ── Functional status system (derived from the accents) ──
  // For status badges / KPI stat values / chips only. Pair each with its tinted bg + darker text.
  onTrack:   "ADD4C1",  // green  — On Track  (bg "E8F3EC", text "4A7A5A")
  atRisk:    "FF9922",  // orange — At Risk   (bg "FFF0DD", text "C47A10")
  behind:    "776CB0",  // purple — Behind    (bg "EAE7F3", text "5A5094")
  complete:  "99CCEE",  // blue   — Complete  (bg "E0F0FA", text "5A9ABE")
  warnBg:    "FFF0DD",  // tip/warning callout background (orange-accent tint)
};
```

> **Removed to stay on-brand:** earlier versions used `navy` (#1B2A3D) and `medBlue` (#5CACEE). Neither is in the official guide, and the guide explicitly says *"Do not use any other shade of blue."* Dark panels/headers now use `charcoal` (#3F454D — the darkest official neutral); all generated icons are `domoBlue` (#99CCEE) on every slide.

**Color rules:**
- **Domo Blue must dominate** as the brand takeaway; never substitute a similar-but-different blue.
- NEVER use Domo Blue (#99CCEE) as a text colour — it's too light on white. Use it for fills, shapes, accents, and the logo box only.
- Heading text is always `charcoal` (#3F454D) on light slides, `white` on dark slides.
- Body text is `neutral4` (#68737F) for secondary content, `charcoal` for primary.
- **All generated icons are Domo Blue (#99CCEE)** — every react-icons glyph is rendered in `domoBlue`, on both light and dark slides. The circle/tile behind the icon must be chosen so the domoBlue icon is actually visible:
  - **Light background** → seat icon in a **charcoal** circle. Never a domoBlue circle — domoBlue icon on domoBlue fill is invisible.
  - **Dark background (charcoal card/panel)** → seat icon in a **neutral1** (#F1F6FA) circle. DomoBlue icon on charcoal is acceptable without a circle, but a neutral1 chip adds polish.
  - **Never** use a domoBlue-filled circle for a domoBlue icon, regardless of the slide background — the icon will disappear.
- Dark panels / table headers / comparison headers / dark checklist slides use `charcoal` (#3F454D).
- Accents are decorative/functional only and used sparingly — `domoBlue` remains the primary accent.

## Typography

The official Domo typeface is **Open Sans** (weights: Light · Regular · Semibold · Bold · Extrabold). Map weights per the brand guide:
- **Open Sans Bold** → titles
- **Open Sans Light** → subtitles
- **Open Sans Regular** → body copy

```javascript
const FONT_H = "Open Sans";  // Headings (Bold)
const FONT_B = "Open Sans";  // Body (Regular) / subtitles (Light)
// Fallback for environments without Open Sans installed: "Calibri".
```

| Element          | Size    | Weight             | Color      |
|------------------|---------|--------------------|------------|
| Slide title      | 30pt    | Bold               | charcoal   |
| Section header   | 14–16pt | Bold/Semibold      | charcoal   |
| Subtitle         | 20pt    | Light              | neutral4   |
| Body text        | 10–12pt | Regular            | neutral4   |
| Card title       | 11–13pt | Bold/Semibold      | charcoal   |
| Card body        | 9.5–11pt| Regular            | neutral4   |
| Footer text      | 8pt     | Regular            | neutral4   |
| Page number      | 9pt     | Bold               | charcoal   |

## Slide Types & Layout Rules

### 1. Title / Section Divider Slides

Used for: opening slide, section transitions (e.g. between Workbench and Magic ETL sections), closing/thank-you slide.

```javascript
slide.background = { data: bgGradient };  // Always use the gradient background image
slide.addImage({ data: logoPrimary, x: 0.0, y: -0.1, w: 1.2, h: 1.2 }); // Primary logo top-left
slide.addText("Slide Title", {
  x: 0.7, y: 2.2, w: 7.0, h: 1.2,
  fontFace: "Open Sans", fontSize: 42, bold: true,
  color: "3F454D", margin: 0,
  lineSpacingMultiple: 0.95,   // tight leading — avoid the double-spaced look on wrapped titles
});
slide.addText("Subtitle text here", {
  x: 0.7, y: 3.3, w: 7.0, h: 0.6,
  fontFace: "Open Sans", fontSize: 20,   // Open Sans Light feel
  color: "68737F", margin: 0, lineSpacingMultiple: 1.0,
});
```

**Rules:**
- Gradient background image (from assets), NOT a solid color
- Primary Domo logo top-left: **always `x: 0.0, y: 0.0`** — never negative y. Negative y clips the logo outside the slide boundary.
- Logo size: `w: 1.0, h: 1.0` (keep w === h to avoid squishing). Logo bottom = y + h = **1.0"**.
- **Title y must be ≥ logo_bottom + 0.05"**. On a closing/title slide where text sits below the logo, the title must start at y ≥ 1.05 — never earlier or it will overlap the logo in the top-left corner.
- Title text: ~36–42pt bold (Open Sans) charcoal, left-aligned at x=0.65–0.7
- Subtitle: 17–20pt neutral4, below title
- **Always set `lineSpacingMultiple`** (≈0.92–0.95) on titles so wrapped lines use tight leading
- No footer on title/divider/closing slides

### 2. Content Slides

Used for: all informational slides.

```javascript
slide.background = { color: "FFFFFF" };  // Always white background
addFooter(slide, pres, pageNumber, logoIcon);  // Always add footer
slide.addText("Slide Title", {
  x: 0.6, y: 0.35, w: 9, h: 0.7,
  fontFace: "Open Sans", fontSize: 30, bold: true,
  color: "3F454D", margin: 0, lineSpacingMultiple: 0.95,
});
```

**Rules:**
- White background, never use the gradient on content slides
- Footer on every content slide (see Footer section below)
- Title: 30pt bold charcoal at y≈0.25–0.35
- Content area: y=1.0 to y=5.0 (above footer) — **no element may have y+h > 5.0**
- Requires `LAYOUT_16x9` (see top of skill) — do NOT use `LAYOUT_WIDE`

### 3. Dark Checklist/Summary Slides

Used for: best practices, key takeaways, summary checklists.

```javascript
slide.background = { color: "3F454D" };  // Charcoal background (darkest official neutral)
// White text, Domo blue (#99CCEE) check icons
// No footer on these slides
```

### Footer (Content Slides Only)

Every content slide gets this footer. Implement as a helper function:

```javascript
function addFooter(slide, pres, pageNum, logoIconData) {
  // Thin separator line
  slide.addShape(pres.shapes.LINE, {
    x: 0.5, y: 5.15, w: 9.0, h: 0,
    line: { color: "B7C1CB", width: 0.5 }
  });
  // Domo icon badge — bottom left. The icon PNG is square (400x400), so KEEP w === h
  // or it renders squished. Never use unequal width/height here.
  slide.addImage({ data: logoIconData, x: 0.35, y: 5.2, w: 0.38, h: 0.38 });
  // CONFIDENTIAL label — spaced letters
  slide.addText("CONFIDENTIAL", {
    x: 7.0, y: 5.2, w: 2.0, h: 0.3,
    fontFace: "Open Sans", fontSize: 8, color: "68737F",
    align: "right", charSpacing: 0, margin: 0
  });
  // Page number
  slide.addText(String(pageNum), {
    x: 9.2, y: 5.2, w: 0.4, h: 0.3,
    fontFace: "Open Sans", fontSize: 9, bold: true, color: "3F454D",
    align: "center", margin: 0
  });
}
```

## Common UI Patterns

### Row alignment rule (icons/glyphs/labels beside text) — READ FIRST

The most common alignment bug in these decks: a small icon, ✓/✕ glyph, cadence label, or
numbered circle sits **too high or too low** relative to the text it labels. The cause is always
the same — the glyph box and the text box have **different heights and/or different `y` origins**,
so their vertical centers don't line up. PptxGenJS centers content on the box's own midpoint, so
two boxes only align when their midpoints match.

**Rule: every element sharing a horizontal row must use the SAME `y`, the SAME `h`, and
`valign: "middle"`.** Only `x` and `w` differ between them. Never give the glyph a shorter box
(e.g. `h: 0.3`) than the text (`h: 0.5`), and never hand-nudge with offsets like `y: y - 0.02`
or `y: y + 0.12` to "eyeball" it — that's what breaks when the text wraps to two lines.

```javascript
// ✗ WRONG — mismatched heights + a nudge → glyph floats above the text
s.addText("✓", { x: 0.6,  y,          w: 0.3, h: 0.3, valign omitted });
s.addText(label, { x: 0.95, y: y - 0.02, w: 3.9, h: 0.5, valign: "middle" });

// ✓ RIGHT — identical y and h, both middle-aligned → centers line up exactly
const rh = 0.5;
s.addText("✓",   { x: 0.6,  y, w: 0.3, h: rh, valign: "middle", fontFace: F, fontSize: 13, bold: true, color: C.aGreen, margin: 0 });
s.addText(label, { x: 0.95, y, w: 3.9, h: rh, valign: "middle", fontFace: F, fontSize: 11, color: C.white, margin: 0 });
```

For a multi-element tile (icon + label + bold title + body in one row, as on engagement/cadence
slides), give **all** of them `y` and `h: rh` with `valign: "middle"` over the full tile height.
Add `lineSpacingMultiple: 0.95` to any title that may wrap to two lines so it stays centered.
`iconChip(...)` is already centered by computing `y + (rh - chip) / 2` — match the text to it,
don't fight it.

### Cards / Feature Tiles

```javascript
// Shadow helper — always create fresh objects (never reuse)
const mkShadow = () => ({
  type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.10
});

slide.addShape(pres.shapes.RECTANGLE, {
  x, y, w: 2.8, h: 1.75,
  fill: { color: "F1F6FA" },  // neutral1 card background
  shadow: mkShadow()
});
```

### Numbered Steps

```javascript
// Circle with number
slide.addShape(pres.shapes.OVAL, {
  x, y, w: 0.45, h: 0.45,
  fill: { color: "99CCEE" }  // domoBlue
});
slide.addText("1", {
  x, y, w: 0.45, h: 0.45,
  fontFace: "Open Sans", fontSize: 16, bold: true, color: "3F454D",  // charcoal — see contrast rule below
  align: "center", valign: "middle", margin: 0
});
```

> **Text on a Domo Blue fill must be `charcoal` (#3F454D), never white.** Domo Blue (#99CCEE) is light, so white text/numbers/icons on it are near-invisible. This applies to numbered circles, domoBlue column/section headers, tier-card header bars, and any chip filled with domoBlue. White text is only for *dark* fills (charcoal #3F454D, or the darker purple accent #776CB0). **Icon circle colour rule:** on a light background use a **charcoal** circle; on a dark (charcoal) background use a **neutral1** circle. Never fill an icon circle with domoBlue — a domoBlue icon on a domoBlue circle is invisible (see Icons).

### Tables

```javascript
// Header row
{ text: "Header", options: {
  bold: true, fontFace: "Open Sans", fontSize: 11,
  color: "FFFFFF",
  fill: { color: "3F454D" },  // charcoal
  valign: "middle"
}}
// Alternating rows
fill: { color: i % 2 === 0 ? "F1F6FA" : "FFFFFF" }
```

### Comparison Columns (Side-by-Side)

```javascript
// Left column header
fill: { color: "99CCEE" }  // domoBlue  → charcoal header text (blue is light)
// Right column header
fill: { color: "3F454D" }  // charcoal   → white header text
```

### Tip/Warning Callouts

```javascript
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0.6, y, w: 8.8, h: 0.8,
  fill: { color: "FFF0DD" }  // warm warning background
});
// Add warning icon + text inside
```

## Icons

Use `react-icons` rendered to PNG via `sharp`. Always render at size 256 for crispness. **Every generated icon is rendered in Domo Blue (#99CCEE)** — pass `"#99CCEE"` as the colour for all icons regardless of slide background.

```javascript
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");

function renderIconSvg(Ic, color, size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(Ic, { color, size: String(size) })
  );
}
async function icon64(Ic, color, size = 256) {
  const svg = renderIconSvg(Ic, color, size);
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// All generated icons are Domo Blue (#99CCEE) — same colour on light and dark slides
const iconData = await icon64(FaDatabase, "#99CCEE");
// Circle colour behind the icon depends on the card/panel background:
//   Light background  → charcoal circle  (domoBlue icon is visible on dark fill)
//   Dark background   → neutral1 circle  (domoBlue icon is visible on light fill)
//   NEVER domoBlue circle — domoBlue icon on domoBlue fill = invisible
```

Install if needed: `npm install -g pptxgenjs react react-dom react-icons sharp`

## Workflow

### Creating from a brief / document / notes

1. Read the input content and identify logical sections
2. Plan slide structure: Title → Content slides → Section dividers if multi-topic → Summary/Checklist → Closing
3. Apply Domo branding per the rules above
4. Generate the deck using PptxGenJS
5. Run QA per the base pptx skill (convert to images, inspect, fix)

### Rebranding an existing .pptx

1. Extract text from the uploaded file: `extract-text uploaded.pptx`
2. Thumbnail it: `python3 /Users/arminas.juknevicius/.claude/skills/pptx/scripts/thumbnail.py uploaded.pptx /tmp/orig_thumb --cols 1`
3. **Extract embedded images** — screenshots and diagrams from the original deck add context that text alone cannot replace; always extract and carry them over:
   ```python
   from pptx import Presentation
   import os
   prs = Presentation("uploaded.pptx")
   os.makedirs("/tmp/orig_images", exist_ok=True)
   for si, slide in enumerate(prs.slides, 1):
       for shape in slide.shapes:
           if shape.shape_type == 13:
               try:
                   img = shape.image
                   path = f"/tmp/orig_images/slide{si:02d}_{shape.shape_id}.{img.ext}"
                   open(path, "wb").write(img.blob)
                   print(f"Slide {si}: {path}  ({shape.width//914400}\" x {shape.height//914400}\")")
               except Exception:
                   pass  # linked images — skip
   ```
   Copy any screenshots you intend to embed to a stable path (e.g. the skill's `assets/screenshots/` directory) before writing the generation script.
4. Map the original content to Domo-branded slide layouts
5. **Retain screenshots on their original slides** — if the original slide had a UI screenshot, product screenshot, or diagram image, embed it in the rebranded slide. Use a two-column layout (text/description on one side, screenshot on the other) or a text-above/screenshot-below layout. Never drop a screenshot just because the content is being rebranded; screenshots provide concrete examples and product proof that text summaries cannot replace.
   ```javascript
   // Load the screenshot as base64 (do this at the top of the async block)
   const scrImg = img64("/path/to/assets/screenshots/slideXX_screenshot.png");
   // Embed it on the slide — use sizing:"contain" to avoid distortion
   s.addImage({ data: scrImg, x: 0.55, y: 1.32, w: 5.5, h: 3.0,
                sizing: { type: "contain", w: 5.5, h: 3.0 } });
   ```
   Aspect-ratio reference when sizing:
   - Square (1:1) screenshot at content-area height 3.72" → use w:3.6, h:3.6
   - 2:1 wide screenshot (e.g. 6"×3" form UI) → scale to fit: w:5.5, h:2.75
   - 5:2 wide screenshot (e.g. activity log table) → w:8.9, h:3.56
   - 7:3 wide screenshot (e.g. output table) → at w:7.0, h:3.0 centered: x:1.5
6. Rebuild slides using PptxGenJS with Domo branding, screenshots embedded
7. Run QA

### Page numbering

- Title and section divider slides do NOT get page numbers or footers
- Content slides are numbered sequentially (start at 1 for the first content slide, not the title)
- Dark checklist slides do NOT get footers

### Speaker notes: documenting changes from the original

**When rebranding an existing deck, every slide where text was changed must include a speaker note** documenting exactly what was changed, why, and where the content came from. This preserves the audit trail so the presenter understands how the rebranded version differs from the source.

Add notes with `slide.addNotes(text)` immediately after all `addShape`/`addText`/`addImage` calls for that slide. Use this format:

```javascript
s.addNotes(
  "CHANGES FROM ORIGINAL (slide N):\n" +
  "- [Element]: [what changed]. Reason: [Domo brand rule / layout constraint / etc.]\n" +
  "- [Element]: [what changed]. Reason: [...].\n" +
  "SOURCE: [original filename], slide [N]."
);
```

**What to document:**
- Text that was shortened, reworded, or split (e.g. "title shortened from 60 chars to 40 to fit 30pt bold in 8.8\" width")
- Em-dashes replaced (brand rule: no em-dashes)
- Colours changed (e.g. "column headers changed from orange to domoBlue per brand guide")
- Logos or images that were dropped, added, or repositioned
- Layout restructuring (e.g. "4-column grid replaced with left-screenshot + right-list to retain original screenshots")
- Any slide where text was NOT changed may omit the note, or include a brief "No text changes; colour/layout only." line

**Slides with logos or diagrams:** always note which logos were retained from the original, which (if any) could not be embedded, and why.

Example:

```javascript
s.addNotes(
  "CHANGES FROM ORIGINAL (slide 4):\n" +
  "- Colour: column header fills changed from orange (#FF6600) to Domo Blue (#99CCEE). Reason: Domo brand guide specifies domoBlue as the primary accent.\n" +
  "- Icon circles: changed from domoBlue fill to neutral1 fill. Reason: domoBlue icon on domoBlue circle is invisible (brand rule).\n" +
  "- Logos retained: Snowflake, Databricks, GCP, AWS, Azure (cloud); Anthropic, OpenAI, Vertex AI, Gemini (AI models) — all embedded from the source file.\n" +
  "SOURCE: OriginalDeck.pptx, slide 4."
);
```

## Design Don'ts

- **Don't use cream/beige backgrounds** — content slides are always white (#FFFFFF)
- **Don't use accent lines under titles** — the Domo style uses whitespace separation
- **Don't use colored bars/ribbons** as decorative elements — only the thin 5px domoBlue top accent on cards
- **Don't use em-dashes (—) in slide text** — they read as AI-generated. Use a comma, a colon, or split into two sentences instead (e.g. "Build from a brief, notes, or an existing deck" not "Build from a brief — notes — or an existing deck"). This applies to titles, subtitles, body copy, captions, and bullets. **This rule applies even when ingesting source content verbatim** — original briefs, notes, or existing decks will often contain em-dashes; strip them during ingestion before writing any string to the generation script.
- **Don't use emoji or raw Unicode symbols as slide content** — they render inconsistently across platforms (e.g. `▶` U+25B6 renders as an emoji on some systems). Use react-icons rendered to PNG via `icon64()` for all directional arrows, checkmarks, and decorative glyphs. Plain text characters like ✓ are acceptable only inside `addText()` for checkmark bullets on dark slides where the react-icon approach is overkill; for any arrow between columns or section elements, always use `icArrow` (or another react-icon image).
- **Don't use any font other than Open Sans** — the entire deck is Open Sans (Calibri only as a system fallback)
- **Don't use the Domo blue for text** — it's an accent/fill color only
- **Don't use any other shade of blue** — only Domo Blue #99CCEE; never a similar-but-different blue
- **Don't let accents overpower Domo Blue** — orange/purple/magenta/pink/green are used sparingly
- **Don't skip the footer** on content slides — it's mandatory for brand consistency
- **Don't place the logo elsewhere** — top-left on gradient slides, bottom-left badge on content slides
- **Don't misalign elements in a row** — an icon/glyph/label and the text it sits beside must share the same `y`, same `h`, and `valign: "middle"` (see "Row alignment rule" above). No mismatched box heights, no hand-nudged `y` offsets.
- **Don't use domoBlue as an icon-circle fill** — a domoBlue icon inside a domoBlue circle is completely invisible. Use charcoal circle on light backgrounds, neutral1 circle on dark backgrounds (charcoal cards/panels).
- **Don't set OVAL `w` and `h` to different values** — unequal dimensions render as an ellipse, not a circle. Always use `w === h` for every circle shape, and center it vertically within its row using `y + (rowHeight - circleSize) / 2`.
- **Don't let a title overflow into a LENS chip** — `titleBar()` accepts a fifth argument `hasChip`. Pass `true` on every slide that also calls `lensChip()`: `titleBar(s, pres, title, sub, true)`. This caps the title and subtitle to `w:7.5` (right edge `0.6+7.5=8.1"`), leaving 0.4" clearance before the chip at `x:8.5`. Never pass the default (`false`) after a `lensChip()` call — even a single-line 30pt title that appears to fit can render right on top of the chip at certain zoom levels.
- **Don't let the blue accent bar cut through a wrapped title — pass `lines=N` to `titleBar()`.** The accent bar is placed at `y = 0.28 + titleH + 0.02` where `titleH = 0.58 + (lines-1)*0.42`. For a title that wraps to 2 lines, use `titleBar(s, pres, title, sub, hasChip, 2)`. Content on that slide must then start at `y ≥ 1.72` (versus `y ≥ 1.28` for 1-line). **Character-count thresholds (count the title string including spaces):**
  - `hasChip=true` (w:7.5"): safe ≤ 32 chars. **Treat 33+ chars as a confirmed 2-liner — always pass `lines=2`.**
  - `hasChip=false` (w:8.8"): safe ≤ 38 chars. Treat 39+ chars as a confirmed 2-liner.
  - The "35–40 chars per line" estimate is too loose — a 36-char title with wide glyphs (M, W, capitals) will wrap. Use 32/38 as the hard cutoff, not an estimate. When in doubt, count and round down.
  - Slides with 2-line titles that use pillarList must also reduce `ih` and adjust `y`: `pillarList(s, pres, x, 1.72, w, items, hi, 0.73, 0.06)` keeps the 4-row list within `y:5.0`. For other content (stage cards, step tiles, screenshots), set `sy:1.72` and reduce heights so the last element ends at `y ≤ 4.82` if a callout/banner follows, or `y ≤ 5.0` otherwise.
- **Don't indent body text in tall card columns — use `fullWidthBody:true`.** `card()` accepts a `fullWidthBody` option (default false). When true, the body text spans the full card width (`x+0.18, w-0.36`) starting below the icon row (`y+0.72`) rather than being indented to align with the title. Use this on tall cards (`h ≥ 2.5"`) where the indented body creates a visually misaligned column of text in the lower two-thirds of the card.
- **Don't drop logos when rebranding** — if the original slide contains third-party logos (cloud providers, AI vendors, partner logos), extract and re-embed them in the rebranded version. Keep the same relative positions and proportions. Apply Domo colours only to backgrounds, borders, and text — never alter the logos themselves. If a logo genuinely cannot be embedded, document it in the slide note.
- **Don't use domoBlue-filled icon circles on dark slides** — applies to ALL slides with `charcoal` or dark panel backgrounds: use `neutral1` circle fill so the domoBlue icon glyph is actually visible.
- **Don't let arrows overlap excessively into adjacent column boxes** — before placing an arrow between columns, calculate the gap: `gap = col2_x - (col1_x + col1_w)`. Always use react-icon images (`icArrow`) for arrows, never Unicode characters. If the gap is less than 0.20", center a 0.16"×0.16" image at the gap midpoint — it will extend ~0.04" into each adjacent column, which is acceptable. Never use `▶` (U+25B6) or similar Unicode symbols — they render as emoji on many platforms.
- **Don't let multi-row content overflow the footer** — for N rows with height `rh` and gap `g`, calculate the last row's bottom: `y0 + (N-1)*(rh+g) + rh`. This must be ≤ 5.0. Always compute this before writing the loop. Example: 5 rows at `rh:0.62, g:0.10, y0:1.38` → last row ends at `1.38 + 4×0.72 + 0.62 = 4.88` (safe). At `rh:0.68, g:0.16` the same 5 rows end at `5.42` — 1.42" overflow. **This applies to the `card()` helper too** — `card(s, pres, x, y0 + i*(ch+gap), w, ch, ...)` loops produce the same pattern. 3 cards at `ch:1.22, gap:0.22, y0:1.32` end at `5.42`; fix to `ch:1.10, gap:0.18` → last card ends at `4.98`.
- **Don't position the primary logo where it overlaps other content** — on dark (gradient/charcoal) slides the primary logo sits at `x:0.0, y:0.0, w:0.92, h:0.92`. Any banner, shape, or text placed below it must start at `y ≥ 0.94` (logo bottom 0.92 + 0.02 clearance). Placing a banner at `y:0.84` while the logo bottom is `y:0.92` causes an 0.08" overlap.
- **Don't place bar chart bars that exceed the panel or slide right margin** — compute max bar right edge: `bx + maxBarWidth ≤ panel_x + panel_w` AND `≤ 9.5`. Scale `maxBarWidth` to the available space: `maxBarWidth = min(panel_w - (bx - panel_x) - 0.05, 9.5 - bx - 0.05)`. Both the background track and the filled bar must use this capped width.
- **Don't write caption text boxes with `h:0` or `fontSize:0`** — PptxGenJS passes the text to PowerPoint, which auto-sizes zero-height boxes using its own default font size. The resulting text can appear anywhere and overflow the footer. Always give caption text boxes a real positive height (`h ≥ 0.12`) and a real font size (`fontSize ≥ 7`).
- **Don't use hardcoded icon offsets in tall card columns — vertically center the content block instead.** When an icon, title, and body are stacked vertically inside a card, compute the total content height and distribute padding symmetrically:
  1. Measure: `contentH = iconSize + titleGap + titleH + bodyGap + bodyH`
  2. Available: `availH = cardH - headerBandH`
  3. Padding: `pad = (availH - contentH) / 2`
  4. `circY = cardTopY + headerBandH + pad`
  A fixed offset like `circY = cy + 0.58` only works at the exact card height it was written for. When `cardH` changes, the icon drifts top-heavy and leaves empty space at the card bottom.
- **Don't over-size body text boxes in card columns.** A body box `h:1.4` for 1–2 lines of text leaves ~1" of invisible empty space at the bottom. Estimate body height from line count: at 10pt with `lineSpacingMultiple:1.2`, each line ≈ 0.25". Use that to size `h` instead of a blanket "big enough" value. Oversized body boxes compound with a non-centered icon offset to make content look clustered at the top of the card.

## QA (Required)

Follow the base pptx skill's QA process exactly:
1. Convert to PDF and then to images
2. Visually inspect every slide for overlaps, overflow, alignment issues
3. Check that every content slide has the footer (logo + CONFIDENTIAL + page number)
4. Verify gradient background is visible on title/divider/closing slides
5. Fix issues and re-verify
