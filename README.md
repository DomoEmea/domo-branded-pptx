# domo-branded-pptx

A Claude Code skill that generates professional, on-brand Domo PowerPoint presentations from a brief, document, or notes — using PptxGenJS under the hood.

---

## What it does

- Builds complete `.pptx` decks from a prompt or structured brief
- Applies official Domo brand colors, typography, logo, and layout rules
- Generates react-icons rendered as PNG (via `sharp`) for slide icons
- Runs full QA: converts to PDF, renders to images, inspects every slide
- Supports rebranding existing decks in Domo style

## When to invoke

Trigger this skill whenever the user asks to:
- Create a presentation, deck, or slides with Domo branding
- Rebrand an existing `.pptx` in Domo style
- Turn a brief, notes, or document into a Domo-branded deck

Keywords: *Domo slides, Domo deck, Domo presentation, Domo branding, Domo template, Domo PowerPoint, branded slides.*

---

## Repository structure

```
domo-branded-pptx/
├── SKILL.md                              # Full skill instructions (read this first)
├── README.md                             # This file
├── gen_ai_can_should.js                  # Generation script: AI Can vs Should deck (27 slides)
├── .gitignore
├── assets/
│   ├── domo_gradient_bg.png              # Title/closing slide gradient background
│   ├── domo_logo_blue_800.png            # Primary logo: blue box + white DOMO wordmark
│   ├── domo_icon_400.png                 # Footer badge: small square DOMO icon
│   └── screenshots/                     # Extracted UI screenshots from source decks
│       ├── s04_cloud_logos.png           # Snowflake, Databricks, GCP, AWS, Azure strip
│       ├── s04_ai_logos.png              # Anthropic, OpenAI, Vertex AI, Gemini strip
│       ├── s10_domo_text_gen.png         # Domo Text Generation tile UI
│       ├── s11_post_summary.png          # PostSummary output example
│       ├── s12_agent_general.png         # AI Agent General tab
│       ├── s13_agent_knowledge.png       # AI Agent Knowledge tab
│       ├── s14_agent_tools.png           # AI Agent Tools tab
│       └── s18_ai_activity_log.png       # Domo AI Activity Log screenshot
└── references/
    └── Domo Brand Guide - Colours.pptx  # Official Domo brand guide (source of truth)
```

### Running the generation script

```bash
cd /Users/arminas.juknevicius/.claude/skills/domo-branded-pptx
NODE_PATH=$(npm root -g) node gen_ai_can_should.js
# Output: ~/Downloads/AI_Can_vs_Should_Domo.pptx
```

---

## Critical setup rules

### 1. Slide layout — always LAYOUT_16x9

```javascript
const pres = new PptxGenJS();
pres.layout = "LAYOUT_16x9";  // 10" wide x 5.625" tall
```

`LAYOUT_WIDE` (13.33" x 7.5") will cause all content to compress into 75% of the slide. `LAYOUT_16x9` is the only layout whose coordinate system matches this skill.

### 2. Logo on gradient slides — never negative y

```javascript
// Correct
slide.addImage({ data: logoPrimary, x: 0.0, y: 0.0, w: 1.0, h: 1.0 });

// Wrong — clips the logo above the slide edge
slide.addImage({ data: logoPrimary, x: 0.0, y: -0.1, w: 1.2, h: 1.2 });
```

On closing/title slides, the first text element must start at `y >= logo_bottom + 0.05` (i.e. `y >= 1.05` with a 1.0" logo).

### 3. Content bounds

| Zone            | Y range       |
|-----------------|---------------|
| Title           | 0.2 – 0.9     |
| Content         | 1.0 – **5.0** |
| Footer line     | 5.15          |
| Footer elements | 5.18 – 5.55   |
| Slide bottom    | 5.625         |

No content element may have `y + h > 5.0` (footer zone starts at 5.15).

---

## Dependencies

Install before running any generation script:

```bash
npm install -g pptxgenjs react react-dom react-icons sharp
```

---

## Brand colors (key)

| Token        | Hex       | Use                                  |
|--------------|-----------|--------------------------------------|
| `domoBlue`   | `#99CCEE` | Fills, accents, icons, logo box      |
| `charcoal`   | `#3F454D` | Heading text, dark panels, icon text |
| `neutral1`   | `#F1F6FA` | Card backgrounds                     |
| `neutral4`   | `#68737F` | Body text, captions, footer          |
| `white`      | `#FFFFFF` | Content slide backgrounds            |

> Never use `domoBlue` as a text color — it is too light on white. Use it for fills only.

Full palette, typography table, icon rules, and QA process are in [`SKILL.md`](./SKILL.md).

---

## Changelog

| Date       | Change                                                                              |
|------------|-------------------------------------------------------------------------------------|
| 2026-05-29 | Skill installed; replaced old python-pptx domo-pptx skill                           |
| 2026-06-16 | Fixed layout bug (LAYOUT_WIDE → LAYOUT_16x9)                                        |
| 2026-06-16 | Fixed logo clipping (y: -0.08 → y: 0.0 on gradient slides)                          |
| 2026-06-16 | Added title clear-logo rule for closing slides                                       |
| 2026-06-16 | Added coordinate bounds table and content limit warnings to SKILL.md                 |
| 2026-06-16 | Added AI Can vs Should deck (gen_ai_can_should.js, 27 slides); retained Domo UI screenshots and original logos |
| 2026-06-16 | Fixed 7 layout overflow/overlap issues: slide 3 footer clearance, slide 4 logo/banner overlap and column arrows, slide 5 5-row overflow, slide 8 gap arrow/text, slide 11 zero-height caption bug, slide 18 field chip overflow, slide 21 bar chart right margin, slide 27 row overflow |
| 2026-06-16 | Added 5 anti-pattern rules to SKILL.md: arrow gap math, multi-row overflow formula, logo clearance, bar chart bounds, zero-height text boxes |
| 2026-06-16 | Fixed slide 2 card overflow (ch:1.22→1.10, gap:0.22→0.18 — last card now ends at y:4.98); extended multi-row rule in SKILL.md to cover card() helper |
| 2026-06-16 | Fixed title/LENS-chip overlap on slide 20 (and all 13 other LENS slides): titleBar() now accepts hasChip flag, caps title+subtitle to w:7.5 when true; SKILL.md rule updated |
| 2026-06-16 | Fixed icon centering on slides 3, 16, 22: vertically centered content blocks using padding formula; tightened oversized body text boxes (h:1.4→h:0.55, h:1.0→h:0.80, h:1.45→h:0.85); added 2 SKILL.md rules (center-content-block formula, body-box sizing) |
| 2026-06-16 | Fixed blue accent bar overlapping 2-line titles on slides 12, 13, 17, 18, 20: titleBar() now accepts lines param (default 1); each extra line adds 0.42" to title box and shifts accent+subtitle accordingly. Converted slides 12/13 from manual title code to titleBar(lines=2). Added fullWidthBody option to card() for slide 15 body alignment. Added 3 SKILL.md rules (lines param, content y offset, fullWidthBody) |
