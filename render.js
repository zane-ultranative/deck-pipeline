'use strict';
const PptxGenJS = require('pptxgenjs');

// ── Brand tokens ──────────────────────────────────────────────
const T = {
  BG:     '0F1923',   // dark navy  — page background
  CARD:   '162130',   // card surface
  RULE:   '1E2E3D',   // divider / border
  WHITE:  'F0EDE6',   // warm white — primary text
  MUTED:  '8090A0',   // blue-grey  — secondary text
  ORANGE: 'E8622A',   // brand accent
  FONT:   'Arial',
};

// Slide canvas: LAYOUT_WIDE = 13.33" × 7.5"
const W = 13.33;
const H = 7.5;

// ── Low-level helpers ─────────────────────────────────────────

function setBg(slide, color = T.BG) {
  slide.background = { color };
}

/** Thin horizontal rule — defaults to full-width orange */
function rule(slide, y, opts = {}) {
  const { x = 0.6, w = W - 1.2, h = 0.035, color = T.ORANGE } = opts;
  slide.addShape('rect', { x, y, w, h, fill: { color }, line: { color } });
}

/** Small eyebrow label above the rule */
function eyebrow(slide, label, y = 0.52) {
  if (!label) return;
  slide.addText(label.toUpperCase(), {
    x: 0.6, y, w: W - 1.2, h: 0.28,
    fontFace: T.FONT, fontSize: 9.5, bold: true,
    color: T.ORANGE, charSpacing: 2.5, align: 'left',
  });
}

/** Slide headline — large, white, bold */
function headline(slide, text, y = 1.05, overrides = {}) {
  if (!text) return;
  slide.addText(text, {
    x: 0.6, y, w: W - 1.2, h: 1.5,
    fontFace: T.FONT, fontSize: 32, bold: true,
    color: T.WHITE, lineSpacingMultiple: 1.1,
    ...overrides,
  });
}

/** Body paragraph — muted, readable */
function body(slide, text, y = 2.65, overrides = {}) {
  if (!text) return;
  slide.addText(text, {
    x: 0.6, y, w: W - 1.2, h: 3.6,
    fontFace: T.FONT, fontSize: 17, color: T.MUTED,
    lineSpacingMultiple: 1.55, wrap: true,
    ...overrides,
  });
}

/** Standard header block: eyebrow + rule + headline */
function header(slide, c, headlineY = 1.05, headlineOpts = {}) {
  eyebrow(slide, c.eyebrow);
  rule(slide, 0.9);
  headline(slide, c.headline, headlineY, headlineOpts);
}

// ── Bullet helpers ────────────────────────────────────────────

function bulletItems(points, color = T.WHITE, size = 16) {
  return (points || []).map(p => ({
    text: p,
    options: {
      bullet: { code: '25A0', color: T.ORANGE },
      color,
      fontSize: size,
    },
  }));
}

function checkItems(points, color = T.WHITE, size = 16) {
  return (points || []).map(p => ({
    text: p,
    options: {
      bullet: { code: '2713', color: T.ORANGE },
      color,
      fontSize: size,
    },
  }));
}

// ── Slide type renderers ──────────────────────────────────────

/**
 * Title
 * content: { eyebrow?, title, subtitle?, meta? }
 */
function slideTitle(slide, c) {
  setBg(slide);
  // Orange strip at bottom
  slide.addShape('rect', {
    x: 0, y: H - 0.1, w: W, h: 0.1,
    fill: { color: T.ORANGE }, line: { color: T.ORANGE },
  });
  if (c.eyebrow) eyebrow(slide, c.eyebrow, 2.05);
  slide.addText(c.title || 'Untitled', {
    x: 0.7, y: 2.45, w: 10.5, h: 2.2,
    fontFace: T.FONT, fontSize: 54, bold: true,
    color: T.WHITE, lineSpacingMultiple: 1.05,
  });
  if (c.subtitle) {
    slide.addText(c.subtitle, {
      x: 0.7, y: 4.9, w: 9, h: 0.65,
      fontFace: T.FONT, fontSize: 19, color: T.MUTED,
    });
  }
  if (c.meta) {
    slide.addText(c.meta, {
      x: W - 3.9, y: H - 0.55, w: 3.5, h: 0.3,
      fontFace: T.FONT, fontSize: 11, color: T.MUTED, align: 'right',
    });
  }
}

/**
 * Statement — centred bold assertion, optional attribution
 * content: { statement, attribution? }
 */
function slideStatement(slide, c) {
  setBg(slide);
  slide.addText(c.statement || '', {
    x: 1.0, y: 1.7, w: W - 2, h: 3.8,
    fontFace: T.FONT, fontSize: 38, bold: true,
    color: T.WHITE, align: 'left', lineSpacingMultiple: 1.25,
  });
  rule(slide, 5.65, { x: 1.0, w: 1.6 });
  if (c.attribution) {
    slide.addText('— ' + c.attribution, {
      x: 1.0, y: 5.82, w: W - 2, h: 0.5,
      fontFace: T.FONT, fontSize: 14, color: T.MUTED,
    });
  }
}

/**
 * Problem
 * content: { eyebrow?, headline, body?, points? }
 */
function slideProblem(slide, c) {
  setBg(slide);
  header(slide, { eyebrow: c.eyebrow || 'The Problem', headline: c.headline });
  if (c.points && c.points.length) {
    slide.addText(bulletItems(c.points), {
      x: 0.6, y: 2.65, w: W - 1.2, h: 4.0,
      fontFace: T.FONT, lineSpacingMultiple: 1.7, paraSpaceAfter: 4,
    });
  } else {
    body(slide, c.body);
  }
}

/**
 * Opportunity
 * content: { eyebrow?, headline, body?, points? }
 */
function slideOpportunity(slide, c) {
  setBg(slide);
  header(slide, { eyebrow: c.eyebrow || 'The Opportunity', headline: c.headline });
  if (c.points && c.points.length) {
    slide.addText(bulletItems(c.points), {
      x: 0.6, y: 2.65, w: W - 1.2, h: 4.0,
      fontFace: T.FONT, lineSpacingMultiple: 1.7, paraSpaceAfter: 4,
    });
  } else {
    body(slide, c.body);
  }
}

/**
 * Bullets — headline + bulleted list
 * content: { eyebrow?, headline, bullets/points }
 */
function slideBullets(slide, c) {
  setBg(slide);
  header(slide, c);
  const pts = c.bullets || c.points || [];
  slide.addText(bulletItems(pts), {
    x: 0.6, y: 2.65, w: W - 1.2, h: 4.0,
    fontFace: T.FONT, lineSpacingMultiple: 1.7, paraSpaceAfter: 4,
  });
}

/**
 * Framework — labelled boxes in a row
 * content: { eyebrow?, headline, items: [{ number?, label, body? }] }
 */
function slideFramework(slide, c) {
  setBg(slide);
  header(slide, c, 1.05, { fontSize: 26, h: 0.9 });
  const items = c.items || [];
  const n     = Math.max(items.length, 1);
  const gap   = 0.22;
  const boxW  = (W - 1.2 - (n - 1) * gap) / n;
  const boxY  = 2.2;
  const boxH  = 4.85 - boxY + 0.3;

  items.forEach((item, i) => {
    const x = 0.6 + i * (boxW + gap);
    slide.addShape('rect', {
      x, y: boxY, w: boxW, h: boxH,
      fill: { color: T.CARD }, line: { color: T.RULE, pt: 0.5 },
    });
    slide.addShape('rect', {
      x, y: boxY, w: boxW, h: 0.055,
      fill: { color: T.ORANGE }, line: { color: T.ORANGE },
    });
    if (item.number) {
      slide.addText(String(item.number), {
        x, y: boxY + 0.18, w: boxW, h: 0.65,
        fontFace: T.FONT, fontSize: 26, bold: true,
        color: T.ORANGE, align: 'center',
      });
    }
    if (item.label) {
      slide.addText(item.label, {
        x, y: boxY + (item.number ? 0.88 : 0.2), w: boxW, h: 0.5,
        fontFace: T.FONT, fontSize: 14, bold: true,
        color: T.WHITE, align: 'center',
      });
    }
    if (item.body) {
      slide.addText(item.body, {
        x: x + 0.18, y: boxY + (item.number ? 1.5 : 0.82),
        w: boxW - 0.36, h: 2.4,
        fontFace: T.FONT, fontSize: 13, color: T.MUTED,
        align: 'left', lineSpacingMultiple: 1.5, wrap: true,
      });
    }
  });
}

/**
 * Solution — headline, optional body, optional checklist
 * content: { eyebrow?, headline, body?, points? }
 */
function slideSolution(slide, c) {
  setBg(slide);
  header(slide, { eyebrow: c.eyebrow || 'The Solution', headline: c.headline });
  if (c.body) body(slide, c.body, 2.65, { h: 1.4, fontSize: 17 });
  if (c.points && c.points.length) {
    slide.addText(checkItems(c.points), {
      x: 0.6, y: c.body ? 4.25 : 2.65,
      w: W - 1.2, h: 3.0,
      fontFace: T.FONT, lineSpacingMultiple: 1.7,
    });
  }
}

/**
 * Case Study — company name, challenge / outcome two-pane, optional stat
 * content: { eyebrow?, company, challenge, outcome, stat?: { value, label } }
 */
function slideCaseStudy(slide, c) {
  setBg(slide);
  eyebrow(slide, c.eyebrow || 'Case Study');
  rule(slide, 0.9);
  slide.addText(c.company || '', {
    x: 0.6, y: 1.05, w: W - 1.2, h: 0.85,
    fontFace: T.FONT, fontSize: 30, bold: true, color: T.WHITE,
  });

  const half  = (W - 1.6) / 2;
  const colY  = 2.15;
  const rx    = 0.6 + half + 0.4;

  // Challenge
  slide.addText('CHALLENGE', {
    x: 0.6, y: colY, w: half, h: 0.28,
    fontFace: T.FONT, fontSize: 9.5, bold: true,
    color: T.ORANGE, charSpacing: 2.5,
  });
  slide.addText(c.challenge || '', {
    x: 0.6, y: colY + 0.38, w: half, h: 3.6,
    fontFace: T.FONT, fontSize: 16, color: T.MUTED,
    lineSpacingMultiple: 1.55, wrap: true,
  });

  // Outcome
  slide.addText('OUTCOME', {
    x: rx, y: colY, w: half, h: 0.28,
    fontFace: T.FONT, fontSize: 9.5, bold: true,
    color: T.ORANGE, charSpacing: 2.5,
  });
  slide.addText(c.outcome || '', {
    x: rx, y: colY + 0.38, w: half, h: c.stat ? 2.4 : 3.6,
    fontFace: T.FONT, fontSize: 16, color: T.MUTED,
    lineSpacingMultiple: 1.55, wrap: true,
  });

  if (c.stat) {
    slide.addText(c.stat.value, {
      x: rx, y: colY + 3.0, w: half, h: 1.0,
      fontFace: T.FONT, fontSize: 44, bold: true, color: T.ORANGE,
    });
    slide.addText(c.stat.label, {
      x: rx, y: colY + 4.0, w: half, h: 0.4,
      fontFace: T.FONT, fontSize: 13, color: T.MUTED,
    });
  }
}

/**
 * Data — stat cards in a row
 * content: { eyebrow?, headline?, stats: [{ value, label, sub? }] }
 */
function slideData(slide, c) {
  setBg(slide);
  header(slide, c, 1.05, { fontSize: 24, h: 0.8 });
  const stats = c.stats || [];
  const n     = Math.max(stats.length, 1);
  const gap   = 0.3;
  const cardW = (W - 1.2 - (n - 1) * gap) / n;
  const cardY = 2.2;
  const cardH = 4.0;

  stats.forEach((stat, i) => {
    const x = 0.6 + i * (cardW + gap);
    slide.addShape('rect', {
      x, y: cardY, w: cardW, h: cardH,
      fill: { color: T.CARD }, line: { color: T.RULE, pt: 0.5 },
    });
    slide.addShape('rect', {
      x, y: cardY, w: cardW, h: 0.055,
      fill: { color: T.ORANGE }, line: { color: T.ORANGE },
    });
    slide.addText(stat.value, {
      x, y: cardY + 0.55, w: cardW, h: 1.35,
      fontFace: T.FONT, fontSize: 52, bold: true,
      color: T.WHITE, align: 'center',
    });
    slide.addText(stat.label, {
      x, y: cardY + 2.0, w: cardW, h: 0.45,
      fontFace: T.FONT, fontSize: 13, bold: true,
      color: T.ORANGE, align: 'center',
    });
    if (stat.sub) {
      slide.addText(stat.sub, {
        x: x + 0.15, y: cardY + 2.55, w: cardW - 0.3, h: 1.2,
        fontFace: T.FONT, fontSize: 13, color: T.MUTED,
        align: 'center', lineSpacingMultiple: 1.5, wrap: true,
      });
    }
  });
}

/**
 * Two-column — headline + left/right columns with optional bullets or body
 * content: {
 *   eyebrow?, headline,
 *   left:  { label?, body?, bullets? },
 *   right: { label?, body?, bullets? }
 * }
 */
function slideTwoColumn(slide, c) {
  setBg(slide);
  header(slide, c);
  const half = (W - 1.6) / 2;
  const colY = 2.45;
  const rx   = 0.6 + half + 0.4;

  function renderCol(col, x) {
    if (!col) return;
    if (col.label) {
      slide.addText(col.label.toUpperCase(), {
        x, y: colY, w: half, h: 0.28,
        fontFace: T.FONT, fontSize: 9.5, bold: true,
        color: T.ORANGE, charSpacing: 2,
      });
    }
    const contentY = colY + (col.label ? 0.4 : 0);
    if (col.bullets && col.bullets.length) {
      slide.addText(bulletItems(col.bullets, T.WHITE, 15), {
        x, y: contentY, w: half, h: 4.3,
        fontFace: T.FONT, lineSpacingMultiple: 1.65,
      });
    } else if (col.body) {
      slide.addText(col.body, {
        x, y: contentY, w: half, h: 4.3,
        fontFace: T.FONT, fontSize: 16, color: T.MUTED,
        lineSpacingMultiple: 1.55, wrap: true,
      });
    }
  }

  renderCol(c.left,  0.6);
  renderCol(c.right, rx);
}

/**
 * CTA — strong call to action, contact line
 * content: { eyebrow?, headline, sub?, contact? }
 */
function slideCTA(slide, c) {
  setBg(slide);
  // Left orange bar
  slide.addShape('rect', {
    x: 0, y: 0, w: 0.2, h: H,
    fill: { color: T.ORANGE }, line: { color: T.ORANGE },
  });
  if (c.eyebrow) eyebrow(slide, c.eyebrow, 1.85);
  slide.addText(c.headline || '', {
    x: 0.75, y: 2.2, w: W - 1.5, h: 2.2,
    fontFace: T.FONT, fontSize: 42, bold: true,
    color: T.WHITE, lineSpacingMultiple: 1.1,
  });
  if (c.sub) {
    slide.addText(c.sub, {
      x: 0.75, y: 4.55, w: W - 1.5, h: 0.65,
      fontFace: T.FONT, fontSize: 18, color: T.MUTED,
    });
  }
  if (c.contact) {
    slide.addText(c.contact, {
      x: 0.75, y: H - 0.9, w: 8, h: 0.4,
      fontFace: T.FONT, fontSize: 14, color: T.ORANGE,
    });
  }
}

/**
 * Quote — decorative quote mark, italic text, attribution
 * content: { quote, attribution?, role? }
 */
function slideQuote(slide, c) {
  setBg(slide);
  // Decorative open-quote
  slide.addText('\u201C', {
    x: 0.3, y: 0.3, w: 2.5, h: 2,
    fontFace: T.FONT, fontSize: 180, bold: true,
    color: T.ORANGE, align: 'left', valign: 'top',
  });
  slide.addText(c.quote || '', {
    x: 1.0, y: 1.55, w: W - 2.1, h: 3.7,
    fontFace: T.FONT, fontSize: 28, color: T.WHITE,
    lineSpacingMultiple: 1.4, italic: true,
  });
  rule(slide, 5.52, { x: 1.0, w: 1.4 });
  if (c.attribution) {
    slide.addText('— ' + c.attribution, {
      x: 1.0, y: 5.7, w: W - 2, h: 0.42,
      fontFace: T.FONT, fontSize: 15, color: T.MUTED,
    });
  }
  if (c.role) {
    slide.addText(c.role, {
      x: 1.0, y: 6.15, w: W - 2, h: 0.35,
      fontFace: T.FONT, fontSize: 12, color: T.MUTED,
    });
  }
}

/**
 * Timeline — horizontal dot-and-line with labelled events
 * content: { eyebrow?, headline?, items: [{ date?, label, body? }] }
 */
function slideTimeline(slide, c) {
  setBg(slide);
  header(slide, c, 1.05, { fontSize: 26, h: 0.8 });
  const items  = c.items || [];
  const n      = Math.max(items.length, 1);
  const dotCY  = 3.35;
  const dotR   = 0.14;
  const itemW  = (W - 1.2) / n;

  // Connecting line
  slide.addShape('rect', {
    x: 0.6, y: dotCY - 0.02, w: W - 1.2, h: 0.04,
    fill: { color: T.RULE }, line: { color: T.RULE },
  });

  items.forEach((item, i) => {
    const cx = 0.6 + i * itemW + itemW / 2;
    // Dot
    slide.addShape('ellipse', {
      x: cx - dotR, y: dotCY - dotR,
      w: dotR * 2, h: dotR * 2,
      fill: { color: T.ORANGE }, line: { color: T.ORANGE },
    });
    // Date above dot
    if (item.date) {
      slide.addText(item.date, {
        x: cx - itemW * 0.45, y: dotCY - 0.72,
        w: itemW * 0.9, h: 0.32,
        fontFace: T.FONT, fontSize: 10.5, bold: true,
        color: T.ORANGE, align: 'center', charSpacing: 1,
      });
    }
    // Label below dot
    if (item.label) {
      slide.addText(item.label, {
        x: cx - itemW * 0.45, y: dotCY + 0.28,
        w: itemW * 0.9, h: 0.45,
        fontFace: T.FONT, fontSize: 14, bold: true,
        color: T.WHITE, align: 'center',
      });
    }
    if (item.body) {
      slide.addText(item.body, {
        x: cx - itemW * 0.45, y: dotCY + 0.82,
        w: itemW * 0.9, h: 2.4,
        fontFace: T.FONT, fontSize: 13, color: T.MUTED,
        align: 'center', lineSpacingMultiple: 1.5, wrap: true,
      });
    }
  });
}

/**
 * Comparison — two titled panels (hero vs. alternative)
 * content: {
 *   eyebrow?, headline?,
 *   left:  { label, points[] },
 *   right: { label, points[] }
 * }
 */
function slideComparison(slide, c) {
  setBg(slide);
  header(slide, c, 1.05, { fontSize: 26, h: 0.8 });
  const panelW = (W - 1.6) / 2;
  const panelY = 2.15;
  const panelH = 4.9;

  // Left panel — highlighted
  slide.addShape('rect', {
    x: 0.6, y: panelY, w: panelW, h: panelH,
    fill: { color: T.CARD }, line: { color: T.RULE, pt: 0.5 },
  });
  slide.addShape('rect', {
    x: 0.6, y: panelY, w: panelW, h: 0.055,
    fill: { color: T.ORANGE }, line: { color: T.ORANGE },
  });

  // Right panel — dimmed
  const rx = 0.6 + panelW + 0.4;
  slide.addShape('rect', {
    x: rx, y: panelY, w: panelW, h: panelH,
    fill: { color: T.CARD }, line: { color: T.RULE, pt: 0.5 },
  });
  slide.addShape('rect', {
    x: rx, y: panelY, w: panelW, h: 0.055,
    fill: { color: T.RULE }, line: { color: T.RULE },
  });

  if (c.left) {
    slide.addText(c.left.label || '', {
      x: 0.6, y: panelY + 0.15, w: panelW, h: 0.52,
      fontFace: T.FONT, fontSize: 16, bold: true,
      color: T.WHITE, align: 'center',
    });
    slide.addText(bulletItems(c.left.points || [], T.WHITE, 14), {
      x: 0.78, y: panelY + 0.82, w: panelW - 0.36, h: 3.7,
      fontFace: T.FONT, lineSpacingMultiple: 1.7,
    });
  }

  if (c.right) {
    slide.addText(c.right.label || '', {
      x: rx, y: panelY + 0.15, w: panelW, h: 0.52,
      fontFace: T.FONT, fontSize: 16, bold: true,
      color: T.MUTED, align: 'center',
    });
    slide.addText(bulletItems(c.right.points || [], T.MUTED, 14), {
      x: rx + 0.18, y: panelY + 0.82, w: panelW - 0.36, h: 3.7,
      fontFace: T.FONT, lineSpacingMultiple: 1.7,
    });
  }
}

// ── Renderer dispatch ─────────────────────────────────────────

const RENDERERS = {
  'title':        slideTitle,
  'statement':    slideStatement,
  'problem':      slideProblem,
  'opportunity':  slideOpportunity,
  'bullets':      slideBullets,
  'framework':    slideFramework,
  'solution':     slideSolution,
  'case study':   slideCaseStudy,
  'casestudy':    slideCaseStudy,
  'data':         slideData,
  'two-column':   slideTwoColumn,
  'twocolumn':    slideTwoColumn,
  'two column':   slideTwoColumn,
  'cta':          slideCTA,
  'quote':        slideQuote,
  'timeline':     slideTimeline,
  'comparison':   slideComparison,
};

// ── Public API ────────────────────────────────────────────────

/**
 * Render an outline JSON object to a .pptx Buffer.
 * @param {object} outline  - { title, author?, date?, slides: [{type, content}] }
 * @returns {Promise<Buffer>}
 */
async function renderDeck(outline) {
  const pptx    = new PptxGenJS();
  pptx.layout   = 'LAYOUT_WIDE';
  pptx.title    = outline.title  || 'Deck';
  pptx.author   = outline.author || 'Ultranative';
  pptx.subject  = outline.title  || '';

  for (const spec of (outline.slides || [])) {
    const key      = (spec.type || '').toLowerCase().trim();
    const renderer = RENDERERS[key];
    const slide    = pptx.addSlide();

    if (renderer) {
      renderer(slide, spec.content || {});
    } else {
      // Graceful fallback
      setBg(slide);
      slide.addText(`Unknown slide type: "${spec.type}"`, {
        x: 0.6, y: 3.3, w: W - 1.2, h: 0.8,
        fontFace: T.FONT, fontSize: 18, color: T.MUTED, align: 'center',
      });
    }
  }

  return pptx.write({ outputType: 'nodebuffer' });
}

module.exports = { renderDeck };
