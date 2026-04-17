'use strict';
const express    = require('express');
const cors       = require('cors');
const { renderDeck } = require('./render');

const app = express();
app.use(cors());
app.use(express.json({ limit: '2mb' }));

// ── Health check ──────────────────────────────────────────────
app.get('/health', (_req, res) => res.json({ status: 'ok' }));

// ── Render endpoint ───────────────────────────────────────────
app.post('/render', async (req, res) => {
  const outline = req.body;

  if (!outline || !Array.isArray(outline.slides) || outline.slides.length === 0) {
    return res.status(400).json({ error: 'Request body must include a non-empty "slides" array.' });
  }

  try {
    const buffer   = await renderDeck(outline);
    const safeName = (outline.title || 'deck')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-|-$/g, '');
    const filename = `${safeName}.pptx`;

    res.setHeader('Content-Type',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);
    res.send(buffer);
  } catch (err) {
    console.error('[render error]', err);
    res.status(500).json({ error: err.message });
  }
});

// ── Start ─────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Deck renderer listening on :${PORT}`);
  console.log(`  POST /render  — accepts outline JSON, returns .pptx`);
  console.log(`  GET  /health  — liveness check`);
});
