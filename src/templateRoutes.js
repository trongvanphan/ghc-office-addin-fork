/**
 * Template Library API routes for PowerPoint slide templates.
 * Mounted at /api/templates on both dev and production servers.
 *
 * Storage layout:
 *   ~/.copilot-office-addin/templates/{id}.pptx        — PPTX binary
 *   ~/.copilot-office-addin/templates/{id}.json        — Metadata
 *   ~/.copilot-office-addin/templates/{id}/thumb-{n}.png — Thumbnails
 */
const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');
const crypto = require('crypto');

// ---- helpers ----------------------------------------------------------------

const TEMPLATES_DIR = path.join(os.homedir(), '.copilot-office-addin', 'templates');
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function ensureDir() {
  fs.mkdirSync(TEMPLATES_DIR, { recursive: true });
}

function isValidId(id) {
  return typeof id === 'string' && UUID_RE.test(id);
}

function metaPath(id) {
  return path.join(TEMPLATES_DIR, `${id}.json`);
}

function pptxPath(id) {
  return path.join(TEMPLATES_DIR, `${id}.pptx`);
}

function thumbPath(id, slideIndex) {
  const thumbDir = path.join(TEMPLATES_DIR, id);
  fs.mkdirSync(thumbDir, { recursive: true });
  return path.join(thumbDir, `thumb-${slideIndex}.png`);
}

function listTemplates() {
  ensureDir();
  return fs.readdirSync(TEMPLATES_DIR)
    .filter(f => f.endsWith('.json'))
    .map(f => {
      try {
        return JSON.parse(fs.readFileSync(path.join(TEMPLATES_DIR, f), 'utf8'));
      } catch {
        return null;
      }
    })
    .filter(Boolean)
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
}

// ---- router -----------------------------------------------------------------

function createTemplateRouter() {
  const router = express.Router();
  router.use(express.json({ limit: '200mb' }));

  // GET /api/templates — list all templates (metadata only, no binary)
  router.get('/', (req, res) => {
    try {
      res.json(listTemplates());
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // POST /api/templates/upload — upload new template
  // Body: { name: string, data: base64string, slideCount: number }
  router.post('/upload', (req, res) => {
    try {
      ensureDir();
      const { name, data, slideCount } = req.body || {};

      if (!name || typeof name !== 'string' || name.trim().length === 0) {
        return res.status(400).json({ error: 'Missing or invalid name' });
      }
      if (!data || typeof data !== 'string') {
        return res.status(400).json({ error: 'Missing data (base64 PPTX)' });
      }
      if (typeof slideCount !== 'number' || slideCount < 1 || slideCount > 500) {
        return res.status(400).json({ error: 'Invalid slideCount' });
      }

      // Validate base64
      const buffer = Buffer.from(data, 'base64');
      // PPTX magic bytes: PK (zip) starts with 0x50 0x4B
      if (buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
        return res.status(400).json({ error: 'Uploaded file does not appear to be a valid PPTX file' });
      }

      const id = crypto.randomUUID();
      const safeName = name.trim().substring(0, 100);

      const metadata = {
        id,
        name: safeName,
        slideCount,
        slides: Array.from({ length: slideCount }, (_, i) => ({
          index: i,
          type: 'other',
          label: '',
        })),
        createdAt: new Date().toISOString(),
      };

      fs.writeFileSync(pptxPath(id), buffer);
      fs.writeFileSync(metaPath(id), JSON.stringify(metadata, null, 2));

      res.status(201).json(metadata);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // GET /api/templates/:id — get metadata + base64 binary
  router.get('/:id', (req, res) => {
    try {
      const { id } = req.params;
      if (!isValidId(id)) return res.status(400).json({ error: 'Invalid template id' });

      const mPath = metaPath(id);
      const fPath = pptxPath(id);
      if (!fs.existsSync(mPath) || !fs.existsSync(fPath)) {
        return res.status(404).json({ error: 'Template not found' });
      }

      const metadata = JSON.parse(fs.readFileSync(mPath, 'utf8'));
      const data = fs.readFileSync(fPath).toString('base64');

      res.json({ ...metadata, data });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // PUT /api/templates/:id/tags — update slide type tags
  // Body: { slides: [{ index, type, label }] }
  router.put('/:id/tags', (req, res) => {
    try {
      const { id } = req.params;
      if (!isValidId(id)) return res.status(400).json({ error: 'Invalid template id' });

      const mPath = metaPath(id);
      if (!fs.existsSync(mPath)) return res.status(404).json({ error: 'Template not found' });

      const { slides } = req.body || {};
      if (!Array.isArray(slides)) return res.status(400).json({ error: 'slides must be an array' });

      const VALID_TYPES = [
        'intro', 'agenda', 'content', 'two_column', 'image_text',
        'chart', 'table', 'qa', 'thank_you', 'other',
      ];

      const sanitized = slides.map(s => ({
        index: Number(s.index),
        type: VALID_TYPES.includes(s.type) ? s.type : 'other',
        label: typeof s.label === 'string' ? s.label.substring(0, 100) : '',
      }));

      const metadata = JSON.parse(fs.readFileSync(mPath, 'utf8'));
      metadata.slides = sanitized;
      fs.writeFileSync(mPath, JSON.stringify(metadata, null, 2));

      res.json(metadata);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // DELETE /api/templates/:id — delete template
  router.delete('/:id', (req, res) => {
    try {
      const { id } = req.params;
      if (!isValidId(id)) return res.status(400).json({ error: 'Invalid template id' });

      const mPath = metaPath(id);
      const fPath = pptxPath(id);
      const thumbDir = path.join(TEMPLATES_DIR, id);

      if (!fs.existsSync(mPath)) return res.status(404).json({ error: 'Template not found' });

      if (fs.existsSync(fPath)) fs.unlinkSync(fPath);
      if (fs.existsSync(mPath)) fs.unlinkSync(mPath);
      if (fs.existsSync(thumbDir)) fs.rmSync(thumbDir, { recursive: true, force: true });

      res.sendStatus(204);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // POST /api/templates/:id/thumbnails — save slide thumbnails from Office.js capture
  // Body: { thumbnails: [{ slideIndex: number, imageData: base64png }] }
  router.post('/:id/thumbnails', (req, res) => {
    try {
      const { id } = req.params;
      if (!isValidId(id)) return res.status(400).json({ error: 'Invalid template id' });
      if (!fs.existsSync(metaPath(id))) return res.status(404).json({ error: 'Template not found' });

      const { thumbnails } = req.body || {};
      if (!Array.isArray(thumbnails)) return res.status(400).json({ error: 'thumbnails must be an array' });

      for (const thumb of thumbnails) {
        const idx = Number(thumb.slideIndex);
        if (!Number.isFinite(idx) || idx < 0 || idx > 499) continue;
        if (typeof thumb.imageData !== 'string') continue;

        const png = Buffer.from(thumb.imageData, 'base64');
        // Validate PNG magic bytes: 0x89 0x50 0x4E 0x47
        if (png[0] !== 0x89 || png[1] !== 0x50 || png[2] !== 0x4e || png[3] !== 0x47) continue;

        fs.writeFileSync(thumbPath(id, idx), png);
      }

      res.sendStatus(204);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // GET /api/templates/:id/slide-thumbnail/:slideIndex — serve thumbnail
  router.get('/:id/slide-thumbnail/:slideIndex', (req, res) => {
    try {
      const { id, slideIndex } = req.params;
      if (!isValidId(id)) return res.status(400).json({ error: 'Invalid template id' });

      const idx = Number(slideIndex);
      if (!Number.isFinite(idx) || idx < 0 || idx > 499) {
        return res.status(400).json({ error: 'Invalid slideIndex' });
      }

      const tPath = path.join(TEMPLATES_DIR, id, `thumb-${idx}.png`);
      if (!fs.existsSync(tPath)) {
        return res.status(404).json({ error: 'Thumbnail not found' });
      }

      res.setHeader('Content-Type', 'image/png');
      res.setHeader('Cache-Control', 'public, max-age=86400');
      fs.createReadStream(tPath).pipe(res);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  return router;
}

module.exports = { createTemplateRouter };
