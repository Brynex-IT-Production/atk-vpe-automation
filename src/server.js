require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

// Initialize Google Apps Script Compatibility Layer
require('./compat');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Request Logger
app.use((req, res, next) => {
  const start = Date.now();
  res.on('finish', () => {
    const duration = Date.now() - start;
    if (process.env.NODE_ENV !== 'test') {
      console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl} - ${res.statusCode} (${duration}ms)`);
    }
  });
  next();
});

// Root / Health check
app.get('/', (req, res) => {
  res.json({
    status: 'online',
    service: 'ATK VPE Automation Engine',
    platform: 'Node.js / Express (MERN Stack)',
    version: '1.0.0',
    system_mode: globalThis.configValue_ ? globalThis.configValue_('SYSTEM_MODE', 'DEV') : 'DEV',
    timezone: globalThis.configValue_ ? globalThis.configValue_('TIMEZONE', 'Asia/Kolkata') : 'Asia/Kolkata'
  });
});

app.get('/health', (req, res) => {
  res.json({ ok: true, uptime: process.uptime(), timestamp: new Date().toISOString() });
});

// Meta Webhook Verification (doGet)
app.get('/webhook', (req, res) => {
  try {
    const e = {
      parameter: req.query || {}
    };

    // Special handler: Programme review view
    if (e.parameter.view === 'programme' && e.parameter.token) {
      const output = globalThis.renderProgrammeReviewPage_(String(e.parameter.token || ''));
      return res.type('text/html').send(output.getContent());
    }

    const output = globalThis.doGet(e);
    res.type(output.getMimeType()).send(output.getContent());
  } catch (err) {
    console.error('Error in GET /webhook:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Meta Webhook & General Event Ingestion (doPost)
app.post('/webhook', (req, res) => {
  try {
    const bodyStr = typeof req.body === 'string' ? req.body : JSON.stringify(req.body);
    const e = {
      postData: {
        contents: bodyStr
      }
    };

    const output = globalThis.doPost(e);
    res.type(output.getMimeType()).send(output.getContent());
  } catch (err) {
    console.error('Error in POST /webhook:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// Dedicated Web API v1 endpoint (Flutter / Web clients)
app.post('/api/v1', (req, res) => {
  try {
    const payload = Object.assign({ api_version: 'v1' }, req.body);
    const result = globalThis.apiHandleRequest_(payload);
    res.status(result.status || 200).json({
      data: result.data || null,
      error: result.error || null
    });
  } catch (err) {
    console.error('Error in POST /api/v1:', err);
    res.status(500).json({ data: null, error: { message: err.message } });
  }
});

// Programme Review Web Page
app.get('/programme-review', (req, res) => {
  try {
    const token = String(req.query.token || '');
    if (!token) {
      return res.status(400).send('<h3>ATK Programme Planner</h3><p>Missing review token.</p>');
    }

    const output = globalThis.renderProgrammeReviewPage_(token);
    let html = output.getContent();

    // Inject a client-side bridge for google.script.run so ProgrammeReview.html works directly in any browser
    const bridgeScript = `
    <script>
      window.google = window.google || {};
      window.google.script = window.google.script || {};
      window.google.script.run = {
        _successHandler: null,
        _failureHandler: null,
        withSuccessHandler: function(fn) { this._successHandler = fn; return this; },
        withFailureHandler: function(fn) { this._failureHandler = fn; return this; },
        getProgrammeReviewData: function(token) {
          var self = this;
          fetch('/api/programme-review/data?token=' + encodeURIComponent(token))
            .then(function(res) { return res.json(); })
            .then(function(data) {
              if (data.error && self._failureHandler) self._failureHandler(data.error);
              else if (self._successHandler) self._successHandler(data);
            })
            .catch(function(err) { if (self._failureHandler) self._failureHandler(err); });
        },
        saveProgrammeReview: function(token, actions) {
          var self = this;
          fetch('/api/programme-review/save', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ token: token, actions: actions })
          })
          .then(function(res) { return res.json(); })
          .then(function(data) {
            if (data.error && self._failureHandler) self._failureHandler(data.error);
            else if (self._successHandler) self._successHandler(data);
          })
          .catch(function(err) { if (self._failureHandler) self._failureHandler(err); });
        }
      };
    </script>
    `;

    html = html.replace('</head>', `${bridgeScript}\n</head>`);
    res.type('text/html').send(html);
  } catch (err) {
    console.error('Error in GET /programme-review:', err);
    res.status(500).send('Error rendering review: ' + err.message);
  }
});

// Programme Review API Data
app.get('/api/programme-review/data', (req, res) => {
  try {
    const token = String(req.query.token || '');
    const data = globalThis.getProgrammeReviewData(token);
    res.json(data);
  } catch (err) {
    res.status(400).json({ error: { message: err.message } });
  }
});

// Programme Review API Save
app.post('/api/programme-review/save', (req, res) => {
  try {
    const { token, actions } = req.body;
    const result = globalThis.saveProgrammeReview(token, actions);
    res.json(result || { ok: true, messages: [] });
  } catch (err) {
    res.status(400).json({ error: { message: err.message } });
  }
});

// Dev Master Regression Suite
app.get('/api/dev/regression', (req, res) => {
  try {
    const verifyToken = req.query.verify_token;
    const expectedToken = process.env.VERIFY_TOKEN || 'ATK_VPE_VERIFY_2026_X7K9M4P2';

    if (verifyToken && verifyToken !== expectedToken) {
      return res.status(403).json({ ok: false, error: 'Unauthorized verify token' });
    }

    const report = globalThis.apiDevMasterRegressionSuite();
    res.json(report);
  } catch (err) {
    console.error('Regression suite error:', err);
    res.status(500).json({ ok: false, error: err.message, stack: err.stack });
  }
});

// 404 Handler
app.use((req, res) => {
  res.status(404).json({ error: 'Endpoint not found', path: req.path });
});

// Export app for testing and server startup
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`\n=================================================`);
    console.log(`  🚀 ATK VPE Automation Node.js Server Started   `);
    console.log(`  Port: http://localhost:${PORT}                 `);
    console.log(`  Health: http://localhost:${PORT}/health        `);
    console.log(`  Webhook: http://localhost:${PORT}/webhook      `);
    console.log(`  API v1: http://localhost:${PORT}/api/v1        `);
    console.log(`=================================================\n`);
  });
}

module.exports = app;
