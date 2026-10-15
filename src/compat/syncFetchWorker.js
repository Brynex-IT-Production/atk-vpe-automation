const https = require('https');
const http = require('http');
const { URL } = require('url');

let input = '';
process.stdin.setEncoding('utf8');
process.stdin.on('data', chunk => { input += chunk; });
process.stdin.on('end', () => {
  try {
    const config = JSON.parse(input);
    const u = new URL(config.url);
    const client = u.protocol === 'https:' ? https : http;
    const method = (config.method || 'GET').toUpperCase();
    const headers = config.headers || {};
    const payload = config.payload;

    if (payload && !headers['Content-Length'] && !headers['content-length']) {
      headers['Content-Length'] = Buffer.byteLength(payload);
    }

    const req = client.request(u, { method, headers }, res => {
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => {
        const body = Buffer.concat(chunks).toString('utf8');
        const output = {
          status: res.statusCode,
          headers: res.headers,
          body
        };
        process.stdout.write(JSON.stringify(output));
        process.exit(0);
      });
    });

    req.on('error', err => {
      process.stderr.write(err.message);
      process.exit(1);
    });

    if (payload) {
      req.write(payload);
    }
    req.end();
  } catch (err) {
    process.stderr.write(err.message);
    process.exit(1);
  }
});
