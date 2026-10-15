const { spawnSync } = require('child_process');
const path = require('path');

const WORKER_SCRIPT = path.resolve(__dirname, 'syncFetchWorker.js');

class HTTPResponse {
  constructor(status, headers, body) {
    this._status = Number(status || 200);
    this._headers = headers || {};
    this._body = String(body || '');
  }

  getResponseCode() {
    return this._status;
  }

  getContentText(optCharset) {
    return this._body;
  }

  getHeaders() {
    return this._headers;
  }

  getBlob() {
    const body = this._body;
    return {
      getBytes: () => Array.from(Buffer.from(body)),
      getDataAsString: () => body,
      getContentType: () => this._headers['content-type'] || 'text/plain'
    };
  }
}

class UrlFetchApp {
  static fetch(url, options = {}) {
    const method = (options.method || 'get').toLowerCase();
    const headers = Object.assign({}, options.headers || {});
    if (options.contentType) {
      headers['Content-Type'] = options.contentType;
    }

    const payload = options.payload !== undefined
      ? (typeof options.payload === 'object' && !(options.payload instanceof Buffer)
          ? JSON.stringify(options.payload)
          : String(options.payload))
      : undefined;

    const requestPayload = JSON.stringify({
      url,
      method,
      headers,
      payload
    });

    const res = spawnSync(process.execPath, [WORKER_SCRIPT], {
      input: requestPayload,
      encoding: 'utf8',
      maxBuffer: 20 * 1024 * 1024
    });

    if (res.status !== 0) {
      const errMsg = (res.stderr || '').trim() || 'UrlFetchApp error';
      if (!options.muteHttpExceptions) {
        throw new Error(errMsg);
      }
      return new HTTPResponse(500, {}, JSON.stringify({ error: errMsg }));
    }

    let parsed;
    try {
      parsed = JSON.parse(res.stdout);
    } catch (e) {
      if (!options.muteHttpExceptions) {
        throw new Error('Failed to parse UrlFetchApp response: ' + res.stdout);
      }
      return new HTTPResponse(500, {}, res.stdout);
    }

    if (parsed.status >= 400 && !options.muteHttpExceptions) {
      throw new Error(`Request failed for ${url} returned code ${parsed.status}. Server response: ${parsed.body}`);
    }

    return new HTTPResponse(parsed.status, parsed.headers, parsed.body);
  }
}

module.exports = { UrlFetchApp, HTTPResponse };
