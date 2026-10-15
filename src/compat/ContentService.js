class TextOutput {
  constructor(content = '') {
    this._content = String(content || '');
    this._mimeType = 'text/plain';
  }

  setMimeType(mimeType) {
    this._mimeType = mimeType;
    return this;
  }

  getContent() {
    return this._content;
  }

  getMimeType() {
    return this._mimeType;
  }
}

class ContentService {
  static createTextOutput(content) {
    return new TextOutput(content);
  }
}

ContentService.MimeType = {
  TEXT: 'text/plain',
  JSON: 'application/json',
  JAVASCRIPT: 'application/javascript',
  XML: 'application/xml',
  CSV: 'text/csv'
};

module.exports = { ContentService, TextOutput };
