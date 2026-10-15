const fs = require('fs');
const path = require('path');

class HtmlOutput {
  constructor(html = '') {
    this._html = String(html || '');
    this._title = '';
    this._metaTags = [];
  }

  setTitle(title) {
    this._title = title;
    return this;
  }

  getTitle() {
    return this._title;
  }

  addMetaTag(name, content) {
    this._metaTags.push({ name, content });
    return this;
  }

  getContent() {
    let out = this._html;
    if (this._title && !out.includes('<title>')) {
      out = out.replace(/<head>/i, `<head>\n  <title>${this._title}</title>`);
    }
    if (this._metaTags.length > 0) {
      const metaStr = this._metaTags.map(m => `  <meta name="${m.name}" content="${m.content}">`).join('\n');
      out = out.replace(/<head>/i, `<head>\n${metaStr}`);
    }
    return out;
  }
}

class HtmlTemplate {
  constructor(content) {
    this._raw = content;
  }

  evaluate() {
    // Replace scriptlets like <?!= JSON.stringify(token) ?> with evaluated properties
    let evaluated = this._raw;
    
    // Evaluate <?!= expression ?>
    evaluated = evaluated.replace(/<\?!=\s*([\s\S]*?)\s*\?>/g, (match, expr) => {
      try {
        const fn = new Function(...Object.keys(this), `return (${expr});`);
        return String(fn(...Object.values(this)));
      } catch (e) {
        return '';
      }
    });

    // Evaluate <?= expression ?> (escaped)
    evaluated = evaluated.replace(/<\?=\s*([\s\S]*?)\s*\?>/g, (match, expr) => {
      try {
        const fn = new Function(...Object.keys(this), `return (${expr});`);
        const val = String(fn(...Object.values(this)));
        return val.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      } catch (e) {
        return '';
      }
    });

    return new HtmlOutput(evaluated);
  }
}

class HtmlService {
  static createHtmlOutput(content) {
    return new HtmlOutput(content);
  }

  static createHtmlOutputFromFile(filename) {
    const file = filename.endsWith('.html') ? filename : `${filename}.html`;
    const fullPath = path.resolve(process.cwd(), file);
    const content = fs.existsSync(fullPath) ? fs.readFileSync(fullPath, 'utf8') : '';
    return new HtmlOutput(content);
  }

  static createTemplateFromFile(filename) {
    const file = filename.endsWith('.html') ? filename : `${filename}.html`;
    const fullPath = path.resolve(process.cwd(), file);
    const content = fs.existsSync(fullPath) ? fs.readFileSync(fullPath, 'utf8') : '';
    return new HtmlTemplate(content);
  }
}

module.exports = { HtmlService, HtmlOutput, HtmlTemplate };
