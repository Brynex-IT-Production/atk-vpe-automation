const fs = require('fs');
const path = require('path');

class PropertiesStore {
  constructor(filePath) {
    this.filePath = filePath;
    this.data = {};
    this._load();
  }

  _load() {
    try {
      if (fs.existsSync(this.filePath)) {
        const raw = fs.readFileSync(this.filePath, 'utf8');
        this.data = JSON.parse(raw);
      }
    } catch (e) {
      this.data = {};
    }
  }

  _save() {
    try {
      const dir = path.dirname(this.filePath);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
      fs.writeFileSync(this.filePath, JSON.stringify(this.data, null, 2), 'utf8');
    } catch (e) {
      // ignore write errors
    }
  }

  getProperty(key) {
    if (this.data[key] !== undefined) {
      return String(this.data[key]);
    }
    // Check direct process.env
    if (process.env[key] !== undefined) {
      return String(process.env[key]);
    }
    // Aliases
    if (key === 'META_ACCESS_TOKEN' && process.env.WA_TOKEN !== undefined) {
      return String(process.env.WA_TOKEN);
    }
    if (key === 'WA_TOKEN' && process.env.META_ACCESS_TOKEN !== undefined) {
      return String(process.env.META_ACCESS_TOKEN);
    }
    return null;
  }

  setProperty(key, value) {
    this.data[key] = String(value);
    this._save();
  }

  getProperties() {
    // Merge process.env with local data
    const merged = {};
    for (const [k, v] of Object.entries(process.env)) {
      if (typeof v === 'string') merged[k] = v;
    }
    for (const [k, v] of Object.entries(this.data)) {
      merged[k] = String(v);
    }
    return merged;
  }

  setProperties(props, deleteAllOthers = false) {
    if (deleteAllOthers) {
      this.data = {};
    }
    for (const [k, v] of Object.entries(props || {})) {
      this.data[k] = String(v);
    }
    this._save();
  }

  deleteProperty(key) {
    delete this.data[key];
    this._save();
  }

  deleteAllProperties() {
    this.data = {};
    this._save();
  }
}

const dataDir = path.resolve(__dirname, '../../data');
const scriptStore = new PropertiesStore(path.join(dataDir, 'script_properties.json'));
const userStore = new PropertiesStore(path.join(dataDir, 'user_properties.json'));

class PropertiesService {
  static getScriptProperties() {
    return scriptStore;
  }

  static getUserProperties() {
    return userStore;
  }
}

module.exports = { PropertiesService };
