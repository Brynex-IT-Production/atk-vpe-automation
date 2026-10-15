class Cache {
  constructor() {
    this._store = new Map();
  }

  get(key) {
    const item = this._store.get(key);
    if (!item) return null;
    if (Date.now() > item.expiresAt) {
      this._store.delete(key);
      return null;
    }
    return item.value;
  }

  put(key, value, expirationInSeconds = 600) {
    const expiresAt = Date.now() + (expirationInSeconds * 1000);
    this._store.set(key, { value: String(value), expiresAt });
  }

  remove(key) {
    this._store.delete(key);
  }

  getAll(keys) {
    const res = {};
    for (const k of keys) {
      const v = this.get(k);
      if (v !== null) res[k] = v;
    }
    return res;
  }

  putAll(values, expirationInSeconds = 600) {
    for (const [k, v] of Object.entries(values || {})) {
      this.put(k, v, expirationInSeconds);
    }
  }

  removeAll(keys) {
    for (const k of keys) {
      this.remove(k);
    }
  }
}

const scriptCache = new Cache();

class CacheService {
  static getScriptCache() {
    return scriptCache;
  }
}

module.exports = { CacheService, Cache };
