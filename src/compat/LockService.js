class Lock {
  constructor() {
    this._locked = false;
  }

  tryLock(timeoutMs = 0) {
    if (!this._locked) {
      this._locked = true;
      return true;
    }
    const start = Date.now();
    while (this._locked) {
      if (Date.now() - start > timeoutMs) {
        return false;
      }
      Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 20);
    }
    this._locked = true;
    return true;
  }

  waitLock(timeoutMs = 10000) {
    const ok = this.tryLock(timeoutMs);
    if (!ok) {
      throw new Error('Could not obtain lock after ' + timeoutMs + 'ms');
    }
  }

  releaseLock() {
    this._locked = false;
  }

  hasLock() {
    return this._locked;
  }
}

const scriptLock = new Lock();
const userLock = new Lock();

class LockService {
  static getScriptLock() {
    return scriptLock;
  }

  static getUserLock() {
    return userLock;
  }
}

module.exports = { LockService, Lock };
