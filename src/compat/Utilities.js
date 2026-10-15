const crypto = require('crypto');

const MONTH_NAMES_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_NAMES_FULL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAY_NAMES_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const DAY_NAMES_FULL = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

class Utilities {
  static getUuid() {
    return crypto.randomUUID();
  }

  static sleep(milliseconds) {
    if (milliseconds <= 0) return;
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, milliseconds);
  }

  static formatDate(date, timeZone, format) {
    if (!date) return '';
    const d = (date instanceof Date) ? date : new Date(date);
    if (isNaN(d.getTime())) return '';

    const tz = timeZone || 'Asia/Kolkata';

    // Format parts in target timezone using Intl
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: tz,
      year: 'numeric',
      month: 'numeric',
      day: 'numeric',
      hour: 'numeric',
      minute: 'numeric',
      second: 'numeric',
      hour12: false,
      weekday: 'short'
    });

    const parts = formatter.formatToParts(d);
    const map = {};
    for (const p of parts) {
      map[p.type] = p.value;
    }

    const year = parseInt(map.year, 10);
    const month = parseInt(map.month, 10);
    const day = parseInt(map.day, 10);
    const hour = parseInt(map.hour === '24' ? '0' : map.hour, 10);
    const minute = parseInt(map.minute, 10);
    const second = parseInt(map.second, 10);
    
    // Day of week index
    const weekdayShort = map.weekday;
    const dayOfWeek = DAY_NAMES_SHORT.indexOf(weekdayShort);

    let result = format;
    result = result.replace(/yyyy/g, String(year));
    result = result.replace(/yy/g, String(year).slice(-2));
    result = result.replace(/MMMM/g, MONTH_NAMES_FULL[month - 1]);
    result = result.replace(/MMM/g, MONTH_NAMES_SHORT[month - 1]);
    result = result.replace(/MM/g, String(month).padStart(2, '0'));
    result = result.replace(/\bM\b/g, String(month));
    result = result.replace(/dd/g, String(day).padStart(2, '0'));
    result = result.replace(/\bd\b/g, String(day));
    result = result.replace(/EEEE/g, DAY_NAMES_FULL[dayOfWeek >= 0 ? dayOfWeek : 0]);
    result = result.replace(/EEE/g, weekdayShort || '');
    result = result.replace(/HH/g, String(hour).padStart(2, '0'));
    result = result.replace(/\bH\b/g, String(hour));
    result = result.replace(/mm/g, String(minute).padStart(2, '0'));
    result = result.replace(/\bm\b/g, String(minute));
    result = result.replace(/ss/g, String(second).padStart(2, '0'));
    result = result.replace(/\bs\b/g, String(second));

    return result;
  }

  static computeDigest(algorithm, value, charset) {
    const alg = (algorithm === Utilities.DigestAlgorithm.MD5) ? 'md5' : 'sha256';
    const hash = crypto.createHash(alg).update(value, 'utf8').digest();
    // Return signed 8-bit integers (Java byte[] behavior in Google Apps Script)
    return Array.from(new Int8Array(hash.buffer, hash.byteOffset, hash.byteLength));
  }

  static computeHmacSha256Signature(value, key) {
    const hmac = crypto.createHmac('sha256', key).update(value).digest();
    return Array.from(new Int8Array(hmac.buffer, hmac.byteOffset, hmac.byteLength));
  }

  static base64Encode(data) {
    if (Array.isArray(data)) {
      return Buffer.from(new Uint8Array(data)).toString('base64');
    }
    return Buffer.from(String(data), 'utf8').toString('base64');
  }

  static base64Decode(str) {
    const buf = Buffer.from(str, 'base64');
    return Array.from(new Int8Array(buf.buffer, buf.byteOffset, buf.byteLength));
  }

  static base64EncodeWebSafe(data) {
    if (Array.isArray(data)) {
      return Buffer.from(new Uint8Array(data)).toString('base64url');
    }
    return Buffer.from(String(data), 'utf8').toString('base64url');
  }

  static base64DecodeWebSafe(str) {
    const buf = Buffer.from(str, 'base64url');
    return Array.from(new Int8Array(buf.buffer, buf.byteOffset, buf.byteLength));
  }

  static newBlob(data, contentType, name) {
    return {
      getBytes: () => Array.from(Buffer.from(data)),
      getDataAsString: () => String(data),
      getContentType: () => contentType || 'application/octet-stream',
      getName: () => name || 'blob'
    };
  }
}

Utilities.DigestAlgorithm = {
  SHA_256: 'SHA_256',
  MD5: 'MD5'
};

Utilities.Charset = {
  UTF_8: 'UTF_8'
};

module.exports = { Utilities };
