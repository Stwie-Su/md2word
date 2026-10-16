/*
 * zip.js — 极简 ZIP 写入器（仅存储压缩 method=0，Word 完全兼容）
 * 不依赖任何第三方库，浏览器与 Node 通用。
 */
(function (global) {
  'use strict';

  // CRC32 表
  let crcTable = null;
  function makeCRCTable() {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) {
        c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      }
      table[n] = c >>> 0;
    }
    return table;
  }
  function crc32(buf) {
    if (!crcTable) crcTable = makeCRCTable();
    let c = 0xffffffff;
    for (let i = 0; i < buf.length; i++) {
      c = (c >>> 8) ^ crcTable[(c ^ buf[i]) & 0xff];
    }
    return (c ^ 0xffffffff) >>> 0;
  }

  const enc = new TextEncoder();

  /**
   * 创建 ZIP 文件的二进制数据。
   * @param {Array<{name:string, data:Uint8Array}>} files
   * @returns {Uint8Array}
   */
  function createZip(files) {
    const now = new Date();
    const time =
      (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const date =
      ((now.getFullYear() - 1980) << 9) |
      ((now.getMonth() + 1) << 5) |
      now.getDate();

    const localParts = [];
    const central = [];
    let offset = 0;

    for (const f of files) {
      const nameBytes = enc.encode(f.name);
      const data = f.data;
      const crc = crc32(data);

      const lh = new Uint8Array(30 + nameBytes.length);
      const dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true); // version needed
      dv.setUint16(6, 0x0800, true); // flags: UTF-8 filename
      dv.setUint16(8, 0, true); // method: stored
      dv.setUint16(10, time, true);
      dv.setUint16(12, date, true);
      dv.setUint32(14, crc, true);
      dv.setUint32(18, data.length, true);
      dv.setUint32(22, data.length, true);
      dv.setUint16(26, nameBytes.length, true);
      dv.setUint16(28, 0, true);
      lh.set(nameBytes, 30);
      localParts.push(lh, data);

      const ch = new Uint8Array(46 + nameBytes.length);
      const cdv = new DataView(ch.buffer);
      cdv.setUint32(0, 0x02014b50, true);
      cdv.setUint16(4, 20, true); // version made by
      cdv.setUint16(6, 20, true); // version needed
      cdv.setUint16(8, 0x0800, true);
      cdv.setUint16(10, 0, true);
      cdv.setUint16(12, time, true);
      cdv.setUint16(14, date, true);
      cdv.setUint32(16, crc, true);
      cdv.setUint32(20, data.length, true);
      cdv.setUint32(24, data.length, true);
      cdv.setUint16(28, nameBytes.length, true);
      cdv.setUint16(30, 0, true);
      cdv.setUint16(32, 0, true);
      cdv.setUint16(34, 0, true);
      cdv.setUint16(36, 0, true);
      cdv.setUint32(38, 0, true);
      cdv.setUint32(42, offset, true);
      ch.set(nameBytes, 46);
      central.push(ch);

      offset += lh.length + data.length;
    }

    const centralSize = central.reduce((s, c) => s + c.length, 0);
    const cdStart = offset;

    const end = new Uint8Array(22);
    const edv = new DataView(end.buffer);
    edv.setUint32(0, 0x06054b50, true);
    edv.setUint16(4, 0, true);
    edv.setUint16(6, 0, true);
    edv.setUint16(8, files.length, true);
    edv.setUint16(10, files.length, true);
    edv.setUint32(12, centralSize, true);
    edv.setUint32(16, cdStart, true);
    edv.setUint16(20, 0, true);

    const all = [...localParts, ...central, end];
    const total = all.reduce((s, a) => s + a.length, 0);
    const out = new Uint8Array(total);
    let pos = 0;
    for (const a of all) {
      out.set(a, pos);
      pos += a.length;
    }
    return out;
  }

  const API = { createZip, crc32 };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  else global.ZipUtil = API;
})(typeof self !== 'undefined' ? self : this);
