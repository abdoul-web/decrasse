/* Décrasse — lecture et écriture d'archives ZIP (format conteneur des fichiers .xlsx).
   Utilise les flux natifs du navigateur (CompressionStream / DecompressionStream). */
(function (D) {
  'use strict';

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(u8) {
    let c = 0xffffffff;
    for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  async function pipe(u8, stream) {
    const out = new Response(new Blob([u8]).stream().pipeThrough(stream));
    return new Uint8Array(await out.arrayBuffer());
  }
  const inflateRaw = (u8) => pipe(u8, new DecompressionStream('deflate-raw'));
  const deflateRaw = (u8) => pipe(u8, new CompressionStream('deflate-raw'));

  /** Lit le répertoire central d'une archive ZIP. Retourne { names, read(name) → Promise<Uint8Array> }. */
  function readZip(input) {
    const u8 = input instanceof Uint8Array ? input : new Uint8Array(input);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("Ce fichier n'est pas un classeur Excel valide (archive illisible).");
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const dec = new TextDecoder();
    const entries = new Map();
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Archive Excel endommagée.');
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true);
      const xlen = dv.getUint16(p + 30, true);
      const clen = dv.getUint16(p + 32, true);
      const local = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
      entries.set(name.replace(/^\/+/, ''), { method, csize, local });
      p += 46 + nlen + xlen + clen;
    }
    return {
      names: [...entries.keys()],
      has: (n) => entries.has(n),
      async read(name) {
        const e = entries.get(name);
        if (!e) return null;
        const lnlen = dv.getUint16(e.local + 26, true);
        const lxlen = dv.getUint16(e.local + 28, true);
        const start = e.local + 30 + lnlen + lxlen;
        const data = u8.subarray(start, start + e.csize);
        if (e.method === 0) return data;
        if (e.method === 8) return inflateRaw(data);
        throw new Error('Méthode de compression non prise en charge (' + e.method + ').');
      },
      async text(name) {
        const b = await this.read(name);
        return b ? dec.decode(b) : null;
      },
    };
  }

  /** Construit une archive ZIP. entries : [{ name, data: Uint8Array | string }] */
  async function writeZip(entries) {
    const enc = new TextEncoder();
    const parts = [];
    const central = [];
    let offset = 0;
    // Date DOS fixe : 1er janvier 2026, 00:00
    const dosTime = 0, dosDate = ((2026 - 1980) << 9) | (1 << 5) | 1;
    for (const e of entries) {
      const raw = typeof e.data === 'string' ? enc.encode(e.data) : e.data;
      const name = enc.encode(e.name);
      const crc = crc32(raw);
      let data = raw, method = 0;
      if (raw.length > 64) {
        const z = await deflateRaw(raw);
        if (z.length < raw.length) { data = z; method = 8; }
      }
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
      lh.setUint16(8, method, true); lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, raw.length, true);
      lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), name, data);

      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true); ch.setUint16(10, method, true); ch.setUint16(12, dosTime, true);
      ch.setUint16(14, dosDate, true); ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true);
      ch.setUint32(24, raw.length, true); ch.setUint16(28, name.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + data.length;
    }
    const csize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
    end.setUint32(12, csize, true); end.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((a, b) => a + b.length, 0));
    let p = 0;
    for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }

  D.zip = { crc32, readZip, writeZip, inflateRaw, deflateRaw };
})(globalThis.Decrasse || (globalThis.Decrasse = {}));
