// ─────────────────────────────────────────────────────────────────────────────
// ziplib.js — Self-contained ZIP read/write for .docx files
// Implements the subset of ZIP spec needed to read and rewrite .docx files.
// No external dependencies. Drop this file in your repo and it just works.
// ─────────────────────────────────────────────────────────────────────────────

const ZipLib = (() => {

  // ── Read a 2 or 4 byte little-endian integer from a DataView ─────────────
  function r16(v, o) { return v.getUint16(o, true); }
  function r32(v, o) { return v.getUint32(o, true); }

  // ── Write little-endian integers into a byte array ───────────────────────
  function w16(a, o, n) { a[o]=n&0xFF; a[o+1]=(n>>8)&0xFF; }
  function w32(a, o, n) { a[o]=n&0xFF; a[o+1]=(n>>8)&0xFF; a[o+2]=(n>>16)&0xFF; a[o+3]=(n>>24)&0xFF; }

  // ── CRC-32 ────────────────────────────────────────────────────────────────
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let i=0;i<256;i++) {
      let c=i;
      for (let j=0;j<8;j++) c = (c&1) ? (0xEDB88320^(c>>>1)) : (c>>>1);
      t[i]=c;
    }
    return t;
  })();

  function crc32(data) {
    let c = 0xFFFFFFFF;
    for (let i=0; i<data.length; i++) c = CRC_TABLE[(c^data[i])&0xFF] ^ (c>>>8);
    return (c^0xFFFFFFFF)>>>0;
  }

  // ── UTF-8 encode/decode ───────────────────────────────────────────────────
  const enc = new TextEncoder();
  const dec = new TextDecoder('utf-8');

  function toBytes(str) { return enc.encode(str); }
  function toString(bytes) { return dec.decode(bytes); }

  // ── Deflate compression using CompressionStream (built into modern browsers)
  async function deflate(data) {
    const cs = new CompressionStream('deflate-raw');
    const writer = cs.writable.getWriter();
    writer.write(data);
    writer.close();
    const chunks = [];
    const reader = cs.readable.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const total = chunks.reduce((s, c) => s + c.length, 0);
    const out = new Uint8Array(total);
    let pos = 0;
    for (const chunk of chunks) { out.set(chunk, pos); pos += chunk.length; }
    return out;
  }

  // ── Inflate decompression ─────────────────────────────────────────────────
  async function inflate(data) {
    const ds = new DecompressionStream('deflate-raw');
    const writer = ds.writable.getWriter();
    writer.write(data);
    writer.close();
    const chunks = [];
    const reader = ds.readable.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const total = chunks.reduce((s, c) => s + c.length, 0);
    const out = new Uint8Array(total);
    let pos = 0;
    for (const chunk of chunks) { out.set(chunk, pos); pos += chunk.length; }
    return out;
  }

  // ── Parse a ZIP archive ───────────────────────────────────────────────────
  async function readZip(arrayBuffer) {
    const buf  = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);
    const files = {};

    // Find End of Central Directory record (scan from end)
    let eocdPos = -1;
    for (let i = buf.length - 22; i >= 0; i--) {
      if (view.getUint32(i, true) === 0x06054B50) { eocdPos = i; break; }
    }
    if (eocdPos < 0) throw new Error('Not a valid ZIP file');

    const cdOffset = r32(view, eocdPos + 16);
    const cdCount  = r16(view, eocdPos + 8);

    // Parse central directory
    let pos = cdOffset;
    for (let i = 0; i < cdCount; i++) {
      if (r32(view, pos) !== 0x02014B50) break;
      const method     = r16(view, pos + 10);
      const crc        = r32(view, pos + 16);
      const compressed = r32(view, pos + 20);
      const uncompressed = r32(view, pos + 24);
      const nameLen    = r16(view, pos + 28);
      const extraLen   = r16(view, pos + 30);
      const commentLen = r16(view, pos + 32);
      const localOffset= r32(view, pos + 42);
      const name = toString(buf.slice(pos + 46, pos + 46 + nameLen));

      // Read local file header to find actual data offset
      const localView  = new DataView(arrayBuffer, localOffset);
      const localNameLen  = r16(localView, 26);
      const localExtraLen = r16(localView, 28);
      const dataOffset = localOffset + 30 + localNameLen + localExtraLen;
      const compressedData = buf.slice(dataOffset, dataOffset + compressed);

      files[name] = { method, crc, compressed, uncompressed, data: compressedData };
      pos += 46 + nameLen + extraLen + commentLen;
    }

    return files;
  }

  // ── Get a file from the parsed zip as a string ────────────────────────────
  async function getFileText(files, name) {
    const f = files[name];
    if (!f) throw new Error(`File not found in zip: ${name}`);
    if (f.method === 0) {
      // Stored (no compression)
      return toString(f.data);
    } else if (f.method === 8) {
      // Deflated
      const decompressed = await inflate(f.data);
      return toString(decompressed);
    }
    throw new Error(`Unsupported compression method: ${f.method}`);
  }

  // ── Write a new ZIP archive ───────────────────────────────────────────────
  async function writeZip(files) {
    // files = { name: string|Uint8Array, ... }
    const localHeaders = [];
    const centralDir   = [];
    let offset = 0;

    for (const [name, content] of Object.entries(files)) {
      const nameBytes = toBytes(name);
      const rawData   = typeof content === 'string' ? toBytes(content) : content;

      // Compress with deflate
      const compressed = await deflate(rawData);
      // Only use compression if it actually makes it smaller
      const useCompression = compressed.length < rawData.length;
      const finalData = useCompression ? compressed : rawData;
      const method    = useCompression ? 8 : 0;
      const crc       = crc32(rawData);

      // Local file header (30 bytes + name)
      const local = new Uint8Array(30 + nameBytes.length);
      w32(local, 0,  0x04034B50); // signature
      w16(local, 4,  20);          // version needed
      w16(local, 6,  0);           // flags
      w16(local, 8,  method);
      w16(local, 10, 0); w16(local, 12, 0); // mod time/date
      w32(local, 14, crc);
      w32(local, 18, finalData.length);
      w32(local, 22, rawData.length);
      w16(local, 26, nameBytes.length);
      w16(local, 28, 0); // extra length
      local.set(nameBytes, 30);

      // Central directory entry (46 bytes + name)
      const central = new Uint8Array(46 + nameBytes.length);
      w32(central, 0,  0x02014B50); // signature
      w16(central, 4,  20);          // version made by
      w16(central, 6,  20);          // version needed
      w16(central, 8,  0);           // flags
      w16(central, 10, method);
      w16(central, 12, 0); w16(central, 14, 0); // mod time/date
      w32(central, 16, crc);
      w32(central, 20, finalData.length);
      w32(central, 24, rawData.length);
      w16(central, 28, nameBytes.length);
      w16(central, 30, 0); // extra
      w16(central, 32, 0); // comment
      w16(central, 34, 0); // disk start
      w16(central, 36, 0); // internal attr
      w32(central, 38, 0); // external attr
      w32(central, 42, offset); // local header offset
      central.set(nameBytes, 46);

      localHeaders.push(local);
      localHeaders.push(finalData);
      centralDir.push(central);
      offset += local.length + finalData.length;
    }

    // End of central directory
    const cdSize = centralDir.reduce((s, c) => s + c.length, 0);
    const eocd = new Uint8Array(22);
    w32(eocd, 0,  0x06054B50);
    w16(eocd, 4,  0); w16(eocd, 6, 0);
    w16(eocd, 8,  centralDir.length);
    w16(eocd, 10, centralDir.length);
    w32(eocd, 12, cdSize);
    w32(eocd, 16, offset);
    w16(eocd, 20, 0);

    // Concatenate everything
    const allParts = [...localHeaders, ...centralDir, eocd];
    const total = allParts.reduce((s, p) => s + p.length, 0);
    const out = new Uint8Array(total);
    let pos = 0;
    for (const part of allParts) { out.set(part, pos); pos += part.length; }
    return out;
  }

  // ── Get a file as raw bytes (for passthrough of non-XML files) ──────────────
  async function getFileAsBytes(files, name) {
    const f = files[name];
    if (!f) throw new Error('File not found: ' + name);
    if (f.method === 0) return f.data;
    if (f.method === 8) return await inflate(f.data);
    throw new Error('Unsupported compression: ' + f.method);
  }

  // ── Public API ────────────────────────────────────────────────────────────
  return { readZip, getFileText, getFileAsBytes, writeZip };
})();
