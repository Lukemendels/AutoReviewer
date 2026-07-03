// ZIP writer (spec §9.3). For every entry in the original archive, untouched parts are
// copied byte-for-byte compressed as-is (reuse the original compressed bytes + CRC --
// M3a's whole scope, since ooxml/inject.js doesn't exist yet); mutated parts (M3b) are
// deflated fresh via CompressionStream and get freshly-computed CRC32/sizes. Local file
// headers are always rebuilt rather than copied verbatim from the original, so both paths
// go through one code path -- extra-field bytes (Zip64/NTFS/Unix extra data) are not
// round-tripped, so output isn't guaranteed byte-identical to the input container, only
// content-identical (real .docx files from Word/python-docx don't rely on extra fields for
// correctness).

const CRC_TABLE = buildCrcTable();
function buildCrcTable() {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
}
export function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

const MAX_ENTRIES = 65535;
const MAX_TOTAL_SIZE = 4 * 1024 * 1024 * 1024; // 4 GiB, no ZIP64 support at this size

async function deflateRaw(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

// Where an entry's compressed data actually starts: past its own local header + name +
// extra field (which can differ in length from the central directory's copy).
function localDataStart(zip, name) {
  const e = zip.entries[name];
  const nameLen = zip.dv.getUint16(e.localOff + 26, true);
  const extraLen = zip.dv.getUint16(e.localOff + 28, true);
  return e.localOff + 30 + nameLen + extraLen;
}

export async function writeZip(zip, mutatedParts = {}) {
  const order = zip.order || Object.keys(zip.entries);
  if (order.length > MAX_ENTRIES) {
    throw new Error(`writeZip: ${order.length} entries exceeds the ${MAX_ENTRIES}-entry limit (no ZIP64 support)`);
  }

  const encoder = new TextEncoder();
  const prepared = order.map((name) => {
    const meta = zip.entries[name];
    if (Object.prototype.hasOwnProperty.call(mutatedParts, name)) {
      return { name, mutated: true, meta, textBytes: encoder.encode(mutatedParts[name]) };
    }
    return { name, mutated: false, meta };
  });

  // Pre-flight the size guard using each entry's own declared size, before touching any
  // real byte data (mutated parts' pre-deflate size is a conservative upper bound).
  let estimatedTotal = 0;
  for (const item of prepared) estimatedTotal += item.mutated ? item.textBytes.length : item.meta.compSize;
  if (estimatedTotal >= MAX_TOTAL_SIZE) {
    throw new Error(`writeZip: total size >= 4 GiB exceeds the ZIP64-free limit`);
  }

  // Resolve each entry's final compressed bytes + CRC/sizes/method.
  for (const item of prepared) {
    if (item.mutated) {
      item.compBytes = await deflateRaw(item.textBytes);
      item.compSize = item.compBytes.length;
      item.uncompSize = item.textBytes.length;
      item.crc32 = crc32(item.textBytes);
      item.method = 8;
    } else {
      const start = localDataStart(zip, item.name);
      item.compBytes = zip.bytes.subarray(start, start + item.meta.compSize);
      item.compSize = item.meta.compSize;
      item.uncompSize = item.meta.uncompSize;
      item.crc32 = item.meta.crc32;
      item.method = item.meta.method;
    }
  }

  // Local file headers + data, tracking each entry's new offset for the central directory.
  const localChunks = [];
  let offset = 0;
  for (const item of prepared) {
    const nameBytes = encoder.encode(item.name);
    const header = new Uint8Array(30);
    const hv = new DataView(header.buffer);
    hv.setUint32(0, 0x04034b50, true);
    hv.setUint16(4, 20, true); // version needed to extract
    hv.setUint16(6, 0, true); // general purpose flag: no data descriptor, sizes known upfront
    hv.setUint16(8, item.method, true);
    hv.setUint16(10, item.mutated ? 0 : item.meta.dosTime, true);
    hv.setUint16(12, item.mutated ? 0 : item.meta.dosDate, true);
    hv.setUint32(14, item.crc32, true);
    hv.setUint32(18, item.compSize, true);
    hv.setUint32(22, item.uncompSize, true);
    hv.setUint16(26, nameBytes.length, true);
    hv.setUint16(28, 0, true); // extra field length
    item.localOffset = offset;
    localChunks.push(header, nameBytes, item.compBytes);
    offset += header.length + nameBytes.length + item.compSize;
  }

  // Central directory.
  const cdStart = offset;
  const cdChunks = [];
  for (const item of prepared) {
    const nameBytes = encoder.encode(item.name);
    const header = new Uint8Array(46);
    const hv = new DataView(header.buffer);
    hv.setUint32(0, 0x02014b50, true);
    hv.setUint16(4, item.mutated ? 20 : item.meta.versionMadeBy, true);
    hv.setUint16(6, 20, true); // version needed to extract
    hv.setUint16(8, 0, true); // general purpose flag
    hv.setUint16(10, item.method, true);
    hv.setUint16(12, item.mutated ? 0 : item.meta.dosTime, true);
    hv.setUint16(14, item.mutated ? 0 : item.meta.dosDate, true);
    hv.setUint32(16, item.crc32, true);
    hv.setUint32(20, item.compSize, true);
    hv.setUint32(24, item.uncompSize, true);
    hv.setUint16(28, nameBytes.length, true);
    hv.setUint16(30, 0, true); // extra field length
    hv.setUint16(32, 0, true); // comment length
    hv.setUint16(34, 0, true); // disk number start
    hv.setUint16(36, item.mutated ? 0 : item.meta.internalAttrs, true);
    hv.setUint32(38, item.mutated ? 0 : item.meta.externalAttrs, true);
    hv.setUint32(42, item.localOffset, true);
    cdChunks.push(header, nameBytes);
  }
  const cdSize = cdChunks.reduce((sum, c) => sum + c.length, 0);

  // End of central directory record.
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true); // disk number
  ev.setUint16(6, 0, true); // disk with central directory start
  ev.setUint16(8, prepared.length, true);
  ev.setUint16(10, prepared.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, cdStart, true);
  ev.setUint16(20, 0, true); // comment length

  const allChunks = [...localChunks, ...cdChunks, eocd];
  const totalLen = allChunks.reduce((sum, c) => sum + c.length, 0);
  if (totalLen >= MAX_TOTAL_SIZE) {
    throw new Error(`writeZip: output size >= 4 GiB exceeds the ZIP64-free limit`);
  }

  const out = new Uint8Array(totalLen);
  let pos = 0;
  for (const chunk of allChunks) {
    out.set(chunk, pos);
    pos += chunk.length;
  }
  return out.buffer;
}
