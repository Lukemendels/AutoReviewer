// Ported verbatim from ref/redline-to-markdown.html (lines 123-164). Pure functions,
// no DOM -- runs in Node (native DecompressionStream/Blob/Response, Node >= 18) and
// in the browser identically.

export async function unzip(buf) {
  const dv = new DataView(buf), bytes = new Uint8Array(buf), len = bytes.length;
  let eocd = -1;
  for (let i = len - 22; i >= Math.max(0, len - 22 - 65535); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("not a valid .docx (no ZIP end record)");
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const entries = {};
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true);
    const compSize = dv.getUint32(p + 20, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commLen = dv.getUint16(p + 32, true);
    const localOff = dv.getUint32(p + 42, true);
    const name = new TextDecoder().decode(bytes.subarray(p + 46, p + 46 + nameLen));
    entries[name] = { method, compSize, localOff };
    p += 46 + nameLen + extraLen + commLen;
  }
  return { dv, bytes, entries };
}

async function inflateRaw(u8) {
  const stream = new Blob([u8]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

export async function readEntry(zip, name) {
  const e = zip.entries[name];
  if (!e) return null;
  if (zip.dv.getUint32(e.localOff, true) !== 0x04034b50) throw new Error("bad local header: " + name);
  const nameLen = zip.dv.getUint16(e.localOff + 26, true);
  const extraLen = zip.dv.getUint16(e.localOff + 28, true);
  const start = e.localOff + 30 + nameLen + extraLen;
  const comp = zip.bytes.subarray(start, start + e.compSize);
  let out;
  if (e.method === 0) out = comp;
  else if (e.method === 8) out = await inflateRaw(comp);
  else throw new Error("unsupported compression in " + name);
  return new TextDecoder("utf-8").decode(out);
}
