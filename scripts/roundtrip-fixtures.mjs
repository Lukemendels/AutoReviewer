// Round-trips every committed fixture through unzip -> writeZip({}) -> disk, so an
// independent tool (Python's zipfile, in verify_zip_roundtrip.py) can verify the output
// byte-for-byte against the source, outside of Vitest/Node's own zip reader.
import { readFileSync, readdirSync, writeFileSync, mkdirSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { unzip } from "../src/zip/reader.js";
import { writeZip } from "../src/zip/writer.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");
const outDir = path.resolve(process.argv[2] || path.join(tmpdir(), "ar-zip-roundtrip"));

mkdirSync(outDir, { recursive: true });

const fixtures = readdirSync(fixturesDir).filter((f) => f.endsWith(".docx"));
for (const name of fixtures) {
  const buf = readFileSync(path.join(fixturesDir, name));
  const bytes = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const zip = await unzip(bytes);
  const rewritten = await writeZip(zip, {});
  writeFileSync(path.join(outDir, name), Buffer.from(rewritten));
}

console.log(`Round-tripped ${fixtures.length} fixture(s) into ${outDir}`);
console.log(outDir);
