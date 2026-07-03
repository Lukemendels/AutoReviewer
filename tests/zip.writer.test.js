// M3a scope: writeZip mutates nothing -- every entry takes the "reuse original compressed
// bytes + CRC as-is" path (spec §9.3). These tests are written before the implementation
// (writer.js is still a stub at this point).
import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { unzip, readEntryBytes } from "../src/zip/reader.js";
import { writeZip, crc32 } from "../src/zip/writer.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");
const FIXTURES = readdirSync(fixturesDir).filter((f) => f.endsWith(".docx"));

function loadFixture(name) {
  const buf = readFileSync(path.join(fixturesDir, name));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

describe("crc32", () => {
  it("matches the standard CRC-32/ISO-HDLC check value for the ASCII test vector \"123456789\"", () => {
    const bytes = new TextEncoder().encode("123456789");
    expect(crc32(bytes)).toBe(0xcbf43926);
  });
  it("crc32 of an empty input is 0", () => {
    expect(crc32(new Uint8Array(0))).toBe(0);
  });
  it("is sensitive to every byte (single-bit change produces a different CRC)", () => {
    const a = new TextEncoder().encode("The rule shall apply.");
    const b = new TextEncoder().encode("The rule shall apply!");
    expect(crc32(a)).not.toBe(crc32(b));
  });
});

describe("writeZip: no-op round trip (mutate nothing) across the fixture corpus", () => {
  for (const fixtureName of FIXTURES) {
    it(`${fixtureName}: entry order/names are preserved and every entry's decompressed content is byte-identical`, async () => {
      const original = await unzip(loadFixture(fixtureName));
      const rewritten = await writeZip(original, {});
      const roundTripped = await unzip(rewritten);

      expect(roundTripped.order).toEqual(original.order);
      expect(roundTripped.order.length).toBeGreaterThan(0);

      for (const name of original.order) {
        const a = await readEntryBytes(original, name);
        const b = await readEntryBytes(roundTripped, name);
        expect(b, `entry "${name}" content changed after round-trip`).toEqual(a);
      }
    });

    it(`${fixtureName}: recorded CRCs match the original for every reused entry`, async () => {
      const original = await unzip(loadFixture(fixtureName));
      const rewritten = await writeZip(original, {});
      const roundTripped = await unzip(rewritten);

      for (const name of original.order) {
        expect(roundTripped.entries[name].crc32, `CRC mismatch for "${name}"`).toBe(original.entries[name].crc32);
      }
    });
  }
});

describe("writeZip: guard limits (fake small inputs, not a real 65,536-entry archive)", () => {
  it("rejects more than 65535 entries", async () => {
    const fakeZip = {
      entries: {},
      order: Array.from({ length: 65536 }, (_, i) => `f${i}.xml`),
      bytes: new Uint8Array(0),
      dv: new DataView(new ArrayBuffer(0)),
    };
    await expect(writeZip(fakeZip, {})).rejects.toThrow(/65535|entries/i);
  });

  it("rejects when total declared size is at or beyond the 4 GiB ZIP64-free limit", async () => {
    const hugeSize = 4 * 1024 * 1024 * 1024 + 1;
    const fakeZip = {
      entries: {
        "big.bin": {
          compSize: hugeSize,
          uncompSize: hugeSize,
          crc32: 0,
          method: 0,
          dosTime: 0,
          dosDate: 0,
          versionMadeBy: 20,
          internalAttrs: 0,
          externalAttrs: 0,
        },
      },
      order: ["big.bin"],
      bytes: new Uint8Array(0),
      dv: new DataView(new ArrayBuffer(0)),
    };
    await expect(writeZip(fakeZip, {})).rejects.toThrow(/4 ?gi?b|size/i);
  });

  it("does not reject a small, well within-limits input", async () => {
    const original = await unzip(loadFixture("plain-paragraphs.docx"));
    await expect(writeZip(original, {})).resolves.toBeInstanceOf(ArrayBuffer);
  });
});
