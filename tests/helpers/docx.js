// Shared fixture-loading helpers for the M3b injection test files. Not itself a spec
// deliverable -- just avoids repeating the same unzip/parse boilerplate across
// tests/inject.*.test.js and tests/comments.test.js.
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser } from "@xmldom/xmldom";
import { exportDocx } from "../../src/ooxml/export.js";
import { unzip, readEntry } from "../../src/zip/reader.js";
import { parseXml } from "../../src/ooxml/parse.js";

export const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
export const fixturesDir = path.join(root, "fixtures");

export function loadDocxBytes(name) {
  const buf = readFileSync(path.join(fixturesDir, `${name}.docx`));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

export async function exportFixture(name) {
  return exportDocx(loadDocxBytes(name), { DOMParserImpl: DOMParser, filename: name });
}

// Unzips a fixture and parses word/document.xml into a live DOM, for tests that need to
// hand real paragraph/run elements to inject.js's primitives.
export async function loadDocumentXmlDom(name) {
  const zip = await unzip(loadDocxBytes(name));
  const docXml = await readEntry(zip, "word/document.xml");
  const docDoc = parseXml(docXml, DOMParser);
  const body = [...docDoc.documentElement.children].find((c) => c.localName === "body");
  return { zip, docDoc, body };
}

export { DOMParser };
