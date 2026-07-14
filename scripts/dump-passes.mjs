// Dumps clustered reviewer passes for a .docx (reviewer-pass-slicer step 2). Usage:
//   node scripts/dump-passes.mjs <path-to.docx> [out.json]
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { DOMParser } from "@xmldom/xmldom";
import { extractObservations } from "../src/ooxml/observations.js";
import { clusterPasses } from "../src/passes.js";

const inputPath = process.argv[2];
if (!inputPath) {
  console.error("Usage: node scripts/dump-passes.mjs <path-to.docx> [out.json]");
  process.exit(1);
}
const outPath = process.argv[3];

const buf = readFileSync(inputPath);
const bytes = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
const filename = path.basename(inputPath).replace(/\.docx$/i, "");

const { observations } = await extractObservations(bytes, { DOMParserImpl: DOMParser, filename });
const { passes, metadataStripped } = clusterPasses(observations);

console.log(`${passes.length} pass(es)${metadataStripped ? " -- METADATA STRIPPED" : ""}:`);
for (const p of passes) {
  console.log(`  ${p.label}`);
}

const payload = JSON.stringify({ metadataStripped, passes }, null, 2);
if (outPath) {
  writeFileSync(outPath, payload);
  console.log(`Wrote ${passes.length} pass(es) to ${outPath}`);
} else {
  console.log(payload);
}
