// Renders one reviewer-pass slice for a .docx (reviewer-pass-slicer step 3). Usage:
//   node scripts/dump-slice.mjs <path-to.docx> <author> <passDate|undated> [out.md]
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { DOMParser } from "@xmldom/xmldom";
import { extractObservations } from "../src/ooxml/observations.js";
import { clusterPasses } from "../src/passes.js";
import { renderSlice, sliceFilename } from "../src/ooxml/slice.js";

const [inputPath, author, passDateArg, outPath] = process.argv.slice(2);
if (!inputPath || !author || !passDateArg) {
  console.error("Usage: node scripts/dump-slice.mjs <path-to.docx> <author> <passDate|undated> [out.md]");
  process.exit(1);
}

const buf = readFileSync(inputPath);
const bytes = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
const filename = path.basename(inputPath).replace(/\.docx$/i, "");

const { observations } = await extractObservations(bytes, { DOMParserImpl: DOMParser, filename });
const { passes } = clusterPasses(observations);
const wantUndated = passDateArg === "undated";
const pass = passes.find((p) => p.author === author && (wantUndated ? p.undated : p.passDate === passDateArg));
if (!pass) {
  console.error(`No pass found for author=${author} pass=${passDateArg}. Available: ${passes.map((p) => p.label).join(" | ")}`);
  process.exit(1);
}

const md = await renderSlice(bytes, pass, { DOMParserImpl: DOMParser, filename });
const target = outPath || sliceFilename(filename, pass);
writeFileSync(target, md);
console.log(`Wrote ${target}`);
