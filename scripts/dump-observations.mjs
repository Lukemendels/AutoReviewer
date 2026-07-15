// Dumps the raw per-tracked-change/comment observation table for a .docx (reviewer-pass-
// slicer, step 1 "audit commit" -- validates extractObservations against a real file
// before any pass-clustering or UI work). Usage:
//   node scripts/dump-observations.mjs <path-to.docx> [out.json]
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { DOMParser } from "@xmldom/xmldom";
import { extractObservations } from "../src/ooxml/observations.js";

const inputPath = process.argv[2];
if (!inputPath) {
  console.error("Usage: node scripts/dump-observations.mjs <path-to.docx> [out.json]");
  process.exit(1);
}
const outPath = process.argv[3];

const buf = readFileSync(inputPath);
const bytes = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
const filename = path.basename(inputPath).replace(/\.docx$/i, "");

const { observations } = await extractObservations(bytes, { DOMParserImpl: DOMParser, filename });

const byAuthor = {};
for (const o of observations) {
  const key = o.author || "(none)";
  byAuthor[key] = byAuthor[key] || { insertion: 0, deletion: 0, comment: 0, "comment-reply": 0 };
  byAuthor[key][o.kind] = (byAuthor[key][o.kind] || 0) + 1;
}

console.log(`${observations.length} observation(s) across ${Object.keys(byAuthor).length} author(s):`);
for (const [author, counts] of Object.entries(byAuthor)) {
  console.log(
    `  ${author}: ${counts.insertion} insertion(s), ${counts.deletion} deletion(s), ` +
      `${counts.comment} comment(s), ${counts["comment-reply"]} repl(y/ies)`
  );
}

const payload = JSON.stringify(observations, null, 2);
if (outPath) {
  writeFileSync(outPath, payload);
  console.log(`Wrote ${observations.length} observation(s) to ${outPath}`);
} else {
  console.log(payload);
}
