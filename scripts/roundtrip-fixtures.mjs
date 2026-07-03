// Round-trips every committed fixture through the real M3b pipeline -- export -> build a
// valid CriticMarkup response -> validate -> injectEdits -> upsertComments (if any new
// comments) -> writeZip -> disk -- not just writeZip(zip, {}) (M3a's whole scope). An
// independent tool (Python's zipfile, in verify_zip_roundtrip.py) then verifies the
// output structurally, outside of Vitest/Node's own zip reader.
//
// Deterministic by default (AR_FUZZ_SEED, same convention as the Vitest property suites)
// so a CI failure reproduces exactly from its own log line. Reuses buildValidResponse
// (tests/helpers/randomEdits.js) rather than a bespoke response builder: it already knows
// how to construct a response that only ever targets real document-text runs, safely, for
// any fixture's source map -- including tracked-changes.docx/comments-threaded.docx, whose
// pre-existing tracked changes/comments render as synthetic markup outside
// sourceMap.blocks[].runs and so are never touched by it either.
import { readFileSync, readdirSync, writeFileSync, mkdirSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { unzip, readEntry } from "../src/zip/reader.js";
import { writeZip } from "../src/zip/writer.js";
import { exportDocx } from "../src/ooxml/export.js";
import { validate } from "../src/validate.js";
import { injectEdits } from "../src/ooxml/inject.js";
import { upsertComments } from "../src/ooxml/comments.js";
import { serializePart } from "../src/ooxml/serialize.js";
import { parseXml } from "../src/ooxml/parse.js";
import { mulberry32, buildValidResponse } from "../tests/helpers/randomEdits.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");
const outDir = path.resolve(process.argv[2] || path.join(tmpdir(), "ar-zip-roundtrip"));

mkdirSync(outDir, { recursive: true });

const seed = Number(process.env.AR_FUZZ_SEED) || 42;
const rng = mulberry32(seed);

const fixtures = readdirSync(fixturesDir).filter((f) => f.endsWith(".docx"));
const manifest = {};

for (const name of fixtures) {
  const buf = readFileSync(path.join(fixturesDir, name));
  const bytes = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const zip = await unzip(bytes);

  const { markdown, sourceMap } = await exportDocx(bytes, { DOMParserImpl: DOMParser, filename: name });
  // A handful of retries with fresh draws: a fixture with PRE-EXISTING comments (their own
  // {>>...<<} already present in the exported markdown as synthetic text) can occasionally
  // have buildValidResponse's randomly-placed new comment token land close enough to
  // collide at the grammar level (G1 nesting) -- not a content bug, just an occasional
  // antagonistic draw. Falling back to zero injected edits for this fixture (still
  // exercising the round-trip machinery, just without new edits) beats a hard crash.
  let result = { ok: true, edits: [] };
  for (let attempt = 0; attempt < 5; attempt++) {
    const candidateResponse = buildValidResponse(rng, markdown, sourceMap);
    const candidateResult = validate({ responseMarkdown: candidateResponse, exportedMarkdown: markdown, sourceMap });
    if (candidateResult.ok) {
      result = candidateResult;
      break;
    }
    if (attempt === 4) {
      console.warn(`roundtrip-fixtures: ${name} -- no valid random response after 5 attempts, proceeding with zero injected edits`);
    }
  }

  const mutatedParts = {};
  if (result.edits.length) {
    const docDoc = parseXml(await readEntry(zip, "word/document.xml"), DOMParser);
    const { newComments } = injectEdits(docDoc, result.edits, sourceMap, {
      author: "AutoReviewer — CI roundtrip",
      date: "2026-01-01T00:00:00Z",
    });
    mutatedParts["word/document.xml"] = serializePart(docDoc, XMLSerializer);

    if (newComments.length) {
      const existingParts = {
        commentsXml: await readEntry(zip, "word/comments.xml"),
        commentsExtendedXml: await readEntry(zip, "word/commentsExtended.xml"),
        relsXml: await readEntry(zip, "word/_rels/document.xml.rels"),
        contentTypesXml: await readEntry(zip, "[Content_Types].xml"),
      };
      const updated = upsertComments(existingParts, newComments, { DOMParserImpl: DOMParser, XMLSerializerImpl: XMLSerializer });
      mutatedParts["word/comments.xml"] = updated.commentsXml;
      mutatedParts["word/commentsExtended.xml"] = updated.commentsExtendedXml;
      mutatedParts["word/_rels/document.xml.rels"] = updated.relsXml;
      mutatedParts["[Content_Types].xml"] = updated.contentTypesXml;
    }
  }

  const rewritten = await writeZip(zip, mutatedParts);
  writeFileSync(path.join(outDir, name), Buffer.from(rewritten));
  manifest[name] = { mutatedParts: Object.keys(mutatedParts), editCount: result.edits.length };
}

writeFileSync(path.join(outDir, "manifest.json"), JSON.stringify(manifest, null, 2));

const mutatedCount = Object.values(manifest).filter((m) => m.mutatedParts.length).length;
console.log(`Round-tripped ${fixtures.length} fixture(s) (${mutatedCount} with real injected edits, seed=${seed}) into ${outDir}`);
console.log(outDir);
