// Round-trips every committed fixture through the real M3b pipeline -- export -> build a
// valid CriticMarkup response -> validate -> injectEdits -> upsertComments (if any new
// comments) -> writeZip -> disk -- not just writeZip(zip, {}) (M3a's whole scope). An
// independent tool (Python's zipfile, in verify_zip_roundtrip.py) then verifies the
// output structurally, outside of Vitest/Node's own zip reader.
//
// Deterministic by default (AR_FUZZ_SEED, same convention as the Vitest property suites)
// so a CI failure reproduces exactly from its own log line. Reuses buildValidResponse
// (tests/helpers/randomEdits.js) rather than a bespoke response builder.
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

// Fixtures whose PRISTINE export already contains CriticMarkup-shaped synthetic text
// (pre-existing tracked changes and/or comments render as {++...++}/{--...--}/{~~...~~}/
// {>>...<<} in the exported markdown -- see fixtures/generate.py's tracked_changes(),
// comments_threaded(), and stressor()) are structurally incompatible with
// buildValidResponse, for reasons that go deeper than token *placement* (an earlier
// version of this script assumed a placement/adjacency issue and tried a retry-with-
// fallback strategy; investigating that assumption is what turned up the real mechanisms
// below, each confirmed directly against the actual export.js/validate.js behavior):
//
// - tracked-changes.docx / stressor.docx: buildValidResponse constructs a "response" by
//   splicing new tokens into the pristine EXPORTED markdown, which still contains the
//   pre-existing tokens verbatim. Re-tokenizing that text treats those pre-existing
//   tokens as if they were newly-proposed edits too, and strip()'s "accepted"
//   interpretation of them (ins -> "", del -> its old text, sub -> its old text) does NOT
//   match what the pristine markdown itself shows at that position (the raw, un-stripped
//   token syntax) -- so even echoing the pristine markdown completely unedited already
//   fails G2 (confirmed: tokenize(pristineMarkdown) fails immediately for stressor.docx
//   once ANY edit is spliced in elsewhere, because the stripped/original coordinate
//   spaces diverge downstream of every pre-existing token; tracked-changes.docx fails G2
//   outright on every attempt for the same underlying reason). This is the identical,
//   already-understood limitation tests/helpers/randomEdits.js's own CLEAN_FIXTURES list
//   was built to exclude for M2's property-fuzz suite -- not new to M3b.
// - comments-threaded.docx: additionally exposes a genuine, separate export.js defect
//   found while investigating this: its pristine export contains a REPLY comment whose
//   thread gets rendered *inside* its parent's {==highlight==} span (confirmed:
//   tokenize() of the untouched, zero-edit pristine export already fails G1 nesting at a
//   fixed offset). serializeSegsTracked tracks only one open highlight span at a time
//   (`let openId = null`), so a second, overlapping comment range's thread is emitted
//   inline mid-span instead of after it, producing a `{==...{>>...<<}...==}` nesting.
//   That's a pre-existing rendering defect independent of any edit buildValidResponse
//   could construct -- not something a response-construction strategy can route around.
//
// These are excluded from the "must produce a real injected edit" requirement below with
// zero injected edits (still exercising the round-trip machinery, just unmutated) rather
// than silently warned past, since the reason is understood and documented, not a
// transient/occasional draw.
const STRUCTURALLY_INCOMPATIBLE_FIXTURES = new Set(["tracked-changes.docx", "comments-threaded.docx", "stressor.docx"]);

const fixtures = readdirSync(fixturesDir).filter((f) => f.endsWith(".docx"));
const manifest = {};

for (const name of fixtures) {
  const buf = readFileSync(path.join(fixturesDir, name));
  const bytes = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const zip = await unzip(bytes);

  const { markdown, sourceMap } = await exportDocx(bytes, { DOMParserImpl: DOMParser, filename: name });

  let result = { ok: true, edits: [] };
  if (!STRUCTURALLY_INCOMPATIBLE_FIXTURES.has(name)) {
    // A handful of retries with fresh draws is defensive-only: every fixture actually
    // exercised here has passed on the first attempt in practice. If a fixture that's
    // NOT in the exclusion list above ever fails to produce a real edit set, that's a
    // genuine regression (in buildValidResponse, validate, or the fixture itself) and
    // must fail the build loudly, not get silently skipped.
    let found = false;
    for (let attempt = 0; attempt < 5; attempt++) {
      const candidateResponse = buildValidResponse(rng, markdown, sourceMap);
      const candidateResult = validate({ responseMarkdown: candidateResponse, exportedMarkdown: markdown, sourceMap });
      if (candidateResult.ok && candidateResult.edits.length) {
        result = candidateResult;
        found = true;
        break;
      }
    }
    if (!found) {
      throw new Error(
        `roundtrip-fixtures: ${name} produced no valid random response with at least one edit after 5 attempts ` +
          `(seed=${seed}). This fixture is not in STRUCTURALLY_INCOMPATIBLE_FIXTURES, so this is a regression -- ` +
          `investigate rather than silencing it.`
      );
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
