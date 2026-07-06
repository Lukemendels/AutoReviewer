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

const AUTHOR = "AutoReviewer — CI roundtrip";

/* ------------------------------------------------------------------ *
 * EXPECTED.md -- the human-verification oracle (pre-Phase-2 hardening).
 *
 * A human checking a round-tripped .docx in real Word was given only "check that
 * Accept All reads cleanly," with no record of what was actually proposed --
 * and correctly flagged healthy, character-level fuzz edits as corruption (see
 * the PR this shipped in). This computes, from the SAME data injectEdits itself
 * consumes (validate()'s resolved edit list + the winning response markdown) --
 * never by re-extracting from the written-out .docx -- what a human should see:
 * the raw tokens in context, and the accepted-view text they produce. Computing
 * it from the injected document instead would let the oracle and the system
 * under test share a failure mode (the golden-dataset principle).
 * ------------------------------------------------------------------ */

// Which block (paragraph) an edit anchors to, by md offset -- ins/bare-comment use a
// zero-width point (mdPos); del/sub/anchored-comment use their span's start (mdStart).
// randomEditForRun anchors every edit entirely inside one run, so this never straddles a
// block boundary.
function editAnchorPos(edit) {
  if (edit.type === "ins" || (edit.type === "comment" && !edit.anchored)) return edit.mdPos;
  return edit.mdStart;
}

function blockIndexForEdit(blocks, edit) {
  const pos = editAnchorPos(edit);
  for (let i = 0; i < blocks.length; i++) {
    if (pos >= blocks[i].mdStart && pos <= blocks[i].mdEnd) return i;
  }
  return -1;
}

// Applies the accepted-view of a paragraph's text-changing edits (ins/del/sub) to that
// paragraph's own original markdown slice -- descending by position so each splice only
// affects text to its own left already spliced, exactly buildValidResponse's own ordering
// discipline, applied to acceptance instead of raw-token insertion.
function acceptedParagraphText(markdown, block, edits) {
  const textEdits = edits.filter((e) => e.type === "ins" || e.type === "del" || e.type === "sub");
  const sorted = [...textEdits].sort((a, b) => editAnchorPos(b) - editAnchorPos(a));
  let text = markdown.slice(block.mdStart, block.mdEnd);
  for (const edit of sorted) {
    if (edit.type === "ins") {
      const local = edit.mdPos - block.mdStart;
      text = text.slice(0, local) + edit.newText + text.slice(local);
    } else {
      const s = edit.mdStart - block.mdStart;
      const e = edit.mdEnd - block.mdStart;
      const replacement = edit.type === "sub" ? edit.newText : "";
      text = text.slice(0, s) + replacement + text.slice(e);
    }
  }
  return text;
}

// The raw CriticMarkup token plus ~15 chars of surrounding text on each side, taken
// directly from the winning response (outside the token, response == exportedMarkdown
// byte-for-byte per G2, so this doubles as "surrounding exported markdown").
const CONTEXT_CHARS = 15;
function tokenInContext(response, edit) {
  const before = response.slice(Math.max(0, edit.rawStart - CONTEXT_CHARS), edit.rawStart);
  const raw = response.slice(edit.rawStart, edit.rawEnd);
  const after = response.slice(edit.rawEnd, edit.rawEnd + CONTEXT_CHARS);
  return `${before}${raw}${after}`;
}

// Word's Review pane has no "substitution" revision type -- inject.js emits a sub as one
// w:del (old text) + one w:ins (new text), see ooxml/inject.js's applyOrdinaryEdits. The
// counts a human checks on-screen must reflect that, not validate()'s own gate-level
// per-token counts.
function reviewPaneCounts(counts) {
  return {
    insertions: counts.ins + counts.sub,
    deletions: counts.del + counts.sub,
    comments: counts.comment,
  };
}

function renderFixtureSection(name, markdown, sourceMap, response, result) {
  const lines = [`## ${name}`, ""];
  if (!result.edits.length) {
    lines.push(
      "No edits were injected into this fixture (either structurally excluded from " +
        "random-edit generation -- see this script's STRUCTURALLY_INCOMPATIBLE_FIXTURES -- " +
        "or it legitimately drew zero valid edits this run).",
      "",
      "**Expected: opens clean in Word -- zero revisions, zero comments present.**",
      ""
    );
    return lines.join("\n");
  }

  const { insertions, deletions, comments } = reviewPaneCounts(result.counts);
  lines.push(
    `**Expected Review pane counts:** ${insertions} insertion(s), ${deletions} deletion(s), ${comments} comment(s).`,
    `**Author:** \`${AUTHOR}\``,
    ""
  );

  // A D1 whole-paragraph INSERT anchors to the gap BETWEEN two blocks (or before the
  // first/after the last), never inside one -- blockIndexForEdit's containment check can
  // never resolve one to a real block index (by design: that's exactly how resolvePoint
  // tells a paragraphBoundary point apart from an ordinary in-run one). A whole-paragraph
  // DELETE is the opposite case: its declared span IS a block's own full [mdStart,mdEnd),
  // so blockIndexForEdit already resolves it correctly and it flows through the normal
  // per-block path below unchanged.
  const blocks = sourceMap.blocks || [];
  const wholeParagraphInserts = result.edits.filter(
    (e) => e.anchor && !Array.isArray(e.anchor) && e.anchor.kind === "paragraphBoundary"
  );
  const byBlock = new Map();
  for (const edit of result.edits) {
    if (wholeParagraphInserts.includes(edit)) continue;
    const idx = blockIndexForEdit(blocks, edit);
    if (!byBlock.has(idx)) byBlock.set(idx, []);
    byBlock.get(idx).push(edit);
  }

  for (const edit of wholeParagraphInserts.sort((a, b) => a.rawStart - b.rawStart)) {
    lines.push(
      `### Whole-paragraph insert (${edit.anchor.edge} paragraph \`${JSON.stringify(edit.anchor.bodyPath)}\`)`,
      "",
      "Proposed edit(s) in context:",
      `- \`${tokenInContext(response, edit)}\``,
      "",
      `Expected after Accept All: a NEW paragraph containing "${edit.newText}"`,
      ""
    );
  }

  for (const idx of [...byBlock.keys()].sort((a, b) => a - b)) {
    const block = blocks[idx];
    const edits = byBlock.get(idx).sort((a, b) => a.rawStart - b.rawStart);
    lines.push(`### Paragraph \`${JSON.stringify(block.bodyPath)}\` (block #${idx})`, "", "Proposed edit(s) in context:");
    for (const edit of edits) {
      lines.push(`- \`${tokenInContext(response, edit)}\``);
      if (edit.type === "comment") lines.push(`  - Comment text: "${edit.commentText}"`);
    }
    const hasTextChange = edits.some((e) => e.type === "ins" || e.type === "del" || e.type === "sub");
    if (hasTextChange) {
      const accepted = acceptedParagraphText(markdown, block, edits);
      lines.push("", `Expected after Accept All: "${accepted}"`);
    }
    lines.push("");
  }
  return lines.join("\n");
}

function renderExpectedMd(seed, sections) {
  const preamble = [
    "# Verification Oracle -- Expected Post-Accept-All State",
    "",
    `Generated by \`scripts/roundtrip-fixtures.mjs\` (seed=${seed}).`,
    "",
    "**This is the sheet, not a suggestion.** The check is \"does the reviewed document " +
      "match the tables below,\" never \"does the accepted text read cleanly.\" Edits are " +
      "generated at the *character* level by a fuzz generator " +
      "(`tests/helpers/randomEdits.js`), so accepted text is frequently not natural " +
      "English -- e.g. a substitution landing mid-word. That is expected and correct; " +
      "compare byte-for-byte against this sheet, not by eye for readability.",
  ].join("\n");
  const chunks = [preamble, ...sections].map((c) => c.replace(/\s+$/, ""));
  return chunks.join("\n\n---\n\n") + "\n";
}

const fixtures = readdirSync(fixturesDir).filter((f) => f.endsWith(".docx"));
const manifest = {};
const expectedSections = [];

for (const name of fixtures) {
  const buf = readFileSync(path.join(fixturesDir, name));
  const bytes = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const zip = await unzip(bytes);

  const { markdown, sourceMap } = await exportDocx(bytes, { DOMParserImpl: DOMParser, filename: name });

  let result = { ok: true, edits: [] };
  let response = markdown;
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
        response = candidateResponse;
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

  expectedSections.push(renderFixtureSection(name, markdown, sourceMap, response, result));

  const mutatedParts = {};
  if (result.edits.length) {
    const docDoc = parseXml(await readEntry(zip, "word/document.xml"), DOMParser);
    const { newComments } = injectEdits(docDoc, result.edits, sourceMap, {
      author: AUTHOR,
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
writeFileSync(path.join(outDir, "EXPECTED.md"), renderExpectedMd(seed, expectedSections));

const mutatedCount = Object.values(manifest).filter((m) => m.mutatedParts.length).length;
console.log(`Round-tripped ${fixtures.length} fixture(s) (${mutatedCount} with real injected edits, seed=${seed}) into ${outDir}`);
console.log(outDir);
