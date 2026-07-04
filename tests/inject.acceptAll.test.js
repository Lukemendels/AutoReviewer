// spec §13 invariant #3, the primary correctness net for the whole M3b milestone: for
// generated random valid edit sets -- inject -> programmatically apply accept-all
// (transform w:ins unwrap, w:del remove, delText drop) -> extract text -- must equal an
// applyEdits-style reconstruction of what the model's response actually proposed. Reuses
// the exact same random-valid-edit generator as tests/validate.property.test.js (M3b
// plan's own instruction), so this test is written and red before inject.js has a real
// implementation, going green only once the whole ordinary-edit pipeline (run-splitting ->
// wrapping -> accept-all) is correct end to end.
//
// "Extract text" is a plain-text walk (concatenating every w:t/w:delText in document
// order), not a re-export through exportDocx()'s markdown renderer -- deliberately. An
// earlier version of this test compared re-exported MARKDOWN against a naive patch of the
// original markdown string, and kept finding *legitimate, non-injection* mismatches:
// export.js trims leading/trailing whitespace at each block's own edge (so a deletion that
// exposes new edge whitespace re-exports differently than a literal string splice
// predicts), writes its own single-space placeholder for a table cell that becomes
// genuinely empty, and never re-merges two adjacent same-formatted runs back into one
// emphasis span after a mid-run split. All three are real, correct properties of the
// *markdown renderer*, not of injection -- comparing markdown-to-markdown was testing the
// wrong layer. Plain-text extraction sidesteps all three: both sides are built by
// concatenating raw run text with no rendering step in between, so the comparison is
// exactly "does the accepted document contain the text the edits describe," independent
// computations (real DOM mutation vs. string-splicing on run-local text) than either.
import { DOMParser } from "@xmldom/xmldom";
import { beforeAll, describe, expect, it } from "vitest";
import { loadDocxBytes, loadDocumentXmlDom } from "./helpers/docx.js";
import { CLEAN_FIXTURES, mulberry32, pick, buildValidResponse } from "./helpers/randomEdits.js";
import { exportDocx } from "../src/ooxml/export.js";
import { validate } from "../src/validate.js";
import { injectEdits, locateParagraph } from "../src/ooxml/inject.js";

async function exportFixture(name) {
  return exportDocx(loadDocxBytes(name), { DOMParserImpl: DOMParser, filename: name });
}

// Unwraps every w:ins (splicing its children into its parent's place) and removes every
// w:del (including its content) throughout the whole document tree -- recursing into
// table cells and any nodes newly exposed by unwrapping, so nesting depth never matters.
function acceptAll(node) {
  for (const child of [...node.children]) {
    if (child.localName === "del") {
      node.removeChild(child);
    } else if (child.localName === "ins") {
      const ref = child.nextSibling;
      const innerNodes = [...child.children];
      for (const inner of innerNodes) node.insertBefore(inner, ref);
      node.removeChild(child);
      for (const inner of innerNodes) acceptAll(inner);
    } else {
      acceptAll(child);
    }
  }
}

// Concatenates every w:t/w:delText in document order, recursing through everything
// (paragraphs, tables, any wrapper element) -- the "actual" side of the comparison.
function extractPlainText(el) {
  let s = "";
  for (const c of el.children) {
    if (c.localName === "t" || c.localName === "delText") s += c.textContent;
    else s += extractPlainText(c);
  }
  return s;
}

function runText(el) {
  let s = "";
  for (const c of el.children) if (c.localName === "t") s += c.textContent;
  return s;
}

// Every paragraph's bodyPath in document order, descending into table cells -- mirrors
// locateParagraph's own bodyPath semantics (length 1 for a top-level paragraph, length 4
// for a table-cell paragraph: [bodyIdx, rowIdx, cellIdx, pIdx]).
function listParagraphBodyPaths(body) {
  const paths = [];
  let bodyIdx = 0;
  for (const el of body.children) {
    if (el.localName === "p") {
      paths.push([bodyIdx]);
    } else if (el.localName === "tbl") {
      const rows = [...el.children].filter((c) => c.localName === "tr");
      rows.forEach((tr, rowIdx) => {
        const cells = [...tr.children].filter((c) => c.localName === "tc");
        cells.forEach((tc, cellIdx) => {
          const cellParas = [...tc.children].filter((c) => c.localName === "p");
          cellParas.forEach((_, pIdx) => paths.push([bodyIdx, rowIdx, cellIdx, pIdx]));
        });
      });
    }
    bodyIdx++;
  }
  return paths;
}

// Decomposes a paragraph's edits into per-run text-splice operations, in the SAME shape
// injectEdits' own applyOrdinaryEdits builds internally (del/sub per-triple, sub's newText
// landing once on the last triple as its own zero-width op) -- an independent
// *description* of the same intent, not a call into inject.js's own machinery.
function editsToRunOps(edits) {
  const byRun = new Map();
  function push(runIndex, op) {
    if (!byRun.has(runIndex)) byRun.set(runIndex, []);
    byRun.get(runIndex).push(op);
  }
  for (const edit of edits) {
    if (edit.type === "ins") {
      const a = edit.anchor;
      push(a.runIndex, { charStart: a.charOffset, charEnd: a.charOffset, insertText: edit.newText });
    } else if (edit.type === "del") {
      for (const t of edit.anchor) push(t.runIndex, { charStart: t.charStart, charEnd: t.charEnd, insertText: "" });
    } else if (edit.type === "sub") {
      const triples = edit.anchor;
      for (const t of triples) push(t.runIndex, { charStart: t.charStart, charEnd: t.charEnd, insertText: "" });
      const last = triples[triples.length - 1];
      push(last.runIndex, { charStart: last.charEnd, charEnd: last.charEnd, insertText: edit.newText });
    }
    // comment: no text change.
  }
  return byRun;
}

function applyOpsToText(text, ops) {
  const sorted = [...ops].sort((a, b) => b.charStart - a.charStart || b.charEnd - a.charEnd);
  let result = text;
  for (const op of sorted) {
    result = result.slice(0, op.charStart) + op.insertText + result.slice(op.charEnd);
  }
  return result;
}

// Mirrors inject.js's own walkForRun traversal exactly (same wrapper set, same fldSimple/
// sdt handling, same counter semantics) -- necessary because runIndex isn't densely
// packed among direct-child runs (a run nested inside w:sdt/w:fldSimple/w:hyperlink/etc.
// still consumes a counter slot without ever being a direct-child match), AND because
// locked/nested content (a field's cached text, a content control's inner run) is real
// document text that extractPlainText's real-DOM walk will include -- it's never a valid
// edit target (G4 protects it), but it must still appear, unedited, in the "expected"
// output, or the two sides would be comparing different sets of included text entirely.
const RUN_WRAPPER_NAMES = new Set(["ins", "del", "hyperlink", "smartTag"]);
function walkExpectedText(el, direct, counter, ops) {
  let s = "";
  for (const c of el.children) {
    const ln = c.localName;
    if (ln === "r") {
      const idx = counter.next++;
      const text = runText(c);
      if (direct) {
        const runOps = ops.get(idx);
        s += runOps ? applyOpsToText(text, runOps) : text;
      } else {
        s += text; // locked/nested -- never an edit target, always passed through as-is.
      }
    } else if (RUN_WRAPPER_NAMES.has(ln)) {
      s += walkExpectedText(c, false, counter, ops);
    } else if (ln === "fldSimple") {
      s += walkExpectedText(c, false, counter, ops);
    } else if (ln === "sdt") {
      const content = [...c.children].find((cc) => cc.localName === "sdtContent");
      if (content) s += walkExpectedText(content, false, counter, ops);
    }
  }
  return s;
}

// The "expected" side: for every paragraph in the ORIGINAL (unmutated) document, apply
// each candidate run's own edits directly to that run's own text via string splicing,
// using the edit's run-local charStart/charEnd/charOffset (independent of however
// injectEdits actually performs the DOM surgery), then concatenate everything -- locked
// content included, unedited -- in document order.
function buildExpectedPlainText(originalBody, edits) {
  const byBodyPath = new Map();
  for (const edit of edits) {
    const bodyPath = Array.isArray(edit.anchor) ? edit.anchor[0].bodyPath : edit.anchor.bodyPath;
    const key = JSON.stringify(bodyPath);
    if (!byBodyPath.has(key)) byBodyPath.set(key, []);
    byBodyPath.get(key).push(edit);
  }

  let result = "";
  for (const bodyPath of listParagraphBodyPaths(originalBody)) {
    const p = locateParagraph(originalBody, bodyPath);
    const group = byBodyPath.get(JSON.stringify(bodyPath)) || [];
    const ops = editsToRunOps(group);
    result += walkExpectedText(p, true, { next: 0 }, ops);
  }
  return result;
}

describe("accept-all equivalence", () => {
  const seed = Number(process.env.AR_FUZZ_SEED) || Date.now();
  const rng = mulberry32(seed);
  const exportedByFixture = {};

  beforeAll(async () => {
    for (const name of CLEAN_FIXTURES) exportedByFixture[name] = await exportFixture(name);
  });

  it(`random valid edit sets: inject -> accept-all -> extract text equals the run-local applyEdits reconstruction (seed=${seed})`, async () => {
    const ITERATIONS = 60;
    for (let i = 0; i < ITERATIONS; i++) {
      const fixtureName = pick(rng, CLEAN_FIXTURES);
      const { markdown: exported, sourceMap } = exportedByFixture[fixtureName];
      const response = buildValidResponse(rng, exported, sourceMap);
      const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
      const ctx = `seed=${seed} iter=${i} fixture=${fixtureName}`;
      expect(result.ok, `${ctx}${result.ok ? "" : ` gate=${result.gate} message=${result.message}`}`).toBe(true);

      // Comments are scoped out here: accepting tracked changes doesn't remove or resolve
      // a comment (they're orthogonal mechanisms in OOXML), and a comment never changes
      // document text anyway -- there's nothing for this text-equivalence check to say
      // about it. Round-tripping a comment's own content is tests/comments.test.js's job.
      //
      // Whole-paragraph anchors are also scoped out: a random span from randomEditForRun
      // can coincidentally land exactly on a full block's boundaries (a short list item,
      // say), tripping D3's wholeParagraphDelete resolution or D2's paragraphBoundary
      // kind instead of the ordinary per-run triples this test's run-local comparison is
      // built around. Whole-paragraph insert/delete correctness is already covered
      // thoroughly by tests/inject.wholeParagraph.test.js; this test is scoped to ordinary
      // run-splitting.
      const trackedEdits = result.edits.filter(
        (e) => e.type !== "comment" && (Array.isArray(e.anchor) || e.anchor.kind === "run")
      );
      if (!trackedEdits.length) continue; // buildValidResponse can legitimately produce none

      const { body: originalBody } = await loadDocumentXmlDom(fixtureName);
      const expectedText = buildExpectedPlainText(originalBody, trackedEdits);

      const { body } = await loadDocumentXmlDom(fixtureName);
      injectEdits(body.ownerDocument, trackedEdits, sourceMap, { author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z" });
      acceptAll(body);
      const actualText = extractPlainText(body);

      expect(actualText, `${ctx}\nresponse=${JSON.stringify(response)}`).toBe(expectedText);
    }
  });
});
