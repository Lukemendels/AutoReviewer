// Regression test for the diffWords quadratic blowup found in code review: validate()'s
// G2 failure path used to call diffWords() eagerly on every failure. For a real, large
// document where drift is scattered across many separate points (not one localized
// paraphrase) -- e.g. re-tokenizing an export that already contains CriticMarkup-shaped
// text from pre-existing tracked changes/comments -- prefix/suffix trimming can't shrink
// the divergent window, leaving something close to the full O(n*m) LCS computation.
// Measured on fixtures/stressor.docx (a real ~65K-character export) before the fix: ~13s
// and ~2.1GB heap for a SINGLE diffWords() call inside one validate() call -- enough to
// OOM a retry loop (the roundtrip script, the property fuzz suite) or hang a browser tab.
//
// The fix (see validate.js's G2 comments and src/ui/diff.js) replaces the eager diff with
// a cheap O(n) first-divergence finder; the full diff is now computed lazily, on demand,
// only by the ratification UI's failure view, and even then is capped (diffWords returns
// null past a safe cell count rather than attempting the full DP).
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { exportDocx } from "../src/ooxml/export.js";
import { validate } from "../src/validate.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");

function loadDocx(name) {
  const buf = readFileSync(path.join(fixturesDir, `${name}.docx`));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

describe("validate(): G2 failure performance (regression -- must never run an unbounded diff)", () => {
  it("a G2 failure on stressor.docx completes within budget, even with drift scattered throughout the document", async () => {
    const exported = await exportDocx(loadDocx("stressor"), { DOMParserImpl: DOMParser, filename: "stressor" });
    const md = exported.markdown;
    const sourceMap = exported.sourceMap;

    // Corrupt one word roughly in the middle of the document -- mirrors the exact repro
    // that measured ~13s/~2.1GB before the fix (stressor.docx has pre-existing tracked
    // changes, so its pristine export already contains CriticMarkup-shaped synthetic text
    // scattered throughout; re-tokenizing it as a "response" causes drift at many points,
    // not just the one word corrupted here -- that's what defeats prefix/suffix trimming).
    const allRuns = [];
    for (const block of sourceMap.blocks) for (const run of block.runs) allRuns.push(run);
    const run = allRuns[Math.floor(allRuns.length / 2)];
    const runText = md.slice(run.mdStart, run.mdEnd);
    const match = runText.match(/\S+/);
    const wordStart = run.mdStart + match.index;
    const word = match[0];
    const response = md.slice(0, wordStart) + word + "X" + md.slice(wordStart + word.length);

    const memBefore = process.memoryUsage().heapUsed;
    const start = performance.now();
    const result = validate({ responseMarkdown: response, exportedMarkdown: md, sourceMap });
    const elapsedMs = performance.now() - start;
    const heapDeltaMB = (process.memoryUsage().heapUsed - memBefore) / (1024 * 1024);

    expect(result.ok, "expected this corruption to fail validation").toBe(false);
    expect(elapsedMs, `validate() took ${elapsedMs.toFixed(1)}ms`).toBeLessThan(200);
    expect(heapDeltaMB, `validate() allocated ${heapDeltaMB.toFixed(1)}MB`).toBeLessThan(50);

    // The cheap diagnostic must still be present and correct -- the fix isn't allowed to
    // just drop the failure information, only the expensive eager full diff.
    expect(result.firstDivergence).toBeTruthy();
    expect(typeof result.firstDivergence.offset).toBe("number");
    expect(result.diffInputs).toBeTruthy();
  });
});
