// Randomized property suite layered on top of tests/validate.test.js's hand-crafted gate
// cases: instead of a fixed set of corruptions I thought to write by hand, this generates
// many random VALID edit sets (spans chosen only from real document-text runs, per the
// source map) and asserts every gate passes, then constructs targeted random corruptions
// in each of four categories and asserts the correctly-named gate fires.
//
// Reproducibility: the PRNG is seeded from AR_FUZZ_SEED (or the current time if unset,
// so different CI runs explore different cases over time); the seed is embedded in every
// assertion message, so a red run's own output is enough to reproduce it locally:
//   AR_FUZZ_SEED=<seed> npx vitest run tests/validate.property.test.js
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser } from "@xmldom/xmldom";
import { beforeAll, describe, expect, it } from "vitest";
import { exportDocx } from "../src/ooxml/export.js";
import { validate } from "../src/validate.js";
import { CLEAN_FIXTURES, mulberry32, randInt, pick, allRuns, buildValidResponse } from "./helpers/randomEdits.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");

function loadDocx(name) {
  const buf = readFileSync(path.join(fixturesDir, `${name}.docx`));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}
async function exportFixture(name) {
  return exportDocx(loadDocx(name), { DOMParserImpl: DOMParser, filename: name });
}

// Each corruption constructor returns a targeted bad response built directly from the
// pristine export (not a perturbation of an arbitrary valid response, which would need
// its own overlap bookkeeping for little benefit), or null if this fixture doesn't have
// the ingredients the category needs (caller skips that iteration for that category).
function corruptParaphrase(rng, exportedMarkdown, sourceMap) {
  const runs = allRuns(sourceMap);
  if (!runs.length) return null;
  const run = pick(rng, runs);
  const text = exportedMarkdown.slice(run.mdStart, run.mdEnd);
  const match = text.match(/\S+/);
  if (!match) return null;
  const wordStart = run.mdStart + match.index;
  const word = match[0];
  return exportedMarkdown.slice(0, wordStart) + word + "X" + exportedMarkdown.slice(wordStart + word.length);
}
function corruptNesting(rng, exportedMarkdown, sourceMap) {
  const runs = allRuns(sourceMap);
  if (!runs.length) return null;
  const run = pick(rng, runs);
  const text = exportedMarkdown.slice(run.mdStart, run.mdEnd);
  const nested = `{++${text.slice(0, 1)}{--${text.slice(1)}--}++}`;
  return exportedMarkdown.slice(0, run.mdStart) + nested + exportedMarkdown.slice(run.mdEnd);
}
function corruptBlockCross(rng, exportedMarkdown, sourceMap) {
  const blocks = sourceMap.blocks;
  if (blocks.length < 2) return null;
  const i = randInt(rng, blocks.length - 1);
  const a = blocks[i], b = blocks[i + 1];
  const margin = 3;
  const start = a.mdEnd - margin, end = b.mdStart + margin;
  if (start < a.mdStart || end > b.mdEnd) return null;
  const oldText = exportedMarkdown.slice(start, end);
  return exportedMarkdown.slice(0, start) + `{--${oldText}--}` + exportedMarkdown.slice(end);
}
function corruptLockedExtension(rng, exportedMarkdown, sourceMap) {
  const locked = sourceMap.locked;
  if (!locked.length) return null;
  const [ls, le] = pick(rng, locked);
  // Stay within the block that contains this locked range -- a locked range sitting right
  // at the start of its own block (e.g. an image alone in its paragraph) has no room for a
  // "just before it" margin without crossing into the previous block instead, which would
  // hit G1 (block-crossing) rather than exercising the G4 check this category targets.
  const containingBlock = sourceMap.blocks.find((b) => b.mdStart <= ls && le <= b.mdEnd);
  if (!containingBlock) return null;
  const margin = 3;
  const start = ls - margin;
  const end = ls + Math.min(2, le - ls);
  if (start < containingBlock.mdStart || end > containingBlock.mdEnd) return null;
  const oldText = exportedMarkdown.slice(start, end);
  return exportedMarkdown.slice(0, start) + `{--${oldText}--}` + exportedMarkdown.slice(end);
}

describe("randomized property suite", () => {
  const seed = Number(process.env.AR_FUZZ_SEED) || Date.now();
  const rng = mulberry32(seed);
  const exportedByFixture = {};

  beforeAll(async () => {
    for (const name of CLEAN_FIXTURES) exportedByFixture[name] = await exportFixture(name);
  });

  it(`~200 random valid edit sets all pass every gate (seed=${seed})`, () => {
    const ITERATIONS = 200;
    for (let i = 0; i < ITERATIONS; i++) {
      const fixtureName = pick(rng, CLEAN_FIXTURES);
      const { markdown: exported, sourceMap } = exportedByFixture[fixtureName];
      const response = buildValidResponse(rng, exported, sourceMap);
      const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
      expect(
        result.ok,
        `seed=${seed} iter=${i} fixture=${fixtureName}${result.ok ? "" : ` gate=${result.gate} message=${result.message}`}`
      ).toBe(true);
    }
  });

  it(`each corruption category is caught by its named gate (seed=${seed})`, () => {
    const ITERATIONS = 40;
    const ran = { paraphrase: 0, nesting: 0, blockCross: 0, lockedExtension: 0 };

    function assertGate(category, response, expectedGate, iter, fixtureName) {
      if (!response) return;
      ran[category]++;
      const result = validate({ responseMarkdown: response, exportedMarkdown: exportedByFixture[fixtureName].markdown, sourceMap: exportedByFixture[fixtureName].sourceMap });
      const ctx = `${category} seed=${seed} iter=${iter} fixture=${fixtureName}`;
      expect(result.ok, `${ctx} expected a failure but got ok:true`).toBe(false);
      expect(result.gate, `${ctx} message=${result.message}`).toBe(expectedGate);
    }

    for (let i = 0; i < ITERATIONS; i++) {
      const fixtureName = pick(rng, CLEAN_FIXTURES);
      const { markdown: exported, sourceMap } = exportedByFixture[fixtureName];
      assertGate("paraphrase", corruptParaphrase(rng, exported, sourceMap), "G2", i, fixtureName);
      assertGate("nesting", corruptNesting(rng, exported, sourceMap), "G1", i, fixtureName);
      assertGate("blockCross", corruptBlockCross(rng, exported, sourceMap), "G1", i, fixtureName);
      assertGate("lockedExtension", corruptLockedExtension(rng, exported, sourceMap), "G4", i, fixtureName);
    }

    for (const [category, count] of Object.entries(ran)) {
      expect(count, `seed=${seed}: corruption category "${category}" never ran -- coverage gap`).toBeGreaterThan(0);
    }
  });
});
