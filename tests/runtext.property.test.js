// M4d PR-2, F-2: export.js and inject.js now share ONE run-text accumulator
// (src/ooxml/runtext.js) instead of two that quietly disagreed. This property test guards
// against future drift if someone re-forks the logic: for arbitrary mixes of w:t, w:tab,
// w:br, w:cr, w:delText children, the shared function's output is (trivially, by
// construction) the same value both call sites would see -- the real regression it guards
// is a future edit to export.js's or inject.js's own copy going out of sync again.
import { describe, expect, it } from "vitest";
import { DOMParser } from "@xmldom/xmldom";
import { parseXml, NS } from "../src/ooxml/parse.js";
import { runPlainText } from "../src/ooxml/runtext.js";

const CHILD_KINDS = ["t", "tab", "br", "cr", "delText"];

function buildRunXml(spec) {
  const parts = spec.map(({ kind, text }) => {
    if (kind === "t") return `<w:t xml:space="preserve">${text}</w:t>`;
    if (kind === "delText") return `<w:delText xml:space="preserve">${text}</w:delText>`;
    return `<w:${kind}/>`;
  });
  return `<w:r xmlns:w="${NS.w}">${parts.join("")}</w:r>`;
}

function expectedText(spec) {
  let s = "";
  for (const { kind, text } of spec) {
    if (kind === "t" || kind === "delText") s += text;
    else if (kind === "tab") s += " ";
    else if (kind === "br" || kind === "cr") s += "  \n";
  }
  return s;
}

// Deterministic pseudo-random generator (no external fuzz dependency, matching this repo's
// existing validate.property.test.js style) -- a fixed seed makes failures reproducible.
function mulberry32(seed) {
  let a = seed;
  return function () {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function genSpec(rand, length) {
  const spec = [];
  for (let i = 0; i < length; i++) {
    const kind = CHILD_KINDS[Math.floor(rand() * CHILD_KINDS.length)];
    const text = kind === "t" || kind === "delText" ? `w${Math.floor(rand() * 1000)}` : "";
    spec.push({ kind, text });
  }
  return spec;
}

describe("runtext.js: runPlainText matches a from-scratch reimplementation of the accumulation rule", () => {
  it("agrees on 200 generated mixes of w:t/w:tab/w:br/w:cr/w:delText", () => {
    const rand = mulberry32(20260710);
    for (let i = 0; i < 200; i++) {
      const spec = genSpec(rand, 1 + Math.floor(rand() * 8));
      const xml = buildRunXml(spec);
      const doc = parseXml(xml, DOMParser);
      const run = doc.documentElement;
      expect(runPlainText(run)).toBe(expectedText(spec));
    }
  });

  it("a run with only w:t children behaves exactly like inject.js's old w:t-only reader", () => {
    const spec = [
      { kind: "t", text: "hello " },
      { kind: "t", text: "world" },
    ];
    const doc = parseXml(buildRunXml(spec), DOMParser);
    expect(runPlainText(doc.documentElement)).toBe("hello world");
  });
});
