import { describe, expect, it } from "vitest";
import { diffWords, findFirstDivergence, MAX_DIFF_CELLS } from "../src/ui/diff.js";

function reconstruct(segments, side) {
  return segments
    .filter((s) => (side === "a" ? s.type !== "add" : s.type !== "del"))
    .map((s) => s.text)
    .join("");
}

describe("diffWords: correctness", () => {
  it("identical strings produce a single same segment", () => {
    const segs = diffWords("The quick fox jumps.", "The quick fox jumps.");
    expect(segs).toEqual([{ type: "same", text: "The quick fox jumps." }]);
  });

  it("a single-word replacement in the middle", () => {
    const segs = diffWords("The quick fox jumps.", "The slow fox jumps.");
    expect(segs.some((s) => s.type === "del" && s.text.includes("quick"))).toBe(true);
    expect(segs.some((s) => s.type === "add" && s.text.includes("slow"))).toBe(true);
    expect(reconstruct(segs, "a")).toBe("The quick fox jumps.");
    expect(reconstruct(segs, "b")).toBe("The slow fox jumps.");
  });

  it("a pure insertion", () => {
    const segs = diffWords("The fox jumps.", "The quick fox jumps.");
    expect(segs.some((s) => s.type === "add" && s.text.includes("quick"))).toBe(true);
    expect(reconstruct(segs, "a")).toBe("The fox jumps.");
    expect(reconstruct(segs, "b")).toBe("The quick fox jumps.");
  });

  it("a pure deletion", () => {
    const segs = diffWords("The quick fox jumps.", "The fox jumps.");
    expect(segs.some((s) => s.type === "del" && s.text.includes("quick"))).toBe(true);
    expect(reconstruct(segs, "a")).toBe("The quick fox jumps.");
    expect(reconstruct(segs, "b")).toBe("The fox jumps.");
  });

  it("reconstructing both sides from segments always round-trips, for an arbitrary drift case", () => {
    const a = "Regulatory text discussing cost-benefit considerations under the proposed rule.";
    const b = "Regulatory text discussing cost and benefit tradeoffs under this proposed rule.";
    const segs = diffWords(a, b);
    expect(reconstruct(segs, "a")).toBe(a);
    expect(reconstruct(segs, "b")).toBe(b);
    expect(segs.some((s) => s.type !== "same")).toBe(true);
  });
});

describe("diffWords: stays fast on a large document with a small localized drift", () => {
  it("completes quickly on a large shared prefix/suffix with a tiny differing core", () => {
    const bigShared = Array.from({ length: 20000 }, (_, i) => `word${i}`).join(" ");
    const a = `${bigShared} ALPHA ${bigShared}`;
    const b = `${bigShared} BETA ${bigShared}`;
    const start = performance.now();
    const segs = diffWords(a, b);
    const elapsedMs = performance.now() - start;
    expect(elapsedMs).toBeLessThan(2000);
    expect(segs.some((s) => s.type === "del" && s.text === "ALPHA")).toBe(true);
    expect(segs.some((s) => s.type === "add" && s.text === "BETA")).toBe(true);
    expect(reconstruct(segs, "a")).toBe(a);
    expect(reconstruct(segs, "b")).toBe(b);
  });
});

// A localized single drift stays fast regardless of overall document size (prefix/suffix
// trim shrinks the divergent window to almost nothing) -- but real G2 failures can have
// drift scattered across MANY separate points (e.g. re-tokenizing a document that already
// contains CriticMarkup-shaped text), which prefix/suffix trimming alone can't shrink:
// after trimming, the divergent middle can span nearly the whole document. Without a
// cap, that's an O(n*m) computation on the full document -- measured at ~13s/~2.1GB for a
// single diffWords call on a real ~65K-character document with scattered drift (see the
// M3b plan's lessons section). diffWords must refuse to run the DP past a safe cell count.
describe("diffWords: hard cap on the divergent middle window (scattered drift, not one localized point)", () => {
  it("returns null instead of running the full O(n*m) DP when the divergent window exceeds MAX_DIFF_CELLS", () => {
    // No shared prefix/suffix at all -- every token differs, so the whole tokenized
    // length is the "divergent middle." Sized comfortably past MAX_DIFF_CELLS.
    const side = Math.ceil(Math.sqrt(MAX_DIFF_CELLS)) + 200;
    const a = Array.from({ length: side }, (_, i) => `a${i}`).join(" ");
    const b = Array.from({ length: side }, (_, i) => `b${i}`).join(" ");
    const start = performance.now();
    const result = diffWords(a, b);
    const elapsedMs = performance.now() - start;
    expect(result).toBeNull();
    expect(elapsedMs).toBeLessThan(500);
  });

  it("still computes a real diff for a window right at the boundary of the cap", () => {
    // tokenizeWords' whitespace-capturing split produces ~2*side tokens for `side` words
    // joined by single spaces, so halve the target side length to land just under the cap.
    const side = Math.floor(Math.sqrt(MAX_DIFF_CELLS) / 2) - 10;
    const a = Array.from({ length: side }, (_, i) => `a${i}`).join(" ");
    const b = Array.from({ length: side }, (_, i) => `b${i}`).join(" ");
    const result = diffWords(a, b);
    expect(result).not.toBeNull();
    expect(reconstruct(result, "a")).toBe(a);
    expect(reconstruct(result, "b")).toBe(b);
  });
});

describe("findFirstDivergence: cheap, O(min(n,m)), never runs the O(n*m) DP", () => {
  it("finds the offset and context around the first differing character", () => {
    const a = "The quick brown fox jumps over the lazy dog.";
    const b = "The quick brown fox leaps over the lazy dog.";
    const result = findFirstDivergence(a, b, 10);
    expect(result.offset).toBe(a.indexOf("jumps"));
    expect(result.before).toBe(a.slice(result.offset - 10, result.offset));
    expect(a.slice(result.offset, result.offset + 5)).toBe("jumps");
    expect(b.slice(result.offset, result.offset + 5)).toBe("leaps");
  });

  it("reports no divergence (offset at end) for identical strings", () => {
    const result = findFirstDivergence("same text", "same text");
    expect(result.offset).toBe("same text".length);
  });

  it("stays fast on a huge document with drift scattered throughout (the case that breaks diffWords without a cap)", () => {
    const side = Math.ceil(Math.sqrt(MAX_DIFF_CELLS)) + 500;
    const a = Array.from({ length: side }, (_, i) => `word${i}`).join(" ");
    const b = Array.from({ length: side }, (_, i) => `WORD${i}`).join(" ");
    const start = performance.now();
    const result = findFirstDivergence(a, b);
    const elapsedMs = performance.now() - start;
    expect(elapsedMs).toBeLessThan(200);
    expect(result.offset).toBe(0);
  });
});
