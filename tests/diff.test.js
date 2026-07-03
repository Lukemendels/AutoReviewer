import { describe, expect, it } from "vitest";
import { diffWords } from "../src/ui/diff.js";

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
