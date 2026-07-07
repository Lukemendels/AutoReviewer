import { describe, expect, it } from "vitest";
import { composeRepair } from "../src/repair.js";

describe("composeRepair: G1 (grammar)", () => {
  const responseText = "...end of paragraph one.{--\n\nParagraph two--}...";
  const failure = {
    ok: false,
    gate: "G1",
    message: "a CriticMarkup token crosses a paragraph/block boundary (position 10-40)",
    detail: { rawStart: 25 },
  };

  it("reuses failure.message verbatim and shows the corrected-pattern cue", () => {
    const out = composeRepair(failure, responseText, 0);
    expect(out).toContain(failure.message);
    expect(out).toContain("one token per block");
  });
});

describe("composeRepair: G2 (fidelity)", () => {
  const failure = {
    ok: false,
    gate: "G2",
    message: "the response's underlying text does not byte-match the exported document outside CriticMarkup tokens",
    firstDivergence: {
      offset: 1204,
      before: "...",
      afterA: "the carrier shall provide",
      afterB: "the carrier must provide",
      truncatedBefore: true,
      truncatedAfterA: true,
      truncatedAfterB: true,
    },
  };

  it("reuses failure.message verbatim, shows the corrected-pattern cue, and quotes both afterA and afterB", () => {
    const out = composeRepair(failure, "irrelevant response text", 0);
    expect(out).toContain(failure.message);
    expect(out).toContain(failure.firstDivergence.afterA);
    expect(out).toContain(failure.firstDivergence.afterB);
    expect(out).toContain("{~~old~>new~~}");
  });
});

describe("composeRepair: G3 (anchor resolution)", () => {
  const responseText = "# {~~Old Title~>New Title~~}";
  const failure = {
    ok: false,
    gate: "G3",
    message: "edit [0,2) is not fully covered by document-text runs (resolved 0/2 chars)",
    detail: { rawStart: 0 },
  };

  it("reuses failure.message verbatim and shows the corrected-pattern cue", () => {
    const out = composeRepair(failure, responseText, 0);
    expect(out).toContain(failure.message);
    expect(out).toContain("exporter invented");
  });
});

describe("composeRepair: G4 (locked content)", () => {
  const responseText = "See ⟦{~~x~>y~~}⟧ for details.";
  const failure = {
    ok: false,
    gate: "G4",
    message: "edit overlaps a locked range",
    detail: { rawStart: 4 },
  };

  it("reuses failure.message verbatim and shows the corrected-pattern cue", () => {
    const out = composeRepair(failure, responseText, 0);
    expect(out).toContain(failure.message);
    expect(out).toContain("locked content");
  });
});

describe("composeRepair: restart policy", () => {
  const failure = { ok: false, gate: "G2", message: "drift", firstDivergence: { offset: 0, afterA: "a", afterB: "b" } };

  it("is absent below attemptCount 2", () => {
    expect(composeRepair(failure, "x", 0)).not.toContain("start clean");
    expect(composeRepair(failure, "x", 1)).not.toContain("start clean");
  });

  it("prepends the restart block at attemptCount >= 2", () => {
    const out = composeRepair(failure, "x", 2);
    expect(out.startsWith("You've hit the same issue twice.")).toBe(true);
    expect(out).toContain(failure.message);

    const outHigher = composeRepair(failure, "x", 5);
    expect(outHigher.startsWith("You've hit the same issue twice.")).toBe(true);
  });
});
