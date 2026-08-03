import { describe, expect, it } from "vitest";
import { createDecisionState, validateDecisionSpec } from "../src/a2ui/decision.js";

function fixture() {
  return {
    schema: "autoreviewer.a2ui.decision/v1",
    documentHash: "sha256-demo",
    title: "Major feedback choices",
    themes: [
      {
        id: "scope",
        title: "Scope of the estimate",
        summary: "The feedback asks whether indirect costs belong in scope.",
        feedbackIds: ["C1", "R1"],
        stakes: "The choice affects the central estimate and its defensibility.",
        options: [
          { id: "broad", label: "Include indirect costs", description: "Expand the estimate.", pros: ["Comprehensive"], cons: ["More assumptions"] },
          { id: "narrow", label: "Keep direct costs", description: "Clarify the boundary.", pros: ["Stable scope"], cons: ["May not satisfy reviewer"] },
          { id: "sensitivity", label: "Add sensitivity", description: "Retain the core estimate and add a case.", pros: ["Balanced"], cons: ["More work"] },
        ],
      },
    ],
  };
}

describe("A2UI decision specification", () => {
  it("validates exact feedback coverage and bounded options", () => {
    const spec = validateDecisionSpec(fixture(), {
      documentHash: "sha256-demo",
      feedbackIds: ["C1", "R1"],
    });
    expect(spec.themes).toHaveLength(1);
    expect(spec.themes[0].options).toHaveLength(3);
  });

  it("rejects omitted feedback", () => {
    expect(() =>
      validateDecisionSpec(fixture(), {
        documentHash: "sha256-demo",
        feedbackIds: ["C1", "R1", "C2"],
      })
    ).toThrow(/coverage incomplete/i);
  });

  it("requires every theme decision and custom text for Other", () => {
    const spec = validateDecisionSpec(fixture(), {
      documentHash: "sha256-demo",
      feedbackIds: ["C1", "R1"],
    });
    const state = createDecisionState(spec);
    expect(state.isComplete()).toBe(false);
    state.select("scope", "other");
    expect(state.isComplete()).toBe(false);
    state.setOtherText("scope", "Keep the estimate narrow but add a qualitative discussion.");
    expect(state.isComplete()).toBe(true);
    const payload = state.confirm();
    expect(payload.schema).toBe("autoreviewer.a2ui.selection/v1");
    expect(payload.decisions[0].otherText).toMatch(/qualitative discussion/);
  });

  it("locks decisions after confirmation", () => {
    const spec = validateDecisionSpec(fixture(), {
      documentHash: "sha256-demo",
      feedbackIds: ["C1", "R1"],
    });
    const state = createDecisionState(spec);
    state.select("scope", "broad");
    state.confirm();
    expect(() => state.select("scope", "narrow")).toThrow(/immutable/i);
  });
});
