import { describe, expect, it } from "vitest";
import {
  buildDecisionAnalysisPrompt,
  buildIncorporationPrompt,
  buildPersonaPrompt,
  transformReviewPrompt,
} from "../src/ui/integrated-workbench.js";

describe("integrated workbench pure contracts", () => {
  it("replaces the model-authored CriticMarkup contract with clean revised Markdown", () => {
    const prompt = [
      "[PERSONA]",
      "Default",
      "",
      "[TASK]",
      "Return CriticMarkup.",
      "",
      "[CRITICMARKUP RULES]",
      "{++text++}",
      "",
      "[HARD CONSTRAINTS]",
      "Old constraints.",
      "",
      "[DOCUMENT]",
      "```markdown",
      "<!-- AutoReviewer -->",
      "Text.",
      "```",
    ].join("\n");
    const transformed = transformReviewPrompt(prompt);
    expect(transformed).toContain("clean revised Markdown");
    expect(transformed).not.toContain("[CRITICMARKUP RULES]");
    expect(transformed).toContain("[DOCUMENT]");
  });

  it("builds a decision packet bound to the active document hash", () => {
    const prompt = buildDecisionAnalysisPrompt({
      documentHash: "sha256-test",
      feedback: [{ feedbackId: "C1", kind: "comment", author: "Reviewer", text: "Clarify scope." }],
      exported: { markdown: "Annotated text" },
    });
    expect(prompt).toContain('"documentHash": "sha256-test"');
    expect(prompt).toContain("Every feedback ID must appear exactly once");
  });

  it("builds a self-contained incorporation packet from confirmed decisions", () => {
    const prompt = buildIncorporationPrompt(
      {
        baselineMarkdown: "Baseline",
        exported: { markdown: "Annotated" },
        feedback: [{ feedbackId: "C1", text: "Clarify." }],
      },
      { schema: "autoreviewer.a2ui.selection/v1", decisions: [{ themeId: "scope", selectedOptionId: "narrow" }] }
    );
    expect(prompt).toContain("[HUMAN DECISIONS]");
    expect(prompt).toContain("[PRE-REVIEW BASELINE]");
    expect(prompt).toContain("narrow");
  });

  it("builds a grounded persona-synthesis packet", () => {
    const prompt = buildPersonaPrompt("reviewer pass corpus");
    expect(prompt).toContain("portable AutoReviewer persona");
    expect(prompt).toContain("[TRAINING CORPUS]");
  });
});
