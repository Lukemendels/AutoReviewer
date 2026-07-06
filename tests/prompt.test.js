import { describe, expect, it } from "vitest";
import { buildPrompt, PROMPT_TEMPLATE_VERSION, CHUNK_WORD_THRESHOLD } from "../src/prompt.js";
import { parsePersona, DEFAULT_PERSONA } from "../src/persona.js";

const EXPORTED = [
  "<!-- Redline export from: policy-draft.docx -->",
  "<!-- no tracked changes or comments detected -->",
  "<!-- CriticMarkup legend: {++ins++} {--del--} {~~old~>new~~} {==highlighted==} {>>comment<<} -->",
  "",
  "This is the first paragraph.",
  "",
  "This is the second paragraph.",
].join("\n");

describe("buildPrompt: section assembly (spec §6.2)", () => {
  it("assembles sections in the spec-mandated order", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "policy-draft" });
    const order = ["[PERSONA]", "[TASK]", "[CRITICMARKUP RULES]", "[HARD CONSTRAINTS]", "[DOCUMENT]"];
    const positions = order.map((tag) => text.indexOf(tag));
    for (const pos of positions) expect(pos).toBeGreaterThan(-1);
    for (let i = 1; i < positions.length; i++) expect(positions[i]).toBeGreaterThan(positions[i - 1]);
  });

  it("embeds the exported markdown exactly once, verbatim (single source of truth)", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "policy-draft" });
    expect(text.split(EXPORTED).length - 1).toBe(1);
  });

  it("records the current template version", () => {
    const { promptVersion } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "policy-draft" });
    expect(promptVersion).toBe(PROMPT_TEMPLATE_VERSION);
  });

  it("defaults to DEFAULT_PERSONA when no persona is given", () => {
    const withNull = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    const withDefault = buildPrompt({ persona: DEFAULT_PERSONA, exportedMarkdown: EXPORTED, filename: "x" });
    expect(withNull.text).toBe(withDefault.text);
  });
});

describe("buildPrompt: issue #10 (header echo) fix", () => {
  it("[HARD CONSTRAINTS] explicitly quotes the document's own 3-line header and requires verbatim echo", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "policy-draft" });
    const hardConstraints = text.slice(text.indexOf("[HARD CONSTRAINTS]"), text.indexOf("[DOCUMENT]"));
    expect(hardConstraints).toContain("<!-- Redline export from: policy-draft.docx -->");
    expect(hardConstraints).toContain("<!-- CriticMarkup legend:");
    expect(hardConstraints.toLowerCase()).toContain("your response must begin with");
  });
});

describe("buildPrompt: m4-scope-notes rulings (D1 + trailing newline)", () => {
  it("teaches the whole-paragraph insert shape by worked example", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    const rules = text.slice(text.indexOf("[CRITICMARKUP RULES]"), text.indexOf("[HARD CONSTRAINTS]"));
    expect(rules).toContain("Existing paragraph one.");
    expect(rules).toContain("{++This entire line is a new inserted paragraph.++}");
    expect(rules).toContain("Existing paragraph two.");
    expect(rules.toLowerCase()).toContain("alone on its own line");
  });

  it("states the final trailing newline is part of the byte-preservation rule", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    const rules = text.slice(text.indexOf("[CRITICMARKUP RULES]"), text.indexOf("[HARD CONSTRAINTS]"));
    expect(rules.toLowerCase()).toContain("final trailing newline");
  });
});

describe("buildPrompt: do-not-touch compilation", () => {
  const PERSONA_MD = `---
name: Strict Reviewer
---
## Role and voice
Neutral.
## Review priorities
Clarity.
## Style exemplars
Some guidance.
## Do-not-touch rules
- Do not alter Section 4.2.
- Do not remove the disclaimer.
## Comment conventions
Be terse.
`;

  it("compiles persona do-not-touch rules into [HARD CONSTRAINTS]", () => {
    const persona = parsePersona(PERSONA_MD);
    const { text } = buildPrompt({ persona, exportedMarkdown: EXPORTED, filename: "x" });
    const hardConstraints = text.slice(text.indexOf("[HARD CONSTRAINTS]"), text.indexOf("[DOCUMENT]"));
    expect(hardConstraints).toContain("Do not alter Section 4.2.");
    expect(hardConstraints).toContain("Do not remove the disclaimer.");
  });

  it("omits a do-not-touch block for the default persona (it has none)", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    const hardConstraints = text.slice(text.indexOf("[HARD CONSTRAINTS]"), text.indexOf("[DOCUMENT]"));
    expect(hardConstraints).not.toContain("Do-not-touch rules from the persona");
  });
});

describe("buildPrompt: token estimate + chunk threshold", () => {
  it("estimates tokens as chars/4", () => {
    const { text, tokenEstimate } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    expect(tokenEstimate).toBe(Math.ceil(text.length / 4));
  });

  it("is not overThreshold for a short document", () => {
    const { overThreshold, documentWordCount } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    expect(overThreshold).toBe(false);
    expect(documentWordCount).toBeLessThan(CHUNK_WORD_THRESHOLD);
  });

  it("flags overThreshold honestly for a document past the word count threshold", () => {
    const bigDoc = EXPORTED + "\n\n" + "word ".repeat(CHUNK_WORD_THRESHOLD + 1);
    const { overThreshold, documentWordCount } = buildPrompt({ persona: null, exportedMarkdown: bigDoc, filename: "x" });
    expect(documentWordCount).toBeGreaterThan(CHUNK_WORD_THRESHOLD);
    expect(overThreshold).toBe(true);
  });
});
