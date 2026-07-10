import { describe, expect, it } from "vitest";
import { buildPrompt, PROMPT_TEMPLATE_VERSION, CHUNK_PROMPT_TEMPLATE_VERSION, CHUNK_WORD_THRESHOLD } from "../src/prompt.js";
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

describe("buildPrompt: chunk mode (M4c, spec §6.4 / architecture doc §7)", () => {
  // A later chunk's own exportedMarkdown never starts with the document header --
  // chunk.js's splitIntoChunks folds the header into chunk 0 only.
  const CHUNK_1_MARKDOWN = ["# Middle Section", "", "This is the middle chunk's own paragraph."].join("\n");

  it("single-doc calls (no chunk arg) are completely unaffected -- same text, same version", () => {
    const withoutChunk = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    expect(withoutChunk.promptVersion).toBe(PROMPT_TEMPLATE_VERSION);
    expect(withoutChunk.text).not.toContain("part 1 of");
  });

  it("prepends a part-N-of-M preamble to [TASK] when chunk is given", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x", chunk: { index: 0, total: 3 } });
    const task = text.slice(text.indexOf("[TASK]"), text.indexOf("[CRITICMARKUP RULES]"));
    expect(task).toContain("This is part 1 of 3 of a larger document.");
    expect(task.toLowerCase()).toContain("return only this part");
  });

  it("indexes the preamble 1-based from a 0-based chunk.index", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: CHUNK_1_MARKDOWN, filename: "x", chunk: { index: 1, total: 3 } });
    expect(text).toContain("This is part 2 of 3 of a larger document.");
  });

  it("chunk-mode prompts get their own distinct promptVersion, never the single-doc one (D1)", () => {
    const { promptVersion } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x", chunk: { index: 0, total: 2 } });
    expect(promptVersion).toBe(CHUNK_PROMPT_TEMPLATE_VERSION);
    expect(promptVersion).not.toBe(PROMPT_TEMPLATE_VERSION);
  });

  it("chunk 0 still carries the header-echo [HARD CONSTRAINTS] rule (it carries the real header)", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x", chunk: { index: 0, total: 2 } });
    const hardConstraints = text.slice(text.indexOf("[HARD CONSTRAINTS]"), text.indexOf("[DOCUMENT]"));
    expect(hardConstraints).toContain("<!-- Redline export from: policy-draft.docx -->");
    expect(hardConstraints.toLowerCase()).toContain("your response must begin with");
  });

  it("a later chunk has no header to echo, so the header-echo bullet is omitted entirely", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: CHUNK_1_MARKDOWN, filename: "x", chunk: { index: 1, total: 3 } });
    const hardConstraints = text.slice(text.indexOf("[HARD CONSTRAINTS]"), text.indexOf("[DOCUMENT]"));
    expect(hardConstraints.toLowerCase()).not.toContain("your response must begin with");
    expect(hardConstraints).not.toContain("<!--");
    // The rest of [HARD CONSTRAINTS] (locked content, no nesting, single fenced block) is
    // unaffected -- only the header-echo bullet is conditional.
    expect(hardConstraints).toContain("Text inside ⟦…⟧ and [image: …] is locked");
  });
});

describe("buildPrompt: M4d PR-3 (prompt fold from the hand-tuned DHSChat test prompt)", () => {
  it("records the m4d template version", () => {
    expect(PROMPT_TEMPLATE_VERSION).toBe("m4d-2026.07-1");
  });

  it("carries the structure-sensitive D1 worked example AND its WRONG counterpart", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    const rules = text.slice(text.indexOf("[CRITICMARKUP RULES]"), text.indexOf("[HARD CONSTRAINTS]"));
    expect(rules).toContain("Worked example (structure-sensitive):");
    expect(rules).toContain("Existing paragraph one.");
    expect(rules).toContain("Existing paragraph three.");
    expect(rules.toLowerCase()).toContain("wrong");
    // The anti-example must actually differ from the correct one -- wrapped in its own
    // blank lines -- not just repeat the same worked example under a WRONG label.
    const wrongIdx = rules.toLowerCase().indexOf("wrong");
    const anti = rules.slice(wrongIdx);
    expect(anti).toContain("{++This entire line is a new inserted paragraph.++}\n\n  Existing paragraph three.");
  });

  it("carries the [WHITESPACE AND LINE BREAKS] section, including the closing-fence rule", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    const rules = text.slice(text.indexOf("[CRITICMARKUP RULES]"), text.indexOf("[HARD CONSTRAINTS]"));
    expect(rules).toContain("[WHITESPACE AND LINE BREAKS]");
    expect(rules.toLowerCase()).toContain("last character before");
    expect(rules.toLowerCase()).toContain("closing ``` fence");
  });

  it("contains no smart quotes (U+2019) anywhere in the assembled prompt", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    expect(text).not.toContain("’");
  });
});

describe("buildPrompt: M4d PR-3 (F-6, header derivation hardened against drift)", () => {
  it("quotes exactly the sourceMap-derived header, not a hardcoded 3-line guess", () => {
    // A synthetic header shaped differently from the real 3-line export header -- proves
    // the derivation is content-anchored (blocks[0].mdStart), not a `split("\n").slice(0,3)`
    // guess that would silently misquote a header of a different shape.
    const doc = "<!-- one header line -->\n\nFirst real paragraph.\n";
    const sourceMap = { blocks: [{ mdStart: doc.indexOf("First real paragraph.") }] };
    const { text } = buildPrompt({ persona: null, exportedMarkdown: doc, filename: "x", sourceMap });
    const hardConstraints = text.slice(text.indexOf("[HARD CONSTRAINTS]"), text.indexOf("[DOCUMENT]"));
    expect(hardConstraints).toContain("<!-- one header line -->");
    expect(hardConstraints).not.toContain("First real paragraph.");
  });

  it("falls back to the old first-3-lines guess when no sourceMap is given (back-compat)", () => {
    const { text } = buildPrompt({ persona: null, exportedMarkdown: EXPORTED, filename: "x" });
    const hardConstraints = text.slice(text.indexOf("[HARD CONSTRAINTS]"), text.indexOf("[DOCUMENT]"));
    expect(hardConstraints).toContain("<!-- Redline export from: policy-draft.docx -->");
    expect(hardConstraints).toContain("<!-- CriticMarkup legend:");
  });
});
