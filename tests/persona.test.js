import { describe, expect, it } from "vitest";
import { parsePersona, DEFAULT_PERSONA } from "../src/persona.js";

const FULL_PERSONA = `---
okf: persona
name: RIA Economist Reviewer
description: >-
  Reviews regulatory impact analyses for economic soundness, cost-benefit
  completeness, and OMB Circular A-4 compliance.
version: 1.0
updated: 2026-07-01
---
## Role and voice
A rigorous, neutral economist reviewer.

## Review priorities        (ordered)
1. Cost-benefit completeness
2. OMB Circular A-4 compliance

## Style exemplars          (2-5 before/after pairs)
Before: The cost is high.
After: The estimated cost is $4.2M (2026 dollars).

## Do-not-touch rules
- Do not alter statutory citations.
- Do not remove the OMB disclaimer paragraph.

## Comment conventions
Use comments for open questions; edits for clear numeric fixes.
`;

describe("parsePersona: a fully-formed OKF persona file", () => {
  it("reads the frontmatter name over any H1/filename fallback", () => {
    const p = parsePersona(FULL_PERSONA);
    expect(p.name).toBe("RIA Economist Reviewer");
  });

  it("extracts each §10 section body, tolerating a trailing heading parenthetical", () => {
    const p = parsePersona(FULL_PERSONA);
    expect(p.roleAndVoice).toContain("rigorous, neutral economist reviewer");
    expect(p.reviewPriorities).toContain("Cost-benefit completeness");
    expect(p.commentConventions).toContain("Use comments for open questions");
  });

  it("parses do-not-touch rules into a clean array (bullet markers stripped)", () => {
    const p = parsePersona(FULL_PERSONA);
    expect(p.doNotTouch).toEqual(["Do not alter statutory citations.", "Do not remove the OMB disclaimer paragraph."]);
  });

  it("parses a Before/After style exemplar into a structured pair", () => {
    const p = parsePersona(FULL_PERSONA);
    expect(p.styleExemplars).toEqual([{ before: "The cost is high.", after: "The estimated cost is $4.2M (2026 dollars)." }]);
  });

  it("has no warnings when every section is present", () => {
    const p = parsePersona(FULL_PERSONA);
    expect(p.warnings).toEqual([]);
    expect(p.isDefault).toBe(false);
  });

  it("keeps the raw source for audit hashing", () => {
    const p = parsePersona(FULL_PERSONA);
    expect(p.raw).toBe(FULL_PERSONA);
  });
});

describe("parsePersona: missing sections warn, never block", () => {
  const MISSING_COMMENTS = FULL_PERSONA.replace(/## Comment conventions[\s\S]*$/, "");

  it("still returns a usable model with a warning naming the missing section", () => {
    const p = parsePersona(MISSING_COMMENTS);
    expect(p.warnings).toContain("Missing section: Comment conventions");
    expect(p.commentConventions).toBe("");
    // Everything else present is unaffected.
    expect(p.roleAndVoice).toContain("rigorous, neutral economist reviewer");
  });
});

describe("parsePersona: name fallback chain", () => {
  it("falls back to the H1 title when frontmatter has no name", () => {
    const md = "---\nokf: persona\n---\n# Fallback Title\n## Role and voice\nx\n";
    expect(parsePersona(md).name).toBe("Fallback Title");
  });

  it("falls back to the filename when neither frontmatter name nor H1 exist", () => {
    const md = "## Role and voice\nx\n";
    expect(parsePersona(md, { filename: "jim-editorial.md" }).name).toBe("jim-editorial");
  });
});

describe("parsePersona: unstructured style exemplars", () => {
  it("falls back to one raw entry when the Before/After shape isn't followed", () => {
    const md = FULL_PERSONA.replace("Before: The cost is high.\nAfter: The estimated cost is $4.2M (2026 dollars).", "Just some free-form guidance here.");
    const p = parsePersona(md);
    expect(p.styleExemplars).toEqual([{ raw: "Just some free-form guidance here." }]);
  });
});

describe("DEFAULT_PERSONA", () => {
  it("is clearly labeled as the built-in default, with no persona content to leak into do-not-touch", () => {
    expect(DEFAULT_PERSONA.isDefault).toBe(true);
    expect(DEFAULT_PERSONA.doNotTouch).toEqual([]);
    expect(typeof DEFAULT_PERSONA.roleAndVoice).toBe("string");
    expect(DEFAULT_PERSONA.roleAndVoice.length).toBeGreaterThan(0);
  });
});
