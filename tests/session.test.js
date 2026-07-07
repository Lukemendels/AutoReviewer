import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { exportDocx } from "../src/ooxml/export.js";
import { validate } from "../src/validate.js";
import { buildPrompt } from "../src/prompt.js";
import { parsePersona } from "../src/persona.js";
import { createAppState, STATES } from "../src/ui/state.js";
import { saveSession, loadSession, SESSION_SCHEMA_VERSION } from "../src/session.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");

function loadDocx(name) {
  const buf = readFileSync(path.join(fixturesDir, `${name}.docx`));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}
async function exportFixture(name) {
  return exportDocx(loadDocx(name), { DOMParserImpl: DOMParser, filename: name });
}
function withEdit(markdown, target, replacement) {
  const first = markdown.indexOf(target);
  const last = markdown.lastIndexOf(target);
  if (first === -1 || first !== last) throw new Error(`withEdit: target not found exactly once: ${JSON.stringify(target)}`);
  return markdown.slice(0, first) + replacement + markdown.slice(first + target.length);
}

// Drives a real app state machine through to RATIFYING against a real fixture, with a
// mix of edit types (so decisions have something real to mix accept/reject over).
async function ratifyingFixture() {
  const docxBytes = loadDocx("plain-paragraphs");
  const exported = await exportDocx(docxBytes, { DOMParserImpl: DOMParser, filename: "plain-paragraphs" });
  let response = exported.markdown;
  response = withEdit(response, "first", "{~~first~>initial~~}");
  response = withEdit(response, "second", "{--second--}");
  response = withEdit(response, "third paragraph", "{++really ++}third paragraph");
  response = withEdit(response, "compare against.", "compare against.{>>Double-check this framing.<<}");

  const s = createAppState();
  s.loadDocument({ docxBytes, filename: "plain-paragraphs", exported });
  const built = buildPrompt({ persona: null, exportedMarkdown: exported.markdown, filename: "plain-paragraphs" });
  s.setPrompt(built);
  s.copyPrompt();
  s.submitResponse(response);
  const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
  expect(result.ok).toBe(true);
  s.validationPassed(result);
  expect(s.state).toBe(STATES.RATIFYING);
  return { s, docxBytes, exported };
}

describe("session round trip (spec §5.4, M4 doc §6.2 mandated case)", () => {
  it("save at RATIFYING with mixed accept/reject decisions -> load -> decisions intact, state === RATIFYING", async () => {
    const { s } = await ratifyingFixture();
    const decisions = s.context.validation.edits.map((_, id) => ({ id, decision: id % 2 === 0 ? "accept" : "reject", reviewed: true }));

    const saved = saveSession({ state: s.state, context: s.context, decisions });
    expect(saved.schemaVersion).toBe(SESSION_SCHEMA_VERSION);
    expect(saved.state).toBe(STATES.RATIFYING);

    const json = JSON.parse(JSON.stringify(saved)); // simulate the download/re-read round trip
    const { state, context } = loadSession(json);

    expect(state).toBe(STATES.RATIFYING);
    expect(context.pendingDecisions).toEqual(decisions);
    // export + validate are deterministic, so the re-derived edit list matches in both
    // length and type/order -- decisions map onto it by `id` (the validate-order index).
    expect(context.validation.ok).toBe(true);
    expect(context.validation.edits.map((e) => e.type)).toEqual(s.context.validation.edits.map((e) => e.type));
  });
});

describe("session round trip: PROMPT_READY (no response yet)", () => {
  it("round trips filename, exported, and a regenerated (not serialized) prompt", async () => {
    const docxBytes = loadDocx("plain-paragraphs");
    const exported = await exportFixture("plain-paragraphs");
    const s = createAppState();
    s.loadDocument({ docxBytes, filename: "plain-paragraphs", exported });
    const built = buildPrompt({ persona: null, exportedMarkdown: exported.markdown, filename: "plain-paragraphs" });
    s.setPrompt(built);
    expect(s.state).toBe(STATES.PROMPT_READY);

    const saved = saveSession({ state: s.state, context: s.context, decisions: [] });
    expect(saved.response).toBeNull();

    const { state, context } = loadSession(JSON.parse(JSON.stringify(saved)));
    expect(state).toBe(STATES.PROMPT_READY);
    expect(context.promptText).toBe(built.text);
    expect(context.tokenEstimate).toBe(built.tokenEstimate);
    expect(context.validation).toBeNull(); // no response yet -- nothing to re-validate
  });
});

describe("session round trip: docxBase64", () => {
  it("decodes back to the exact original bytes", async () => {
    const docxBytes = loadDocx("plain-paragraphs");
    const exported = await exportFixture("plain-paragraphs");
    const s = createAppState();
    s.loadDocument({ docxBytes, filename: "plain-paragraphs", exported });

    const saved = saveSession({ state: s.state, context: s.context, decisions: [] });
    const { context } = loadSession(JSON.parse(JSON.stringify(saved)));

    expect(context.docxBytes.byteLength).toBe(docxBytes.byteLength);
    expect(new Uint8Array(context.docxBytes)).toEqual(new Uint8Array(docxBytes));
  });
});

describe("session round trip: persona", () => {
  it("DEFAULT persona (context.persona === null) saves as personaRef: null and reloads to null", async () => {
    const docxBytes = loadDocx("plain-paragraphs");
    const exported = await exportFixture("plain-paragraphs");
    const s = createAppState();
    s.loadDocument({ docxBytes, filename: "plain-paragraphs", exported });
    expect(s.context.persona).toBeNull();

    const saved = saveSession({ state: s.state, context: s.context, decisions: [] });
    expect(saved.personaRef).toBeNull();

    const { context } = loadSession(JSON.parse(JSON.stringify(saved)));
    expect(context.persona).toBeNull();
  });

  it("a custom persona round trips via raw re-parse (name + do-not-touch rules intact)", async () => {
    const personaMd = `---\nname: Strict Reviewer\n---\n## Role and voice\nNeutral.\n## Review priorities\nClarity.\n## Do-not-touch rules\n- Do not alter Section 4.2.\n## Comment conventions\nBe terse.\n`;
    const persona = parsePersona(personaMd, { filename: "strict-reviewer.md" });

    const docxBytes = loadDocx("plain-paragraphs");
    const exported = await exportFixture("plain-paragraphs");
    const s = createAppState();
    s.loadDocument({ docxBytes, filename: "plain-paragraphs", exported });
    s.loadPersona(persona);

    const saved = saveSession({ state: s.state, context: s.context, decisions: [] });
    expect(saved.personaRef).toEqual({ name: "Strict Reviewer", raw: personaMd });

    const { context } = loadSession(JSON.parse(JSON.stringify(saved)));
    expect(context.persona.name).toBe("Strict Reviewer");
    expect(context.persona.doNotTouch).toEqual(["Do not alter Section 4.2."]);
  });
});
