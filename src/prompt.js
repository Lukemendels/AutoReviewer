// Versioned prompt assembly (spec §6.2, §6.1 step 2). Pure: takes a persona + the export
// leg's output and returns the exact text the user copies into DHSChat, plus the metadata
// the audit trail (M4b) needs to answer "which prompt produced this review."
//
// Bump PROMPT_TEMPLATE_VERSION on ANY change to the assembled template text (section
// order, wording, worked examples) -- it's recorded in the audit record specifically so a
// past review's exact prompt is reconstructable later.
import { DEFAULT_PERSONA } from "./persona.js";

export const PROMPT_TEMPLATE_VERSION = "m4d-2026.07-1";

// Chunk-mode prompts (M4c) prepend a part-N-of-M preamble to [TASK] -- distinct template
// text from the single-doc prompt, so it gets its own version rather than silently sharing
// PROMPT_TEMPLATE_VERSION (m4-scope-notes.md D1: audit provenance should record which
// template actually ran). Single-doc's own version is UNCHANGED by chunk mode's existence.
export const CHUNK_PROMPT_TEMPLATE_VERSION = "m4c-chunk-2026.07-1";

// Word-count threshold above which a full-document round trip is abandoned in favor of
// chunk mode (spec §6.4). M4a only ever reports whether a document is over this line --
// chunk.js (M4c) is what actually acts on it. Kept here, not duplicated, so both stay in
// sync by construction.
export const CHUNK_WORD_THRESHOLD = 12000;

function wordCount(text) {
  return text.trim().split(/\s+/).filter(Boolean).length;
}

// chars/4 heuristic (spec §6.4) -- sufficient for a threshold warning, not a token-exact count.
function estimateTokens(text) {
  return Math.ceil(text.length / 4);
}

function buildPersonaSection(persona) {
  if (persona.isDefault) {
    return (
      "[PERSONA]\n" +
      "No custom persona was loaded -- using the built-in default persona.\n\n" +
      `Role and voice: ${persona.roleAndVoice}\n` +
      `Review priorities: ${persona.reviewPriorities}\n` +
      `Comment conventions: ${persona.commentConventions}`
    );
  }
  const lines = [
    "[PERSONA]",
    `Persona: ${persona.name}`,
    "",
    "Role and voice:",
    persona.roleAndVoice || "(not specified)",
    "",
    "Review priorities:",
    persona.reviewPriorities || "(not specified)",
  ];
  if (persona.styleExemplars && persona.styleExemplars.length) {
    lines.push("", "Style exemplars:");
    for (const ex of persona.styleExemplars) {
      if (ex.raw != null) lines.push(ex.raw);
      else lines.push(`  before: ${ex.before}\n  after:  ${ex.after}`);
    }
  }
  lines.push("", "Comment conventions:", persona.commentConventions || "(not specified)");
  return lines.join("\n");
}

// `chunk`, when present, is { index, total } (0-based index) -- M4c's chunk mode splits an
// oversized document into independent prompt/paste units (spec §6.4), and the model needs
// to be told it's only seeing one part, or it has no way to know the document continues.
function buildTaskSection(chunk) {
  const preamble = chunk
    ? `This is part ${chunk.index + 1} of ${chunk.total} of a larger document. Review ONLY ` +
      "the text below and return ONLY this part, beginning with its first line exactly.\n\n"
    : "";
  return (
    "[TASK]\n" +
    preamble +
    "Review the document below according to the persona above. Propose your edits and " +
    "comments using ONLY the CriticMarkup syntax defined next. Return the ENTIRE document " +
    "with your proposed changes expressed as CriticMarkup tokens -- never rewrite, " +
    "paraphrase, or reformat anything outside those tokens."
  );
}

// Folded from the repo's hand-tuned "DHSChat Prompt (adjusted after Test 2)" (M4d PR-3),
// which one-shot D1 (whole-paragraph insertion) on GPT-5.5 where the shipped template
// hadn't. The worked example, its WRONG counterpart, and the [WHITESPACE AND LINE BREAKS]
// rules are folded in verbatim (smart quotes converted to ASCII apostrophes; the
// test-harness-specific task line and the pre-existing-CriticMarkup language in that file
// are deliberately NOT folded here -- the latter is M6a scope, made unreachable by the
// M4d PR-2 annotation fence).
function buildCriticMarkupRulesSection() {
  return (
    "[CRITICMARKUP RULES]\n" +
    "Support exactly these five constructs:\n" +
    "  {++text++}            insertion\n" +
    "  {--text--}             deletion\n" +
    "  {~~old~>new~~}         substitution (~> is the only arrow form)\n" +
    "  {==text==}{>>comment<<} anchored comment\n" +
    "  {>>comment<<}          bare (point) comment\n\n" +
    "No nesting: no CriticMarkup token may contain another. No block-crossing: a token must " +
    "open and close within one markdown block (paragraph, heading, list item, or table cell) " +
    "-- express a multi-paragraph change as one token per paragraph.\n\n" +
    "Worked example:\n" +
    "  original: The rule shall apply to all carriers.\n" +
    "  edited:   The rule {~~shall~>must~~} apply to all {--air--} carriers.\n" +
    "            {==all carriers==}{>>Confirm scope includes indirect air carriers.<<}\n\n" +
    "Whole-paragraph insert: to insert an entirely new paragraph, put the {++...++} token " +
    "ALONE on its own line, immediately after the prior paragraph line and immediately " +
    "before the following paragraph line, with no extra blank lines added above or below it.\n\n" +
    "Worked example (structure-sensitive):\n\n" +
    "  Existing paragraph one.\n\n" +
    "  Existing paragraph two.\n" +
    "  {++This entire line is a new inserted paragraph.++}\n" +
    "  Existing paragraph three.\n\n" +
    "WRONG -- do not wrap the inserted line in blank lines of its own:\n\n" +
    "  Existing paragraph two.\n\n" +
    "  {++This entire line is a new inserted paragraph.++}\n\n" +
    "  Existing paragraph three.\n\n" +
    "[WHITESPACE AND LINE BREAKS]\n" +
    "- Do not add or remove any blank lines except where required to place the CriticMarkup " +
    "token itself.\n" +
    "- Treat the existing blank-line structure as fixed: preserve all existing blank lines " +
    "exactly as they are, unless a CriticMarkup token must appear on a line between two " +
    "existing lines.\n" +
    "- When inserting a new paragraph, do not introduce additional blank lines before or " +
    "after the {++...++} line beyond what already exists in the input.\n" +
    "- Do not add a new blank line at the end of the document. The last character before " +
    "the closing ``` fence must match the source document exactly.\n\n" +
    "Byte preservation: every character outside your CriticMarkup tokens -- including the " +
    "document's leading header comment lines and its FINAL trailing newline at the " +
    "very end -- must be returned exactly as given, unchanged."
  );
}

// M4d PR-3, F-6: matches validate.js's validateText D7 derivation exactly (headerContent,
// content-anchored to the export's own header text) instead of a static "first 3 lines"
// slice, so the prompt's quoted header and G2's own header check can never drift apart
// again. Falls back to the old line-count guess only when no sourceMap is available (e.g.
// a caller that hasn't been updated yet) -- every production call site passes one.
function deriveHeaderContent(exportedMarkdown, sourceMap) {
  const blocks = sourceMap && sourceMap.blocks;
  if (!blocks) return exportedMarkdown.split("\n").slice(0, 3).join("\n");
  const rawHeaderPrefix = blocks.length ? exportedMarkdown.slice(0, blocks[0].mdStart) : exportedMarkdown;
  return rawHeaderPrefix.replace(/\s+$/, "");
}

function buildHardConstraintsSection(persona, exportedMarkdown, chunk, sourceMap) {
  // FABLE-REVIEW (M4 milestone): re-emits the export's header lines here in addition to
  // the verbatim [DOCUMENT] embedding below -- §4 single-source vs Issue #10, two conformant
  // readings. Deferred ruling; see docs/plans/m4-scope-notes.md -> "Deferred to Fable".
  // Do not "resolve" this ad hoc.
  //
  // Only chunk 0 (or the single-doc case, chunk == null) actually carries the document's
  // header -- chunk.js's splitIntoChunks folds the header into chunk 0 and starts every
  // later chunk's own slice directly at its top-level heading (M4c, architecture doc §7).
  // Instructing a later chunk to echo header lines it was never given would be an
  // instruction the model can't satisfy, and validateText's own header check (derived from
  // that chunk's own sourceMap.blocks[0].mdStart, which is 0) doesn't require it either.
  const isFirstChunk = !chunk || chunk.index === 0;
  const lines = [
    "[HARD CONSTRAINTS]",
    "- Return the ENTIRE document below, not an excerpt or summary.",
    "- Change nothing outside your own CriticMarkup tokens.",
  ];
  if (isFirstChunk) {
    const headerLines = deriveHeaderContent(exportedMarkdown, sourceMap);
    lines.push(
      "- The document below begins with HTML comment lines (export header and legend). " +
        "These are part of the document, not metadata to skip. Your response must begin with " +
        "those exact lines, unmodified:\n" +
        headerLines
          .split("\n")
          .map((l) => "    " + l)
          .join("\n")
    );
  }
  lines.push(
    "- Text inside ⟦…⟧ and [image: …] is locked: never edit, move, or comment inside it.",
    "- No nesting. No block-crossing.",
    "- Return your entire response inside ONE fenced code block: ```markdown ... ```, with " +
      "no other output before or after it."
  );
  if (!persona.isDefault && persona.doNotTouch && persona.doNotTouch.length) {
    lines.push("- Do-not-touch rules from the persona (never edit or comment on this content):");
    for (const rule of persona.doNotTouch) lines.push(`    - ${rule}`);
  }
  return lines.join("\n");
}

function buildDocumentSection(exportedMarkdown) {
  return "[DOCUMENT]\n```markdown\n" + exportedMarkdown + "\n```";
}

// `chunk`, when present, is { index, total } (0-based `index`) identifying which chunk of a
// larger document `exportedMarkdown` is (M4c, spec §6.4). Every section here is built from
// exactly the same {persona, exportedMarkdown} a single-document call would use -- chunk
// mode doesn't restructure the template, it only (a) adds the part-N-of-M preamble to
// [TASK] and (b) drops the header-echo bullet from [HARD CONSTRAINTS] for any chunk after
// the first, which never carries the header. See CHUNK_PROMPT_TEMPLATE_VERSION's comment
// for why chunk mode gets its own promptVersion.
export function buildPrompt({ persona, exportedMarkdown, filename, chunk = null, sourceMap = null }) {
  const p = persona || DEFAULT_PERSONA;

  const sections = [
    buildPersonaSection(p),
    buildTaskSection(chunk),
    buildCriticMarkupRulesSection(),
    buildHardConstraintsSection(p, exportedMarkdown, chunk, sourceMap),
    buildDocumentSection(exportedMarkdown),
  ];
  const text = sections.join("\n\n");
  const words = wordCount(exportedMarkdown);

  return {
    text,
    tokenEstimate: estimateTokens(text),
    promptVersion: chunk ? CHUNK_PROMPT_TEMPLATE_VERSION : PROMPT_TEMPLATE_VERSION,
    documentWordCount: words,
    overThreshold: words > CHUNK_WORD_THRESHOLD,
    filename,
  };
}

export const RESPOND_PROMPT_TEMPLATE_VERSION = "m6b-respond-2026.07-1";

export function buildRespondPrompt({ persona, exportedMarkdown, filename, sourceMap = null }) {
  const p = persona || DEFAULT_PERSONA;

  const personaSection = buildPersonaSection(p);

  const taskText =
    "[TASK]\n" +
    "You are given an exported document below that contains pre-existing comments and tracked changes " +
    "represented as double-bracketed sentinels:\n" +
    "  - Revisions are labeled as ⟦R1: ...⟧, ⟦R2: ...⟧, etc.\n" +
    "  - Comments are labeled as ⟦C1: ... >> ...⟧, ⟦C2: ... >> ...⟧, etc.\n\n" +
    "You must respond to every single comment and revision. Your response must consist ONLY of a " +
    "structured response block (do not return the full document). Do not write any preamble, chatter, " +
    "or postamble outside the fenced block.";

  const grammarText =
    "[RESPOND GRAMMAR]\n" +
    "For every label C1...Cn and R1...Rn present in the document, you must output exactly one line in the " +
    "following format:\n\n" +
    "For comments [Cn]:\n" +
    "  [Cn] {>>reply text<<}                    <-- to reply to the comment\n" +
    "  [Cn] {>>[AR:resolve] reply text<<}      <-- to reply and recommend resolving/closing the comment\n\n" +
    "For revisions [Rn]:\n" +
    "  [Rn] {>>[AR:accept] rationale<<}         <-- to recommend accepting the tracked change\n" +
    "  [Rn] {>>[AR:reject] rationale<<}         <-- to recommend rejecting the tracked change\n" +
    "  [Rn] {>>[AR:discuss] rationale<<}        <-- to recommend discussing the tracked change\n\n" +
    "Constraints:\n" +
    "- You must address every single label exactly once. Zero skips, zero duplicates, zero invented labels.\n" +
    "- Write your response inside CriticMarkup comments: {>>...<<}.\n" +
    "- Rationale/reply texts must be concise, economic, and compliant with the persona above (max 1000 characters per reply).";

  const examplesText =
    "[WORKED EXAMPLES]\n" +
    "If the document contains [C1], [C2], [R1], and [R2], your response must look exactly like this:\n" +
    "```markdown\n" +
    "[C1] {>>Agreed -- will clarify this in the next draft.<<}\n" +
    "[C2] {>>[AR:resolve] Citation added in the bibliography.<<}\n" +
    "[R1] {>>[AR:accept] This aligns with the updated statutory definitions.<<}\n" +
    "[R2] {>>[AR:reject] The original phrasing is required by OMB guidelines.<<}\n" +
    "```";

  const docText = buildDocumentSection(exportedMarkdown);

  const sections = [
    personaSection,
    taskText,
    grammarText,
    examplesText,
    docText
  ];

  const text = sections.join("\n\n");
  const words = wordCount(exportedMarkdown);

  return {
    text,
    tokenEstimate: estimateTokens(text),
    promptVersion: RESPOND_PROMPT_TEMPLATE_VERSION,
    documentWordCount: words,
    filename,
  };
}
