// Versioned prompt assembly (spec §6.2, §6.1 step 2). Pure: takes a persona + the export
// leg's output and returns the exact text the user copies into DHSChat, plus the metadata
// the audit trail (M4b) needs to answer "which prompt produced this review."
//
// Bump PROMPT_TEMPLATE_VERSION on ANY change to the assembled template text (section
// order, wording, worked examples) -- it's recorded in the audit record specifically so a
// past review's exact prompt is reconstructable later.
import { DEFAULT_PERSONA } from "./persona.js";

export const PROMPT_TEMPLATE_VERSION = "m4a-2026.07-1";

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
    "ALONE on its own line, between the two existing paragraphs it separates -- never place " +
    "a newline character inside a mid-paragraph insertion's own text. Worked example:\n" +
    "  Existing paragraph one.\n" +
    "  {++This entire line is a new inserted paragraph.++}\n" +
    "  Existing paragraph two.\n\n" +
    "Byte preservation: every character outside your CriticMarkup tokens -- including the " +
    "document's three leading header comment lines and its FINAL trailing newline at the " +
    "very end -- must be returned exactly as given, unchanged."
  );
}

function buildHardConstraintsSection(persona, exportedMarkdown, chunk) {
  // FABLE-REVIEW (M4 milestone): re-emits the export's 3 header lines here in addition to
  // the verbatim [DOCUMENT] embedding below -- §4 single-source vs Issue #10, two conformant
  // readings. Deferred ruling; see docs/plans/m4-scope-notes.md -> "Deferred to Fable".
  // Do not "resolve" this ad hoc.
  //
  // Only chunk 0 (or the single-doc case, chunk == null) actually carries the document's
  // 3-line header -- chunk.js's splitIntoChunks folds the header into chunk 0 and starts
  // every later chunk's own slice directly at its top-level heading (M4c, architecture doc
  // §7). Instructing a later chunk to echo 3 header lines it was never given would be an
  // instruction the model can't satisfy, and validateText's own header check (derived from
  // that chunk's own sourceMap.blocks[0].mdStart, which is 0) doesn't require it either.
  const isFirstChunk = !chunk || chunk.index === 0;
  const lines = [
    "[HARD CONSTRAINTS]",
    "- Return the ENTIRE document below, not an excerpt or summary.",
    "- Change nothing outside your own CriticMarkup tokens.",
  ];
  if (isFirstChunk) {
    const headerLines = exportedMarkdown.split("\n").slice(0, 3).join("\n");
    lines.push(
      "- The document below begins with three HTML comment lines (export header and legend). " +
        "These are part of the document, not metadata to skip. Your response must begin with " +
        "those exact three lines, unmodified:\n" +
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
export function buildPrompt({ persona, exportedMarkdown, filename, chunk = null }) {
  const p = persona || DEFAULT_PERSONA;

  const sections = [
    buildPersonaSection(p),
    buildTaskSection(chunk),
    buildCriticMarkupRulesSection(),
    buildHardConstraintsSection(p, exportedMarkdown, chunk),
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
