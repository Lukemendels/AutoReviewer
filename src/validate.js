// Validation gates G1-G5 (spec §7). Runs in order, short-circuiting on the first G1-G4
// failure; G5 collects warnings and never blocks.
import { tokenize } from "./criticmarkup/grammar.js";
import { strip } from "./criticmarkup/strip.js";
import { parseEdits } from "./criticmarkup/parse.js";
import { snapBoundary, resolveRange, resolvePoint, SourceMapError } from "./sourcemap.js";
import { diffWords } from "./ui/diff.js";

function normalizeNewlines(s) {
  return s.replace(/\r\n/g, "\n");
}

// Blocks a token's [start,end) span overlaps. For zero-width point tokens (ins, bare
// comment) this never returns more than one block, so the block-crossing check below only
// ever fires for real spans (del/sub/anchored-comment highlight), matching spec §4's
// "a token must open and close within one markdown block".
function overlappingBlocks(blocks, start, end) {
  return (blocks || []).filter((b) => start < b.mdEnd && b.mdStart < end);
}

function buildRepairPrompt() {
  return (
    "Your last response modified or omitted text outside your own CriticMarkup tokens " +
    "(this includes the three leading <!-- ... --> header comment lines at the very top " +
    "of the document, which are also part of the document text). Re-emit the ENTIRE " +
    "document exactly as given -- including those header lines -- changing nothing except " +
    "your own {++...++} {--...--} {~~old~>new~~} {==...==}{>>...<<} edits. Do not " +
    "paraphrase, reformat, reorder, or drop any text outside those tokens. Return the " +
    "corrected document inside a single ```markdown fenced block, with no other output."
  );
}

function wordCount(text) {
  return text.trim().split(/\s+/).filter(Boolean).length;
}

// True if [start,end) overlaps ANY locked range, at all -- checked against the edit's
// ORIGINAL declared span, before any boundary snapping. This matters: snapBoundary trims a
// boundary that merely touches the *edge* of a locked run (correct -- see sourcemap.js),
// but if left unchecked, that same trimming can silently shrink away a small overlap that
// reaches *into* a locked range's interior (e.g. an edit whose declared span covers 2 of a
// 13-character locked placeholder), leaving an edit that looks accepted even though the
// model's own token claimed to touch locked content. Spec §7's G4 is unconditional ("no
// edit overlaps a locked range") -- it doesn't carve out an exception for overlaps small
// enough to snap away, so this check runs on the pre-snap span specifically.
function overlapsLocked(sourceMap, start, end) {
  for (const [ls, le] of sourceMap.locked || []) {
    if (start < le && ls < end) return true;
  }
  return false;
}

// Resolves one edit's anchor against the source map. Span edits (del/sub/anchored comment)
// are checked against locked ranges on their original span first (see overlapsLocked
// above), then snap and resolve strictly; point edits (ins/bare comment) use resolvePoint
// instead, since a zero-width position has no synthetic span to snap around (M2 plan
// decision 1). Returns { anchor } on success, or { gateFailure } naming the gate: a
// SourceMapError of kind "locked" is G4 (protection); kind "synthetic" -- meaning the span
// doesn't resolve to document text, whether because it never had any (a genuinely
// zero-width/point case) or because its interior can't be covered -- is G3 (anchor
// resolution), except the explicit "collapsed to empty on snap" case, which is G4's
// "entirely synthetic" per the plan's exact wording.
function resolveEditAnchor(edit, sourceMap) {
  if (edit.type === "ins" || (edit.type === "comment" && !edit.anchored)) {
    const pos = edit.type === "ins" ? edit.mdPos : edit.mdPos;
    try {
      return { anchor: resolvePoint(sourceMap, pos) };
    } catch (err) {
      if (!(err instanceof SourceMapError)) throw err;
      return { gateFailure: { gate: err.kind === "locked" ? "G4" : "G3", message: err.message } };
    }
  }

  if (overlapsLocked(sourceMap, edit.mdStart, edit.mdEnd)) {
    return { gateFailure: { gate: "G4", message: "edit overlaps a locked range" } };
  }
  const [snapStart, snapEnd] = snapBoundary(sourceMap, edit.mdStart, edit.mdEnd);
  if (snapStart >= snapEnd) {
    return { gateFailure: { gate: "G4", message: "edit is entirely synthetic (no document text remains after boundary snapping)" } };
  }
  try {
    return { anchor: resolveRange(sourceMap, snapStart, snapEnd) };
  } catch (err) {
    if (!(err instanceof SourceMapError)) throw err;
    return { gateFailure: { gate: err.kind === "locked" ? "G4" : "G3", message: err.message } };
  }
}

export function validate({ responseMarkdown, exportedMarkdown, sourceMap, largestDeletionWords = 50 }) {
  const response = normalizeNewlines(responseMarkdown);
  const exported = normalizeNewlines(exportedMarkdown);

  // The header comment block (spec §5.1) is required reading for the model and literally
  // contains a CriticMarkup legend as example syntax (e.g. "{==highlighted==} {>>comment<<}")
  // -- without an exemption, every single document's own header would be misparsed as real
  // tokens. Exempt it from opener-scanning (see grammar.js's tokenize() doc comment); G2's
  // full byte-equality check still catches any actual drift there regardless. A valid edit
  // can legitimately anchor immediately after the last block's own content (e.g. a bare
  // point-comment on the final sentence), so unlike the header there's no similarly-safe
  // trailing region to exempt -- the far rarer risk of a human-authored comment's text in
  // the "## Unanchored comments" trailer accidentally containing brace sequences isn't
  // addressed here.
  const blocks = sourceMap.blocks || [];
  const tokenizeOpts = {
    skipBefore: blocks.length ? blocks[0].mdStart : exported.length,
  };

  // G1 -- grammar
  const tokenized = tokenize(response, tokenizeOpts);
  if (!tokenized.ok) {
    return { ok: false, gate: "G1", message: tokenized.error.message, detail: { rawStart: tokenized.error.rawStart } };
  }
  for (const t of tokenized.tokens) {
    if (t.strippedEnd <= t.strippedStart) continue; // zero-width point tokens can't cross a block
    if (overlappingBlocks(sourceMap.blocks, t.strippedStart, t.strippedEnd).length > 1) {
      return {
        ok: false,
        gate: "G1",
        message: `a CriticMarkup token crosses a paragraph/block boundary (position ${t.strippedStart}-${t.strippedEnd})`,
        detail: { rawStart: t.rawStart },
      };
    }
  }

  // G2 -- fidelity (the fabrication gate)
  const strippedResponse = strip(response, tokenizeOpts);
  if (strippedResponse !== exported) {
    return {
      ok: false,
      gate: "G2",
      message: "the response's underlying text does not byte-match the exported document outside CriticMarkup tokens",
      diff: diffWords(strippedResponse, exported),
      repairPrompt: buildRepairPrompt(),
    };
  }

  // G3/G4 -- anchor resolution + protection
  const edits = parseEdits(response, tokenizeOpts);
  const resolvedEdits = [];
  for (const edit of edits) {
    const { anchor, gateFailure } = resolveEditAnchor(edit, sourceMap);
    if (gateFailure) {
      return { ok: false, gate: gateFailure.gate, message: gateFailure.message, detail: { rawStart: edit.rawStart } };
    }
    resolvedEdits.push({ ...edit, anchor });
  }

  // G5 -- sanity report (warnings only, never blocks)
  const warnings = [];
  for (const edit of resolvedEdits) {
    if ((edit.type === "del" || edit.type === "sub") && wordCount(edit.oldText) > largestDeletionWords) {
      warnings.push({
        code: "oversized-deletion",
        message: `a ${edit.type === "sub" ? "substitution's old text" : "deletion"} is ${wordCount(edit.oldText)} words (> ${largestDeletionWords})`,
        rawStart: edit.rawStart,
      });
    }
  }
  const byCommentText = new Map();
  for (const edit of resolvedEdits) {
    if (edit.type !== "comment") continue;
    const key = edit.commentText.trim();
    if (!byCommentText.has(key)) byCommentText.set(key, []);
    byCommentText.get(key).push(edit);
  }
  for (const [text, group] of byCommentText) {
    if (group.length > 1) {
      warnings.push({
        code: "duplicate-comment",
        message: `${group.length} comments share identical text: "${text}"`,
        rawStarts: group.map((e) => e.rawStart),
      });
    }
  }

  const counts = { ins: 0, del: 0, sub: 0, comment: 0 };
  for (const edit of resolvedEdits) counts[edit.type]++;

  return { ok: true, edits: resolvedEdits, warnings, counts };
}
