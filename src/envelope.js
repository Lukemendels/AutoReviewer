// Response envelope (spec §6.3, §6.1 step 3-4): pasted reply -> candidate fenced blocks.
// Pure. Models routinely preface a reply with chatter ("Sure, here's the revised
// document:") before the actual fenced block, and DHSChat's copy-out can occasionally
// include more than one fence -- this module finds every fenced block in the paste and
// lets the caller (app.js) decide what to do with the result; it never guesses which one
// is "the" answer when more than one is plausible.
//
// Verbatim rule: a candidate's `content` is the EXACT bytes between the fence's opening
// line's trailing newline and its closing "```" -- never trimmed, never appended to. If
// the model's response is missing its final newline, that's a G2 failure for the repair
// prompt to name (M4b), not something this extractor silently papers over.

// Scans for every ```-fenced block in `text`, in document order. An opening fence is
// "```" followed by an optional language tag then a newline; its content ends at the next
// literal "```", wherever that falls -- including immediately after the last content
// character, with no newline in between (the fence-with-no-trailing-newline regression
// case). An unterminated opening fence (no closing "```" anywhere after it) is not a fence
// at all for our purposes and scanning stops there, matching "chatter, then the real fence"
// rather than "chatter that happens to contain three backticks".
function findFences(text) {
  const fences = [];
  let i = 0;
  while (i < text.length) {
    const open = text.indexOf("```", i);
    if (open === -1) break;
    const lineEnd = text.indexOf("\n", open + 3);
    if (lineEnd === -1) break;
    const lang = text.slice(open + 3, lineEnd).trim();
    const contentStart = lineEnd + 1;
    const close = text.indexOf("```", contentStart);
    if (close === -1) break;
    fences.push({ start: open, end: close + 3, lang, content: text.slice(contentStart, close) });
    i = close + 3;
  }
  return fences;
}

// `exportedLength` is the exported markdown's own length; a candidate must be at least
// half that to plausibly be a full-document echo rather than an excerpt the model quoted
// while explaining itself.
export function extractCandidates(pastedText, { exportedLength = 0 } = {}) {
  const fences = findFences(pastedText);

  if (!fences.length) {
    return {
      fences: [],
      candidates: [{ content: pastedText, fenceInfo: null }],
      noFencesFound: true,
    };
  }

  const threshold = exportedLength * 0.5;
  const documentSized = fences.filter((f) => f.content.length >= threshold);
  // Presented last-found-first: the model's actual final answer is usually the last fence
  // in the paste, so that's the natural first option in a picker.
  const candidates = documentSized
    .slice()
    .reverse()
    .map((f) => ({ content: f.content, fenceInfo: f }));

  return { fences, candidates, noFencesFound: false };
}

// Convenience for callers: null unless exactly one plausible candidate exists (never
// guesses between multiple -- spec §6.3, "never a silent guess").
export function selectSingleCandidate(pastedText, opts) {
  const { candidates } = extractCandidates(pastedText, opts);
  return candidates.length === 1 ? candidates[0] : null;
}
