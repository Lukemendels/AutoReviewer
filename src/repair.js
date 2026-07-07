// Mistake-specific repair prompt composer (M4 doc §6.1, M4b build plan §1). Pure string
// function: given a blocking validate() failure, the raw response text that produced it,
// and how many times this same gate has failed in a row, returns the exact text a human
// pastes back to the model to fix ONLY that mistake.
//
// `responseText` is the same string passed as validate()'s `responseMarkdown` (context.response
// in app.js) -- NOT the exported document. This matters for G1/G3/G4: their `detail.rawStart`
// is an offset into that raw response (see grammar.js's tokenize() and criticmarkup/parse.js's
// parseEdits(), both called as `fn(response, tokenizeOpts)` in validate.js), so quoting context
// around it means slicing the response, not the export.
//
// G-1 (M4b session brief): every message reuses `failure.message` verbatim -- it IS the
// validator's teaching text -- rather than re-typing a paraphrase of it.

function contextAround(text, rawStart, pad = 80) {
  const start = Math.max(0, rawStart - pad);
  const end = Math.min(text.length, rawStart + pad);
  return (start > 0 ? "…" : "") + text.slice(start, end) + (end < text.length ? "…" : "");
}

function composeG1(failure, responseText) {
  const snippet = contextAround(responseText, failure.detail.rawStart);
  return (
    `${failure.message}\n` +
    `Near: "${snippet}"\n` +
    "A CriticMarkup token must open and close inside ONE paragraph/block, and an insertion's " +
    "own new text must never contain a newline character. Split a multi-paragraph change into " +
    "one token per block, e.g.:\n" +
    "  ...end of paragraph one.{--...--}\n" +
    "  {--Paragraph two...--}\n" +
    "and use the whole-paragraph insert shape -- the {++...++} token ALONE on its own line -- " +
    "instead of a newline character inside a token's own text."
  );
}

function composeG2(failure) {
  const fd = failure.firstDivergence;
  return (
    `${failure.message}\n` +
    `Your response diverges at offset ${fd.offset}.\n` +
    `You wrote:    "${fd.afterA}"\n` +
    `Expected:     "${fd.afterB}"\n` +
    "Everything outside your own CriticMarkup tokens must be returned byte-for-byte -- including " +
    "the document's three leading header comment lines and its final trailing newline. Re-emit " +
    "the full response with that exact text restored; to make an intentional wording change, " +
    "wrap it in a token instead, e.g. {~~old~>new~~}."
  );
}

function composeG3(failure, responseText) {
  const snippet = contextAround(responseText, failure.detail.rawStart);
  return (
    `${failure.message}\n` +
    `Near: "${snippet}"\n` +
    "The edit's span falls on markdown the exporter invented (a heading \"# \", a bullet \"- \", " +
    "or emphasis \"**\"), not on real document text. Move the token so it wraps only the actual " +
    "words, e.g. {~~old~>new~~} inside the heading text, not around the \"# \" prefix."
  );
}

function composeG4(failure, responseText) {
  const snippet = contextAround(responseText, failure.detail.rawStart);
  return (
    `${failure.message}\n` +
    `Near: "${snippet}"\n` +
    "This edit touches locked content (text inside ⟦…⟧ or an [image: …] placeholder), " +
    "which must never be edited, moved, or commented on. Remove that token, or re-scope it to " +
    "text outside the locked range."
  );
}

const RESTART_BLOCK =
  "You've hit the same issue twice. Rather than patch again, start clean: re-copy the full " +
  "prompt below and paste a fresh response -- partial fixes tend to compound structural " +
  "mistakes.\n\n---\n";

export function composeRepair(failure, responseText, attemptCount) {
  let body;
  switch (failure.gate) {
    case "G1":
      body = composeG1(failure, responseText);
      break;
    case "G2":
      body = composeG2(failure, responseText);
      break;
    case "G3":
      body = composeG3(failure, responseText);
      break;
    case "G4":
      body = composeG4(failure, responseText);
      break;
    default:
      body = failure.message;
  }
  return attemptCount >= 2 ? RESTART_BLOCK + body : body;
}
