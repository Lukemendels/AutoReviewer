// markup -> ordered edit list with md offsets. Offsets are in the *stripped* coordinate
// space (== exportedMarkdown's coordinate space once G2 confirms alignment) -- see
// grammar.js's tokenize() for why each token carries both a raw and a stripped position.
import { tokenize } from "./grammar.js";

// D1 (M3b plan): whole-paragraph detection, applied only to ins/del tokens (spec doesn't
// define whole-paragraph substitution/comment). A token is whole-paragraph if it is alone
// on its own raw-text line -- spec §4's own definition ("a line consisting solely of one
// token") applied literally: everything from the nearest single "\n" (or string start)
// up to rawStart is blank, and everything from rawEnd to the nearest single "\n" (or
// string end) is blank.
//
// This is a single "\n" boundary, not "\n\n" -- deliberately, even though blocks are
// always separated by "\n\n" in a pristine export. G2's byte-equality gate forces the
// only valid raw-text shape for "insert a new paragraph between existing block A and
// block B" to be A's text + "\n" + the token + "\n" + B's text: the token sits spliced
// *inside* the original "\n\n" gap, splitting it into two single "\n"s (strip() removes
// the token, reassembling exactly "\n\n" -- the double-newline never survives intact in
// the raw response next to the token itself, since adding a fresh "\n\n" alongside the
// existing one would reconstruct four newlines and fail G2). A "\n\n" search would never
// match this real, G2-valid construction; a single-"\n" search does, and still correctly
// rejects a token that merely shares an internal single-newline (hard-break) sub-line
// within one paragraph's own content, since real (non-blank) text sits on that line too.
//
// Two (or more) consecutive whole-paragraph tokens sharing ONE original "\n\n" gap need
// zero raw characters between them for the same G2 reason (that gap only has 2 newlines
// to spend total, one on each outer edge, none left over between the tokens themselves)
// -- so the blank-check treats an immediately-adjacent token (rawEnd/rawStart touching,
// no gap at all) as transparent and walks past it to keep looking for the real line
// boundary, rather than seeing its raw text and concluding "not blank".
function isWholeParagraph(markdown, tokens, index) {
  let pos = tokens[index].rawStart;
  let i = index;
  for (;;) {
    const prevNl = markdown.lastIndexOf("\n", pos - 1);
    const boundary = prevNl === -1 ? 0 : prevNl + 1;
    if (/^\s*$/.test(markdown.slice(boundary, pos))) break;
    if (i > 0 && tokens[i - 1].rawEnd === pos) {
      pos = tokens[i - 1].rawStart;
      i--;
      continue;
    }
    return false;
  }

  pos = tokens[index].rawEnd;
  i = index;
  for (;;) {
    const nextNl = markdown.indexOf("\n", pos);
    const boundary = nextNl === -1 ? markdown.length : nextNl;
    if (/^\s*$/.test(markdown.slice(pos, boundary))) break;
    if (i < tokens.length - 1 && tokens[i + 1].rawStart === pos) {
      pos = tokens[i + 1].rawEnd;
      i++;
      continue;
    }
    return false;
  }
  return true;
}

export function parseEdits(markdown, opts) {
  const result = tokenize(markdown, opts);
  if (!result.ok) {
    throw new Error(`parseEdits: invalid CriticMarkup grammar -- ${result.error.message} (at ${result.error.rawStart})`);
  }
  return result.tokens.map((t, index) => {
    if (t.type === "ins") {
      return {
        type: "ins",
        mdPos: t.strippedStart,
        newText: t.text,
        rawStart: t.rawStart,
        rawEnd: t.rawEnd,
        wholeParagraph: isWholeParagraph(markdown, result.tokens, index),
      };
    }
    if (t.type === "del") {
      return {
        type: "del",
        mdStart: t.strippedStart,
        mdEnd: t.strippedEnd,
        oldText: t.text,
        rawStart: t.rawStart,
        rawEnd: t.rawEnd,
        wholeParagraph: isWholeParagraph(markdown, result.tokens, index),
      };
    }
    if (t.type === "sub") {
      return { type: "sub", mdStart: t.strippedStart, mdEnd: t.strippedEnd, oldText: t.oldText, newText: t.newText, rawStart: t.rawStart, rawEnd: t.rawEnd };
    }
    // comment
    return t.anchored
      ? {
          type: "comment",
          anchored: true,
          mdStart: t.strippedStart,
          mdEnd: t.strippedEnd,
          highlightText: t.highlightText,
          commentText: t.commentText,
          rawStart: t.rawStart,
          rawEnd: t.rawEnd,
        }
      : { type: "comment", anchored: false, mdPos: t.strippedStart, commentText: t.commentText, rawStart: t.rawStart, rawEnd: t.rawEnd };
  });
}
