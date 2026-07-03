// markup -> ordered edit list with md offsets. Offsets are in the *stripped* coordinate
// space (== exportedMarkdown's coordinate space once G2 confirms alignment) -- see
// grammar.js's tokenize() for why each token carries both a raw and a stripped position.
import { tokenize } from "./grammar.js";

export function parseEdits(markdown, opts) {
  const result = tokenize(markdown, opts);
  if (!result.ok) {
    throw new Error(`parseEdits: invalid CriticMarkup grammar -- ${result.error.message} (at ${result.error.rawStart})`);
  }
  return result.tokens.map((t) => {
    if (t.type === "ins") {
      return { type: "ins", mdPos: t.strippedStart, newText: t.text, rawStart: t.rawStart, rawEnd: t.rawEnd };
    }
    if (t.type === "del") {
      return { type: "del", mdStart: t.strippedStart, mdEnd: t.strippedEnd, oldText: t.text, rawStart: t.rawStart, rawEnd: t.rawEnd };
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
