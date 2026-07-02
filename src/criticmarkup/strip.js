// markup -> base text, for the G2 fidelity gate (spec §4/§7). Built on grammar.js's
// tokens rather than regex-only, so malformed input is rejected (G1's job) instead of
// silently mis-stripped.
import { tokenize } from "./grammar.js";

export function strip(markdown, opts) {
  const result = tokenize(markdown, opts);
  if (!result.ok) {
    throw new Error(`strip: invalid CriticMarkup grammar -- ${result.error.message} (at ${result.error.rawStart})`);
  }
  let out = "";
  let raw = 0;
  for (const t of result.tokens) {
    out += markdown.slice(raw, t.rawStart);
    if (t.type === "ins") out += "";
    else if (t.type === "del") out += t.text;
    else if (t.type === "sub") out += t.oldText;
    else if (t.type === "comment") out += t.anchored ? t.highlightText : "";
    raw = t.rawEnd;
  }
  out += markdown.slice(raw);
  return out;
}
