// The one accumulator for "the plain text of a docx run" (M4d PR-2, fixing F-2).
// export.js and inject.js each computed this independently and disagreed: export counted
// w:tab/w:br/w:cr as extra characters while inject's splitRun counted only w:t, so any run
// containing one of those children made export's source-map offsets unreachable by inject
// (crashes like "splitRun: invalid range [34,34) for run text length of 31"). Both now call
// this single function, so their lengths agree by construction rather than by convention.
//
// Byte-identical to export.js's pre-M4d behavior on purpose -- this is a refactor, not a
// behavior change; the M4d structural preflight fence (src/load.js) rejects any document
// containing w:br/w:cr/w:tab/w:delText before inject ever needs to split one of these runs.
export function runPlainText(run) {
  let s = "";
  for (const c of run.children) {
    if (c.localName === "t" || c.localName === "delText") s += c.textContent;
    else if (c.localName === "tab") s += " ";
    else if (c.localName === "br" || c.localName === "cr") s += "  \n";
  }
  return s;
}
