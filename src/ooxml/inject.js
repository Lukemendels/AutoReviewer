// Accepted edits -> mutated document.xml DOM (spec §9.1; M3b plan's run-location +
// run-splitting algorithm). Built incrementally, primitive-first: this file grows a
// function at a time, each landing with the test file that drove it (see
// docs/plans/m3b-plan.md's Implementation protocol).
import { NS, kid, kids } from "./parse.js";

const XML_NS = "http://www.w3.org/XML/1998/namespace";

/* ------------------------------------------------------------------ *
 * Locating the target w:p from bodyPath (replicates walkBodyTracked's exact traversal:
 * every body child increments the counter, regardless of type).
 * ------------------------------------------------------------------ */
export function locateParagraph(body, bodyPath) {
  let counter = 0;
  for (const el of body.children) {
    if (counter === bodyPath[0]) {
      if (bodyPath.length === 1) return el;
      if (bodyPath.length === 4) {
        const tr = kids(el, "tr")[bodyPath[1]];
        const tc = tr && kids(tr, "tc")[bodyPath[2]];
        const p = tc && kids(tc, "p")[bodyPath[3]];
        if (!p) throw new Error(`locateParagraph: table cell path [${bodyPath}] not found`);
        return p;
      }
      throw new Error(`locateParagraph: unsupported bodyPath length ${bodyPath.length}`);
    }
    counter++;
  }
  throw new Error(`locateParagraph: bodyPath [${bodyPath}] not found in body`);
}

/* ------------------------------------------------------------------ *
 * Locating the target w:r from runIndex (replicates buildSegments' + collectRunsTextIndexed's
 * exact counting order in export.js: every w:r increments the counter, whether it's a direct
 * child of p or nested inside w:ins/w:del/w:hyperlink/w:smartTag/w:fldSimple/w:sdt, but only a
 * direct child of p is ever a candidate return value.)
 * ------------------------------------------------------------------ */
const RUN_WRAPPERS = new Set(["ins", "del", "hyperlink", "smartTag"]);

function walkForRun(el, state, direct) {
  for (const c of el.children) {
    if (state.found) return;
    const ln = c.localName;
    if (ln === "r") {
      if (direct && state.count === state.target) state.found = c;
      state.count++;
    } else if (RUN_WRAPPERS.has(ln)) {
      walkForRun(c, state, false);
    } else if (ln === "fldSimple") {
      walkForRun(c, state, false);
    } else if (ln === "sdt") {
      const content = kid(c, "sdtContent");
      if (content) walkForRun(content, state, false);
    }
  }
}

export function locateRun(p, runIndex) {
  const state = { count: 0, target: runIndex, found: null };
  walkForRun(p, state, true);
  return state.found;
}

/* ------------------------------------------------------------------ *
 * splitRun: the peel-from-the-right primitive. Splits `run`'s w:t text at
 * [charStart, charEnd) into up to three fresh <w:r> siblings -- before/core/after --
 * inserted in `run`'s place (insertBefore + removeChild, not replaceWith). Each piece
 * deep-copies run's rPr and gets xml:space="preserve" on its own w:t iff that piece's own
 * text starts or ends with whitespace. charStart === charEnd is a valid zero-width split
 * (the insertion-point case): before/after are produced with no core.
 * ------------------------------------------------------------------ */
function runPlainText(run) {
  let s = "";
  for (const c of run.children) if (c.localName === "t") s += c.textContent;
  return s;
}

function cloneRunPiece(doc, run, text) {
  const newRun = doc.createElementNS(NS.w, "w:r");
  const rPr = kid(run, "rPr");
  if (rPr) newRun.appendChild(rPr.cloneNode(true));
  const t = doc.createElementNS(NS.w, "w:t");
  if (/^\s|\s$/.test(text)) t.setAttributeNS(XML_NS, "xml:space", "preserve");
  t.appendChild(doc.createTextNode(text));
  newRun.appendChild(t);
  return newRun;
}

export function splitRun(doc, run, charStart, charEnd) {
  const text = runPlainText(run);
  if (charStart < 0 || charEnd < charStart || charEnd > text.length) {
    throw new Error(`splitRun: invalid range [${charStart},${charEnd}) for run text of length ${text.length}`);
  }
  const parent = run.parentNode;
  const reference = run.nextSibling;

  const before = charStart > 0 ? cloneRunPiece(doc, run, text.slice(0, charStart)) : null;
  const core = charEnd > charStart ? cloneRunPiece(doc, run, text.slice(charStart, charEnd)) : null;
  const after = charEnd < text.length ? cloneRunPiece(doc, run, text.slice(charEnd)) : null;

  for (const piece of [before, core, after]) {
    if (piece) parent.insertBefore(piece, reference);
  }
  parent.removeChild(run);

  return { parent, reference, before, core, after };
}

export function injectEdits(_documentXmlDoc, _acceptedEdits, _sourceMap) {
  throw new Error("ooxml/inject: injectEdits not yet implemented (M3b in progress)");
}
