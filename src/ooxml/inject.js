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

/* ------------------------------------------------------------------ *
 * Construction helpers for the wrapping/insertion step (spec §9.1 steps 3-6).
 * ------------------------------------------------------------------ */
function buildRunWithText(doc, rPrEl, text) {
  const r = doc.createElementNS(NS.w, "w:r");
  if (rPrEl) r.appendChild(rPrEl.cloneNode(true));
  const t = doc.createElementNS(NS.w, "w:t");
  if (/^\s|\s$/.test(text)) t.setAttributeNS(XML_NS, "xml:space", "preserve");
  t.appendChild(doc.createTextNode(text));
  r.appendChild(t);
  return r;
}

function setRevisionAttrs(el, { id, author, date }) {
  el.setAttributeNS(NS.w, "w:id", String(id));
  el.setAttributeNS(NS.w, "w:author", author);
  el.setAttributeNS(NS.w, "w:date", date);
}

// Allocates the next revision id from the shared counter, bundled with author/date --
// used by both the ordinary and whole-paragraph edit paths.
function nextRevFor(revState) {
  return { id: revState.next++, author: revState.author, date: revState.date };
}

// Wraps `runEl` (already positioned in the DOM, e.g. a splitRun `core` piece) in a fresh
// <w:del>, renaming its w:t -> w:delText in place (spec §9.1 step 3). Mutates runEl's own
// position: replaced in its parent by the new <w:del> wrapper.
function wrapAsDel(doc, runEl, rev) {
  const t = kid(runEl, "t");
  if (t) {
    const delText = doc.createElementNS(NS.w, "w:delText");
    const space = t.getAttributeNS(XML_NS, "space");
    if (space) delText.setAttributeNS(XML_NS, "xml:space", space);
    while (t.firstChild) delText.appendChild(t.firstChild);
    runEl.insertBefore(delText, t);
    runEl.removeChild(t);
  }
  const delEl = doc.createElementNS(NS.w, "w:del");
  setRevisionAttrs(delEl, rev);
  const parent = runEl.parentNode;
  parent.insertBefore(delEl, runEl);
  parent.removeChild(runEl);
  delEl.appendChild(runEl);
  return delEl;
}

// Wraps `runEl` in a fresh <w:ins>, unchanged otherwise (spec §9.1 step 4).
function wrapAsIns(doc, runEl, rev) {
  const insEl = doc.createElementNS(NS.w, "w:ins");
  setRevisionAttrs(insEl, rev);
  insEl.appendChild(runEl);
  return insEl;
}

// Inserts `node` at the position a splitRun() call's (now-consumed) core would have
// occupied: immediately after whatever `before`/prior insertions left in place, and
// immediately before `after` (or the captured reference node if there's no `after`).
// Callable multiple times against the same splitResult to build up an ordered sequence
// (each subsequent node lands immediately before the same reference, i.e. right after the
// previous one) -- used for comment marker pairs/triples that land at one point.
function insertAtGap(splitResult, node) {
  const target = splitResult.after || splitResult.reference;
  splitResult.parent.insertBefore(node, target);
}

function buildCommentMarker(doc, localName, id) {
  const el = doc.createElementNS(NS.w, "w:" + localName);
  el.setAttributeNS(NS.w, "w:id", String(id));
  return el;
}
function buildCommentReferenceRun(doc, id) {
  const r = doc.createElementNS(NS.w, "w:r");
  const ref = doc.createElementNS(NS.w, "w:commentReference");
  ref.setAttributeNS(NS.w, "w:id", String(id));
  r.appendChild(ref);
  return r;
}

/* ------------------------------------------------------------------ *
 * Revision/comment id allocation: single counters, each seeded from the pre-mutation
 * document's own existing max id + 1 (spec §9.1 step 8; separate namespaces per the
 * M3b plan's comment-id note).
 * ------------------------------------------------------------------ */
function maxIdAmong(documentXmlDoc, localNames) {
  let max = 0; // no existing ids -> counter starts at 1, matching a fresh document's convention
  for (const el of documentXmlDoc.getElementsByTagName("*")) {
    if (!localNames.has(el.localName)) continue;
    const idAttr = el.getAttributeNS(NS.w, "id") ?? el.getAttribute("w:id");
    const id = parseInt(idAttr, 10);
    if (Number.isFinite(id) && id > max) max = id;
  }
  return max;
}
const REVISION_TAGS = new Set(["ins", "del"]);
const COMMENT_MARKER_TAGS = new Set(["commentRangeStart", "commentRangeEnd", "commentReference"]);

function findBody(documentXmlDoc) {
  const body = [...documentXmlDoc.documentElement.children].find((c) => c.localName === "body");
  if (!body) throw new Error("injectEdits: no document body found");
  return body;
}

function anchorBodyPath(anchor) {
  if (Array.isArray(anchor)) return anchor.length ? anchor[0].bodyPath : null;
  return anchor ? anchor.bodyPath : null;
}
function isWholeParagraphAnchor(anchor) {
  return !Array.isArray(anchor) && !!anchor && (anchor.kind === "wholeParagraphDelete" || anchor.kind === "paragraphBoundary");
}

/* ------------------------------------------------------------------ *
 * Per-paragraph ordinary-edit application: builds a set of per-run "tasks" from the
 * paragraph's accepted+resolved edits, then applies them run by run via the peel-from-
 * the-right mechanic (splitRun), sorted right-to-left by charStart (ties broken by
 * charEnd descending -- the wider span at a shared boundary point is spliced first,
 * matching the M3b plan's run-splitting algorithm section).
 * ------------------------------------------------------------------ */
function applyOrdinaryEdits(doc, p, edits, revState, commentState) {
  const tasksByRun = new Map();
  function pushTask(runIndex, task) {
    if (!tasksByRun.has(runIndex)) tasksByRun.set(runIndex, []);
    tasksByRun.get(runIndex).push(task);
  }

  // Revision/comment ids are allocated here, eagerly, in `edits`' own (document) order --
  // not lazily inside each task's `run` closure -- so id numbering reads top-to-bottom in
  // the Review pane. The plan notes id *ordering* isn't load-bearing (only uniqueness is),
  // but document order is trivial to get here and much friendlier to a human reviewer than
  // the right-to-left order tasks actually execute in below.
  function nextRev() {
    return nextRevFor(revState);
  }

  for (const edit of edits) {
    if (edit.type === "ins") {
      const a = edit.anchor;
      const rev = nextRev();
      pushTask(a.runIndex, {
        charStart: a.charOffset,
        charEnd: a.charOffset,
        run: (rPr, splitResult) => {
          const insRun = buildRunWithText(doc, rPr, edit.newText);
          insertAtGap(splitResult, wrapAsIns(doc, insRun, rev));
        },
      });
    } else if (edit.type === "del") {
      for (const triple of edit.anchor) {
        const rev = nextRev();
        pushTask(triple.runIndex, {
          charStart: triple.charStart,
          charEnd: triple.charEnd,
          run: (rPr, splitResult) => {
            wrapAsDel(doc, splitResult.core, rev);
          },
        });
      }
    } else if (edit.type === "sub") {
      const triples = edit.anchor;
      for (const triple of triples) {
        const rev = nextRev();
        pushTask(triple.runIndex, {
          charStart: triple.charStart,
          charEnd: triple.charEnd,
          run: (rPr, splitResult) => {
            wrapAsDel(doc, splitResult.core, rev);
          },
        });
      }
      const last = triples[triples.length - 1];
      const insRev = nextRev();
      pushTask(last.runIndex, {
        charStart: last.charEnd,
        charEnd: last.charEnd,
        run: (rPr, splitResult) => {
          const insRun = buildRunWithText(doc, rPr, edit.newText);
          insertAtGap(splitResult, wrapAsIns(doc, insRun, insRev));
        },
      });
    } else if (edit.type === "comment" && edit.anchored) {
      const commentId = commentState.next++;
      commentState.newComments.push({ id: commentId, author: revState.author, date: revState.date, text: edit.commentText });
      const triples = edit.anchor;
      const first = triples[0];
      const last = triples[triples.length - 1];
      pushTask(first.runIndex, {
        charStart: first.charStart,
        charEnd: first.charStart,
        run: (_rPr, splitResult) => {
          insertAtGap(splitResult, buildCommentMarker(doc, "commentRangeStart", commentId));
        },
      });
      pushTask(last.runIndex, {
        charStart: last.charEnd,
        charEnd: last.charEnd,
        run: (_rPr, splitResult) => {
          insertAtGap(splitResult, buildCommentMarker(doc, "commentRangeEnd", commentId));
          insertAtGap(splitResult, buildCommentReferenceRun(doc, commentId));
        },
      });
    } else if (edit.type === "comment" && !edit.anchored) {
      const commentId = commentState.next++;
      commentState.newComments.push({ id: commentId, author: revState.author, date: revState.date, text: edit.commentText });
      const a = edit.anchor;
      pushTask(a.runIndex, {
        charStart: a.charOffset,
        charEnd: a.charOffset,
        run: (_rPr, splitResult) => {
          insertAtGap(splitResult, buildCommentMarker(doc, "commentRangeStart", commentId));
          insertAtGap(splitResult, buildCommentMarker(doc, "commentRangeEnd", commentId));
          insertAtGap(splitResult, buildCommentReferenceRun(doc, commentId));
        },
      });
    } else {
      throw new Error(`injectEdits: unrecognized edit type "${edit.type}"`);
    }
  }

  // Runs are processed highest-runIndex-first (right-to-left across the whole paragraph,
  // not just within one run -- plan: "in practice this falls out automatically by
  // processing runs from highest runIndex to lowest"). This isn't just a nicety: locateRun
  // re-counts every <w:r> from the paragraph's start on each call, and splitting an
  // earlier (lower-index) run adds extra sibling nodes *before* any later run -- corrupting
  // that later run's own from-scratch count if it hasn't been located yet. Processing
  // right-to-left guarantees every run is located while everything to its own left is
  // still in its pristine, unsplit shape.
  const runIndices = [...tasksByRun.keys()].sort((a, b) => b - a);
  for (const runIndex of runIndices) {
    const tasks = tasksByRun.get(runIndex);
    const originalRun = locateRun(p, runIndex);
    if (!originalRun) throw new Error(`injectEdits: runIndex ${runIndex} not found in paragraph`);
    const rPr = kid(originalRun, "rPr");
    const sorted = [...tasks].sort((a, b) => b.charStart - a.charStart || b.charEnd - a.charEnd);

    let current = originalRun;
    for (const task of sorted) {
      if (!current) {
        throw new Error(`injectEdits: runIndex ${runIndex} has more edits than remaining text can support (overlapping edits?)`);
      }
      const splitResult = splitRun(doc, current, task.charStart, task.charEnd);
      task.run(rPr, splitResult);
      current = splitResult.before;
    }
  }
}

/* ------------------------------------------------------------------ *
 * Paragraph-mark flagging (shared by whole-paragraph insert and delete): ensures pPr/rPr
 * exist, inserting each at the schema-correct position -- w:rPr must appear before
 * w:sectPr/w:pPrChange if either exists on that paragraph (CT_PPr's own child order; this
 * is the concrete mechanism behind the last-paragraph/sectPr known risk, spec §16 --
 * blindly appending would produce schema-invalid XML on exactly the paragraph most likely
 * to carry a sectPr, the last one in the body). Within rPr, ins/del comes before any other
 * formatting properties (CT_ParaRPr's own order), so it's always prepended.
 * ------------------------------------------------------------------ */
function ensurePPr(doc, p) {
  let pPr = kid(p, "pPr");
  if (!pPr) {
    pPr = doc.createElementNS(NS.w, "w:pPr");
    p.insertBefore(pPr, p.firstChild);
  }
  return pPr;
}
function ensureParaRPr(doc, pPr) {
  let rPr = kid(pPr, "rPr");
  if (!rPr) {
    rPr = doc.createElementNS(NS.w, "w:rPr");
    const before = kid(pPr, "sectPr") || kid(pPr, "pPrChange");
    if (before) pPr.insertBefore(rPr, before);
    else pPr.appendChild(rPr);
  }
  return rPr;
}
function flagParagraphMark(doc, p, tag, rev) {
  const pPr = ensurePPr(doc, p);
  const rPr = ensureParaRPr(doc, pPr);
  const markEl = doc.createElementNS(NS.w, "w:" + tag);
  setRevisionAttrs(markEl, rev);
  rPr.insertBefore(markEl, rPr.firstChild);
}

const D4_ERROR_MESSAGE =
  "This paragraph contains existing tracked changes. Accept or reject them in Word first, then re-export and re-run the review.";

/* ------------------------------------------------------------------ *
 * Whole-paragraph delete (D3/D4): wrap every direct-child w:r (and w:hyperlink's inner
 * runs) in a fresh <w:del>, then flag the paragraph mark deleted.
 * ------------------------------------------------------------------ */
function applyWholeParagraphDelete(doc, p, revState) {
  for (const c of [...p.children]) {
    const ln = c.localName;
    if (ln === "pPr") continue;
    if (ln === "r") {
      wrapAsDel(doc, c, nextRevFor(revState));
    } else if (ln === "hyperlink") {
      for (const inner of [...c.children]) {
        if (inner.localName === "r") wrapAsDel(doc, inner, nextRevFor(revState));
      }
    } else if (ln === "ins" || ln === "del" || ln === "fldSimple" || ln === "sdt") {
      // D4: locked content (fldSimple/sdt) can never actually reach here -- its span
      // would overlap a locked range and G4 rejects it upstream -- but the check stays as
      // defense in depth. Pre-existing w:ins/w:del is a real, reachable edge case
      // (deferred by ruling): Word's delete-of-an-insertion semantics are non-trivial, so
      // this throws a clear, user-actionable error instead of guessing at corruption.
      throw new Error(D4_ERROR_MESSAGE);
    }
    // bookmarkStart/End, commentRangeStart/End, proofErr, etc. left alone.
  }
  flagParagraphMark(doc, p, "del", nextRevFor(revState));
}

/* ------------------------------------------------------------------ *
 * Whole-paragraph insert (D2/D5): build a fresh <w:p> and insert it as a sibling of the
 * anchor paragraph.
 * ------------------------------------------------------------------ */
const LIST_PREFIX_RE = /^(?: ?- |\d+\.\s+)/;

// Copies the anchor's pPr for style continuity (heading/list style, indent), minus
// sectPr/pPrChange (those never propagate to a new paragraph) and minus numPr (D5: numPr
// is NOT unconditionally inherited -- it's copied back only when the inserted text itself
// has a recognized bullet/numbered prefix AND the anchor is itself a list item; otherwise
// a plain paragraph merely adjacent to a list must not silently become a list item too).
function buildWholeParagraphInsertPPr(doc, anchorP, text) {
  const anchorPPr = kid(anchorP, "pPr");
  const newPPr = doc.createElementNS(NS.w, "w:pPr");
  let bodyText = text;
  if (anchorPPr) {
    for (const c of anchorPPr.children) {
      const ln = c.localName;
      if (ln === "sectPr" || ln === "pPrChange" || ln === "numPr") continue;
      newPPr.appendChild(c.cloneNode(true));
    }
    const anchorNumPr = kid(anchorPPr, "numPr");
    const m = anchorNumPr ? LIST_PREFIX_RE.exec(text) : null;
    if (m) {
      bodyText = text.slice(m[0].length);
      const numPrClone = anchorNumPr.cloneNode(true);
      const pStyleInNew = kid(newPPr, "pStyle");
      newPPr.insertBefore(numPrClone, pStyleInNew ? pStyleInNew.nextSibling : newPPr.firstChild);
    }
  }
  return { pPr: newPPr, bodyText };
}

function buildWholeParagraphInsertParagraph(doc, anchorP, text, revState) {
  const { pPr: newPPr, bodyText } = buildWholeParagraphInsertPPr(doc, anchorP, text);

  const rPr = doc.createElementNS(NS.w, "w:rPr");
  const markIns = doc.createElementNS(NS.w, "w:ins");
  setRevisionAttrs(markIns, nextRevFor(revState));
  rPr.appendChild(markIns);
  newPPr.appendChild(rPr);

  const newP = doc.createElementNS(NS.w, "w:p");
  newP.appendChild(newPPr);

  const contentIns = doc.createElementNS(NS.w, "w:ins");
  setRevisionAttrs(contentIns, nextRevFor(revState));
  if (bodyText) contentIns.appendChild(buildRunWithText(doc, null, bodyText));
  newP.appendChild(contentIns);

  return newP;
}

// Groups a paragraph-boundary insert group's edits by (bodyPath, edge) and inserts each
// group's new paragraphs as siblings of the anchor, in the edits' own left-to-right
// document order (spec/plan's one explicit exception to right-to-left processing --
// inserting several new siblings at one point is order-sensitive in a way in-run
// splitting isn't). A single fixed reference node per group is enough regardless of edge:
// each subsequent insertBefore(newP, ref) with ref held constant naturally accumulates
// new paragraphs in call order immediately before ref.
function applyWholeParagraphInserts(doc, body, bodyPath, edits, revState) {
  const anchorP = locateParagraph(body, bodyPath);
  const byEdge = new Map();
  for (const edit of edits) {
    const edge = edit.anchor.edge;
    if (!byEdge.has(edge)) byEdge.set(edge, []);
    byEdge.get(edge).push(edit);
  }
  for (const [edge, group] of byEdge) {
    const sorted = [...group].sort((a, b) => a.mdPos - b.mdPos);
    const ref = edge === "before" ? anchorP : anchorP.nextSibling;
    for (const edit of sorted) {
      const newP = buildWholeParagraphInsertParagraph(doc, anchorP, edit.newText, revState);
      anchorP.parentNode.insertBefore(newP, ref);
    }
  }
}

function applyWholeParagraphEdits(doc, body, bodyPath, edits, revState) {
  const deletes = edits.filter((e) => e.anchor.kind === "wholeParagraphDelete");
  const inserts = edits.filter((e) => e.anchor.kind === "paragraphBoundary");
  for (const edit of deletes) {
    applyWholeParagraphDelete(doc, locateParagraph(body, bodyPath), revState);
  }
  if (inserts.length) applyWholeParagraphInserts(doc, body, bodyPath, inserts, revState);
}

/* ------------------------------------------------------------------ *
 * Public entry point (spec §9.1).
 * ------------------------------------------------------------------ */
export function injectEdits(documentXmlDoc, acceptedEdits, _sourceMap, opts = {}) {
  const author = opts.author || "AutoReviewer";
  const date = opts.date || new Date().toISOString();
  const body = findBody(documentXmlDoc);

  const revState = { next: maxIdAmong(documentXmlDoc, REVISION_TAGS) + 1, author, date };
  const commentState = { next: maxIdAmong(documentXmlDoc, COMMENT_MARKER_TAGS) + 1, newComments: [] };

  const byBodyPath = new Map();
  for (const edit of acceptedEdits) {
    const bodyPath = anchorBodyPath(edit.anchor);
    if (bodyPath == null) throw new Error(`injectEdits: edit has no resolvable bodyPath (type ${edit.type})`);
    const key = JSON.stringify(bodyPath);
    if (!byBodyPath.has(key)) byBodyPath.set(key, { bodyPath, edits: [] });
    byBodyPath.get(key).edits.push(edit);
  }

  for (const { bodyPath, edits } of byBodyPath.values()) {
    const p = locateParagraph(body, bodyPath);
    const wholeParagraph = edits.filter((e) => isWholeParagraphAnchor(e.anchor));
    const ordinary = edits.filter((e) => !isWholeParagraphAnchor(e.anchor));
    if (ordinary.length) applyOrdinaryEdits(documentXmlDoc, p, ordinary, revState, commentState);
    if (wholeParagraph.length) applyWholeParagraphEdits(documentXmlDoc, body, bodyPath, wholeParagraph, revState);
  }

  return { newComments: commentState.newComments };
}
