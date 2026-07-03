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
    return { id: revState.next++, author: revState.author, date: revState.date };
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

  for (const [runIndex, tasks] of tasksByRun) {
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

function applyWholeParagraphEdits() {
  throw new Error(
    "injectEdits: whole-paragraph insert/delete is not yet implemented (M3b plan D3/D4/D5, in progress)"
  );
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
    if (wholeParagraph.length) applyWholeParagraphEdits(documentXmlDoc, body, p, bodyPath, wholeParagraph, revState);
  }

  return { newComments: commentState.newComments };
}
