// Reviewer-pass slice rendering (reviewer-pass-slicer step 3, task spec's "Slice export
// format"). Renders ONE (reviewer, pass)'s CriticMarkup body: that reviewer's own edits
// from that pass stay live tokens; every other tracked change (other reviewers, AND this
// same reviewer's other passes) renders as if accepted -- insertions become plain text,
// deletions vanish, substitutions become their replacement text. Comments/replies follow
// the same rule, except a reply whose parent lives in a different pass/author still gets
// its parent embedded read-only for context (never orphan a reply from its parent).
//
// Reuses export.js's segment extraction (buildSegments/normalize) and comment data
// (buildCommentsData) -- the parsing stays in one place. What's genuinely different here
// is the per-segment CriticMarkup emission (filtering by pass membership), so this file
// has its own small body/table walk and its own token formatting; it does not need
// export.js's Composer/source-map machinery at all, since slices are read-only mining
// artifacts, never re-anchored for injection.
import { unzip, readEntry } from "../zip/reader.js";
import { wAttr, kid, kids, parseXml, parseRels, fmtDate } from "./parse.js";
import { buildCommentsData, buildSegments, normalize, emph, cendAfter, segPlainText, containingSentence } from "./export.js";

function paragraphStyle(p) {
  const pPr = kid(p, "pPr");
  let style = "", isList = false, ilvl = 0;
  if (pPr) {
    const ps = kid(pPr, "pStyle"); if (ps) style = wAttr(ps, "val") || "";
    const np = kid(pPr, "numPr");
    if (np) { isList = true; const il = kid(np, "ilvl"); if (il) ilvl = parseInt(wAttr(il, "val") || "0", 10) || 0; }
  }
  return { style, isList, ilvl };
}
function paragraphPrefix({ style, isList, ilvl }) {
  const h = /^Heading([1-6])$/.exec(style);
  if (h) return "#".repeat(+h[1]) + " ";
  if (style === "Title") return "# ";
  if (isList) return "  ".repeat(ilvl) + "- ";
  return "";
}

// A thread node belongs in this slice if it's directly one of this pass's own
// comment/reply observations, OR it's an ancestor of one (walking parentId up to the
// root) -- the parent-context rule, generalized to arbitrarily deep reply chains.
function threadIncludedIds(ctx, rootId) {
  const included = new Set();
  function markAncestors(id) {
    let cur = id;
    while (cur != null && !included.has(cur)) {
      included.add(cur);
      cur = ctx.comments[cur] && ctx.comments[cur].parentId;
    }
  }
  function walk(id) {
    if (ctx.passCommentIds.has(id)) markAncestors(id);
    for (const childId of ctx.childrenMap[id] || []) walk(childId);
  }
  walk(rootId);
  return included;
}

function formatCommentToken(c, parent, context) {
  const dateStr = c.date ? fmtDate(c.date) : "undated";
  const resolvedTag = c.resolved ? " [resolved]" : "";
  const contextTag = context ? ` (context: "${context}")` : "";
  if (parent) {
    return `{>>↳ reply to ${parent.author} — ${c.author} [${dateStr}]${resolvedTag}${contextTag}: ${c.text}<<}`;
  }
  return `{>>${c.author} [${dateStr}]${resolvedTag}${contextTag}: ${c.text}<<}`;
}

// Emits every included node in the thread, root first then depth-first replies (so a
// reply always immediately follows its parent, matching the spec's "nested inline
// immediately after the parent"), regardless of whether the root itself is in this pass.
function renderThreadTokens(ctx, rootId, included, isPoint, paraPlainText, pointPos) {
  const context = isPoint ? containingSentence(paraPlainText, pointPos) : null;
  let s = "";
  function emitNode(id, parentId) {
    if (included.has(id)) {
      const c = { ...ctx.comments[id], resolved: ctx.comments[id].done };
      const parent = parentId != null ? ctx.comments[parentId] : null;
      s += formatCommentToken(c, parent, context);
    }
    for (const childId of ctx.childrenMap[id] || []) emitNode(childId, id);
  }
  emitNode(rootId, null);
  return s;
}

function renderSliceSegs(segs, ctx) {
  let out = "";
  let openId = null;
  let plainPos = 0;
  let spanStartOutLen = 0;
  const paraPlainText = segs.map(segPlainText).join("");

  function closeSpan(id) {
    const included = threadIncludedIds(ctx, id);
    if (!included.size) return; // thread irrelevant to this slice -- leave the plain text as-is
    const spanContent = out.slice(spanStartOutLen);
    const isPoint = spanContent === "";
    out = out.slice(0, spanStartOutLen) + "{==" + (isPoint ? "¶" : spanContent) + "==}";
    out += renderThreadTokens(ctx, id, included, isPoint, paraPlainText, plainPos);
  }

  for (let i = 0; i < segs.length; i++) {
    const s = segs[i];
    if (s.t === "text") { out += emph(s.raw, { b: s.b, i: s.i }); plainPos += s.raw.length; }
    else if (s.t === "locked" || s.t === "opaque") { out += s.s; plainPos += s.s.length; }
    else if (s.t === "ins") {
      out += ctx.passIds.has(s.id) ? "{++" + s.s + "++}" : s.s; // accepted view: insertion kept plain
    }
    else if (s.t === "del") {
      if (ctx.passIds.has(s.id)) out += "{--" + s.s + "--}"; // accepted view: deletion vanishes
    }
    else if (s.t === "sub") {
      // delId/insId always agree on pass membership (normalize() only pairs a sub when
      // both sides share one author+date, i.e. one pass) -- checking either is enough.
      out += ctx.passIds.has(s.delId) || ctx.passIds.has(s.insId) ? "{~~" + s.del + "~>" + s.ins + "~~}" : s.ins;
    }
    else if (s.t === "cstart") {
      if (openId === null && cendAfter(segs, i, s.id)) {
        openId = s.id; spanStartOutLen = out.length;
      } else {
        // No matching cend in this paragraph (point comment or a range spanning multiple
        // paragraphs -- see export.js's cstart comment); either way there's no span here
        // to wrap, so render like a point: {==¶==} + containing-sentence context.
        closeSpan(s.id);
      }
    } else if (s.t === "cend") {
      if (openId === s.id) {
        closeSpan(s.id);
        openId = null;
      }
    }
  }
  if (openId !== null) closeSpan(openId);
  return out;
}

function renderSliceParagraph(p, rels, ctx) {
  const { style, isList, ilvl } = paragraphStyle(p);
  const runCounter = { next: 0 };
  const segs = normalize(buildSegments(p, rels, runCounter));
  const text = renderSliceSegs(segs, ctx).trim();
  const h = /^Heading([1-6])$/.exec(style);
  const forceKeep = !!h || style === "Title" || isList;
  if (!text && !forceKeep) return null;
  return paragraphPrefix({ style, isList, ilvl }) + text;
}

function renderSliceTable(tbl, rels, ctx) {
  const rows = kids(tbl, "tr");
  if (!rows.length) return null;
  const matrix = rows.map((tr) =>
    kids(tr, "tc").map((tc) => {
      const cellText = kids(tc, "p").map((p) => renderSliceParagraph(p, rels, ctx)).filter(Boolean).join(" ");
      return (cellText || " ").replace(/\|/g, "\\|");
    })
  );
  const cols = Math.max(...matrix.map((r) => r.length));
  const writeRow = (row) => "| " + Array.from({ length: cols }, (_, c) => row[c] ?? " ").join(" | ") + " |";
  const lines = [writeRow(matrix[0]), "| " + Array(cols).fill("---").join(" | ") + " |"];
  for (let r = 1; r < matrix.length; r++) lines.push(writeRow(matrix[r]));
  return lines.join("\n");
}

function renderSliceBody(body, rels, ctx) {
  const blocks = [];
  for (const el of body.children) {
    if (el.localName === "p") {
      const rendered = renderSliceParagraph(el, rels, ctx);
      if (rendered != null) blocks.push(rendered);
    } else if (el.localName === "tbl") {
      const rendered = renderSliceTable(el, rels, ctx);
      if (rendered != null) blocks.push(rendered);
    } else if (el.localName === "sdt") {
      const content = kid(el, "sdtContent");
      if (content) {
        for (const ip of kids(content, "p")) {
          const rendered = renderSliceParagraph(ip, rels, ctx);
          if (rendered != null) blocks.push(rendered);
        }
      }
    }
    // sectPr and other body children skipped, matching export.js.
  }
  return blocks.join("\n\n");
}

function slugify(name) {
  return String(name).trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "unknown";
}

// "2026-05-14T09:12:00Z" -> "2026-05-14T09:12Z" (spec's pass_window example omits seconds).
function fmtWindow(iso) {
  if (!iso) return null;
  return iso.replace(/^(.*T\d{2}:\d{2}):\d{2}(Z?)$/, "$1$2");
}

export function sliceFilename(docBaseName, pass) {
  const dateSlug = pass.undated ? "undated" : pass.passDate;
  return `${slugify(docBaseName)}_${slugify(pass.author)}_${dateSlug}.md`;
}

function buildFrontmatter({ docFilename, pass, generated }) {
  const passField = pass.undated ? "undated" : pass.passDate;
  const windowField = pass.undated ? "n/a" : `${fmtWindow(pass.windowStart)} – ${fmtWindow(pass.windowEnd)}`;
  return [
    "---",
    `doc: ${docFilename}.docx`,
    `reviewer: ${pass.author}`,
    `pass: ${passField}`,
    `pass_window: ${windowField}`,
    `edits: ${pass.counts.insertions + pass.counts.deletions}`,
    `comments: ${pass.counts.comments}`,
    `replies: ${pass.counts.replies}`,
    `generated: ${generated}`,
    "---",
  ].join("\n");
}

// Renders one reviewer/pass slice (task spec's "Slice export format") for a given pass
// object from clusterPasses (src/passes.js). options.generated defaults to today (UTC);
// pass a fixed value in tests for determinism.
export async function renderSlice(docxBytes, pass, options = {}) {
  const { DOMParserImpl, filename = "document", generated = new Date().toISOString().slice(0, 10) } = options;

  const zip = await unzip(docxBytes);
  const docXml = await readEntry(zip, "word/document.xml");
  if (!docXml) throw new Error("not a Word document (no word/document.xml)");

  const rels = parseRels(await readEntry(zip, "word/_rels/document.xml.rels"), DOMParserImpl);
  const { comments, childrenMap } = buildCommentsData(
    await readEntry(zip, "word/comments.xml"),
    await readEntry(zip, "word/commentsExtended.xml"),
    DOMParserImpl
  );

  const docDoc = parseXml(docXml, DOMParserImpl);
  const body = [...docDoc.documentElement.children].find((c) => c.localName === "body");
  if (!body) throw new Error("no document body found");

  const passIds = new Set(pass.observations.filter((o) => o.id != null).map((o) => o.id));
  const passCommentIds = new Set(pass.observations.filter((o) => o.commentId != null).map((o) => o.commentId));
  const ctx = { comments, childrenMap, passIds, passCommentIds };

  const bodyMd = renderSliceBody(body, rels, ctx);
  const frontmatter = buildFrontmatter({ docFilename: filename, pass, generated });
  return frontmatter + "\n\n" + bodyMd + "\n";
}
