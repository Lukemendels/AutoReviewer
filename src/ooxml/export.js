// docx -> {markdown, sourceMap, comments, counts}.
//
// The rendering algorithm (segment building, comment threading, table/list/heading
// handling, CriticMarkup emission for pre-existing tracked changes) is ported unchanged
// in behavior from ref/redline-to-markdown.html. What's new here is the Composer: every
// render function writes into it instead of returning a plain string, so we can record,
// alongside the markdown text itself, which byte ranges are real document text (mapped
// back to a run + char offset) vs. synthetic markup the exporter invented (spec §5.2),
// and which ranges are locked protected islands (spec §5.3). Concatenation is
// associative, so this restructuring does not change the emitted markdown itself for any
// construct the reference impl actually handles -- verified directly against the
// reference impl's own algorithm (all fixtures but one are byte-identical; the exception
// is fields-and-content-controls.docx, where this port adds the locked ⟦field: ...⟧
// placeholder spec §5.3 requires and the reference impl doesn't implement at all).
import { unzip, readEntry } from "../zip/reader.js";
import { NS, wAttr, w14Attr, w15Attr, rAttr, kids, kid, basename, fmtDate, parseXml, parseRels } from "./parse.js";

/* ------------------------------------------------------------------ *
 * Composer: builds markdown text while tracking synthetic/locked/doc-text spans.
 * ------------------------------------------------------------------ */
function pushRange(arr, start, end) {
  if (start === end) return;
  const last = arr[arr.length - 1];
  if (last && last[1] === start) last[1] = end;
  else arr.push([start, end]);
}
function clipRanges(arr, pos) {
  const out = [];
  for (const [s, e] of arr) {
    if (s >= pos) continue;
    out.push([s, Math.min(e, pos)]);
  }
  return out;
}

class Composer {
  constructor() {
    this.out = "";
    this.docRuns = []; // populated when this composer is used as a per-paragraph scratch buffer
    this.synthetic = [];
    this.locked = [];
    this.blocks = []; // populated when this composer is used as the top-level accumulator
  }
  get pos() {
    return this.out.length;
  }
  writeSynthetic(str) {
    if (!str) return;
    const s = this.pos;
    this.out += str;
    pushRange(this.synthetic, s, this.pos);
  }
  writeLocked(str) {
    if (!str) return;
    const s = this.pos;
    this.out += str;
    pushRange(this.synthetic, s, this.pos);
    pushRange(this.locked, s, this.pos);
  }
  writeDocText(str, runIndex, charOffset) {
    if (!str) return;
    const s = this.pos;
    this.out += str;
    this.docRuns.push({ mdStart: s, mdEnd: this.pos, runIndex, charOffset });
  }
  rollbackTo(pos) {
    this.out = this.out.slice(0, pos);
    this.synthetic = clipRanges(this.synthetic, pos);
    this.locked = clipRanges(this.locked, pos);
    this.blocks = this.blocks.filter((b) => b.mdStart < pos);
  }
  mergeFrom(other) {
    const shift = this.pos;
    this.out += other.out;
    for (const r of other.synthetic) pushRange(this.synthetic, r[0] + shift, r[1] + shift);
    for (const r of other.locked) pushRange(this.locked, r[0] + shift, r[1] + shift);
    for (const b of other.blocks) {
      this.blocks.push({
        ...b,
        mdStart: b.mdStart + shift,
        mdEnd: b.mdEnd + shift,
        runs: b.runs.map((r) => ({ ...r, mdStart: r.mdStart + shift, mdEnd: r.mdEnd + shift })),
      });
    }
  }
}

/* ------------------------------------------------------------------ *
 * Comments + threads (ported; state passed explicitly instead of module globals)
 * ------------------------------------------------------------------ */
export function buildCommentsData(commentsXml, extXml, DOMParserImpl) {
  const comments = {};
  const childrenMap = {};
  const paraIdToId = {};
  if (commentsXml) {
    const doc = parseXml(commentsXml, DOMParserImpl);
    for (const cm of doc.getElementsByTagName("*")) {
      if (cm.localName !== "comment") continue;
      const id = wAttr(cm, "id");
      const author = wAttr(cm, "author") || "Unknown";
      // Full ISO timestamp kept here (spec: reviewer-pass-slicer needs hour/minute precision
      // for 48h-gap pass clustering); truncate to date-only only at render time (commentToken).
      const date = wAttr(cm, "date") || null;
      const ps = [];
      for (const p of cm.getElementsByTagName("*")) {
        if (p.localName === "p") {
          ps.push(p);
          const pid = w14Attr(p, "paraId");
          if (pid) paraIdToId[pid] = id;
        }
      }
      const text = ps.map((p) => collectRunsText(p).trim()).filter(Boolean).join(" ") || "(empty comment)";
      // done stays null (unknown) rather than false when commentsExtended.xml itself is
      // absent from the package -- "not resolved" and "no resolution data at all" are
      // different facts the slicer's raw-observation model needs to keep apart.
      comments[id] = { id, author, date, text, done: extXml ? false : null, parentId: null };
    }
  }
  if (extXml) {
    const doc = parseXml(extXml, DOMParserImpl);
    for (const ce of doc.getElementsByTagName("*")) {
      if (ce.localName !== "commentEx") continue;
      const pid = w15Attr(ce, "paraId");
      const parent = w15Attr(ce, "paraIdParent");
      const done = w15Attr(ce, "done");
      const id = paraIdToId[pid];
      if (id == null || !comments[id]) continue;
      comments[id].done = done === "1" || done === "true";
      if (parent) {
        const par = paraIdToId[parent];
        if (par != null) comments[id].parentId = par;
      }
    }
  }
  for (const id in comments) {
    const pid = comments[id].parentId;
    if (pid != null) (childrenMap[pid] = childrenMap[pid] || []).push(id);
  }
  return { comments, childrenMap };
}
function commentToken(c, depth) {
  const arrow = depth > 0 ? "↳".repeat(depth) + " " : "";
  const date = c.date ? " (" + fmtDate(c.date) + ")" : "";
  const res = c.done ? " [resolved]" : "";
  return "{>>" + arrow + c.author + date + res + ": " + c.text + "<<}";
}
function renderReplies(comments, childrenMap, parentId, depth) {
  let s = "";
  for (const childId of childrenMap[parentId] || []) {
    s += commentToken(comments[childId], depth);
    s += renderReplies(comments, childrenMap, childId, depth + 1);
  }
  return s;
}
function renderThread(comments, childrenMap, id) {
  const c = comments[id];
  if (!c) return "";
  return commentToken(c, 0) + renderReplies(comments, childrenMap, id, 1);
}

/* ------------------------------------------------------------------ *
 * Raw observation collection (reviewer-pass-slicer step 1, spec-workbench.md).
 * Woven into the same walk that builds the markdown -- see serializeSegsTracked below --
 * so the slicer never diverges from the one canonical exporter. Only active when
 * ctx.observations is non-null (exportDocx's collectObservations option).
 * ------------------------------------------------------------------ */

// Point comments (commentRangeStart immediately followed by commentRangeEnd, nothing
// between) have no real span to quote -- approximate with the sentence containing the
// point, found by nearest sentence-ending punctuation on each side. Good enough for the
// audit table; not meant to handle abbreviations/decimals precisely.
function containingSentence(text, offset) {
  if (!text) return "";
  const before = text.slice(0, offset);
  const after = text.slice(offset);
  const lastBreak = [...before.matchAll(/[.!?]\s+/g)].pop();
  const start = lastBreak ? lastBreak.index + lastBreak[0].length : 0;
  const nextBreak = /[.!?](\s|$)/.exec(after);
  const end = nextBreak ? offset + nextBreak.index + 1 : text.length;
  return text.slice(start, end).trim();
}

// Visible plain-text contribution of a segment, for anchor-text accumulation only --
// deliberately excludes ins/del (a reviewer's own tracked edits aren't "the resolved text
// being commented on"; true accepted-view resolution is deferred to the later
// slice-rendering step, out of scope here).
function segPlainText(s) {
  if (s.t === "text") return s.raw;
  if (s.t === "locked" || s.t === "opaque") return s.s;
  return "";
}

// One raw observation per node in a comment thread (root + every reply, depth-first),
// all sharing the root's anchorText since replies comment on the thread, not on a new
// document span of their own.
function pushReplyObservations(ctx, parentId, anchorText) {
  for (const childId of ctx.childrenMap[parentId] || []) {
    const c = ctx.comments[childId];
    ctx.observations.push({
      kind: "comment-reply",
      author: c.author,
      date: c.date,
      anchorText,
      text: c.text,
      parentCommentId: parentId,
      resolved: c.done,
      docOrder: ctx.docOrder.next++,
    });
    pushReplyObservations(ctx, childId, anchorText);
  }
}
function pushThreadObservations(ctx, rootId, anchorText) {
  const root = ctx.comments[rootId];
  if (!root) return;
  ctx.observations.push({
    kind: "comment",
    author: root.author,
    date: root.date,
    anchorText,
    text: root.text,
    parentCommentId: null,
    resolved: root.done,
    docOrder: ctx.docOrder.next++,
  });
  pushReplyObservations(ctx, rootId, anchorText);
}

/* ------------------------------------------------------------------ *
 * Run / text extraction (ported)
 * ------------------------------------------------------------------ */
function runText(r) {
  let s = "";
  for (const c of r.children) {
    if (c.localName === "t" || c.localName === "delText") s += c.textContent;
    else if (c.localName === "tab") s += " ";
    else if (c.localName === "br" || c.localName === "cr") s += "  \n";
  }
  return s;
}
function runEmph(r) {
  const rPr = kid(r, "rPr");
  let b = false, i = false;
  if (rPr) {
    const be = kid(rPr, "b"); if (be) { const v = wAttr(be, "val"); b = !(v === "0" || v === "false"); }
    const ie = kid(rPr, "i"); if (ie) { const v = wAttr(ie, "val"); i = !(v === "0" || v === "false"); }
  }
  return { b, i };
}
// Splits an emphasis-wrapped string into its real-text (lead/core/trail) and
// synthetic-marker (openMark/closeMark) parts, matching the original single-string
// emph() byte-for-byte when reassembled: lead+openMark+core+closeMark+trail.
function emphParts(s, { b, i }) {
  if (!s || !s.trim()) return { lead: s || "", core: "", trail: "", openMark: "", closeMark: "" };
  const lead = s.match(/^\s*/)[0], trail = s.match(/\s*$/)[0];
  const core = s.slice(lead.length, s.length - trail.length);
  const openMark = (b ? "**" : "") + (i ? "*" : "");
  const closeMark = (i ? "*" : "") + (b ? "**" : "");
  return { lead, core, trail, openMark, closeMark };
}
function emph(s, be) {
  const { lead, core, trail, openMark, closeMark } = emphParts(s, be);
  return lead + openMark + core + closeMark + trail;
}
// Non-indexed run-text collector, used only for comments.xml bodies (not part of the
// exported document, so no run-index bookkeeping is needed there).
function collectRunsText(el) {
  let s = "";
  for (const c of el.children) {
    if (c.localName === "r") s += emph(runText(c), runEmph(c));
    else if (["ins", "del", "hyperlink", "smartTag"].includes(c.localName)) s += collectRunsText(c);
  }
  return s;
}
// Same, but for the document body: advances the shared per-paragraph run counter for
// every w:r it consumes so run indices stay re-derivable later by a plain doc-order walk.
function collectRunsTextIndexed(el, runCounter) {
  let s = "";
  for (const c of el.children) {
    if (c.localName === "r") { s += emph(runText(c), runEmph(c)); runCounter.next++; }
    else if (["ins", "del", "hyperlink", "smartTag"].includes(c.localName)) s += collectRunsTextIndexed(c, runCounter);
  }
  return s;
}
function findImage(r, rels) {
  for (const e of r.getElementsByTagName("*")) {
    if (e.localName === "blip") {
      const id = rAttr(e, "embed") || rAttr(e, "link");
      if (id && rels[id]) return basename(rels[id]);
      if (id) return "image";
    }
    if (e.localName === "imagedata") {
      const id = rAttr(e, "id");
      if (id && rels[id]) return basename(rels[id]);
    }
  }
  return null;
}

/* ------------------------------------------------------------------ *
 * Paragraph -> ordered segment list (ported + extended: run indices for
 * document-text mapping; locked placeholders for fields/content controls/
 * images/math/objects per spec §5.3; hyperlinks marked "opaque" -- synthetic
 * but not a spec-defined protected island).
 * ------------------------------------------------------------------ */
function buildSegments(p, rels, runCounter) {
  const segs = [];
  let fieldState = null; // {buf, instr} while inside a begin/separate/end fldChar sequence
  for (const c of p.children) {
    const ln = c.localName;
    if (ln === "commentRangeStart") segs.push({ t: "cstart", id: wAttr(c, "id") });
    else if (ln === "commentRangeEnd") segs.push({ t: "cend", id: wAttr(c, "id") });
    else if (ln === "fldSimple") {
      const instr = wAttr(c, "instr") || "";
      const cached = collectRunsTextIndexed(c, runCounter).trim();
      segs.push({ t: "locked", s: "⟦field: " + (instr.trim() || cached || "field") + "⟧" });
    } else if (ln === "sdt") {
      const content = kid(c, "sdtContent");
      let inner = "";
      if (content) {
        for (const cc of content.children) {
          if (cc.localName === "r") { inner += emph(runText(cc), runEmph(cc)); runCounter.next++; }
          else if (["ins", "del", "hyperlink", "smartTag"].includes(cc.localName)) inner += collectRunsTextIndexed(cc, runCounter);
        }
      }
      if (inner.trim()) segs.push({ t: "locked", s: inner });
    } else if (ln === "r") {
      const runIndex = runCounter.next++;
      const fldChar = kid(c, "fldChar");
      if (fldChar) {
        const type = wAttr(fldChar, "fldCharType");
        if (type === "begin") fieldState = { buf: "", instr: "" };
        else if (type === "end") {
          if (fieldState) {
            const shown = fieldState.instr.trim() || fieldState.buf.trim() || "field";
            segs.push({ t: "locked", s: "⟦field: " + shown + "⟧" });
          }
          fieldState = null;
        }
        continue;
      }
      const instrText = kid(c, "instrText");
      if (instrText && fieldState) { fieldState.instr += instrText.textContent; continue; }
      const img = findImage(c, rels);
      if (img) { segs.push({ t: "locked", s: "[image: " + img + "]" }); continue; }
      const raw = runText(c);
      if (fieldState) { if (raw) fieldState.buf += raw; continue; }
      if (!raw) {
        const hasEmbed = [...c.getElementsByTagName("*")].some((e) => ["oMath", "object", "pict"].includes(e.localName));
        if (hasEmbed) { segs.push({ t: "locked", s: "⟦object⟧" }); continue; }
      }
      const { b, i } = runEmph(c);
      if (raw) segs.push({ t: "text", raw, b, i, runIndex });
    } else if (ln === "ins") {
      segs.push({ t: "ins", s: collectRunsTextIndexed(c, runCounter), author: wAttr(c, "author"), date: wAttr(c, "date") });
    } else if (ln === "del") {
      segs.push({ t: "del", s: collectRunsTextIndexed(c, runCounter), author: wAttr(c, "author"), date: wAttr(c, "date") });
    } else if (ln === "hyperlink") {
      const txt = collectRunsTextIndexed(c, runCounter);
      const url = rels[rAttr(c, "id")] || "";
      segs.push({ t: "opaque", s: url ? "[" + txt + "](" + url + ")" : txt });
    }
    // bookmarkStart/End, proofErr and other no-text nodes are ignored (matches reference impl).
  }
  return segs;
}
// Merge fragmented ins/del runs; collapse adjacent del->ins into a substitution. Unlike
// the reference impl, "text" segments are intentionally NOT merged here -- each stays
// tagged with its own runIndex for the source map. String concatenation is associative,
// so leaving them unmerged does not change the emitted markdown.
function normalize(segs) {
  const merged = [];
  for (const s of segs) {
    const last = merged[merged.length - 1];
    // Only merge adjacent same-type runs when they share one author/date -- otherwise
    // they're two distinct tracked-change actions that happen to be text-adjacent, and
    // merging would silently discard whichever side's attribution didn't win (spec:
    // reviewer-pass-slicer needs every edit's own author/date intact).
    if (last && last.t === s.t && (s.t === "ins" || s.t === "del") && last.author === s.author && last.date === s.date) {
      last.s += s.s;
    } else merged.push({ ...s });
  }
  const out = [];
  for (let i = 0; i < merged.length; i++) {
    const curr = merged[i], next = merged[i + 1];
    if (curr.t === "del" && next && next.t === "ins" && curr.author === next.author && curr.date === next.date) {
      out.push({ t: "sub", del: curr.s, ins: next.s, author: next.author, date: next.date });
      i++;
    } else out.push(curr);
  }
  return out;
}
function authorTag(s, annotate) {
  return annotate && s.author ? "{>>—" + s.author + "<<}" : "";
}
function cendAfter(segs, i, id) {
  for (let j = i + 1; j < segs.length; j++) if (segs[j].t === "cend" && segs[j].id === id) return true;
  return false;
}

function writeTextSegTracked(local, seg) {
  const { lead, core, trail, openMark, closeMark } = emphParts(seg.raw, { b: seg.b, i: seg.i });
  let charOffset = 0;
  if (lead) { local.writeDocText(lead, seg.runIndex, charOffset); charOffset += lead.length; }
  if (openMark) local.writeSynthetic(openMark);
  if (core) { local.writeDocText(core, seg.runIndex, charOffset); charOffset += core.length; }
  if (closeMark) local.writeSynthetic(closeMark);
  if (trail) local.writeDocText(trail, seg.runIndex, charOffset);
}
function serializeSegsTracked(local, segs, ctx) {
  let openId = null;
  let anchorStart = 0;
  let plainPos = 0;
  const paraPlainText = ctx.observations ? segs.map(segPlainText).join("") : "";

  function commentObserved(id, anchorText) {
    if (!ctx.observations) return;
    const text = anchorText && anchorText.trim() ? anchorText.trim() : containingSentence(paraPlainText, plainPos);
    pushThreadObservations(ctx, id, text);
  }

  for (let i = 0; i < segs.length; i++) {
    const s = segs[i];
    if (s.t === "text") { writeTextSegTracked(local, s); plainPos += s.raw.length; }
    else if (s.t === "locked") { local.writeLocked(s.s); plainPos += s.s.length; }
    else if (s.t === "opaque") { local.writeSynthetic(s.s); plainPos += s.s.length; }
    else if (s.t === "ins") {
      local.writeSynthetic("{++" + s.s + "++}" + authorTag(s, ctx.annotate));
      ctx.counts.ins++;
      if (ctx.observations) {
        ctx.observations.push({
          kind: "insertion", author: s.author || "Unknown", date: s.date || null,
          anchorText: null, text: s.s, parentCommentId: null, resolved: null, docOrder: ctx.docOrder.next++,
        });
      }
    }
    else if (s.t === "del") {
      local.writeSynthetic("{--" + s.s + "--}" + authorTag(s, ctx.annotate));
      ctx.counts.del++;
      if (ctx.observations) {
        ctx.observations.push({
          kind: "deletion", author: s.author || "Unknown", date: s.date || null,
          anchorText: null, text: s.s, parentCommentId: null, resolved: null, docOrder: ctx.docOrder.next++,
        });
      }
    }
    else if (s.t === "sub") {
      local.writeSynthetic("{~~" + s.del + "~>" + s.ins + "~~}" + authorTag(s, ctx.annotate));
      ctx.counts.sub++;
      // Raw observation model (spec-workbench.md) has no "substitution" kind -- a sub is
      // just an adjacent del+ins pair the renderer merges into one CriticMarkup token, so
      // unpack it back into its two constituent observations here.
      if (ctx.observations) {
        ctx.observations.push({
          kind: "deletion", author: s.author || "Unknown", date: s.date || null,
          anchorText: null, text: s.del, parentCommentId: null, resolved: null, docOrder: ctx.docOrder.next++,
        });
        ctx.observations.push({
          kind: "insertion", author: s.author || "Unknown", date: s.date || null,
          anchorText: null, text: s.ins, parentCommentId: null, resolved: null, docOrder: ctx.docOrder.next++,
        });
      }
    }
    else if (s.t === "cstart") {
      // cendAfter only looks within THIS paragraph's segs (buildSegments runs per
      // paragraph) -- a comment range that legitimately spans multiple paragraphs lands
      // here too, indistinguishable from a genuinely rangeless/malformed one. Either way
      // commentObserved falls back to containingSentence() over this one paragraph's plain
      // text, not the comment's true (possibly multi-paragraph) span. Known approximation:
      // the slice renderer (later step) must not assume anchorText is always the complete
      // anchor for a cross-paragraph comment.
      if (openId === null && cendAfter(segs, i, s.id)) { openId = s.id; anchorStart = plainPos; local.writeSynthetic("{=="); }
      else if (!ctx.emitted.has(s.id)) {
        local.writeSynthetic(renderThread(ctx.comments, ctx.childrenMap, s.id));
        ctx.emitted.add(s.id);
        commentObserved(s.id, "");
      }
    } else if (s.t === "cend") {
      if (openId === s.id) {
        local.writeSynthetic("==}");
        openId = null;
        if (!ctx.emitted.has(s.id)) {
          local.writeSynthetic(renderThread(ctx.comments, ctx.childrenMap, s.id));
          ctx.emitted.add(s.id);
          commentObserved(s.id, paraPlainText.slice(anchorStart, plainPos));
        }
      }
    }
  }
  if (openId !== null) {
    local.writeSynthetic("==}");
    if (!ctx.emitted.has(openId)) {
      local.writeSynthetic(renderThread(ctx.comments, ctx.childrenMap, openId));
      ctx.emitted.add(openId);
      commentObserved(openId, paraPlainText.slice(anchorStart, plainPos));
    }
  }
}

function renderParagraphContent(p, rels, ctx) {
  const pPr = kid(p, "pPr");
  let style = "", isList = false, ilvl = 0;
  if (pPr) {
    const ps = kid(pPr, "pStyle"); if (ps) style = wAttr(ps, "val") || "";
    const np = kid(pPr, "numPr");
    if (np) { isList = true; const il = kid(np, "ilvl"); if (il) ilvl = parseInt(wAttr(il, "val") || "0", 10) || 0; }
  }
  const runCounter = { next: 0 };
  const segs = normalize(buildSegments(p, rels, runCounter));
  const local = new Composer();
  serializeSegsTracked(local, segs, ctx);
  return { local, style, isList, ilvl };
}

// Commits a rendered paragraph's trimmed content into `main`, registering a source-map
// block. Mirrors the reference impl's trim() semantics exactly: plain paragraphs that
// are entirely blank contribute nothing (return false); headings/lists always contribute
// at least their prefix, even when the paragraph's own text is empty (forceKeep).
function commitParagraphBlock(main, local, prefix, kind, bodyPath, forceKeep = false) {
  const text = local.out;
  const leadWs = text.match(/^\s*/)[0].length;
  const trailWs = text.match(/\s*$/)[0].length;
  const trimEnd = Math.max(leadWs, text.length - trailWs);
  const isBlank = leadWs >= trimEnd;
  if (isBlank && !forceKeep) return false;

  const trimmedText = isBlank ? "" : text.slice(leadWs, trimEnd);
  const blockStart = main.pos;
  if (prefix) main.writeSynthetic(prefix);
  const afterPrefix = main.pos;
  if (trimmedText) main.out += trimmedText;

  if (!isBlank) {
    for (const r of local.synthetic) {
      const s = Math.max(r[0], leadWs), e = Math.min(r[1], trimEnd);
      if (e > s) pushRange(main.synthetic, afterPrefix + s - leadWs, afterPrefix + e - leadWs);
    }
    for (const r of local.locked) {
      const s = Math.max(r[0], leadWs), e = Math.min(r[1], trimEnd);
      if (e > s) pushRange(main.locked, afterPrefix + s - leadWs, afterPrefix + e - leadWs);
    }
  }
  const runs = [];
  if (!isBlank) {
    for (const dr of local.docRuns) {
      const s = Math.max(dr.mdStart, leadWs), e = Math.min(dr.mdEnd, trimEnd);
      if (e <= s) continue;
      const dropStart = s - dr.mdStart;
      runs.push({ mdStart: afterPrefix + s - leadWs, mdEnd: afterPrefix + e - leadWs, runIndex: dr.runIndex, charOffset: dr.charOffset + dropStart });
    }
  }
  main.blocks.push({ mdStart: blockStart, mdEnd: main.pos, kind, bodyPath, runs });
  return true;
}
// Body-level content control (w:sdt) support, v1: wraps whole paragraphs as a single
// locked block each (no doc-text runs). Nested tables inside a body-level sdt are not
// supported in v1 (rare; documented gap).
function commitLockedParagraph(main, local, bodyPath) {
  const text = local.out.trim();
  if (!text) return false;
  const blockStart = main.pos;
  main.writeLocked(text);
  main.blocks.push({ mdStart: blockStart, mdEnd: main.pos, kind: "locked", bodyPath, runs: [] });
  return true;
}

/* ------------------------------------------------------------------ *
 * Table rendering (ported layout; per-cell source-map blocks)
 * ------------------------------------------------------------------ */
function renderTableCellTracked(tc, rels, ctx, bodyPathPrefix) {
  const paras = kids(tc, "p");
  const pieces = [];
  for (let pIdx = 0; pIdx < paras.length; pIdx++) {
    const { local } = renderParagraphContent(paras[pIdx], rels, ctx);
    const t = local.out.trim();
    if (t) pieces.push({ pIdx, local, text: t });
  }
  const joinedPlain = pieces.map((p) => p.text).join(" ");
  const hasPipe = joinedPlain.includes("|");
  return function write(main) {
    if (!pieces.length) { main.writeSynthetic(" "); return; }
    if (hasPipe) {
      // Rare: a literal "|" inside a cell forces a coarser, whole-cell-synthetic
      // fallback rather than tracking per-run offsets through pipe-escaping.
      main.writeSynthetic(joinedPlain.replace(/\|/g, "\\|"));
      return;
    }
    pieces.forEach((piece, idx) => {
      if (idx > 0) main.writeSynthetic(" ");
      commitParagraphBlock(main, piece.local, "", "tableCell", [...bodyPathPrefix, piece.pIdx]);
    });
  };
}
function renderTableTracked(main, tbl, rels, ctx, bodyIdx) {
  const rows = kids(tbl, "tr");
  if (!rows.length) return false;
  const matrixWriters = rows.map((tr, rowIdx) =>
    kids(tr, "tc").map((tc, cellIdx) => renderTableCellTracked(tc, rels, ctx, [bodyIdx, rowIdx, cellIdx]))
  );
  const cols = Math.max(...matrixWriters.map((r) => r.length));

  function writeRow(cellWriters) {
    main.writeSynthetic("| ");
    for (let c = 0; c < cols; c++) {
      if (c > 0) main.writeSynthetic(" | ");
      if (cellWriters[c]) cellWriters[c](main); else main.writeSynthetic(" ");
    }
    main.writeSynthetic(" |");
  }

  writeRow(matrixWriters[0]);
  main.writeSynthetic("\n" + "| " + Array(cols).fill("---").join(" | ") + " |");
  for (let r = 1; r < matrixWriters.length; r++) {
    main.writeSynthetic("\n");
    writeRow(matrixWriters[r]);
  }
  return true;
}

/* ------------------------------------------------------------------ *
 * Body walk + header + full render
 * ------------------------------------------------------------------ */
function walkBodyTracked(main, body, rels, ctx) {
  let bodyIdx = 0;
  let wroteAny = false;
  for (const el of body.children) {
    if (el.localName === "p") {
      const { local, style, isList, ilvl } = renderParagraphContent(el, rels, ctx);
      const h = /^Heading([1-6])$/.exec(style);
      let prefix = "", kind = "p", forceKeep = false;
      if (h) { prefix = "#".repeat(+h[1]) + " "; kind = "heading"; forceKeep = true; }
      else if (style === "Title") { prefix = "# "; kind = "heading"; forceKeep = true; }
      else if (isList) { prefix = "  ".repeat(ilvl) + "- "; kind = "list"; forceKeep = true; }

      const beforePos = main.pos;
      if (wroteAny) main.writeSynthetic("\n\n");
      const wrote = commitParagraphBlock(main, local, prefix, kind, [bodyIdx], forceKeep);
      if (!wrote) main.rollbackTo(beforePos); else wroteAny = true;
    } else if (el.localName === "tbl") {
      const beforePos = main.pos;
      if (wroteAny) main.writeSynthetic("\n\n");
      const wrote = renderTableTracked(main, el, rels, ctx, bodyIdx);
      if (!wrote) main.rollbackTo(beforePos); else wroteAny = true;
    } else if (el.localName === "sdt") {
      const content = kid(el, "sdtContent");
      if (content) {
        for (const ip of kids(content, "p")) {
          const { local } = renderParagraphContent(ip, rels, ctx);
          const beforePos = main.pos;
          if (wroteAny) main.writeSynthetic("\n\n");
          const wrote = commitLockedParagraph(main, local, [bodyIdx]);
          if (!wrote) main.rollbackTo(beforePos); else wroteAny = true;
        }
      }
    }
    // sectPr and other body children are skipped (matches reference impl).
    bodyIdx++;
  }
}
function buildHeaderTracked(filename, comments, counts) {
  const total = Object.keys(comments).length;
  const unresolved = Object.values(comments).filter((c) => !c.done).length;
  const parts = [];
  if (counts.ins) parts.push(counts.ins + " insertion" + (counts.ins !== 1 ? "s" : ""));
  if (counts.del) parts.push(counts.del + " deletion" + (counts.del !== 1 ? "s" : ""));
  if (counts.sub) parts.push(counts.sub + " substitution" + (counts.sub !== 1 ? "s" : ""));
  if (total) parts.push(total + " comment" + (total !== 1 ? "s" : "") + (unresolved ? " (" + unresolved + " unresolved)" : ""));
  if (counts.fmt) parts.push(counts.fmt + " formatting change" + (counts.fmt !== 1 ? "s" : ""));
  const summary = parts.length ? parts.join(" · ") : "no tracked changes or comments detected";
  return (
    "<!-- Redline export from: " + filename + ".docx -->\n" +
    "<!-- " + summary + " -->\n" +
    "<!-- CriticMarkup legend: {++ins++} {--del--} {~~old~>new~~} {==highlighted==} {>>comment<<} -->"
  );
}
function render({ body, rels, comments, childrenMap, docDoc, filename }, annotate, collectObservations) {
  const ctx = {
    comments, childrenMap, annotate,
    counts: { ins: 0, del: 0, sub: 0, fmt: 0 },
    emitted: new Set(),
    observations: collectObservations ? [] : null,
    docOrder: { next: 0 },
  };
  for (const e of docDoc.getElementsByTagName("*")) {
    if (e.localName === "rPrChange" || e.localName === "pPrChange") ctx.counts.fmt++;
  }

  const bodyComposer = new Composer();
  walkBodyTracked(bodyComposer, body, rels, ctx);

  const final = new Composer();
  final.writeSynthetic(buildHeaderTracked(filename, ctx.comments, ctx.counts));
  final.writeSynthetic("\n\n");
  final.mergeFrom(bodyComposer);

  // Comments never reached via a commentRangeStart/End pair in the body walk (a data
  // anomaly -- every Word-authored comment has both) get no real document position, so
  // their observations (if collected) carry a null anchorText rather than a guessed one.
  const orphans = Object.values(ctx.comments).filter((c) => !c.parentId && !ctx.emitted.has(c.id));
  if (orphans.length) {
    final.writeSynthetic(
      "\n\n## Unanchored comments\n\n" +
        orphans
          .map((c) => {
            ctx.emitted.add(c.id);
            if (ctx.observations) pushThreadObservations(ctx, c.id, null);
            return "- " + renderThread(ctx.comments, ctx.childrenMap, c.id);
          })
          .join("\n")
    );
  }
  final.writeSynthetic("\n");

  return {
    markdown: final.out, blocks: final.blocks, synthetic: final.synthetic, locked: final.locked,
    comments: ctx.comments, counts: ctx.counts, observations: ctx.observations,
  };
}

async function sha256Hex(arrayBuffer) {
  const digest = await crypto.subtle.digest("SHA-256", arrayBuffer);
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

/* ------------------------------------------------------------------ *
 * Public API
 * ------------------------------------------------------------------ */
export async function exportDocx(docxBytes, options = {}) {
  const { DOMParserImpl, annotate = false, filename = "document", collectObservations = false } = options;

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

  const result = render({ body, rels, comments, childrenMap, docDoc, filename }, annotate, collectObservations);
  const docHash = "sha256-" + (await sha256Hex(docxBytes));

  return {
    markdown: result.markdown,
    sourceMap: { docHash, blocks: result.blocks, synthetic: result.synthetic, locked: result.locked },
    comments: result.comments,
    counts: result.counts,
    ...(collectObservations ? { observations: result.observations } : {}),
  };
}

export const _internal = { NS };
