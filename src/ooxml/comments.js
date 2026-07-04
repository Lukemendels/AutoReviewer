// Create/extend the in-scope comment parts (spec §9.2; M3b plan D6 -- comments.xml,
// commentsExtended.xml, the document rels entries, and [Content_Types].xml overrides
// ONLY. commentsIds.xml and commentsExtensible.xml are OUT of scope for M3b: they are a
// follow-up milestone implemented against a Word-authored fixture so the schemas are
// copied from XML Word itself wrote rather than reconstructed from memory -- this module
// never writes those two parts, though upsertComments's return shape leaves room for them).
import { parseXml, wAttr, w14Attr, w15Attr } from "./parse.js";
import { serializePart } from "./serialize.js";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

const REL_TYPE_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const REL_TYPE_COMMENTS_EXT = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const CT_COMMENTS_EXT = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";

const COMMENTS_TEMPLATE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"></w:comments>`;
const COMMENTS_EXT_TEMPLATE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + `<w15:commentsEx xmlns:w15="${W15_NS}"></w15:commentsEx>`;
const RELS_TEMPLATE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + `<Relationships xmlns="${PKG_REL_NS}"></Relationships>`;
const CONTENT_TYPES_TEMPLATE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + `<Types xmlns="${CT_NS}"></Types>`;

function existingParaIds(commentsDoc) {
  const set = new Set();
  for (const el of commentsDoc.getElementsByTagName("*")) {
    if (el.localName !== "p") continue;
    const pid = w14Attr(el, "paraId");
    if (pid) set.add(pid);
  }
  return set;
}

// 8-hex uppercase, random and checked for uniqueness against every paraId already in the
// document (collision astronomically unlikely but checked defensively -- matching
// fixtures/generate.py's add_comment_threading's own paraId-generation spirit).
function freshParaId(used) {
  for (;;) {
    let hex = "";
    for (let i = 0; i < 8; i++) hex += Math.floor(Math.random() * 16).toString(16);
    hex = hex.toUpperCase();
    if (!used.has(hex)) {
      used.add(hex);
      return hex;
    }
  }
}

function buildCommentElement(doc, comment, paraId) {
  const commentEl = doc.createElementNS(W_NS, "w:comment");
  commentEl.setAttributeNS(W_NS, "w:id", String(comment.id));
  commentEl.setAttributeNS(W_NS, "w:author", comment.author);
  commentEl.setAttributeNS(W_NS, "w:date", comment.date);
  commentEl.setAttributeNS(W_NS, "w:initials", "");

  const p = doc.createElementNS(W_NS, "w:p");
  p.setAttributeNS(W14_NS, "w14:paraId", paraId);

  const pPr = doc.createElementNS(W_NS, "w:pPr");
  const pStyle = doc.createElementNS(W_NS, "w:pStyle");
  pStyle.setAttributeNS(W_NS, "w:val", "CommentText");
  pPr.appendChild(pStyle);
  p.appendChild(pPr);

  // The comment-reference marker run Word itself always emits (matches real Word-authored
  // comments.xml, per spec §9.2's repair-prompt-linkage concern) -- collectRunsText in
  // export.js safely contributes "" for it (no w:t/w:delText inside).
  const refRun = doc.createElementNS(W_NS, "w:r");
  const refRPr = doc.createElementNS(W_NS, "w:rPr");
  const rStyle = doc.createElementNS(W_NS, "w:rStyle");
  rStyle.setAttributeNS(W_NS, "w:val", "CommentReference");
  refRPr.appendChild(rStyle);
  refRun.appendChild(refRPr);
  refRun.appendChild(doc.createElementNS(W_NS, "w:annotationRef"));
  p.appendChild(refRun);

  const textRun = doc.createElementNS(W_NS, "w:r");
  const t = doc.createElementNS(W_NS, "w:t");
  t.appendChild(doc.createTextNode(comment.text));
  textRun.appendChild(t);
  p.appendChild(textRun);

  commentEl.appendChild(p);
  return commentEl;
}

function buildCommentExElement(doc, paraId, parentParaId) {
  const el = doc.createElementNS(W15_NS, "w15:commentEx");
  el.setAttributeNS(W15_NS, "w15:paraId", paraId);
  el.setAttributeNS(W15_NS, "w15:done", "0");
  if (parentParaId) el.setAttributeNS(W15_NS, "w15:paraIdParent", parentParaId);
  return el;
}

// Existing comment id (string, as read from w:id) -> paraId, seeded from whatever's
// already in commentsDoc so a reply's parentId (an id from either an existing comment or
// one earlier in this same batch) always resolves.
function buildIdToParaId(commentsDoc) {
  const map = {};
  for (const el of commentsDoc.getElementsByTagName("*")) {
    if (el.localName !== "comment") continue;
    const id = wAttr(el, "id");
    const p = [...el.children].find((c) => c.localName === "p");
    const pid = p && w14Attr(p, "paraId");
    if (id != null && pid) map[id] = pid;
  }
  return map;
}

function ensureRelationship(doc, type, target) {
  let maxId = 0;
  for (const el of doc.getElementsByTagName("*")) {
    if (el.localName !== "Relationship") continue;
    const m = /^rId(\d+)$/.exec(el.getAttribute("Id") || "");
    if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
    if (el.getAttribute("Type") === type) return; // already linked
  }
  const rel = doc.createElementNS(PKG_REL_NS, "Relationship");
  rel.setAttribute("Id", `rId${maxId + 1}`);
  rel.setAttribute("Type", type);
  rel.setAttribute("Target", target);
  doc.documentElement.appendChild(rel);
}

function ensureContentTypeOverride(doc, partName, contentType) {
  for (const el of doc.getElementsByTagName("*")) {
    if (el.localName === "Override" && el.getAttribute("PartName") === partName) return; // already present
  }
  const override = doc.createElementNS(CT_NS, "Override");
  override.setAttribute("PartName", partName);
  override.setAttribute("ContentType", contentType);
  doc.documentElement.appendChild(override);
}

export function upsertComments(existingParts, newComments, options = {}) {
  if (!newComments || !newComments.length) {
    return { ...existingParts };
  }
  const { DOMParserImpl, XMLSerializerImpl } = options;

  const commentsDoc = parseXml(existingParts.commentsXml || COMMENTS_TEMPLATE, DOMParserImpl);
  const extDoc = parseXml(existingParts.commentsExtendedXml || COMMENTS_EXT_TEMPLATE, DOMParserImpl);
  const relsDoc = parseXml(existingParts.relsXml || RELS_TEMPLATE, DOMParserImpl);
  const ctDoc = parseXml(existingParts.contentTypesXml || CONTENT_TYPES_TEMPLATE, DOMParserImpl);

  const usedParaIds = existingParaIds(commentsDoc);
  const idToParaId = buildIdToParaId(commentsDoc);

  for (const comment of newComments) {
    const paraId = freshParaId(usedParaIds);
    idToParaId[String(comment.id)] = paraId;

    commentsDoc.documentElement.appendChild(buildCommentElement(commentsDoc, comment, paraId));

    const parentParaId = comment.parentId != null ? idToParaId[String(comment.parentId)] : null;
    extDoc.documentElement.appendChild(buildCommentExElement(extDoc, paraId, parentParaId));
  }

  ensureRelationship(relsDoc, REL_TYPE_COMMENTS, "comments.xml");
  ensureRelationship(relsDoc, REL_TYPE_COMMENTS_EXT, "commentsExtended.xml");
  ensureContentTypeOverride(ctDoc, "/word/comments.xml", CT_COMMENTS);
  ensureContentTypeOverride(ctDoc, "/word/commentsExtended.xml", CT_COMMENTS_EXT);

  return {
    ...existingParts,
    commentsXml: serializePart(commentsDoc, XMLSerializerImpl),
    commentsExtendedXml: serializePart(extDoc, XMLSerializerImpl),
    relsXml: serializePart(relsDoc, XMLSerializerImpl),
    contentTypesXml: serializePart(ctDoc, XMLSerializerImpl),
  };
}
