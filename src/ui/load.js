// Document + persona load (spec §6.1 step 1, M4 architecture doc §2). Drag-drop AND
// file-picker both funnel through the same pure core, per the "pure core, thin DOM"
// principle -- the DOM adapters at the bottom exist only to get a File into that core.
import { exportDocx } from "../ooxml/export.js";
import { D4_ERROR_MESSAGE } from "../ooxml/inject.js";
import { parsePersona } from "../persona.js";

// Issue #16's second bullet: block comment-reply documents until #15 (export.js's
// overlapping-reply nesting bug) is fixed, since a reply thread fails G1 on a byte-perfect
// echo with zero edits -- a confusing mid-flow failure this preflight turns into a clear
// one at the door instead.
export const COMMENT_REPLY_MESSAGE =
  "This document contains a comment reply (a threaded response to another comment). Resolve or remove replies in Word first, then re-export and re-run the review.";

export { D4_ERROR_MESSAGE };

// M4d PR-2, F-2: structural fence. inject.js's splitRun can only reposition a plain w:t
// run; a run containing a soft break, tab, or tracked deletion crashes it later with
// "splitRun: invalid range". Stays until child-aware splitRun lands (v2 backlog) -- this is
// NOT the temporary annotation fence below.
export const STRUCTURAL_FENCE_MESSAGE =
  "This document contains soft line breaks, tabs, or tracked deletions inside paragraphs, which this version cannot reposition safely. Remove them (Ctrl+Shift+8 in Word shows soft breaks as a bent arrow) or use a copy saved without them.";

// M4d PR-2, F-3 (fenced, not fixed here): export.js renders pre-existing comments and
// tracked changes AS CriticMarkup, which the validator can't yet tell apart from a model's
// real edits (sentinelization is M6a scope, alongside the comment-reply export fix). Until
// then, any pre-existing annotation at all is rejected at the door. Broader than -- and
// TEMPORARY unlike -- the structural fence and the US-7 checks above: this is the first
// thing M6a removes.
export const ANNOTATION_FENCE_MESSAGE =
  "This document already contains comments or tracked changes. This version reviews clean documents only; support for annotated documents is the next milestone.";

export const NOTHING_TO_RESPOND_MESSAGE =
  "This document does not contain any comments or tracked changes. There is nothing here to respond to.";

export function baseName(filename) {
  return (filename || "document").replace(/\.docx$/i, "");
}

// Reject non-docx before ever attempting to parse (spec: "Reject non-docx with a clear
// message").
export function checkFileExtension(filename) {
  if (!/\.docx$/i.test(filename || "")) {
    return { ok: false, message: `"${filename || "This file"}" is not a .docx file. Choose a Word .docx document.` };
  }
  return { ok: true };
}

// Issue #16: pre-existing tracked changes or comment replies must block at load time, not
// surface later as a confusing G1/G2 failure. `counts.ins/del/sub` > 0 means the export
// found pre-existing w:ins/w:del/substitution in the document (exportDocx renders those as
// CriticMarkup for display, per spec §5.1 -- their presence here is exactly the "already
// has tracked changes" signal); any comment with a parentId is a reply.
//
// M4d adds two more checks (F-2, F-3) and reports every applicable reason together, not
// just the first hit, so a document failing multiple checks gets one complete dialog
// instead of a fix-one-reload-hit-the-next loop.
export function checkPreflight(exportResult, flowType = "run-review") {
  const { counts, comments, structuralHazard } = exportResult;
  const reasons = [];

  if (structuralHazard) reasons.push(STRUCTURAL_FENCE_MESSAGE);

  if (flowType === "run-review") {
    // Structural fence only! All other annotation fences come down.
  } else if (flowType === "respond-review") {
    const hasAnnotation = counts.ins > 0 || counts.del > 0 || counts.sub > 0 || Object.keys(comments || {}).length > 0;
    if (!hasAnnotation) {
      reasons.push(NOTHING_TO_RESPOND_MESSAGE);
    }
  }

  if (reasons.length) return { ok: false, message: reasons.join("\n\n") };
  return { ok: true };
}

// Pure core: raw docx bytes + the original filename -> {ok, exported, filename} or
// {ok:false, message}. `originalFilename` keeps its extension for the checks above; the
// export itself is tagged with the extension-stripped base name (exportDocx's header
// appends ".docx" itself -- see ooxml/export.js's buildHeaderTracked).
export async function loadDocxFromBytes(docxBytes, { originalFilename, DOMParserImpl, flowType = "run-review" } = {}) {
  const extCheck = checkFileExtension(originalFilename);
  if (!extCheck.ok) return extCheck;

  const filename = baseName(originalFilename);
  let exported;
  try {
    exported = await exportDocx(docxBytes, { filename, DOMParserImpl });
  } catch (err) {
    return { ok: false, message: `Could not read "${originalFilename}": ${err.message}` };
  }

  const preflight = checkPreflight(exported, flowType);
  if (!preflight.ok) return preflight;

  return { ok: true, exported, filename, docxBytes };
}

export function loadPersonaFromText(markdown, { filename } = {}) {
  return parsePersona(markdown, { filename });
}

/* ------------------------------------------------------------------ *
 * DOM adapters -- wire a drop zone / file input to the pure core above.
 * ------------------------------------------------------------------ */

export async function readFileAsArrayBuffer(file) {
  if (typeof file.arrayBuffer === "function") return file.arrayBuffer();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

export function readFileAsText(file) {
  if (typeof file.text === "function") return file.text();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsText(file);
  });
}

// Wires both drag-and-drop and the OS file-picker's "click to browse" affordance onto the
// same drop-zone element, since both must land on `onFiles` identically (spec: "Drag-drop
// AND file-picker ... make the copy true").
export function attachDropZone(el, inputEl, { onFiles }) {
  const stop = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };
  el.addEventListener("dragover", (e) => {
    stop(e);
    el.classList.add("ar-drag-over");
  });
  el.addEventListener("dragleave", (e) => {
    stop(e);
    el.classList.remove("ar-drag-over");
  });
  el.addEventListener("drop", (e) => {
    stop(e);
    el.classList.remove("ar-drag-over");
    const files = e.dataTransfer && e.dataTransfer.files ? [...e.dataTransfer.files] : [];
    if (files.length) onFiles(files);
  });
  if (inputEl) {
    el.addEventListener("click", () => inputEl.click());
    inputEl.addEventListener("change", () => {
      const files = inputEl.files ? [...inputEl.files] : [];
      if (files.length) onFiles(files);
      inputEl.value = "";
    });
  }
}
