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
export function checkPreflight(exportResult) {
  const { counts, comments } = exportResult;
  if (counts.ins || counts.del || counts.sub) {
    return { ok: false, message: D4_ERROR_MESSAGE };
  }
  if (Object.values(comments || {}).some((c) => c.parentId != null)) {
    return { ok: false, message: COMMENT_REPLY_MESSAGE };
  }
  return { ok: true };
}

// Pure core: raw docx bytes + the original filename -> {ok, exported, filename} or
// {ok:false, message}. `originalFilename` keeps its extension for the checks above; the
// export itself is tagged with the extension-stripped base name (exportDocx's header
// appends ".docx" itself -- see ooxml/export.js's buildHeaderTracked).
export async function loadDocxFromBytes(docxBytes, { originalFilename, DOMParserImpl } = {}) {
  const extCheck = checkFileExtension(originalFilename);
  if (!extCheck.ok) return extCheck;

  const filename = baseName(originalFilename);
  let exported;
  try {
    exported = await exportDocx(docxBytes, { filename, DOMParserImpl });
  } catch (err) {
    return { ok: false, message: `Could not read "${originalFilename}": ${err.message}` };
  }

  const preflight = checkPreflight(exported);
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
