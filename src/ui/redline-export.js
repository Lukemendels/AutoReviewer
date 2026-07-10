// Redline export shell (M6 "Respond to Review", first slice, spec-workbench.md). A
// read-only tool: docx in, CriticMarkup markdown out. No injection, no ratification, no
// preflight gate -- this surface exists precisely so a document WITH pre-existing tracked
// changes/comments (which src/ui/load.js's checkPreflight would reject) can still be
// exported for pasting into an external LLM chat.
//
// Deliberately does not import src/ui/load.js, src/ui/state.js, src/ui/app.js,
// src/validate.js, src/ooxml/inject.js, src/ratify*, or src/audit.js -- see the task
// spec's HARD CONSTRAINT. exportDocx() is the entire pipeline this module drives.
import { exportDocx } from "../ooxml/export.js";
// CHUNK_WORD_THRESHOLD is a standalone constant in prompt.js (12000) with a single,
// light, non-circular dependency (persona.js, itself dependency-free) -- tests/
// app.chunkFlow.test.js already imports it the same way, so this is an established,
// safe import rather than a new risk.
import { CHUNK_WORD_THRESHOLD } from "../prompt.js";

/* ------------------------------------------------------------------ *
 * Pure core
 * ------------------------------------------------------------------ */

export function baseName(filename) {
  return (filename || "document").replace(/\.docx$/i, "");
}

export function wordCount(text) {
  const trimmed = (text || "").trim();
  return trimmed ? trimmed.split(/\s+/).length : 0;
}

// Top-level threads vs. total nodes (spec: "count top-level threads and total nodes
// separately if that is cheap") -- cheap here, since buildCommentsData already gives us
// parentId per node.
export function buildSummary(counts, comments) {
  const nodes = Object.values(comments || {});
  const total = nodes.length;
  const topLevel = nodes.filter((c) => c.parentId == null).length;

  const plural = (n, word) => `${n} ${word}${n === 1 ? "" : "s"}`;
  const parts = [
    plural(counts.ins, "insertion"),
    plural(counts.del, "deletion"),
    plural(counts.sub, "substitution"),
    total === topLevel ? plural(total, "comment") : `${plural(topLevel, "comment thread")} (${total} total)`,
  ];
  return parts.join(" · ");
}

export function checkWordCountWarning(markdown) {
  const words = wordCount(markdown);
  if (words <= CHUNK_WORD_THRESHOLD) return null;
  return (
    `This document is ${words.toLocaleString()} words, which is likely too long for a single ` +
    `chat paste. Consider splitting it by top-level heading before pasting.`
  );
}

/* ------------------------------------------------------------------ *
 * DOM adapter
 * ------------------------------------------------------------------ */

async function readFileAsArrayBuffer(file) {
  if (typeof file.arrayBuffer === "function") return file.arrayBuffer();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function downloadText(text, filename, mimeType) {
  const blob = new Blob([text], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// copyText mirrors template.html's own copyText/legacyCopy: the async Clipboard API is
// unavailable on file://, so hidden-textarea execCommand('copy') is the PRIMARY path, not
// a fallback (spec's Non-negotiables).
function legacyCopy(text) {
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.style.position = "fixed";
  ta.style.opacity = "0";
  document.body.appendChild(ta);
  ta.focus();
  ta.select();
  let ok = false;
  try {
    ok = document.execCommand("copy");
  } catch (e) {
    ok = false;
  }
  document.body.removeChild(ta);
  return ok;
}
function copyText(text) {
  if (navigator.clipboard && window.isSecureContext) {
    return navigator.clipboard.writeText(text).then(() => true).catch(() => legacyCopy(text));
  }
  return Promise.resolve(legacyCopy(text));
}

function attachDropZone(el, inputEl, { onFiles }) {
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

function mount(root) {
  const dropzone = root.querySelector("#rl-dropzone");
  const fileInput = root.querySelector("#rl-file");
  const errorEl = root.querySelector("#rl-error");
  const summaryEl = root.querySelector("#rl-summary");
  const warningEl = root.querySelector("#rl-warning");
  const outputEl = root.querySelector("#rl-output");
  const resultsEl = root.querySelector("#rl-results");
  const copyBtn = root.querySelector("#rl-copy");
  const downloadMdBtn = root.querySelector("#rl-download-md");
  const downloadSourceMapBtn = root.querySelector("#rl-download-sourcemap");

  let current = null; // { markdown, sourceMap, comments, filename }

  function showError(message) {
    errorEl.textContent = message;
    errorEl.style.display = message ? "block" : "none";
  }

  function reset() {
    current = null;
    resultsEl.style.display = "none";
    summaryEl.textContent = "";
    warningEl.textContent = "";
    outputEl.value = "";
  }

  async function handleFiles(files) {
    const file = files[0];
    showError("");
    reset();
    if (!/\.docx$/i.test(file.name || "")) {
      showError(`"${file.name || "This file"}" is not a .docx file. Choose a Word .docx document.`);
      return;
    }
    const filename = baseName(file.name);
    let bytes;
    try {
      bytes = await readFileAsArrayBuffer(file);
    } catch (err) {
      showError(`Could not read "${file.name}": ${err.message}`);
      return;
    }
    let exported;
    try {
      exported = await exportDocx(bytes, { DOMParserImpl: DOMParser, annotate: true, filename });
    } catch (err) {
      showError(err.message);
      return;
    }

    current = { ...exported, filename };
    outputEl.value = exported.markdown;
    summaryEl.textContent = buildSummary(exported.counts, exported.comments);
    const warning = checkWordCountWarning(exported.markdown);
    warningEl.textContent = warning || "";
    warningEl.style.display = warning ? "block" : "none";
    resultsEl.style.display = "block";
  }

  attachDropZone(dropzone, fileInput, { onFiles: handleFiles });

  copyBtn.addEventListener("click", async () => {
    if (!current) return;
    const ok = await copyText(current.markdown);
    const original = copyBtn.textContent;
    copyBtn.textContent = ok ? "Copied ✓" : "Copy failed";
    setTimeout(() => {
      copyBtn.textContent = original;
    }, 1500);
  });

  downloadMdBtn.addEventListener("click", () => {
    if (!current) return;
    downloadText(current.markdown, `${current.filename}.redline.md`, "text/markdown");
  });

  downloadSourceMapBtn.addEventListener("click", () => {
    if (!current) return;
    const payload = { sourceMap: current.sourceMap, comments: current.comments };
    downloadText(JSON.stringify(payload, null, 2), `${current.filename}.sourcemap.json`, "application/json");
  });
}

function init() {
  const root = document.getElementById("rl-app");
  if (!root) return;
  mount(root);
}

if (typeof document !== "undefined") {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
}

export { mount };
