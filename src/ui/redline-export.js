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
import { clusterPasses } from "../passes.js";
import { renderSlice, sliceFilename } from "../ooxml/slice.js";

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
 * Reviewer-pass slicer (task spec: "Discovery" table + slice export). aliasMap is never
 * mutated in place -- every merge produces a new Map so the caller can keep the observation
 * array untouched and just re-cluster from scratch, matching this file's existing
 * pure-core/DOM-adapter split.
 * ------------------------------------------------------------------ */

export function applyAuthorAliases(observations, aliasMap) {
  if (!aliasMap || !aliasMap.size) return observations;
  return observations.map((o) => {
    const canonical = aliasMap.get(o.author);
    return canonical && canonical !== o.author ? { ...o, author: canonical } : o;
  });
}

// Merges 2+ author strings (as currently shown in the discovery table, i.e. already
// passed through any prior aliasing) into one canonical name -- the first in the list.
// Composes with existing aliases: anything already mapped to one of the merged names gets
// redirected too, so repeated merges of overlapping groups stay consistent.
export function mergeAuthorsInto(aliasMap, authors) {
  if (!authors || authors.length < 2) return aliasMap;
  const canonical = authors[0];
  const next = new Map(aliasMap);
  for (const [orig, mapped] of aliasMap) {
    if (authors.includes(mapped)) next.set(orig, canonical);
  }
  for (const a of authors) next.set(a, canonical);
  return next;
}

export function passesFor(observations, aliasMap) {
  return clusterPasses(applyAuthorAliases(observations, aliasMap));
}

export function metadataStrippedWarning() {
  return (
    "Every tracked change and comment in this document shows the same author with no date " +
    '(Word’s "Remove personal information from file properties on save," or similar, ' +
    "strips this before you can separate reviewer passes). All edits and comments are shown " +
    "below as a single group."
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

// Writes every pass's slice directly into a user-picked folder via the File System Access
// API (Chromium-only by design -- Firefox lacks showDirectoryPicker; no zip/sequential-
// download fallback, see the reviewer-pass-slicer task's Export UX decision).
async function writeSlicesToDirectory(dirHandle, bytes, passes, options) {
  let count = 0;
  for (const pass of passes) {
    const md = await renderSlice(bytes, pass, options);
    const name = sliceFilename(options.filename, pass);
    const fileHandle = await dirHandle.getFileHandle(name, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(md);
    await writable.close();
    count++;
  }
  return count;
}

function escapeHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
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

  const passesCard = root.querySelector("#rl-passes");
  const passesWarningEl = root.querySelector("#rl-passes-warning");
  const passesTableBody = root.querySelector("#rl-passes-tbody");
  const mergeBtn = root.querySelector("#rl-merge-selected");
  const exportAllBtn = root.querySelector("#rl-export-all");
  const passesStatusEl = root.querySelector("#rl-passes-status");
  const passesErrorEl = root.querySelector("#rl-passes-error");

  let current = null; // { markdown, sourceMap, comments, observations, filename, bytes }
  let aliasMap = new Map();

  function showError(message) {
    errorEl.textContent = message;
    errorEl.style.display = message ? "block" : "none";
  }
  function showPassesError(message) {
    passesErrorEl.textContent = message;
    passesErrorEl.style.display = message ? "block" : "none";
  }

  function reset() {
    current = null;
    aliasMap = new Map();
    resultsEl.style.display = "none";
    summaryEl.textContent = "";
    warningEl.textContent = "";
    outputEl.value = "";
    passesCard.style.display = "none";
    passesTableBody.innerHTML = "";
    passesStatusEl.textContent = "";
    showPassesError("");
  }

  function selectedAuthors() {
    const checked = [...passesTableBody.querySelectorAll('input[type="checkbox"]:checked')];
    return [...new Set(checked.map((cb) => cb.dataset.author))];
  }

  function renderPasses() {
    if (!current) return;
    const { passes, metadataStripped } = passesFor(current.observations, aliasMap);
    current.passes = passes;

    passesWarningEl.textContent = metadataStripped ? metadataStrippedWarning() : "";
    passesWarningEl.style.display = metadataStripped ? "block" : "none";

    passesTableBody.innerHTML = "";
    passes.forEach((pass, idx) => {
      const tr = document.createElement("tr");
      const passLabel = pass.undated ? "undated" : pass.passDate;
      tr.innerHTML = `
        <td><input type="checkbox" data-author="${escapeHtml(pass.author)}"></td>
        <td>${escapeHtml(pass.author)}</td>
        <td>${escapeHtml(passLabel)}</td>
        <td>${pass.counts.insertions + pass.counts.deletions}</td>
        <td>${pass.counts.comments}</td>
        <td>${pass.counts.replies}</td>
        <td><button type="button" data-pass-idx="${idx}" class="rl-export-slice">Export slice</button></td>
      `;
      passesTableBody.appendChild(tr);
    });

    passesTableBody.querySelectorAll(".rl-export-slice").forEach((btn) => {
      btn.addEventListener("click", async () => {
        showPassesError("");
        const pass = current.passes[Number(btn.dataset.passIdx)];
        try {
          const md = await renderSlice(current.bytes, pass, { DOMParserImpl: DOMParser, filename: current.filename });
          downloadText(md, sliceFilename(current.filename, pass), "text/markdown");
        } catch (err) {
          showPassesError(`Could not render slice for ${pass.label}: ${err.message}`);
        }
      });
    });

    passesCard.style.display = "block";
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
      exported = await exportDocx(bytes, { DOMParserImpl: DOMParser, annotate: true, filename, collectObservations: true });
    } catch (err) {
      showError(err.message);
      return;
    }

    current = { ...exported, filename, bytes };
    outputEl.value = exported.markdown;
    summaryEl.textContent = buildSummary(exported.counts, exported.comments);
    const warning = checkWordCountWarning(exported.markdown);
    warningEl.textContent = warning || "";
    warningEl.style.display = warning ? "block" : "none";
    resultsEl.style.display = "block";

    renderPasses();
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

  mergeBtn.addEventListener("click", () => {
    if (!current) return;
    showPassesError("");
    const authors = selectedAuthors();
    if (authors.length < 2) {
      showPassesError("Select two or more reviewer rows (from different names) to merge them into one reviewer.");
      return;
    }
    aliasMap = mergeAuthorsInto(aliasMap, authors);
    renderPasses();
  });

  exportAllBtn.addEventListener("click", async () => {
    if (!current || !current.passes || !current.passes.length) return;
    showPassesError("");
    if (typeof window.showDirectoryPicker !== "function") {
      showPassesError(
        "Export all slices needs the File System Access API, available in Chrome or Edge. " +
          "Use \"Export slice\" on individual rows instead."
      );
      return;
    }
    try {
      const dirHandle = await window.showDirectoryPicker();
      passesStatusEl.textContent = "Writing slices…";
      const count = await writeSlicesToDirectory(dirHandle, current.bytes, current.passes, {
        DOMParserImpl: DOMParser,
        filename: current.filename,
      });
      passesStatusEl.textContent = `Wrote ${count} slice${count === 1 ? "" : "s"} to the selected folder.`;
    } catch (err) {
      passesStatusEl.textContent = "";
      if (err && err.name === "AbortError") return; // user cancelled the picker
      showPassesError(`Could not write slices: ${err.message}`);
    }
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
