import { validate, validateText, resolveEdits } from "../validate.js";
import { splitIntoChunks, translateEdits } from "../chunk.js";
import { buildPrompt } from "../prompt.js";
import { composeRepair } from "../repair.js";
import { buildAuditRecord } from "../audit.js";
import { saveSession, loadSession } from "../session.js";
import { createRatificationState, renderRatificationUI } from "./ratify.js";
import { createAppState, STATES } from "./state.js";
import { loadDocxFromBytes, loadPersonaFromText, attachDropZone, readFileAsArrayBuffer, readFileAsText } from "./load.js";
import { extractCandidates } from "../envelope.js";
import { DEMO_DOCX_BASE64 } from "./demo-doc.js";
import { unzip, readEntry } from "../zip/reader.js";
import { writeZip } from "../zip/writer.js";
import { parseXml } from "../ooxml/parse.js";
import { injectEdits } from "../ooxml/inject.js";
import { upsertComments } from "../ooxml/comments.js";
import { serializePart } from "../ooxml/serialize.js";
import { diffWords } from "./diff.js";

const FLOWS = [
  { id: "run-review", label: "Run Review" },
  { id: "respond-review", label: "Respond to Review" },
  { id: "train-persona", label: "Train Persona" },
];

function renderShell(root) {
  const nav = document.createElement("nav");
  nav.className = "ar-tabs";

  const panels = document.createElement("div");
  panels.className = "ar-panels";

  for (const flow of FLOWS) {
    const tab = document.createElement("button");
    tab.type = "button";
    tab.className = "ar-tab";
    tab.dataset.flow = flow.id;
    tab.textContent = flow.label;
    nav.appendChild(tab);

    const panel = document.createElement("section");
    panel.className = "ar-panel";
    panel.dataset.flow = flow.id;
    if (flow.id === "run-review") {
      renderRunReviewPanel(panel);
    } else {
      panel.innerHTML = `<p class="ar-coming-soon">${flow.label} is not implemented yet.</p>`;
    }
    panels.appendChild(panel);

    tab.addEventListener("click", () => selectFlow(flow.id));
  }

  root.appendChild(nav);
  root.appendChild(panels);
  selectFlow(FLOWS[0].id);
}

function selectFlow(flowId) {
  for (const tab of document.querySelectorAll(".ar-tab")) {
    tab.classList.toggle("active", tab.dataset.flow === flowId);
  }
  for (const panel of document.querySelectorAll(".ar-panel")) {
    panel.classList.toggle("active", panel.dataset.flow === flowId);
  }
}

function base64ToArrayBuffer(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function renderDiff(container, segments) {
  container.innerHTML = "";
  container.className = "ar-diff";
  for (const seg of segments) {
    const el = document.createElement(seg.type === "same" ? "span" : seg.type === "del" ? "del" : "ins");
    el.textContent = seg.text;
    container.appendChild(el);
  }
}

// Pure core of the Inject flow (spec §9's write-back): re-unzip the ORIGINAL docx bytes
// (bodyPath/runIndex are stable, deterministic offsets into that same original structure
// the source map was built from -- M3b plan's wiring note) rather than reusing anything
// already parsed, inject accepted edits, upsert any newly-created comments, and re-zip.
// Kept separate from the DOM-triggering download step below so it's independently
// testable without a browser Blob/URL implementation. DOMParserImpl/XMLSerializerImpl are
// optional, defaulting to the browser globals (matching every other ooxml/* module) --
// tests pass @xmldom/xmldom explicitly, since happy-dom's DOMParser doesn't reliably
// parse a real, namespace-heavy document.xml (verified: it silently drops the entire body).
export async function buildReviewedDocx({
  docxBytes,
  acceptedEdits,
  sourceMap,
  author,
  date,
  DOMParserImpl,
  XMLSerializerImpl,
}) {
  const zip = await unzip(docxBytes);
  const docDoc = parseXml(await readEntry(zip, "word/document.xml"), DOMParserImpl);
  const { newComments } = injectEdits(docDoc, acceptedEdits, sourceMap, { author, date });
  const mutatedParts = { "word/document.xml": serializePart(docDoc, XMLSerializerImpl) };

  if (newComments.length) {
    const existingParts = {
      commentsXml: await readEntry(zip, "word/comments.xml"),
      commentsExtendedXml: await readEntry(zip, "word/commentsExtended.xml"),
      relsXml: await readEntry(zip, "word/_rels/document.xml.rels"),
      contentTypesXml: await readEntry(zip, "[Content_Types].xml"),
    };
    const updated = upsertComments(existingParts, newComments, { DOMParserImpl, XMLSerializerImpl });
    mutatedParts["word/comments.xml"] = updated.commentsXml;
    mutatedParts["word/commentsExtended.xml"] = updated.commentsExtendedXml;
    mutatedParts["word/_rels/document.xml.rels"] = updated.relsXml;
    mutatedParts["[Content_Types].xml"] = updated.contentTypesXml;
  }

  return writeZip(zip, mutatedParts);
}

function downloadBlob(bytes, filename) {
  const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function downloadJson(obj, filename) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ------------------------------------------------------------------ *
 * Run Review flow (spec §6.1): load -> persona (optional) -> compose prompt -> copy ->
 * paste response -> validate -> ratify -> inject. Driven entirely by state.js's app state
 * machine; every render() is a full re-render from state/context, same pattern as
 * ratify.js's renderRatificationUI.
 * ------------------------------------------------------------------ */
function renderRunReviewPanel(panel) {
  const appState = createAppState();
  let loadError = null;
  // Set by renderRatifyStep; read by handleDownloadSession so a session saved mid- or
  // post-ratification carries the human's real decisions (ratify.js's state is ephemeral,
  // recreated fresh each time RATIFYING is entered -- see ratify.js -- so it's never part
  // of state.js's own context).
  let currentRatifyState = null;

  function currentDecisions() {
    if (!currentRatifyState) return [];
    return currentRatifyState.rows.map((r) => ({ id: r.id, decision: r.decision, reviewed: r.reviewed }));
  }

  function handleDownloadSession() {
    const ctx = appState.context;
    const saved = saveSession({ state: appState.state, context: ctx, decisions: currentDecisions() });
    downloadJson(saved, `${ctx.filename} — session.json`);
  }

  async function handleResumeFiles(files) {
    const file = files[0];
    if (!file) return;
    if (!/\.json$/i.test(file.name)) {
      loadError = `"${file.name}" is not a session .json file.`;
      render();
      return;
    }
    let parsed;
    try {
      parsed = JSON.parse(await readFileAsText(file));
    } catch (err) {
      loadError = `Could not parse "${file.name}" as a session file: ${err.message}`;
      render();
      return;
    }
    try {
      const { state, context } = loadSession(parsed);
      appState.hydrate({ state, context });
    } catch (err) {
      loadError = `Could not resume session: ${err.message}`;
      render();
    }
  }

  function render() {
    panel.innerHTML = "";
    const state = appState.state;
    const ctx = appState.context;

    if (state === STATES.EMPTY) {
      currentRatifyState = null;
      renderLoadStep(panel);
      return;
    }

    renderPersonaControls(panel, ctx);
    renderSessionBar(panel);

    if (state === STATES.DOC_LOADED || state === STATES.PROMPT_READY) {
      renderPromptStep(panel, ctx);
    } else if (state === STATES.AWAITING_RESPONSE || state === STATES.VALIDATION_FAILED) {
      renderPromptStep(panel, ctx, { collapsed: true });
      renderResponseStep(panel, ctx);
    } else if (state === STATES.RATIFYING) {
      renderRatifyStep(panel, ctx);
    } else if (state === STATES.INJECTED) {
      renderInjectedStep(panel, ctx);
    }
  }

  // ---- doc + persona loading ----

  // Composes the next prompt to show. Chunk mode (M4c, spec §6.4) is entered here, the
  // first time a document's full-document prompt comes back overThreshold -- `chunkMode`
  // is only ever set true once, from DOC_LOADED (state.js's enterChunkMode guard), which is
  // exactly the state right after loadDocument() and before this function's very first
  // call for that document. A LATER persona swap mid-chunk-review (loadPersona() already
  // transitioned to PROMPT_READY by the time app.js gets here) takes the other branch:
  // rebuild only the CURRENT chunk's own prompt with the new persona, leaving chunk
  // progress (chunkIndex/chunkEdits) untouched -- it doesn't re-enter chunk mode, so it
  // never hits enterChunkMode's DOC_LOADED-only guard.
  function regeneratePrompt(persona) {
    const ctx = appState.context;

    if (ctx.chunkMode) {
      const chunk = ctx.chunks[ctx.chunkIndex];
      const built = buildPrompt({
        persona,
        exportedMarkdown: chunk.exportedMarkdown,
        filename: ctx.filename,
        chunk: { index: ctx.chunkIndex, total: ctx.chunks.length },
      });
      appState.setPrompt(built);
      return;
    }

    const built = buildPrompt({ persona, exportedMarkdown: ctx.exported.markdown, filename: ctx.filename });
    if (!built.overThreshold) {
      appState.setPrompt(built);
      return;
    }

    const chunks = splitIntoChunks(ctx.exported.markdown, ctx.exported.sourceMap);
    appState.enterChunkMode({ chunks });
    const chunk0Built = buildPrompt({
      persona,
      exportedMarkdown: chunks[0].exportedMarkdown,
      filename: ctx.filename,
      chunk: { index: 0, total: chunks.length },
    });
    appState.setPrompt(chunk0Built);
  }

  async function handleDocFiles(files) {
    const file = files[0];
    if (!file) return;
    loadError = null;
    render();
    const statusEl = panel.querySelector("#ar-load-status");
    if (statusEl) statusEl.textContent = `Reading "${file.name}"...`;
    let bytes;
    try {
      bytes = await readFileAsArrayBuffer(file);
    } catch (err) {
      loadError = `Could not read "${file.name}": ${err.message}`;
      render();
      return;
    }
    const result = await loadDocxFromBytes(bytes, { originalFilename: file.name });
    if (!result.ok) {
      loadError = result.message;
      render();
      return;
    }
    appState.loadDocument({ docxBytes: result.docxBytes, filename: result.filename, exported: result.exported });
    regeneratePrompt(null); // default persona
  }

  async function handleTryDemo() {
    const bytes = base64ToArrayBuffer(DEMO_DOCX_BASE64);
    const result = await loadDocxFromBytes(bytes, { originalFilename: "plain-paragraphs.docx" });
    if (!result.ok) {
      loadError = result.message;
      render();
      return;
    }
    appState.loadDocument({ docxBytes: result.docxBytes, filename: result.filename, exported: result.exported });
    regeneratePrompt(null);
  }

  async function handlePersonaFiles(files) {
    const file = files[0];
    if (!file) return;
    if (!/\.md$/i.test(file.name)) {
      loadError = `"${file.name}" is not a persona .md file.`;
      render();
      return;
    }
    const text = await readFileAsText(file);
    const persona = loadPersonaFromText(text, { filename: file.name });
    appState.loadPersona(persona);
    regeneratePrompt(persona);
  }

  // ---- step renderers ----

  function renderLoadStep(container) {
    const wrap = document.createElement("div");
    wrap.innerHTML = `
      <p class="ar-hint">Drop a Word <code>.docx</code> document here, or click to choose one from disk.</p>
      <div id="ar-dropzone" class="ar-dropzone" tabindex="0">
        <p>Drop a <code>.docx</code> here, or click to browse.</p>
      </div>
      <input type="file" id="ar-doc-input" accept=".docx" style="display:none" />
      <p id="ar-load-status" class="ar-hint">${loadError ? escapeHtml(loadError) : ""}</p>
      <p class="ar-hint">No document handy? <button type="button" id="ar-try-demo" class="ar-link">Try the demo</button> instead.</p>
      <p class="ar-hint">Resuming a review? <button type="button" id="ar-resume-session" class="ar-link">Resume session</button>
      from a saved session <code>.json</code>.</p>
      <input type="file" id="ar-resume-input" accept=".json" style="display:none" />
    `;
    container.appendChild(wrap);
    const dz = wrap.querySelector("#ar-dropzone");
    const input = wrap.querySelector("#ar-doc-input");
    attachDropZone(dz, input, { onFiles: handleDocFiles });
    wrap.querySelector("#ar-try-demo").addEventListener("click", handleTryDemo);
    const resumeInput = wrap.querySelector("#ar-resume-input");
    wrap.querySelector("#ar-resume-session").addEventListener("click", () => resumeInput.click());
    resumeInput.addEventListener("change", () => {
      const files = resumeInput.files ? [...resumeInput.files] : [];
      if (files.length) handleResumeFiles(files);
      resumeInput.value = "";
    });
  }

  // Available in every non-EMPTY state (spec §5.4 / M4b build plan §3).
  function renderSessionBar(container) {
    const wrap = document.createElement("div");
    wrap.className = "ar-session-bar";
    wrap.innerHTML = `<button type="button" id="ar-download-session" class="ar-link">Download session</button>`;
    container.appendChild(wrap);
    wrap.querySelector("#ar-download-session").addEventListener("click", handleDownloadSession);
  }

  function renderPersonaControls(container, ctx) {
    const wrap = document.createElement("div");
    wrap.className = "ar-persona-controls";
    const persona = ctx.persona;
    const warningsHtml =
      persona && persona.warnings && persona.warnings.length
        ? `<ul class="ar-persona-warnings">${persona.warnings.map((w) => `<li>${escapeHtml(w)}</li>`).join("")}</ul>`
        : "";
    wrap.innerHTML = `
      <p class="ar-hint">
        Persona: <strong>${escapeHtml(persona ? persona.name : "Default Persona (built-in)")}</strong>
        ${persona ? "" : "&mdash; drop an OKF persona .md to use a custom one."}
      </p>
      ${warningsHtml}
      <div id="ar-persona-dropzone" class="ar-dropzone ar-dropzone-small" tabindex="0">
        <p>Drop a persona <code>.md</code> here, or click to browse.</p>
      </div>
      <input type="file" id="ar-persona-input" accept=".md" style="display:none" />
    `;
    container.appendChild(wrap);
    const dz = wrap.querySelector("#ar-persona-dropzone");
    const input = wrap.querySelector("#ar-persona-input");
    attachDropZone(dz, input, { onFiles: handlePersonaFiles });
  }

  // M4c: the "Part i of N" indicator shown throughout chunk mode, both on the live prompt
  // step and its collapsed (AWAITING_RESPONSE/VALIDATION_FAILED) form -- the one visible
  // reminder of which chunk the user is currently copying/pasting for.
  function chunkPartHint(ctx) {
    if (!ctx.chunkMode) return "";
    return `<p class="ar-hint ar-chunk-indicator"><strong>Part ${ctx.chunkIndex + 1} of ${ctx.chunks.length}</strong> of this document (chunk mode, spec §6.4).</p>`;
  }

  function renderPromptStep(container, ctx, { collapsed = false } = {}) {
    const wrap = document.createElement("div");
    wrap.className = "ar-prompt-step";

    if (ctx.tokenEstimate == null) {
      wrap.innerHTML = `<p class="ar-hint">Composing prompt...</p>`;
      container.appendChild(wrap);
      return;
    }

    if (collapsed) {
      wrap.innerHTML = `
        ${chunkPartHint(ctx)}
        <details class="ar-prompt-details">
          <summary>Composed prompt (${ctx.tokenEstimate} tokens, ~${ctx.documentWordCount} words) &mdash; click to re-view</summary>
          <textarea class="ar-prompt-text" readonly>${escapeHtml(ctx.promptText)}</textarea>
          <button type="button" id="ar-recopy-prompt" class="ar-primary">Copy prompt again</button>
        </details>
      `;
      container.appendChild(wrap);
      wrap.querySelector("#ar-recopy-prompt").addEventListener("click", (e) => {
        if (typeof copyWithFeedback === "function") copyWithFeedback(e.currentTarget, ctx.promptText);
      });
      return;
    }

    // Not reached in practice once a document has gone through regeneratePrompt() (which
    // enters chunk mode itself the moment it sees overThreshold) -- kept as a defensive
    // fallback rather than assuming that invariant holds for every future caller.
    if (ctx.overThreshold && !ctx.chunkMode) {
      wrap.innerHTML = `
        <div class="ar-gate-failure">
          <p><strong>This document is over the single-prompt word threshold</strong> and chunk
          mode was not entered for it. A full round trip on a document this long is known to
          fail the fidelity gate (G2), so the prompt below is withheld rather than handing
          you one known not to work.</p>
        </div>
      `;
      container.appendChild(wrap);
      return;
    }

    wrap.innerHTML = `
      ${chunkPartHint(ctx)}
      <p class="ar-hint">Token estimate: ~${ctx.tokenEstimate} (prompt version ${escapeHtml(ctx.promptVersion)})</p>
      <textarea class="ar-prompt-text" readonly>${escapeHtml(ctx.promptText)}</textarea>
      <div class="ar-controls">
        <button type="button" id="ar-copy-prompt" class="ar-primary">Copy prompt</button>
      </div>
      <p class="ar-hint">Paste this into DHSChat, copy the reply, then paste it below.</p>
    `;
    container.appendChild(wrap);
    wrap.querySelector("#ar-copy-prompt").addEventListener("click", (e) => {
      if (typeof copyWithFeedback === "function") copyWithFeedback(e.currentTarget, ctx.promptText);
      appState.copyPrompt();
    });
  }

  function renderResponseStep(container, ctx) {
    const wrap = document.createElement("div");
    wrap.className = "ar-response-step";
    container.appendChild(wrap);

    if (ctx.validation && !ctx.validation.ok) {
      wrap.appendChild(ctx.validation.global ? buildGlobalGateFailureEl(ctx.validation) : buildGateFailureEl(ctx.validation, ctx));
    }

    const fieldWrap = document.createElement("div");
    fieldWrap.innerHTML = `
      <div class="ar-field">
        <label for="ar-response">Response</label>
        <textarea id="ar-response" class="ar-response" spellcheck="false"></textarea>
      </div>
      <div class="ar-controls">
        <button type="button" id="ar-validate" class="ar-primary">Validate</button>
      </div>
      <div id="ar-picker"></div>
    `;
    wrap.appendChild(fieldWrap);

    const responseEl = fieldWrap.querySelector("#ar-response");
    const pickerEl = fieldWrap.querySelector("#ar-picker");

    function runEnvelope(pastedText) {
      // Chunk mode's "document-sized" reference is THIS chunk's own exportedMarkdown, not
      // the full document's -- a real per-chunk response is only ever a fraction of the
      // full document's length, so using ctx.exported.markdown.length here would make the
      // envelope's own half-length threshold reject every legitimate chunk response.
      const referenceLength = ctx.chunkMode ? ctx.chunks[ctx.chunkIndex].exportedMarkdown.length : ctx.exported.markdown.length;
      const { candidates, noFencesFound } = extractCandidates(pastedText, { exportedLength: referenceLength });
      if (candidates.length === 1 && !noFencesFound) {
        runValidation(candidates[0].content);
        return;
      }
      pickerEl.innerHTML = noFencesFound
        ? `<p class="ar-hint">No fenced code block was found -- using the entire paste. If that's wrong, edit the text above and click Validate.</p>`
        : `<p class="ar-hint">${candidates.length} plausible document-sized fenced blocks were found -- pick the right one:</p>`;
      if (!noFencesFound) {
        candidates.forEach((c, i) => {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.textContent = `Use block ${i + 1} (${c.content.length} chars)`;
          btn.addEventListener("click", () => runValidation(c.content));
          pickerEl.appendChild(btn);
        });
      } else {
        runValidation(pastedText);
      }
    }

    function runValidation(responseText) {
      pickerEl.innerHTML = "";
      appState.submitResponse(responseText);
      if (ctx.chunkMode) {
        runChunkValidation(responseText);
        return;
      }
      const result = validate({ responseMarkdown: responseText, exportedMarkdown: ctx.exported.markdown, sourceMap: ctx.exported.sourceMap });
      if (result.ok) appState.validationPassed(result);
      else appState.validationFailed(result);
    }

    // M4c two-phase chunk validation (spec §6.4, architecture doc §7): phase 1
    // (validateText) runs against the CURRENT chunk's own slice only -- a failure here is
    // this chunk's own G1/G2 problem, reported and repaired exactly like the single-doc
    // path. Once phase 1 passes, this chunk's edits are translated into the full
    // document's coordinate space (chunk.js's translateEdits): if more chunks remain,
    // that's chunkAdvance's job (loop back to PROMPT_READY for the next chunk's prompt);
    // on the LAST chunk, every chunk's translated edits are merged and resolved ONCE
    // (resolveEdits, phase 2) against the full document's source map -- exactly the
    // single-document path's own G3/G4/G5 pass, just fed a merged edit list instead of one
    // parsed straight from a single response.
    function runChunkValidation(responseText) {
      const chunk = ctx.chunks[ctx.chunkIndex];
      const isLastChunk = ctx.chunkIndex === ctx.chunks.length - 1;

      const phase1 = validateText({
        responseMarkdown: responseText,
        exportedMarkdown: chunk.exportedMarkdown,
        sourceMap: chunk.sourceMap,
      });
      if (!phase1.ok) {
        appState.validationFailed(phase1);
        return;
      }
      const translated = translateEdits(phase1.edits, chunk.baseOffset);

      if (!isLastChunk) {
        const nextChunk = ctx.chunks[ctx.chunkIndex + 1];
        const nextBuilt = buildPrompt({
          persona: ctx.persona,
          exportedMarkdown: nextChunk.exportedMarkdown,
          filename: ctx.filename,
          chunk: { index: ctx.chunkIndex + 1, total: ctx.chunks.length },
        });
        appState.chunkAdvance({
          promptText: nextBuilt.text,
          promptVersion: nextBuilt.promptVersion,
          tokenEstimate: nextBuilt.tokenEstimate,
          translatedEdits: translated,
        });
        return;
      }

      const merged = [...ctx.chunkEdits, ...translated];
      const resolved = resolveEdits({ edits: merged, sourceMap: ctx.exported.sourceMap });
      if (resolved.ok) {
        appState.validationPassed(resolved);
      } else {
        // Rare by design (chunk boundaries stop edits spanning chunks) -- but the failing
        // edit may be from an EARLIER chunk's own response, not this (last) chunk's, so a
        // per-chunk repair prompt (which quotes ctx.response, only THIS chunk's text) would
        // point at the wrong text entirely. Tag it distinctly so the UI surfaces a global
        // restart notice instead of the ordinary repair flow.
        appState.validationFailed({ ...resolved, global: true });
      }
    }

    wrap.querySelector("#ar-validate").addEventListener("click", () => runEnvelope(responseEl.value));
    responseEl.addEventListener("paste", () => {
      // Let the paste land in the textarea first (spec §6.1: "Validator runs automatically
      // on paste"), then extract on the very next tick.
      setTimeout(() => runEnvelope(responseEl.value), 0);
    });
  }

  // M4c edge case: every chunk passed its own phase-1 (G1/G2) check, but the merged edit
  // list failed phase 2 (G3/G4) against the FULL document -- rare by design, since chunk
  // boundaries are chosen so an edit can never span two chunks. The failing edit may
  // originate from an EARLIER chunk's own response, not the one just pasted, so a
  // per-chunk repair prompt (composeRepair quotes ctx.response, which is only the LAST
  // chunk's text) could point entirely at the wrong text. Surface the gate + message and
  // let the human restart the review instead.
  function buildGlobalGateFailureEl(result) {
    const box = document.createElement("div");
    box.className = "ar-gate-failure";
    box.innerHTML = `
      <p><strong>${escapeHtml(result.gate)} failed after merging all chunks:</strong> ${escapeHtml(result.message)}</p>
      <p class="ar-hint">This is rare -- chunk boundaries are chosen so an edit can never span
      two chunks -- but it means the merged edit list doesn't resolve cleanly against the full
      document, and the failing edit may not even be from the chunk you just pasted. Rather
      than patch one chunk, start the review over.</p>
      <button type="button" id="ar-restart-review" class="ar-primary">Start a new review</button>
    `;
    box.querySelector("#ar-restart-review").addEventListener("click", () => appState.reset());
    return box;
  }

  // Mirrors the demo panel's failure view (first-divergence context is always cheap/O(n);
  // the full word-level diff is computed lazily, on demand, only if the human clicks the
  // button -- see validate.js's G2 comment for why an eager diffWords() call is unsafe).
  function buildGateFailureEl(result, ctx) {
    const box = document.createElement("div");
    box.className = "ar-gate-failure";
    const title = document.createElement("p");
    title.innerHTML = `<strong>${result.gate} failed:</strong> ${escapeHtml(result.message)}`;
    box.appendChild(title);

    if (result.gate === "G2" && result.firstDivergence) {
      const fd = result.firstDivergence;
      const fdEl = document.createElement("div");
      fdEl.className = "ar-first-divergence";
      fdEl.innerHTML =
        `<p>First divergence at offset ${fd.offset}:</p>` +
        `<pre>${fd.truncatedBefore ? "…" : ""}${escapeHtml(fd.before)}<mark>${escapeHtml(fd.afterA)}</mark>${fd.truncatedAfterA ? "…" : ""}\n` +
        `vs.\n` +
        `${fd.truncatedBefore ? "…" : ""}${escapeHtml(fd.before)}<mark>${escapeHtml(fd.afterB)}</mark>${fd.truncatedAfterB ? "…" : ""}</pre>`;
      box.appendChild(fdEl);
    }

    if (result.diffInputs) {
      const showDiffBtn = document.createElement("button");
      showDiffBtn.type = "button";
      showDiffBtn.textContent = "Show full diff";
      showDiffBtn.addEventListener("click", () => {
        const segments = diffWords(result.diffInputs.a, result.diffInputs.b);
        const diffEl = document.createElement("div");
        if (segments) {
          renderDiff(diffEl, segments);
        } else {
          diffEl.textContent = "The divergence is too large/scattered to diff in full -- see the first-divergence context above.";
        }
        box.replaceChild(diffEl, showDiffBtn);
      });
      box.appendChild(showDiffBtn);
    }

    // Mistake-specific repair prompt (M4b binding ruling: named rule + quoted divergence +
    // corrected pattern, not validate.js's generic "re-emit exactly" repairPrompt).
    const attemptCount = ctx.repairAttempts[result.gate] || 0;
    const repairText = composeRepair(result, ctx.response, attemptCount);
    const repairEl = document.createElement("div");
    repairEl.className = "ar-repair";
    const repairPre = document.createElement("pre");
    repairPre.textContent = repairText;
    repairEl.appendChild(repairPre);
    const copyBtn = document.createElement("button");
    copyBtn.type = "button";
    copyBtn.textContent = "Copy repair prompt";
    copyBtn.addEventListener("click", (e) => {
      if (typeof copyWithFeedback === "function") copyWithFeedback(e.currentTarget, repairText);
    });
    repairEl.appendChild(copyBtn);
    box.appendChild(repairEl);

    return box;
  }

  function renderRatifyStep(container, ctx) {
    const wrap = document.createElement("div");
    container.appendChild(wrap);

    if (ctx.validation.warnings && ctx.validation.warnings.length) {
      const warnBox = document.createElement("div");
      warnBox.className = "ar-warnings";
      const heading = document.createElement("p");
      heading.innerHTML = "<strong>G5 sanity warnings</strong> (do not block):";
      warnBox.appendChild(heading);
      const list = document.createElement("ul");
      for (const w of ctx.validation.warnings) {
        const li = document.createElement("li");
        li.textContent = w.message;
        list.appendChild(li);
      }
      warnBox.appendChild(list);
      wrap.appendChild(warnBox);
    }

    const authorField = document.createElement("div");
    authorField.className = "ar-field";
    const personaName = ctx.persona ? ctx.persona.name : "Default Persona";
    authorField.innerHTML = `
      <label for="ar-author">Author (tracked-change attribution)</label>
      <input type="text" id="ar-author" class="ar-author" value="AutoReviewer — ${escapeHtml(personaName)}" />
    `;
    wrap.appendChild(authorField);
    const authorEl = authorField.querySelector("#ar-author");

    const ratifyContainer = document.createElement("div");
    wrap.appendChild(ratifyContainer);
    const state = createRatificationState(ctx.validation.edits);
    currentRatifyState = state; // read by handleDownloadSession
    // Resume flow (M4b): loadSession() staged the saved decisions on ctx.pendingDecisions
    // rather than the (freshly re-derived) ratify state -- re-apply them here, once, by id.
    if (ctx.pendingDecisions && ctx.pendingDecisions.length) {
      for (const d of ctx.pendingDecisions) {
        state.setDecision(d.id, d.decision);
        if (d.reviewed) state.markReviewed(d.id);
      }
    }
    renderRatificationUI(ratifyContainer, state, { sourceText: ctx.exported.markdown });

    const statusEl = document.createElement("p");
    statusEl.className = "ar-hint";
    wrap.appendChild(statusEl);

    const injectBtn = ratifyContainer.querySelector('[data-action="inject"]');
    if (injectBtn) {
      injectBtn.addEventListener("click", async () => {
        injectBtn.disabled = true;
        statusEl.textContent = "Injecting accepted edits...";
        const author = authorEl.value.trim() || "AutoReviewer";
        let rewritten;
        try {
          const acceptedEdits = state.acceptedEdits();
          rewritten = await buildReviewedDocx({
            docxBytes: ctx.docxBytes,
            acceptedEdits,
            sourceMap: ctx.exported.sourceMap,
            author,
            date: new Date().toISOString(),
          });
          downloadBlob(rewritten, `${ctx.filename} — reviewed.docx`);
          appState.inject(); // stamps ctx.timestamps.injected, read by buildAuditRecord below
        } catch (err) {
          statusEl.textContent = `Injection failed: ${err.message}`;
          injectBtn.disabled = false;
          return;
        }

        // Audit sidecar (spec §12): assembled here, after the docx is actually serialized
        // and written, so output.sha256 hashes the real written bytes. A failure here is
        // reported separately -- the reviewed docx above already downloaded successfully.
        try {
          const auditRecord = await buildAuditRecord({
            promptVersion: ctx.promptVersion,
            timestamps: ctx.timestamps,
            filename: ctx.filename,
            docxBytes: ctx.docxBytes,
            outputBytes: rewritten,
            response: ctx.response,
            sourceMap: ctx.exported.sourceMap,
            persona: ctx.persona,
            validationAttempts: ctx.validationAttempts,
            rows: state.rows,
            author,
          });
          downloadJson(auditRecord, `${ctx.filename} — review-audit.json`);
        } catch (err) {
          statusEl.textContent = `Reviewed document downloaded, but the audit sidecar failed to build: ${err.message}`;
        }
      });
    }
  }

  function renderInjectedStep(container) {
    const wrap = document.createElement("div");
    wrap.innerHTML = `
      <p class="ar-hint">Reviewed document downloaded. Open it in Word to check the tracked changes and comments.</p>
      <button type="button" id="ar-new-review" class="ar-primary">Start a new review</button>
    `;
    container.appendChild(wrap);
    wrap.querySelector("#ar-new-review").addEventListener("click", () => appState.reset());
  }

  appState.onChange(render);
  render();
}

function init() {
  const root = document.getElementById("ar-app");
  if (!root) return;
  renderShell(root);
}

if (typeof document !== "undefined") {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
}

export { renderShell, selectFlow };
