import { validate } from "../validate.js";
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

  function regeneratePrompt(persona) {
    const ctx = appState.context;
    const built = buildPrompt({ persona, exportedMarkdown: ctx.exported.markdown, filename: ctx.filename });
    appState.setPrompt(built);
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

    if (ctx.overThreshold) {
      wrap.innerHTML = `
        <div class="ar-gate-failure">
          <p><strong>This document is over the single-prompt word threshold.</strong>
          Chunk mode required &mdash; coming in M4c. A full round trip on a document this
          long is known to fail the fidelity gate (G2), so the prompt below is withheld
          rather than handing you one known not to work.</p>
        </div>
      `;
      container.appendChild(wrap);
      return;
    }

    wrap.innerHTML = `
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
      wrap.appendChild(buildGateFailureEl(ctx.validation, ctx));
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
      const { candidates, noFencesFound } = extractCandidates(pastedText, { exportedLength: ctx.exported.markdown.length });
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
      const result = validate({ responseMarkdown: responseText, exportedMarkdown: ctx.exported.markdown, sourceMap: ctx.exported.sourceMap });
      if (result.ok) appState.validationPassed(result);
      else appState.validationFailed(result);
    }

    wrap.querySelector("#ar-validate").addEventListener("click", () => runEnvelope(responseEl.value));
    responseEl.addEventListener("paste", () => {
      // Let the paste land in the textarea first (spec §6.1: "Validator runs automatically
      // on paste"), then extract on the very next tick.
      setTimeout(() => runEnvelope(responseEl.value), 0);
    });
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
