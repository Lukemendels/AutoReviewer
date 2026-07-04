import { exportDocx } from "../ooxml/export.js";
import { validate } from "../validate.js";
import { buildPrompt } from "../prompt.js";
import { parsePersona } from "../persona.js";
import { buildAuditRecord } from "../audit.js";
import { saveSession } from "../session.js";
import { createRatificationState, renderRatificationUI } from "./ratify.js";
import { DEMO_DOCX_BASE64 } from "./demo-doc.js";
import { unzip, readEntry } from "../zip/reader.js";
import { writeZip } from "../zip/writer.js";
import { parseXml } from "../ooxml/parse.js";
import { injectEdits } from "../ooxml/inject.js";
import { upsertComments } from "../ooxml/comments.js";
import { serializePart } from "../ooxml/serialize.js";

const FLOWS = [
  { id: "run-review", label: "Run Review" },
  { id: "respond-review", label: "Respond to Review" },
  { id: "train-persona", label: "Train Persona" },
];

// Not yet wired into any flow's UI (land in M4/M5/M6).
const NOT_YET_WIRED = { buildPrompt, parsePersona, buildAuditRecord, saveSession };

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

// The Run Review flow's file-picker/persona/prompt wiring is M4's job (spec §6). This
// panel demonstrates the part M2/M3b actually build -- paste a response, run G1-G5,
// ratify, inject -- against an embedded demo document so it's usable standalone right now.
function renderRunReviewPanel(panel) {
  panel.innerHTML = `
    <p class="ar-hint">
      Demo document: <code>plain-paragraphs.docx</code> (embedded). Paste an LLM's
      CriticMarkup response below and click Validate. The exported markdown is
      pre-filled as a starting point -- edit it to try a valid response, or paste your
      own to test a failure.
    </p>
    <div class="ar-field">
      <label for="ar-response">Response</label>
      <textarea id="ar-response" class="ar-response" spellcheck="false"></textarea>
    </div>
    <div class="ar-field">
      <label for="ar-author">Author (tracked-change attribution)</label>
      <input type="text" id="ar-author" class="ar-author" value="AutoReviewer — Demo Persona" />
    </div>
    <div class="ar-controls">
      <button type="button" id="ar-validate" class="ar-primary">Validate</button>
      <span id="ar-export-status" class="ar-hint"></span>
    </div>
    <div id="ar-results"></div>
  `;

  const responseEl = panel.querySelector("#ar-response");
  const authorEl = panel.querySelector("#ar-author");
  const statusEl = panel.querySelector("#ar-export-status");
  const resultsEl = panel.querySelector("#ar-results");
  const validateBtn = panel.querySelector("#ar-validate");

  let exported = null; // { markdown, sourceMap }
  let docxBytes = null;
  validateBtn.disabled = true;
  statusEl.textContent = "Loading demo document...";

  docxBytes = base64ToArrayBuffer(DEMO_DOCX_BASE64);
  exportDocx(docxBytes, { filename: "plain-paragraphs" })
    .then((result) => {
      exported = result;
      responseEl.value = result.markdown;
      statusEl.textContent = "Demo document loaded.";
      validateBtn.disabled = false;
    })
    .catch((err) => {
      statusEl.textContent = `Failed to load demo document: ${err.message}`;
    });

  validateBtn.addEventListener("click", () => {
    if (!exported) return;
    const result = validate({ responseMarkdown: responseEl.value, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    renderValidationResult(resultsEl, result, exported, {
      docxBytes,
      filename: "plain-paragraphs",
      getAuthor: () => authorEl.value.trim() || "AutoReviewer",
      statusEl,
    });
  });
}

function renderValidationResult(container, result, exported, injectCtx) {
  container.innerHTML = "";

  if (!result.ok) {
    const box = document.createElement("div");
    box.className = "ar-gate-failure";
    const title = document.createElement("p");
    title.innerHTML = `<strong>${result.gate} failed:</strong> ${result.message}`;
    box.appendChild(title);

    if (result.gate === "G2") {
      const diffEl = document.createElement("div");
      renderDiff(diffEl, result.diff);
      box.appendChild(diffEl);

      const copyBtn = document.createElement("button");
      copyBtn.type = "button";
      copyBtn.textContent = "Copy repair prompt";
      copyBtn.addEventListener("click", (e) => {
        if (typeof copyWithFeedback === "function") copyWithFeedback(e.currentTarget, result.repairPrompt);
      });
      box.appendChild(copyBtn);
    }

    container.appendChild(box);
    return;
  }

  if (result.warnings.length) {
    const warnBox = document.createElement("div");
    warnBox.className = "ar-warnings";
    const heading = document.createElement("p");
    heading.innerHTML = "<strong>G5 sanity warnings</strong> (do not block):";
    warnBox.appendChild(heading);
    const list = document.createElement("ul");
    for (const w of result.warnings) {
      const li = document.createElement("li");
      li.textContent = w.message;
      list.appendChild(li);
    }
    warnBox.appendChild(list);
    container.appendChild(warnBox);
  }

  const ratifyContainer = document.createElement("div");
  container.appendChild(ratifyContainer);
  const state = createRatificationState(result.edits);
  renderRatificationUI(ratifyContainer, state, { sourceText: exported.markdown });

  // ratify.js renders the Inject button itself (data-action="inject") but deliberately
  // leaves it unwired -- injection is business logic the caller owns. Per the plan's
  // wiring section: on click, re-unzip -> injectEdits -> upsertComments (if new comments)
  // -> writeZip -> Blob download named "{original name} — reviewed.docx".
  const injectBtn = ratifyContainer.querySelector('[data-action="inject"]');
  if (injectBtn && injectCtx) {
    injectBtn.addEventListener("click", async () => {
      injectBtn.disabled = true;
      injectCtx.statusEl.textContent = "Injecting accepted edits...";
      try {
        const acceptedEdits = state.acceptedEdits();
        const rewritten = await buildReviewedDocx({
          docxBytes: injectCtx.docxBytes,
          acceptedEdits,
          sourceMap: exported.sourceMap,
          author: injectCtx.getAuthor(),
          date: new Date().toISOString(),
        });
        const downloadName = `${injectCtx.filename} — reviewed.docx`;
        downloadBlob(rewritten, downloadName);
        injectCtx.statusEl.textContent = `Injected ${acceptedEdits.length} accepted edit(s) and downloaded "${downloadName}".`;
      } catch (err) {
        injectCtx.statusEl.textContent = `Injection failed: ${err.message}`;
      } finally {
        injectBtn.disabled = false;
      }
    });
  }
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

export { NOT_YET_WIRED, renderShell, selectFlow };
