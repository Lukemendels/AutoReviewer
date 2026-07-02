import { exportDocx } from "../ooxml/export.js";
import { validate } from "../validate.js";
import { buildPrompt } from "../prompt.js";
import { parsePersona } from "../persona.js";
import { buildAuditRecord } from "../audit.js";
import { saveSession } from "../session.js";
import { createRatificationState, renderRatificationUI } from "./ratify.js";
import { DEMO_DOCX_BASE64 } from "./demo-doc.js";

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

// The Run Review flow's file-picker/persona/prompt wiring is M4's job (spec §6). This
// panel demonstrates the part M2 actually builds -- paste a response, run G1-G5, ratify --
// against an embedded demo document so it's usable standalone right now.
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
    <div class="ar-controls">
      <button type="button" id="ar-validate" class="ar-primary">Validate</button>
      <span id="ar-export-status" class="ar-hint"></span>
    </div>
    <div id="ar-results"></div>
  `;

  const responseEl = panel.querySelector("#ar-response");
  const statusEl = panel.querySelector("#ar-export-status");
  const resultsEl = panel.querySelector("#ar-results");
  const validateBtn = panel.querySelector("#ar-validate");

  let exported = null; // { markdown, sourceMap }
  validateBtn.disabled = true;
  statusEl.textContent = "Loading demo document...";

  exportDocx(base64ToArrayBuffer(DEMO_DOCX_BASE64), { filename: "plain-paragraphs" })
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
    renderValidationResult(resultsEl, result, exported);
  });
}

function renderValidationResult(container, result, exported) {
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
      copyBtn.addEventListener("click", () => {
        if (typeof copyText === "function") copyText(result.repairPrompt);
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
