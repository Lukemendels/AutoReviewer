import { exportDocx } from "../ooxml/export.js";
import { validate } from "../validate.js";
import { buildPrompt } from "../prompt.js";
import { parsePersona } from "../persona.js";
import { buildAuditRecord } from "../audit.js";
import { saveSession } from "../session.js";
import { renderRatificationUI } from "./ratify.js";
import { diffWords } from "./diff.js";

const FLOWS = [
  { id: "run-review", label: "Run Review" },
  { id: "respond-review", label: "Respond to Review" },
  { id: "train-persona", label: "Train Persona" },
];

// Referenced by each flow's not-yet-implemented placeholder so the module graph
// this milestone builds against is proven end-to-end by the M0 build/bundle.
const NOT_YET_WIRED = {
  exportDocx,
  validate,
  buildPrompt,
  parsePersona,
  buildAuditRecord,
  saveSession,
  renderRatificationUI,
  diffWords,
};

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
    panel.innerHTML = `<p class="ar-coming-soon">${flow.label} is not implemented yet.</p>`;
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
