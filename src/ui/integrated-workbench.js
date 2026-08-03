import { compileMarkdownDiff, deriveProtectedPrefixLength, MarkdownDiffError } from "../markdown-diff.js";
import { extractCandidates } from "../envelope.js";
import { exportDocx } from "../ooxml/export.js";
import { clusterPasses } from "../passes.js";
import { renderSlice } from "../ooxml/slice.js";
import { strip } from "../criticmarkup/strip.js";
import { parsePersona } from "../persona.js";
import { createDecisionState, renderDecisionPane, validateDecisionSpec } from "../a2ui/decision.js";

const FLOW_LABELS = {
  "run-review": "Review a Draft",
  "respond-review": "Incorporate Feedback",
  "train-persona": "Build a Persona",
};

function escapeHtml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function baseName(filename) {
  return (filename || "document").replace(/\.docx$/i, "");
}

async function readBytes(file) {
  if (typeof file.arrayBuffer === "function") return file.arrayBuffer();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function downloadText(text, filename, type = "text/plain") {
  const blob = new Blob([text], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function legacyCopy(text) {
  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.style.position = "fixed";
  textarea.style.opacity = "0";
  document.body.appendChild(textarea);
  textarea.focus();
  textarea.select();
  let ok = false;
  try {
    ok = document.execCommand("copy");
  } catch {
    ok = false;
  }
  textarea.remove();
  return ok;
}

async function copyText(text) {
  if (navigator.clipboard && window.isSecureContext) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      return legacyCopy(text);
    }
  }
  return legacyCopy(text);
}

function installStyles() {
  if (document.getElementById("ar-integrated-styles")) return;
  const style = document.createElement("style");
  style.id = "ar-integrated-styles";
  style.textContent = `
    .ar-task-home{border:1px solid #c9d7e5;background:#f7fbff;border-radius:12px;padding:1rem;margin:0 0 1rem}
    .ar-task-home h2{margin:.1rem 0 .75rem}.ar-task-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:.75rem}
    .ar-task-card{display:flex;flex-direction:column;align-items:flex-start;text-align:left;border:1px solid #b8cce0;background:#fff;color:#17324d;border-radius:10px;padding:1rem;min-height:120px}
    .ar-task-card:hover{border-color:#005ea8}.ar-task-card strong{font-size:1.05rem}.ar-task-card span{margin-top:.35rem;color:#465b70;font-weight:400}
    .ar-integrated-status{padding:.7rem;border-radius:6px;background:#eef6ff;margin:.75rem 0}.ar-integrated-error{background:#fde7e9;color:#8b0000}
    .ar-workflow-step{border:1px solid #d8d8d8;border-radius:10px;padding:1rem;margin:1rem 0;background:#fff}.ar-workflow-step h2,.ar-workflow-step h3{margin-top:0}
    .ar-a2ui-theme{border:1px solid #cfd8e3;border-radius:10px;padding:1rem;margin:1rem 0}.ar-a2ui-stakes{background:#f7f7f7;padding:.65rem;border-radius:6px}
    .ar-a2ui-options{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:.75rem}.ar-a2ui-option{display:grid;grid-template-columns:auto 1fr;gap:.65rem;border:1px solid #c9c9c9;border-radius:8px;padding:.75rem;cursor:pointer}
    .ar-a2ui-option:has(input:checked){border-color:#005ea8;box-shadow:0 0 0 2px rgba(0,94,168,.15)}.ar-a2ui-option p{margin:.35rem 0}
    .ar-a2ui-tradeoffs{display:grid;grid-template-columns:1fr 1fr;gap:.75rem}.ar-a2ui-tradeoffs ul{margin:.35rem 0;padding-left:1.2rem}
    .ar-a2ui-note,.ar-a2ui-other{width:100%;min-height:70px;margin-top:.65rem}.ar-pass-table{width:100%;border-collapse:collapse}.ar-pass-table th,.ar-pass-table td{border-bottom:1px solid #ddd;padding:.5rem;text-align:left}
    .ar-corpus-list{max-height:420px;overflow:auto}.ar-json-output{width:100%;min-height:260px;font-family:Consolas,monospace}.ar-inline-controls{display:flex;gap:.5rem;flex-wrap:wrap;align-items:center}
    @media(max-width:850px){.ar-task-grid,.ar-a2ui-options{grid-template-columns:1fr}}
  `;
  document.head.appendChild(style);
}

function installTaskHome(root) {
  if (root.querySelector(".ar-task-home")) return;
  const home = document.createElement("section");
  home.className = "ar-task-home";
  home.innerHTML = `
    <h2>What are you trying to accomplish?</h2>
    <div class="ar-task-grid">
      <button type="button" class="ar-task-card" data-flow="run-review"><strong>Review a Draft</strong><span>Apply a reviewer persona and return human-ratified tracked changes.</span></button>
      <button type="button" class="ar-task-card" data-flow="respond-review"><strong>Incorporate Feedback</strong><span>Turn redlines and comments into bounded choices and a revised draft.</span></button>
      <button type="button" class="ar-task-card" data-flow="train-persona"><strong>Build a Persona</strong><span>Mine selected reviewer passes into a portable review profile.</span></button>
    </div>`;
  const nav = root.querySelector(".ar-tabs");
  root.insertBefore(home, nav || root.firstChild);
  home.addEventListener("click", (event) => {
    const button = event.target.closest("[data-flow]");
    if (!button) return;
    root.querySelector(`.ar-tab[data-flow="${button.dataset.flow}"]`)?.click();
  });
}

function relabelFlows(root) {
  for (const [id, label] of Object.entries(FLOW_LABELS)) {
    const tab = root.querySelector(`.ar-tab[data-flow="${id}"]`);
    if (tab) tab.textContent = label;
  }
}

function transformReviewPrompt(prompt) {
  if (!prompt || !prompt.includes("[CRITICMARKUP RULES]") || !prompt.includes("[DOCUMENT]")) return prompt;
  const taskAt = prompt.indexOf("[TASK]");
  const documentAt = prompt.indexOf("[DOCUMENT]");
  if (taskAt === -1 || documentAt === -1) return prompt;
  const persona = prompt.slice(0, taskAt).trimEnd();
  const documentSection = prompt.slice(documentAt);
  const chunkMatch = prompt.match(/This is part (\d+) of (\d+) of a larger document\./);
  const chunk = chunkMatch
    ? `This is part ${chunkMatch[1]} of ${chunkMatch[2]} of a larger document. Revise only this part and return the complete revised part.\n\n`
    : "";
  return [
    persona,
    `[TASK]\n${chunk}Review the document according to the persona. Return a clean revised Markdown document. Perform semantic revision only; do not write CriticMarkup, JSON, explanations, or a review memo.`,
    `[OUTPUT CONTRACT]\n- Return the complete revised document inside exactly one fenced \`\`\`markdown block.\n- Preserve every unchanged character and structural marker that should remain in the document.\n- The browser will derive tracked changes deterministically from the baseline and revised versions.`,
    `[HARD CONSTRAINTS]\n- Preserve the leading AutoReviewer HTML-comment preface byte-for-byte.\n- Preserve text inside ⟦…⟧ and [image: …] exactly.\n- Do not remove or rewrite locked placeholders.\n- Do not add commentary before or after the fenced document.`,
    documentSection,
  ].join("\n\n");
}

function extractDocumentFromPrompt(prompt) {
  const marker = "[DOCUMENT]";
  const at = prompt.indexOf(marker);
  if (at === -1) return null;
  const after = prompt.slice(at + marker.length);
  const open = after.indexOf("```markdown");
  if (open === -1) return null;
  const start = open + "```markdown".length;
  const bodyStart = after[start] === "\n" ? start + 1 : start;
  const close = after.lastIndexOf("```");
  if (close === -1 || close < bodyStart) return null;
  let body = after.slice(bodyStart, close);
  // buildDocumentSection adds one framing newline before the closing fence. Remove only
  // that framing newline; preserve the document's own final newline when it has one.
  if (body.endsWith("\n")) body = body.slice(0, -1);
  return body;
}

function visibleRunPrompt() {
  const panel = document.querySelector('.ar-panel[data-flow="run-review"]');
  return panel?.querySelector(".ar-prompt-text")?.value || null;
}

function transformVisibleRunPrompt() {
  const panel = document.querySelector('.ar-panel[data-flow="run-review"]');
  if (!panel) return;
  for (const textarea of panel.querySelectorAll(".ar-prompt-text")) {
    if (textarea.dataset.cleanRevisionPrompt === "1") continue;
    const transformed = transformReviewPrompt(textarea.value);
    if (transformed !== textarea.value) {
      textarea.value = transformed;
      textarea.dataset.cleanRevisionPrompt = "1";
    }
  }
  const responseLabel = panel.querySelector('label[for="ar-response"]');
  if (responseLabel) responseLabel.textContent = "Revised Markdown response";
}

function revisedCandidate(text, baselineLength) {
  const { candidates, noFencesFound } = extractCandidates(text, { exportedLength: baselineLength });
  if (candidates.length === 1 && !noFencesFound) return candidates[0].content;
  if (noFencesFound || candidates.length === 0) return text;
  throw new MarkdownDiffError("Multiple plausible fenced documents were found. Keep only the revised Markdown block and try again.");
}

function showRunAdapterStatus(message, error = false) {
  const panel = document.querySelector('.ar-panel[data-flow="run-review"]');
  if (!panel) return;
  let status = panel.querySelector(".ar-clean-revision-status");
  if (!status) {
    status = document.createElement("div");
    status.className = "ar-integrated-status ar-clean-revision-status";
    const responseStep = panel.querySelector(".ar-response-step");
    (responseStep || panel).prepend(status);
  }
  status.classList.toggle("ar-integrated-error", error);
  status.textContent = message;
}

function compileRunResponse(textarea) {
  if (!textarea || textarea.dataset.compiledCriticMarkup === "1") return true;
  const prompt = visibleRunPrompt();
  const baseline = prompt ? extractDocumentFromPrompt(prompt) : null;
  if (baseline == null) {
    showRunAdapterStatus("Could not recover the authoritative Markdown from the active prompt.", true);
    return false;
  }
  try {
    const revised = revisedCandidate(textarea.value, baseline.length);
    const result = compileMarkdownDiff(baseline, revised, {
      protectedPrefixLength: deriveProtectedPrefixLength(baseline),
    });
    textarea.dataset.cleanRevisionMarkdown = revised;
    textarea.value = result.criticMarkup;
    textarea.dataset.compiledCriticMarkup = "1";
    showRunAdapterStatus(`Compiled revised Markdown into ${result.editGroups} deterministic change group${result.editGroups === 1 ? "" : "s"}.`);
    return true;
  } catch (error) {
    showRunAdapterStatus(error.message || "Could not compile the revised Markdown.", true);
    return false;
  }
}

function installRunReviewAdapter() {
  if (window.__AR_CLEAN_REVIEW_ADAPTER__) return;
  window.__AR_CLEAN_REVIEW_ADAPTER__ = true;

  const originalCopy = window.copyWithFeedback;
  if (typeof originalCopy === "function") {
    window.copyWithFeedback = function wrappedCopyWithFeedback(button, text) {
      return originalCopy(button, transformReviewPrompt(text));
    };
  }

  document.addEventListener(
    "paste",
    (event) => {
      const textarea = event.target.closest?.('.ar-panel[data-flow="run-review"] #ar-response');
      if (!textarea) return;
      event.stopImmediatePropagation();
      textarea.dataset.compiledCriticMarkup = "";
      setTimeout(() => {
        if (compileRunResponse(textarea)) {
          textarea.closest(".ar-response-step")?.querySelector("#ar-validate")?.click();
        }
      }, 0);
    },
    true
  );

  document.addEventListener(
    "click",
    (event) => {
      const button = event.target.closest?.('.ar-panel[data-flow="run-review"] #ar-validate');
      if (!button) return;
      const textarea = button.closest(".ar-response-step")?.querySelector("#ar-response");
      if (!compileRunResponse(textarea)) event.stopImmediatePropagation();
    },
    true
  );

  document.addEventListener(
    "input",
    (event) => {
      const textarea = event.target.closest?.('.ar-panel[data-flow="run-review"] #ar-response');
      if (textarea) textarea.dataset.compiledCriticMarkup = "";
    },
    true
  );
}

function feedbackTable(observations) {
  let revision = 0;
  let comment = 0;
  return [...observations]
    .sort((a, b) => a.docOrder - b.docOrder)
    .map((observation) => {
      const isComment = observation.kind === "comment" || observation.kind === "comment-reply";
      const feedbackId = isComment ? `C${++comment}` : `R${++revision}`;
      return { ...observation, feedbackId };
    });
}

function buildDecisionAnalysisPrompt(session) {
  const feedback = session.feedback.map((item) => ({
    id: item.feedbackId,
    kind: item.kind,
    author: item.author,
    date: item.date,
    anchorText: item.anchorText,
    text: item.text,
  }));
  return `You are the decision-framing layer for AutoReviewer. Analyze the review feedback below and compress it into major edit-theme decisions.\n\nReturn JSON only, inside one fenced \`\`\`json block, using this exact schema:\n{\n  "schema": "autoreviewer.a2ui.decision/v1",\n  "documentHash": ${JSON.stringify(session.documentHash)},\n  "title": "...",\n  "themes": [{\n    "id": "theme-01",\n    "title": "...",\n    "summary": "...",\n    "feedbackIds": ["C1", "R1"],\n    "stakes": "...",\n    "options": [{\n      "id": "...",\n      "label": "...",\n      "description": "...",\n      "pros": ["..."],\n      "cons": ["..."]\n    }]\n  }]\n}\n\nRules:\n- Create the smallest useful set of major themes.\n- Every feedback ID must appear exactly once.\n- Give each theme three or four genuinely distinct options.\n- Include concrete pros and cons for every option.\n- Do not add an Other option; the browser adds it deterministically.\n- Do not decide for the user.\n\n[FEEDBACK OBSERVATIONS]\n${JSON.stringify(feedback, null, 2)}\n\n[ANNOTATED DOCUMENT]\n\`\`\`markdown\n${session.exported.markdown}\n\`\`\``;
}

function buildIncorporationPrompt(session, payload) {
  return `You are revising a document after human decisions on clustered review feedback.\n\nReturn the complete clean revised Markdown inside exactly one fenced \`\`\`markdown block. Do not use CriticMarkup and do not explain your work.\n\nThe human decisions are authoritative. Implement them faithfully, preserve locked placeholders, and resolve local wording in the most coherent way consistent with those decisions.\n\n[HUMAN DECISIONS]\n${JSON.stringify(payload, null, 2)}\n\n[FEEDBACK EVIDENCE]\n${JSON.stringify(session.feedback, null, 2)}\n\n[PRE-REVIEW BASELINE]\n\`\`\`markdown\n${session.baselineMarkdown}\n\`\`\`\n\n[ANNOTATED REVIEW VIEW]\n\`\`\`markdown\n${session.exported.markdown}\n\`\`\``;
}

function parseJsonEnvelope(text) {
  const matches = [...String(text).matchAll(/```(?:json)?\s*\n([\s\S]*?)```/gi)];
  const candidate = matches.length ? matches[matches.length - 1][1] : text;
  return JSON.parse(candidate.trim());
}

function mountRespondWorkflow(panel) {
  if (!panel || panel.dataset.integratedOwned === "1") return;
  panel.dataset.integratedOwned = "1";
  const session = { file: null, bytes: null, exported: null, feedback: [], documentHash: null, baselineMarkdown: null, decisionSpec: null, decisionPayload: null };

  function renderLoad() {
    panel.innerHTML = `<section class="ar-workflow-step"><h2>Incorporate Feedback</h2><p>Load a reviewed Word document. AutoReviewer will extract its comments and redlines, ask the model to organize them into major decisions, and render those choices through a trusted A2UI interface.</p><input id="ar-incorporate-file" type="file" accept=".docx"><div id="ar-incorporate-status"></div></section>`;
    panel.querySelector("#ar-incorporate-file").addEventListener("change", async (event) => {
      const file = event.target.files?.[0];
      if (!file) return;
      const status = panel.querySelector("#ar-incorporate-status");
      status.className = "ar-integrated-status";
      status.textContent = `Reading ${file.name}…`;
      try {
        const bytes = await readBytes(file);
        const filename = baseName(file.name);
        const exported = await exportDocx(bytes, { DOMParserImpl: DOMParser, annotate: true, filename, collectObservations: true });
        const prefix = exported.sourceMap?.blocks?.length ? exported.sourceMap.blocks[0].mdStart : deriveProtectedPrefixLength(exported.markdown);
        session.file = file;
        session.bytes = bytes;
        session.exported = exported;
        session.feedback = feedbackTable(exported.observations || []);
        session.documentHash = exported.sourceMap?.docHash || `${filename}:${bytes.byteLength}`;
        session.baselineMarkdown = strip(exported.markdown, { skipBefore: prefix });
        renderAnalysisPrompt();
      } catch (error) {
        status.classList.add("ar-integrated-error");
        status.textContent = error.message;
      }
    });
  }

  function renderAnalysisPrompt() {
    const prompt = buildDecisionAnalysisPrompt(session);
    panel.innerHTML = `<section class="ar-workflow-step"><h2>1. Frame the decisions</h2><p>${session.feedback.length} feedback observations were extracted. Copy this packet into a model, then paste the returned A2UI JSON below.</p><textarea id="ar-decision-prompt" class="ar-json-output" readonly>${escapeHtml(prompt)}</textarea><div class="ar-controls"><button id="ar-copy-decision-prompt" class="ar-primary">Copy decision packet</button></div><label for="ar-decision-response"><strong>A2UI decision specification</strong></label><textarea id="ar-decision-response" class="ar-json-output" placeholder="Paste the model's JSON response"></textarea><div class="ar-controls"><button id="ar-render-decisions" class="ar-primary">Render decisions</button><button id="ar-restart-incorporate">Start over</button></div><div id="ar-decision-error"></div></section>`;
    panel.querySelector("#ar-copy-decision-prompt").addEventListener("click", async (event) => {
      const ok = await copyText(prompt);
      event.currentTarget.textContent = ok ? "Copied ✓" : "Copy failed";
    });
    panel.querySelector("#ar-restart-incorporate").addEventListener("click", renderLoad);
    panel.querySelector("#ar-render-decisions").addEventListener("click", () => {
      const errorEl = panel.querySelector("#ar-decision-error");
      try {
        const raw = parseJsonEnvelope(panel.querySelector("#ar-decision-response").value);
        session.decisionSpec = validateDecisionSpec(raw, {
          documentHash: session.documentHash,
          feedbackIds: session.feedback.map((item) => item.feedbackId),
          requireFullCoverage: true,
        });
        renderDecisions();
      } catch (error) {
        errorEl.className = "ar-integrated-status ar-integrated-error";
        errorEl.textContent = error.message;
      }
    });
  }

  function renderDecisions() {
    panel.innerHTML = `<section class="ar-workflow-step"><div id="ar-decision-pane"></div><div id="ar-decision-next"></div></section>`;
    const feedbackById = new Map(session.feedback.map((item) => [item.feedbackId, item]));
    const state = createDecisionState(session.decisionSpec);
    renderDecisionPane(panel.querySelector("#ar-decision-pane"), session.decisionSpec, state, {
      feedbackById,
      onConfirm(payload) {
        session.decisionPayload = payload;
        renderRevisionPrompt();
      },
    });
  }

  function renderRevisionPrompt() {
    const prompt = buildIncorporationPrompt(session, session.decisionPayload);
    panel.innerHTML = `<section class="ar-workflow-step"><h2>2. Produce the revised draft</h2><p>The confirmed decision payload is now the authoritative instruction. Copy the packet, then paste the complete revised Markdown.</p><textarea id="ar-incorporation-prompt" class="ar-json-output" readonly>${escapeHtml(prompt)}</textarea><div class="ar-controls"><button id="ar-copy-incorporation" class="ar-primary">Copy revision packet</button><button id="ar-download-decisions">Download decisions.json</button></div><label for="ar-incorporated-markdown"><strong>Revised Markdown</strong></label><textarea id="ar-incorporated-markdown" class="ar-json-output"></textarea><div class="ar-controls"><button id="ar-compile-incorporation" class="ar-primary">Compile tracked-change diff</button></div><div id="ar-incorporation-error"></div></section>`;
    panel.querySelector("#ar-copy-incorporation").addEventListener("click", async (event) => {
      const ok = await copyText(prompt);
      event.currentTarget.textContent = ok ? "Copied ✓" : "Copy failed";
    });
    panel.querySelector("#ar-download-decisions").addEventListener("click", () => downloadText(JSON.stringify(session.decisionPayload, null, 2), `${baseName(session.file.name)}.decisions.json`, "application/json"));
    panel.querySelector("#ar-compile-incorporation").addEventListener("click", () => {
      try {
        const revised = revisedCandidate(panel.querySelector("#ar-incorporated-markdown").value, session.baselineMarkdown.length);
        const compiled = compileMarkdownDiff(session.baselineMarkdown, revised, { protectedPrefixLength: deriveProtectedPrefixLength(session.baselineMarkdown) });
        renderIncorporationResult(revised, compiled);
      } catch (error) {
        const target = panel.querySelector("#ar-incorporation-error");
        target.className = "ar-integrated-status ar-integrated-error";
        target.textContent = error.message;
      }
    });
  }

  function renderIncorporationResult(revised, compiled) {
    panel.innerHTML = `<section class="ar-workflow-step"><h2>3. Deterministic change artifact</h2><p>The decisions produced clean revised Markdown and a deterministic CriticMarkup diff. Word write-back requires the planned pre-review OOXML flattening boundary; this branch exposes the complete decision and revision artifacts without pretending that flattening already occurred.</p><textarea class="ar-json-output" readonly>${escapeHtml(compiled.criticMarkup)}</textarea><div class="ar-controls"><button id="ar-download-revised" class="ar-primary">Download revised Markdown</button><button id="ar-download-diff">Download CriticMarkup</button><button id="ar-new-incorporation">Start another document</button></div></section>`;
    panel.querySelector("#ar-download-revised").addEventListener("click", () => downloadText(revised, `${baseName(session.file.name)}.revised.md`, "text/markdown"));
    panel.querySelector("#ar-download-diff").addEventListener("click", () => downloadText(compiled.criticMarkup, `${baseName(session.file.name)}.criticmarkup.md`, "text/markdown"));
    panel.querySelector("#ar-new-incorporation").addEventListener("click", renderLoad);
  }

  renderLoad();
}

function buildPersonaPrompt(packet) {
  return `Synthesize a portable AutoReviewer persona from the selected reviewer-pass corpus.\n\nReturn Markdown only inside one fenced \`\`\`markdown block. The persona must contain recognizable sections for: Persona name; Role and voice; Review priorities; Structural preferences; Substantive standards; Common intervention categories; Comment conventions; Pushback patterns; Do-not-touch rules; Style exemplars; Limits and uncertainty.\n\nGround every claim in repeated evidence from the corpus. Distinguish strong recurring patterns from tentative patterns. Do not describe or evaluate the individual reviewer as a person; describe a reusable review function.\n\n[TRAINING CORPUS]\n${packet}`;
}

function mountPersonaWorkflow(panel) {
  if (!panel || panel.dataset.integratedOwned === "1") return;
  panel.dataset.integratedOwned = "1";
  const docs = [];
  const selected = new Set();

  function render() {
    panel.innerHTML = `<section class="ar-workflow-step"><h2>Build a Persona</h2><p>Add reviewed Word documents, select the reviewer passes that represent the target review function, and export one grounded training packet.</p><input id="ar-persona-docs" type="file" accept=".docx" multiple><div id="ar-persona-status"></div><div class="ar-corpus-list"><table class="ar-pass-table"><thead><tr><th>Use</th><th>Document</th><th>Reviewer</th><th>Pass</th><th>Edits</th><th>Comments</th><th>Replies</th><th></th></tr></thead><tbody id="ar-persona-pass-rows"></tbody></table></div><div class="ar-controls"><button id="ar-build-corpus" class="ar-primary" ${selected.size ? "" : "disabled"}>Build training packet (${selected.size})</button></div><div id="ar-persona-output"></div></section>`;
    const rows = panel.querySelector("#ar-persona-pass-rows");
    docs.forEach((doc, docIndex) => {
      doc.passes.forEach((pass, passIndex) => {
        const key = `${docIndex}:${passIndex}`;
        const tr = document.createElement("tr");
        tr.innerHTML = `<td><input type="checkbox" data-key="${key}" ${selected.has(key) ? "checked" : ""}></td><td>${escapeHtml(doc.filename)}</td><td>${escapeHtml(pass.author)}</td><td>${escapeHtml(pass.undated ? "undated" : pass.passDate)}</td><td>${pass.counts.insertions + pass.counts.deletions}</td><td>${pass.counts.comments}</td><td>${pass.counts.replies}</td><td><button type="button" data-preview="${key}">Preview</button></td>`;
        rows.appendChild(tr);
      });
    });
    rows.addEventListener("change", (event) => {
      const key = event.target.dataset.key;
      if (!key) return;
      if (event.target.checked) selected.add(key);
      else selected.delete(key);
      render();
    });
    rows.addEventListener("click", async (event) => {
      const key = event.target.dataset.preview;
      if (!key) return;
      const [docIndex, passIndex] = key.split(":").map(Number);
      const doc = docs[docIndex];
      const slice = await renderSlice(doc.bytes, doc.passes[passIndex], { DOMParserImpl: DOMParser, filename: baseName(doc.filename) });
      const output = panel.querySelector("#ar-persona-output");
      output.innerHTML = `<h3>Slice preview</h3><textarea class="ar-json-output" readonly>${escapeHtml(slice)}</textarea>`;
    });
    panel.querySelector("#ar-persona-docs").addEventListener("change", async (event) => {
      const status = panel.querySelector("#ar-persona-status");
      status.className = "ar-integrated-status";
      for (const file of [...(event.target.files || [])]) {
        try {
          status.textContent = `Reading ${file.name}…`;
          const bytes = await readBytes(file);
          const exported = await exportDocx(bytes, { DOMParserImpl: DOMParser, annotate: true, filename: baseName(file.name), collectObservations: true });
          const clustered = clusterPasses(exported.observations || []);
          docs.push({ filename: file.name, bytes, passes: clustered.passes, metadataStripped: clustered.metadataStripped });
        } catch (error) {
          status.classList.add("ar-integrated-error");
          status.textContent = `${file.name}: ${error.message}`;
          return;
        }
      }
      render();
    });
    panel.querySelector("#ar-build-corpus").addEventListener("click", buildCorpus);
  }

  async function buildCorpus() {
    const slices = [];
    for (const key of selected) {
      const [docIndex, passIndex] = key.split(":").map(Number);
      const doc = docs[docIndex];
      const pass = doc.passes[passIndex];
      const slice = await renderSlice(doc.bytes, pass, { DOMParserImpl: DOMParser, filename: baseName(doc.filename) });
      slices.push(`<!-- SOURCE: ${doc.filename}; PASS: ${pass.label} -->\n${slice}`);
    }
    const packet = slices.join("\n\n---\n\n");
    const prompt = buildPersonaPrompt(packet);
    const output = panel.querySelector("#ar-persona-output");
    output.innerHTML = `<h3>Persona synthesis packet</h3><textarea id="ar-persona-training-prompt" class="ar-json-output" readonly>${escapeHtml(prompt)}</textarea><div class="ar-controls"><button id="ar-copy-persona-prompt" class="ar-primary">Copy training packet</button><button id="ar-download-corpus">Download corpus.md</button></div><label for="ar-persona-response"><strong>Generated persona Markdown</strong></label><textarea id="ar-persona-response" class="ar-json-output"></textarea><div class="ar-controls"><button id="ar-validate-persona" class="ar-primary">Validate and download persona</button></div><div id="ar-persona-validation"></div>`;
    output.querySelector("#ar-copy-persona-prompt").addEventListener("click", async (event) => {
      const ok = await copyText(prompt);
      event.currentTarget.textContent = ok ? "Copied ✓" : "Copy failed";
    });
    output.querySelector("#ar-download-corpus").addEventListener("click", () => downloadText(packet, "autoreviewer-persona-corpus.md", "text/markdown"));
    output.querySelector("#ar-validate-persona").addEventListener("click", () => {
      const validation = output.querySelector("#ar-persona-validation");
      try {
        const raw = output.querySelector("#ar-persona-response").value;
        const matches = [...raw.matchAll(/```(?:markdown)?\s*\n([\s\S]*?)```/gi)];
        const personaText = (matches.length ? matches[matches.length - 1][1] : raw).trim();
        const persona = parsePersona(personaText, { filename: "generated-persona.md" });
        validation.className = "ar-integrated-status";
        validation.textContent = `Validated persona: ${persona.name}`;
        downloadText(personaText + "\n", `${persona.name.replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "autoreviewer-persona"}.md`, "text/markdown");
      } catch (error) {
        validation.className = "ar-integrated-status ar-integrated-error";
        validation.textContent = error.message;
      }
    });
  }

  render();
}

function mountIntegratedSurfaces() {
  const root = document.getElementById("ar-app");
  if (!root) return;
  installStyles();
  relabelFlows(root);
  installTaskHome(root);
  installRunReviewAdapter();
  transformVisibleRunPrompt();
  mountRespondWorkflow(root.querySelector('.ar-panel[data-flow="respond-review"]'));
  mountPersonaWorkflow(root.querySelector('.ar-panel[data-flow="train-persona"]'));

  const observer = new MutationObserver(() => {
    relabelFlows(root);
    transformVisibleRunPrompt();
  });
  observer.observe(root, { subtree: true, childList: true });
}

if (typeof document !== "undefined") {
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", mountIntegratedSurfaces);
  else mountIntegratedSurfaces();
}

export { buildDecisionAnalysisPrompt, buildIncorporationPrompt, buildPersonaPrompt, mountIntegratedSurfaces, transformReviewPrompt };
