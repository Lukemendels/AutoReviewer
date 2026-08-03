// Trusted A2UI decision layer for AutoReviewer. The model supplies data only;
// deterministic code owns validation, rendering, selection state, and export.

export const DECISION_SPEC_SCHEMA = "autoreviewer.a2ui.decision/v1";
export const DECISION_SELECTION_SCHEMA = "autoreviewer.a2ui.selection/v1";

export class DecisionSpecError extends Error {
  constructor(message, details = {}) {
    super(message);
    this.name = "DecisionSpecError";
    this.details = details;
  }
}

function nonEmptyString(value, label, max = 4000) {
  if (typeof value !== "string" || !value.trim()) throw new DecisionSpecError(`${label} must be non-empty text.`);
  if (value.length > max) throw new DecisionSpecError(`${label} exceeds ${max} characters.`);
  return value.trim();
}

function uniqueIds(items, label) {
  const seen = new Set();
  for (const item of items) {
    if (seen.has(item.id)) throw new DecisionSpecError(`Duplicate ${label} id: ${item.id}`);
    seen.add(item.id);
  }
}

export function validateDecisionSpec(spec, options = {}) {
  if (!spec || typeof spec !== "object" || Array.isArray(spec)) throw new DecisionSpecError("Decision specification must be an object.");
  if (spec.schema !== DECISION_SPEC_SCHEMA) throw new DecisionSpecError(`Unsupported decision schema: ${spec.schema || "missing"}.`);
  const expectedHash = options.documentHash;
  if (expectedHash && spec.documentHash !== expectedHash) throw new DecisionSpecError("Decision specification is bound to a different document.");
  if (!Array.isArray(spec.themes) || spec.themes.length === 0) throw new DecisionSpecError("Decision specification must contain at least one theme.");
  if (spec.themes.length > (options.maxThemes || 12)) throw new DecisionSpecError("Decision specification contains too many themes.");

  const knownFeedback = new Set(options.feedbackIds || []);
  const usedFeedback = new Map();
  uniqueIds(spec.themes, "theme");

  const normalizedThemes = spec.themes.map((theme, themeIndex) => {
    const id = nonEmptyString(theme.id, `themes[${themeIndex}].id`, 80);
    const title = nonEmptyString(theme.title, `themes[${themeIndex}].title`, 180);
    const summary = nonEmptyString(theme.summary, `themes[${themeIndex}].summary`, 1200);
    const stakes = nonEmptyString(theme.stakes, `themes[${themeIndex}].stakes`, 1200);
    if (!Array.isArray(theme.feedbackIds) || !theme.feedbackIds.length) throw new DecisionSpecError(`${id} must reference at least one feedback item.`);
    const feedbackIds = [...new Set(theme.feedbackIds.map((value) => nonEmptyString(value, `${id}.feedbackIds`, 80)))];
    for (const feedbackId of feedbackIds) {
      if (knownFeedback.size && !knownFeedback.has(feedbackId)) throw new DecisionSpecError(`${id} references unknown feedback item ${feedbackId}.`);
      if (usedFeedback.has(feedbackId)) throw new DecisionSpecError(`${feedbackId} appears in both ${usedFeedback.get(feedbackId)} and ${id}.`);
      usedFeedback.set(feedbackId, id);
    }
    if (!Array.isArray(theme.options) || theme.options.length < 3 || theme.options.length > 4) {
      throw new DecisionSpecError(`${id} must contain three or four model-supplied options.`);
    }
    uniqueIds(theme.options, `option in ${id}`);
    const optionsNormalized = theme.options.map((option, optionIndex) => {
      if (!Array.isArray(option.pros) || !option.pros.length || !Array.isArray(option.cons) || !option.cons.length) {
        throw new DecisionSpecError(`${id} option ${optionIndex + 1} must contain both pros and cons.`);
      }
      return {
        id: nonEmptyString(option.id, `${id}.options[${optionIndex}].id`, 80),
        label: nonEmptyString(option.label, `${id}.options[${optionIndex}].label`, 160),
        description: nonEmptyString(option.description, `${id}.options[${optionIndex}].description`, 1000),
        pros: option.pros.map((value, index) => nonEmptyString(value, `${id}.options[${optionIndex}].pros[${index}]`, 500)),
        cons: option.cons.map((value, index) => nonEmptyString(value, `${id}.options[${optionIndex}].cons[${index}]`, 500)),
      };
    });
    return { id, title, summary, stakes, feedbackIds, options: optionsNormalized };
  });

  if (options.requireFullCoverage !== false && knownFeedback.size) {
    const omitted = [...knownFeedback].filter((id) => !usedFeedback.has(id));
    if (omitted.length) throw new DecisionSpecError(`Feedback coverage incomplete: ${omitted.join(", ")}`, { omitted });
  }

  return {
    schema: DECISION_SPEC_SCHEMA,
    documentHash: spec.documentHash || null,
    title: typeof spec.title === "string" ? spec.title.trim() : "Feedback decisions",
    themes: normalizedThemes,
  };
}

export function createDecisionState(spec) {
  const decisions = new Map();
  let confirmed = false;

  function ensureEditable() {
    if (confirmed) throw new Error("Confirmed decisions are immutable. Start a new decision pass to change them.");
  }

  return {
    get confirmed() {
      return confirmed;
    },
    select(themeId, optionId) {
      ensureEditable();
      const theme = spec.themes.find((item) => item.id === themeId);
      if (!theme) throw new Error(`Unknown theme ${themeId}.`);
      if (optionId !== "other" && !theme.options.some((option) => option.id === optionId)) throw new Error(`Unknown option ${optionId}.`);
      const prior = decisions.get(themeId) || {};
      decisions.set(themeId, { ...prior, themeId, selectedOptionId: optionId, otherText: optionId === "other" ? prior.otherText || "" : null });
    },
    setOtherText(themeId, text) {
      ensureEditable();
      const prior = decisions.get(themeId);
      if (!prior || prior.selectedOptionId !== "other") throw new Error("Select Other before entering custom direction.");
      decisions.set(themeId, { ...prior, otherText: String(text) });
    },
    setNote(themeId, text) {
      ensureEditable();
      const prior = decisions.get(themeId) || { themeId, selectedOptionId: null, otherText: null };
      decisions.set(themeId, { ...prior, userNote: String(text) });
    },
    get(themeId) {
      return decisions.get(themeId) || null;
    },
    isComplete() {
      return spec.themes.every((theme) => {
        const value = decisions.get(theme.id);
        if (!value || !value.selectedOptionId) return false;
        return value.selectedOptionId !== "other" || Boolean(value.otherText && value.otherText.trim());
      });
    },
    confirm() {
      if (!this.isComplete()) throw new Error("Every theme must have a decision. Other requires custom text.");
      confirmed = true;
      return this.toPayload();
    },
    toPayload() {
      return {
        schema: DECISION_SELECTION_SCHEMA,
        documentHash: spec.documentHash,
        decisions: spec.themes.map((theme) => {
          const value = decisions.get(theme.id) || {};
          return {
            themeId: theme.id,
            selectedOptionId: value.selectedOptionId || null,
            otherText: value.otherText || null,
            userNote: value.userNote ? value.userNote.trim() : null,
          };
        }),
      };
    },
  };
}

function element(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text != null) node.textContent = text;
  return node;
}

export function renderDecisionPane(container, spec, state, options = {}) {
  container.innerHTML = "";
  const feedbackById = options.feedbackById || new Map();
  const heading = element("div", "ar-a2ui-heading");
  heading.appendChild(element("h2", "", spec.title || "Feedback decisions"));
  heading.appendChild(element("p", "ar-hint", "Choose one direction for each major feedback theme. No choice is preselected."));
  container.appendChild(heading);

  for (const theme of spec.themes) {
    const card = element("section", "ar-a2ui-theme");
    card.dataset.themeId = theme.id;
    card.appendChild(element("h3", "", theme.title));
    card.appendChild(element("p", "ar-a2ui-summary", theme.summary));
    const stakes = element("p", "ar-a2ui-stakes");
    stakes.appendChild(element("strong", "", "What is at stake: "));
    stakes.appendChild(document.createTextNode(theme.stakes));
    card.appendChild(stakes);

    const evidence = element("details", "ar-a2ui-evidence");
    evidence.appendChild(element("summary", "", `Feedback evidence (${theme.feedbackIds.length})`));
    const evidenceList = element("ul");
    for (const id of theme.feedbackIds) {
      const item = feedbackById.get(id);
      const label = item ? `${id} — ${item.author || "Unknown"}: ${item.text || item.anchorText || item.kind}` : id;
      evidenceList.appendChild(element("li", "", label));
    }
    evidence.appendChild(evidenceList);
    card.appendChild(evidence);

    const optionGrid = element("div", "ar-a2ui-options");
    const allOptions = [...theme.options, { id: "other", label: "Other", description: "Provide a different direction.", pros: [], cons: [] }];
    for (const option of allOptions) {
      const label = element("label", "ar-a2ui-option");
      const radio = document.createElement("input");
      radio.type = "radio";
      radio.name = `decision-${theme.id}`;
      radio.value = option.id;
      radio.checked = state.get(theme.id)?.selectedOptionId === option.id;
      radio.addEventListener("change", () => {
        state.select(theme.id, option.id);
        renderDecisionPane(container, spec, state, options);
      });
      label.appendChild(radio);
      const body = element("div");
      body.appendChild(element("strong", "", option.label));
      body.appendChild(element("p", "", option.description));
      if (option.pros.length || option.cons.length) {
        const tradeoffs = element("div", "ar-a2ui-tradeoffs");
        const pros = element("div");
        pros.appendChild(element("strong", "", "Pros"));
        const prosList = element("ul");
        option.pros.forEach((value) => prosList.appendChild(element("li", "", value)));
        pros.appendChild(prosList);
        const cons = element("div");
        cons.appendChild(element("strong", "", "Cons"));
        const consList = element("ul");
        option.cons.forEach((value) => consList.appendChild(element("li", "", value)));
        cons.appendChild(consList);
        tradeoffs.append(pros, cons);
        body.appendChild(tradeoffs);
      }
      label.appendChild(body);
      optionGrid.appendChild(label);
    }
    card.appendChild(optionGrid);

    if (state.get(theme.id)?.selectedOptionId === "other") {
      const other = document.createElement("textarea");
      other.className = "ar-a2ui-other";
      other.placeholder = "Describe the direction that should govern this theme.";
      other.value = state.get(theme.id)?.otherText || "";
      other.addEventListener("input", () => state.setOtherText(theme.id, other.value));
      card.appendChild(other);
    }
    const note = document.createElement("textarea");
    note.className = "ar-a2ui-note";
    note.placeholder = "Optional implementation note";
    note.value = state.get(theme.id)?.userNote || "";
    note.addEventListener("input", () => state.setNote(theme.id, note.value));
    card.appendChild(note);
    container.appendChild(card);
  }

  const controls = element("div", "ar-controls");
  const confirm = element("button", "ar-primary", "Confirm decisions");
  confirm.type = "button";
  confirm.disabled = !state.isComplete() || state.confirmed;
  confirm.addEventListener("click", () => {
    const payload = state.confirm();
    if (typeof options.onConfirm === "function") options.onConfirm(payload);
    renderDecisionPane(container, spec, state, options);
  });
  controls.appendChild(confirm);
  container.appendChild(controls);
}
