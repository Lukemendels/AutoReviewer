// The ratification UI (spec §8). Split into a pure, fully-tested state machine
// (createRatificationState) and a DOM renderer built on top of it (renderRatificationUI),
// since there's no DOM available where most of this project's tests run -- see the M2 plan
// for why (in short: only files that need one pay the happy-dom cost).
import { diffWords } from "./diff.js";

export function createRatificationState(edits) {
  const rows = edits.map((edit, id) => ({ id, edit, decision: "accept", reviewed: false }));
  const listeners = [];

  function notify() {
    for (const fn of listeners) fn();
  }
  function row(id) {
    return rows.find((r) => r.id === id);
  }

  return {
    rows,
    onChange(fn) {
      listeners.push(fn);
    },
    setDecision(id, decision) {
      const r = row(id);
      if (!r) return;
      r.decision = decision;
      notify();
    },
    acceptAll() {
      for (const r of rows) r.decision = "accept";
      notify();
    },
    rejectAll() {
      for (const r of rows) r.decision = "reject";
      notify();
    },
    markReviewed(id) {
      const r = row(id);
      if (!r || r.reviewed) return;
      r.reviewed = true;
      notify();
    },
    // Every row reviewed (an empty list trivially qualifies) -- spec §8: "the human must
    // scroll the full list before Inject enables," no always-accept escape hatch.
    canInject() {
      return rows.every((r) => r.reviewed);
    },
    acceptedEdits() {
      return rows.filter((r) => r.decision === "accept").map((r) => r.edit);
    },
    nextUnreviewedId() {
      const r = rows.find((r) => !r.reviewed);
      return r ? r.id : null;
    },
  };
}

function editSpan(edit) {
  const start = edit.mdStart ?? edit.mdPos;
  const end = edit.mdEnd ?? edit.mdPos;
  return { start, end };
}

function contextSnippet(sourceText, edit, pad = 80) {
  const { start, end } = editSpan(edit);
  const from = Math.max(0, start - pad);
  const to = Math.min(sourceText.length, end + pad);
  return { before: sourceText.slice(from, start), highlighted: sourceText.slice(start, end), after: sourceText.slice(end, to) };
}

function typeBadgeText(edit) {
  if (edit.type === "comment") return edit.anchored ? "COMMENT" : "COMMENT (point)";
  return edit.type.toUpperCase();
}

function buildChangeEl(doc, edit) {
  const el = doc.createElement("div");
  el.className = "ar-row-change";
  if (edit.type === "ins") {
    const ins = doc.createElement("ins");
    ins.textContent = edit.newText;
    el.appendChild(ins);
  } else if (edit.type === "del") {
    const del = doc.createElement("del");
    del.textContent = edit.oldText;
    el.appendChild(del);
  } else if (edit.type === "sub") {
    for (const seg of diffWords(edit.oldText, edit.newText)) {
      const node = seg.type === "same" ? doc.createElement("span") : doc.createElement(seg.type === "del" ? "del" : "ins");
      node.textContent = seg.text;
      el.appendChild(node);
    }
  } else {
    const span = doc.createElement("span");
    span.textContent = edit.commentText;
    el.appendChild(span);
  }
  return el;
}

function buildRowEl(doc, state, row, sourceText) {
  const el = doc.createElement("div");
  el.className = "ar-row";
  el.dataset.rowId = String(row.id);

  const badge = doc.createElement("span");
  badge.className = "ar-badge";
  badge.textContent = typeBadgeText(row.edit);
  el.appendChild(badge);

  const snippet = contextSnippet(sourceText, row.edit);
  const snippetEl = doc.createElement("div");
  snippetEl.className = "ar-snippet";
  snippetEl.append(doc.createTextNode(snippet.before));
  const mark = doc.createElement("mark");
  mark.textContent = snippet.highlighted || "•";
  snippetEl.appendChild(mark);
  snippetEl.append(doc.createTextNode(snippet.after));
  el.appendChild(snippetEl);

  el.appendChild(buildChangeEl(doc, row.edit));

  const controls = doc.createElement("div");
  controls.className = "ar-row-controls";
  const acceptBtn = doc.createElement("button");
  acceptBtn.type = "button";
  acceptBtn.dataset.action = "accept";
  acceptBtn.textContent = "Accept";
  acceptBtn.addEventListener("click", () => state.setDecision(row.id, "accept"));
  const rejectBtn = doc.createElement("button");
  rejectBtn.type = "button";
  rejectBtn.dataset.action = "reject";
  rejectBtn.textContent = "Reject";
  rejectBtn.addEventListener("click", () => state.setDecision(row.id, "reject"));
  controls.append(acceptBtn, rejectBtn);
  el.appendChild(controls);

  return el;
}

function observeForReview(doc, rowEls, state) {
  const IO = typeof IntersectionObserver !== "undefined" ? IntersectionObserver : null;
  if (!IO) return; // no fallback: target browser (Chromium Edge) always has this.
  const observer = new IO(
    (entries) => {
      for (const entry of entries) {
        if (entry.isIntersecting) state.markReviewed(Number(entry.target.dataset.rowId));
      }
    },
    { threshold: 0.5 }
  );
  for (const el of rowEls) observer.observe(el);
}

export function renderRatificationUI(container, state, { sourceText = "" } = {}) {
  const doc = container.ownerDocument || document;
  container.innerHTML = "";

  const bulk = doc.createElement("div");
  bulk.className = "ar-bulk-controls";
  const acceptAllBtn = doc.createElement("button");
  acceptAllBtn.type = "button";
  acceptAllBtn.dataset.action = "accept-all";
  acceptAllBtn.textContent = "Accept all";
  acceptAllBtn.addEventListener("click", () => state.acceptAll());
  const rejectAllBtn = doc.createElement("button");
  rejectAllBtn.type = "button";
  rejectAllBtn.dataset.action = "reject-all";
  rejectAllBtn.textContent = "Reject all";
  rejectAllBtn.addEventListener("click", () => state.rejectAll());
  const jumpNextBtn = doc.createElement("button");
  jumpNextBtn.type = "button";
  jumpNextBtn.dataset.action = "jump-next";
  jumpNextBtn.textContent = "Jump to next unreviewed";
  jumpNextBtn.addEventListener("click", () => {
    const id = state.nextUnreviewedId();
    if (id == null) return;
    container.querySelector(`[data-row-id="${id}"]`)?.scrollIntoView({ block: "center" });
  });
  bulk.append(acceptAllBtn, rejectAllBtn, jumpNextBtn);
  container.appendChild(bulk);

  const list = doc.createElement("div");
  list.className = "ar-row-list";
  const rowEls = state.rows.map((row) => buildRowEl(doc, state, row, sourceText));
  for (const el of rowEls) list.appendChild(el);
  container.appendChild(list);

  const injectBtn = doc.createElement("button");
  injectBtn.type = "button";
  injectBtn.className = "ar-inject";
  injectBtn.dataset.action = "inject";
  injectBtn.textContent = "Inject accepted edits";
  container.appendChild(injectBtn);

  function syncInjectEnabled() {
    injectBtn.disabled = !state.canInject();
  }
  syncInjectEnabled();
  state.onChange(syncInjectEnabled);

  observeForReview(doc, rowEls, state);
}
