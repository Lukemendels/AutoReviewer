// @vitest-environment happy-dom
import { describe, expect, it } from "vitest";
import { createRatificationState, renderRatificationUI } from "../src/ui/ratify.js";

const edits = [
  { type: "sub", mdStart: 10, mdEnd: 15, oldText: "shall", newText: "must", rawStart: 0, rawEnd: 10 },
  { type: "del", mdStart: 20, mdEnd: 26, oldText: "second", rawStart: 20, rawEnd: 30 },
  { type: "ins", mdPos: 40, newText: "really ", rawStart: 40, rawEnd: 50 },
  { type: "comment", anchored: false, mdPos: 60, commentText: "Double-check this.", rawStart: 60, rawEnd: 70 },
];

describe("createRatificationState: pure state machine", () => {
  it("defaults every row to accepted and unreviewed", () => {
    const state = createRatificationState(edits);
    expect(state.rows.length).toBe(4);
    expect(state.rows.every((r) => r.decision === "accept")).toBe(true);
    expect(state.rows.every((r) => r.reviewed === false)).toBe(true);
    expect(state.canInject()).toBe(false);
  });

  it("canInject() only becomes true once every row is marked reviewed", () => {
    const state = createRatificationState(edits);
    for (const row of state.rows.slice(0, -1)) state.markReviewed(row.id);
    expect(state.canInject()).toBe(false);
    state.markReviewed(state.rows[state.rows.length - 1].id);
    expect(state.canInject()).toBe(true);
  });

  it("setDecision toggles a single row without affecting others", () => {
    const state = createRatificationState(edits);
    state.setDecision(state.rows[1].id, "reject");
    expect(state.rows[1].decision).toBe("reject");
    expect(state.rows[0].decision).toBe("accept");
  });

  it("acceptAll / rejectAll set every row", () => {
    const state = createRatificationState(edits);
    state.rejectAll();
    expect(state.rows.every((r) => r.decision === "reject")).toBe(true);
    state.acceptAll();
    expect(state.rows.every((r) => r.decision === "accept")).toBe(true);
  });

  it("acceptedEdits() returns only accepted rows' edits, in order", () => {
    const state = createRatificationState(edits);
    state.setDecision(state.rows[1].id, "reject");
    const accepted = state.acceptedEdits();
    expect(accepted.length).toBe(3);
    expect(accepted).not.toContain(edits[1]);
  });

  it("an empty edit list trivially allows injection", () => {
    const state = createRatificationState([]);
    expect(state.canInject()).toBe(true);
  });

  it("notifies onChange listeners on every mutation", () => {
    const state = createRatificationState(edits);
    let calls = 0;
    state.onChange(() => calls++);
    state.setDecision(state.rows[0].id, "reject");
    state.markReviewed(state.rows[0].id);
    state.acceptAll();
    expect(calls).toBe(3);
  });
});

describe("renderRatificationUI: DOM smoke test", () => {
  const sourceText =
    "The rule shall apply to all second carriers, and third notes follow after this point in the document for padding.";
  const domEdits = [
    { type: "sub", mdStart: 9, mdEnd: 14, oldText: "shall", newText: "must", rawStart: 0, rawEnd: 0 },
    { type: "del", mdStart: 28, mdEnd: 34, oldText: "second", rawStart: 0, rawEnd: 0 },
  ];

  it("renders one row per edit, with a type badge and an inline diff for sub", () => {
    const state = createRatificationState(domEdits);
    const container = document.createElement("div");
    renderRatificationUI(container, state, { sourceText });

    const rows = container.querySelectorAll("[data-row-id]");
    expect(rows.length).toBe(2);
    expect(rows[0].textContent).toMatch(/sub/i);
    expect(rows[1].textContent).toMatch(/del/i);
    // the sub row's inline diff should show both the old and new word
    expect(rows[0].textContent).toContain("shall");
    expect(rows[0].textContent).toContain("must");
  });

  it("wires a row's reject button to the state", () => {
    const state = createRatificationState(domEdits);
    const container = document.createElement("div");
    renderRatificationUI(container, state, { sourceText });

    const rejectBtn = container.querySelector('[data-row-id="0"] [data-action="reject"]');
    rejectBtn.dispatchEvent(new Event("click", { bubbles: true }));
    expect(state.rows[0].decision).toBe("reject");
  });

  it("Inject stays disabled until every row is marked reviewed, then re-enables reactively", () => {
    const state = createRatificationState(domEdits);
    const container = document.createElement("div");
    renderRatificationUI(container, state, { sourceText });

    const injectBtn = container.querySelector('[data-action="inject"]');
    expect(injectBtn.disabled).toBe(true);
    for (const row of state.rows) state.markReviewed(row.id);
    expect(injectBtn.disabled).toBe(false);
  });

  it("bulk accept-all / reject-all buttons drive the state", () => {
    const state = createRatificationState(domEdits);
    const container = document.createElement("div");
    renderRatificationUI(container, state, { sourceText });

    container.querySelector('[data-action="reject-all"]').dispatchEvent(new Event("click", { bubbles: true }));
    expect(state.rows.every((r) => r.decision === "reject")).toBe(true);

    container.querySelector('[data-action="accept-all"]').dispatchEvent(new Event("click", { bubbles: true }));
    expect(state.rows.every((r) => r.decision === "accept")).toBe(true);
  });
});
