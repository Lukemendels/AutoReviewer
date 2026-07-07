import { describe, expect, it } from "vitest";
import { createAppState, STATES } from "../src/ui/state.js";

function loadedState() {
  const s = createAppState();
  s.loadDocument({ docxBytes: new ArrayBuffer(0), filename: "doc", exported: { markdown: "md", sourceMap: {} } });
  return s;
}

describe("createAppState: the happy path", () => {
  it("walks EMPTY -> DOC_LOADED -> PROMPT_READY -> AWAITING_RESPONSE -> VALIDATING -> RATIFYING -> INJECTED", () => {
    const s = createAppState();
    expect(s.state).toBe(STATES.EMPTY);

    s.loadDocument({ docxBytes: new ArrayBuffer(0), filename: "doc", exported: { markdown: "md", sourceMap: {} } });
    expect(s.state).toBe(STATES.DOC_LOADED);

    s.setPrompt({ text: "prompt text", tokenEstimate: 10, promptVersion: "v1", documentWordCount: 2, overThreshold: false });
    expect(s.state).toBe(STATES.PROMPT_READY);

    s.copyPrompt();
    expect(s.state).toBe(STATES.AWAITING_RESPONSE);

    s.submitResponse("the response");
    expect(s.state).toBe(STATES.VALIDATING);
    expect(s.context.response).toBe("the response");

    s.validationPassed({ ok: true, edits: [] });
    expect(s.state).toBe(STATES.RATIFYING);

    s.inject();
    expect(s.state).toBe(STATES.INJECTED);
  });
});

describe("createAppState: the repair loop", () => {
  it("bounces VALIDATING -> VALIDATION_FAILED -> AWAITING_RESPONSE -> VALIDATING until it passes", () => {
    const s = loadedState();
    s.setPrompt({ text: "p", tokenEstimate: 1, promptVersion: "v1" });
    s.copyPrompt();
    s.submitResponse("bad response");
    s.validationFailed({ ok: false, gate: "G2", message: "drift" });
    expect(s.state).toBe(STATES.VALIDATION_FAILED);
    expect(s.context.repairAttempts.G2).toBe(1);

    s.acknowledgeFailure();
    expect(s.state).toBe(STATES.AWAITING_RESPONSE);

    s.submitResponse("bad response again");
    s.validationFailed({ ok: false, gate: "G2", message: "still drifting" });
    expect(s.context.repairAttempts.G2).toBe(2);

    s.acknowledgeFailure();
    s.submitResponse("good response");
    s.validationPassed({ ok: true, edits: [] });
    expect(s.state).toBe(STATES.RATIFYING);
  });
});

describe("createAppState: persona load is legal >= DOC_LOADED and regenerates the prompt", () => {
  it("throws if loaded before any document", () => {
    const s = createAppState();
    expect(() => s.loadPersona({ name: "x" })).toThrow();
  });

  it("jumps straight back to PROMPT_READY from any later state, discarding the in-flight response", () => {
    const s = loadedState();
    s.setPrompt({ text: "p1", tokenEstimate: 1, promptVersion: "v1" });
    s.copyPrompt();
    s.submitResponse("some response");
    expect(s.state).toBe(STATES.VALIDATING);

    s.loadPersona({ name: "New Persona" });
    expect(s.state).toBe(STATES.PROMPT_READY);
    expect(s.context.persona.name).toBe("New Persona");
    expect(s.context.response).toBeNull();
  });
});

describe("createAppState: illegal transitions throw rather than silently corrupting state", () => {
  it("rejects submitResponse before a prompt has been copied", () => {
    const s = loadedState();
    s.setPrompt({ text: "p", tokenEstimate: 1, promptVersion: "v1" });
    expect(() => s.submitResponse("x")).toThrow();
  });

  it("rejects inject before ratifying", () => {
    const s = createAppState();
    expect(() => s.inject()).toThrow();
  });

  it("rejects loading a second document without a reset", () => {
    const s = loadedState();
    expect(() => s.loadDocument({ docxBytes: new ArrayBuffer(0), filename: "doc2", exported: {} })).toThrow();
  });
});

describe("createAppState: onChange notifies on every transition", () => {
  it("fires once per state.set() call", () => {
    const s = createAppState();
    let calls = 0;
    s.onChange(() => calls++);
    s.loadDocument({ docxBytes: new ArrayBuffer(0), filename: "doc", exported: { markdown: "md", sourceMap: {} } });
    s.setPrompt({ text: "p", tokenEstimate: 1, promptVersion: "v1" });
    expect(calls).toBe(2);
  });
});

describe("createAppState: reset returns to a clean EMPTY", () => {
  it("clears every context field", () => {
    const s = loadedState();
    s.setPrompt({ text: "p", tokenEstimate: 1, promptVersion: "v1", documentWordCount: 5, overThreshold: false });
    s.reset();
    expect(s.state).toBe(STATES.EMPTY);
    expect(s.context.exported).toBeNull();
    expect(s.context.promptText).toBeNull();
    expect(s.context.documentWordCount).toBeNull();
    expect(s.context.overThreshold).toBeNull();
    expect(s.context.validationAttempts).toEqual([]);
    expect(s.context.timestamps).toEqual({ loaded: null, injected: null });
  });
});

describe("createAppState: validationAttempts log (M4b, audit provenance)", () => {
  it("appends an 'ok' entry on validationPassed and a gate+offset entry on validationFailed", () => {
    const s = loadedState();
    s.setPrompt({ text: "p", tokenEstimate: 1, promptVersion: "v1" });
    s.copyPrompt();

    s.submitResponse("bad response");
    s.validationFailed({ ok: false, gate: "G2", message: "drift", firstDivergence: { offset: 42 } });
    expect(s.context.validationAttempts).toHaveLength(1);
    expect(s.context.validationAttempts[0].result).toBe("G2");
    expect(s.context.validationAttempts[0].offset).toBe(42);
    expect(typeof s.context.validationAttempts[0].ts).toBe("string");

    s.acknowledgeFailure();
    s.submitResponse("good response");
    s.validationPassed({ ok: true, edits: [] });
    expect(s.context.validationAttempts).toHaveLength(2);
    expect(s.context.validationAttempts[1].result).toBe("ok");
    expect(s.context.validationAttempts[1].offset).toBeUndefined();
  });

  it("records rawStart (not offset) on a non-G2 gate failure", () => {
    const s = loadedState();
    s.setPrompt({ text: "p", tokenEstimate: 1, promptVersion: "v1" });
    s.copyPrompt();

    s.submitResponse("bad response");
    // G1/G3/G4 carry detail.rawStart (offset into the raw response), not firstDivergence.
    s.validationFailed({ ok: false, gate: "G3", message: "anchor on synthetic markdown", detail: { rawStart: 17 } });
    const entry = s.context.validationAttempts[0];
    expect(entry.result).toBe("G3");
    expect(entry.rawStart).toBe(17);
    expect(entry.offset).toBeUndefined();
  });

  it("starts empty on a fresh document load", () => {
    const s = loadedState();
    expect(s.context.validationAttempts).toEqual([]);
  });
});

describe("createAppState: timestamps (M4b, audit provenance)", () => {
  it("stamps loaded on loadDocument and injected on inject, leaving injected null until then", () => {
    const s = loadedState();
    expect(typeof s.context.timestamps.loaded).toBe("string");
    expect(s.context.timestamps.injected).toBeNull();

    s.setPrompt({ text: "p", tokenEstimate: 1, promptVersion: "v1" });
    s.copyPrompt();
    s.submitResponse("the response");
    s.validationPassed({ ok: true, edits: [] });
    expect(s.context.timestamps.injected).toBeNull();

    s.inject();
    expect(typeof s.context.timestamps.injected).toBe("string");
    expect(s.context.timestamps.loaded).not.toBeNull();
  });
});

describe("createAppState: hydrate (M4b Resume)", () => {
  it("restores state + context wholesale from EMPTY", () => {
    const s = createAppState();
    const restoredContext = {
      docxBytes: new ArrayBuffer(0),
      filename: "resumed-doc",
      exported: { markdown: "md", sourceMap: {} },
      persona: null,
      promptText: "p",
      promptVersion: "v1",
      tokenEstimate: 1,
      documentWordCount: 2,
      overThreshold: false,
      response: "the response",
      validation: { ok: true, edits: [] },
      repairAttempts: {},
      validationAttempts: [{ ts: "2026-07-07T00:00:00.000Z", result: "ok" }],
      timestamps: { loaded: "2026-07-07T00:00:00.000Z", injected: null },
    };
    s.hydrate({ state: STATES.RATIFYING, context: restoredContext });
    expect(s.state).toBe(STATES.RATIFYING);
    expect(s.context.filename).toBe("resumed-doc");
    expect(s.context.response).toBe("the response");
  });

  it("throws if called from any state other than EMPTY", () => {
    const s = loadedState();
    expect(() => s.hydrate({ state: STATES.RATIFYING, context: {} })).toThrow();
  });

  it("notifies listeners once", () => {
    const s = createAppState();
    let calls = 0;
    s.onChange(() => calls++);
    s.hydrate({ state: STATES.DOC_LOADED, context: { filename: "x" } });
    expect(calls).toBe(1);
  });
});
