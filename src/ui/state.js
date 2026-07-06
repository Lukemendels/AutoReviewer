// The Run Review flow's app state machine (M4 architecture doc §1), replacing the demo
// panel's implicit flags. Pure state + listener notification, same split as ratify.js's
// createRatificationState: the DOM only ever renders from `state`/`context`, never stores
// its own.
//
//   EMPTY -> DOC_LOADED -> PROMPT_READY -> AWAITING_RESPONSE -> VALIDATING
//          -> (VALIDATION_FAILED <-> AWAITING_RESPONSE)   [repair loop]
//          -> RATIFYING -> INJECTED
//
// Persona load is legal in any state >= DOC_LOADED and always regenerates the prompt
// (-> PROMPT_READY); discarding an un-ratified response on that jump is the caller's call
// to make (e.g. an explicit confirm dialog) -- this module just performs the reset once
// asked.
export const STATES = Object.freeze({
  EMPTY: "EMPTY",
  DOC_LOADED: "DOC_LOADED",
  PROMPT_READY: "PROMPT_READY",
  AWAITING_RESPONSE: "AWAITING_RESPONSE",
  VALIDATING: "VALIDATING",
  VALIDATION_FAILED: "VALIDATION_FAILED",
  RATIFYING: "RATIFYING",
  INJECTED: "INJECTED",
});

export function createAppState() {
  let state = STATES.EMPTY;
  const context = {
    docxBytes: null,
    filename: null,
    exported: null, // { markdown, sourceMap, comments, counts }
    persona: null, // null while default persona is in effect
    promptText: null,
    promptVersion: null,
    tokenEstimate: null,
    documentWordCount: null,
    overThreshold: null,
    response: null,
    validation: null,
    repairAttempts: {}, // gate -> count, for M4b's "advise restarting after 2" rule
  };
  const listeners = [];

  function notify() {
    for (const fn of listeners) fn();
  }
  function set(next, patch) {
    state = next;
    Object.assign(context, patch);
    notify();
  }
  function illegal(action) {
    throw new Error(`state.js: cannot ${action} from state ${state}`);
  }

  return {
    onChange(fn) {
      listeners.push(fn);
    },
    get state() {
      return state;
    },
    get context() {
      return context;
    },

    // EMPTY -> DOC_LOADED. Loading a second document mid-review is out of scope for M4a;
    // reset() first (a fresh page load also starts at EMPTY).
    loadDocument({ docxBytes, filename, exported }) {
      if (state !== STATES.EMPTY) illegal("loadDocument");
      set(STATES.DOC_LOADED, {
        docxBytes,
        filename,
        exported,
        persona: null,
        promptText: null,
        promptVersion: null,
        tokenEstimate: null,
        documentWordCount: null,
        overThreshold: null,
        response: null,
        validation: null,
        repairAttempts: {},
      });
    },

    // Legal from any state >= DOC_LOADED; always regenerates the prompt and lands back on
    // PROMPT_READY, discarding any in-flight response/validation.
    loadPersona(persona) {
      if (state === STATES.EMPTY) illegal("loadPersona");
      set(STATES.PROMPT_READY, { persona, response: null, validation: null, repairAttempts: {} });
    },

    // DOC_LOADED or PROMPT_READY (e.g. re-generated after a persona swap) -> PROMPT_READY.
    setPrompt({ text, tokenEstimate, promptVersion, documentWordCount, overThreshold }) {
      if (state !== STATES.DOC_LOADED && state !== STATES.PROMPT_READY) illegal("setPrompt");
      set(STATES.PROMPT_READY, { promptText: text, tokenEstimate, promptVersion, documentWordCount, overThreshold });
    },

    // PROMPT_READY -> AWAITING_RESPONSE: the human has copied the prompt and is off to
    // paste it into DHSChat.
    copyPrompt() {
      if (state !== STATES.PROMPT_READY) illegal("copyPrompt");
      set(STATES.AWAITING_RESPONSE, {});
    },

    // AWAITING_RESPONSE -> VALIDATING: the human pasted a reply back in.
    submitResponse(responseText) {
      if (state !== STATES.AWAITING_RESPONSE) illegal("submitResponse");
      set(STATES.VALIDATING, { response: responseText });
    },

    // VALIDATING -> RATIFYING on a passing validate() result.
    validationPassed(result) {
      if (state !== STATES.VALIDATING) illegal("validationPassed");
      set(STATES.RATIFYING, { validation: result });
    },

    // VALIDATING -> VALIDATION_FAILED on a blocking gate failure. Bumps the per-gate
    // repair-attempt counter (M4b's composeRepair reads this for the "restart advice
    // after 2" rule).
    validationFailed(result) {
      if (state !== STATES.VALIDATING) illegal("validationFailed");
      const repairAttempts = { ...context.repairAttempts };
      repairAttempts[result.gate] = (repairAttempts[result.gate] || 0) + 1;
      set(STATES.VALIDATION_FAILED, { validation: result, repairAttempts });
    },

    // VALIDATION_FAILED -> AWAITING_RESPONSE: ready for the next repair paste.
    acknowledgeFailure() {
      if (state !== STATES.VALIDATION_FAILED) illegal("acknowledgeFailure");
      set(STATES.AWAITING_RESPONSE, {});
    },

    // RATIFYING -> INJECTED.
    inject() {
      if (state !== STATES.RATIFYING) illegal("inject");
      set(STATES.INJECTED, {});
    },

    // Back to EMPTY, e.g. "start a new review."
    reset() {
      set(STATES.EMPTY, {
        docxBytes: null,
        filename: null,
        exported: null,
        persona: null,
        promptText: null,
        promptVersion: null,
        tokenEstimate: null,
        response: null,
        validation: null,
        repairAttempts: {},
      });
    },
  };
}
