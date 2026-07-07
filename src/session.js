// Resumable session file (spec §5.4, M4 doc §6.2). Pure, sync -- base64 of the docx
// ArrayBuffer needs no async (only audit.js hashes; session.js's M4b session brief G-2:
// audit and session are separate artifacts with separate schemas -- never a shared
// serializer, and unlike audit, session round-trips).
import { STATES } from "./ui/state.js";
import { buildPrompt } from "./prompt.js";
import { validate } from "./validate.js";
import { parsePersona } from "./persona.js";

export const SESSION_SCHEMA_VERSION = 1;

const STATE_ORDER = Object.values(STATES);
function atLeast(state, floor) {
  return STATE_ORDER.indexOf(state) >= STATE_ORDER.indexOf(floor);
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function base64ToArrayBuffer(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

// `appStateSnapshot` is `{ state, context }` -- createAppState()'s own `.state`/`.context`
// -- plus `decisions`, the ratify rows' current `{ id, decision, reviewed }` (owned by
// app.js's ephemeral ratify.js state; never part of `context`, per build plan §3).
export function saveSession({ state, context, decisions = [] }) {
  return {
    schemaVersion: SESSION_SCHEMA_VERSION,
    state,
    filename: context.filename,
    docxBase64: context.docxBytes ? arrayBufferToBase64(context.docxBytes) : null,
    exported: context.exported,
    personaRef: context.persona ? { name: context.persona.name, raw: context.persona.raw } : null,
    promptVersion: context.promptVersion,
    response: context.response,
    decisions,
    savedAt: new Date().toISOString(),
  };
}

// Rebuilds `{ state, context }` for state.hydrate(). Deterministic re-derivation, not
// serialized storage: promptText/tokenEstimate are regenerated via buildPrompt (never
// serialized -- doc §6.2), and if a response was recorded, validate() is re-run against it
// to reproduce the IDENTICAL edit list (export + validate are pure functions of their
// inputs), which `decisions` then re-applies onto by `id` (the edit's validate-order
// index) via `context.pendingDecisions` -- app.js's ratify step applies these once, to the
// freshly-created ratification state, on entering RATIFYING.
export function loadSession(saved) {
  const persona = saved.personaRef ? parsePersona(saved.personaRef.raw, { filename: saved.personaRef.name }) : null;

  const context = {
    docxBytes: saved.docxBase64 ? base64ToArrayBuffer(saved.docxBase64) : null,
    filename: saved.filename,
    exported: saved.exported,
    persona,
    promptText: null,
    promptVersion: saved.promptVersion,
    tokenEstimate: null,
    documentWordCount: null,
    overThreshold: null,
    response: saved.response,
    validation: null,
    repairAttempts: {},
    validationAttempts: [],
    timestamps: { loaded: null, injected: null },
    pendingDecisions: saved.decisions || [],
  };

  if (atLeast(saved.state, STATES.PROMPT_READY)) {
    const built = buildPrompt({ persona, exportedMarkdown: saved.exported.markdown, filename: saved.filename });
    context.promptText = built.text;
    context.tokenEstimate = built.tokenEstimate;
    context.documentWordCount = built.documentWordCount;
    context.overThreshold = built.overThreshold;
  }

  if (atLeast(saved.state, STATES.RATIFYING) && saved.response != null) {
    const result = validate({
      responseMarkdown: saved.response,
      exportedMarkdown: saved.exported.markdown,
      sourceMap: saved.exported.sourceMap,
    });
    if (!result.ok) {
      throw new Error(
        `session: re-validating the saved response on load produced a failing result (gate ${result.gate}) though the saved ` +
          `state (${saved.state}) implies it had passed -- export/validate should be deterministic; something upstream isn't.`
      );
    }
    context.validation = result;
  }

  return { state: saved.state, context };
}
