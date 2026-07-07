// Audit sidecar assembly (spec §12, M4 doc §6.2). Async only for its own SHA-256 hashing --
// a build path only, no load path (M4b session brief G-3: audit is immutable provenance;
// session.js is the separate, mutable, round-tripping artifact -- G-2, the two are never
// merged, no shared serializer).
//
// `digestImpl` is injected, defaulting to crypto.subtle.digest -- mirrors load.js's
// DOMParserImpl pattern (G-4) so tests stay pure/deterministic and a file:// / Node
// webcrypto gap can't break the suite.

// Matches package.json's version -- hand-maintained the same way prompt.js's
// PROMPT_TEMPLATE_VERSION is, since this is provenance metadata (what shipped this
// review), not a build-injected constant.
export const APP_VERSION = "0.1.0";

function toBytes(input) {
  return typeof input === "string" ? new TextEncoder().encode(input) : input;
}

async function sha256Hex(digestImpl, input) {
  const digest = await digestImpl(toBytes(input));
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("");
}
async function sha256Tag(digestImpl, input) {
  return "sha256-" + (await sha256Hex(digestImpl, input));
}

function excerpt(edit, pad = 60) {
  const text =
    edit.type === "ins"
      ? edit.newText
      : edit.type === "del"
        ? edit.oldText
        : edit.type === "sub"
          ? `${edit.oldText} -> ${edit.newText}`
          : edit.commentText;
  return text.length > pad ? text.slice(0, pad) + "…" : text;
}

// `details`: { promptVersion, timestamps, filename, docxBytes, outputBytes, response,
// sourceMap, persona, validationAttempts, rows, author }. `rows` are ratify.js's rows --
// `{ id, edit, decision, reviewed }` -- covering EVERY ratified edit (accept and reject
// alike; spec §12 wants the full decision record). `edit.anchor` (set by validate.js's
// resolveEditAnchor for every resolved edit, accepted or not) is the resolvedAnchor.
export async function buildAuditRecord(details, { digestImpl = (data) => crypto.subtle.digest("SHA-256", data) } = {}) {
  const { promptVersion, timestamps, filename, docxBytes, outputBytes, response, sourceMap, persona, validationAttempts, rows, author } =
    details;

  const sourceSha = await sha256Tag(digestImpl, docxBytes);
  if (sourceSha !== sourceMap.docHash) {
    throw new Error(
      `audit: source.sha256 (${sourceSha}) does not match sourceMap.docHash (${sourceMap.docHash}) -- both hash the same ` +
        `original docx bytes and must agree`
    );
  }

  const personaRecord = persona ? { name: persona.name, sha256: await sha256Tag(digestImpl, persona.raw) } : { name: "Default Persona (built-in)" };

  const edits = rows.map((row) => ({
    id: `e${row.id}`,
    type: row.edit.type,
    excerpt: excerpt(row.edit),
    decision: row.decision,
    resolvedAnchor: row.edit.anchor,
  }));

  const counts = { ins: 0, del: 0, sub: 0, comment: 0 };
  for (const row of rows) if (row.decision === "accept") counts[row.edit.type]++;

  return {
    schemaVersion: 1,
    appVersion: APP_VERSION,
    promptVersion,
    timestamps,
    source: { filename, sha256: sourceSha },
    output: { sha256: await sha256Tag(digestImpl, outputBytes) },
    response: { sha256: await sha256Tag(digestImpl, response) },
    persona: personaRecord,
    validationAttempts,
    edits,
    injection: { author, counts },
  };
}
