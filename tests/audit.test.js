import { describe, expect, it } from "vitest";
import { buildAuditRecord, APP_VERSION } from "../src/audit.js";

// Deterministic stand-in for crypto.subtle.digest (M4b session brief G-4): a fixed-size
// digest that's a pure function of the input bytes, so hashes are assertable for
// equality/inequality without pinning real SHA-256 values or touching WebCrypto.
function fakeDigestImpl(bytes) {
  const arr = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let h = 2166136261 >>> 0;
  for (const b of arr) {
    h ^= b;
    h = Math.imul(h, 16777619) >>> 0;
  }
  const out = new Uint8Array(32);
  for (let i = 0; i < 32; i++) out[i] = ((h >>> ((i % 4) * 8)) & 0xff) ^ i;
  return Promise.resolve(out.buffer);
}
async function fakeSha256Tag(input) {
  const bytes = typeof input === "string" ? new TextEncoder().encode(input) : input;
  const digest = await fakeDigestImpl(bytes);
  return "sha256-" + [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

const DOCX_BYTES = new TextEncoder().encode("stand-in original docx bytes").buffer;
const OUTPUT_BYTES = new TextEncoder().encode("stand-in output docx bytes").buffer;
const RESPONSE = "the pasted response text";

async function baseDetails(overrides = {}) {
  return {
    promptVersion: "m4a-2026.07-1",
    timestamps: { loaded: "2026-07-07T00:00:00.000Z", injected: "2026-07-07T00:05:00.000Z" },
    filename: "policy-draft",
    docxBytes: DOCX_BYTES,
    outputBytes: OUTPUT_BYTES,
    response: RESPONSE,
    sourceMap: { docHash: await fakeSha256Tag(DOCX_BYTES) },
    persona: null,
    validationAttempts: [{ ts: "2026-07-07T00:04:00.000Z", result: "ok" }],
    rows: [
      { id: 0, edit: { type: "ins", newText: "This entire line is a new inserted paragraph.", anchor: { kind: "paragraphBoundary", edge: "before", bodyPath: [0] } }, decision: "accept", reviewed: true },
      { id: 1, edit: { type: "comment", commentText: "Double-check this framing.", anchor: { bodyPath: [1], runIndex: 0, charStart: 4, charEnd: 4 } }, decision: "reject", reviewed: true },
    ],
    author: "AutoReviewer — Default Persona",
    ...overrides,
  };
}

describe("buildAuditRecord: every spec-§12 field is present", () => {
  it("assembles the full record", async () => {
    const record = await buildAuditRecord(await baseDetails(), { digestImpl: fakeDigestImpl });
    expect(record.schemaVersion).toBe(1);
    expect(record.appVersion).toBe(APP_VERSION);
    expect(record.promptVersion).toBe("m4a-2026.07-1");
    expect(record.timestamps).toEqual({ loaded: "2026-07-07T00:00:00.000Z", injected: "2026-07-07T00:05:00.000Z" });
    expect(record.source.filename).toBe("policy-draft");
    expect(record.source.sha256).toMatch(/^sha256-[0-9a-f]{64}$/);
    expect(record.output.sha256).toMatch(/^sha256-[0-9a-f]{64}$/);
    expect(record.response.sha256).toMatch(/^sha256-[0-9a-f]{64}$/);
    expect(record.persona).toBeTruthy();
    expect(record.validationAttempts).toEqual([{ ts: "2026-07-07T00:04:00.000Z", result: "ok" }]);
    expect(record.injection.author).toBe("AutoReviewer — Default Persona");
    expect(Array.isArray(record.edits)).toBe(true);
  });

  it("source.sha256 === sourceMap.docHash (same original bytes, cheap consistency check)", async () => {
    const details = await baseDetails();
    const record = await buildAuditRecord(details, { digestImpl: fakeDigestImpl });
    expect(record.source.sha256).toBe(details.sourceMap.docHash);
  });

  it("throws if source.sha256 and sourceMap.docHash disagree (would mean different original bytes)", async () => {
    const details = await baseDetails({ sourceMap: { docHash: "sha256-" + "0".repeat(64) } });
    await expect(buildAuditRecord(details, { digestImpl: fakeDigestImpl })).rejects.toThrow(/does not match/);
  });

  it("distinct inputs (source/output/response) hash to distinct digests", async () => {
    const record = await buildAuditRecord(await baseDetails(), { digestImpl: fakeDigestImpl });
    const hashes = new Set([record.source.sha256, record.output.sha256, record.response.sha256]);
    expect(hashes.size).toBe(3);
  });
});

describe("buildAuditRecord: persona", () => {
  it("DEFAULT persona (persona: null) omits sha256, records only the built-in name", async () => {
    const record = await buildAuditRecord(await baseDetails({ persona: null }), { digestImpl: fakeDigestImpl });
    expect(record.persona).toEqual({ name: "Default Persona (built-in)" });
  });

  it("a custom persona records name + a hash of its raw markdown", async () => {
    const persona = { name: "Jim — editorial", raw: "---\nname: Jim\n---\n## Role and voice\n..." };
    const record = await buildAuditRecord(await baseDetails({ persona }), { digestImpl: fakeDigestImpl });
    expect(record.persona.name).toBe("Jim — editorial");
    expect(record.persona.sha256).toBe(await fakeSha256Tag(persona.raw));
  });
});

describe("buildAuditRecord: edits[] mirrors ratify decisions, resolvedAnchor per §5 decision", () => {
  it("includes every ratified edit (accept AND reject), each carrying its decision + resolvedAnchor", async () => {
    const details = await baseDetails();
    const record = await buildAuditRecord(details, { digestImpl: fakeDigestImpl });
    expect(record.edits).toHaveLength(2);
    expect(record.edits[0]).toMatchObject({ id: "e0", type: "ins", decision: "accept" });
    expect(record.edits[0].resolvedAnchor).toEqual(details.rows[0].edit.anchor);
    expect(record.edits[1]).toMatchObject({ id: "e1", type: "comment", decision: "reject" });
    expect(record.edits[1].resolvedAnchor).toEqual(details.rows[1].edit.anchor);
  });
});

describe("buildAuditRecord: injection.counts matches what was actually injected", () => {
  it("counts only ACCEPTED edits by type, ignoring rejected ones", async () => {
    const details = await baseDetails({
      rows: [
        { id: 0, edit: { type: "ins", newText: "a", anchor: {} }, decision: "accept", reviewed: true },
        { id: 1, edit: { type: "ins", newText: "b", anchor: {} }, decision: "reject", reviewed: true },
        { id: 2, edit: { type: "del", oldText: "c", anchor: [] }, decision: "accept", reviewed: true },
        { id: 3, edit: { type: "sub", oldText: "d", newText: "e", anchor: [] }, decision: "accept", reviewed: true },
        { id: 4, edit: { type: "comment", commentText: "f", anchor: {} }, decision: "reject", reviewed: true },
      ],
    });
    const record = await buildAuditRecord(details, { digestImpl: fakeDigestImpl });
    expect(record.injection.counts).toEqual({ ins: 1, del: 1, sub: 1, comment: 0 });
  });
});

describe("buildAuditRecord: no load path (immutability, G-3)", () => {
  it("the module exposes no loadAudit / parse function", async () => {
    const auditModule = await import("../src/audit.js");
    expect(auditModule.loadAudit).toBeUndefined();
    expect(Object.keys(auditModule).sort()).toEqual(["APP_VERSION", "buildAuditRecord"]);
  });
});
