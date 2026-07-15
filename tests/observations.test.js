// Step 1 of the reviewer-pass-slicer task: validates extractObservations's raw-observation
// table (kind/author/date/anchorText/docOrder/resolved/parentCommentId) against a synthetic
// two-reviewer, two-pass docx before any pass-clustering or UI work is built on top of it.
import { describe, expect, it } from "vitest";
import { DOMParser } from "@xmldom/xmldom";
import { extractObservations } from "../src/ooxml/observations.js";
import { buildAuditFixtureDocx } from "./helpers/auditFixture.js";

async function extractFixture() {
  const bytes = await buildAuditFixtureDocx();
  return extractObservations(bytes, { DOMParserImpl: DOMParser, filename: "audit-fixture" });
}

describe("extractObservations: two-reviewer, two-pass audit fixture", () => {
  it("returns one observation per tracked change and per comment/reply node", async () => {
    const { observations } = await extractFixture();
    expect(observations).toHaveLength(7);
    const kinds = observations.map((o) => o.kind).sort();
    expect(kinds).toEqual(["comment", "comment", "comment", "comment-reply", "deletion", "insertion", "insertion"].sort());
  });

  it("assigns docOrder strictly increasing in document-flow order", async () => {
    const { observations } = await extractFixture();
    const orders = observations.map((o) => o.docOrder);
    expect(orders).toEqual([...orders].sort((a, b) => a - b));
    expect(new Set(orders).size).toBe(orders.length);
  });

  it("keeps full ISO timestamps (not truncated to date-only) for both edits and comments", async () => {
    const { observations } = await extractFixture();
    const jimIns = observations.find((o) => o.kind === "insertion" && o.author === "Jim Smith");
    const jimComment = observations.find((o) => o.kind === "comment" && o.text === "Is this the right threshold?");
    expect(jimIns.date).toBe("2026-05-14T09:12:00Z");
    expect(jimComment.date).toBe("2026-05-14T09:30:00Z");
  });

  it("Jim Smith has two edits and two comments spanning a >48h gap (two distinct passes)", async () => {
    const { observations } = await extractFixture();
    const jim = observations.filter((o) => o.author === "Jim Smith");
    expect(jim).toHaveLength(4);
    const dates = jim.map((o) => new Date(o.date).getTime()).sort((a, b) => a - b);
    const gapHours = (dates[dates.length - 1] - dates[0]) / 36e5;
    expect(gapHours).toBeGreaterThan(48);
    // The pass-1 pair (insertion + first comment) and pass-2 pair (deletion + resolved
    // comment) are each within a few minutes of each other, i.e. the gap above is between
    // passes, not spread evenly across all four observations.
    const early = dates.filter((d) => d < new Date("2026-05-15T00:00:00Z").getTime());
    const late = dates.filter((d) => d >= new Date("2026-05-15T00:00:00Z").getTime());
    expect(early).toHaveLength(2);
    expect(late).toHaveLength(2);
  });

  it("Katie Chen has one insertion, one point comment, and one reply, all in one pass", async () => {
    const { observations } = await extractFixture();
    const katie = observations.filter((o) => o.author === "Katie Chen");
    expect(katie).toHaveLength(3);
    expect(katie.map((o) => o.kind).sort()).toEqual(["comment", "comment-reply", "insertion"]);
  });

  it("captures the anchored comment's span as anchorText", async () => {
    const { observations } = await extractFixture();
    const c = observations.find((o) => o.text === "Is this the right threshold?");
    expect(c.anchorText).toBe("This sentence has a discussion thread attached to it.");
  });

  it("falls back to the containing sentence for a point comment (no range)", async () => {
    const { observations } = await extractFixture();
    const c = observations.find((o) => o.text === "Does this point still hold after the revision above?");
    expect(c.anchorText).toBe("The deadline sits within this second sentence for testing.");
  });

  it("threads a reply to its parent via parentCommentId, sharing the parent's anchorText", async () => {
    const { observations } = await extractFixture();
    const parent = observations.find((o) => o.text === "Is this the right threshold?");
    const reply = observations.find((o) => o.kind === "comment-reply");
    expect(reply.author).toBe("Katie Chen");
    expect(reply.parentCommentId).toBe("0");
    expect(reply.anchorText).toBe(parent.anchorText);
  });

  it("carries resolved status per comment (true/false), null/not-applicable for edits", async () => {
    const { observations } = await extractFixture();
    const resolved = observations.find((o) => o.text === "Please cite the authority here.");
    const unresolved = observations.find((o) => o.text === "Is this the right threshold?");
    expect(resolved.resolved).toBe(true);
    expect(unresolved.resolved).toBe(false);
    const edit = observations.find((o) => o.kind === "insertion");
    expect(edit.resolved).toBeNull();
    expect(edit.parentCommentId).toBeNull();
    expect(edit.anchorText).toBeNull();
  });

  it("exports comment text verbatim, with no summarization or truncation", async () => {
    const { observations } = await extractFixture();
    const reply = observations.find((o) => o.kind === "comment-reply");
    expect(reply.text).toBe("Agreed -- flagged for legal review.");
  });
});
