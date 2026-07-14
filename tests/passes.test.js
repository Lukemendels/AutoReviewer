// Step 2 of the reviewer-pass-slicer task: pass clustering on top of extractObservations.
import { describe, expect, it } from "vitest";
import { DOMParser } from "@xmldom/xmldom";
import { extractObservations } from "../src/ooxml/observations.js";
import { clusterPasses, PASS_GAP_HOURS } from "../src/passes.js";
import { buildAuditFixtureDocx } from "./helpers/auditFixture.js";
import { buildStrippedMetadataFixtureDocx } from "./helpers/strippedMetadataFixture.js";
import { buildUndatedObservationsFixtureDocx } from "./helpers/undatedObservationsFixture.js";

async function clustersFor(buildFixture) {
  const bytes = await buildFixture();
  const { observations } = await extractObservations(bytes, { DOMParserImpl: DOMParser, filename: "fixture" });
  return clusterPasses(observations);
}

describe("clusterPasses: two-reviewer, two-pass audit fixture", () => {
  it("splits Jim's observations into two passes >48h apart and keeps Katie in one", async () => {
    const { passes, metadataStripped } = await clustersFor(buildAuditFixtureDocx);
    expect(metadataStripped).toBe(false);

    const jimPasses = passes.filter((p) => p.author === "Jim Smith");
    const katiePasses = passes.filter((p) => p.author === "Katie Chen");
    expect(jimPasses).toHaveLength(2);
    expect(katiePasses).toHaveLength(1);

    expect(jimPasses.map((p) => p.passDate)).toEqual(["2026-05-14", "2026-05-16"]);
    expect(jimPasses.every((p) => !p.undated)).toBe(true);
  });

  it("labels each pass with author, earliest date, and edit/comment counts", async () => {
    const { passes } = await clustersFor(buildAuditFixtureDocx);
    const pass1 = passes.find((p) => p.author === "Jim Smith" && p.passDate === "2026-05-14");
    expect(pass1.label).toBe("Jim Smith — 2026-05-14 (1 edit, 1 comment, 0 replies)");
    expect(pass1.counts).toEqual({ insertions: 1, deletions: 0, comments: 1, replies: 0 });

    const katie = passes.find((p) => p.author === "Katie Chen");
    // insertion + point comment + reply to Jim's pass-1 thread, all one pass for Katie
    expect(katie.counts).toEqual({ insertions: 1, deletions: 0, comments: 1, replies: 1 });
    expect(katie.label).toBe("Katie Chen — 2026-05-15 (1 edit, 1 comment, 1 reply)");
  });

  it("keeps each pass's observations in document order, not date order", async () => {
    const { passes } = await clustersFor(buildAuditFixtureDocx);
    for (const pass of passes) {
      const orders = pass.observations.map((o) => o.docOrder);
      expect(orders).toEqual([...orders].sort((a, b) => a - b));
    }
  });

  it("PASS_GAP_HOURS is tunable via options.gapHours", async () => {
    const bytes = await buildAuditFixtureDocx();
    const { observations } = await extractObservations(bytes, { DOMParserImpl: DOMParser, filename: "fixture" });
    expect(PASS_GAP_HOURS).toBe(48);
    // A 72h threshold swallows Jim's 48.5h gap into a single pass.
    const { passes } = clusterPasses(observations, { gapHours: 72 });
    const jimPasses = passes.filter((p) => p.author === "Jim Smith");
    expect(jimPasses).toHaveLength(1);
  });
});

describe("clusterPasses: metadata-stripped document (Word's 'remove personal info')", () => {
  it("collapses to one pseudo-pass and sets metadataStripped", async () => {
    const { passes, metadataStripped } = await clustersFor(buildStrippedMetadataFixtureDocx);
    expect(metadataStripped).toBe(true);
    expect(passes).toHaveLength(1);
    expect(passes[0].author).toBe("Author");
    expect(passes[0].label).toBe("Author (metadata stripped)");
    expect(passes[0].undated).toBe(true);
    expect(passes[0].observations).toHaveLength(2); // one insertion + one comment
  });

  it("does not crash and does not guess dates", async () => {
    const { passes } = await clustersFor(buildStrippedMetadataFixtureDocx);
    expect(passes[0].observations.every((o) => o.date === null)).toBe(true);
    expect(passes[0].passDate).toBeNull();
  });
});

describe("clusterPasses: undated observations for an author with dated passes", () => {
  it("buckets an undated observation into its own pseudo-pass when there are 2+ dated passes", async () => {
    const { passes, metadataStripped } = await clustersFor(buildUndatedObservationsFixtureDocx);
    expect(metadataStripped).toBe(false);

    const al = passes.filter((p) => p.author === "Al");
    expect(al).toHaveLength(3);
    const undatedAl = al.find((p) => p.undated);
    expect(undatedAl.observations).toHaveLength(1);
    expect(undatedAl.observations[0].kind).toBe("comment");
    expect(undatedAl.label).toBe("Al — undated (0 edits, 1 comment, 0 replies)");
  });

  it("attaches an undated observation to the single existing pass when there's only one", async () => {
    const { passes } = await clustersFor(buildUndatedObservationsFixtureDocx);
    const bo = passes.filter((p) => p.author === "Bo");
    expect(bo).toHaveLength(1);
    expect(bo[0].undated).toBe(false);
    expect(bo[0].observations).toHaveLength(2); // the dated insertion + the undated comment
    expect(bo[0].counts).toEqual({ insertions: 1, deletions: 0, comments: 1, replies: 0 });
  });
});
