// Writes the reviewer-pass-slicer step-1 audit fixture (tests/helpers/auditFixture.js) to
// disk for manual inspection via scripts/dump-observations.mjs. Not a committed fixture --
// see auditFixture.js's header for why (hand-authored XML, not fixtures/generate.py).
import { writeFileSync } from "node:fs";
import { buildAuditFixtureDocx } from "../tests/helpers/auditFixture.js";

const outPath = process.argv[2];
if (!outPath) {
  console.error("Usage: node scripts/gen-audit-fixture.mjs <out.docx>");
  process.exit(1);
}

const bytes = await buildAuditFixtureDocx();
writeFileSync(outPath, Buffer.from(bytes));
console.log(`Wrote audit fixture to ${outPath}`);
